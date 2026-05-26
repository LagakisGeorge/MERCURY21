VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form uFind_pelat 
   Caption         =   "Form2"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   9870
   Begin VB.CommandButton Command1 
      Caption         =   "Εξοδος"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select left(EPO+AFM+DIE+THL,30) as xx from pel where eidos='e';"
      Top             =   2640
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton Option1 
         Caption         =   "Προμηθευτής"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Πελάτης"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "uFind_Pelat.frx":0000
      Height          =   2415
      Left            =   3240
      OleObjectBlob   =   "uFind_Pelat.frx":0014
      TabIndex        =   10
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "A.Φ.Μ."
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Επωνυμία"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Κωδικός"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "uFind_pelat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
          gKod_pel = gbuff + Right(Trim(Data1.Recordset("xx")), 4)
          
          
          Unload Me
     End If
End Sub

Private Sub Form_Activate()
   Data1.DatabaseName = gDir
   Data1.Connect = gConnect
   Data1.RecordSource = "select LEFT(EPO+'-'+DIE+'-'+KOD,60) as xx from pel where eidos='" + gbuff + "';"
   DBGrid1.Columns(0).Width = 3000
   DBGrid1.Columns(0).DataField = "xx"
   Data1.Refresh
End Sub


Private Sub Text2_LostFocus(Index As Integer)
Dim x As String

If Len(Trim(Text2(Index))) = 0 Then Exit Sub


x = gbuff 'If Option1(1) Then x = "e" Else x = "r"
If Index = 1 Then  'eponumo
    
    ' "select LEFT(EPO+AFM+DIE+THL,30) as xx from pel where eidos='" + gbuff + "';"
    Data1.RecordSource = "select LEFT(EPO+'-'+DIE+'-'+KOD,60) as xx from pel where eidos='" + x + "' and epo like '" + Text2(1).Text + "*';"
End If
If Index = 0 Then 'kodikos
    Data1.RecordSource = "select  LEFT(EPO+'-'+DIE+'-'+KOD,60) as xx from pel where kod='" + x + "' and epo like '" + Text2(0).Text + "*';"
End If
If Index = 2 Then 'afm
    Data1.RecordSource = "select  LEFT(EPO+'-'+DIE+'-'+KOD,60) AS XX from pel where eidos='" + x + "' and afm like '" + Text2(2).Text + "*';"
End If
DBGrid1.Columns(0).Width = 3000
Data1.Refresh

End Sub

