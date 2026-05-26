VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form apot_searching 
   Caption         =   "Form1"
   ClientHeight    =   11625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   ScaleHeight     =   11625
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox listes 
      BackColor       =   &H00FFFF80&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   14715
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14775
      Begin VB.ListBox List 
         Height          =   2985
         Index           =   4
         Left            =   11040
         TabIndex        =   7
         Top             =   0
         Width           =   2775
      End
      Begin VB.ListBox List 
         Height          =   2985
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2775
      End
      Begin VB.ListBox List 
         Height          =   2985
         Index           =   3
         Left            =   8280
         TabIndex        =   5
         Top             =   0
         Width           =   2775
      End
      Begin VB.ListBox List 
         Height          =   2985
         Index           =   2
         Left            =   5520
         TabIndex        =   4
         Top             =   0
         Width           =   2775
      End
      Begin VB.ListBox List 
         Height          =   2985
         Index           =   1
         Left            =   2760
         TabIndex        =   3
         Top             =   0
         Width           =   2775
      End
      Begin VB.ListBox List 
         Height          =   2985
         Index           =   5
         Left            =   13800
         TabIndex        =   2
         Top             =   0
         Width           =   2775
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   1215
         Left            =   0
         TabIndex        =   1
         Top             =   3120
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   2143
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483638
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "apot_searching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub MOVELIST(Index As Integer)
    Dim i As Long
    Dim R As New ADODB.Recordset
 On Error GoTo MOVELIST
102 If Index <= 3 And List(Index).ListCount > 0 Then
104  R.Open "SELECT * FROM CATEGTEL where AA=" + str(Index + 2) + " AND PARENT=" + str(List(Index).ItemData(List(Index).ListIndex)) + " ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic
     'Dim roulis(5000, 4)
     'Dim I As Long
106  i = 0
108  List(Index + 1).Clear
110  If (Index + 2) <= 3 Then List(Index + 2).Clear
112  If (Index + 3) <= 3 Then List(Index + 3).Clear
 
114  Do While Not R.EOF
116       List(Index + 1).AddItem R!Name, i
118       List(Index + 1).ItemData(i) = R!ID
120       R.MoveNext
     Loop
122 R.Close
   
124 If List(Index).ListCount > 0 Then
126     ' list(INDEX + 1).SetFocus
128     ' list(INDEX + 1).ListIndex = 0
    End If








    End If

    Exit Sub '
MOVELIST:
130 MsgBox "MOVELIST" + str(Erl)
Resume Next
End Sub


Private Sub list_Click(Index As Integer)
   MOVELIST Index
   
   
   
'124    If list(INDEX + 1).ListCount > 0 Then
'126      list(INDEX + 1).SetFocus
'128      list(INDEX + 1).ListIndex = 0
'    End If
End Sub
