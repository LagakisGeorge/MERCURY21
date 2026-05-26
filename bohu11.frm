VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form bohu11 
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   11385
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "bohu11.frx":0000
      Height          =   2415
      Left            =   480
      OleObjectBlob   =   "bohu11.frx":0014
      TabIndex        =   5
      Top             =   5280
      Width           =   9495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "bohu11.frx":09E7
      Left            =   3120
      List            =   "bohu11.frx":09F7
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ейтупысг"
      Height          =   975
      Left            =   7560
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   4320
      Width           =   4695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "bohu11.frx":0A0F
      Height          =   3855
      Left            =   480
      OleObjectBlob   =   "bohu11.frx":0A23
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
   Begin VB.Label Label1 
      Caption         =   "table"
      Height          =   255
      Left            =   8280
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "bohu11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
  Dim dum
  ' CommonDialog1.ShowPrinter
  
  
  dum = Data1.Recordset("aa")
 ' g_ektyp dum
 
End Sub

Private Sub Command2_Click()
   Data2.DatabaseName = gDir
   Data2.Connect = gConnect
   Data2.RecordSource = "FIELDS"
   Data2.Refresh
   

End Sub

Private Sub Form_Load()
 Me.Picture = LoadPicture(gPicture)
   Data1.DatabaseName = gDir
   Data1.Connect = gConnect
   Data1.RecordSource = "VBREPEID"
End Sub
