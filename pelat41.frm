VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PELAT41 
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   10050
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "еНОДОР"
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   5040
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "pelat41.frx":0000
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "pelat41.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   9855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "PELAT41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Me.Hide
  ' Unload Me
End Sub

Private Sub Form_Activate()
Dim d1 As Date, d2 As Date

On Error Resume Next


Data1.DatabaseName = gDir


d1 = Label2.Caption 'pelat4.egg.Recordset("hme")


d2 = DateAdd("d", 1, d1)

Data1.RecordSource = "select atim,format(hme,'dd/mm/yy') as shme,kode,onoma,poso,timm,timm*poso as ajia from eggtim where atim='" + Label1.Caption + "' and hme>=#" + Format(d1, "mm/dd/yyyy") + "# and hme<=#" + Format(d2, "mm/dd/yyyy") + "#;"
Data1.Connect = gConnect
Data1.Refresh



DBGrid1.Columns(0).Caption = "пая/йо"
DBGrid1.Columns(0).Width = 800

DBGrid1.Columns(1).Width = 800
DBGrid1.Columns(1).DataField = "shme"
DBGrid1.Columns(1).Caption = "глея/миа"

DBGrid1.Columns(2).Caption = "йыд.еидоус"


DBGrid1.Columns(3).Width = 2600
DBGrid1.Columns(3).Caption = "пеяиц.еидоус"

DBGrid1.Columns(4).Width = 900
DBGrid1.Columns(4).Caption = "посотгта"


DBGrid1.Columns(5).Width = 800
DBGrid1.Columns(5).Caption = "тил.лом"


DBGrid1.Columns(6).Width = 800
DBGrid1.Columns(6).Caption = "аниа"

Data1.Refresh



End Sub

