VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form bohu5 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   9480
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "bohu5.frx":0000
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8916
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÐñïâïëÞ"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "bohu5.frx":0014
      Height          =   5295
      Left            =   240
      OleObjectBlob   =   "bohu5.frx":0028
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   60817409
      CurrentDate     =   38348
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Width           =   7095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   60817409
      CurrentDate     =   38348
   End
End
Attribute VB_Name = "bohu5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 Dim d As String
  Me.MousePointer = vbHourglass
  d = "select hme AS [ÇÌÅÑ/ÍÉÁ],kod AS [ÊÙÄÉÊÏÓ],apa AS [ÁÑ_ÐÁÑ],"
  d = d + "ait AS [ÁÉÔÉÏËÏÃÉÁ],"
  d = d + "iif(xpi='×',format(XRE,'###,##0.00'),'') AS [×ÑÅÙÓÇ],"
  d = d + "iif(xpi='Ð',format(XRE,'###,##0.00'),'') AS [ÐÉÓÔÙÓÇ] from  temp "
  d = d + " where hme>=#" + Format(DTPicker1.Value, "mm/dd/yyyy") + "# and hme<#" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "#  order by hme,apa;"
  
On Error Resume Next
MSFlexGrid1.ColWidth(3) = 500
DBGrid1.Columns(4).NumberFormat = "000000.00"
Data1.RecordSource = d
Data1.Refresh
 MSFlexGrid1.ColAlignment(6) = 7
 MSFlexGrid1.ColAlignment(5) = 7
MSFlexGrid1.ColWidth(3) = 1500
Dim k
Do While k < MSFlexGrid1.Rows - 1
  k = k + 1
If MSFlexGrid1.TextMatrix(k, 3) = MSFlexGrid1.TextMatrix(k - 1, 3) Then
   Else
        MSFlexGrid1.AddItem "    ", k
        k = k + 1
   End If
Loop
 
 Me.MousePointer = vbNormal

End Sub

Private Sub Form_Load()
 Me.Picture = LoadPicture(gPicture)
Data1.DatabaseName = gdirlog
Data1.Connect = gConnect

End Sub
