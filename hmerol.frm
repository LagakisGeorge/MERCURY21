VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form hmerol 
   Caption         =   "Form2"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10596
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   10596
   Begin VB.CommandButton Command2 
      Caption         =   "Επόμενη"
      Height          =   480
      Left            =   4800
      TabIndex        =   5
      Top             =   6435
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Προηγούμενη"
      Height          =   480
      Left            =   525
      TabIndex        =   4
      Top             =   6435
      Width           =   1215
   End
   Begin VB.PictureBox Εξοδος 
      Height          =   480
      Left            =   6750
      ScaleHeight     =   432
      ScaleWidth      =   2160
      TabIndex        =   3
      Top             =   6435
      Width           =   2205
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   5790
      Picture         =   "hmerol.frx":0000
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   1800
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   3615
      _ExtentX        =   6371
      _ExtentY        =   656
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   103350272
      CurrentDate     =   38733
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   5385
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   765
      Width           =   8490
   End
End
Attribute VB_Name = "hmerol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    DTPicker1.value = DateAdd("d", -1, DTPicker1.value)
    UPDATE_TEXT

End Sub

Private Sub Command2_Click()

    DTPicker1.value = DateAdd("d", 1, DTPicker1.value)
    UPDATE_TEXT

End Sub

Private Sub DTPicker1_Change()
    UPDATE_TEXT

End Sub

Private Sub DTPicker1_Click()

    UPDATE_TEXT
End Sub

Private Sub DTPicker1_CloseUp()

    UPDATE_TEXT
End Sub

Private Sub DTPicker1_GotFocus()
    Gdb.Execute "UPDATE HMEROL SET SHM2='" + Text1.Text + "' WHERE HME='" + Format(DTPicker1.value, "MM/DD/YYYY") + "'"
End Sub

Private Sub Form_Load()
    hmerol.Picture = LoadPicture(gPicture)
    DTPicker1.value = Now
    'Adodc1.ConnectionString = gConnect
    'Adodc1.RecordSource = "select TOP 1 * FROM HMEROL"



    'On Error GoTo CREATE_ETAIR
    UPDATE_TEXT

    'Adodc1.Refresh
    On Error Resume Next


    ' Dim R As New ADODB.Recordset




    Exit Sub


CREATE_ETAIR:
    MsgBox Err.Description

    Gdb.Execute "CREATE TABLE HMEROL ( HME  DATETIME ,SHM2  NCHAR(120) )"
    Gdb.Execute "INSERT INTO HMEROL (HME,SHM2) VALUES ('01/01/2006','')"
    Resume Next


End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim G

    On Error Resume Next
    Gdb.Execute "UPDATE HMEROL SET SHM2='" + Text1.Text + "' WHERE HME='" + Format(DTPicker1.value, "MM/DD/YYYY") + "'", G
    If G = 0 Then
        Exit Sub
    End If




End Sub
Sub UPDATE_TEXT()
'  DHMIOYRGEI THN HMERA
    Dim R As New ADODB.Recordset
    dum2 = ADD_FIELD("HMEROL", "SHM2", "TEXT")

    On Error GoTo 0
    R.Open "SELECT * FROM HMEROL WITH (ROWLOCK) WHERE HME='" + Format(DTPicker1.value, "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockPessimistic

    If R.EOF Then
        Gdb.Execute "INSERT INTO HMEROL (HME,SHM2) VALUES ('" + Format(DTPicker1.value, "MM/DD/YYYY") + "',' ')"
    Else
        '    If IsNull(R("SHM2")) = True Then
        '          Text1.Text = " "
        '    Else
        On Error Resume Next

        Text1.Text = R("SHM2")
        '    End If

    End If

    'Adodc1.RecordSource = "SELECT * FROM HMEROL WHERE HME='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "'"
    'Adodc1.Refresh





End Sub

Private Sub MyButton1_Click()
    Unload Me

End Sub

Private Sub Εξοδος_Click()
    Unload Me


End Sub
