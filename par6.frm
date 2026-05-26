VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form par6 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   11250
   WindowState     =   2  'Maximized
   Begin VB.TextBox apot 
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Text            =   "2"
      Top             =   720
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   3360
      TabIndex        =   4
      Top             =   300
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      Format          =   225378305
      CurrentDate     =   39209
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Åêôýðùóç"
      Height          =   675
      Left            =   2340
      TabIndex        =   3
      Top             =   4785
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Õðïëïãéóìüò"
      Height          =   675
      Left            =   180
      TabIndex        =   2
      Top             =   4785
      Width           =   1605
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid g1 
      Height          =   3090
      Left            =   165
      TabIndex        =   1
      Top             =   1530
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   5450
      _Version        =   393216
      Rows            =   40
      Cols            =   3
      FixedRows       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton m_Exit 
      Caption         =   "Åîïäïò"
      Height          =   675
      Left            =   6780
      TabIndex        =   0
      Top             =   4785
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ÁðïèÞêç ÔáìåéáêÞò"
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   1920
      TabIndex        =   7
      Top             =   720
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Çìåñïìçíßá "
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   1950
      TabIndex        =   5
      Top             =   345
      Width           =   1320
   End
End
Attribute VB_Name = "par6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim f_flag As Integer

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     f_flag = 1
110     mCalc

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par6.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par6.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '====================================================================
        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     utDialog.Command1.Enabled = False
110     utDialog.Command3.Enabled = False
120     utDialog.Command4.Enabled = False

130     utDialog.SHOW 1

        ' MDIForm1.StatusBar1.Panels(6).Text = 1

140     Printer.FontSize = 9
150     Printer.FontName = "Courier New"
160     Printer.Font.Charset = 161

170     Printer.Print

        Dim k

180     For k = 1 To 30
190         Printer.Print G1.TextMatrix(k, 0);
200         Printer.Print Tab(30); Right(Space(15) + Format(G1.TextMatrix(k, 1), "###,##0.00"), 14)
        Next

210     Printer.EndDoc

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par6.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par6.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Activate()

        '<EhHeader>
        On Error GoTo Form_Activate_Err

        '</EhHeader>

100     DTPicker1.Value = Now

110     If Len(Dir("C:\KENTRIKO.TXT")) > 1 Then
120         apot.Text = "1"
130     ElseIf Len(Dir("C:\YPOK.TXT")) > 1 Then
140         apot.Text = "2"
150     ElseIf Len(Dir("C:\YPOK2.TXT")) > 1 Then
160         apot.Text = "3"
        End If

        '<EhFooter>
        Exit Sub

Form_Activate_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par6.Form_Activate " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par6.Form_Activate " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then m_Exit_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par6.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par6.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     f_flag = 0
110     mCalc

120     mForm_Load Me, fh, fw, ft, fl

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par6.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par6.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub mCalc()

        '<EhHeader>
        On Error GoTo mCalc_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

        Dim X

        Dim n As Integer

        Dim Oceos(100, 2), nOceos As Integer

        Dim Daneia(100, 2), mDaneia As Integer

100     par6.MousePointer = vbHourglass
110     R.Open "SELECT * FROM PINAKES WHERE TYPOS=12 ORDER BY AYJON", Gdb, adOpenForwardOnly, adLockReadOnly

        Dim k

120     k = 0

130     Do While Not R.EOF
140         k = k + 1
            '  Oceos(k, 1) = R("PERIGRAFH")
            '  Oceos(k, 2) = R("TIMH")
150         R.MoveNext
        Loop

160     nOceos = k
170     R.Close

180     Me.KeyPreview = True
190     Me.Picture = LoadPicture(gPicture)
200     G1.ColWidth(0) = 2000
210     G1.ColWidth(1) = 2000
        G1.ColWidth(2) = 12000
220     G1.row = 0
230     G1.TextMatrix(0, 0) = "ÅÓÏÄÁ"
240     G1.CellBackColor = vbGreen
250     G1.TextMatrix(1, 0) = "ËÏÃÁÑÉÁÓÌÏÓ ÔÁÌÅÉÏ "

        Dim m_now As Date

260     m_now = DTPicker1.Value

        On Error Resume Next

270     n = 2

280     For k = 1 To nOceos
290         n = n + 1
300         G1.TextMatrix(n, 0) = Oceos(k, 1)

310         G1.TextMatrix(n, 1) = Format(Ypoloipo_today("e", Oceos(k, 2), m_now), "#####0.00")
        Next

320     n = n + 1

        Dim Q      As String

        Dim SUMTAM As Single

330     Q = "select SUM(MIK_AJIA) " & " from EGGTIM WHERE LEFT(ATIM,1)='L'" & " and HME>='" + Format(m_now, "mm/dd/yyyy") + "'  " & " and HME<'" + Format(DateAdd("d", 1, m_now), "mm/dd/yyyy") + "' " & " AND APOT=" + apot.Text

340     R.Open Q, Gdb, adOpenDynamic, adLockOptimistic
350     G1.TextMatrix(n, 0) = "ÔÁÌÅÉÁÊÇ ÌÇ×"

360     SUMTAM = IIf(IsNull(R(0)), 0, R(0))
370     G1.TextMatrix(n, 1) = SUMTAM

380     R.Close

390     n = n + 1
        ' Dim Q As String
        ' Dim SUMTAM As Single

400     Q = "select SUM(MIK_AJIA) AS MIK " & " from EGGTIM WHERE LEFT(ATIM,1) in ('l','R')" & " and HME>='" + Format(m_now, "mm/dd/yyyy") + "'  " & " and HME<'" + Format(DateAdd("d", 1, m_now), "mm/dd/yyyy") + "' " & " AND APOT=" + apot.Text

410     R.Open Q, Gdb, adOpenDynamic, adLockOptimistic
420     G1.TextMatrix(n, 0) = "internet"

430     SUMTAM = IIf(IsNull(R(0)), 0, R(0))
440     G1.TextMatrix(n, 1) = SUMTAM

450     R.Close

460     n = n + 1

        Dim ayp As Single

        Dim sql As String

470     If f_flag = 1 Then

            Dim Y As Single

480         sql = "select sum(AJI) from TIM WHERE LEFT(ATIM,1)='T'  and HME>='" + Format(DTPicker1, "mm/dd/yyyy") + "'  and HME<'" + Format(DateAdd("d", 1, DTPicker1), "mm/dd/yyyy") + "'"
490         R.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

500         If R.EOF Then
510             Y = 0
            Else

520             If IsNull(R(0)) Then
530                 Y = 0
                Else
540                 Y = R(0)
                End If
            End If

550         R.Close
560         G1.TextMatrix(n, 0) = "TIMÏËÏÃÉÁ"
570         G1.TextMatrix(n, 1) = Y

        End If

        Dim f  As Integer

        Dim SU As Single

580     SU = 0

590     For f = 1 To n
600         SU = SU + Val(G1.TextMatrix(f, 1))
        Next

610     n = n + 1
620     G1.TextMatrix(n, 0) = "óõíïëá"

630     G1.TextMatrix(n, 1) = SU

640     Me.MousePointer = vbNormal

        '<EhFooter>
        Exit Sub

mCalc_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par6.mCalc " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par6.mCalc " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 11490, 7245, fh, fw, ft, fl
End Sub

Private Sub g1_Click()

        '<EhHeader>
        On Error GoTo g1_Click_Err

        '</EhHeader>

100     f_flag = 1

        '<EhFooter>
        Exit Sub

g1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par6.g1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par6.g1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub m_Exit_Click()

        '<EhHeader>
        On Error GoTo m_Exit_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

m_Exit_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par6.m_Exit_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par6.m_Exit_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
