VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form par5 
   BackColor       =   &H00FF0000&
   Caption         =   "тАЛЕъО гЛщЯАР"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   9975
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "еЙТЩПЫСГ"
      Height          =   675
      Left            =   2400
      TabIndex        =   3
      Top             =   8760
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "уПОКОЦИСЛЭР"
      Height          =   675
      Left            =   240
      TabIndex        =   2
      Top             =   8760
      Width           =   1605
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid g1 
      Height          =   7920
      Left            =   120
      TabIndex        =   1
      Top             =   615
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   13970
      _Version        =   393216
      Rows            =   40
      FixedRows       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton m_Exit 
      Caption         =   "еНОДОР"
      Height          =   675
      Left            =   6840
      TabIndex        =   0
      Top             =   8760
      Width           =   1605
   End
End
Attribute VB_Name = "par5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
Dim F_MONOENERGOS As Integer   '= Val(FINDPARAMETROI(1, "PAR1", "F_MONOENERGOS", "0", "деивмеи ломо тоус емеяцоус=1 окоус=0")) '

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
190         Printer.Print g1.TextMatrix(k, 0);
200         Printer.Print Tab(30); Right(Space(15) + Format(g1.TextMatrix(k, 1), "#####0.00"), 14)
        Next

210     Printer.EndDoc

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par5.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par5.Command2_Click " & "at line " & Erl

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
         "in ADOMERCNEW.par5.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par5.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

        Dim x

        Dim N As Integer

        Dim Oceos(100, 2), nOceos As Integer

        Dim Daneia(100, 2), mDaneia As Integer

100     mForm_Load Me, fh, fw, ft, fl

110     par5.SHOW
120     par5.MousePointer = vbHourglass

130     R.Open "SELECT * FROM PINAKES WHERE TYPOS=16 ORDER BY AYJON", Gdb, adOpenForwardOnly, adLockReadOnly
 F_MONOENERGOS = Val(FINDPARAMETROI(1, "PAR1", "F_MONOENERGOS", "0", "деивмеи ломо тоус емеяцоус=1 окоус=0")) '
        ' For K = 1 To 100
        Dim k

140     k = 0

150     Do While Not R.EOF
160         k = k + 1
170         Oceos(k, 1) = R("PERIGRAFH")
180         Oceos(k, 2) = R("TIMH")
190         R.MoveNext
        Loop

200     nOceos = k
210     R.Close

        Dim Tamieyt(100, 2), nTam As Integer

220     R.Open "SELECT * FROM PINAKES WHERE TYPOS=17 ORDER BY AYJON", Gdb, adOpenForwardOnly, adLockReadOnly

        ' For K = 1 To 100
        'Dim K

230     k = 0

240     Do While Not R.EOF
250         k = k + 1
260         Tamieyt(k, 1) = R("PERIGRAFH")
270         Tamieyt(k, 2) = R("TIMH")
280         R.MoveNext
        Loop

290     nTam = k
300     R.Close

310     R.Open "SELECT * FROM PINAKES WHERE TYPOS=18 ORDER BY AYJON", Gdb, adOpenForwardOnly, adLockReadOnly
320     k = 0

330     Do While Not R.EOF
340         k = k + 1
350         Daneia(k, 1) = R("PERIGRAFH")
360         Daneia(k, 2) = R("TIMH")
370         R.MoveNext
        Loop

380     mDaneia = k
390     R.Close

400     Me.KeyPreview = True
410     Me.Picture = LoadPicture(gPicture)
420     g1.ColWidth(0) = 2000
430     g1.ColWidth(1) = 2000
440     g1.row = 0
450     g1.TextMatrix(0, 0) = "есода"
460     g1.CellBackColor = vbGreen
470     g1.TextMatrix(1, 0) = "талеио "

480     N = 2

490     For k = 1 To nOceos
500         N = N + 1
510         g1.TextMatrix(N, 0) = Oceos(k, 1)
520         g1.TextMatrix(N, 1) = Format(Ypoloipo_Pel("e", Oceos(k, 2), F_MONOENERGOS), "###,##0.00")
        Next

530     N = N + 1

540     For k = 1 To nTam
550         N = N + 1
560         g1.TextMatrix(N, 0) = Tamieyt(k, 1)
570         g1.TextMatrix(N, 1) = Format(Ypoloipo_Pel("e", Tamieyt(k, 2), F_MONOENERGOS), "###,##0.00")
        Next

580     N = N + 1
590     N = N + 1

600     g1.TextMatrix(N, 0) = "епитацес еиспяайтеес"

610     R.Open "SELECT SUM(POS) FROM GRA WHERE ARX=1", Gdb, adOpenForwardOnly, adLockReadOnly

620     If R.EOF Then
630         x = 0
        Else

640         If IsNull(R(0)) Then
650             x = 0
            Else
660             x = R(0)
            End If
        End If

670     R.Close
680     g1.TextMatrix(N, 1) = x
        '

690     N = N + 1
700     g1.TextMatrix(N, 0) = "упокоипа пекатым"
710     g1.TextMatrix(N, 1) = Format(ypoloipa_pel("e", gEnarjh, Now, F_MONOENERGOS), "###,##0.00")

        '
720     N = N + 2
730     g1.TextMatrix(N, 0) = "аккгковяеос "

740     For k = 1 To mDaneia
750         N = N + 1
760         g1.TextMatrix(N, 0) = Daneia(k, 1)
770         g1.TextMatrix(N, 1) = Format(Ypoloipo_Pel("e", Format(Daneia(k, 2), "0000"), F_MONOENERGOS), "###,##0.00")
        Next

780     N = N + 2

790     g1.TextMatrix(N, 0) = "епитацес пкгяытеес"
800     R.Open "SELECT SUM(POS) FROM GRA WHERE ARX=2", Gdb, adOpenForwardOnly, adLockReadOnly

810     If R.EOF Then
820         x = 0
        Else

830         If IsNull(R(0)) Then
840             x = 0
            Else
850             x = R(0)
            End If
        End If

860     R.Close
870     g1.TextMatrix(N, 1) = Format(x, "###,##0.00")

880     N = N + 1
890     g1.TextMatrix(N, 0) = "упокоипа пяолгхеутым"
900     g1.TextMatrix(N, 1) = -ypoloipa_pel("r", gEnarjh, Now, F_MONOENERGOS)

910     N = N + 1
920     N = N + 1
930     g1.TextMatrix(N, 0) = "апотекесла "

        '
        Dim SUMA

940     SUMA = 0

950     g1.TextMatrix(1, 1) = Format(Ypoloipo_Pel("e", "9990", F_MONOENERGOS), "###,###,##0.00")  '"талеио "

960     For k = 1 To N
970         SUMA = SUMA + Val(Replace(g1.TextMatrix(k, 1), ",", ""))
        Next

980     g1.TextMatrix(N, 1) = Format(SUMA, "###,##0.00")

        '

        'g1.TextMatrix(3, 0) = "оьеыс ецматиа"
        'g1.TextMatrix(4, 0) = "оьеыс ехмийг"
        'g1.TextMatrix(5, 0) = "оьеыс аттийгс"
        ''
        'g1.TextMatrix(7, 0) = "талиеутгяио ецматиа"
        'g1.TextMatrix(8, 0) = "талиеутгяио ехмийг"
        'g1.TextMatrix(9, 0) = "талиеутгяио акжа"
        'g1.TextMatrix(10, 0) = "талиеутгяио пеияаиыс"
        ''
        'g1.TextMatrix(12, 0) = "епитацес еиспяайтеес"
        ''
        'g1.TextMatrix(13, 0) = "упокоипа пекатым"
        ''
        'g1.TextMatrix(16, 0) = "аккгковяеос "
        'g1.TextMatrix(17, 0) = "епитацес пкгяытеес"
        ''
        'g1.TextMatrix(18, 0) = "упокоипа пяолгхеутым"
        'g1.TextMatrix(20, 0) = "апотекесла "

990     Me.MousePointer = vbNormal

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par5.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par5.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 10215, 10725, fh, fw, ft, fl
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
         "in ADOMERCNEW.par5.m_Exit_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par5.m_Exit_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
