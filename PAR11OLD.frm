VERSION 5.00
Begin VB.Form PAR11 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox label11 
      Height          =   375
      Left            =   1545
      TabIndex        =   19
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Áêýñùóç Êïóôïëüãçóçò"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Åêôýðùóç"
      Height          =   285
      Left            =   6600
      TabIndex        =   13
      Top             =   0
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   4560
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   2970
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   4560
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2955
      TabIndex        =   0
      Top             =   1185
      Width           =   1215
   End
   Begin VB.Label TimXonMeFpa 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label POSO 
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   6345
      TabIndex        =   16
      Top             =   105
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   510
      TabIndex        =   15
      Top             =   2370
      Width           =   900
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Åîïäá  %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   540
      TabIndex        =   14
      Top             =   2055
      Width           =   900
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ôéì.Xoíäñ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ðïó.ÊÝñä%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   7
      Left            =   3030
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ÍÝá Ôéì.Áã"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ðïó.ÊÝñä%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   2955
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ôéì.Ëéáí.ÌÖ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ðáë.Ôéì.Áã"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "PAR11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_XONDRIKH As Integer    ' ôï üíïìá ôïõ áñ÷åßïõ ðïõ ãßíåôáé ç åêôýðùóç ð.÷. c:\print
Dim F_ETTIKETES    ' 1=ÁÕÔÏÌÁÔÇ ÓÅ ÅÊÔÕÐÙÓÇ ÓÔÇÍ ÊÏÓÔÏËÏÃÇÓÇ


'

Sub LAG_ETIK()
        '<EhHeader>
        On Error GoTo LAG_ETIK_Err
        '</EhHeader>

        Dim k, mtimes, MTIMM, timologio

        Dim rec As New ADODB.Recordset

        Dim MPER
100     MPER = PAR1.EID.Recordset("ONO")





110     mtimes = InputBox("Ðüóåò åôéêÝôåò;" + PAR1.EID.Recordset("ONO"), , POSO.Caption)
120     TIMM = Format(Val(Text1(4).Text), "##0.00")
130     timologio = PAR1.EID.Recordset("KOD") + Left(PAR1.Text1.Text, 6) + "=" + PAR1.PEL.Recordset("KOD")
140     timologio = timologio + Format(Val(Label11.Text) * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(PAR1.DTPicker1, "dd/mm/yyyy")
        Dim mlabel

150     If IsNull(PAR1.EID.Recordset("ERG")) Then
160         GoSub FIND_N_NUMBER
        Else
170         If Val(PAR1.EID.Recordset("ERG")) = 0 Then
180             GoSub FIND_N_NUMBER
            Else
190             mlabel = PAR1.EID.Recordset("erg")
            End If

        End If







200     For k = 1 To Val(mtimes)

210         Printer.FontName = "128"
220         Printer.FONTSIZE = 5
230         Printer.CurrentY = 0
240         Printer.Print Tab(13); mlabel
250         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
260         Printer.FONTSIZE = 7
270         Printer.CurrentY = 110
280         Printer.Print

290         Printer.Print Tab(ar_per + 8); mlabel
300         Printer.Print Tab(ar_per + 8); MPER
310         Printer.Print Tab(ar_per + 8); timologio
320         Printer.FONTSIZE = 16
330         Printer.FontBold = True
340         Printer.Print Tab(ar_per + 8); Trim(LTrim(TIMM)) + " €"

350         Printer.EndDoc
        Next    '00005589





        '-----------------  MEGALH ETIKETA ------------------------
        'For k = 1 To Val(mtimes)
        '
        '
        '
        '  Printer.FontName = "128"
        '  Printer.FontSize = 6
        '  Printer.Print Tab(12); mlabel
        '
        '  Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
        '  Printer.FontSize = 8
        '  Printer.Print
        '  Printer.Print
        '  Printer.Print
        '  Printer.Print
        '
        '  Printer.Print Tab(10); mlabel
        '  Printer.Print Tab(10); MPER
        '  Printer.Print Tab(10); timologio
        '  Printer.FontSize = 16
        '  Printer.FontBold = True
        '  Printer.Print Tab(10); Trim(LTrim(TIMM)) + " €"
        '
        '  Printer.EndDoc
        '
        'Next
        '
        Exit Sub


FIND_N_NUMBER:
        '************************

        On Error GoTo LAG_ETIK_Err
360     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

370     rec.MoveFirst
380     rec.Move 10
        'rec.Edit
390     If IsNull(rec("epel")) Then
400         rec("epel") = 1
        Else
410         rec("epel") = rec("epel") + 1
        End If
420     rec.Update
430     rec.Close


440     mlabel = Right("00000000" + LTrim(Str(rec("EPEL"))), 8)
450     Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + PAR1.EID.Recordset("KOD") + "'"
460     Return








        '<EhFooter>
        Exit Sub

LAG_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.LAG_ETIK " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.LAG_ETIK " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub MAKIS_ETIK()
        '<EhHeader>
        On Error GoTo MAKIS_ETIK_Err
        '</EhHeader>

        Dim k, timologio

        Dim rec As New ADODB.Recordset



        't = "c:\lageuro\makisget.exe "
        't = t + par1.EID.Recordset("KOD") + " "
        't = t + Replace(par1.EID.Recordset("ONO"), " ", "_") + " "
        't = t + Format(Val(Text1(4).Text), "##0.00") + " "   ' LTI5  LIANIKI
        't = t + Format(Val(Text1(2).Text), "##0") + " "  ' POS_KERD
        'If Len(Trim(mERG)) > 0 Then
        '   t = t + mERG
        'Else
        '   t = t + "00"
        'End If





100     MPER = PAR1.EID.Recordset("ONO")



        Dim mtimes, TIMM, pOS_KERD


110     mtimes = 1
120     TIMM = Format(Val(Text1(4).Text), "##0.00")
130     MPER = PAR1.EID.Recordset("ONO")
140     pOS_KERD = Format(Val(Text1(2).Text), "##0") + " -"    ' POS_KERD

        '   timologio = par1.EID.Recordset("KOD") + Left(par1.Text1.Text, 6) + "=" + par1.PEL.Recordset("KOD")
        '  timologio = timologio + Format(Val(label11.text) * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(par1.DTPicker1, "dd/mm/yyyy")
        Dim mlabel

150     Printer.FontBold = True

160     Printer.CurrentY = 30

170     Printer.FONTSIZE = 10
180     Printer.Print Tab(2); MPER

190     Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
200     Printer.FONTSIZE = 26

210     Printer.CurrentY = 600
220     Printer.Print "      " + TIMM + "€"
230     Printer.FONTSIZE = 10
240     Printer.CurrentY = 1450
250     Printer.Print "  " + PAR1.EID.Recordset("ERG") + "-" + Format(500 + Val(Text1(2).Text), "####0") + " -" + PAR1.EID.Recordset("KOD")





260     Printer.EndDoc





        Exit Sub


FIND_N_NUMBER:
        '************************

        On Error GoTo MAKIS_ETIK_Err
270     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

280     rec.MoveFirst
290     rec.Move 10
        'rec.Edit
300     If IsNull(rec("epel")) Then
310         rec("epel") = 1
        Else
320         rec("epel") = rec("epel") + 1
        End If
330     rec.Update
340     rec.Close


350     mlabel = Right("00000000" + LTrim(Str(rec("EPEL"))), 8)
360     Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + PAR1.EID.Recordset("KOD") + "'"
370     Return








        '<EhFooter>
        Exit Sub

MAKIS_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.MAKIS_ETIK " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.MAKIS_ETIK " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub





Private Sub Command1_Click()
        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>
        Dim DUM
        Dim mERG


        Dim X As Printer

100     For Each X In Printers
110         If InStr(X.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
120             Set Printer = X
                ' Stop looking for a printer.
                Exit For
            End If
        Next



130     If Len(Dir("C:\LAGEURO\LAGAKIS.TXT", vbNormal)) > 0 Then
140         LAG_ETIK
150         Text1(4).SetFocus

            Exit Sub
        End If





160     If IsNull(PAR1.EID.Recordset("ERG")) Then
170         mERG = "--"
        Else
180         mERG = PAR1.EID.Recordset("ERG")
        End If


190     If IsNull(PAR1.EID.Recordset("KOD")) Then
200         Me.Caption = "NULL KOD"
        End If



210     If IsNull(Replace(PAR1.EID.Recordset("ONO"), " ", "_")) Then
220         Me.Caption = "NULL ONO"
            Exit Sub
        End If



230     If IsNull(PAR1.EID.Recordset("LTI5")) Then
240         Me.Caption = "NULL LTI5"
            Exit Sub
        End If

250     If IsNull(PAR1.EID.Recordset("POS_KERD")) Then
260         Me.Caption = "NULL POS_KERD"
            Exit Sub
        End If

270     If IsNull(PAR1.EID.Recordset("ERG")) Then
280         Me.Caption = "NULL ERG"
            Exit Sub
        Else
290         If Len(PAR1.EID.Recordset("ERG")) = 0 Then
300             mERG = "--"
            End If
        End If


        Dim t
310     t = "c:\lageuro\makisget.exe "
320     t = t + PAR1.EID.Recordset("KOD") + " "
330     t = t + Replace(PAR1.EID.Recordset("ONO"), " ", "_") + " "
340     t = t + Format(Val(Text1(4).Text), "##0.00") + " "   ' LTI5  LIANIKI
350     t = t + Format(Val(Text1(2).Text), "##0") + " "  ' POS_KERD
360     If Len(Trim(mERG)) > 0 Then
370         t = t + mERG
        Else
380         t = t + "00"
        End If


390     Me.Caption = t


400     If Len(Dir("C:\LAGEURO\MAKISGET.EXE", vbNormal)) > 0 Then
            ' DUM = Shell(t)
410         MAKIS_ETIK
        Else
420         PRINT_ETIK
        End If


        '   DUM = Shell("c:\lageuro\makisget.exe " + par1.EID.Recordset("KOD") + " " + Replace(par1.EID.Recordset("ONO"), " ", "_") + " " + Format(par1.EID.Recordset("LTI5"), "##0.00") + " " + Format(par1.EID.Recordset("POS_KERD"), "##0") + " " + mERG)


        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.Command1_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.Command1_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Command2_Click()
        '<EhHeader>
        On Error GoTo Command2_Click_Err
        '</EhHeader>
100     PAR1.kostologhsh.Value = False
110     Unload Me

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.Command2_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.Command2_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_DblClick()
        '<EhHeader>
        On Error GoTo Form_DblClick_Err
        '</EhHeader>
100     PARAMETROI.PARAM.Caption = "PAR11"

110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.Form_DblClick " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.Form_DblClick " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     F_XONDRIKH = Val(FINDPARAMETROI("PAR11", "F_XONDRIKH", "1", "Å÷åé ÷ïíäñéêÞ ôéìÞ=1/Ï÷é=0"))    '
110     F_ETTIKETES = Val(FINDPARAMETROI("PAR11", "F_ETTIKETES", "0", "AYTOMATA ÓE EKTYÐÙÓÇ=1/Ï÷é=0"))    '
120     If F_XONDRIKH = 0 Then
130         Text1(6).Visible = False
140         Text1(7).Visible = False
150         Label1(7).Visible = False
160         Label1(8).Visible = False
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.Form_Load " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Text1_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo Text1_Change_Err
        '</EhHeader>
        Dim S

100     If Index = 7 Then
110         S = 0
        End If

        '<EhFooter>
        Exit Sub

Text1_Change_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.Text1_Change " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.Text1_Change " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Text1_GotFocus(Index As Integer)
        '<EhHeader>
        On Error GoTo Text1_GotFocus_Err
        '</EhHeader>
        On Error Resume Next

        Dim exoda As Single    ' exoda
        Dim TimLianMefpa As Single, TimXon As Single
        Dim SyntFpa As Single
        ' p.x. 1.19  gia 19%
100     SyntFpa = (1 + g_Fpa(PAR1.EID.Recordset("fpa")) / 100)


110     Text1(Index).BackColor = vbYellow
120     Text1(Index).SelStart = 0
130     Text1(Index).SelLength = Len(Text1(Index).Text)
        'label4  % ÅÎÏÄÁ
140     Label4.Caption = metaf.Text / 3
150     Label4.Caption = Replace(Label4.Caption, ",", ".")

160     exoda = Val(Label4.Caption)

170     Text1(4).Text = Replace(Text1(4).Text, ",", ".")
180     TimLianMefpa = Val(Text1(4).Text)

190     Label11.Text = Replace(Label11.Text, ",", ".")
200     TimXon = Val(Label11.Text)




        'LABEL10  ' PALIA TIMH AGORAS
        'LABEL11  ' NEA TIMH AGORAS


        ' ÔÅ×Ô1(4).ÔÅ×Ô=TIMH     LIANIKHS ME FPA
        ' ÔÅ×Ô1(2).ÔÅ×Ô=POSOSTO  KERDOYS LIANIKHS

        ' ÔÅ×Ô1(7).ÔÅ×Ô=TIMH     XONDRIKHS ME FPA
        ' ÔÅ×Ô1(6).ÔÅ×Ô=POSOSTO  KERDOYS XONDRIKHS
        'TimXonMeFpa

        'label4  % ÅÎÏÄÁ

        'ÐÑÏÔÅÉÍÏÌÅÍÇ ËÉÁÍÉÊÇ    LABEL2.CAPTION

        'ÌÏËÉÓ ÐÁÅÉ ÓÔÏ ÌÁÑÊÁÐ ×ÏÍÄÑÉÊÇÓ ÔÏÔÅ
210     If Index = 6 Then    ' íá õðïëïãßæåé ôï ðñáãìáôéêü ìáñêáð ëéáíéêÞò
220         If Val(TimXon) > 0 Then
230             Text1(2).Text = Format(-100 + 100 * TimLianMefpa / SyntFpa / Val(TimXon), "###0.00")
240             Text1(2).Text = Replace(Text1(2).Text, ",", ".")

                ' âáæù ôï éäéï ìáñêáð êáé óôçí ÷ïíäñéêç
250             If Val(Text1(6).Text) = 0 Then
                    'OXI DIONYSHS Text1(6).Text = Text1(2).Text
                End If


            End If
        End If

        'proteinomeni lianikh me fpa
260     Label2.Caption = Format((100 + exoda) / 100 * (100 + Val(Text1(2).Text)) / 100 * TimXon * SyntFpa, "#####.00")
        'Label2.Caption = Format((100 + exoda) / 100 * (100 + TimLianMefpa) / 100 * TimXon * SyntFpa, "#####.00")

        ' TimXonMeFpa = Format((100 + Val(Replace(Label4.Caption, ",", "."))) / 100 * (100 + Val(Replace(Text1(7).Text, ",", "."))) / 100 * Val(Replace(TimXon, ",", ".")) * (100 + g_Fpa(PAR1.EID.Recordset("fpa"))) / 100, "#####.00")


        ' ÔÅ×Ô1(7).ÔÅ×Ô=TIMH LIANIKHS ME FPA

        ' Text1(7).Text

270     TimXonMeFpa = (100 + Val(Replace(Label4.Caption, ",", "."))) / 100 * (100 + Val(Replace(Text1(6).Text, ",", "."))) / 100 * Val(Replace(TimXon, ",", "."))

        ' Text1(7).Text = (100 + Val(Replace(Label4.Caption, ",", "."))) / 100 * (100 + Val(Replace(Text1(6).Text, ",", "."))) / 100 * Val(Replace(TimXon, ",", "."))



        '<EhFooter>
        Exit Sub

Text1_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.Text1_GotFocus " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.Text1_GotFocus " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    '------------------------------------
    'LABEL11  ' NEA TIMH AGORAS
    ' ÔÅ×Ô1(4).ÔÅ×Ô=TIMH     LIANIKHS ME FPA
    ' ÔÅ×Ô1(2).ÔÅ×Ô=POSOSTO  KERDOYS LIANIKHS
    ' ÔÅ×Ô1(7).ÔÅ×Ô=TIMH XONDRIKHS
    ' ÔÅ×Ô1(6).ÔÅ×Ô=POSOSTO  KERDOYS XONDRIKHS
    'TimXonMeFpa
    'label4  % ÅÎÏÄÁ
        '<EhHeader>
        On Error GoTo Text1_KeyPress_Err
        '</EhHeader>

        Dim S
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         If Index = 7 Then    ' terma
130             Text1(2).SetFocus
140             Me.Hide
150         ElseIf Index = 4 And F_XONDRIKH = 0 Then    ' terma
160             Text1(2).SetFocus
170             Me.Hide
            Else
180             keybd_event VK_TAB, 0, 0, 0    'send a tab
            End If
        Else
190         If Index = 7 Then
200             S = 0
                '' ÔÅ×Ô1(6).ÔÅ×Ô=POSOSTO  KERDOYS XONDRIKHS
                'LABEL11  ' NEA TIMH AGORAS
                ' ' ÔÅ×Ô1(7).ÔÅ×Ô=TIMH XONDRIKHS
                'POSOSTO  KERDOYS XONDRIKHS       (TIMH XONDRIKHS )    (NEA TIMH AGORAS)
210             Text1(6).Text = Format(-100 + 100 * Val(Text1(7).Text) / Val(Label11.Text), "###0.00")
            End If
        End If
        '<EhFooter>
        Exit Sub

Text1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.Text1_KeyPress " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.Text1_KeyPress " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Private Sub Text1_LostFocus(Index As Integer)
        '<EhHeader>
        On Error GoTo Text1_LostFocus_Err
        '</EhHeader>
100     Text1(Index).BackColor = vbWhite
110     Text1(Index).Text = Replace(Text1(Index).Text, ",", ".")
        '<EhFooter>
        Exit Sub

Text1_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.Text1_LostFocus " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.Text1_LostFocus " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub PRINT_ETIK()
        '<EhHeader>
        On Error GoTo PRINT_ETIK_Err
        '</EhHeader>
        Dim ll, apoSCANNER As Long
        Dim db
        Dim PSIFIA As Integer
        Dim fores As Long, k As Integer
        Dim X As Printer
        Dim barc(200)
        Dim mONOMA, mBarcode, mERG, mLTI5
        Dim R As New ADODB.Recordset
        Dim RD As DAO.Recordset
100     mONOMA = PAR1.EID.Recordset("ONO")    ' Text1(1).Text 'ONOMA
110     mBarcode = PAR1.EID.Recordset("ERG")    'Text1(0).Text '  ' BARCODE
120     mERG = PAR1.EID.Recordset("KOD")    ' Text1(6).Text ' erg
130     mLTI5 = Format(Val(Replace(Text1(4).Text, ",", ".")), "###0.00")    ' LTI5

        On Error Resume Next
140     For Each X In Printers
150         If InStr(X.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
160             Set Printer = X
                ' Stop looking for a printer  5200016000239
                Exit For
            End If
        Next
170     apoSCANNER = 1
180     fores = 1  '   Val(InputBox("Ðüóåò åôôéêÝôåò; ", "", 1))
        Dim mFORES As Long
        Dim Z As Integer

190     For ll = 1 To apoSCANNER
200         mFORES = fores
210         Z = 10    ' InputBox("METATOPISH")
220         For k = 1 To mFORES
230             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992
240             Printer.FONTSIZE = 8
250             Printer.FontBold = True
260             Printer.Print Tab(4 + Z); mONOMA    'Text1(1).Text 'ONOMA
270             Printer.FONTSIZE = 8
280             Printer.Print Tab(4 + Z); mBarcode;    '  Text1(0).Text;  ' BARCODE
290             If Len(Trim(Text1(0).Text)) = 12 Then
300                 Printer.FontName = "UPCA"    ' "128"
                Else
310                 Printer.FontName = "EAN13SMALL"    ' "128"
                End If
320             If Len(Dir("C:\LAGEURO\NOPONTOI.TXT", vbNormal)) > 0 Then
                Else
330                 Printer.FONTSIZE = 5
340                 Printer.Print Tab(15 + Z); mBarcode    'Text1(0).Text 'eggtim.Recordset("kode") ' mlabel
                End If
350             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
360             Printer.FontBold = True
370             Printer.FONTSIZE = 8
380             Printer.Print Tab(4 + Z); mERG    ' Text1(6).Text ' erg
390             Printer.FONTSIZE = 24
400             If Len(mLTI5) = 6 Then    ' 208.35
410                 Printer.Print Tab(7 + Z - 7); " € " + mLTI5; ""
420             ElseIf Len(mLTI5) = 5 Then    '18.35
430                 Printer.Print Tab(8 + Z - 7); "€ " + mLTI5; ""
                Else
440                 Printer.Print Tab(10 + Z - 7); "€ " + mLTI5; ""
                End If
450             Printer.FONTSIZE = 6
460             Printer.Print
470             Printer.FONTSIZE = 8
480             Printer.Print
490             If Len(Dir("C:\LAGEURO\NOPONTOI.TXT", vbNormal)) > 0 Then
500                 Printer.Print Tab(Z + 3); "5-" + Text1(2).Text    'POSOSTO KERDOYS
                Else
510                 Printer.CurrentX = 200
520                 Printer.CurrentY = 500
530                 Printer.FONTSIZE = 14
540                 If Val(mLTI5) > 2.99 Then
550                     Printer.Print Tab(Z - 2); "ÐÏÍÔÏÉ"
560                     PSIFIA = Len(Format(Int(Val(mLTI5) / 3), "##0"))
570                     If PSIFIA = 3 Then Printer.Print ; Tab(5 + Z - 5); Format(Int(Val(mLTI5) / 3), "##0")
580                     If PSIFIA = 1 Then Printer.Print ; Tab(6 + Z - 4); Format(Int(Val(mLTI5) / 3), "##0")
590                     If PSIFIA = 2 Then Printer.Print ; Tab(5 + Z - 4); Format(Int(Val(mLTI5) / 3), "##0")
                    End If
                End If
600             Printer.EndDoc
610         Next k
620     Next ll
        '<EhFooter>
        Exit Sub

PRINT_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.PAR11.PRINT_ETIK " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PAR11.PRINT_ETIK " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

