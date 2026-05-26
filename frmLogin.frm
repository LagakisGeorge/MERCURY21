VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3015
   ClientLeft      =   12465
   ClientTop       =   6345
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1781.361
   ScaleMode       =   0  'User
   ScaleWidth      =   3873.152
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¡ ’—Ÿ”«"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1140
      Picture         =   "frmLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1740
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080C0FF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1140
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   930
      Width           =   2625
   End
   Begin MSComCtl2.DTPicker Workday 
      Height          =   285
      Left            =   1140
      TabIndex        =   6
      Top             =   1380
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   503
      _Version        =   393216
      Format          =   442957825
      CurrentDate     =   39828
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1170
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2655
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.TextBox txtUserName 
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1140
      TabIndex        =   1
      Top             =   135
      Width           =   2625
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FF80&
      Caption         =   "OK"
      Height          =   660
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1740
      Width           =   1425
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1140
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2625
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      ForeColor       =   &H8000000E&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      ForeColor       =   &H8000000E&
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim F_UPDATESTRUCTURES As Integer

Dim COUNTER           As Integer

Dim ETAIREIES(10)     As String

Public LoginSucceeded As Boolean

'Private Sub cmdCancel_Click()
''set the global var to false
''to denote a failed login
'    LoginSucceeded = False
'
'
'
'
'
'    gWorkDay = Workday.Value
'
'    End
'
'
'    Me.Hide
'End Sub

Private Sub cmdOK_Click()

        '<EhHeader>
        On Error GoTo cmdOK_Click_Err

        '</EhHeader>

        Dim R As New ADODB.Recordset

        ' If Len(Trim(Text1.Text)) > 0 Then
        '    gConnect = gConnect + Text1.Text
        ' End If

        On Error GoTo NOCONNECT

        '
        'Gdb.Open gConnect

        On Error GoTo CREATE_USERS

        'Dim R As New ADODB.Recordset
100     R.Open "SELECT TOP 1 * FROM USERS", Gdb, adOpenForwardOnly, adLockReadOnly
110     R.Close

120     R.Open "select * FROM USERS WHERE NAME ='" + Trim(txtUserName) + "' AND CODE ='" + Trim(txtPassword) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

130     If R.EOF Then
140         COUNTER = COUNTER + 1
150         MsgBox "À¡»œ”  Ÿƒ… œ”"
160         txtUserName = ""
170         txtPassword = ""
180         txtUserName.SetFocus

190         If COUNTER > 3 Then End

            Exit Sub

        End If

200     gXEIRISTHS = R("DIKAIOMA")    '  1  'ÂÒÈÔÒÈÛÏÂÌÔ ÏÂÌÔı
210     gUserId = R("ID_NUM")
220     R.Close

        Dim N As Integer

230     gPicture = "C:\MERCVB\BACKGROUND.BMP"

240     If Combo1.ListIndex = -1 Then
250         N = 0
        Else
260         N = Combo1.ListIndex

270         If N > 0 Then
280             gPicture = "C:\MERCVB\PALIA.BMP"
            End If
        End If

290     If N > 0 Then
300         gConnect = gConnect + ";DATABASE=" + ETAIREIES(N)
        End If

310     Gdb.Close

        On Error GoTo OUTSIDE

320     Gdb.ConnectionTimeout = 50000

330     Gdb.Open gConnect

PARAMETROI_TO_ARRAY

        '           332     r.Open "select * FROM  WHERE NAME ='" + Trim(txtUserName) + "' AND CODE ='" + Trim(txtPassword) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

        '290   '

        On Error Resume Next

Dim F_VER As String ' VERSION DATABASE
Dim P_VER As String ' VERSION PROGRAMMATOS
P_VER = Format(App.Major, "00") + "." + Format(App.Minor, "00") + "." + Format(App.Revision, "00")


F_VER = FINDPARAMETROI(1, "MDIFORM1", "F_VER", "..", "MERCURY VERSION ")

 F_UPDATESTRUCTURES = Val(FINDPARAMETROI(1, "MDIFORM1", "F_UPDATESTRUCTURES", "1", "Õ¡  ¡Õ≈… UPDATE STRUCTURES=1 OXI=0"))


If F_UPDATESTRUCTURES = 1 Then

'If Trim(F_VER) = Trim(P_VER) Then
        'OK
'Else
340     Update_Structures
        Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + P_VER + "' WHERE  FORMA='MDIFORM1' AND VAR='F_VER' "
'End If

End If



350     LOADPARAMETERS 1

360     MDIForm1.Picture = LoadPicture(gPicture)

370    'MDIForm1.StatusBar1.Panels(1).Text = Combo1.Text

380     MDIForm1.StatusBar1.Panels(4).Text = gUserId

390     LoginSucceeded = True

400     R.Close

410     gWorkDay = Workday.Value

420     Me.Hide

        Exit Sub

NOCONNECT:
430     MsgBox "¡‰˝Ì·ÙÁ Á ÛıÌ›˜ÂÈ·"
440     killmenu
450     gConnect = "end"

460     Unload Me

OUTSIDE:
470     MsgBox "¡‰˝Ì·ÙÁ Á Û˝Ì‰ÂÛÁ"
480     gConnect = "end"

490     killmenu

500     Unload Me

        'Adodc1.Refresh
        On Error Resume Next

        ' Dim R As New ADODB.Recordset

        Exit Sub

CREATE_USERS:

510     Gdb.Execute "CREATE TABLE USERS (  ID_NUM INT IDENTITY(1,1),NAME  NCHAR(30),CODE NCHAR(20),DIKAIOMA  INT )"
520     Gdb.Execute "INSERT INTO USERS ( NAME,CODE,DIKAIOMA ) VALUES ('LAGAKIS','2304603921',99)"
530     Gdb.Execute "INSERT INTO USERS ( NAME,CODE,DIKAIOMA ) VALUES ('1','1',1)"
540     Gdb.Execute "INSERT INTO USERS ( NAME,CODE,DIKAIOMA ) VALUES ('2','2',1)"
550     Gdb.Execute "INSERT INTO USERS ( NAME,CODE,DIKAIOMA ) VALUES ('3','3',9)"
560     Gdb.Execute "INSERT INTO USERS ( NAME,CODE,DIKAIOMA ) VALUES ('4','4',1)"
570     Gdb.Execute "INSERT INTO USERS ( NAME,CODE,DIKAIOMA ) VALUES ('5','5',1)"
580     Gdb.Execute "INSERT INTO USERS ( NAME,CODE,DIKAIOMA ) VALUES ('6','6',1)"
590     Gdb.Execute "INSERT INTO USERS ( NAME,CODE,DIKAIOMA ) VALUES ('7','7',1)"

        On Error Resume Next

600     Resume Next

        '
610     R.Open "select *FROM DIK WHERE LEFT(KOD,4)='0701' AND CODE='" + txtUserName + "'", Gdb, adOpenForwardOnly, adLockReadOnly

620     If R.EOF Then
630         gXEIRISTHS = 1  'ÂÒÈÔÒÈÛÏÂÌÔ ÏÂÌÔı
        Else

640         If Right(R("KOD"), 1) = "8" Then  ' 070108'
650             gXEIRISTHS = 8    ''ssuser
            Else
660             gXEIRISTHS = 9  'admin
            End If

        End If

        ' gXEIRISTHS = 8

        'check for correct password
670     If txtPassword = "" Then
            'place code to here to pass the
            'success to the calling sub
            'setting a global var is the easiest

680         LoginSucceeded = True
690         Me.Hide
        Else
700         MsgBox "Invalid Password, try again!", , "Login"
710         txtPassword.SetFocus
720         SendKeys "{Home}+{End}"
        End If

730     SendKeys "%"

        '<EhFooter>
        Exit Sub

cmdOK_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.frmLogin.cmdOK_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.frmLogin.cmdOK_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub killmenu()

        '<EhHeader>
        On Error GoTo killmenu_Err

        '</EhHeader>
100     With MDIForm1
            'bohu.Visible = False
110         .mapot1(21).Visible = False
120         .pareid(23).Visible = False
130         .apo4(24).Visible = False
140         .apo5(25).Visible = False
150         .mapot6(26).Visible = False
160         .epit.Visible = False

170         .mBohu1(51).Visible = False
180         .mbohu2(52).Visible = False

190         .par3(53).Visible = False
200         .parkin(54).Visible = False
210         .pros(55).Visible = False
220         .genlog(56).Visible = False

            ' .apoemporiko(1).Visible = False 'True
230         .DBF(57).Visible = False
240         .pontoif(58).Visible = False
250         .apoemporiko(59).Visible = False
260         .calculator(510).Visible = False
270         .mpelat1.Visible = False
280         .mpelat2.Visible = False
290         .mpelat3.Visible = False
300         .mpelat4.Visible = False
310         .mpelat5.Visible = False

320         .diagrpar(42).Visible = False
330         .parenos(43).Visible = False
340         .mpar5.Visible = False

350         .mPar6.Visible = False

360         .mbohu12.Visible = False
370         .timspar(44).Visible = False

380         .mapot2(22).Visible = False

            ' .mapot7(27).Visible = False

        End With

        '<EhFooter>
        Exit Sub

killmenu_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.frmLogin.killmenu " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.frmLogin.killmenu " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command1_Click()
  End
End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        Dim mess As String
       
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2
        
        mess = ""

        Dim k As Integer

100     MDIForm1.Timer1.Enabled = False

110     Me.Picture = LoadPicture(gPicture)
120     COUNTER = 0

        '    'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        '   ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        '    'SkinFramework1.ApplyWindow Me.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
        '
        '
        Dim R As New ADODB.Recordset

        'R.Open "select * from ETAIREIES", Gdb, adOpenForwardOnly, adLockReadOnly
130     k = 0
        'Do While Not R.EOF
        '  Combo1.AddItem Left(R("TITLOS") + Space(20), 20) '+ Left(R("DSN") + Space(20), 20)
        '  ETAIREIES(k) = R("DSN")
        '  k = k + 1
        '  R.MoveNext
        'Loop
        'Combo1.Text = Combo1.List(0)

140     Workday = Now

150     If Len(Trim(Text1.Text)) > 0 Then
160         gConnect = gConnect + Text1.Text
        End If

        ' On Error GoTo NOCONNECT
        mess = "¡‰˝Ì·ÙÁ Á Û˝Ì‰ÂÛÁ ÏÂ Ù· ‰Â‰ÔÏ›Ì·"
170     Gdb.Open gConnect

        '≈À≈√◊Ÿ ¡Õ ’–¡—◊≈… œ –…Õ¡ ¡” –≈À¡‘≈” ¡Õ œ◊… ‘œ‘≈ ”«Ã¡…Õ≈… œ‘… ≈…Ã¡… ”≈ À¡»œ” ¬¡”« (MASTER)
175     R.Open "SELECT COUNT(*) AS N  FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME='PEL';", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next
      
        If R(0) = 0 Then
180         Gdb.Execute "CREATE DATABASE  MERCURY    ON (FILENAME = 'C:\MERCVB\MERCURY_DATA.MDF'),(FILENAME = 'C:\MERCVB\MERCURY_log.LDF') FOR ATTACH;"
190         MILSEC 2000
200         Gdb.Close
   
210         gConnect = Replace(gConnect, "MASTER", "MERCURY") ' ' + ";DATABASE=MERCURY;"
   
220         Open "C:\MERCPATH.TXT" For Output As #1
230         Print #1, ".."
250         Print #1, gConnect
260         Close #1
   
270         Gdb.Open gConnect
        End If
      
        Exit Sub

Form_Load_Err:
        MsgBox mess & vbCrLf & Err.Description & vbCrLf & "in ADOMERCNEW.frmLogin.Form_Load " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.frmLogin.Form_Load " & "at line " & Erl

        End

        '  Resume Next
        '</EhFooter>

End Sub

Private Sub Form_Unload(Cancel As Integer)

        '<EhHeader>
        On Error GoTo Form_Unload_Err

        '</EhHeader>

100     MDIForm1.Timer1.Enabled = True

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.frmLogin.Form_Unload " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.frmLogin.Form_Unload " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtPassword_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then

110         KeyAscii = 0    'suppress the beep

120         keybd_event VK_TAB, 0, 0, 0    'send a tab

        End If

        '<EhFooter>
        Exit Sub

txtPassword_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.frmLogin.txtPassword_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.frmLogin.txtPassword_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtUserName_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then

110         KeyAscii = 0    'suppress the beep

120         keybd_event VK_TAB, 0, 0, 0    'send a tab

            End If

        '<EhFooter>
        Exit Sub

txtUserName_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.frmLogin.txtUserName_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.frmLogin.txtUserName_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

