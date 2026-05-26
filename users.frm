VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form bohu11 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   13620
   WindowState     =   2  'Maximized
   Begin VB.CommandButton enhm 
      Caption         =   "Αποθήκευση αλλαγών"
      Height          =   360
      Left            =   5880
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ενεργοποίηση μενού χρήστη"
      Height          =   360
      Index           =   3
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Αλλαγή δικαιωμάτων χρήστη"
      Height          =   360
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Αλλαγή κωδικού χρήστη"
      Height          =   360
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Αλλαγή Ονόματος χρήστη"
      Height          =   360
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   3135
   End
   Begin VB.ComboBox uSERS 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3720
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "MENUS"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "users.frx":0000
      Height          =   2415
      Left            =   360
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Εξοδος"
      Height          =   600
      Left            =   5760
      TabIndex        =   4
      Top             =   7320
      Width           =   1950
   End
   Begin VB.TextBox OLDCODE 
      Enabled         =   0   'False
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   5280
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox NEWCODE 
      BackColor       =   &H0080FF80&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   6840
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label HELP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4080
      TabIndex        =   12
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Παλιό"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Νέο"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "bohu11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim F_BUTTON As Integer
Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
Dim IDS(20) As Integer

Dim F_USERID As Integer





'Private Sub Command1_Click()
'
'        '<EhHeader>
'        On Error GoTo Command1_Click_Err
'
'        '</EhHeader>
'        Dim n As Long
'
''100     If NEWCODE.Text = EPIBEB.Text Then
''            'OK
''        Else
''110         MsgBox "Δεν επιβεβαιώθηκε ο νέος κωδικός"
''
''            Exit Sub
''
''        End If
'
'120     Gdb.Execute "update USERS SET DIKAIOMA=" + DIKAIOMA.Text + ", CODE='" + Trim(NEWCODE.Text) + "'," + " NAME='" + Trim(mNAME.Text) + "'" + " WHERE CODE='" + OLDCODE + "' AND ID_NUM=" + USERID, n
'
'130     If n > 0 Then
'140         MsgBox "ΕΝΗΜΕΡΩΘΗΚΕ"
'        Else
'150         MsgBox "ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ"
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'Command1_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.bohu11.Command1_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu11.Command1_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu11.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu11.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click(Index As Integer)

If uSERS.ListIndex < 0 Then
   MsgBox "Επιλέξτε χρήστη"
   Exit Sub
 End If
   

uSERS.Enabled = False
F_BUTTON = Index
   

 
 Dim N As Integer: N = uSERS.ItemData(uSERS.ListIndex)
 
 F_USERID = N
 
 Dim R As New ADODB.Recordset
R.Open "SELECT * FROM USERS WHERE ID_NUM=" + str(N), Gdb, adOpenDynamic, adLockOptimistic




If Index = 0 Then ' USERNAME
  Label2.Visible = True
  Label4.Visible = True
  OLDCODE.Visible = True
  NEWCODE.Visible = True
  enhm.Visible = True
  NEWCODE.PasswordChar = ""
  HELP.Caption = "Αλλάξτε το όνομα του χρήστη και πατήστε αποθήκευση"
  OLDCODE.Text = R!Name
  
  NEWCODE.SetFocus

End If
If Index = 1 Then ' PASSWORD
  Label2.Visible = True
  Label4.Visible = True
  OLDCODE.Visible = True
  NEWCODE.Visible = True
  enhm.Visible = True
  NEWCODE.PasswordChar = "*"
  OLDCODE.Text = R!code
  HELP.Caption = "Αλλάξτε τον κωδικό του χρήστη και πατήστε αποθήκευση"
  NEWCODE.SetFocus
End If
If Index = 2 Then ' DIKAIOMATA
  Label2.Visible = True
  Label4.Visible = True
  OLDCODE.Visible = True
  NEWCODE.Visible = True
  NEWCODE.SetFocus
  NEWCODE.PasswordChar = ""
  OLDCODE.Text = R!DIKAIOMA
  enhm.Visible = True
  HELP.Caption = "Αλλάξτε τα δικαώματα του χρήστη και πατήστε αποθήκευση.Το 9 σημαίνει διαχειριστής ενώ 1 απλός χρήστης"
End If
If Index = 3 Then ' MENOYS
   Label2.Visible = False
   Label4.Visible = False
   OLDCODE.Visible = False
   NEWCODE.Visible = False
   enhm.Visible = False
   HELP.Caption = "Αλλάξτε τα μενού του χρήστη το 1 σημαίνει ενεργοποιημένο ενώ 0 απενεργοποιημένο"
   
   DataGrid1.Visible = True
   
   Dim RR As New ADODB.Recordset
     
     For N = 1 To MDIForm1.Controls.Count - 1
       If TypeName(MDIForm1.Controls(N)) = "Menu" Then
         RR.Open "SELECT * FROM ACTIVE_MENUS WHERE MENU_ID=" + str(N) + " AND USER_ID=" + str(F_USERID), Gdb, adOpenDynamic, adLockOptimistic
         If RR.EOF Then
            'On Error Resume Next
            Gdb.Execute "INSERT INTO ACTIVE_MENUS (MENU_ID,NAME,ENABLED,USER_ID) VALUES (" + str(N) + ",'" + Left(MDIForm1.Controls(N).Caption, 35) + "',1," + str(F_USERID) + ")"
         End If
         RR.Close
        End If
         
         
     Next
     
     
    Adodc1.RecordSource = "SELECT NAME,ENABLED,MENU_ID,ID FROM ACTIVE_MENUS where USER_ID=" + str(F_USERID)
    Adodc1.ConnectionString = gConnect

    Adodc1.Refresh
    DataGrid1.columns(2).Visible = False
    DataGrid1.columns(3).Visible = False
    
     
     
     
     DataGrid1.SetFocus
End If




















End Sub

Private Sub enhm_Click()
 Dim Index As Integer
 Index = F_BUTTON
 
 If uSERS.ListIndex < 0 Then
   Exit Sub
 End If
 
 Dim N As Integer: N = uSERS.ItemData(uSERS.ListIndex)
 
 
 
If Index = 0 Then ' USERNAME
   Gdb.Execute "update USERS set NAME='" + NEWCODE.Text + "' WHERE ID_NUM=" + str(N)


End If
If Index = 1 Then ' PASSWORD
  Gdb.Execute "update USERS set CODE='" + NEWCODE.Text + "' WHERE ID_NUM=" + str(N)


End If
If Index = 2 Then ' DIKAIOMATA
  'DIKAIOMA
  Gdb.Execute "update USERS set DIKAIOMA=" + NEWCODE.Text + " WHERE ID_NUM=" + str(N)

End If
If Index = 3 Then ' MENOYS



End If

NEWCODE.Text = ""
uSERS.Enabled = True



















End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "BOHU11"
110     PARAMETROI.SHOW 1
        
        
        
        
        
        
        
        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu11.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu11.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err
 F_BUTTON = -1
        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl

        Dim R As New ADODB.Recordset
 
 
 
 load_Combo uSERS, "select NAME,ID_NUM FROM USERS"
 
If gXEIRISTHS < 8 Then
     MsgBox "δεν έχετε δικαιώματα για μεταβολές"
     Unload Me
End If


   'Dim R As New ADODB.Recordset
   Dim N As Integer
 '================================ PRODUCT_TO_CATEGORIES ================================
     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'ACTIVE_MENUS'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

     If R(0) = 0 Then
        Gdb.Execute "CREATE TABLE ACTIVE_MENUS ( ID INT IDENTITY(1,1),ENABLED INT NULL,KOD VARCHAR(16) NULL,MENU_ID int  NULL,NAME VARCHAR(35) NULL ,USER_ID INT null) "


     End If

        R.Close



'110     r.Open "select * from USERS WHERE ID_NUM=" + str(gUserId), Gdb, adOpenForwardOnly, adLockReadOnly
'
'120     If r("DIKAIOMA") = 99 Then
'130        ' DIKAIOMA.Enabled = True
'140        ' USERID.Enabled = True
'150        ' OLDCODE.PasswordChar = ""
'
'        End If
'
'160   '  mNAME = r("NAME")
'170   '  DIKAIOMA = r("DIKAIOMA")
'180    ' USERID = gUserId
'r.Close
'
'
'
'
'
'














        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu11.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu11.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

   ' mForm_Resize Me, 8, 9480, 6120, fh, fw, ft, fl

End Sub

Private Sub USERID_LostFocus()

'        '<EhHeader>
'        On Error GoTo USERID_LostFocus_Err
'
'        '</EhHeader>
'
'        Dim R As New ADODB.Recordset
'
'100     R.Open "select * from USERS WHERE ID_NUM=" + str(USERID), Gdb, adOpenForwardOnly, adLockReadOnly
'
'110     OLDCODE.Text = R("CODE")
'
'        If R("DIKAIOMA") = 99 Then
'            DIKAIOMA.Enabled = True
'            USERID.Enabled = True
'        End If
'
'120     mNAME = R("NAME")
'130     DIKAIOMA = R("DIKAIOMA")
'        'USERID = gUserId
'
'        '<EhFooter>
'        Exit Sub
'
'USERID_LostFocus_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.bohu11.USERID_LostFocus " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu11.USERID_LostFocus " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>

End Sub



Sub PAINT()
'------------------ αλλαζει τα δικαιωματα μενου

 'If gXEIRISTHS < 8 Then
 '  DataGrid1.Visible = False
 'End If
 
        
    Adodc1.RecordSource = "SELECT NAME,ENABLED,MENU_ID,ID FROM ACTIVE_MENUS where USER_ID=" + str(F_USERID)
    Adodc1.ConnectionString = gConnect

    Adodc1.Refresh
    DataGrid1.columns(2).Visible = False
    DataGrid1.columns(3).Visible = False
    


End Sub

