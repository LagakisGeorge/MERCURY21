VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form bohu8 
   BackColor       =   &H00FF0000&
   Caption         =   "¡ÌÙﬂ„Ò·ˆ·"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   12936
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12936
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "·ÌÙÈ„Ò·ˆ· ÛÂ Access "
      Height          =   315
      Left            =   240
      TabIndex        =   21
      Top             =   5620
      Width           =   5175
   End
   Begin VB.CommandButton Command12 
      Caption         =   "·Ô access"
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   6000
      Width           =   5175
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FFFF00&
      Caption         =   "ANTI√—¡÷œ ¬¡”«”"
      Height          =   315
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5240
      Width           =   5175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "RESTORE A–œ BACKUP"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   6360
      Width           =   5175
   End
   Begin VB.CommandButton NeaBash 
      Caption         =   "ƒÁÏÈÔıÒ„È· Õ›·Ú ¬‹ÛÁÚ (˜˘ÒÈÚ Â„„Ò·ˆ›Ú)"
      Height          =   480
      Left            =   4200
      TabIndex        =   17
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdNeaXrhsh 
      Caption         =   "Õ›· ◊ÒﬁÛÁ(ÃÂ Â·Ì·ˆÔÒ‹)"
      Height          =   495
      Left            =   6720
      TabIndex        =   16
      Top             =   480
      Width           =   1500
   End
   Begin VB.ListBox List3 
      Height          =   1200
      Left            =   6600
      TabIndex        =   15
      Top             =   7080
      Width           =   4095
   End
   Begin VB.CommandButton fortosh 
      Caption         =   "ˆÔÒÙ˘ÛÁ"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "·ÌÙÈ„Ò·ˆ· „È· ˆÔÒÙÁ„Ô"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton NEAXRONIA 
      Caption         =   "N›· ·ÎÎ·„ﬁ ˜ÒﬁÛÁÚ(sql2005)"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   8040
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Ã≈‘¡÷œ—¡ ’–œÀœ…–ŸÕ ≈…ƒŸÕ"
      Height          =   495
      Left            =   8520
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton PEL 
      Caption         =   "Ã≈‘¡÷œ—¡ ’–œÀœ…–ŸÕ –≈À¡‘ŸÕ"
      Height          =   495
      Left            =   8520
      TabIndex        =   9
      ToolTipText     =   "–ÒÔÛÔ˜ﬁ Á ÏÂÙ·ˆÔÒ‹ „ﬂÌÂÙ·È ÛÙÁÌ ÙÒ›˜ÔıÛ· ˜ÒÔÌÈ‹,·ˆÔı ÒÔÁ„ÔıÏ›Ì˘Ú ıÔÎÔ„ﬂÛ˘ Ù· ı¸ÎÔÈ· ÛÙÁÌ ·ÎÈ‹ ˜ÒÔÌÈ‹."
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command14 
      Caption         =   "¡ÌÙÈ„Ò·ˆﬁ ·Ò˜Âﬂ˘Ì ÛÂ ·ÎÈ‹ ˜ÒÔÌÈ‹"
      Height          =   495
      Left            =   8520
      TabIndex        =   7
      Top             =   480
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5160
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
      Caption         =   "¡ÔËﬁÍÂıÛÁ"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   6000
      Width           =   2175
   End
   Begin VB.ListBox List2 
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   10455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Text            =   "C:\mercVB\back1.MDB"
      Top             =   4740
      Width           =   3495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Height          =   4080
      Left            =   240
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      Height          =   6615
      Left            =   6480
      Top             =   120
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   6615
      Left            =   240
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label ALLAGH 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "»›ÛÁ dbf ·Ò˜Âﬂ˘Ì"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4740
      Width           =   3525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "bohu8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim F_xre_APOU

Dim F_axre_APOU

Dim F_pis_APOU

Dim F_apis_APOU

Dim Col As TrueOleDBGrid80.Column





Dim F_PATH    ' = FindParametroi(1,"BOHU8", "F_PATH", "C:\MERCVB", "ƒ…¡ƒ—œÃ« ¡Õ‘…√—¡÷ŸÕ") 'posa psifia tha exei h kathe seira

Dim F_PATHFORT As String

Dim F_BACKUPTYPE    ' 1=SQL  0=ACCESS

Dim ColS As TrueOleDBGrid80.columns

Dim f_psifia_atim    ' = Val(FindParametroi(1,"PAR1", "F_PSIFIA_ATIM", "5", "√. 9. ÿ«÷…¡ ¡—…»Ãœ’ ‘…ÃœÀœ√…œ’ (5)"))

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)



Private Sub cmdNeaXrhsh_Click()

        Dim c As String
        Dim TREX As String, PALIA As String
        
        Dim MKOD As String
        
140     MKOD = InputBox("¡Ì ÂﬂÛ·È Ûﬂ„ÔıÒÔÚ ‰˛ÛÂ Í˘‰ÈÍ¸ ")

150     If Val(MKOD) = 12345 Then
            'ok
        Else

            Exit Sub

        End If
        
        
        
        
        
        
        
170             TREX = InputBox("NEA” ˜ÒÔÌÈ‹", , Year(Date))

180     PALIA = InputBox("œÌÔÏ· ·ÎÈ‹Ú ˜ÒÔÌÈ‹Ú", , Year(Date) - 1)

190     If TREX <= PALIA Or TREX - PALIA > 1 Then
200         ANS = MsgBox("À¡»œ” ◊—œÕ…≈”. ≈…”‘≈ ”…√œ’—œ…;", vbYesNo)

210         If ANS = vbNo Then

                Exit Sub

            End If

        End If
        
Dim RR As New ADODB.Recordset
    ' RR.Open "SELECT  COUNT(*) FROM  ETAIREIES WHERE DSN='" + "D" + Format(PALIA, "0000") + "'", Gdb, adOpenDynamic, adLockOptimistic
     RR.Open "SELECT  COUNT(*) FROM  sys.databases WHERE name='" + "D" + Format(PALIA, "0000") + "'", Gdb, adOpenDynamic, adLockOptimistic
     ' RR.Open "SELECT  COUNT(*) FROM  ETAIREIES WHERE DSN='" + "D" + Format(PALIA, "0000") + "'", Gdb, adOpenDynamic, adLockOptimistic
   If RR(0) > 0 Then
      MsgBox "≈◊≈… «ƒ« ƒ«Ã…œ’—√«»≈… « ◊—œÕ…¡ " + "D" + Format(PALIA, "0000") + " ¡ƒ’Õ¡‘« ”’Õ≈◊≈…¡"
      Exit Sub
   End If
   
     
  
        

100     'C = InputBox("ƒŸ”≈ œÕœÃ¡ BACKUP ‘«” ¬¡”«” " + Gdb.DefaultDatabase, , "C:\MERCVB\BACK" + Replace(time$, ":", "") + ".bak")


         c = InputBox("ƒŸ”≈ œÕœÃ¡ BACKUP ‘«” ¬¡”«” " + Gdb.DefaultDatabase, , "C:\MERCVB\BACK" + Replace(time$, ":", "") + ".bak")
Gdb.CommandTimeout = 0

         Gdb.Execute "BACKUP DATABASE [" + Gdb.DefaultDatabase + "] TO  DISK ='" + c + "' WITH NOFORMAT, NOINIT, SKIP, NOREWIND, NOUNLOAD,  STATS = 10", k
         



Dim c2 As String
c2 = Gdb.DefaultDatabase
List3.AddItem "BACKUP OK"


Dim CN As String
CN = ""









'‰ÁÏÈÔıÒ„Âﬂ ÙÁÌ Ì›· ‚‹ÛÁ ant2 ·Ô ÙÔ backup  C:\mercvb\ant.bak
' Í·È Ù· ·Ò˜Âﬂ· mercury_data & mercury_log Ù· ·ÌÙÈ„Ò‹ˆÂÈ ÛÙ· C:\mercvb\ant2.mdf & C:\mercvb\ant2_log.ldf

'RESTORE DATABASE ant2 FROM DISK='C:\mercvb\ant.bak'
'WITH
'Move 'mercury_data' TO 'C:\mercvb\ant2.mdf',
'Move 'mercury_log' TO 'C:\mercvb\ant2_log.ldf'


'RESTORE DATABASE [emp2] FROM DISK = 'c:\mercvb\emp2'
'WITH
'Move 'EMP_Data' TO 'c:\MERCVB\EMP2_Data.mdf',
'Move 'EMP_Log' TO 'c:\MERCVB\EMP2_Log.ldf',
'RECOVERY, REPLACE, STATS = 10;






Dim log_name As String
log_name = InputBox("ƒŸ”≈ œÕœÃ¡ (LOGICAL) ‘«” ¬¡”«” " + Gdb.DefaultDatabase, , "MERCURY")


        Dim CPALIA

220     CPALIA = "D" + Format(PALIA, "0000")


c2 = "RESTORE  DATABASE " + CPALIA + " FROM DISK='" + c + "'  WITH "
c2 = c2 + " Move '" + log_name + "_Data' TO 'C:\mercvb\" + CPALIA + "_Data.mdf',  "
c2 = c2 + " Move '" + log_name + "_Log' TO 'C:\mercvb\" + CPALIA + "_log.ldf',"
c2 = c2 + " RECOVERY, REPLACE, STATS = 10; "
On Error GoTo LATOS
Gdb.Execute c2
'Exit Sub


List3.AddItem "RESTORE OK"


        Dim sql  As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

        Dim CONNECT_PALIA As String

240     If InStr(gConnect, "DATABASE") = 0 Then
250         CONNECT_PALIA = gConnect + ";" + "DATABASE=" + CPALIA + ";"
        Else
260         CONNECT_PALIA = Left(gConnect, InStr(gConnect, "DATABASE") - 1) + ";" + "DATABASE=" + CPALIA + ";"
        End If

        Dim N


        Dim GDBPALIA As New ADODB.Connection

        GDBPALIA.Open CONNECT_PALIA

       

        Dim SQL3     As String



 Gdb.Execute "INSERT INTO ETAIREIES (TITLOS,DSN)  VALUES ('" + Format(PALIA, "0000") + "','" + "D" + Format(PALIA, "0000") + "')"

510     Gdb.Execute "DELETE  FROM EGG WHERE YEAR(HME)<=" + str(TREX - 1), N
520     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)<7 AND YEAR(HME)<=" + str(TREX - 1), N
530     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=7 AND YEAR(HME)<=" + str(TREX - 1), N
540     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=8 AND YEAR(HME)<=" + str(TREX - 1), N
550     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=9 AND YEAR(HME)<=" + str(TREX - 1), N
560     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=10 AND YEAR(HME)<=" + str(TREX - 1), N
570     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=11 AND YEAR(HME)<=" + str(TREX - 1), N
580     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=12  AND YEAR(HME)<=" + str(TREX - 1), N

590     Gdb.Execute "DELETE  FROM TIM WHERE YEAR(HME)<=" + str(TREX - 1), N

       List3.AddItem "ƒ…¡√—¡÷« –¡À…ŸÕ OK"
 
 
 
     GDBPALIA.Execute "DELETE  FROM EGG     WHERE YEAR(HME)>=" + str(TREX), N
    GDBPALIA.Execute "DELETE   FROM EGGTIM  WHERE YEAR(HME)>=" + str(TREX), N
    GDBPALIA.Execute "DELETE   FROM TIM     WHERE YEAR(HME)>=" + str(TREX), N
     
  List3.AddItem " –¡À…¡ OK"
    






        List3.AddItem "‘≈Àœ” Ã≈‘¡÷œ—¡” "


MsgBox "‘≈Àœ” Ã≈‘¡÷œ—¡” "


Exit Sub


LATOS:

MsgBox "‰ÂÌ Â„ÈÌÂ Á ·ÎÎ·„Á ˜ÒÔÌÈ‹Ú " + Chr(13) + Err.Description























End Sub

'Private Sub cmdClearFilter_Click()
'
''Clears filter from grid
'
'    For Each Col In TDBGrid.Columns
'
'        Col.FilterText = ""
'
'    Next Col
'
'    Adodc2.Recordset.Filter = adFilterNone
'
'End Sub

Private Sub Command11_Click()

        '<EhHeader>
        On Error GoTo Command11_Click_Err

        '</EhHeader>

100     cd1.InitDir = "c:\mercvb\queries"

110     cd1.ShowSave

        Dim F

120     F = cd1.FILENAME

130     Open F For Output As #5
140     Print #5, Text2.Text
150     Close #5

        '+ ".txt"

        '<EhFooter>
        Exit Sub

Command11_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Command11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Command11_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command12_Click()

        'Dim Gdb As New ADODB.Connection
        'Dim GR As New ADODB.Recordset
        '<EhHeader>
        On Error GoTo Command12_Click_Err

        '</EhHeader>

        Dim k As Integer

        Dim R As New ADODB.Recordset

100     ANS = MsgBox("–ÒÔÛÔ˜ﬁ Ë· Û‚ÁÛÙÔ˝Ì Ù· ·Ò˜Âﬂ· ÙÔı SERVER  " + text1text + "  ≈ﬂÛ·È Ûﬂ„ÔıÒÔÚ;", vbYesNo)

110     If ANS = vbNo Then Exit Sub

        Dim MKOD As String

120     If Len(Dir(Text1.Text)) > 2 Then
            'ok
        Else
130         MsgBox "Î·ËÔÚ Í·Ù‹ÎÔ„ÔÚ"

            Exit Sub

        End If

140     If Not ANS = vbYes Then
150         MsgBox " ¡ ’—ø»« ≈ « ≈—√¡”∫¡"

            Exit Sub

        End If

160     MKOD = InputBox("¡Ì ÂﬂÛ·È Ûﬂ„ÔıÒÔÚ ‰˛ÛÂ Í˘‰ÈÍ¸ ")

170     If Val(MKOD) = 12345 Then
            'ok
        Else

            Exit Sub

        End If

        Dim DUM

180     DD = Text1.Text

        On Error Resume Next

        On Error GoTo LATOS2

        On Error Resume Next

190     If F_BACKUPTYPE = 1 Then

            Dim CC As String

            Dim nn As Long

200         CC = Gdb.DefaultDatabase  '"EMP"  ' Gdb.Properties.Item(0).value   '

            '  Gdb.Close
            On Error GoTo Command12_Click_Err

            ' Gdb.Open gConnect + ";DATABASE=MASTER"
210         Gdb.Close
220         MDIForm1.TELEFON.ConnectionString = ""
230         Gdb.Open gConnect + ";DATABASE=MASTER"

            On Error GoTo LATOS3

            'Unload MDIForm1

240         Gdb.Execute "RESTORE DATABASE [" + CC + "] FROM  DISK = '" + Text1.Text + "' WITH  FILE = 1,  NOUNLOAD,  REPLACE,  STATS = 10", k

250         Gdb.Close
260         Gdb.Open gConnect

        End If

        Dim DB As Database

270     Set DB = OpenDatabase("", False, False, gConnect)

        Dim RR As Recordset

        On Error GoTo EXO

        On Error GoTo Command12_Click_Err

280     Me.MousePointer = vbHourglass

290     List2.Visible = True
        'Gdb.Open "DSN=2004" 1.058
300     List2.Clear

310     For k = 0 To List1.ListCount - 1
320         LATOS = 0

            On Error GoTo LATOS

330         DoEvents
340         Me.Caption = List1.List(k)

350         If List1.Selected(k) Then

                ' Set rr = db.OpenRecordset(List1.List(K))
                On Error GoTo EXO

                On Error Resume Next

360             Gdb.Execute "DROP TABLE " + List1.List(k)

370             DB.Execute "SELECT * INTO " + List1.List(k) + "  FROM  " + List1.List(k) + "  IN '" + DD + "'"

380             If List1.List(k) = "TIM" Then

                    On Error GoTo Command12_Click_Err

                    On Error Resume Next

390                 Gdb.Execute "ALTER TABLE TIM ADD  ID_NUM2 int "
400                 Gdb.Execute "UPDATE TIM SET ID_NUM2 = ID_NUM"
410                 Gdb.Execute "ALTER TABLE TIM DROP COLUMN ID_NUM "
420                 Gdb.Execute "ALTER TABLE TIM ADD  ID_NUM int IDENTITY(1,1)"
                    '      Gdb.Execute "UPDATE TIM SET ID_NUM = ID_NUM2"

430                 Gdb.Execute "ALTER TABLE GRA DROP COLUMN ID "
440                 Gdb.Execute "ALTER TABLE PARAMETROI DROP COLUMN ID "
450                 Gdb.Execute "ALTER TABLE EGG DROP COLUMN ID "
460                 Gdb.Execute "ALTER TABLE PARASTAT DROP COLUMN ID "

                End If

470             If LATOS = 0 Then
480                 List2.AddItem "œ     " + List1.List(k)
                Else
490                 List2.AddItem "À¡»œ” " + List1.List(k)
                End If

            End If

        Next

500     Me.MousePointer = vbNormal

        Exit Sub

LATOS:
510     List2.AddItem Err.Description
520     LATOS = 1

530     Resume Next

540     Return

LATOS3:
550     MsgBox "·‰ıÌ·Ïﬂ· ÂÍÙ›ÎÂÛÁÚ Â·Ì·ˆÔÒ‹Ú" + Chr(13) + Err.Description
560     LATOS = 1

570     Resume Next

580     Return

LATOS2:
590     MsgBox "À¡»œ” œÕœÃ¡ ¬¡”«” ACCESS  ƒŸ”‘≈ –.◊.  C:\BACK.MDB"
600     Return

EXO:
610     MsgBox "ƒŸ”¡‘≈ À¡»œ” ƒ…” œ"

        Exit Sub

        '<EhFooter>
        Exit Sub

Command12_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Command12_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Command12_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Command13_Click()
'' METAFORA YPOLOIPOY EIDON
'    Dim arxeio
'    Dim a As Long
'    Dim DBF As Database
'    Dim sql As New ADODB.Connection
'    Dim rSQLEGGT As New ADODB.Recordset
'    Dim rSQL As New ADODB.Recordset
'
'    Dim conSQL As String
'    Dim k As Long
'    Dim Fname As String
'    Dim hm As String
'
'    hm = InputBox("‰˘ÛÂ ÁÏÂÒÔÏÁÌﬂ· ·Ô„Ò·ˆﬁÚ", , "01/01/2006")
'    If IsDate(hm) Then
'        'ok
'    Else
'        MsgBox "Î·ËÔÚ ÁÏÂÒÔÏÁÌﬂ·"
'    End If
'
'
'    a = GetTickCount()
'    Me.Caption = "ENHME—Ÿ”« Ã≈ ‘≈À≈’‘¡…¡ ’–œÀœ…–¡"
'    UPDATE_YPOLOIPA2
'    arxeio = "EID"
'    sql.Open gConnect
'    On Error GoTo 0
'    rSQL.Open "SELECT * FROM EID WHERE POS>0 ", Gdb, adOpenForwardOnly, adLockReadOnly
'    rSQLEGGT.Open "SELECT * FROM EGGTIM ", Gdb, adOpenDynamic, adLockBatchOptimistic
'
'    rSQL.MoveFirst
'    Dim Z
'    Z = 0
'    On Error GoTo WRITEERROR    ' Resume Next
'
'    Do While Not rSQL.EOF
'
'        If Not IsNull(rSQL("POS01")) Then
'            If rSQL("POS01") > 0 Then
'                rSQLEGGT.AddNew
'                rSQLEGGT("hme") = CDate(hm)
'                rSQLEGGT("ATIM") = "Î00001"
'                rSQLEGGT("KODE") = rSQL("KOD")
'                rSQLEGGT("ONOMA") = rSQL("ONO")
'                rSQLEGGT("TIMM") = rSQL("XTI")
'                rSQLEGGT("FPA") = rSQL("FPA")
'                rSQLEGGT("POS0") = rSQL("POS01")
'                rSQLEGGT("APOT") = 1
'                rSQLEGGT.Update
'            End If
'        End If
'        If Not IsNull(rSQL("POS02")) Then
'            If rSQL("POS02") > 0 Then
'                rSQLEGGT.AddNew
'                rSQLEGGT("hme") = CDate(hm)
'                rSQLEGGT("ATIM") = "Î00001"
'                rSQLEGGT("KODE") = rSQL("KOD")
'                rSQLEGGT("ONOMA") = rSQL("ONO")
'                rSQLEGGT("TIMM") = rSQL("XTI")
'                rSQLEGGT("FPA") = rSQL("FPA")
'                rSQLEGGT("POS0") = rSQL("POS01")
'                rSQLEGGT("APOT") = 2
'                rSQLEGGT.Update
'            End If
'        End If
'        Z = Z + 1
'        If Z Mod 10 = 0 Then
'            Me.Caption = Z
'            DoEvents
'        End If
'        rSQL.MoveNext
'    Loop
'
'    rSQL.Close
'    rSQLEGGT.Close
'
'    Me.Caption = (GetTickCount() - a) / 1000
'
'
'
'
'    Exit Sub
'
'
'WRITEERROR:
'    List2.Visible = True
'
'    List2.AddItem Err.Description
'    Resume Next
'
'End Sub

'Private Sub File1_Click()
'    Dim a
'    a = 0
'
'    Dim F As String
'
'    F = File1.FILENAME
'
'
'    Dim ss As String
'    Dim b As String
'
'    ss = ""
'    Open File1.Path + "\" + F For Input As #1
'    Do While Not EOF(1)
'
'        Line Input #1, b
'        ss = ss + b + " "
'    Loop
'    Close #1
'
'
'    Text2.Text = ss
'
'
'
'End Sub

'Private Sub TDBGrid_FilterChange()
'
''Gets called when an action is performed on the filter bar
'
'    On Error GoTo ErrHandler
'
'
'
'    Set ColS = TDBGrid.Columns
'
'    Dim C As Integer
'
'    C = TDBGrid.Col
'
'    TDBGrid.HoldFields
'
'    Adodc2.Recordset.Filter = getFilter()
'
'    TDBGrid.Col = C
'
'    TDBGrid.EditActive = True
'
'    Exit Sub
'
'
'
'ErrHandler:
'
'    MsgBox Err.Source & ":" & vbCrLf & Err.Description
'
'    Call cmdClearFilter_Click
'
'End Sub

'Private Function getFilter() As String
''Creates the SQL statement in adodc1.recordset.filter
''and only filters text currently. It must be modified to filter other data types.
'
'    Dim tmp As String
'    Dim N As Integer
'    For Each Col In ColS
'        If Trim(Col.FilterText) <> "" Then
'            N = N + 1
'            If N > 1 Then
'                tmp = tmp & " AND "
'            End If
'
'
'            If Adodc2.Recordset(Col.ColIndex).Type = 5 Then    '
'                If InStr(">< >= <=  = ", Left(Col.FilterText, 1)) > 0 And Len(Col.FilterText) > 1 And IsNumeric(Right(Col.FilterText, 1)) Then
'                    tmp = tmp & Col.DataField & Col.FilterText
'                Else
'                    tmp = tmp & Col.DataField & " >  " & Val(Col.FilterText)
'                End If
'
'            Else
'                tmp = tmp & Col.DataField & " LIKE '" & Col.FilterText & "*'"
'            End If
'
'        End If
'    Next Col
'
'    getFilter = tmp
'End Function

'Private Sub APOTSETIM_Click()
'
'    Gdb.Execute "CREATE VIEW dbo.[UPD_TIM] as " _
'              & " SELECT    dbo.TIM.B_N1 , dbo.EGGTIM.APOT " _
'              & "  FROM         dbo.TIM LEFT JOIN" _
'              & "  dbo.EGGTIM ON dbo.TIM.HME = dbo.EGGTIM.HME AND dbo.TIM.ATIM = dbo.EGGTIM.ATIM"
'
'
'
'
'
'    Gdb.Execute "UPDATE UPD_TIM  SET B_N1=APOT"
'
'End Sub

Private Sub Command1_Click()

        '·ÌÙÈ„Ò·ˆ· „È· ˆÔÒÙÁ„Ô
        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim k As Integer

        Dim R As New ADODB.Recordset

        Dim DUM

        Dim mBUFF

100     mBUFF = "e"

        '
        '
        '' ’–œÀœ√…∆Ÿ ‘¡ ‘≈À≈’‘¡…A ’–œÀœ…–¡ ¡–œ ‘œ EGG
        'Gdb.Execute "DROP TABLE EGG1"
        'Dim L As Long
        'Gdb.Execute "UPDATE PEL  SET TYP=0,XRESYN=0,PISSYN=0 FROM" _
        '    & " PEL WHERE EIDOS='" + mBUFF + "'", L
        '
        'Gdb.Execute "SELECT EIDOS as MEIDOS,KOD,SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END ) AS SX," _
        '& " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END ) AS SP INTO EGG1" _
        '& " FROM EGG where EIDOS='" + mBUFF + "' AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "'  GROUP BY EIDOS,KOD ", L
        '
        '
        'Gdb.Execute "UPDATE PEL  SET TYP=AYP WHERE EIDOS='" + mBUFF + "'"
        '
        'Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" _
        '    & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP " _
        '    & " from PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.MEIDOS WHERE PEL.EIDOS='e' ", L
        '

110     DD = F_PATHFORT    ' "C:\LAGEURO\FORTHGO.MDB"

        On Error Resume Next

120     Kill DD

        On Error GoTo LATOS2

130     DUM = CREATE_DATABASE(DD)

        On Error Resume Next

        Dim c As String

140     c = InputBox("ƒŸ”≈ ¡—…»Ãœ ”’√ ≈Õ‘—Ÿ‘… œ’ ÷œ—‘Ÿ”«”")

150     c = "Ù" + Right(String(f_psifia_atim, "0") + c, f_psifia_atim)

        On Error GoTo Command1_Click_Err

160     R.Open "select count(*)  from EGGTIM WHERE LEFT(ATIM," + str(1 + f_psifia_atim) + ")='" + c + "'", Gdb, adOpenDynamic, adLockOptimistic

170     If R(0) = 0 Then
180         MsgBox "ƒ≈Õ ¬—≈»« ≈ ‘œ ”’√ ≈Õ‘—Ÿ‘… œ"

            Exit Sub

        End If

        Dim DB As Database

190     Set DB = OpenDatabase("", False, False, gConnect)

200     DB.Execute "SELECT * INTO EGGTIM   IN '" + DD + "'  FROM  EGGTIM WHERE LEFT(ATIM," + str(1 + F_PSIFIATIMOLOGIOY) + ")='" + c + "'"

210     Me.MousePointer = vbHourglass

220     List2.Visible = True
        'Gdb.Open "DSN=2004" 1.058
230     List2.Clear

240     DB.Execute "SELECT * INTO EID  IN '" + DD + "'  FROM  EID", ANS
250     DB.Execute "SELECT * INTO EGG  IN '" + DD + "'  FROM  EGG", ANS

260     DB.Execute "SELECT * INTO BARCODES  IN '" + DD + "'  FROM  BARCODES ", ANS
270     DB.Execute "SELECT * INTO PEL   IN '" + DD + "'  FROM  PEL ", ANS

280     Me.MousePointer = vbNormal

290     MsgBox "≈√…Õ≈ « ≈Œ¡√Ÿ√« ”‘œ " + F_PATHFORT

        Exit Sub

LATOS:
300     List2.AddItem Err.Description
310     LATOS = 1

320     Resume Next

330     Return

LATOS2:
340     MsgBox "À¡»œ” œÕœÃ¡ ¬¡”«” ACCESS  ƒŸ”‘≈ –.◊.  C:\BACK.MDB"

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub update_sql_from_dbf(arxeio As String)

        '<EhHeader>
        On Error GoTo update_sql_from_dbf_Err

        '</EhHeader>
        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim conDBF As String

        Dim conSQL As String

        Dim k      As Long

        Dim Fname  As String

100     g2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
        '2004 G2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
110     conSQL = gConnect    ' "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP;Data Source=LAGAKIS"
120     conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO"
        '2004 conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO\04"

130     sql.Open conSQL
        'SET DBF=DATA.Open conDBF

140     Data1.DatabaseName = Text1.Text

150     Data1.RecordSource = arxeio
160     Data1.Refresh

170     Set rDBF = Data1.Recordset

180     Label1.Caption = arxeio + " " + Format(rDBF.RecordCount, "######")
190     sql.Execute "DELETE FROM " + arxeio

200     rSQL.Open "SELECT *FROM " + arxeio, sql, adOpenDynamic, adLockOptimistic
        ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

210     rDBF.MoveFirst

        Dim Z

220     Z = 0

        On Error GoTo WRITEERROR    ' Resume Next

230     Do While Not rDBF.EOF

240         rSQL.AddNew

250         For k = 0 To rDBF.FIELDS.Count - 1
260             Fname = rDBF.FIELDS(k).Name    ' p.x. FNAME=epo    R(0).NAME

270             If IsNull(rDBF(k)) Then
                Else
280                 rSQL(Fname) = rDBF(k)    ' rsqk("epo")=r(0)
                End If

                ' ¡’‘¡ –œ’ ≈…Õ¡… ÕULL ‘¡ √≈Ã…∆≈…
290             If IsNull(rDBF(k)) Then
300                 If rDBF(k).Type = dbText Then    '10=STRING
310                     rSQL(Fname) = " "
320                 ElseIf rDBF(k).Type = dbDate Then
330                     rSQL(Fname) = CDate("01/01/1990")
340                 ElseIf rDBF(k).Type = dbBoolean Then
350                     rSQL(Fname) = False
360                 ElseIf rDBF(k).Type = dbDouble Then
370                     rSQL(Fname) = 0
                    Else
380                     conSQL = ""
                    End If
                End If

            Next

            ' Ã≈‘¡‘—≈–≈… ‘«Õ "◊R≈OSI" ”≈ "XRE" KAI "PISTOSI"
390         If arxeio = "EGG" Then
400             rSQL("PISTOSI") = 0: rSQL("XREOSI") = 0

410             If Not IsNull(rSQL("XRE")) Then
420                 If rSQL("EIDOS") = "e" Then
430                     If rSQL("XRE") > 0 Then rSQL("XREOSI") = rSQL("XRE")
440                     If rSQL("XRE") < 0 Then rSQL("PISTOSI") = -rSQL("XRE")
                    Else

450                     If rSQL("XRE") > 0 Then rSQL("PISTOSI") = rSQL("XRE")
460                     If rSQL("XRE") < 0 Then rSQL("XREOSI") = -rSQL("XRE")
                    End If
                End If
            End If

            ' MET¡‘—≈–≈… ‘œ 'POSO' ”≈ 'XRE' KAI 'PIS'
470         If arxeio = "EGGTIM" Then
480             rSQL("PIS") = 0: rSQL("XRE") = 0

490             If Not IsNull(rSQL("POSO")) Then
500                 If InStr(F_xre_APOU + F_axre_APOU, Left(rSQL("ATIM"), 1)) > 0 Then
510                     rSQL("XRE") = rSQL("POSO")
                    End If

520                 If InStr(F_pis_APOU + F_apis_APOU, Left(rSQL("ATIM"), 1)) > 0 Then
530                     rSQL("PIS") = rSQL("POSO")
                    End If
                End If
            End If

540         rSQL.Update
550         Z = Z + 1

560         If Z Mod 10 = 0 Then
570             Me.Caption = Z
            End If

580         rDBF.MoveNext

590         DoEvents

        Loop

600     rSQL.Close
610     rDBF.Close

        Exit Sub

WRITEERROR:
620     List2.Visible = True

630     List2.AddItem Err.Description

640     Resume Next

        '<EhFooter>
        Exit Sub

update_sql_from_dbf_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.update_sql_from_dbf " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.update_sql_from_dbf " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command10_Click()

        '<EhHeader>
        On Error GoTo Command10_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command10_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Command10_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Command10_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command14_Click()

        '<EhHeader>
        On Error GoTo Command14_Click_Err

        '</EhHeader>

        Dim DB       As Database

        Dim CA       As String

        Dim DB_PALIA As New ADODB.Connection

        Dim PALIA    As String

        Dim TREX     As String

        Dim ANS

100     ANS = MsgBox("ÒÔÛÔ˜Á Ë· ÏÂÙ·ˆÂÒËÔ˝Ì Ù· ·Ò˜Âﬂ· ÛÙÁÌ ·ÎÈ‹ ˜ÒÔÌÈ‹" & Chr(13) + " Í·È Ë· ÏÁ‰ÂÌÈÛÙÂﬂ Á ÙÒ›˜ÔıÛ·" + Chr(13) & "≈ﬂÛÙÂ Ûﬂ„ÔıÒÔÈ;", vbYesNo)

110     If ANS = vbYes Then
        Else

            Exit Sub

        End If

120     ALLAGH.Caption = "1.ƒÈ·‰ÈÍ·Ûﬂ· BACKUP"

        ' TO SBHNV GIA NA PARV FRESKO BACKUP
130     DD = "C:\MERCVB\BACK" + Format(Day(Now), "00") + ".MDB"

        On Error Resume Next

140     Kill DD
        ' Kill "C:\MERCVB\BACK" + Format(Day(Now), "00") + ".LDB"

150     MDIForm1.MakeBackup "C:\MERCVB"

160     TREX = InputBox("NEA” ˜ÒÔÌÈ‹", , Year(Date))

170     PALIA = InputBox("œÌÔÏ· ·ÎÈ‹Ú ˜ÒÔÌÈ‹Ú", , Year(Date) - 1)

180     If TREX <= PALIA Or TREX - PALIA > 1 Then
190         MsgBox "À¡»œ” ◊—œÕ…≈”. ≈Œœƒœ”"

            Exit Sub

        End If

        ' Gdb.Execute "create database " + palia
        'DD = Text1.Text
        On Error Resume Next

        '================================================

        On Error GoTo Command14_Click_Err

        Dim CPALIA

200     CPALIA = "D" + Format(PALIA, "0000")

        Dim sql  As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

210     Gdb.Execute "create database " + CPALIA

        'Gdb.DefaultDatabase = CPALIA
        Dim CONNECT_PALIA As String

220     If InStr(gConnect, "DATABASE") = 0 Then
230         CONNECT_PALIA = gConnect + ";" + "DATABASE=" + CPALIA + ";"
        Else
240         CONNECT_PALIA = Left(gConnect, InStr(gConnect, "DATABASE") - 1) + ";" + "DATABASE=" + CPALIA + ";"
        End If

        '---------- RESTORE –¡À…¡” ◊—œÕ…¡” ”‘«Õ DATABASE ‘«” –¡À…¡” ◊—œÕ…¡” -------------------------------------
        On Error GoTo Command14_Click_Err

250     Set DB = OpenDatabase("", False, False, CONNECT_PALIA)
260     Me.MousePointer = vbHourglass

        Dim LISTA(20)

270     LISTA(1) = "PEL"
280     LISTA(2) = "EID"
290     LISTA(3) = "EGG"
300     LISTA(4) = "EGGTIM"
310     LISTA(5) = "TIM"
320     LISTA(6) = "GRA"
330     LISTA(7) = "MEM"
340     LISTA(8) = "PINAKES"
350     LISTA(9) = "XREOPIS"
360     LISTA(10) = "ARITMISI"
370     LISTA(11) = "PARASTAT"
380     LISTA(12) = "BARCODES"
390     LISTA(13) = "HMEROL"
400     LISTA(14) = "HMEROL2"
410     LISTA(15) = "PARAMETROI"
420     LISTA(16) = "CalendarEvents"
430     LISTA(17) = "CalendarRecurrencePatterns"
440     LISTA(18) = "ETAIREIES"
450     ALLAGH.Caption = "2.¡Õ‘…√—¡÷« ”≈ –¡À…¡ ◊—œÕ…¡"

460     DB_PALIA.Open CONNECT_PALIA  ' Left(gConnect, InStr(gConnect, "DATABASE") - 1) + ";DATABASE=" + CPALIA

470     For k = 1 To 18
480         LATOS = 0

            '   On Error GoTo LATOS
490         DoEvents
500         Me.Caption = LISTA(k)

            On Error Resume Next

510         DB_PALIA.Execute "DROP TABLE " + LISTA(k)

520         If LISTA(k) = "EGG" Or LISTA(k) = "EGGTIM" Or LISTA(k) = "TIM" Then
530             DB.Execute "SELECT * INTO " + LISTA(k) + "  FROM  " + LISTA(k) + "  IN '" + DD + "' WHERE YEAR(HME)=" + PALIA
            Else

540             DB.Execute "SELECT * INTO " + LISTA(k) + "  FROM  " + LISTA(k) + "  IN '" + DD + "'"
            End If

        Next

550     Me.MousePointer = vbNormal

560     Gdb.Close

570     Gdb.Open gConnect

        '--- ≈À≈√◊œ” ¡Õ Ã≈‘¡÷≈—»« ¡Õ œ…  …Õ«”≈…” ------------------
580     ALLAGH.Caption = "2¡.≈À≈√◊œ” ¡Õ Ã≈‘¡÷≈—»« ¡Õ œ…  …Õ«”≈…”"

        Dim R  As New ADODB.Recordset

        Dim r2 As New ADODB.Recordset

        On Error GoTo Command14_Click_Err

590     R.Open "SELECT SUM(AJI) FROM TIM WHERE YEAR(HME)=" + PALIA, Gdb, adOpenForwardOnly, adLockReadOnly
600     r2.Open "SELECT SUM(AJI) FROM TIM WHERE YEAR(HME)=" + PALIA, DB_PALIA, adOpenForwardOnly, adLockReadOnly

610     If Abs(R(0) - r2(0)) < 1 Then
            'OK
        Else
620         MsgBox "–—œ¬À«Ã¡ ”‘«Õ Ã≈‘¡÷œ—¡. « ◊—œÕ…¡ –¡—¡Ã≈Õ≈… Ÿ” ≈◊≈…"
630         DB_PALIA.Close

            Exit Sub

        End If

640     R.Close
650     r2.Close

660     R.Open "SELECT SUM(XRE) FROM EGG WHERE YEAR(HME)=" + PALIA, Gdb, adOpenForwardOnly, adLockReadOnly
670     r2.Open "SELECT SUM(XRE) FROM EGG WHERE YEAR(HME)=" + PALIA, DB_PALIA, adOpenForwardOnly, adLockReadOnly

680     If Abs(R(0) - r2(0)) < 1 Then
            'OK
        Else
690         MsgBox "–—œ¬À«Ã¡ ”‘«Õ Ã≈‘¡÷œ—¡. « ◊—œÕ…¡ –¡—¡Ã≈Õ≈… Ÿ” ≈◊≈…"
700         DB_PALIA.Close

            Exit Sub

        End If

710     R.Close
720     r2.Close
730     R.Open "SELECT SUM(POSO) FROM EGGTIM WHERE YEAR(HME)=" + PALIA, Gdb, adOpenForwardOnly, adLockReadOnly
740     r2.Open "SELECT SUM(POSO) FROM EGGTIM WHERE YEAR(HME)=" + PALIA, DB_PALIA, adOpenForwardOnly, adLockReadOnly

750     If Abs(R(0) - r2(0)) < 1 Then
            'OK
        Else
760         MsgBox "–—œ¬À«Ã¡ ”‘«Õ Ã≈‘¡÷œ—¡. « ◊—œÕ…¡ –¡—¡Ã≈Õ≈… Ÿ” ≈◊≈…"
770         DB_PALIA.Close

            Exit Sub

        End If

780     Gdb.Execute "INSERT INTO ETAIREIES (TITLOS,DSN)  VALUES ('" + Format(PALIA, "0000") + "','" + "D" + Format(PALIA, "0000") + "')"

790     ALLAGH.Caption = "3.”¬«”…Ãœ  …Õ«”≈ŸÕ ¡–œ ‘—≈◊œ’”¡ ◊—œÕ…¡"
        '--------------------- Ã«ƒ≈Õ…”Ãœ” ‘…ÃœÀœ√…ŸÕ ----------------------------
800     Gdb.Execute "delete from TIM WHERE YEAR(HME)=" + PALIA
810     Gdb.Execute "delete from EGGTIM WHERE YEAR(HME)=" + PALIA
820     Gdb.Execute "delete from EGG WHERE YEAR(HME)=" + PALIA

        On Error Resume Next

830     DB_PALIA.Execute "ALTER TABLE TIM ADD  ID_NUM2 int "
840     DB_PALIA.Execute "UPDATE TIM SET ID_NUM2 = ID_NUM"
850     DB_PALIA.Execute "ALTER TABLE TIM DROP COLUMN ID_NUM "
860     DB_PALIA.Execute "ALTER TABLE TIM ADD  ID_NUM int IDENTITY(1,1)"

870     MsgBox "‘≈Àœ” Ã≈‘¡÷œ—¡”"

        Exit Sub

        '
        '
        '
        '
        '
        '7776
        '
        '' ===========================================================================
        '' --------------------- Ã≈‘¡÷œ—¡ ’–œÀœ…–ŸÕ ----------------------------------
        'Dim arxeio
        'Dim A As Long
        'Gdb.Close
        'Gdb.Open gConnect + ";DATABASE=" + CPALIA
        'A = GetTickCount()
        'Me.Caption = "ENHME—Ÿ”« Ã≈ ‘≈À≈’‘¡…¡ ’–œÀœ…–¡"
        'ALLAGH.Caption = "3.Ã≈‘¡÷œ—¡ ’–œÀœ…–ŸÕ ”≈ Õ≈¡ ◊—œÕ…¡"
        'UPDATE_YPOLOIPA2
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        'rSQL.Open "SELECT * FROM EID ", Gdb, adOpenDynamic, adLockOptimistic
        '
        '
        '
        '
        'CA = InputBox("ƒ˛ÛÂ Í˘‰ÈÍ¸ ·Ò·ÛÙ·ÙÈÍÔ˝", , "Î")
        '
        '
        'DB_PALIA.Open gConnect
        '
        '
        'DB_PALIA.Execute "DELETE  FROM EGGTIM WHERE LEFT(ATIM,1)='" + CA + "';"
        '
        ''Dim K As Long
        'rSQL.MoveFirst
        'K = 0
        'Do While Not rSQL.EOF
        '   If Not IsNull(rSQL("POS01")) Then
        '     If rSQL("POS01") <> 0 Then
        '       K = K + 1
        '       DB_PALIA.Execute "INSERT INTO EGGTIM (EIDOS,PELKOD,HME,XRE,KODE,ATIM,APOT) VALUES ('e','0000','01/01/" + TREX + "'," + Str(rSQL("POS01")) + ",'" + rSQL("KOD") + "','" + CA + Format(K, "00000") + "',1)"
        '     End If
        '
        '   End If
        '   '   rSQL.MoveNext
        '   If Not IsNull(rSQL("POS02")) Then
        '     If rSQL("POS02") <> 0 Then
        '        K = K + 1
        '        DB_PALIA.Execute "INSERT INTO EGGTIM (EIDOS,PELKOD,HME,XRE,KODE,ATIM,APOT) VALUES ('e','0000','01/01/" + TREX + "'," + Str(rSQL("POS02")) + ",'" + rSQL("KOD") + "','" + CA + Format(K, "00000") + "',2)"
        '     End If
        '   End If
        '   rSQL.MoveNext
        '
        '
        '
        '
        '   If K Mod 10 = 0 Then DoEvents: Me.Caption = Str(K)
        'Loop
        ''  Gdb.Execute "UPDATE EID SET APOS=0"

        '--------------------------------------------------------------------------------
        '
        '7777
        'update_pel
        '
        'rSQL.Open "SELECT * FROM PEL ", Gdb, adOpenDynamic, adLockOptimistic
        '
        ''Dim CA As String
        'Dim SSA As String
        '
        ''Dim K As Long
        'rSQL.MoveFirst
        'K = 0
        'Do While Not rSQL.EOF
        '   If Not IsNull(rSQL("TYP")) Then
        '     If rSQL("TYP") <> 0 Then
        '       K = K + 1
        '       DB_PALIA.Execute "UPDATE PEL SET AYP=" + Str(rSQL("TYP")) + "  WHERE KOD='" + rSQL("KOD") + "' AND EIDOS='" + rSQL("EIDOS") + "'"
        '
        '
        ''        SSA = "insert into EGG (HME,ATIM,APA,AIT,EIDOS,KOD,XPI,XREOSI,PISTOSI) values ("
        ''        SSA = SSA + "'" + Format(Now, "01/01/" + TREX) + "',"
        ''        SSA = SSA + "'Î" + Format(K, "00000") + "'," ' ATIM
        ''        SSA = SSA + Format(K, "00000") + "," 'APA
        ''        SSA = SSA + "'" + "¡–œ√—¡÷« + " ',"  'AIT
        ''        SSA = SSA + "'" + rSQL("Eidos") + "'," ' EIDOS='E'
        ''        SSA = SSA + "'" + rSQL("kod") + "'," ' KOD
        ''        SSA = SSA + "'" + "◊" + "'," ' XPI
        ''        SSA = SSA + Str(rSQL("TYP")) + "," 'XREOSI
        ''        SSA = SSA + Str(0) + ")" ' PISTOSI
        ''        Gdb.Execute SSA
        '
        '
        '     End If
        '
        '   End If
        '   rSQL.MoveNext
        '
        '
        '
        '
        '   If K Mod 10 = 0 Then DoEvents: Me.Caption = Str(K)
        'Loop
        ''  Gdb.Execute
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        'Gdb.Close
        'Gdb.Open gConnect

        Exit Sub

LATOS:
880     LATOS = 1

890     Resume Next

900     Return
LATOS2:
910     MsgBox "À¡»œ” œÕœÃ¡ ¬¡”«” ACCESS  ƒŸ”‘≈ –.◊.  C:\BACK.MDB"

        '<EhFooter>
        Exit Sub

Command14_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Command14_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Command14_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Sub update_pel()
''ypologismos ypoloipon
'    Dim X As String, order As String
'    Dim k As Integer, R As Recordset    ', sum As Single
'    Dim mBUFF
'    Dim execX, file
'    'If Combo1.Text = Combo1.List(0) Then mBUFF = "e" Else mBUFF = "r"
'
'    F_RUN = 1
'
'    'Data1.DatabaseName = gDir
'
'
'    On Error Resume Next
'
'
'    'On Error GoTo 0
'    ' ’–œÀœ√…∆Ÿ ‘¡ ‘≈À≈’‘¡…A ’–œÀœ…–¡ ¡–œ ‘œ EGG
'    Gdb.Execute "DROP TABLE EGG1"
'
'
'    Dim L As Long
'    Gdb.Execute "UPDATE PEL  SET TYP=0,XRESYN=0,PISSYN=0 FROM" _
'              & " PEL ", L
'
'
'    Gdb.Execute "SELECT UPPER(EIDOS) as MEIDOS,KOD,SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END ) AS SX," _
'              & " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END ) AS SP INTO EGG1" _
'              & " FROM EGG  GROUP BY UPPER(EIDOS),KOD "
'
'
'
'
'    'Gdb.Execute "CREATE VIEW dbo.[UPD_YPOL_PEL] as " _
'     '& " SELECT    dbo.PEL.XRESYN,dbo.PEL.PISSYN,dbo.PEL.TYP, dbo.PEL.KOD, dbo.PEL.EIDOS, dbo.EGG1.SP, dbo.EGG1.SX, dbo.PEL.AYP" _
'     '& "  FROM         dbo.PEL INNER JOIN" _
'     '& "  dbo.EGG1 ON dbo.PEL.EIDOS = dbo.EGG1.EIDOS AND dbo.PEL.KOD = dbo.EGG1.KOD"
'
'
'
'
'
'
'    ' If gXEIRISTHS = 8 Then
'    Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" _
'              & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" _
'              & " PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND UPPER(PEL.EIDOS)=EGG1.MEIDOS", L
'    ' Else
'    '  Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" _
'       '   & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" _
'       '   & " PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.EIDOS", L
'    ' End If
'
'    'Else
'    '   On Error GoTo 0
'    Gdb.Execute "UPDATE PEL  SET TYP=AYP-(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )+" _
'              & " (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" _
'              & " PEL INNER JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.EIDOS"
'
'    'End If
'
'
'End Sub

'Private Sub make_backup()
'    Dim k As Integer
'    Dim R As New ADODB.Recordset
'    Dim mkod As String
'    Dim DUM
'    Dim DD
'
'
'    DD = "C:\MERCVB\BACK" + Format(Day(Now), "00") + ".MDB"
'    Dim OK As Boolean
'    OK = False
'
'    'Dim yparxei_palio As Boolean
'    'yparxei_palio = False
'    If Len(Dir(DD, vbNormal)) > 0 Then
'        If Format(FileDateTime(DD), "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then
'            Exit Sub
'        Else
'            Kill DD
'            'yparxei_palio = True
'        End If
'
'    End If
'    Dim LISTA(20)
'    LISTA(1) = "PEL"
'    LISTA(2) = "EID"
'    LISTA(3) = "EGG"
'    LISTA(4) = "EGGTIM"
'    LISTA(5) = "TIM"
'    LISTA(6) = "GRA"
'    LISTA(7) = "MEM"
'    LISTA(8) = "PINAKES"
'    LISTA(9) = "XREOPIS"
'    LISTA(10) = "ARITMISI"
'    LISTA(11) = "PARASTAT"
'    LISTA(12) = "BARCODES"
'    LISTA(13) = "HMEROL"
'    LISTA(14) = "HMEROL2"
'    LISTA(15) = "PARAMETROI"
'    LISTA(16) = "CalendarEvents"
'    LISTA(17) = "CalendarRecurrencePatterns"
'    LISTA(18) = "ETAIREIES"
'
'
'    On Error Resume Next
'    'Kill DD
'
'    On Error GoTo LATOS2
'    Gdb.Execute "CREATE DATABASE " + DD
'    On Error Resume Next
'
'    Dim db As Database
'    Set db = OpenDatabase("", False, False, gConnect)
'    Me.MousePointer = vbHourglass
'    Dim LATOS
'
'    Me.Caption = "·Ò·Í·Î˛ ÂÒÈÏ›ÌÂÙÂ - ‰È·‰ÈÍ·Ûﬂ· ·ÌÙÈ„Ò‹ˆ˘Ì"
'
'    For k = 1 To 18   ' List1.ListCount - 1
'        LATOS = 0
'        On Error GoTo LATOS
'        DoEvents
'        db.Execute "SELECT * INTO " + LISTA(k) + "  IN '" + DD + "'  FROM  " + LISTA(k)
'    Next
'
'
'
'
'
'
'
'
'    Me.MousePointer = vbNormal
'    Exit Sub
'
'LATOS:
'    'List2.AddItem Err.Description
'    LATOS = 1
'    Resume Next
'    Return
'
'LATOS2:
'    MsgBox "À¡»œ” œÕœÃ¡ ¬¡”«” ACCESS  ƒŸ”‘≈ –.◊.  C:\BACK.MDB"
'
'
'End Sub

Private Sub Command15_Click()

        '<EhHeader>
        On Error GoTo Command15_Click_Err

        '</EhHeader>

        Dim DB    As Database

        Dim CA    As String

        Dim DB2   As New ADODB.Connection

        Dim PALIA As String

        Dim TREX  As String

        Dim ANS

        Dim arxeio

        Dim a As Long

        Dim CPALIA

        Dim sql  As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

        On Error GoTo OUTSIDE

100     TREX = InputBox("NEA” ˜ÒÔÌÈ‹", , Year(Date))

110     CPALIA = InputBox("¬‹ÛÁ  ·ÎÈ‹Ú ˜ÒÔÌÈ‹Ú .˜. D2007 ", , "D" + Format(Year(Date) - 1), "0000")

        '
        ' If TREX <= palia Or TREX - palia > 1 Then
        '    MsgBox "À¡»œ” ◊—œÕ…≈”. ≈Œœƒœ”"
        '    Exit Sub
        ' End If
        '

120     '   Gdb.Execute "update PEL SET AYP=0"

        On Error GoTo Command15_Click_Err

        ' CPALIA = "D" + Format(palia, "0000")

130     Gdb.Close

140     Gdb.Open gConnect + ";DATABASE=" + CPALIA

150     DB2.Open gConnect

        ' ===========================================================================
        ' --------------------- Ã≈‘¡÷œ—¡ ’–œÀœ…–ŸÕ ----------------------------------TelHmeAgoras()

160     Gdb.Close
170     Gdb.Open gConnect + ";DATABASE=" + CPALIA
180     a = GetTickCount()
190     Me.Caption = "ENHME—Ÿ”« Ã≈ ‘≈À≈’‘¡…¡ ’–œÀœ…–¡"
200     ALLAGH.Caption = "3.Ã≈‘¡÷œ—¡ ’–œÀœ…–ŸÕ ”≈ Õ≈¡ ◊—œÕ…¡"

        TelHmeAgoras CDate("01/01/2000"), Now() ' ÙÂÎÂıÙ·È· ÙÈÏÁ ·„ÔÒ‹Ú

210     UPDATE_YPOLOIPA2
       
        
        
        



220     rSQL.Open "SELECT * FROM EID ", Gdb, adOpenDynamic, adLockOptimistic

230     CA = InputBox("ƒ˛ÛÂ Í˘‰ÈÍ¸ ·Ò·ÛÙ·ÙÈÍÔ˝", , "Î")

240     DB2.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=1 AND DAY(HME)=1 AND LEFT(ATIM,1)='" + CA + "';"

        'Dim K As Long
250     rSQL.MoveFirst
260     k = 0

Dim nA As Integer
Dim rSQL2 As New ADODB.Recordset
        rSQL2.Open "SELECT COUNT(*) FROM PINAKES WHERE TYPOS=4 ", Gdb, adOpenDynamic, adLockOptimistic
Dim Apot_count As Integer
Apot_count = 0
If rSQL2(0) > 4 Then
   Apot_count = rSQL2(0)
End If

rSQL2.Close

Dim RTemp As New ADODB.Recordset




270     Do While Not rSQL.EOF
           RTemp.Open "select count(*) from EID WHERE KOD='" + rSQL("KOD") + "'", DB2, adOpenDynamic, adLockOptimistic
           If RTemp(0) = 0 Then
               DB2.Execute "insert into EID (KOD,FPA,MON,ONO) VALUES ('" + rSQL!kod + "'," + str(rSQL!FPA) + ",'" + CNull(rSQL!MON) + "','" + CNull(rSQL!ono) + "')"
           End If
           RTemp.Close
           
           
           

          If Apot_count <= 4 Then ' ≈ŸÚ 4 ¡–œ»« ≈Ú

280         If Not IsNull(rSQL("POS01")) Then
290             If rSQL("POS01") <> 0 Then
300                 k = k + 1
310                 DB2.Execute "INSERT INTO EGGTIM (EIDOS,PELKOD,HME,XRE,KODE,ATIM,APOT,TIMM) VALUES ('e','0000','01/01/" + TREX + "'," + str(rSQL("POS01")) + ",'" + rSQL("KOD") + "','" + CA + Format(k, "00000") + "',1," + Replace(Format(rSQL("xti"), "######0.00000"), ",", ".") + ")"
                End If

            End If

            '   rSQL.MoveNext
320         If Not IsNull(rSQL("POS02")) Then
330             If rSQL("POS02") <> 0 Then
340                 k = k + 1
350                 DB2.Execute "INSERT INTO EGGTIM (EIDOS,PELKOD,HME,XRE,KODE,ATIM,APOT,TIMM) VALUES ('e','0000','01/01/" + TREX + "'," + str(rSQL("POS02")) + ",'" + rSQL("KOD") + "','" + CA + Format(k, "00000") + "',2," + Replace(Format(rSQL("xti"), "######0.00000"), ",", ".") + ")"
                End If
            End If

            '   rSQL.MoveNext
360         If Not IsNull(rSQL("POS03")) Then
370             If rSQL("POS03") <> 0 Then
380                 k = k + 1
390                 DB2.Execute "INSERT INTO EGGTIM (EIDOS,PELKOD,HME,XRE,KODE,ATIM,APOT,TIMM) VALUES ('e','0000','01/01/" + TREX + "'," + str(rSQL("POS03")) + ",'" + rSQL("KOD") + "','" + CA + Format(k, "00000") + "',3," + Replace(Format(rSQL("xti"), "######0.00000"), ",", ".") + ")"
                End If
            End If

            '   rSQL.MoveNext
400         If Not IsNull(rSQL("POS04")) Then
410             If rSQL("POS04") <> 0 Then
420                 k = k + 1
430                 DB2.Execute "INSERT INTO EGGTIM (EIDOS,PELKOD,HME,XRE,KODE,ATIM,APOT,TIMM) VALUES ('e','0000','01/01/" + TREX + "'," + str(rSQL("POS04")) + ",'" + rSQL("KOD") + "','" + CA + Format(k, "00000") + "',4," + Replace(Format(rSQL("xti"), "######0.00000"), ",", ".") + ")"
                End If
            End If
            
            
          Else
          
          
          
             For nA = 1 To Apot_count
               rSQL2.Open " SELECT  SUM(CASE APOT  WHEN " + str(nA) + "  THEN XRE ELSE 0  END )- SUM(CASE APOT  WHEN " + str(nA) + "  THEN PIS ELSE 0  END ) AS S1 from EGGTIM WHERE KODE='" + rSQL("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
               
               If rSQL2(0) <> 0 Then
                  k = k + 1
                  DB2.Execute "INSERT INTO EGGTIM (EIDOS,PELKOD,HME,XRE,KODE,ATIM,APOT,TIMM) VALUES ('e','0000','01/01/" + TREX + "'," + str(rSQL2(0)) + ",'" + rSQL("KOD") + "','" + CA + Format(k, "00000") + "'," + str(nA) + "," + Replace(Format(rSQL("xti"), "######0.00000"), ",", ".") + ")"
               
               End If
               rSQL2.Close
               
          
          
             Next
          
          
          
          
          End If
         
            
            

440         rSQL.MoveNext

450         If k Mod 10 = 0 Then DoEvents: Me.Caption = str(k)
        Loop

        '  Gdb.Execute "UPDATE EID SET APOS=0"

460     DB2.Close

470     Gdb.Close
480     Gdb.Open gConnect

490     MsgBox "œÎÔÍÎÁÒ˛ËÁÍÂ Á ÏÂÙ·ˆÔÒ‹ ÂÈ‰˛Ì"

        Exit Sub

OUTSIDE:
500     MsgBox "ƒ≈Õ œÀœ À«—Ÿ»« ≈ « Ã≈‘¡÷œ—¡ ’–œÀœ…–œ’ ≈…ƒŸÕ"

        Exit Sub

LATOS:
510     LATOS = 1

520     Resume Next

530     Return
LATOS2:
540     MsgBox "À¡»œ” œÕœÃ¡ ¬¡”«” ACCESS  ƒŸ”‘≈ –.◊.  C:\BACK.MDB"

        '<EhFooter>
        Exit Sub

Command15_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Command15_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Command15_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command16_Click()

        ' BOHU82.SHOW
        '<EhHeader>
        On Error GoTo Command16_Click_Err

        '</EhHeader>
        Dim k

        Dim c As String

100     c = InputBox("ƒŸ”≈ œÕœÃ¡ BACKUP ‘«” ¬¡”«” " + Gdb.DefaultDatabase, , "C:\MERCVB\BACK1.bak")
110     Gdb.Execute "BACKUP DATABASE [" + Gdb.DefaultDatabase + "] TO  DISK ='" + c + "' WITH NOFORMAT, NOINIT, SKIP, NOREWIND, NOUNLOAD,  STATS = 10", k
120     MsgBox "œÀœ À«—Ÿ»« ≈"

        '<EhFooter>
        Exit Sub

Command16_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Command16_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Command16_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim k As Integer

        Dim R As New ADODB.Recordset

100     ANS = MsgBox("–ÒÔÛÔ˜ﬁ Ë· Û‚ÁÛÙÔ˝Ì Ù· ·Ò˜Âﬂ· ÙÔı SERVER. ≈ﬂÛ·È Ûﬂ„ÔıÒÔÚ;", vbYesNo)

110     If ANS = vbNo Then Exit Sub

        Dim MKOD As String

120     If Not ANS = vbYes Then
130         MsgBox " ¡ ’—ø»« ≈ « ≈—√¡”∫¡"

            Exit Sub

        End If

140     MKOD = InputBox("¡Ì ÂﬂÛ·È Ûﬂ„ÔıÒÔÚ ‰˛ÛÂ Í˘‰ÈÍ¸ ")

150     If Val(MKOD) = 12345 Then
            'ok
        Else

            Exit Sub

        End If

160     R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

170     F_xre_APOU = ""
180     F_axre_APOU = ""
190     F_pis_APOU = ""
200     F_apis_APOU = ""

210     R.MoveFirst

220     Do While Not R.EOF

230         If R("POS_APOU") = 1 Then
240             F_xre_APOU = F_xre_APOU + R("eidos")
250         ElseIf R("POS_APOU") = 2 Then
260             F_axre_APOU = F_axre_APOU + R("eidos")
270         ElseIf R("POS_APOU") = 3 Then
280             F_pis_APOU = F_pis_APOU + R("eidos")
290         ElseIf R("POS_APOU") = 4 Then
300             F_apis_APOU = F_apis_APOU + R("eidos")
            End If

310         R.MoveNext
        Loop

320     R.Close

330     For k = 0 To List1.ListCount - 1

340         If List1.Selected(k) Then
350             update_sql_from_dbf List1.List(k)
            End If

        Next

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        'Dim Gdb As New ADODB.Connection
        'Dim GR As New ADODB.Recordset
        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>

        Dim k As Integer

        Dim R As New ADODB.Recordset

100     If InStr(Text1.Text, "MDB") > 0 Then
            'OK
        Else
110         MsgBox "ƒ«ÀŸ”‘≈ ≈Õ¡ œÕœÃ¡ Ãœ—÷«” ACCESS –.◊. C:\BACK1.MDB"
120         Text1.SelStart = 0
130         Text1.SelLength = Len(Text1.Text)
140         Text1.SetFocus

            Exit Sub

        End If

150     ANS = MsgBox("–ÒÔÛÔ˜ﬁ Ë· Û‚ÁÛÙÔ˝Ì Ù· ·Ò˜Âﬂ· ÙÔı ˆ·Í›ÎÔı " + text1text + "  ≈ﬂÛ·È Ûﬂ„ÔıÒÔÚ;", vbYesNo)

160     If ANS = vbNo Then Exit Sub

        Dim MKOD As String

170     If Not ANS = vbYes Then
180         MsgBox " ¡ ’—ø»« ≈ « ≈—√¡”∫¡"

            Exit Sub

        End If

190     MKOD = InputBox("¡Ì ÂﬂÛ·È Ûﬂ„ÔıÒÔÚ ‰˛ÛÂ Í˘‰ÈÍ¸ ")

200     If Val(MKOD) = 12345 Then
            'ok
        Else

            Exit Sub

        End If

        Dim DUM

210     DD = Text1.Text

        On Error Resume Next

220     Kill DD

        On Error GoTo LATOS2

230     DUM = CREATE_DATABASE(DD)

        On Error Resume Next

        Dim DB As Database

240     Set DB = OpenDatabase("", False, False, gConnect)
250     Me.MousePointer = vbHourglass

260     List2.Visible = True
        'Gdb.Open "DSN=2004" 1.058
270     List2.Clear

280     For k = 0 To List1.ListCount - 1
290         LATOS = 0

            On Error GoTo LATOS

300         DoEvents
310         Me.Caption = List1.List(k)

320         If List1.Selected(k) Then
                ' update_sql_from_dbf List1.List(K)
                ' Kill Text1.Text + "\" + List1.List(K) + ".dbf"
330             DB.Execute "SELECT * INTO " + List1.List(k) + "  IN '" + DD + "'  FROM  " + List1.List(k), ANS

340             If LATOS = 0 Then
350                 List2.AddItem "œ     " + List1.List(k)
                Else
360                 List2.AddItem "À¡»œ” " + List1.List(k)
                End If

            End If

        Next

370     Me.MousePointer = vbNormal

        Exit Sub

LATOS:
380     List2.AddItem Err.Description
390     LATOS = 1

400     Resume Next

410     Return

LATOS2:
420     MsgBox "À¡»œ” œÕœÃ¡ ¬¡”«” ACCESS  ƒŸ”‘≈ –.◊.  C:\BACK.MDB"

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

Gdb.Close
Gdb.Open gConnect + ";DATABASE=MASTER"

'sql = "RESTORE DATABASE TEST   FROM DISK = 'C:\MERCVB\BACK1121.BAK' WITH REPLACE;  With Move 'MERCURY' TO 'C:\MERCVB\TEST.MDF', Move 'MERCURY_LOG' TO 'C:\MERCVB\TEST.LDF' "


sql = "RESTORE DATABASE [MERCURY] From DISK =  N'C:\Mercvb\BACK99.bak' WITH  FILE = 1, NOUNLOAD,REPLACE,STATS = 10 "
Gdb.Execute sql

Gdb.Close

Gdb.Open gConnect


End Sub

'Private Sub Command4_Click()
'    Dim sql As New ADODB.Connection
'    Dim rDBF As New ADODB.Recordset
'    Dim rSQL As New ADODB.Recordset
'
'    rSQL.Open "SELECT * FROM EID WHERE APOS>0", Gdb, adOpenDynamic, adLockOptimistic
'
'
'
'
'    a = InputBox("ƒ˛ÛÂ Í˘‰ÈÍ¸ ·Ò·ÛÙ·ÙÈÍÔ˝", , "Î")
'
'    Gdb.Execute "DELETE  FROM EGGTIM WHERE LEFT(ATIM,1)='" + a + "';"
'
'    Dim k As Long
'    rSQL.MoveFirst
'    k = 0
'    Do While Not rSQL.EOF
'        If Not IsNull(rSQL("APOS")) Then
'            k = k + 1
'            Gdb.Execute "INSERT INTO EGGTIM (EIDOS,PELKOD,HME,XRE,KODE,ATIM) VALUES ('e','0000','01/01/2005'," + Str(rSQL("APOS")) + ",'" + rSQL("KOD") + "','" + a + Format(k, "00000") + "')"
'        End If
'        rSQL.MoveNext
'        If k Mod 10 = 0 Then DoEvents: Me.Caption = Str(k)
'    Loop
'    Gdb.Execute "UPDATE EID SET APOS=0"
'End Sub

'Private Sub Command5_Click()
'    Dim DBF As Database
'    Dim sql As New ADODB.Connection
'    Dim rDBF As Recordset
'    Dim rSQL As New ADODB.Recordset
'    Dim conDBF As String
'    Dim conSQL As String
'    Dim lo As Long
'    Dim r2 As DAO.Recordset
'
'    Set DBF = OpenDatabase(Text1.Text, False, False, "dBase III;")
'    Set r2 = DBF.OpenRecordset("select LEFT(memo_key,5) AS KLEIDI,COUNT(*) from memotext where left(memo_key,1) in ('e','r') group by  LEFT(memo_key,5) order by  LEFT(memo_key,5) ")
'
'    rSQL.Open "select *from PEL", Gdb, adOpenDynamic, adLockOptimistic
'
'
'    Dim a, b, C, D, e, F, G
'
'    r2.MoveFirst
'    Do While Not r2.EOF
'        a = Left(r2(0), 5)
'        Set R = DBF.OpenRecordset("select *from memotext where left(memo_key,5)='" + a + "' order by memo_key ")
'        If R.RecordCount > 0 Then
'            Me.Caption = a
'            DoEvents
'            R.MoveFirst
'            b = ""
'            Do While Not R.EOF
'                b = b + R(1) + Chr$(13) + Chr$(10)    'memoline
'                R.MoveNext
'            Loop
'
'            If Not IsNull(b) Then
'                b = Replace(b, "'", "-")
'                If Not IsNull(b) Then Gdb.Execute "update PEL SET MEMO='" + b + "' WHERE EIDOS='" + Left(a, 1) + "' AND KOD='" + mID$(a, 2, 4) + "'"
'            End If
'            DoEvents
'        End If
'        r2.MoveNext
'    Loop
'
'
'
'
'
'
'
'
'End Sub

'Private Sub Command6_Click()
''F5 ¡–œ»« «Ú
'    Dim DBF As Database
'    Dim sql As New ADODB.Connection
'    Dim rDBF As Recordset
'    Dim rSQL As New ADODB.Recordset
'    Dim conDBF As String
'    Dim conSQL As String
'    Dim lo As Long
'
'    Set DBF = OpenDatabase(Text1.Text, False, False, "dBase III;")
'
'
'    rSQL.Open "select *from EID", Gdb, adOpenDynamic, adLockOptimistic
'
'
'    Dim a, b, C, D, e, F, G
'    G = 0
'    rSQL.MoveFirst
'    Do While Not rSQL.EOF
'        G = G + 1
'        a = "a" + rSQL!kod
'
'        Set R = DBF.OpenRecordset("select *from memotext where left(memo_key,15)='" + a + "' order by memo_key ")
'        If R.RecordCount > 0 Then
'
'            R.MoveFirst
'            b = ""
'            Do While Not R.EOF
'                b = b + R(1) + Chr$(13)    'memoline
'                R.MoveNext
'            Loop
'            rSQL!memo = b
'            rSQL.Update
'        End If
'        G = G + 1
'        If G Mod 20 = 0 Then
'            Me.Caption = a
'            DoEvents
'        End If
'        rSQL.MoveNext
'    Loop
'
'End Sub

'Private Sub Command7_Click()
''------ FALCON ----------
'
'    Dim sql As New ADODB.Connection
'    Dim rDBF As New ADODB.Recordset
'    Dim rSQL As New ADODB.Recordset
'
'    rSQL.Open "SELECT * FROM EID where LEFT(KOD,5) NOT IN ('91393')", Gdb, adOpenDynamic, adLockOptimistic
'
'    Dim db As Database, R As Recordset
'
'    Set db = OpenDatabase("C:\LAGEURO\MHXANAKI", False, False, "dBase III;")
'
'    On Error Resume Next
'
'    Kill "C:\LAGEURO\MHXANAKI\OLA.DBF"
'    On Error GoTo 0
'    db.Execute "SELECT TOP 1 OLAA.* INTO OLA FROM OLAA"
'    Set R = db.OpenRecordset("OLA")
'
'
'    Dim k As Long
'    rSQL.MoveFirst
'    k = 0
'    Do While Not rSQL.EOF
'        R.AddNew
'        R("BARCODE") = Left(rSQL("KOD"), 13)
'        R("LTI") = rSQL("LTI5")
'        '    R("XTI") = rSQL("XTI")
'        R("NAME") = Left(rSQL("ONO"), 25)
'        R("POS01") = rSQL("POS01")
'        R("POS02") = rSQL("POS02")
'
'        '  If Not IsNull(rSQL("POS01")) Then
'        '     K = 0
'        '  End If
'
'
'        R("PROM") = Left(rSQL("PROM"), R("PROM").Size)
'        R("XTI") = rSQL("XTI")
'        R("KODLOG") = rSQL("KODLOG")
'
'
'        R.Update
'        k = k + 1
'        If k Mod 100 = 0 Then
'            Me.Caption = k
'            DoEvents
'        End If
'
'        'Me.Caption = rSQL("kod")
'
'        rSQL.MoveNext
'
'    Loop
'    Gdb.Execute "UPDATE EID SET APOS=0"
'
'    Dim DUM
'
'    DUM = Shell("C:\LAGEURO\MHXANAKI\send.bat", vbHide)
'
'
'
'
'
'
'
'
'
'
'End Sub

'Private Sub Command8_Click()
'' METAFORA SE ARXEIO DBF
'    Dim arxeio
'    Dim a As Long
'    a = GetTickCount()
'    Me.Caption = "ENHME—Ÿ”« Ã≈ ‘≈À≈’‘¡…¡ ’–œÀœ…–¡"
'
'    UPDATE_YPOLOIPA2
'
'    arxeio = "EID"
'
'
'
'    Dim DBF As Database
'    Dim sql As New ADODB.Connection
'    Dim rDBF As Recordset
'    Dim rSQL As New ADODB.Recordset
'
'    Dim conSQL As String
'    Dim k As Long
'    Dim Fname As String
'
'
'    sql.Open gConnect
'
'
'    Dim db As DAO.Database
'    Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
'    On Error Resume Next
'    Kill "C:\LAGEURO\EID.DBF"
'    On Error GoTo 0
'
'    db.Execute "SELECT TOP 1 EID0.* INTO EID  FROM EID0"
'
'
'    Set rDBF = db.OpenRecordset("EID")
'
'
'
'
'
'    'Label1.Caption = ARXEIO + " " + Format(rDBF.RecordCount, "######")
'
'    'SQL.Execute "DELETE FROM " + arxeio
'
'    rSQL.Open "SELECT * FROM EID ", Gdb, adOpenForwardOnly, adLockReadOnly
'
'    ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic
'
'    rSQL.MoveFirst
'
'    Dim Z
'    Z = 0
'    On Error GoTo WRITEERROR    ' Resume Next
'
'    Do While Not rSQL.EOF
'
'        rDBF.AddNew
'        rDBF("KOD") = rSQL("KOD")
'        rDBF("ONO") = rSQL("ONO")
'        rDBF("ERG") = rSQL("ERG")
'        rDBF("LTI") = rSQL("LTI")
'        rDBF("LTI5") = rSQL("LTI5")
'        rDBF("pontoi") = rSQL("pontoi")
'        rDBF("FPA") = rSQL("FPA")
'        rDBF("POS01") = rSQL("POS01")
'        rDBF("POS02") = rSQL("POS02")
'
'        rDBF("prom") = rSQL("prom")
'        rDBF("xti") = rSQL("xti")
'        rDBF("KODLOG") = rSQL("KODLOG")
'
'
'
'        rDBF("pOS_KERD") = rSQL("pOS_KERD")
'        rDBF("POS_KERD2") = rSQL("POS_KERD2")
'        rDBF("SPA") = rSQL("SPA")
'        rDBF("MON") = rSQL("MON")
'
'
'
'
'
'
'
'
'
'        rDBF.Update
'
'        Z = Z + 1
'        If Z Mod 10 = 0 Then
'            Me.Caption = Z
'        End If
'        rSQL.MoveNext
'        DoEvents
'
'    Loop
'
'    rSQL.Close
'    rDBF.Close
'
'    Me.Caption = (GetTickCount() - a) / 1000
'
'
'
'
'    Exit Sub
'
'
'WRITEERROR:
'    List2.Visible = True
'
'    List2.AddItem Err.Description
'    Resume Next
'
'
'
'
'
'
'
'End Sub

'Private Sub Command9_Click()
'' apo SQL SERVER ”≈ DBF ¡—◊≈…¡ –œ’ «ƒ« ’–¡—◊œ’Õ
'    Dim k As Integer
'    Dim R As New ADODB.Recordset
'    ANS = MsgBox("–ÒÔÛÔ˜ﬁ Ë· Û‚ÁÛÙÔ˝Ì Ù· ·Ò˜Âﬂ· DBF. ≈ﬂÛ·È Ûﬂ„ÔıÒÔÚ;", vbYesNo)
'    If ANS = vbNo Then Exit Sub
'    Dim mkod As Integer
'    mkod = InputBox("¡Ì ÂﬂÛ·È Ûﬂ„ÔıÒÔÚ ‰˛ÛÂ Í˘‰ÈÍ¸ ")
'    If Val(mkod) = 12345 Then
'        'ok
'    Else
'        Exit Sub
'    End If
'
'
'
'    R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
'
'
'    F_xre_APOU = ""
'    F_axre_APOU = ""
'    F_pis_APOU = ""
'    F_apis_APOU = ""
'
'    R.MoveFirst
'    Do While Not R.EOF
'        If R("POS_APOU") = 1 Then
'            F_xre_APOU = F_xre_APOU + R("eidos")
'        ElseIf R("POS_APOU") = 2 Then
'            F_axre_APOU = F_axre_APOU + R("eidos")
'        ElseIf R("POS_APOU") = 3 Then
'            F_pis_APOU = F_pis_APOU + R("eidos")
'        ElseIf R("POS_APOU") = 4 Then
'            F_apis_APOU = F_apis_APOU + R("eidos")
'        End If
'
'        R.MoveNext
'    Loop
'    R.Close
'
'
'    For k = 0 To List1.ListCount - 1
'        If List1.Selected(k) Then
'            update_DBF_from_SQL List1.List(k)
'        End If
'
'    Next
'
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command10_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "BOHU8"
110     PARAMETROI.SHOW 1

        'F_TAB = Val(FindParametroi(1,"APOT2", "F_TAB", "3", "”Â ÔÈ¸ TAB ÂﬂÌ·È ÛÙ·Ï·ÙÁÏ›ÌÔ"))
        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()
Dim f_gfk As Long
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl
110     Me.KeyPreview = True
120     Me.Picture = LoadPicture(gPicture)

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

130     f_psifia_atim = Val(FINDPARAMETROI(1, "PAR1", "F_PSIFIA_ATIM", "5", "√. 9. ÿ«÷…¡ ¡—…»Ãœ’ ‘…ÃœÀœ√…œ’ (5)"))
140     F_PATHFORT = FINDPARAMETROI(1, "BOHU8", "F_PATHFORT", "C:\LAGEURO\FORTHGO.MDB", "ƒ…¡ƒ—œÃ«  ¡… œÕœÃ¡  ¡Õ‘…√—¡÷ŸÕ ÷œ—‘«√œ’")   'posa psifia tha exei h kathe seira
f_gfk = FINDPARAMETROI(1, "BOHU8", "F_GFK", 0, "0=OXI ¡ÀÀœ…Ÿ”  Ÿƒ… œ”")
150     List1.AddItem "PEL"
160     List1.AddItem "EID"
170     List1.AddItem "EGG"
180     List1.AddItem "EGGTIM"
190     List1.AddItem "TIM"
200     List1.AddItem "GRA"
210     List1.AddItem "MEM"
220     List1.AddItem "PINAKES"
230     List1.AddItem "XREOPIS"
240     List1.AddItem "ARITMISI"
250     List1.AddItem "PARASTAT"
260     List1.AddItem "BARCODES"
270     List1.AddItem "HMEROL"
280     List1.AddItem "HMEROL2"
290     List1.AddItem "PARAMETROI"
300     List1.AddItem "TIMOKAT"
310     List1.AddItem "CalendarEvents"
320     List1.AddItem "CalendarRecurrencePatterns"

330     If Len(Dir("C:\KENTRIKO.TXT")) > 0 Then
340         Command12.Enabled = False
        End If

350     F_PATH = FINDPARAMETROI(1, "BOHU8", "F_PATH", "C:\MERCVB", "ƒ…¡ƒ—œÃ« ¡Õ‘…√—¡÷ŸÕ")   'posa psifia tha exei h kathe seira
360     F_BACKUPTYPE = FINDPARAMETROI(1, "BOHU8", "F_BACKUPTYPE", 0, "‘’–œ” BACKUP 0=ACCESS 1=SQL DATABASE")   'posa psifia tha exei h kathe seira

370     If F_BACKUPTYPE = 1 Then
380         Text1.Text = "C:\mercVB\back1.BAK"
        End If

        On Error Resume Next

390     File1.Path = "C:\mercvb\queries"

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub update_DBF_from_SQL(arxeio As String)
'
'    Dim DBF As Database
'    Dim sql As New ADODB.Connection
'    Dim rDBF As Recordset
'    Dim rSQL As New ADODB.Recordset
'    Dim conDBF As String
'    Dim conSQL As String
'    Dim k As Long
'    Dim Fname As String
'
'    Dim db As DAO.Database
'    Set db = OpenDatabase(Text1.Text, False, False, "dBase III;")
'    db.Execute "delete * from " + arxeio
'    Set rDBF = db.OpenRecordset(arxeio)
'
'    Label1.Caption = arxeio + " " + Format(rDBF.RecordCount, "######")
'
'
'
'    rSQL.Open "SELECT *FROM " + arxeio, sql, adOpenDynamic, adLockOptimistic
'    ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic
'
'    rSQL.MoveFirst
'
'    Dim Z
'    Z = 0
'    On Error GoTo WRITEERROR    ' Resume Next
'
'    Do While Not rSQL.EOF
'
'        rDBF.AddNew
'
'        For k = 0 To rDBF.FIELDS.Count - 1
'            Fname = rDBF.FIELDS(k).Name    ' p.x. FNAME=epo    R(0).NAME
'
'            '      If IsNull(rDBF(k)) Then
'            '      Else
'            rDBF(Fname) = rSQL(Fname)    ' rsqk("epo")=r(0)
'            '      End If
'        Next
'        rDBF.Update
'        Z = Z + 1
'        If Z Mod 10 = 0 Then
'            Me.Caption = Z
'        End If
'        rSQL.MoveNext
'        DoEvents
'
'    Loop
'
'    rSQL.Close
'    rDBF.Close
'    Exit Sub
'
'
'WRITEERROR:
'    List2.Visible = True
'
'    List2.AddItem Err.Description
'    Resume Next
'
'
'
'
'End Sub

'Private Sub TDBGrid_GroupColMove(ByVal Position As Integer, ByVal ColIndex As Integer, Cancel As Integer)
'    Dim strSort As String
'
'    Dim Col As TrueOleDBGrid80.Column
'
'
'
'    ' Loop through GroupColumns collection and construct
'
'    ' the sort string for the Sort property of the Recordset
'
'    For Each Col In TDBGrid.GroupColumns
'
'        If strSort <> vbNullString Then
'
'            strSort = strSort & ", "
'
'        End If
'
'        strSort = strSort & "[" & Col.DataField & "]"
'
'    Next Col
'
'
'
'    TDBGrid.HoldFields
'
'    Adodc2.Recordset.Sort = strSort
'
'
'End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

'ResizeFormFor Me

    mForm_Resize Me, 8, 13170, 9570, fh, fw, ft, fl

End Sub

Private Sub fortosh_Click()

        '<EhHeader>
        On Error GoTo fortosh_Click_Err

        '</EhHeader>
100     If Len(Dir("C:\KENTRIKO.TXT", vbNormal)) > 0 Then
110         MsgBox "≈…”‘≈ ”≈ À¡»œ” ’–œÀœ√…”‘«"

            Exit Sub

        End If

        Dim ANS

120     ANS = MsgBox("»¡ ƒ…¡√—¡÷œ’Õ ‘¡ ¡—◊≈…¡ ‘œ’ «/’  ¡… »¡ √…Õ≈… ≈…”¡√Ÿ√« ¡–œ ‘œ  ≈Õ‘—… œ", vbYesNo)

130     If ANS = vbNo Then

            Exit Sub

        End If

        '==============================================================================

        Dim DB As Database

140     Set DB = OpenDatabase("", False, False, gConnect)

        Dim RR As Recordset

        On Error GoTo fortosh_Click_Err

150     Me.MousePointer = vbHourglass

160     List2.Visible = True
        'Gdb.Open "DSN=2004" 1.058
170     List2.Clear

        Dim DD As String

180     DD = F_PATHFORT  ' "C:\LAGEURO\FORTHGO.MDB"

        On Error Resume Next

190     Gdb.Execute "DELETE FROM PEL"
200     DB.Execute "INSERT INTO PEL SELECT *  FROM  PEL IN '" + DD + "'"

        On Error GoTo fortosh_Click_Err

210     Gdb.Execute "DELETE FROM EGG"
220     DB.Execute "INSERT INTO EGG SELECT *  FROM  EGG IN '" + DD + "'"

230     Gdb.Execute "DELETE FROM EID"
240     DB.Execute "INSERT INTO EID SELECT *  FROM  EID IN '" + DD + "'"

250     Gdb.Execute "DELETE FROM BARCODES"
260     DB.Execute "INSERT INTO BARCODES SELECT *  FROM  BARCODES  IN '" + DD + "'"

270     Gdb.Execute "DELETE FROM EGGTIM "
280     DB.Execute "INSERT INTO EGGTIM  SELECT *  FROM  EGGTIM IN '" + DD + "'"

290     MsgBox "‘≈Àœ” Ã≈‘¡÷œ—¡”"
300     Me.MousePointer = vbNormal

        '<EhFooter>
        Exit Sub

fortosh_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.fortosh_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.fortosh_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub List1_DblClick()

        '<EhHeader>
        On Error GoTo List1_DblClick_Err

        '</EhHeader>
        Dim k As Integer

100     For k = 0 To List1.ListCount - 1
110         List1.Selected(k) = True
        Next

        '<EhFooter>
        Exit Sub

List1_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.List1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.List1_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub NeaBash_Click()

       Dim c As String
        Dim TREX As String, PALIA As String
        
        Dim MKOD As String
        
140     MKOD = InputBox("¡Ì ÂﬂÛ·È Ûﬂ„ÔıÒÔÚ ‰˛ÛÂ Í˘‰ÈÍ¸ ")

150     If Val(MKOD) = 12345 Then
            'ok
        Else

            Exit Sub

        End If
        
        
        
        
        
        
        
170            ' TREX = InputBox("NEA” ˜ÒÔÌÈ‹", , Year(Date))

180     PALIA = InputBox("œÌÔÏ· Õ›·Ú ¬‹ÛÁÚ", , "")


        
Dim RR As New ADODB.Recordset
   
     RR.Open "SELECT  COUNT(*) FROM  sys.databases WHERE name='" + PALIA + "'", Gdb, adOpenDynamic, adLockOptimistic
    
   If RR(0) > 0 Then
      MsgBox "≈◊≈… «ƒ« ƒ«Ã…œ’—√«»≈… « ¬¡”« " + PALIA + " ¡ƒ’Õ¡‘« ”’Õ≈◊≈…¡"
      Exit Sub
   End If
   
     
  
        




         c = InputBox("ƒŸ”≈ œÕœÃ¡ BACKUP ‘«” ¬¡”«” " + Gdb.DefaultDatabase, , "C:\MERCVB\BACK" + Replace(time$, ":", "") + ".bak")
Gdb.CommandTimeout = 0

         Gdb.Execute "BACKUP DATABASE [" + Gdb.DefaultDatabase + "] TO  DISK ='" + c + "' WITH NOFORMAT, NOINIT, SKIP, NOREWIND, NOUNLOAD,  STATS = 10", k
         



Dim c2 As String
c2 = Gdb.DefaultDatabase
List3.AddItem "BACKUP OK"


Dim CN As String
CN = ""









'‰ÁÏÈÔıÒ„Âﬂ ÙÁÌ Ì›· ‚‹ÛÁ ant2 ·Ô ÙÔ backup  C:\mercvb\ant.bak
' Í·È Ù· ·Ò˜Âﬂ· mercury_data & mercury_log Ù· ·ÌÙÈ„Ò‹ˆÂÈ ÛÙ· C:\mercvb\ant2.mdf & C:\mercvb\ant2_log.ldf

'RESTORE DATABASE ant2 FROM DISK='C:\mercvb\ant.bak'
'WITH
'Move 'mercury_data' TO 'C:\mercvb\ant2.mdf',
'Move 'mercury_log' TO 'C:\mercvb\ant2_log.ldf'


'RESTORE DATABASE [emp2] FROM DISK = 'c:\mercvb\emp2'
'WITH
'Move 'EMP_Data' TO 'c:\MERCVB\EMP2_Data.mdf',
'Move 'EMP_Log' TO 'c:\MERCVB\EMP2_Log.ldf',
'RECOVERY, REPLACE, STATS = 10;






Dim log_name As String
log_name = InputBox("ƒŸ”≈ œÕœÃ¡ (LOGICAL) ‘«” ¬¡”«” " + Gdb.DefaultDatabase, , "MERCURY")


        Dim CPALIA

220     CPALIA = PALIA  '    "D" + Format(PALIA, "0000")


c2 = "RESTORE   DATABASE " + CPALIA + " FROM DISK='" + c + "'  WITH "
c2 = c2 + " Move '" + log_name + "_Data' TO 'C:\mercvb\" + CPALIA + "_Data.mdf',  "
c2 = c2 + " Move '" + log_name + "_Log' TO 'C:\mercvb\" + CPALIA + "_log.ldf',"
c2 = c2 + " RECOVERY, REPLACE, STATS = 10; "

On Error GoTo NOTFINISHED

Gdb.Execute c2
 Gdb.Execute "INSERT INTO ETAIREIES (TITLOS,DSN)  VALUES ('" + CPALIA + "','" + CPALIA + "')"

List3.AddItem "RESTORE OK"



    Dim sql  As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

        Dim CONNECT_PALIA As String

'240     If InStr(gConnect, "DATABASE") = 0 Then
         CONNECT_PALIA = gConnect + ";" + "DATABASE=" + CPALIA + ";"
 '       Else
'260         CONNECT_PALIA = Left(gConnect, InStr(gConnect, "DATABASE") - 1) + ";" + "DATABASE=" + CPALIA + ";"
 '       End If

        Dim N


        Dim GDBPALIA As New ADODB.Connection

        GDBPALIA.Open CONNECT_PALIA


       GDBPALIA.Execute "delete from EGG"
       GDBPALIA.Execute "delete from TIM"
       GDBPALIA.Execute "delete from EGGTIM"

Dim ANS As Integer
ANS = MsgBox("”ÙÁÌ Ì›· ¬‹ÛÁ " + CPALIA + " Õ· ‰È·„Ò·ˆÔıÌ Í·È ÔÈ ÂÎ‹ÙÂÚ;", vbYesNo)
If ANS = vbYes Then
     GDBPALIA.Execute "delete from PEL"
End If

ANS = MsgBox("”ÙÁÌ Ì›· ¬‹ÛÁ " + CPALIA + " Õ· ‰È·„Ò·ˆÔıÌ Í·È ‘¡ ≈…ƒ«;", vbYesNo)
If ANS = vbYes Then
     GDBPALIA.Execute "delete from EID"
End If



















Exit Sub

NOTFINISHED:

MsgBox " ƒ≈Õ œÀœ À«—Ÿ»« ≈ " + Err.Description

 




End Sub

Private Sub NEAXRONIA_Click()

        '<EhHeader>
        On Error GoTo NEAXRONIA_Click_Err

        '</EhHeader>

        Dim DB       As Database

        Dim CA       As String

        Dim DB_PALIA As New ADODB.Connection

        Dim PALIA    As String

        Dim TREX     As String

        Dim ANS

100     ANS = MsgBox("ÒÔÛÔ˜Á Ë· ÏÂÙ·ˆÂÒËÔ˝Ì Ù· ·Ò˜Âﬂ· ÛÙÁÌ ·ÎÈ‹ ˜ÒÔÌÈ‹" & Chr(13) + " Í·È Ë· ÏÁ‰ÂÌÈÛÙÂﬂ Á ÙÒ›˜ÔıÛ·" + Chr(13) & "≈ﬂÛÙÂ Ûﬂ„ÔıÒÔÈ;", vbYesNo)

110     If ANS = vbYes Then
        Else

            Exit Sub

        End If

120     ALLAGH.Caption = "1.ƒÈ·‰ÈÍ·Ûﬂ· BACKUP"

        Dim k

130     Gdb.Execute "BACKUP DATABASE [" + Gdb.DefaultDatabase + "] TO  DISK ='C:\MERCVB\BACK1.bak' WITH NOFORMAT, NOINIT, SKIP, NOREWIND, NOUNLOAD,  STATS = 10", k

        ' TO SBHNV GIA NA PARV FRESKO BACKUP
140     DD = "C:\MERCVB\BACK" + Format(Day(Now), "00") + ".MDB"

        On Error Resume Next

150     Kill DD
        ' Kill "C:\MERCVB\BACK" + Format(Day(Now), "00") + ".LDB"

160     MDIForm1.MakeBackup "C:\MERCVB"

170     TREX = InputBox("NEA” ˜ÒÔÌÈ‹", , Year(Date))

180     PALIA = InputBox("œÌÔÏ· ·ÎÈ‹Ú ˜ÒÔÌÈ‹Ú", , Year(Date) - 1)

190     If TREX <= PALIA Or TREX - PALIA > 1 Then
200         ANS = MsgBox("À¡»œ” ◊—œÕ…≈”. ≈…”‘≈ ”…√œ’—œ…;", vbYesNo)

210         If ANS = vbNo Then

                Exit Sub

            End If

        End If

        ' Gdb.Execute "create database " + palia
        'DD = Text1.Text
        On Error Resume Next

        '================================================

        On Error GoTo NEAXRONIA_Click_Err

        Dim CPALIA

220     CPALIA = "D" + Format(PALIA, "0000")

        Dim sql  As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

230     Gdb.Execute "create database " + CPALIA, k

        'Gdb.DefaultDatabase = CPALIA
        Dim CONNECT_PALIA As String

240     If InStr(gConnect, "DATABASE") = 0 Then
250         CONNECT_PALIA = gConnect + ";" + "DATABASE=" + CPALIA + ";"
        Else
260         CONNECT_PALIA = Left(gConnect, InStr(gConnect, "DATABASE") - 1) + ";" + "DATABASE=" + CPALIA + ";"
        End If

        Dim N

270
        
        
         
         
         
280     Gdb.Execute "select * into " + CPALIA + ".dbo." + "EID FROM EID", N
        List3.AddItem "EID " + Format(N, "####0")
        
         Gdb.Execute "select * into " + CPALIA + ".dbo." + "PEL FROM PEL", N
        List3.AddItem "PEL " + Format(N, "####0")
        
290     Gdb.Execute "select * into " + CPALIA + ".dbo." + "MEM FROM MEM", N
        List3.AddItem "MEM " + Format(N, "####0")
300     Gdb.Execute "select * into " + CPALIA + ".dbo." + "PINAKES FROM PINAKES", N
        List3.AddItem "PINAKES " + Format(N, "####0")
310     Gdb.Execute "select * into " + CPALIA + ".dbo." + "XREOPIS FROM XREOPIS", N
        List3.AddItem "XREOPIS " + Format(N, "####0")
320     Gdb.Execute "select * into " + CPALIA + ".dbo." + "GRA FROM GRA", N
        List3.AddItem "GRA " + Format(N, "####0")
330     Gdb.Execute "select * into " + CPALIA + ".dbo." + "ARITMISI FROM ARITMISI", N
        List3.AddItem "ARITMISI " + Format(N, "####0")

340     Gdb.Execute "select * into " + CPALIA + ".dbo." + "PARASTAT FROM PARASTAT", N
        List3.AddItem "PARASTAT " + Format(N, "####0")

350     Gdb.Execute "select * into " + CPALIA + ".dbo." + "BARCODES FROM BARCODES", N
        List3.AddItem "BARCODES " + Format(N, "####0")

        'On Error GoTo 0
        Gdb.Execute "select * into " + CPALIA + ".dbo." + "USERS FROM USERS", N
        List3.AddItem "USERS " + Format(N, "####0")
     
        Gdb.Execute "select * into " + CPALIA + ".dbo." + "XRHSEIS FROM XRHSEIS", N
        List3.AddItem "XRHSEIS " + Format(N, "####0")

360     Gdb.Execute "select * into " + CPALIA + ".dbo." + "HMEROL FROM HMEROL", N
370     Gdb.Execute "select * into " + CPALIA + ".dbo." + "HMEROL2 FROM HMEROL2", N

380     Gdb.Execute "select * into " + CPALIA + ".DBO." + "PARAMETROI FROM PARAMETROI", N
        List3.AddItem "PARAMETROI " + Format(N, "####0")


Gdb.Execute "select * into " + CPALIA + ".dbo." + "PARAMETROI FROM PARAMETROI", N
        List3.AddItem "PARAMETROI " + Format(N, "####0")



390     On Error Resume Next

        If TABLE_EXIST("CalendarEvents") Then
            Gdb.Execute "select * into " + CPALIA + ".dbo." + "CalendarEvents FROM CalendarEvents", N
        End If

        If TABLE_EXIST("CalendarEvents") Then
400         Gdb.Execute "select * into " + CPALIA + ".dbo." + "CalendarRecurrencePatterns FROM CalendarRecurrencePatterns", N
        End If

410     Gdb.Execute "select * into " + CPALIA + ".dbo." + "EGG FROM EGG WHERE YEAR(HME)=" + str(TREX - 1), N
        List3.AddItem "EGG " + Format(N, "####0")
        'Gdb.Execute "select * into " + CPALIA + ".DBO." + "EGGTIM FROM EGGTIM WHERE MONTH(HME)<7 AND YEAR(HME)=" + Str(TREX - 1), n
        'Gdb.Execute "DROP TABLE " + CPALIA + ".DBO." + "EGGTIM"

        'Gdb.Execute "select * into " + CPALIA + ".DBO." + "EGGTIM FROM EGGTIM WHERE MONTH(HME)<7 AND YEAR(HME)=" + Str(TREX - 1), n
        'Gdb.Execute "INSERT INTO  " + CPALIA + ".DBO." + "EGGTIM    SELECT * FROM EGGTIM WHERE MONTH(HME)>=7 AND YEAR(HME)=" + Str(TREX - 1), n

        On Error Resume Next

420     Gdb.Execute "DROP TABLE " + CPALIA + ".dbo." + "EGGTIM"

        On Error GoTo NEAXRONIA_Click_Err

430     Gdb.Execute "select * into " + CPALIA + ".dbo." + "EGGTIM FROM EGGTIM WHERE YEAR(HME)=" + str(TREX - 1), N
        List3.AddItem "EGGTIM 1-6 " + Format(N, "####0")

        'sql.Open CONNECT_PALIA
        'sql.Execute "ALTER TABLE EGGTIM DROP COLUMN ID_NUM"

440     '   Gdb.Execute "INSERT INTO  " + CPALIA + ".DBO." + "EGGTIM    SELECT * FROM EGGTIM WHERE MONTH(HME)=7 AND YEAR(HME)=" + str(TREX - 1), n
        List3.AddItem "EGGTIM 7 " + Format(N, "####0")

        Dim GDBPALIA As New ADODB.Connection

        'GDBPALIA.Open CONNECT_PALIA

        'GDBPALIA.Execute "alter TABLE EGGTIM DROP

        Dim SQL3     As String

        SQL3 = "[EIDOS] ,[ATIM],[POSO],[MONA],[TIMM],[KERDOS],[KODE],[HME],[ERGO],[FPA],[PROOD],[PROOD_AJ],[EKPT],[KAU_AJIA],[MIK_AJIA],[ONOMA],[MIKTA],[KOLA],[PELKOD],[PROELEYSH],[XRE],[PIS],[APOT],[ATIM2],[FCURRENCY],[EKPT2]"

450     Gdb.Execute "INSERT INTO  " + CPALIA + ".DBO." + "EGGTIM(" + SQL3 + ")  SELECT " + SQL3 + " FROM EGGTIM WHERE MONTH(HME)=7 AND YEAR(HME)=" + str(TREX - 1), N

        Gdb.Execute "INSERT INTO  " + CPALIA + ".DBO." + "EGGTIM(" + SQL3 + ")  SELECT " + SQL3 + " FROM EGGTIM WHERE MONTH(HME)=8 AND YEAR(HME)=" + str(TREX - 1), N
        List3.AddItem "EGGTIM 8 " + Format(N, "####0")
460     Gdb.Execute "INSERT INTO  " + CPALIA + ".DBO." + "EGGTIM(" + SQL3 + ")  SELECT " + SQL3 + " FROM EGGTIM WHERE MONTH(HME)=9 AND YEAR(HME)=" + str(TREX - 1), N
        List3.AddItem "EGGTIM 9 " + Format(N, "####0")

470     Gdb.Execute "INSERT INTO  " + CPALIA + ".DBO." + "EGGTIM(" + SQL3 + ")  SELECT " + SQL3 + " FROM EGGTIM WHERE MONTH(HME)=10 AND YEAR(HME)=" + str(TREX - 1), N
        List3.AddItem "EGGTIM 10 " + Format(N, "####0")
480     Gdb.Execute "INSERT INTO  " + CPALIA + ".DBO." + "EGGTIM(" + SQL3 + ")  SELECT " + SQL3 + " FROM EGGTIM WHERE MONTH(HME)=11 AND YEAR(HME)=" + str(TREX - 1), N
        List3.AddItem "EGGTIM 11 " + Format(N, "####0")
490     Gdb.Execute "INSERT INTO  " + CPALIA + ".DBO." + "EGGTIM(" + SQL3 + ")  SELECT " + SQL3 + " FROM EGGTIM WHERE MONTH(HME)=12 AND YEAR(HME)=" + str(TREX - 1), N
        List3.AddItem "EGGTIM 12 " + Format(N, "####0")

500     Gdb.Execute "select * into " + CPALIA + ".DBO." + "TIM FROM TIM WHERE YEAR(HME)=" + str(TREX - 1), N
        List3.AddItem "TIM  " + Format(N, "####0")

510     Gdb.Execute "DELETE  FROM EGG WHERE YEAR(HME)=" + str(TREX - 1), N
520     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)<7 AND YEAR(HME)=" + str(TREX - 1), N
530     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=7 AND YEAR(HME)=" + str(TREX - 1), N
540     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=8 AND YEAR(HME)=" + str(TREX - 1), N
550     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=9 AND YEAR(HME)=" + str(TREX - 1), N
560     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=10 AND YEAR(HME)=" + str(TREX - 1), N
570     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=11 AND YEAR(HME)=" + str(TREX - 1), N
580     Gdb.Execute "DELETE  FROM EGGTIM WHERE MONTH(HME)=12  AND YEAR(HME)=" + str(TREX - 1), N

590     Gdb.Execute "DELETE  FROM TIM WHERE YEAR(HME)=" + str(TREX - 1), N

600     Gdb.Execute "INSERT INTO ETAIREIES (TITLOS,DSN)  VALUES ('" + Format(PALIA, "0000") + "','" + "D" + Format(PALIA, "0000") + "')"

        List3.AddItem "‘≈Àœ” Ã≈‘¡÷œ—¡” "

        ''---------- RESTORE –¡À…¡” ◊—œÕ…¡” ”‘«Õ DATABASE ‘«” –¡À…¡” ◊—œÕ…¡” -------------------------------------
        'On Error GoTo NEAXRONIA_Click_Err
        '
        'Set db = OpenDatabase("", False, False, CONNECT_PALIA)
        'Me.MousePointer = vbHourglass
        'Dim LISTA(20)
        'LISTA(1) = "PEL"  -
        'LISTA(2) = "EID"-
        'LISTA(3) = "EGG"
        'LISTA(4) = "EGGTIM"
        'LISTA(5) = "TIM"
        'LISTA(6) = "GRA"
        'LISTA(7) = "MEM"
        'LISTA(8) = "PINAKES"
        'LISTA(9) = "XREOPIS"
        'LISTA(10) = "ARITMISI"-
        'LISTA(11) = "PARASTAT"-
        'LISTA(12) = "BARCODES"-
        'LISTA(13) = "HMEROL"
        'LISTA(14) = "HMEROL2"
        'LISTA(15) = "PARAMETROI"
        'LISTA(16) = "CalendarEvents"
        'LISTA(17) = "CalendarRecurrencePatterns"
        '
        'ALLAGH.Caption = "2.¡Õ‘…√—¡÷« ”≈ –¡À…¡ ◊—œÕ…¡"
        '
        '
        'DB_PALIA.Open CONNECT_PALIA  ' Left(gConnect, InStr(gConnect, "DATABASE") - 1) + ";DATABASE=" + CPALIA
        '
        'For k = 1 To 17
        '   LATOS = 0
        ''   On Error GoTo LATOS
        '   DoEvents
        '   Me.Caption = LISTA(k)
        '   On Error Resume Next
        '   DB_PALIA.Execute "DROP TABLE " + LISTA(k)
        '
        '
        '   If LISTA(k) = "EGG" Or LISTA(k) = "EGGTIM" Or LISTA(k) = "TIM" Then
        '      db.Execute "SELECT * INTO " + LISTA(k) + "  FROM  " + LISTA(k) + "  IN '" + DD + "' WHERE YEAR(HME)=" + palia
        '   Else
        '
        '      db.Execute "SELECT * INTO " + LISTA(k) + "  FROM  " + LISTA(k) + "  IN '" + DD + "'"
        '   End If
        '
        'Next
        'Me.MousePointer = vbNormal
        '
        '
        ''Gdb.Close
        '
        'Gdb.Open gConnect
        '
        ''--- ≈À≈√◊œ” ¡Õ Ã≈‘¡÷≈—»« ¡Õ œ…  …Õ«”≈…” ------------------
        'ALLAGH.Caption = "2¡.≈À≈√◊œ” ¡Õ Ã≈‘¡÷≈—»« ¡Õ œ…  …Õ«”≈…”"
        'Dim R As New ADODB.Recordset
        'Dim r2 As New ADODB.Recordset
        '
        '
        'On Error GoTo NEAXRONIA_Click_Err
        'R.Open "SELECT SUM(AJI) FROM TIM WHERE YEAR(HME)=" + palia, Gdb, adOpenForwardOnly, adLockReadOnly
        'r2.Open "SELECT SUM(AJI) FROM TIM WHERE YEAR(HME)=" + palia, DB_PALIA, adOpenForwardOnly, adLockReadOnly
        '
        '
        'If Abs(R(0) - r2(0)) < 1 Then
        '    'OK
        'Else
        '    MsgBox "–—œ¬À«Ã¡ ”‘«Õ Ã≈‘¡÷œ—¡. « ◊—œÕ…¡ –¡—¡Ã≈Õ≈… Ÿ” ≈◊≈…"
        '    DB_PALIA.Close
        '
        '    Exit Sub
        'End If
        '
        'R.Close
        'r2.Close
        '
        'R.Open "SELECT SUM(XRE) FROM EGG WHERE YEAR(HME)=" + palia, Gdb, adOpenForwardOnly, adLockReadOnly
        'r2.Open "SELECT SUM(XRE) FROM EGG WHERE YEAR(HME)=" + palia, DB_PALIA, adOpenForwardOnly, adLockReadOnly
        'If Abs(R(0) - r2(0)) < 1 Then
        '    'OK
        'Else
        '    MsgBox "–—œ¬À«Ã¡ ”‘«Õ Ã≈‘¡÷œ—¡. « ◊—œÕ…¡ –¡—¡Ã≈Õ≈… Ÿ” ≈◊≈…"
        '        DB_PALIA.Close
        '
        '    Exit Sub
        '
        'End If
        '
        'R.Close
        'r2.Close
        'R.Open "SELECT SUM(POSO) FROM EGGTIM WHERE YEAR(HME)=" + palia, Gdb, adOpenForwardOnly, adLockReadOnly
        'r2.Open "SELECT SUM(POSO) FROM EGGTIM WHERE YEAR(HME)=" + palia, DB_PALIA, adOpenForwardOnly, adLockReadOnly
        'If Abs(R(0) - r2(0)) < 1 Then
        '    'OK
        'Else
        '    MsgBox "–—œ¬À«Ã¡ ”‘«Õ Ã≈‘¡÷œ—¡. « ◊—œÕ…¡ –¡—¡Ã≈Õ≈… Ÿ” ≈◊≈…"
        '    DB_PALIA.Close
        '    Exit Sub
        'End If
        '
        '
        '
        'Gdb.Execute "INSERT INTO ETAIREIES (TITLOS,DSN)  VALUES ('" + Format(palia, "0000") + "','" + "D" + Format(palia, "0000") + "')"
        '
        '
        '
        'ALLAGH.Caption = "3.”¬«”…Ãœ  …Õ«”≈ŸÕ ¡–œ ‘—≈◊œ’”¡ ◊—œÕ…¡"
        ''--------------------- Ã«ƒ≈Õ…”Ãœ” ‘…ÃœÀœ√…ŸÕ ----------------------------
        'Gdb.Execute "delete from TIM WHERE YEAR(HME)=" + palia
        'Gdb.Execute "delete from EGGTIM WHERE YEAR(HME)=" + palia
        'Gdb.Execute "delete from EGG WHERE YEAR(HME)=" + palia
        '
        '
        '
        '
        'On Error Resume Next
        'DB_PALIA.Execute "ALTER TABLE TIM ADD  ID_NUM2 int "
        '      DB_PALIA.Execute "UPDATE TIM SET ID_NUM2 = ID_NUM"
        '      DB_PALIA.Execute "ALTER TABLE TIM DROP COLUMN ID_NUM "
        '      DB_PALIA.Execute "ALTER TABLE TIM ADD  ID_NUM int IDENTITY(1,1)"
        '
        '
        'MsgBox "‘≈Àœ” Ã≈‘¡÷œ—¡”"

        '<EhFooter>
        Exit Sub

NEAXRONIA_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.NEAXRONIA_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.NEAXRONIA_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub PEL_Click()

        '<EhHeader>
        On Error GoTo PEL_Click_Err

        '</EhHeader>

        Dim DB    As Database

        Dim CA    As String

        Dim DB2   As New ADODB.Connection

        Dim PALIA As String

        Dim TREX  As String

        Dim ANS

        Dim arxeio

        Dim a As Long

        Dim CPALIA

        Dim N

        Dim sql  As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

        On Error GoTo OUTSIDE

        ANS = MsgBox("≈˜ÂÙÂ ıÔÎÔ„ﬂÛÂÈ Ù· ı¸ÎÔÈ· ÛÙÁÌ ·ÎÈ‹ ˜ÒÔÌÈ‹;", vbYesNo)

        If ANS = vbNo Then
            MsgBox "ƒÂÌ ›„ÈÌÂ Á ÏÂÙ·ˆÔÒ‹ ıÔÎÔﬂ˘Ì"

            Exit Sub

        End If

100     TREX = InputBox("ETO” NEA” ◊—œÕ…¡” ", , Year(Date))

110     PALIA = InputBox("¬¡”« –¡À…¡” ◊—œÕ…¡” –.◊. D2007 ", , "D" + Format(Year(Date) - 1), "0000")

120     If Len(PALIA) = 0 Then Exit Sub

130     Gdb.Execute "update PEL SET AYP=0"

140     sql = "update PEL SET AYP=A.TYP FROM PEL INNER JOIN " + PALIA + ".dbo.PEL A  "
150     sql = sql + " on  PEL.KOD=A.KOD COLLATE Greek_CS_AS AND PEL.EIDOS=A.EIDOS COLLATE Greek_CS_AS  "

160     Gdb.Execute sql, N

        On Error GoTo PEL_Click_Err

170     MsgBox "≈Õ«Ã≈—Ÿ»« ¡Õ " + str(N) + " ≈√√—¡÷≈” –≈À¡‘ŸÕ / –—œÃ«»≈’‘ŸÕ"

        Exit Sub

180     CPALIA = PALIA

190     Gdb.Close

200     Gdb.Open gConnect + ";DATABASE=" + CPALIA

210     DB2.Open gConnect

220     MsgBox "–—œ«√œ’Ã≈ÕŸ” –¡‘«”‘≈ ”‘«Õ –¡À…¡ ◊—œÕ…¡ –¡—œ’”…¡”« –≈À¡‘ŸÕ & –—œÃ«»≈’‘ŸÕ"
        'update_pel2   14-1-2010

230     rSQL.Open "SELECT * FROM PEL ", Gdb, adOpenDynamic, adLockOptimistic

        'Dim CA As String
        Dim ssa As String

        'Dim K As Long
240     rSQL.MoveFirst
250     k = 0

260     Do While Not rSQL.EOF

270         If Not IsNull(rSQL("TYP")) Then
280             If rSQL("TYP") <> 0 Then
290                 k = k + 1
300                 DB2.Execute "UPDATE PEL SET AYP=" + str(rSQL("TYP")) + "  WHERE KOD='" + rSQL("KOD") + "' AND EIDOS='" + rSQL("EIDOS") + "'"

                    '        SSA = "insert into EGG (HME,ATIM,APA,AIT,EIDOS,KOD,XPI,XREOSI,PISTOSI) values ("
                    '        SSA = SSA + "'" + Format(Now, "01/01/" + TREX) + "',"
                    '        SSA = SSA + "'Î" + Format(K, "00000") + "'," ' ATIM
                    '        SSA = SSA + Format(K, "00000") + "," 'APA
                    '        SSA = SSA + "'" + "¡–œ√—¡÷« + " ',"  'AIT
                    '        SSA = SSA + "'" + rSQL("Eidos") + "'," ' EIDOS='E'
                    '        SSA = SSA + "'" + rSQL("kod") + "'," ' KOD
                    '        SSA = SSA + "'" + "◊" + "'," ' XPI
                    '        SSA = SSA + Str(rSQL("TYP")) + "," 'XREOSI
                    '        SSA = SSA + Str(0) + ")" ' PISTOSI
                    '        Gdb.Execute SSA

                End If

            End If

310         rSQL.MoveNext

320         If k Mod 10 = 0 Then DoEvents: Me.Caption = str(k)
        Loop

        '  Gdb.Execute

330     Gdb.Close
340     Gdb.Open gConnect

350     MsgBox "Ù›ÎÔÚ ÏÂÙ·ˆÔÒ‹Ú ÂÎ·Ù˛Ì - ÒÔÏÁËÂıÙ˛Ì"

        Exit Sub

OUTSIDE:
360     MsgBox "ƒ≈Õ œÀœ À«—Ÿ»« ≈ « Ã≈‘¡÷œ—¡ ’–œÀœ…–œ’ –≈À¡‘ŸÕ / –—œÃ«»≈’‘ŸÕ À¡»œ” :" + Chr(13) + Err.Description

        '     Resume Next
        Exit Sub

LATOS:
370     LATOS = 1

380     Resume Next

390     Return
LATOS2:
400     MsgBox "À¡»œ” œÕœÃ¡ ¬¡”«” ACCESS  ƒŸ”‘≈ –.◊.  C:\BACK.MDB"

        '<EhFooter>
        Exit Sub

PEL_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.PEL_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.PEL_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Sub update_pel2()
''ypologismos ypoloipon
'    Dim X As String, order As String
'    Dim k As Integer, R As Recordset    ', sum As Single
'    Dim mBUFF
'    Dim execX, file
'    'If Combo1.Text = Combo1.List(0) Then mBUFF = "e" Else mBUFF = "r"
'
'    F_RUN = 1
'
'    'Data1.DatabaseName = gDir
'
'
'    'On Error GoTo 0
'    ' ’–œÀœ√…∆Ÿ ‘¡ ‘≈À≈’‘¡…A ’–œÀœ…–¡ ¡–œ ‘œ EGG
'    On Error Resume Next
'    Gdb.Execute "DROP TABLE EGG1"
'    On Error GoTo 0
'
'    Dim L As Long
'    Gdb.Execute "UPDATE PEL  SET TYP=0,XRESYN=0,PISSYN=0 FROM" _
'              & " PEL WHERE EIDOS='" + mBUFF + "'", L
'
'
'    Gdb.Execute "SELECT EIDOS,KOD,SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END ) AS SX," _
'              & " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END ) AS SP INTO EGG1" _
'              & " FROM EGG   GROUP BY EIDOS,KOD "
'
'
'
'
'    'Gdb.Execute "CREATE VIEW dbo.[UPD_YPOL_PEL] as " _
'     '& " SELECT    dbo.PEL.XRESYN,dbo.PEL.PISSYN,dbo.PEL.TYP, dbo.PEL.KOD, dbo.PEL.EIDOS, dbo.EGG1.SP, dbo.EGG1.SX, dbo.PEL.AYP" _
'     '& "  FROM         dbo.PEL INNER JOIN" _
'     '& "  dbo.EGG1 ON dbo.PEL.EIDOS = dbo.EGG1.EIDOS AND dbo.PEL.KOD = dbo.EGG1.KOD"
'
'
'
'
'
'    'If Combo1.Text = Combo1.List(0) Then '  "e",
'    'Gdb.Execute "UPDATE UPD_YPOL_PEL  SET TYP=AYP+SX-SP,XRESYN=SX,PISSYN=SP"
'
'    '  If gXEIRISTHS = 8 Then
'    ' Else
'    Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" _
'              & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" _
'              & " PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.EIDOS WHERE PEL.EIDOS='e'"
'    'End If
'
'    'Else
'    '   On Error GoTo 0
'    Gdb.Execute "UPDATE PEL  SET TYP=AYP-(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )+" _
'              & " (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" _
'              & " PEL INNER JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.EIDOS where PEL.EIDOS='r'"
'
'    'End If
'
'
'
'
'
'    Exit Sub
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'    On Error GoTo 0
'
'
'    'On Error GoTo 0
'    ' ’–œÀœ√…∆Ÿ ‘¡ ‘≈À≈’‘¡…A ’–œÀœ…–¡ ¡–œ ‘œ EGG
'    Gdb.Execute "DROP TABLE EGG1"
'
'
'
'    Gdb.Execute "UPDATE PEL  SET TYP=0,XRESYN=0,PISSYN=0 FROM" _
'              & " PEL ", L
'
'
'    Gdb.Execute "SELECT UPPER(EIDOS) as MEIDOS,KOD,SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END ) AS SX," _
'              & " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END ) AS SP INTO EGG1" _
'              & " FROM EGG  GROUP BY UPPER(EIDOS),KOD "
'
'
'    On Error GoTo 0
'
'
'    'Gdb.Execute "CREATE VIEW dbo.[UPD_YPOL_PEL] as " _
'     '& " SELECT    dbo.PEL.XRESYN,dbo.PEL.PISSYN,dbo.PEL.TYP, dbo.PEL.KOD, dbo.PEL.EIDOS, dbo.EGG1.SP, dbo.EGG1.SX, dbo.PEL.AYP" _
'     '& "  FROM         dbo.PEL INNER JOIN" _
'     '& "  dbo.EGG1 ON dbo.PEL.EIDOS = dbo.EGG1.EIDOS AND dbo.PEL.KOD = dbo.EGG1.KOD"
'
'
'
'
'
'
'    ' If gXEIRISTHS = 8 Then
'    Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" _
'              & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" _
'              & " PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND UPPER(PEL.EIDOS)=EGG1.MEIDOS", L
'    ' Else
'    '  Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" _
'       '   & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" _
'       '   & " PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.EIDOS", L
'    ' End If
'
'    'Else
'    '   On Error GoTo 0
'    Gdb.Execute "UPDATE PEL  SET TYP=AYP-(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )+" _
'              & " (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" _
'              & " PEL INNER JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.MEIDOS"
'
'    'End If
'
'End Sub
