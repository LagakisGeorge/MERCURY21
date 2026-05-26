VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form BOHU13 
   BackColor       =   &H00FF0000&
   Caption         =   "Ελεγχος πληρότητας"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10545
   ScaleWidth      =   10980
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "ΛΙΣΤΑ ΠΑΡΑΣΤΑΤΙΚΩΝ ΜΕ ΑΥΤΑ ΤΑ ΣΤΟΙΧΕΙΑ"
      Height          =   930
      Left            =   4320
      TabIndex        =   11
      Top             =   7470
      Width           =   1800
   End
   Begin VB.TextBox ATIM 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   8025
      Width           =   1575
   End
   Begin VB.TextBox HMER 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   7470
      Width           =   1575
   End
   Begin VB.CommandButton EPANAFORA 
      Caption         =   "Επαναφορά Διαγραμμένου Παρ/κού"
      Height          =   930
      Left            =   6270
      TabIndex        =   6
      Top             =   7470
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   840
      TabIndex        =   5
      Top             =   3240
      Width           =   9270
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Επανενημέρωση ποσοτήτων αποθήκης"
      Height          =   915
      Left            =   5460
      TabIndex        =   4
      Top             =   2025
      Width           =   1950
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Επανυπολογισμός Παραστατικών"
      Height          =   885
      Left            =   3255
      TabIndex        =   3
      Top             =   2055
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ελεγχος Αν έχουν περαστεί τα τιμολόγια στις κινήσεις πελατών προμηθευτών"
      Height          =   885
      Left            =   1020
      TabIndex        =   2
      Top             =   2055
      Width           =   1965
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   315490305
      CurrentDate     =   40200
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   315490305
      CurrentDate     =   40200
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   840
      Top             =   7440
      Width           =   9255
   End
   Begin VB.Label Label2 
      Caption         =   "Παρ/κό π.χ. : G000121"
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   8025
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Ημερ/νία: μμ/ηη/εεεε"
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   7470
      Width           =   1455
   End
End
Attribute VB_Name = "BOHU13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
Dim f_id As Long

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim DB   As Database

        Dim Rp   As New ADODB.Recordset

        Dim Rtim As New ADODB.Recordset

        Dim R    As New ADODB.Recordset

        Dim rEgg As New ADODB.Recordset

        '
        'sql = "SELECT * FROM TIM  WHERE HME>='" + Format(DTPicker1.value, "MM/DD/YYYY") + "'"
        'sql = sql + " AND HME<='" + Format(DateAdd("d", 1, DTPicker2.value), "MM/DD/YYYY") + "'"
        'Rtim.Open sql, Gdb, adOpenDynamic, adLockOptimistic
        
        
        Dim ccc As String
        ccc = InputBox("ΣΤΟΝ ΤΡΟΠΟ ΠΛΗΡΩΜΗΣ ΝΑ ΠΕΡΙΕΧΕΤΑΙ Η ΛΕΞΗ ", "")
        
        
        
        
        

100     Gdb.BeginTrans

        Dim nCount

110     nCount = 0

120     sql = "SELECT * FROM TIM  WHERE  LEFT(ATIM,1) IN (SELECT EIDOS FROM PARASTAT WHERE PEL IN (1,2,3,4) ) AND TRP LIKE '%" + ccc + "%' AND HME>='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "'"
130     sql = sql + " AND HME<='" + Format(DateAdd("d", 1, DTPicker2.Value), "MM/DD/YYYY") + "'"
140     Rtim.Open sql, Gdb, adOpenDynamic, adLockOptimistic

        Dim S2   As String

        Dim mfil As String

150     mfil = "C:\EGGRAFES.TXT"

160     Open mfil For Output As #1
170     R.Open "select top 1 * from EGG", Gdb, adOpenDynamic, adLockOptimistic

180     Do While Not Rtim.EOF
190         rEgg.Open "select * from EGG where KOD='" + Rtim("KPE") + "' AND ATIM='" + Rtim("ATIM") + "' AND DAY(HME)=" + str(Day(Rtim("HME"))) + " and month(HME)=" + str(Month(Rtim("HME"))), Gdb, adOpenDynamic, adLockOptimistic

200         If rEgg.EOF Then
210             Rp.Open "SELECT * FROM PARASTAT WHERE EIDOS='" + Left(Rtim("atim"), 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
220             R.AddNew
230             R("hme") = CDate(Format(Rtim("HME"), "DD/MM/YYYY"))    ' DTPicker1.Value
240             R("UserID") = gUserId
250             R("apa") = Val(mID$(Rtim("atim"), 2, 6))
260             R("atim") = Rtim("atim")    ' Left(f_matim, 1 + f_psifia_atim)
270             R("ait") = Left(Rp("TITLOS"), 20) + "**"
280             R("XREOSI") = 0: R("PISTOSI") = 0    ' ΑΠΟΦΥΓΗ ΝULL
                R("IDTIM") = Rtim!id_num
290             If Rp("pel") = "1" Or Rp("pel") = "4" Then    ' XREVSH
300                 R("xre") = Rtim("AJI")

310                 If Rp("pel") = "1" Then R("XREOSI") = Rtim("AJI")
320                 If Rp("pel") = "4" Then R("PISTOSI") = -Rtim("AJI")
                Else
330                 R("xre") = -Rtim("AJI")

340                 If Rp("pel") = "2" Then R("XREOSI") = -Rtim("AJI")
350                 If Rp("pel") = "3" Then R("PISTOSI") = Rtim("AJI")
                End If

360             R("eidos") = Rtim("eidos")
370             R("kod") = Rtim("kpe")
380             S2 = R("ATIM") + Format(R("HME"), "DD/MM/YYYY") + "ΚΩΔ.ΠΕΛ" + R("KOD")
390             List1.AddItem S2
400             Print #1, S2
410             R.Update
420             Rp.Close


If InStr(Rtim("TRP"), "ΠΙΣΤΩΣ") > 0 Or InStr(Rtim("TRP"), "2;") > 0 Then
Else

2210             Rp.Open "SELECT * FROM PARASTAT WHERE EIDOS='" + Left(Rtim("atim"), 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
2220             R.AddNew
2230             R("hme") = CDate(Format(Rtim("HME"), "DD/MM/YYYY"))    ' DTPicker1.Value
2240             R("UserID") = gUserId
2250             R("apa") = Val(mID$(Rtim("atim"), 2, 6))
2260             R("atim") = Rtim("atim")    ' Left(f_matim, 1 + f_psifia_atim)
2270             R("ait") = Left(Rp("TITLOS"), 20) + "**"
2280             R("XREOSI") = 0: R("PISTOSI") = 0    ' ΑΠΟΦΥΓΗ ΝULL

2290             If Rp("pel") = "1" Or Rp("pel") = "4" Then    ' XREVSH
2300                 R("xre") = -Rtim("AJI")

2310                 If Rp("pel") = "1" Then R("PISTOSI") = Rtim("AJI")
2320                 If Rp("pel") = "4" Then R("PISTOSI") = Rtim("AJI")
                Else
2330                 R("xre") = -Rtim("AJI")

2340                 If Rp("pel") = "2" Then R("XREOSI") = Rtim("AJI")
2350                 If Rp("pel") = "3" Then R("PISTOSI") = -Rtim("AJI")
                End If

2360             R("eidos") = Rtim("eidos")
2370             R("kod") = Rtim("kpe")
2380             S2 = R("ATIM") + Format(R("HME"), "DD/MM/YYYY") + "ΚΩΔ.ΠΕΛ" + R("KOD")
2390             List1.AddItem "---------------" + S2
2400             Print #1, S2
2410             R.Update
2420             Rp.Close

End If
















430             nCount = nCount + 1
440             Me.Caption = "Δημιουργήθηκαν " + str(nCount) + " εγγραφές"
            End If

450         rEgg.Close

460         Rtim.MoveNext
        Loop

470     Close #1
    
    
If nCount = 0 Then
      MsgBox " ΔΕΝ ΒΡΕΘΗΚΑΝ"
      Exit Sub
End If
    
    
        Dim ANS

480     ANS = MsgBox("Να αποθηκευτούν οι αλλαγές μόνιμα;", vbYesNo)

490     If ANS = vbYes Then
500         Gdb.CommitTrans
510         MsgBox "ΟΚ Αποθηκεύτηκαν"
520         ANS = MsgBox("Οι εγγραφές που δημιουργήθηκαν είναι επίσης αποθηκευμένες στο " + mfil + Chr(13) + "Θέλετε να το ανοίξετε;", vbYesNo)

530         If ANS = vbYes Then
540             DUM = Shell("notepad.exe " + mfil, vbMaximizedFocus)
            End If

        Else
550         Gdb.RollbackTrans
560         MsgBox "Δεν Αποθηκεύτηκαν"

        End If

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU13.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU13.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim R     As New ADODB.Recordset
        Dim r2     As New ADODB.Recordset
        
        Dim s     As String

       ' Dim R2    As New ADODB.Recordset
        
        Dim ANS As Integer
        ANS = MsgBox("ΕΧΕΤΕ ΠΑΡΕΙ BACKUP ΠΡΙΝ ΣΥΝΕΧΙΣΕΤΕ;", vbYesNo)
        
        If ANS = vbNo Then
             Exit Sub
        
        End If
        
        ANS = MsgBox("ΝΑ ΕΜΦΑΝΙΣΤΟΥΝ ΜΟΝΟ ΟΙ ΔΙΑΦΟΡΕΣ;", vbYesNo)
        
         Dim synt2 As String
        synt2 = " AND  HME>='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "'"
        synt2 = synt2 + " AND HME<='" + Format(DateAdd("d", 1, DTPicker2.Value), "MM/DD/YYYY") + "'"
        
        

        Dim EIDOS As String

100     EIDOS = "G','g','t','T','L','l"
110     EIDOS = InputBox("ΔΩΣΕ ΤΑ ΠΑΡΑΣΤΑΤΙΚΑ ΠΟΥ ΘΑ ΕΝΗΜΕΡΩΘΟΥΝ ", , EIDOS)
'UPDATE EGGTIM KAU_AJIA,MIK_AJIA
120     Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA= POSO*TIMM*(100-EKPT)/100*" + str((100 - g_Fpa(1)) / 100) + " where FPA=1 AND LEFT(ATIM,1) IN ('" + EIDOS + "')  " + synt2
130     Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA= POSO*TIMM*(100-EKPT)/100*" + str((100 - g_Fpa(2)) / 100) + " where FPA=2 AND LEFT(ATIM,1) IN ('" + EIDOS + "')  " + synt2
140     Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA= POSO*TIMM*(100-EKPT)/100*" + str((100 - g_Fpa(3)) / 100) + " where FPA=3 AND LEFT(ATIM,1) IN ('" + EIDOS + "')  " + synt2
150     Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA= POSO*TIMM*(100-EKPT)/100*" + str((100 - g_Fpa(4)) / 100) + " where FPA=4 AND LEFT(ATIM,1) IN ('" + EIDOS + "')  " + synt2
160     Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA= POSO*TIMM*(100-EKPT)/100*" + str((100 - g_Fpa(5)) / 100) + " where FPA=5 AND LEFT(ATIM,1) IN ('" + EIDOS + "')  " + synt2

170     s = "select sum(CASE WHEN FPA=1 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S1,"
180     s = s + " sum(CASE WHEN FPA=2 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S2,"
190     s = s + " sum(CASE WHEN FPA=3 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S3,"
200     s = s + " sum(CASE WHEN FPA=4 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S4,"
210     s = s + " sum(CASE WHEN FPA=5 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S5, "
        s = s + " sum(CASE WHEN FPA=6 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S6, "
        s = s + " sum(CASE WHEN FPA=7 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S7, "
220     s = s + "ATIM,HME,ID_NUM FROM EGGTIM where LEFT(ATIM,1) IN ('" + EIDOS + "') " + synt2 + " GROUP BY ATIM,HME,ID_NUM "

230     R.Open s, Gdb, adOpenDynamic, adLockOptimistic

        Dim Z As Single

240     R.MoveFirst

250     Do While Not R.EOF
260         Z = R("S1") * (100 + g_Fpa(1)) / 100
270         Z = Z + R("S2") * (100 + g_Fpa(2)) / 100
280         Z = Z + R("S3") * (100 + g_Fpa(3)) / 100
290         Z = Z + R("S4") * (100 + g_Fpa(4)) / 100
300         Z = Z + R("S5") * (100 + g_Fpa(5)) / 100
            Z = Z + R("S6") * (100 + g_Fpa(6)) / 100
            Z = Z + R("S7") * (100 + g_Fpa(7)) / 100
            
310         s = "UPDATE TIM SET AJ1=" + str(R("S1")) + ","
320         s = s + " AJ2=" + str(R("S2")) + ","
330         s = s + " AJ3=" + str(R("S3")) + ","
340         s = s + " AJ4=" + str(R("S4")) + ","
350         s = s + " AJ5=" + str(R("S5")) + ","
            s = s + " AJ6=" + str(R("S6")) + ","
            s = s + " AJ7=" + str(R("S7")) + ","




360         s = s + "AJI=" + str(Z) + " WHERE ATIM='" + R("ATIM") + "' AND HME='" + Format(R("HME"), "MM/DD/YYYY") + "'"


        If ANS = vbYes Then  ' MONO BLEPO
            r2.Open "SELECT * FROM TIM WHERE ID_NUM=" + str(R("ID_NUM")), Gdb, adOpenDynamic, adLockOptimistic
            If Abs(r2("AJI") - Z) > 0.01 Then
               List1.AddItem r2("ATIM") + " " + Format(R("HME"), "dd/mm/yyyy") + " ΔΙΑΦΟΡΑ " + Format(Abs(r2("AJI") - Z), "####.00")
            End If
            r2.Close
            
        Else
370         Gdb.Execute s
        End If
            'R2.Open "SELECT * FROM EGG  WHERE ATIM='" + R("ATIM") + "' AND HME='" + Format(R("HME"), "MM/DD/YYYY") + "'"", Gdb, adOpenDynamic, adLockOptimistic"

            'ΜΟΝΟ ΠΙΣΤΟΣΗ ΚΑΝΕΙ
            '  S = "UPDATE EGG SET PISTOSI=" + Str(Z) + " WHERE ATIM='" + R("ATIM") + "' AND HME='" + Format(R("HME"), "MM/DD/YYYY") + "'"
            '  Gdb.Execute S
            Me.Caption = R("HME")
380         R.MoveNext
        Loop

390     R.Close


         MsgBox "ΟΛΟΚΛΗΡΩΘΗΚΕ"
          
        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU13.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU13.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Function find_eid_parastat()
'
'  Dim r As New ADODB.Recordset
'
'  r.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
'  r.MoveFirst
'  Do While Not r.EOF
'     If Right(PARAS.Text, 1) = r("EIDOS") Then
'       F_TITLOS = r("TITLOS")
'       F_POS_APOU = r("POS_APOU")
'       F_AJIA_APOU = r("AJIA_APOU")
'       f_pel = r("pel")
'       f_pol = r("pol") ' 1=POLISEIS 2=AGORES
'       F_STADIO = r("STADIO") ' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
'       If IsNull(r("GEF_P")) Then
'           F_rec_p = 0
'       Else
'           F_rec_p = r("GEF_P") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
'       End If
'       If IsNull(r("XONDR")) Then
'           F_XONDR = 0
'       Else
'           F_XONDR = r("XONDR") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
'       End If
'     End If
'     r.MoveNext
'  Loop
'
'
'End Function

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>
        Dim R      As New ADODB.Recordset

        Dim meidos As String

100     R.Open "SELECT * FROM PARASTAT ", Gdb, adOpenDynamic, adLockOptimistic

110     Do While Not R.EOF

120         meidos = R("eidos")
130         Me.Caption = meidos

140         If R("POS_APOU") = 1 Then
150             Gdb.Execute "update EGGTIM SET XRE=POSO,PIS=0 WHERE LEFT(ATIM,1)='" + meidos + "'"
160         ElseIf R("POS_APOU") = 2 Then
170             Gdb.Execute "update EGGTIM SET XRE=-POSO,PIS=0 WHERE LEFT(ATIM,1)='" + meidos + "'"
180         ElseIf R("POS_APOU") = 3 Then
190             Gdb.Execute "update EGGTIM SET XRE=0,PIS=POSO WHERE LEFT(ATIM,1)='" + meidos + "'"
200         ElseIf R("POS_APOU") = 4 Then
210             Gdb.Execute "update EGGTIM SET XRE=0,PIS=-POSO WHERE LEFT(ATIM,1)='" + meidos + "'"
            End If

220         R.MoveNext
        Loop

230     MsgBox "ΟΚ"

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU13.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU13.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub EPANAFORA_Click()
   Dim R As New ADODB.Recordset
   R.Open "select * from LOGGING WHERE ID>=" + str(f_id) + " AND ENTOLH LIKE '%" + HMER.Text + "%' AND ENTOLH LIKE '%" + ATIM.Text + "%' order by ID", Gdb, adOpenDynamic, adLockOptimistic
   
   Dim FOUND As Boolean
   FOUND = False
   Dim ent(1000) As String
   Dim F As Integer
   F = 0
   Dim IDLAST As Long
   If Not R.EOF Then  'ΓΙΑ ΝΑ ΠΙΑΣΩ ΚΑΙ ΤΑ MYDATA
       R.MoveLast
       IDLAST = R!ID
       R.MoveFirst
       R.Close
       R.Open "select * from LOGGING WHERE ID>=" + str(f_id) + " AND ID<=" + str(IDLAST) + " ORDER BY ID", Gdb, adOpenDynamic, adLockOptimistic
   End If
   
   
   
   Me.MousePointer = vbHourglass
   
   Do While Not R.EOF
      List1.AddItem R!Entolh
      FOUND = True
      
      ' ΑΝ ΣΥΝΑΝΤΗΣΕΙ ΤΗΝ ΕΠΟΜΕΝΗ ΑΛΛΑΓΗ ΤΟΥ ΤΙΜΟΛΟΓΙΟΥ ΝΑ ΣΤΑΜΑΤΗΣΕΙ
      If F > 1 And InStr(R!Entolh, " TIM ") Then
           If InStr(R!Entolh, "UPDATE ") Then
           Else
              Exit Do
           End If
      End If
      
      
      F = F + 1
      ent(F) = R!Entolh
      R.MoveNext
   Loop
   Me.MousePointer = vbNormal
   If FOUND Then
      Dim r2 As New ADODB.Recordset
      Dim ANS As Integer
      Dim MAXID As String
      
      'ΨΑΧΝΩ ΝΑ ΒΡΩ ΠΟΥ ΕΙΝΑΙ ΤΟ ID_NUM ΓΙΑ ΝΑ ΤΟ ΑΛΛΑΞΩ
      If F > 1 Then
         Dim txtLines() As String
         txtLines = Split(ent(2), ",")
      End If
      
      Dim palio_ID_NUM As String
      palio_ID_NUM = txtLines(48)
      
      Gdb.BeginTrans
      On Error GoTo STOP_RESTORE
      
      
      ANS = MsgBox("Να δημιουργηθεί το παραστατικό;", vbYesNo)
      If ANS = vbYes Then
          For k = 1 To F
            
             If k = 1 Then ' βρισκω το νεο ID_NUM
                Gdb.Execute ent(k)
                r2.Open "select max(ID_NUM) FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
                MAXID = str(r2(0))
             Else
                ' ΑΛΛΑΖΩ ΤΟ ID_NUM ΣΤΙΣ ΕΓΓΡΑΦΕΣ ΤΟΥ EGGTIM
                ent(k) = Replace(ent(k), palio_ID_NUM, MAXID)
                Gdb.Execute ent(k)
             End If
             
          Next
        
        Gdb.CommitTrans
        MsgBox "OK ΕΠΑΝΑΦΕΡΘΗΚΕ"
        Unload Me
        
        
      
      End If
   Else
     MsgBox "δεν βρεθηκε"
   End If
      
   Exit Sub
   
STOP_RESTORE:
   Gdb.RollbackTrans
   MsgBox "Δεν εγινε η επαναφορά " + Chr(13) + Err.Description
   
   
End Sub

Private Sub Command5_Click()
 Dim R As New ADODB.Recordset
  On Error Resume Next
  R.Open "select * from LOGGING WHERE  ENTOLH LIKE '% TIM %' AND ENTOLH LIKE '%" + HMER.Text + "%' AND ENTOLH LIKE '%" + ATIM.Text + "%' order by ID", Gdb, adOpenDynamic, adLockOptimistic
   
   Dim FOUND As Boolean
   FOUND = False
   Dim ent(1000) As String
   Dim F As Integer
   F = 0
   Me.MousePointer = vbHourglass
   List1.Clear
    ' Dim txtLines() As String
    ' txtLines = Split(R!Entolh, ",")
         
   Do While Not R.EOF
   
     ' txtLines = Split(R!Entolh, ",")
      CC = GET_SPLIT(R!Entolh, ",", 55)
      List1.AddItem str(R!ID) + " ; " + Format(R!hme, "dd/mm/YYYY HH:MM") + " ; " + CC + " ; " + R!Entolh
      
 
      FOUND = True
      F = F + 1
      ent(F) = R!Entolh
      R.MoveNext
   Loop
   Me.MousePointer = vbNormal
   EPANAFORA.Visible = True
   
   
   
End Sub

Function GET_SPLIT(c As String, DELIM As String, Position As Integer) As String
  Dim txtLines() As String
   txtLines = Split(c, DELIM)
   GET_SPLIT = txtLines(Position)
End Function


Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl
110     DTPicker1.Value = gEnarjh
 
120     DTPicker2.Value = Now

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU13.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU13.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 11100, 9050, fh, fw, ft, fl

End Sub

Private Sub List1_Click()
Dim a As String
a = List1.List(List1.ListIndex)

 Dim txtLines() As String
 txtLines = Split(a, ";")

 f_id = Val(txtLines(0))
 

End Sub
