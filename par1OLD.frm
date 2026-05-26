VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form par1OLD 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11130
   WindowState     =   2  'Maximized
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "par1OLD.frx":0000
      Height          =   3855
      Left            =   0
      OleObjectBlob   =   "par1OLD.frx":0012
      TabIndex        =   24
      Top             =   1200
      Width           =   10335
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "par1OLD.frx":09E9
      Left            =   6960
      List            =   "par1OLD.frx":09EB
      TabIndex        =   14
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Data dok 
      Caption         =   "dok"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Eid"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.CommandButton metasxhmatismos 
      Caption         =   "Μετασχηματισμός"
      Height          =   660
      Left            =   6960
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Data eggtim 
      Caption         =   "Data1"
      Connect         =   "dBASE IV;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "EGGTIM"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CommandButton b_akyr 
      Caption         =   "Εξοδος"
      Height          =   660
      Left            =   9120
      TabIndex        =   10
      Top             =   6600
      Width           =   1140
   End
   Begin VB.CommandButton b_Katax 
      Caption         =   "Καταχώρηση"
      Height          =   660
      Left            =   7920
      TabIndex        =   9
      Top             =   6600
      Width           =   1140
   End
   Begin VB.Data pel 
      Caption         =   "pel"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   4845
      TabIndex        =   4
      Top             =   270
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Text            =   "Combo3"
      Top             =   270
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "par1OLD.frx":09ED
      Left            =   1560
      List            =   "par1OLD.frx":09FA
      TabIndex        =   2
      Top             =   270
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "par1OLD.frx":0A1E
      Left            =   120
      List            =   "par1OLD.frx":0A2B
      TabIndex        =   1
      Top             =   270
      Width           =   1455
   End
   Begin VB.Data EID 
      Caption         =   "EID"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Data TIMEDIT 
      Caption         =   "TIMEDIT"
      Connect         =   "dBASE IV;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Timedit"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "par1OLD.frx":0A4F
      Height          =   3900
      Left            =   0
      OleObjectBlob   =   "par1OLD.frx":0A65
      TabIndex        =   8
      Top             =   1185
      Width           =   10335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      Caption         =   "Είδος Κίνησης         Παραστατικό                                  Σειρά         Αριθμός  "
      Height          =   1170
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10275
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6960
         TabIndex        =   19
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   5040
         TabIndex        =   18
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   17
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   7680
         TabIndex        =   15
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60751873
         CurrentDate     =   38294
      End
      Begin VB.CommandButton selectPEL 
         Caption         =   "Ανεύρεση Πελάτη"
         Height          =   345
         Left            =   5970
         TabIndex        =   6
         Top             =   255
         Width           =   1590
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Διεύθυνση"
         Height          =   255
         Left            =   6960
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Α.Φ.Μ."
         Height          =   255
         Left            =   5040
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Επώνυμο"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδικός"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "par1OLD.frx":2804
      Height          =   2145
      Left            =   0
      OleObjectBlob   =   "par1OLD.frx":2816
      TabIndex        =   0
      Top             =   5160
      Width           =   6885
   End
   Begin VB.Label SynoloG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8520
      TabIndex        =   12
      Top             =   6000
      Width           =   1470
   End
   Begin VB.Label sfpa 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8520
      TabIndex        =   11
      Top             =   5520
      Width           =   1470
   End
   Begin VB.Label Synolo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8520
      TabIndex        =   7
      Top             =   5100
      Width           =   1470
   End
End
Attribute VB_Name = "par1OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FlagDbgrid2
Dim f_matim As String
Dim f_dat As Date
Dim f_mfpa(1 To 9) As Single 'οριζονται στο find_parastat
Dim f_kau(1 To 9) As Single 'καθαρή αξία κατά φπα
Dim F_TITLOS, F_POS_APOU, F_AJIA_APOU, f_pel, f_metasx 'παράμετροι του παραστατικού
Dim f_deltia(1 To 30) 'δελτια που σημαδεύω
Dim f_suma ' teliko poso timologioy
'Dim f_epitr_seir
Dim fDB As Database  ' ανοιγω γενικά την database και με την open_data ανοίγω τον πίνακα
Dim ProthFora As Boolean
Dim fSHMADEMENA As String
Dim mEidos_Par(50) 'ειδος παραστατικού . π.χ.  atim=T
'

Private Sub b_akyr_Click()
  Unload Me
  
  
  
End Sub
Private Sub b_Katax_Click()
'====================================================
Dim wrkDefault As Workspace
Dim R As Recordset
Dim x, dum
Dim f_atim
'ελεγχος για άδειο παραστατικο
x = Open_Data("select sum(poso) as ss from timedit")

If IsNull(x("ss")) Or x("SS") = 0 Then
   MsgBox "Δεν περάστηκαν είδη"
   Exit Sub
End If

If Label1 = "" Then
    MsgBox "Δεν Διαλέξατε " + IIf(gbuff = "e", "Πελάτη", "Προμηθευτή")
    Exit Sub
End If

If Val(Synolo.Caption) = 0 Then Exit Sub




f_matim = Right(Combo2.Text, 1) + Right("00000" + Trim(Text1.Text), 5) + Combo3.Text

   ' Get default Workspace.
   Set wrkDefault = DBEngine.Workspaces(0)
' Start of outer transaction.
   wrkDefault.BeginTrans

'δημιουργία eggtim και ΤΙΜ
enhm_eggtim

'ενημέρωση αποθήκης
enhm_apot
'ενημέρωση πελάτη
enhm_pel

'--------------   ενημέρωση μετρητων  ----------------------
If Combo4.Text = Combo4.List(0) Then ' METRHTOIS
   enhm_met
End If



'ENHMERVSH ARIUMOY PARASTATIKOY
Upd_Parastat

If Len(fSHMADEMENA) > 5 Then
  fDB.Execute "UPDATE tim set art='1' where atim+Format(hme, 'dd/mm/yyyy')  in (" + fSHMADEMENA + ");"
End If




        ' Ask if the user wants to commit to all the changes
        ' made above.
        If MsgBox("Nα αποθηκευθεί το παραστατικό", vbYesNo) = vbYes Then
            wrkDefault.CommitTrans
        Else
            wrkDefault.Rollback
        End If

'εκτύπωση
'dum = print_timol(f_matim, CDate(Text2.Text))



    b_Katax.Enabled = False
    



End Sub
Sub elegxoi()
'================================================================
'=======   DEMO TA  TRANSACTIONS  ===============================
'================================================================

   Dim strName As String
   Dim strMessage As String
   Dim wrkDefault As Workspace
   Dim dbsNorthwind As Database
   Dim rstEmployees As Recordset

   ' Get default Workspace.
   Set wrkDefault = DBEngine.Workspaces(0)
   Set dbsNorthwind = OpenDatabase("Northwind.mdb")
   Set rstEmployees = _
      dbsNorthwind.OpenRecordset("Employees")

   ' Start of outer transaction.
   wrkDefault.BeginTrans
   ' Start of main transaction.
   wrkDefault.BeginTrans

   With rstEmployees

      ' Loop through recordset and ask user if she wants to
      ' change the title for a specified employee.
      Do Until .EOF
         If !Title = "Sales Representative" Then
            strName = !LastName & ", " & !FirstName
            strMessage = "Employee: " & strName & vbCr & _
               "Change title to Account Executive?"

            ' Change the title for the specified employee.
            If MsgBox(strMessage, vbYesNo) = vbYes Then
               .Edit
               !Title = "Account Executive"
               .Update
            End If
         End If

         .MoveNext
      Loop

      ' Ask if the user wants to commit to all the changes
      ' made above.
      If MsgBox("Save all changes?", vbYesNo) = vbYes Then
         wrkDefault.CommitTrans
      Else
         wrkDefault.Rollback
      End If

      ' Print current data in recordset.
      .MoveFirst
      Do While Not .EOF
         Debug.Print !LastName & ", " & !FirstName & _
            " - " & !Title
         .MoveNext
      Loop

      ' Roll back any changes made by the user since this is
      ' a demonstration.
      wrkDefault.Rollback
      .Close
   End With

   dbsNorthwind.Close

'End Sub



'1.να μην υπάρχουν ΦΠΑ>9 και <=0
End Sub
Sub Upd_Parastat()

Dim K As Integer
Dim db As Database
Dim R As Recordset, x As String
Dim dum As Integer


For K = 1 To Combo3.ListCount
   If Combo3.Text = Combo3.List(K) Then
      Exit For
   End If
Next


If K <= Combo3.ListCount Then
   dum = find_parastat(K, Right(Combo2.Text, 1), 1)
End If

End Sub


Sub enhm_eggtim()
'=========================================
'δημιουργία eggtim και ΤΙΜ
'=========================================
Dim R, re, rt As Recordset, db As Database
Dim K, l, M, n As Integer
f_dat = Now

'On Error GoTo 0

For K = 1 To 9: f_kau(K) = 0: Next
  
  
'  If gConnect = "Access" Then
'     Set db = OpenDatabase(gDir, False, False)
'  Else
     Set db = OpenDatabase(gDir, False, False, gConnect)
'  End If
  
  Set re = db.OpenRecordset("eggtim", dbOpenTable)
  Set rt = db.OpenRecordset("tim", dbOpenTable)


'------------------------------eggtim -----------------------------------------
Set R = TIMEDIT.Recordset

On Error Resume Next
R.MoveFirst
Do While Not R.EOF
  If Not IsNull("poso") Then
   If R("poso") > 0 Then
'     Set re = db.OpenRecordset("eggtim", dbOpenTable, dbAppendOnly, dbOptimistic)
'     On Error GoTo 0
     re.AddNew
        re("hme") = f_dat: re("atim") = f_matim: re("poso") = R("poso"): re("timm") = R("timm")
        re("kode") = R("kode")
        re("onoma") = R("onoma"): re("kau_ajia") = R("kau_ajia"): re("fpa") = R("fpa")
        f_kau(R("fpa")) = f_kau(R("fpa")) + re("kau_ajia")  'αξια κατα συντελεστή
        re("pelkod") = pel.Recordset("kod")
     re.Update
       End If
  End If
  R.MoveNext
Loop
'------------------------------ TIM -----------------------------------------
're.Close

rt.AddNew
  rt("hme") = f_dat
  rt("atim") = f_matim
  f_suma = 0
  For K = 1 To 9
     rt("aj" + LTrim(Str(K))) = Round(f_kau(K), 2)
     f_suma = f_suma + rt("aj" + LTrim(Str(K)))
     If K <> 5 Then
        rt("fpa" + LTrim(Str(K))) = Round(f_mfpa(K) * f_kau(K), 2)
        f_suma = f_suma + rt("fpa" + LTrim(Str(K)))
     End If
  Next
  rt("kpe") = pel.Recordset("kod")
  rt("aji") = Round(f_suma, 2)
rt.Update


End Sub

'ενημέρωση αποθήκης
Sub enhm_apot()
Dim db As Database, R As Recordset, e As Recordset
Dim G, M, a, GA, MA, P
G = "G" + Format(Month(Date$), "00")
M = "M" + Format(Month(Date$), "00")
GA = "GA" + Format(Month(Date$), "00")
MA = "MA" + Format(Month(Date$), "00")

Set R = TIMEDIT.Recordset
Set db = OpenDatabase(gDir, False, False, gConnect)
Set e = db.OpenRecordset("EID", dbOpenTable)

R.MoveFirst
'e.Index = "EIDko"

Do While Not R.EOF
  
  
  'e.Seek "=", R("KODE")
  If Not IsNull(R("kode")) Then
     Set e = db.OpenRecordset("select *from EID where kod='" + R("kode") + "'")
  'End If
  
  If e.RecordCount > 0 Then
    P = R("POSO")
    a = R("POSO") * R("TIMM")
    e.Edit
    Select Case F_POS_APOU
    Case "1" ' EISAGOGES=1
         e("POS") = e("POS") + P
         e(G) = e(G) + P
    Case "2" ' -EISAGOGES
         e("POS") = e("POS") - P
         e(G) = e(G) - P
    Case "3" ' EXAGOGES=1
         e("POS") = e("POS") - P
         e(M) = e(M) + P
    Case "4" ' -EXAGOGES
         e("POS") = e("POS") + P
         e(M) = e(M) - P
    End Select
    
    Select Case F_AJIA_APOU
    Case "1" ' EISAGOGES=1
       '  e(GA) = e(GA) + a
    Case "2" ' -EISAGOGES
'         e(GA) = e(GA) - a
    Case "3" ' EXAGOGES=1
'         e(MA) = e(MA) + a
    Case "4" ' -EXAGOGES
         e(MA) = e(MA) - a
    End Select
    
    
    
    
    e.Update
  
  End If
  End If  ' R("kode")=null
  
  R.MoveNext
Loop



End Sub
'ενημέρωση πελάτη
Sub enhm_pel()
'----------------------------------------------------------------------
Dim db As Database
Dim R As Recordset, x As String

If gConnect = "Access" Then
   Set db = OpenDatabase(gDir, False, False)
Else
   Set db = OpenDatabase(gDir, False, False, gConnect)
End If
Set R = db.OpenRecordset("egg")





pel.Recordset.Edit

If pel.Recordset("eidos") = "e" Then
   If f_pel = "1" Then
      pel.Recordset("typ") = pel.Recordset("typ") + f_suma
   End If
   If f_pel = "2" Then
      pel.Recordset("typ") = pel.Recordset("typ") - f_suma
   End If
End If
If pel.Recordset("eidos") = "r" Then
   If f_pel = "1" Then
      pel.Recordset("typ") = pel.Recordset("typ") + f_suma
   End If
   If f_pel = "2" Then
      pel.Recordset("typ") = pel.Recordset("typ") - f_suma
   End If
End If
pel.Recordset.Update


R.AddNew
   R("hme") = DTPicker1.Value
   R("atim") = f_matim
   R("apa") = Val(Mid$(f_matim, 2, 6))
   
   R("ait") = Left(Combo2.Text, 20)
   
   R("xre") = Round(f_suma, 2)
   R("eidos") = pel.Recordset("eidos")
   R("kod") = pel.Recordset("kod")
R.Update
R.Close

End Sub


Sub enhm_met()
'
'ενημέρωση μετρητών
'----------------------------------------------------------------------

Dim db As Database
Dim R As Recordset, x As String

If gConnect = "Access" Then
   Set db = OpenDatabase(gDir, False, False)
Else
   Set db = OpenDatabase(gDir, False, False, gConnect)
End If
Set R = db.OpenRecordset("egg")





pel.Recordset.Edit

If pel.Recordset("eidos") = "e" Then
   
   If f_pel = "1" Then
      pel.Recordset("typ") = pel.Recordset("typ") - f_suma
   End If
   
   If f_pel = "2" Then
      pel.Recordset("typ") = pel.Recordset("typ") + f_suma
   End If
   
End If

If pel.Recordset("eidos") = "r" Then
   If f_pel = "1" Then
      pel.Recordset("typ") = pel.Recordset("typ") - f_suma
   End If
   If f_pel = "2" Then
      pel.Recordset("typ") = pel.Recordset("typ") + f_suma
   End If
End If
pel.Recordset.Update


R.AddNew
   
   R("hme") = DTPicker1.Value
   R("atim") = f_matim
   R("apa") = Val(Mid$(f_matim, 2, 6))
   R("eid") = "ΜΕ"
   R("ait") = Left(Combo2.Text, 20)
   
   R("xre") = Round(-f_suma, 2)
   R("eidos") = pel.Recordset("eidos")
   R("kod") = pel.Recordset("kod")
   
R.Update

R.Close

End Sub

'εκτύπωση
Function print_timol(ByVal atim As String, ByVal hme As Date)

Dim dum

'dum = ektyp_forma("f90", atim, hme)
10


dum = ektyp_forma("c:\mercvb\f90.txt", atim, hme)
'
'GoTo 10

End Function


Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Combo2.SetFocus
 
End Sub

Private Sub Combo1_LostFocus()
' ANALOGA ME TI DIALEJA ΑΝΟΙΓΕΙ ΤΟ PARASTAT KAI GEMIZEI TO COMBO2

Dim db As Database
Dim R As Recordset, x As String, K As Integer


If gConnect = "Access" Then
   Set db = OpenDatabase(gDir, False, False)
Else
   Set db = OpenDatabase(gDir, False, False, gConnect)
End If

If Combo1.Text = Combo1.List(0) Then
    x = "2": gbuff = "r" 'προμηθευτές
  ElseIf Combo1.Text = Combo1.List(1) Then
    x = "1": gbuff = "e" 'pelates
  Else
    x = "3": gbuff = "l" 'λογαριασμοί
End If

Set R = db.OpenRecordset("select pol,titlos,eidos from parastat where pol='" + x + "'")
Combo2.Clear
  R.MoveFirst
  K = 0
  Do While Not R.EOF
     If Not IsNull(R("TITLOS")) And R("pol") = x Then
        Combo2.AddItem R("TITLOS") + "." + R("EIDOS") ' + Space(30) + R("eidos")
        K = K + 1
        mEidos_Par(K) = R("eidos")
     End If
     R.MoveNext
  Loop
Combo2.Enabled = True


End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Combo3.SetFocus
End Sub

Private Sub Combo2_LostFocus()
' STELNEI STO FIND_PARASTAT NA BREI TON A/A KAI PAIRNEI TIS PARAMETROYS APO TO PARASTAT

Dim K As Integer
Dim db As Database
Dim R As Recordset, x As String


If gConnect = "Access" Then
   Set db = OpenDatabase(gDir, False, False)
Else
   Set db = OpenDatabase(gDir, False, False, gConnect)
End If
Set R = db.OpenRecordset("parastat")


  R.MoveFirst
  Do While Not R.EOF
     If Right(Combo2.Text, 1) = R("EIDOS") Then
       F_TITLOS = R("TITLOS")
       F_POS_APOU = R("POS_APOU")
       F_AJIA_APOU = R("AJIA_APOU")
       f_pel = R("pel")
       f_metasx = R("metasx") ' απο ποιό παραστατικό μετασχηματίζεται
     End If
     R.MoveNext
  Loop
'Combo2.Clear
geidos_paras = Combo2.ListIndex
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text1.SetFocus
End Sub

Private Sub Combo3_LostFocus()
' STELNEI STO FIND_PARASTAT NA BREI TON A/A KAI PAIRNEI TIS PARAMETROYS APO TO PARASTAT

Dim K As Integer
Dim db As Database
Dim R As Recordset, x As String


For K = 1 To Combo3.ListCount
   If Combo3.Text = Combo3.List(K) Then
      Exit For
   End If
Next


If K <= Combo3.ListCount Then
  ' Text1.Text = find_parastat(k, Right(Combo2.Text, 1), 0)
    x = mEidos_Par(1 + Combo2.ListIndex)
    Text1.Text = find_parastat(K, x, 0)
End If



If gConnect = "Access" Then
   Set db = OpenDatabase(gDir, False, False)
Else
   Set db = OpenDatabase(gDir, False, False, gConnect)
End If
Set R = db.OpenRecordset("parastat")


  R.MoveFirst
  Do While Not R.EOF
     If Right(Combo2.Text, 1) = R("EIDOS") Then
       F_TITLOS = R("TITLOS")
       F_POS_APOU = R("POS_APOU")
       F_AJIA_APOU = R("AJIA_APOU")
     End If
     R.MoveNext
  Loop
  R.Close
  
'Combo2.Clear
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   b_Katax.SetFocus
End If

End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If LastCol = 5 Then
       
       DBGrid1.Row = DBGrid1.Row + 1: DBGrid1.Col = 0
   End If
End Sub

Private Sub DBGrid2_Click()
Dim delt As Recordset

If DBGrid1.Enabled = False Then ' φάση δελτίων αποστολής
   FlagDbgrid2 = True
  If ProthFora = False Then
     ProthFora = True
     Set delt = EID.Recordset
     If IsNull(delt("metaf")) Or delt("metaf") = 0 Then
        delt.Edit: delt("metaf") = 1: delt.Update
     Else
        delt.Edit: delt("metaf") = 0: delt.Update
     End If
  End If
Else
   DBGrid1.Col = 2
End If

End Sub

Private Sub DBGrid2_KeyPress(KeyAscii As Integer)
Dim a, SS
'a = 1

If KeyAscii = 13 Then
   
   
   TIMEDIT.Recordset.Edit
   TIMEDIT.Recordset("kode") = Left(EID.Recordset("xx"), 13)
   TIMEDIT.Recordset("onoma") = Mid$(EID.Recordset("xx"), 14, 30)
   
   SS = DBGrid1.Bookmark
   
   DBGrid1.Columns(0).Value = Left(EID.Recordset("xx"), 13)
   DBGrid1.Columns(1).Value = Mid$(EID.Recordset("xx"), 14, 25)
   If gbuff = "r" Then
      TIMEDIT.Recordset("timm") = EID.Recordset("xti")
   Else
      TIMEDIT.Recordset("timm") = EID.Recordset("lti")
   End If
   TIMEDIT.Recordset("fpa") = EID.Recordset("fpa")
   TIMEDIT.Recordset.Update
   
   DBGrid1.Refresh
   
   TIMEDIT.Refresh
   
  
   DBGrid1.Bookmark = SS
   DBGrid1.Col = 2
   
   DBGrid1.SetFocus
   
   '          gKod_pel = gbuff + Right(Trim(Data1.Recordset("xx")), 4)
   '         Unload Me
 End If





'
End Sub

Private Sub DBGrid3_DblClick()
Dim a, e
   
   gKod_pel = gbuff + Right(Trim(pel.Recordset("xx")), 4)
   
   
   
   'A = pel.Recordset("KOD")
   'E = Data1.Recordset("EIDOS")
   pel.RecordSource = "select *from pel where eidos='" + Left(gKod_pel, 1) + "' and kod='" + Mid$(gKod_pel, 2, Len(gKod_pel) - 1) + "'"
   pel.Refresh
   
   Text2(0).Text = pel.Recordset("KOD")
   Text2(1).Text = pel.Recordset("EPO")
   DBGrid3.Visible = False
   DBGrid1.SetFocus
   
   
End Sub

Private Sub DBGrid3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  
    DBGrid3_DblClick

End If
End Sub

Private Sub metasxhmatismos_Click()
'===================================
'συγκεντρώνω τα δελτία που σημάδεψε ο χρήστης
'===================================
Dim delt As Recordset, K As Integer, l As Integer
Dim e, s As String, d As String, t As Recordset
Dim x, c, Eidh As Recordset
Dim R As Recordset

K = 0
Set delt = EID.Recordset
delt.MoveFirst
s = ""
Do While Not delt.EOF
  If delt("metaf") = 1 Then
     K = K + 1
     f_deltia(K) = delt("atim") + Format(delt("hme"), "dd/mm/yyyy")
     s = s + "'" + f_deltia(K) + "'" + ","
  End If
  delt.MoveNext
Loop
s = s + "''"
fSHMADEMENA = s
d = EID.RecordSource
' σημάδεψε δελτια // ψάχνω να βρώ τα είδη
EID.RecordSource = "select sum(poso) as sposo,kode from eggtim where atim+Format(hme, 'dd/mm/yyyy')  in (" + s + ") group by kode;"
EID.Refresh

If EID.Recordset.RecordCount > 20 Then
   MsgBox "Πάνω από 20 είδη στο τιμολόγιο. Διαλέξτε λιγότερα δελτία"
   EID.RecordSource = d
   EID.Refresh
   Exit Sub
End If

'dok.Connect = gConnect
'dok.DatabaseName = gDir
'dok.RecordSource = "eid"
'dok.Recordset.Index = "EIDKO"



Set Eidh = fDB.OpenRecordset("eid")
Eidh.Index = "eidko"





Set t = TIMEDIT.Recordset
If EID.Recordset.RecordCount = 0 Then
   EID.RecordSource = "SELECT *FROM EID"
   EID.Refresh
   Exit Sub
End If


EID.Recordset.MoveFirst
t.MoveFirst
Do While Not EID.Recordset.EOF
  Eidh.Seek "=", EID.Recordset("kode")
  t.Edit
  
  t("kode") = EID.Recordset("kode")
  t("poso") = EID.Recordset("sposo")
  t("onoma") = Eidh("ono")
  t("timm") = Eidh("lti")
  t("fpa") = Eidh("fpa")
  t("kau_ajia") = t("poso") * t("timm")
  t("mik_ajia") = t("kau_ajia") + t("kau_ajia") * IIf(t("fpa") > 0, g_Fpa(t("fpa")), 18) / 100
  t.Update
  
  t.MoveNext
  EID.Recordset.MoveNext
Loop
DBGrid1.Enabled = True
't.Close


Set R = fDB.OpenRecordset("select SUM(KAU_AJIA) AS K,sum(mik_ajia) as mik from TIMEDIT")
  Synolo.Caption = Format(R("k"), gDek_ajia)
  sfpa.Caption = Format(R("mik") - R("k"), gDek_ajia)
  SynoloG.Caption = Format(R("mik"), gDek_ajia)




metasxhmatismos.Enabled = False
EID.RecordSource = "SELECT *FROM EID"
EID.Refresh

End Sub

Private Sub selectPEL_Click()
' Διαλέγει πελάτη και δείχνει δελτία σε εκκρεμότητα
'===================================================
  DBGrid3.Visible = True
  Text2(0).Enabled = True
  Text2(1).Enabled = True
  Text2(2).Enabled = True
  Text2(0).SetFocus



' αν είναι τιμολόγιο του δείχνει τα δελτία σε εκκρεμότητα
'If Len(f_metasx) > 0 Then
'
'  EID.RecordSource = "select metaf,atim,hme from tim where left(atim,1)='" + f_metasx + "' and kpe='" + pel.Recordset("kod") + "' and art=null;"
'  EID.Refresh
'  DBGrid1.Enabled = False
'  metasxhmatismos.Visible = True
'  FlagDbgrid2 = False
'End If
  
  
End Sub

Private Sub DBGrid1_AfterColEdit(ByVal ColIndex As Integer)
'---------------------------------------
' ANAZHTHSH TOY ΕΙΔΟΥΣ ΜΕΣΑ ΣΤΟ ΤΙΜΟΛΟΓΙΟ

Dim db As Database, R As Recordset, t As Recordset

Dim l As Integer

Set t = TIMEDIT.Recordset


' ANAZHTHSH TOY KODIKOY EIDOYS
If ColIndex = 0 Then
  EID.RecordSource = "select LEFT(KOD+'                 ',14)+ONO as xx,xti,lti,fpa from EID WHERE KOD LIKE '" + DBGrid1.Columns(0).Text + "*';"
  EID.Refresh
  DBGrid2.Columns(0).DataField = "xx"
  DBGrid2.Columns(0).Width = 6000
'  DBGrid2.Columns(0).FontName = "COURIER"
  DBGrid2.Columns(1).Visible = False
  DBGrid2.Columns(2).Visible = False
  DBGrid2.Columns(3).Visible = False
  DBGrid2.SetFocus
  DBGrid1.Col = 2
  
  
ElseIf ColIndex = 1 Then ' ΚΑΤΑ ΟΝΟΜΑ
  l = Len(DBGrid1.Columns(1).Text)
  If l > 0 Then
     EID.RecordSource = "select LEFT(KOD+'                ',14)+ONO  as xx,xti,lti,fpa from EID WHERE LEFT(ONO," + Str(l) + ")='" + DBGrid1.Columns(1).Text + "';"
     ' LIKE '" + DBGrid1.Columns(1).Text + "*';"
     EID.Refresh
    ' DBGrid2.FontΝΑΜΕ = "COURIER"
     DBGrid2.Columns(0).DataField = "xx"
     DBGrid2.Columns(0).Width = 6000
     DBGrid2.Columns(1).Visible = False
     DBGrid2.Columns(2).Visible = False
     DBGrid2.Columns(3).Visible = False
     
     DBGrid2.SetFocus
     
  End If
  
  ' υπολογισμός αξίας
ElseIf ColIndex = 2 Or ColIndex = 3 Or ColIndex = 4 Then




   DBGrid1.Columns(6).Text = Round(Val(DBGrid1.Columns(2).Text) * Val(DBGrid1.Columns(3).Text) * (100 - Val(DBGrid1.Columns(4).Text)) / 100, 2)
   
   'If InStr(DBGrid1.Columns(COLLINDEX).Text, ",") > 0 Then DBGrid1.Columns(COLLINDEX).Text = 0
   'If InStr(DBGrid1.Columns(3).Text, ",") > 0 Then DBGrid1.Columns(3).Text = 0
   'If InStr(DBGrid1.Columns(4).Text, ",") > 0 Then DBGrid1.Columns(4).Text = 0
   DBGrid1.Columns(ColIndex).Text = Val(DBGrid1.Columns(ColIndex).Text)
   
   
   t.Edit
   t("kau_ajia") = (DBGrid1.Columns(6).Text) ' * (DBGrid1.Columns(3).Text)
   t("mik_ajia") = t("kau_ajia") + t("kau_ajia") * IIf(t("fpa") > 0, g_Fpa(t("fpa")), 18) / 100 ' * (DBGrid1.Columns(3).Text)
   t.Update
  
  If gConnect = "Access" Then
      Set db = OpenDatabase(gDir, False, False)
  Else
      Set db = OpenDatabase(gDir, False, False, gConnect)
  End If
  
  Set R = db.OpenRecordset("select SUM(KAU_AJIA) AS K,sum(mik_ajia) as mik from TIMEDIT")
  Synolo.Caption = Format(R("k"), gDek_ajia)
  sfpa.Caption = Format(R("mik") - R("k"), gDek_ajia)
  SynoloG.Caption = Format(R("mik"), gDek_ajia)

ElseIf ColIndex = 5 Then ' pane sthn epomenh grammh
  If InStr(DBGrid1.Columns(5).Text, ",") > 0 Then DBGrid1.Columns(5).Text = 2


  DBGrid1.Col = 0

End If

End Sub

Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'==================================
'επιλογή είδους ή δελτίων αποστολής
'==================================
Dim delt As Recordset

  On Error GoTo DBGrid2_RowColChange2:

If DBGrid1.Enabled = False Then ' φάση δελτίων αποστολής
 
 If FlagDbgrid2 Then
   ' FlagDbgrid2 = False
    Set delt = EID.Recordset
    If IsNull(delt("metaf")) Or delt("metaf") = 0 Then
       delt.Edit: delt("metaf") = 1: delt.Update
    Else
       delt.Edit: delt("metaf") = 0: delt.Update
    End If
 End If
  
Else

'   TIMEDIT.Recordset.Edit
'   TIMEDIT.Recordset("kode") = EID.Recordset("kod")
'   TIMEDIT.Recordset("onoma") = EID.Recordset("ono")
'   If gbuff = "r" Then
'      TIMEDIT.Recordset("timm") = EID.Recordset("xti")
'   Else
'      TIMEDIT.Recordset("timm") = EID.Recordset("lti")
'   End If
'   TIMEDIT.Recordset("fpa") = EID.Recordset("fpa")
'   TIMEDIT.Recordset.Update
'   DBGrid1.Refresh
   
End If
   

   
   
   Exit Sub
DBGrid2_RowColChange2:
   Exit Sub
End Sub

Private Sub Form_Activate()
'==============================
'==============================
'==============================
'==============================

Dim K, l As Integer, x As String, db, R

ProthFora = False

DTPicker1.Value = Now


If gConnect = "Access" Then
  Set fDB = OpenDatabase(gDir, False, False)
Else
  Set fDB = OpenDatabase(gDir, False, False, gConnect)
End If

Synolo = " " ' ΓΙΑ ΝΑ ΜΗΝ ΒΓΑΖΕΙ NULL


'
'If gConnect = "Access" Then
'  Set DB = OpenDatabase(gDir, False, False)
'Else
'  Set DB = OpenDatabase(gDir, False, False, gConnect)
'End If
'Set r = DB.OpenRecordset("select *from mem")
'r.MoveFirst
'r.Move 3
'
''r.MoveNext
'gEPITR_SEIR = r("c3")
'r.Close


'For k = 1 To Len(gEPITR_SEIR)
'    Combo1.AddItem Mid$(gEPITR_SEIR, k, 1)
'Next


TIMEDIT.DatabaseName = gDir
TIMEDIT.Connect = gConnect

EID.DatabaseName = gDir
EID.Connect = gConnect

pel.DatabaseName = gDir
pel.Connect = gConnect
    
    
TIMEDIT.Enabled = True
EID.Enabled = True
pel.Enabled = True
    
'Text2.Text = Date$
    
    

TIMEDIT.Refresh
 If TIMEDIT.Recordset.RecordCount < 40 Then
    For K = 1 To 40
      TIMEDIT.Recordset.AddNew
      TIMEDIT.Recordset.Update
    Next
 Else
    For K = 1 To 40
      TIMEDIT.Recordset.Edit
      For l = 0 To TIMEDIT.Recordset.Fields.Count - 1
              TIMEDIT.Recordset(l) = Null
      Next
      TIMEDIT.Recordset.Update
      TIMEDIT.Recordset.MoveNext
    Next
 End If
 TIMEDIT.Recordset.MoveFirst
 
 
 Combo3.Clear
 x = find_seires(0)
 For K = 1 To Len(x)
    Combo3.AddItem Mid$(x, K, 1)
 Next
 
 Combo4.Clear
 Combo4.AddItem "ΜΕΤΡΗΤΟΙΣ"
 Combo4.AddItem "ΠΙΣΤΩΣΗ"
 Combo4.AddItem "ΠΙΣΤ.ΚΑΡΤΑ"
 Combo4.Text = Combo4.List(0)
 
 
 
 Combo1.Text = Combo1.List(1)
 Combo1_LostFocus
 
Combo2.Text = Combo2.List(geidos_paras)
 
 
 DBGrid1.Columns(2).NumberFormat = gDek_pos 'δεκαδικά ποσοτήτων
 DBGrid1.Columns(3).NumberFormat = gDek_ajia 'δεκαδικά αξίας
 DBGrid1.Columns(6).NumberFormat = gDek_ajia 'δεκαδικά αξίας
 
 Combo1.SetFocus
 
 
 
End Sub

Function find_parastat(recno As Integer, par As String, ReadWrite As Integer) As Long
'ReadWrite=0   Βρίσκει τον αριθμο παραστατικού          -------------------------
'ReadWrite=1   Aποθηκεύει τον αριθμο παραστατικού + 1   -------------------------
'Βρίσκει τις κατηγορίες φπα

Dim db As Database, K As Integer
Dim R As Recordset, x, Bohuhtiko As String

Bohuhtiko = "**"

If gConnect = "Access" Then
  Set db = OpenDatabase(gDir, False, False)
Else
  Set db = OpenDatabase(gDir, False, False, gConnect)
End If


Set R = db.OpenRecordset("select *from mem")
R.MoveFirst
For K = 1 To 9: f_mfpa(K) = R("fpa" + LTrim(Str(K))): Next
R.Move recno - 1
Select Case par
    Case "P"   ' πιστ. σε πελάτη
      Bohuhtiko = "ps"
    Case "p"    ' δελτ.επιστροφής
      Bohuhtiko = "de"
    Case "T"    ' τιμολόγιο και δελτίο αποστολής
      Bohuhtiko = "tp"
    Case "t"    ' τιμολόγιο σκέτο
      Bohuhtiko = "tpsketo"
    Case "Y"    ' παροχής υπηρεσιών
      Bohuhtiko = "tpy"
    Case "A"    ' δελτίο αποστολής σε πελάτη
      Bohuhtiko = "da"
    Case "e"    ' δελτίο αποστολής σε προμηθευτή
      Bohuhtiko = "da"
    Case "L"   ' λιανική - δα
      Bohuhtiko = "dl"
    Case "l"    ' λιανική
      Bohuhtiko = "dlSKETO"
    Case "k"    ' ενδοδιακίνηση
      Bohuhtiko = "endod"
    Case "R"   ' AGROTIKA
      Bohuhtiko = "AGROT"
   Case "y"   ' αποδ.παρ.υπηρεσιών
     Bohuhtiko = "dpy"
End Select

1 If ReadWrite = 0 Then
   If Bohuhtiko <> "**" Then find_parastat = R(Bohuhtiko) Else find_parastat = 0
Else
   If Bohuhtiko = "**" Then
   Else
      R.Edit: R(Bohuhtiko) = R(Bohuhtiko) + 1: R.Update
   End If
End If


db.Close
End Function
Function find_seires(dum As Integer) As String

'Βρίσκει τις υπάρχουσες σειρές
Dim db As Database
Dim R As Recordset, x As String

If gConnect = "Access" Then
   Set db = OpenDatabase(gDir, False, False)
Else
   Set db = OpenDatabase(gDir, False, False, gConnect)
End If
Set R = db.OpenRecordset("select c3 from mem")
R.MoveFirst
R.Move 3

If IsNull(R("c3")) Then
   find_seires = " "
Else
   find_seires = Trim(R("c3"))
End If
db.Close
End Function

Function Open_Data(ByVal mtable As String) As Recordset

 Set Open_Data = fDB.OpenRecordset(mtable)
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        selectPEL.SetFocus
      
   End If
End Sub


Private Sub Text2_LostFocus(Index As Integer)
Dim x As String

If Len(Trim(Text2(Index))) = 0 Then Exit Sub

x = gbuff 'If Option1(1) Then x = "e" Else x = "r"


If Index = 1 Then  'eponumo
    pel.RecordSource = "select LEFT(EPO+'-'+DIE+'-'+KOD,60) as xx from pel where eidos='" + x + "' and epo like '" + Text2(1).Text + "*';"
End If

If Index = 0 Then 'kodikos
    pel.RecordSource = "select  LEFT(EPO+'-'+DIE+'-'+KOD,60) as xx from pel where kod='" + x + "' and epo like '" + Text2(0).Text + "*';"
End If

If Index = 2 Then 'afm
    pel.RecordSource = "select  LEFT(EPO+'-'+DIE+'-'+KOD,60) AS XX from pel where eidos='" + x + "' and afm like '" + Text2(2).Text + "*';"
End If

DBGrid3.Columns(0).Width = 4000

pel.Refresh
DBGrid3.SetFocus


End Sub
