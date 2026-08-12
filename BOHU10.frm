VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form bohu10 
   BackColor       =   &H80000002&
   Caption         =   "Γεφυρα Λογιστικής"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9960
   ScaleWidth      =   18870
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGeferror 
      Caption         =   "geferror"
      Height          =   360
      Left            =   960
      TabIndex        =   27
      Top             =   8280
      Width           =   990
   End
   Begin VB.CommandButton CMDSTOP 
      Caption         =   "STOP"
      Height          =   360
      Left            =   15600
      TabIndex        =   26
      Top             =   2400
      Width           =   2895
   End
   Begin VB.CommandButton cmdPBSGEFYRA 
      Caption         =   "PBS ΓΕΦΥΡΑ"
      Height          =   360
      Left            =   11280
      TabIndex        =   24
      Top             =   1920
      Width           =   4215
   End
   Begin VB.ListBox List2 
      Height          =   5325
      Left            =   15600
      TabIndex        =   23
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdΓΚατηγορια 
      Caption         =   "Γ Κατηγορια ΕΙΣΠΡΑΞΕΙΣ"
      Height          =   360
      Left            =   11280
      TabIndex        =   22
      Top             =   2400
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Διαγραφή Προσωρινών Αρχείων (SQLTEMP)"
      Height          =   375
      Left            =   11280
      TabIndex        =   21
      Top             =   1440
      Width           =   4215
   End
   Begin VB.CommandButton bkat_kathgories 
      Caption         =   "ΥΠΟΛΟΓΙΣΜΟΣ Β ΚΑΤΗΓ ΜΕ ΚΑΤΗΓΟΡΙΕΣ ΕΙΔΩΝ"
      Height          =   360
      Left            =   6840
      TabIndex        =   20
      Top             =   1440
      Width           =   4335
   End
   Begin VB.CheckBox chkPOL 
      Caption         =   "ΠΩΛΗΣΕΙΣ"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkAG 
      Caption         =   "ΑΓΟΡΕΣ"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chk2 
      Caption         =   "Κωδικός Πελάτη 4 ψηφία"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txt22 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Text            =   "0000"
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Text            =   "30-01-01-"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdG 
      Caption         =   "Γ Κατηγορια"
      Height          =   360
      Left            =   6840
      TabIndex        =   7
      Top             =   2400
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   12600
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   960
      TabIndex        =   6
      Top             =   2880
      Width           =   8055
   End
   Begin VB.CheckBox chkEPAN 
      Caption         =   "Επανενημέρωση"
      Height          =   195
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker eos 
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Format          =   291438593
      CurrentDate     =   44346
   End
   Begin MSComCtl2.DTPicker apo 
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Format          =   291438593
      CurrentDate     =   44346
   End
   Begin VB.CommandButton cmdYPOL 
      BackColor       =   &H0080FF80&
      Caption         =   "ΥΠΟΛΟΓΙΣΜΟΣ Β ΚΑΤΗΓ ΧΩΡΙΣ ΚΑΤΗΓ.ΕΙΔΩΝ"
      Height          =   360
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frame1"
      Height          =   5325
      Left            =   9120
      TabIndex        =   14
      Top             =   2880
      Width           =   6375
      Begin VB.CommandButton cmdImport 
         Caption         =   "import Excel"
         Height          =   360
         Left            =   960
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdAntist 
         Caption         =   "antist"
         Height          =   360
         Left            =   3360
         TabIndex        =   18
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtText1 
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdCommand1 
         Caption         =   "επιλογη Excel"
         Height          =   360
         Left            =   3360
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdCommand2 
         Caption         =   "xml για Epsilon"
         Height          =   360
         Left            =   960
         TabIndex        =   15
         Top             =   1800
         Width           =   1455
      End
   End
   Begin VB.Label lblVER2016 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VER.2016"
      Height          =   195
      Left            =   9120
      TabIndex        =   25
      Top             =   8280
      Width           =   705
   End
   Begin VB.Label lblΚωδικόςΠελάτη 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδικός Πελάτη"
      Height          =   195
      Left            =   960
      TabIndex        =   10
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label lblΕως 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "εως"
      Height          =   195
      Left            =   6480
      TabIndex        =   4
      Top             =   1140
      Width           =   270
   End
   Begin VB.Label lblΑπο 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "απο"
      Height          =   195
      Left            =   6480
      TabIndex        =   3
      Top             =   540
      Width           =   270
   End
End
Attribute VB_Name = "bohu10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim PIN(30, 2)
Dim f_categ As Integer
Dim F_PEL_KOD As Integer '1=OPOY BREI ASTERAKIA BAZEI TON KODIKO  0=OXI
Dim xl       As New excel.Application

Dim xlsheet  As excel.Worksheet

Dim xlsheet3 As excel.Worksheet

Dim xlwbook  As excel.workbook

Dim f_kod_tameio As String
Dim arr_log(100) As String
Dim arr_posa(100) As Single

Private Sub Add_to_Array(log As String, kau As Single, MSIGN As Integer)
Dim k As Integer, FOUND As Integer
FOUND = 0
For k = 1 To 100
  If log = arr_log(k) Then 'το βρηκα βαζω το ποσο
       FOUND = 1
       arr_posa(k) = arr_posa(k) + (MSIGN * kau)
       Exit For
  End If
  If Len(arr_log(k)) = 0 Then  ' αδειο αρα το γεμιζω
       FOUND = 2 'neos logariasmos
       arr_posa(k) = (MSIGN * kau)
       arr_log(k) = log
       Exit For
  End If
Next




End Sub

Private Sub bkat_kathgories_Click()
    selectRecords 1, 1

End Sub

Private Sub cmdAntist_Click()


        '<EhHeader>
        On Error GoTo Label6_Click_Err

        '</EhHeader>
        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        'pinakas pedion
100     UPDATE2_PINAKES 101, lblΑπο  ', Combo3

        ' UPDATE_PINAKES3f 21, Label6  ', Combo1
        '<EhFooter>
        Exit Sub

Label6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu51.Label6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Label6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdCommand1_Click()
       Dim mF As String
       If Len(mF) < 2 Then
          MsgBox "δεν διαλεξατε αρχείο"
          Exit Sub
       End If
      txtText1.Text = mF
      
       
       
End Sub

Private Sub cmdCommand2_Click()

   Dim par As String
   
Dim polepis As String
polepis = "'Π'"
par = "'Α','Λ'"
            
Dim sql As String, sqlwh As String
            sql = "SELECT  'e' AS EIDOS,CONVERT(INTEGER,AFM) AS PELID,1 AS ID_NUM, ISNULL(AJ1,0) AS AJ1  ,ISNULL(AJ2,0) AS AJ2 , ISNULL(AJ3,0) AS AJ3,"
            sql = sql + "ISNULL(AJ4,0) as aj4,isnull(AJ5,0) as AJ5,isnull(AJI,0) as AJI,"
            sql = sql + "isnull(FPA1,0) as FPA1,ISNULL(FPA2,0) AS FPA2,ISNULL(FPA3,0) AS FPA3,ISNULL(FPA4,0) AS FPA4 ,0 AS FPA5,ATIM,"
            
            sql = sql + "HME,EPO,AFM,KPE,DIE,XRVMA"
            sql = sql + ",EPA,POL,ISNULL(AJ6,0) AS AJ6,ISNULL(FPA6,0) AS FPA6,ISNULL(AJ7,0) AS AJ7,ISNULL(FPA7,0) AS FPA7,"
            sql = sql + "ISNULL(AJ1,0)+ISNULL(AJ2,0)+ISNULL(AJ3,0)+ISNULL(AJ4,0)+ISNULL(AJ5,0)+ISNULL(AJ6,0)+ISNULL(AJ7,0) AS SKAU,"
            sql = sql + "ISNULL(FPA1,0)+ISNULL(FPA2,0)+ISNULL(FPA3,0)+ISNULL(FPA4,0)+ISNULL(FPA6,0)+ISNULL(FPA7,0) AS SFPA ,"
            sql = sql + "( CASE WHEN (left(ATIM,1) IN(" + polepis + "  )) THEN -1 ELSE 1 END ) AS MSIGN "
            sql = sql + "   FROM TIMTEMP2  "
            '             sqlwh = sqlwh + "  AND AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7>0  " + synt
            
            
            
       Dim SQLPEL As String
       
            
SQLPEL = "SELECT DISTINCT ISNULL(AFM,'000000000') AS CAFM, CONVERT(INTEGER,AFM) AS ID,EPO,ISNULL(EPA,'') AS CEPA,'' AS CDOY,ISNULL(DIE,'') AS CDIE,ISNULL(XRVMA,'') AS TK," _
                 & "ISNULL(POL,'') AS CPOL,'' AS CTHL  FROM TIMTEMP2 "

sqlwh = sql + sqlwh
  makeXML sql, sqlwh, SQLPEL

End Sub

Private Sub cmdGeferror_Click()
Dim dum, mfil As String
mfil = "C:\MERCVB\LOGGEFYR.TXT"
dum = Shell("notepad.exe " + mfil, vbMaximizedFocus)
End Sub

Private Sub cmdPBSGEFYRA_Click()
 
 
 Dim L As Integer
 For L = 1 To 100
    arr_log(L) = ""
    arr_posa(L) = 0
 Next
 List2.Clear
 selectRecords 0, 2
 
 For L = 1 To 100
    If Len(arr_log(L)) = 0 Then Exit For
    List2.AddItem arr_log(L) + " => " + Format(arr_posa(L), "#####0.00")
    If Left$(arr_log(L), 1) = "7" Then
       pol = pol + arr_posa(L)
    End If
    
   If Left$(arr_log(L), 1) = "2" Then
       ago = ago + arr_posa(L)
    End If
    If Left$(arr_log(L), 1) = "6" Then
       EXO = EXO + arr_posa(L)
    End If
    
    
 Next
 
 
    List2.AddItem " ====== "
    List2.AddItem "7 => " + Format(pol, "#####0.00")
       List2.AddItem arr_log(L) + "2 => " + Format(ago, "#####0.00")
          List2.AddItem arr_log(L) + "6 => " + Format(EXO, "#####0.00")
          
 
End Sub

Private Sub cmdG_Click()





'Στις Τετ, 23 Ιουν 2021 στις 12:43 μ.μ., ο/η Τεχνολογιστική Ι.Κ.Ε. <info@tlgi.gr> έγραψε:
'
'        Π     70-00-01-0036 ΠΩΛΗΣ.ΕΜΠΟΡ.ΕΣΩΤΕΡ.ΧΟΝΔΡΙΚΑ ΥΠΟΚ/ΜΑ ΦΠΑ 13%
'        Π     54-00-70-0036 ΦΠΑ ΕΚΡΟΩΝ ΕΜΠΟΡΕΥΜΑΤΩΝ ΜΕ 13%
'        Π     70-00-01-0087 ΠΩΛΗΣ.ΕΜΠΟΡ.ΕΣΩΤΕΡ.ΧΟΝΔΡΙΚΑ ΥΠΟΚ/ΜΑ ΦΠΑ 24%
'        Π     54-00-70-0087 ΦΠΑ ΕΚΡΟΩΝ ΕΜΠΟΡΕΥΜ.ΕΣΩΤΕΡΙΚΟΥ 24%
'        Π     54-09-18-0087 ΠΕΡΙΒΑΛΛΟΝΤΙΚΟ ΤΕΛΟΣ ΣΤΙΣ ΠΛΑΣΤΙΚΕΣ ΣΑΚΟΥΛΕΣ 24%
'    Χ  30-00-00-1000  ΠΕΛΑΤΕΣ ΕΣΩΤΕΡΙΚΟΥ ΥΠΟΚ/ΤΟΣ
'
'


















Dim R As New ADODB.Recordset

         Dim synt As String
         synt = " and HME>='" + Format(APO, "mm/dd/yyyy") + "'  AND HME<='" + Format(EOS, "mm/dd/yyyy") + "'  "
            If chkEPAN.Value = vbChecked Then
                

            Else
                synt = synt + " and (B_C1 is null or LEFT(B_C1,1)<>'*') "

            End If
            ' ExecuteSQLQuery("update TIM SET AJ7=0 WHERE AJ7 IS NULL")

            
            If chkPOL.Value = vbChecked And chkAG.Value = vbChecked Then
                'ok

            Else
               If chkPOL.Value = vbChecked Then
                  synt = synt + " and TIM.EIDOS='e' "
               Else
                 If chkAG.Value = vbChecked Then
                   synt = synt + " and TIM.EIDOS='r' "
                 Else
                   'ok
                 End If
               End If

            End If












        Dim pol As String: pol = " "
        Dim polepis As String: polepis = " "
        Dim ago As String: ago = " "
        Dim AGOEPIS As String: AGOEPIS = " "
        Get_AJ_ASCII pol, polepis, ago, AGOEPIS
        
        par = ""
        If chkPOL.Value = vbChecked Then
            par = pol + "," + polepis
        End If
            
        
        If chkAG.Value = vbChecked Then
           If Len(par) > 1 Then
             par = par + ","
           End If
           
              par = par + ago + "," + AGOEPIS
        End If
        
        
        
        
        
        

            Gdb.Execute "UPDATE TIM SET AJ7=0  WHERE AJ7 IS NULL"
            Gdb.Execute "UPDATE TIM SET FPA7=0 WHERE FPA7 IS NULL"


            '  Dim XL As DataTable
Dim sql As String, sqlwh As String
            sql = "SELECT TRP,TIM.EIDOS AS EIDOS,PEL.ID AS PELID,ID_NUM, AJ1  ,AJ2 , AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,0 AS FPA5,ATIM,"
            sql = sql + "HME,PEL.EPO,PEL.AFM,KPE,PEL.DIE,PEL.XRVMA"    '"CONVERT(CHAR(10),HME,3) AS HMEP
            sql = sql + ",PEL.EPA,PEL.POL,AJ6,FPA6,AJ7,FPA7,AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7 AS SKAU,FPA1+FPA2+FPA3+FPA4+FPA6+FPA7 AS SFPA ,"
            sql = sql + "( CASE WHEN ascii(left(ATIM,1)) IN(" + polepis + "," + AGOEPIS + "  ) THEN -1 ELSE 1 END ) AS MSIGN ,ISNULL(KODGAL,'1') AS KODGAL  "
            
            sqlwh = "   FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
            sqlwh = sqlwh + " WHERE AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7>0  " + synt
            'SQLWH = SQLWH + " order by HME"
            
            
            
            '  sqlwh = "   FROM EGG INNER JOIN PEL ON EGG.EIDOS=PEL.EIDOS AND EGG.KOD=PEL.KOD "
          
            '  sqlwh = sqlwh + " WHERE  (IDTIM IS NULL)    and HME>='" + Format(apo, "mm/dd/yyyy") + "'  AND HME<='" + Format(eos, "mm/dd/yyyy") + "'  "
            '  sqlwh = sqlwh + "  AND EGG.XREOSI+EGG.PISTOSI>0  "  ' + synt
                       
            
            
            
            
            
            
            
            
            
            
            
            SQLPEL = "SELECT ID,EPO,ISNULL(AFM,'000000000') AS CAFM,ISNULL(EPA,'') AS CEPA,ISNULL(DOY,'') AS CDOY,ISNULL(DIE,'') AS CDIE,ISNULL(XRVMA,'') AS TK," _
                 & "ISNULL(POL,'') AS CPOL,ISNULL(THL,'') AS CTHL " _
                 & " FROM PEL WHERE ID IN ( SELECT DISTINCT PEL.ID " + sqlwh + ")"
            
            
            
    
'   Exit Sub
  ' Gdb.BeginTrans
    GmakeXML sql, sqlwh, SQLPEL, f_categ, 0
    'cmdG.Enabled = False
  '  Gdb.CommitTrans
End Sub

Private Sub cmdImport_Click()
    ' EXCEL_IMPORT
    
        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim R   As New ADODB.Recordset

        Dim r0  As New ADODB.Recordset

        Dim mok As Integer

        Dim PIN(30, 2)

        Dim COUNTER    As Integer

        Dim KOD_OK     As Boolean    ' ΕΧΕΙ ΔΗΛΩΣΕΙ ΤΟΝ ΚΩΔΙΚΟ

        Dim BARCODE_OK As Boolean    ' ΕΧΕΙ ΔΗΛΩΣΕΙ TA BARCODES

        Dim sql        As String

100     KOD_OK = False
110     BARCODE_OK = False

        Dim k As Long

120     COUNTER = 1  ' ARIUMOS PEDION POY THA METAFERTHOYN



        
        'ΕΛΕΓΧΩ ΑΝ ΥΠΑΡΧΕΙ Ο ΠΙΝΑΚΑΣ ΠΕΛΑΤΕΣ ΑΝ ΟΧΙ ΤΟΤΕ ΣΗΜΑΙΝΕΙ ΟΤΙ ΕΙΜΑΙ ΣΕ ΛΑΘΟΣ ΒΑΣΗ (MASTER)
175     R.Open "SELECT COUNT(*) AS N  FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME='TIMTEMP';", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next
      
        If R(0) = 0 Then

        

'           Dim sqlwh As String
'            sql = "SELECT TOP 0 PEL.ID AS PELID,ID_NUM, AJ1  ,AJ2 , AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,0 AS FPA5,ATIM,"
'            sql = sql + "HME,PEL.EPO,PEL.AFM,KPE,PEL.DIE,PEL.XRVMA"    '"CONVERT(CHAR(10),HME,3) AS HMEP
'            sql = sql + ",PEL.EPA,PEL.POL,AJ6,FPA6,AJ7,FPA7,AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7 AS SKAU,FPA1+FPA2+FPA3+FPA4+FPA6+FPA7 AS SFPA ,"
'            sql = sql + "1 AS MSIGN "
'            sqlwh = " INTO TIMTEMP  FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
'
           
'    CREATE TABLE [dbo].[TIMTEMP2](
'    [PELID] [int]  NULL,
'    [ID_NUM] [int] IDENTITY(1,1) NOT NULL,
'    [AJ1] [float] NULL,
'    [AJ2] [float] NULL,
'    [AJ3] [float] NULL,
'    [AJ4] [float] NULL,
'    [AJ5] [float] NULL,
'    [AJI] [float] NULL,
'    [FPA1] [float] NULL,
'    [FPA2] [float] NULL,
'    [FPA3] [float] NULL,
'    [FPA4] [float] NULL,
'    [FPA5] [int]  NULL,
'    [ATIM] [varchar](255) NULL,
'    [HME] [datetime] NULL,
'    [EPO] [varchar](255) NULL,
'    [AFM] [varchar](255) NULL,
'    [KPE] [varchar](255) NULL,
'    [DIE] [nvarchar](60) NULL,
'    [XRVMA] [varchar](255) NULL,
'    [EPA] [varchar](255) NULL,
'    [POL] [varchar](255) NULL,
'    [AJ6] [float] NULL,
'    [FPA6] [float] NULL,
'    [AJ7] [float] NULL,
'    [FPA7] [float] NULL,
'    [SKAU] [float] NULL,
'    [SFPA] [float] NULL,
'    [MSIGN] [int]  NULL
'') ON [PRIMARY]
           
           
           
           
           
           
           
           
           
           
           
           
           
           
           
           
         ' Gdb.Execute sql + sqlwh
        End If
        
        R.Close
        

' Exit Sub
Gdb.Execute "DELETE FROM TIMTEMP2"
        
130     r0.Open "SELECT TOP 0 * FROM TIMTEMP2 ", Gdb, adOpenDynamic, adLockOptimistic

        ' FORTONO TON PINAKA 20 ME TIS ANTISTOIXISEIS TON PEDION
        ' ΚΑΙ ΕΛΕΓΧΩ ΑΝ ΕΧΕΙ ΔΗΛΩΘΕΙ Ο ΚΩΔΙΚΟΣ
140     R.Open "SELECT *FROM PINAKES WHERE TYPOS=101 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

150     Do While Not R.EOF

160         If R("typos") = 101 Then
170             PIN(COUNTER, 1) = R("PERIGRAFH")    'ΠΕΔΙΟ ΤΟΥ ΚΥΡΙΟΥ ΑΡΧΕΙΟΥ (SQLSERVER)
180             PIN(COUNTER, 2) = R("TIMH")    ' ΑΡΙΘΜΟΣ ΣΤΗΛΗΣ ΤΟΥ EXCEL


                ' On Error GoTo LATOSPEDIO
220             mok = 0    'ΕΛΕΓΧΩ ΑΝ ΤΑ ΠΕΔΙΑ ΠΟΥ ΔΗΛΩΣΑ ΣΤΟ <<PINAKES>> ΥΠΑΡΧΟΥΝ ΣΤΟΝ ΠΙΝΑΚΑ PEL
               ' r0.AddNew
230             For k = 0 To r0.FIELDS.Count - 1

240                 If r0(k).Name = R("PERIGRAFH") Then
250                     mok = 1

                        Exit For

                    End If

                Next

260             If mok = 0 Then
270                 MsgBox "Λάθος στην δήλωση των πεδίων." + Chr(13) + "Δεν υπάρχει στο αρχείο ειδών το πεδίο " + R("PERIGRAFH")

                    Exit Sub

                End If

                On Error GoTo Command1_Click_Err

280             COUNTER = COUNTER + 1
            Else

                Exit Do

            End If
           ' r0.Update

290         R.MoveNext
        Loop

        ' kathgories

300     R.Close


330     Set xlwbook = xl.Workbooks.Open(txtText1.Text)
340     Set xlsheet = xlwbook.Sheets.Item(1)

        Dim ko   As String

        Dim mNew As Long, mUpd As Long

350     mNew = 0
360     mUpd = 0

370     'Label4.Caption = "Nέες εγγραφές 0"
380    ' Label5.Caption = "Ενημέρωση εγγραφές 0"

        Dim mRow As Long

        ' data1.Recordset.MoveFirst
390     mRow = 2 'Val(Text2.Text)    '  data1.Recordset.Move

        On Error GoTo error_name

400     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

410         If IsNull(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

420         If IsEmpty(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

430         ko = "" 'Trim(xlsheet.cells(mRow, KOD_COLUMN))    '  ' ΚΩΔΙΚΟς ΤΟΥ ΕΙΔΟΥΣ  EXCEL

440         R.Open "SELECT * FROM TIMTEMP2 ", Gdb, adOpenDynamic, adLockOptimistic


460             R.AddNew

   
480             mNew = mNew + 1
490             Me.Caption = "Nέες εγγραφές " + Format(mNew, "######")

            'ΕΝΗΜΕΡΩΝΩ ΤΑ ΥΠΟΛΟΙΠΑ ΠΕΔΙΑ ΑΠΟ ΤΟ EXCEL
520         For k = 1 To COUNTER - 1


540                 If R(PIN(k, 1)).Type = 202 Then    ' KEIMENO
550                     R(PIN(k, 1)) = Left(xlsheet.cells(mRow, PIN(k, 2)), R(PIN(k, 1)).DefinedSize)
                    Else    'ARITMOS
560                     R(PIN(k, 1)) = xlsheet.cells(mRow, PIN(k, 2))
                    End If
 

            Next

570         R("msign") = 1 ' Left(Combo2.Text, 1)
             ' R("ID_NUM") = 1
             R("PELID") = 1
              R("FPA5") = 0

580         R.Update

590         R.Close

600         DoEvents
610         Me.Caption = ko
620         mRow = mRow + 1    'data1.Recordset.MoveNext
        Loop

630     xl.Quit
640     Set xlwbook = Nothing
650     Set xl = Nothing

660     MsgBox "τέλος ενημέρωσης"

        Exit Sub

error_name:
670     mok = MsgBox("λάθος στην σειρά " + Format(mRow, "#######"), vbYesNo, "Διακοπή ;")

680     If mok = vbYes Then Exit Sub

690     Resume Next

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu51.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>


    
End Sub

'Sub EXCEL_IMPORT()
'
'
'
'
'
'        ' ME EXCEL
'        '<EhHeader>
'        On Error GoTo SOKRATIS_import_Err
'
'        '</EhHeader>
'        Dim xl As New excel.Application
'
'        Dim xlsheet As excel.Worksheet
'
'        Dim xlsheet3 As excel.Worksheet
'
'        Dim xlwbook As excel.workbook
'
'100
'110     Set xlwbook = xl.Workbooks.Open(mF) ' "c:\MERCVB\TIMOL.XLSX")
'        Set xlsheet = xlwbook.Sheets.Item(1)
'
'        Dim k
'
'        Dim a$, B, X
'
'        Dim d
'
'        Dim FPA As Integer
'
'        Dim rSQL As New ADODB.Recordset
'
'        Dim kod, per, MON, POS, TIMH, TIMOL
'
'        Dim F As String, RECS As Integer
'
'120     k = 1
'
'        Dim CH1, CH2, CH3
'
'        Dim XTI, lti
'
'        Dim s As String
'
'        Dim nn As Long
'
'        Dim LL2 As Long
'
'        Dim ono As String
'
'        'Dim FPA As Integer
'        Dim TIMM As Single
'        Dim EKPT As Single
'
'130     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF
'
'140         If IsNull(xlsheet.cells(k, 1)) Then
'
'                Exit Do
'
'            End If
'
'150         If IsEmpty(xlsheet.cells(k, 1)) Then
'
'                Exit Do
'
'            End If
'
'
'
'
'            On Error Resume Next
'
'160         kod = xlsheet.cells(k + 1, 1)  ' LTrim(str(xlsheet.cells(k + 1, 1)))    '
'170         POS = xlsheet.cells(k + 1, 3)    ' PROSFORA
'            ono = xlsheet.cells(k + 1, 2)  ' PERIGRAFH
'
'           If POS <> 0 Then
'
'
'
'
'220         rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'
'230         If rSQL.EOF Then
'                 Gdb.Execute "insert into EID (KOD,ONO,FPA,MON) VALUES ('" + kod + "','" + ono + "',1,'TEM')"
'              '  MsgBox "DENYPARXEI O KVDIKOS " + kod
'                xlsheet.cells(k, 7) = "δημιουργηθηκε ο κωδικός ME KAT ΦΠΑ=1 & ΜΟΝ=ΤΕΜ"
'            End If
'
'                ono = rSQL("ONO")
'                FPA = rSQL("FPA")
'                MON = rSQL("MON")
'                TIMM = rSQL("xti")   ' xlsheet.cells(k + 1, 3)
'                EKPT = 0 '  xlsheet.cells(k + 1, 4)
'
'           ' End If
'
'290         rSQL.Close
'
'300         If k > f_SeiresTimologioy Then Exit Do
'
'
'
'420         Grid1.row = k
'430        ' FindSynolo
'
'           End If
'
'
'440         k = k + 1
'
'        Loop
'
'450     Set xlsheet = Nothing
'460     Set xlwbook = Nothing
'470     xl.Quit
'
'        '<EhFooter>
'        Exit Sub
'
'SOKRATIS_import_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.par1.SOKRATIS_import " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SOKRATIS_import " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub




Private Sub CMDSTOP_Click()
    CMDSTOP.Enabled = False
    
End Sub

Private Sub cmdYPOL_Click()
Dim pol As Double, ago As Double, EXO As Double
 pol = 0: ago = 0: EXO = 0



 Dim L As Integer
 For L = 1 To 100
    arr_log(L) = ""
    arr_posa(L) = 0
 Next
 List2.Clear
 selectRecords 0, 1
 
 For L = 1 To 100
    If Len(arr_log(L)) = 0 Then Exit For
    List2.AddItem arr_log(L) + " => " + Format(arr_posa(L), "#####0.00")
   
   If Left$(arr_log(L), 1) = "7" Then
       pol = pol + arr_posa(L)
    End If
    
   If Left$(arr_log(L), 1) = "2" Then
       ago = ago + arr_posa(L)
    End If
    If Left$(arr_log(L), 1) = "6" Then
       EXO = EXO + arr_posa(L)
    End If
   
   
   
   
 Next
 
   List2.AddItem " ====== "
    List2.AddItem "7 => " + Format(pol, "#####0.00")
       List2.AddItem arr_log(L) + "2 => " + Format(ago, "#####0.00")
          List2.AddItem arr_log(L) + "6 => " + Format(EXO, "#####0.00")
          

End Sub

Private Sub selectRecords(mekathg As Integer, EPSILON1 As Integer) ' 2=PBS
On Error GoTo LATOS

List1.Clear

Dim R As New ADODB.Recordset

         'Dim synt As String
         
          Dim synt As String
         synt = " AND HME>='" + Format(APO, "mm/dd/yyyy") + "'  AND HME<='" + Format(EOS, "mm/dd/yyyy") + "'  "
         
         
            If chkEPAN.Value = vbChecked Then
                'synt = ""

            Else
                synt = " and (B_C1 is null or LEFT(B_C1,1)<>'*') "

            End If
            ' ExecuteSQLQuery("update TIM SET AJ7=0 WHERE AJ7 IS NULL")


        Dim pol As String: pol = " "
        Dim polepis As String: polepis = " "
        Dim ago As String: ago = " "
        Dim AGOEPIS As String: AGOEPIS = " "
        Get_AJ_ASCII pol, polepis, ago, AGOEPIS



        par = ""
        If chkPOL.Value = vbChecked Then
            par = pol + "," + polepis
        End If
            
        
        If chkAG.Value = vbChecked Then
           If Len(par) > 1 Then
             par = par + ","
           End If
           
              par = par + ago + "," + AGOEPIS
        End If
        
        
        




     '   par = pol + "," + POLEPIS + "," + ago + "," + AGOEPIS
        








On Error Resume Next

            Gdb.Execute "IF EXISTS (SELECT * FROM dbo.SQLTEMP) DROP TABLE dbo.SQLTEMP"

            Gdb.Execute "UPDATE TIM SET AJ7=0 WHERE AJ7 IS NULL"
            Gdb.Execute "UPDATE TIM SET FPA7=0 WHERE FPA7 IS NULL"

On Error GoTo LATOS

            '  Dim XL As DataTable
Dim sql As String, sqlwh As String
            sql = "SELECT TIM.EIDOS AS EIDOS,PEL.ID AS PELID,ID_NUM, AJ1  ,AJ2 , AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,0 AS FPA5,ATIM,APALAGIFPA,ENTITYMARK,ENTITYUID,"
            sql = sql + "HME,PEL.EPO,PEL.AFM,KPE,PEL.DIE,PEL.XRVMA"    '"CONVERT(CHAR(10),HME,3) AS HMEP
            sql = sql + ",PEL.EPA,PEL.POL,PEL.DOY,AJ6,FPA6,AJ7,FPA7,AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7 AS SKAU,FPA1+FPA2+FPA3+FPA4+FPA6+FPA7 AS SFPA ,"
            sql = sql + "( CASE WHEN ascii(left(ATIM,1)) IN(" + polepis + "," + AGOEPIS + "  ) THEN -1 ELSE 1 END ) AS MSIGN "
            sqlwh = " FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD  where AKYROMENO=0 AND ascii(left(ATIM,1)) IN(" + par + ") "
            sqlwh = sqlwh + "  AND AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7>0  " + synt
            'SQLWH = SQLWH + " order by HME"
            
            
            
            
            Gdb.Execute sql + " into SQLTEMP " + sqlwh
            
           ' sqlwh =
            
            SQLPEL = "SELECT ID,EPO,ISNULL(AFM,'000000000') AS CAFM,ISNULL(EPA,'') AS CEPA,ISNULL(DOY,'') AS CDOY,ISNULL(DIE,'') AS CDIE,ISNULL(XRVMA,'') AS TK," _
                 & "ISNULL(POL,'') AS CPOL,ISNULL(THL,'') AS CTHL " _
                 & " FROM PEL WHERE ID IN ( SELECT DISTINCT PELID FROM SQLTEMP ) "    '   + sqlwh + ")"
            
            On Error GoTo 0
            
            If EPSILON1 = 2 Then
               makePBSXML sql, "SELECT * FROM SQLTEMP", SQLPEL
            Else
              If mekathg = 0 Then
                makeXML sql, "SELECT * FROM SQLTEMP", SQLPEL
              Else
                makeXML_KATHG sql, "SELECT * FROM SQLTEMP", SQLPEL
              End If
            End If
            
            
            
            Exit Sub
            
            
LATOS:
MsgBox "ΥΠΕΡΒΑΣΗ ΧΡΟΝΟΥ " + Err.Description

End Sub

Private Sub makePBSXML(ByVal sql As String, ByVal sqlwh As String, ByVal SQLPEL As String)
'=============================================================================================================

Dim R As New ADODB.Recordset
            
            R.Open sqlwh + " order by HME", Gdb, adOpenDynamic, adLockOptimistic
Dim mpelP As String
Dim mpela As String
mpelP = InputBox("Κωδ.Πελατη ", "", "30-00-00-0000")

mpela = InputBox("Κωδ.Προμηθευτη ", "", "50-00-00-0000")
     Dim varStock As Variant
       'global       Dim docStock As MSXML2.DOMDocument
        Dim elemDATA As MSXML2.IXMLDOMElement
        Dim elemARTICLES As MSXML2.IXMLDOMElement
        Dim elemARTICLE As MSXML2.IXMLDOMElement
         
         
        Dim elemDETAILS As MSXML2.IXMLDOMElement
        Dim elemDETAIL As MSXML2.IXMLDOMElement
         
          Dim elemCUSTOMERS As MSXML2.IXMLDOMElement
          Dim elemCUSTOMER As MSXML2.IXMLDOMElement
          
        Dim elemField As MSXML2.IXMLDOMElement
        Dim I As Integer
    

102     Set docStock = New MSXML2.DOMDocument
104     With docStock
106         .appendChild .createProcessingInstruction("xml", _
                                                      "version=""1.0"" encoding=""ISO-8859-7""")
108         Set elemDATA = .createElement("GDataXML")
110         With elemDATA
112             '.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
114             '.setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"


                  Set elemDETAILS = docStock.createElement("FileInfo")
                       
                          Set elemField = docStock.createElement("MTYPE"): elemField.Text = "SOFT1": elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("Version"):  elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("CreationDate"): elemField.Text = Format(Now, "dd/MM/yyyy"): elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("DataFromDate"): elemField.Text = Format(APO, "dd/MM/yyyy"): elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("DataToDate"): elemField.Text = Format(EOS, "dd/MM/yyyy"): elemDETAILS.appendChild elemField
                       Set elemField = docStock.createElement("TIN"): elemField.Text = "142205729": elemDETAILS.appendChild elemField
                       Set elemField = docStock.createElement("SerialNumber"): elemField.Text = "10502558420321": elemDETAILS.appendChild elemField
                       Set elemField = docStock.createElement("SXRecCount"): elemField.Text = "71": elemDETAILS.appendChild elemField
                 .appendChild elemDETAILS
                 
                 
                            
                            
                            
                 Set elemARTICLES = docStock.createElement("SXRecords")
                 
                 
                 
                 Dim mType As String
                 Dim pel30 As String
                 Dim kID As Integer: kID = 0
                 Do While Not R.EOF 'For k = 1 To 5
                    Me.Caption = Format(R!hme, "dd/MM/yyyy")
                    
                    kID = kID + 1
                    
                  Set elemARTICLE = docStock.createElement("DocHeader"): elemARTICLE.Text = CStr(k):
                    With elemARTICLE
                    
'                       If Trim(R!EIDOS) = "e" Then
'                          mType = "0"
'                          pel30 = mpelP  '"30-0000"
'                       Else
'                          mType = "1"
'                          pel30 = mpela ' "50-0000"
'                       End If
                       
                     If Trim(R!EIDOS) = "e" Then
                          mType = "0"
                          pel30 = mpelP  '"30-0000"
                       Else
                          mType = "1"
                          pel30 = mpela ' "50-0000"
                       End If
                       
                       
                                              
                        Set elemField = docStock.createElement("ID"): elemField.Text = str(kID): .appendChild elemField
                        Set elemField = docStock.createElement("SxSign"): elemField.Text = "1": .appendChild elemField
                        Set elemField = docStock.createElement("DocDate"): elemField.Text = Format(R!hme, "dd/MM/yyyy"): .appendChild elemField



                       Set elemField = docStock.createElement("DocSeriesId"): elemField.Text = Format(R!hme, "dd/MM/yyyy"): .appendChild elemField
                       Set elemField = docStock.createElement("Branch"): elemField.Text = mType: .appendChild elemField
                       Set elemField = docStock.createElement("TradeCode"): elemField.Text = R!ATIM: .appendChild elemField
                      
                       
                       
                       
                       Set elemField = docStock.createElement("Comment"): elemField.Text = "": .appendChild elemField
                       Set elemField = docStock.createElement("Comment2"): elemField.Text = "": .appendChild elemField
                       
                       
                       Set elemField = docStock.createElement("IsCancelled"): elemField.Text = "0": .appendChild elemField
                       Set elemField = docStock.createElement("IsCancelling"): elemField.Text = "0": .appendChild elemField
                      
                      
                       Set elemField = docStock.createElement("CurrencyISO"): elemField.Text = "EUR": .appendChild elemField
                       
                       
                       '========================================
                       
                       Set elemField = docStock.createElement("OriginId"): elemField.Text = "251": .appendChild elemField
                       Set elemField = docStock.createElement("TaxSeries"): elemField.Text = "ΤΠΥ": .appendChild elemField
                       Set elemField = docStock.createElement("TaxSeriesNum"): elemField.Text = "1": .appendChild elemField
                       
                       Set elemField = docStock.createElement("CurrencyRate"): elemField.Text = "1": .appendChild elemField
                       
                       
                       Set elemField = docStock.createElement("MyfValueKind"): elemField.Text = "1": .appendChild elemField
                       
                       Set elemField = docStock.createElement("MyfSign"): elemField.Text = str(R!MSIGN): .appendChild elemField
                        
                        
                       Set elemField = docStock.createElement("MyfRetailSN"): elemField.Text = "": .appendChild elemField
                        
                        
                        
                        
                        
                        
                        '-------------------------------------------------------------------------------
                       Set elemDETAILS = docStock.createElement("myDataInfo")
                       
'        <myDataSendCode>2.1</myDataSendCode>
'        <myDataGetType>0</myDataGetType>
'        <SelfPricing>0</SelfPricing>

'        <PaymentMethod>5</PaymentMethod>
'        <VatProvisions>0</VatProvisions>
'        <UID>7852BC5F9FA3CF72066137D56D224820A73C5FD4</UID>
'        <MarkId>400002030871613</MarkId>
                       
                       
                      Dim myDataSendCode As String: myDataSendCode = GGET_CVALUE("SELECT MYEID FROM PARASTAT WHERE EIDOS='" + Left$(R!ATIM, 1) + "'")
                          Set elemField = docStock.createElement("myDataSendCode"): elemField.Text = myDataSendCode: elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("myDataGetType"): elemField.Text = "0": elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("SelfPricing"): elemField.Text = "0": elemDETAILS.appendChild elemField
                          
                          Dim pl As String: pl = GGET_CVALUE("select  STR(N1)  from PINAKES WHERE TYPOS=12 AND AYJON=(SELECT CONVERT(INT,LEFT(TRP,1) )  FROM TIM where ATIM='" + R!ATIM + "')")
                          Set elemField = docStock.createElement("PaymentMethod"): elemField.Text = pl: elemDETAILS.appendChild elemField
                          
                          Set elemField = docStock.createElement("VatProvisions"): elemField.Text = nNull(R!APALAGIFPA): elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("UID"): elemField.Text = CNull(R!ENTITYUID): elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("MarkID"): elemField.Text = CNull(R!entitymark): elemDETAILS.appendChild elemField
                        .appendChild elemDETAILS
                        
                        
                        
                         
                       '------------------------------------------------------------------------
                       Set elemDETAILS = docStock.createElement("Trader")
                       
                '        <TraderType>13</TraderType>
                '        <TraderAfm>095618983</TraderAfm>
                '        <TraderMyfStatus>0</TraderMyfStatus>
                '        <TraderCode>30.01.00.0001</TraderCode>
                '        <TraderName>ΙΒS ΣΥΣΤΗΜΑΤΑ ΠΛΗΡΟΦΟΡΙΚΗΣ ΕΠΕ</TraderName>
                '        <TraderIrsdata>1201 Α ΠΕΙΡΑΙΑ</TraderIrsdata>
                
                '        <TraderVatsts>1</TraderVatsts>
                '        <TraderCurrency>EUR</TraderCurrency>
                '        <TraderCountry>GR</TraderCountry>
                
                '        <TraderAddress>ΝΟΤΑΡΑ 108</TraderAddress>
                '        <TraderCity/>
                '        <TraderZip>18535</TraderZip>
                '        <TraderPhone01>210 4224242</TraderPhone01>
                '        <TraderCategory>2</TraderCategory>
                          Set elemField = docStock.createElement("TraderType"): elemField.Text = IIf(R!EIDOS = "e", "13", "12"): elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("TraderAfm"): elemField.Text = CNull(R!AFM): elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("TraderMyfStatus"): elemField.Text = "0": elemDETAILS.appendChild elemField
                          
                          Set elemField = docStock.createElement("TraderCode"): elemField.Text = pel30: elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("TraderName"): elemField.Text = CNull(R!EPO): elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("TraderIrsdata"): elemField.Text = "0": elemDETAILS.appendChild elemField
                           
                           
                          Set elemField = docStock.createElement("TraderVatsts"): elemField.Text = "1": elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("TraderCurrency"): elemField.Text = "EUR": elemDETAILS.appendChild elemField
                          
                          Set elemField = docStock.createElement("TraderCountry"): elemField.Text = "GR": elemDETAILS.appendChild elemField
                          
                          
                          
                          Set elemField = docStock.createElement("TraderAddress"): elemField.Text = CNull(R!AFM): elemDETAILS.appendChild elemField
                          
                          Set elemField = docStock.createElement("TraderCity"): elemField.Text = "": elemDETAILS.appendChild elemField
                          
                          
                          Set elemField = docStock.createElement("TraderZip"): elemField.Text = "": elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("TraderPhone01"): elemField.Text = "": elemDETAILS.appendChild elemField
                          Set elemField = docStock.createElement("TraderCategory"): elemField.Text = "1001": elemDETAILS.appendChild elemField 'Λογιστική κατηγορία, Soft1 -->
                       
                       
                        .appendChild elemDETAILS
                       
                       
                      
                       
                       Set elemDETAILS = docStock.createElement("DocLines")
                        
                         ' <LineNum>1</LineNum>
                         '<AccountCode>70.07.0024</AccountCode>
                         ' <LSubValue>375.00</LSubValue>
                         '<LVatValue>90.00</LVatValue>

                          
                         Dim LLK As Integer
                         For LLK = 1 To 7
                           If R("aj" + Format(LLK, "0")) > 0 Then
                             Set elemDETAIL = docStock.createElement("DocLine")
                                Dim CC9 As String: CC9 = FIND_KOD(R!ATIM, LLK)  'FIND_KOD(R!ATIM, LLK)
                                Set elemField = docStock.createElement("LineNum"): elemField.Text = Format(LLK, "0"): elemDETAIL.appendChild elemField
                                Set elemField = docStock.createElement("AccountCode"): elemField.Text = CC9: elemDETAIL.appendChild elemField
                                Set elemField = docStock.createElement("LSubValue"): elemField.Text = Replace(Format(R("aj" + Format(LLK, "0")), "#####0.00"), ".", ","): elemDETAIL.appendChild elemField
                                Set elemField = docStock.createElement("LVatValue"): elemField.Text = Replace(Format(R("FPA" + Format(LLK, "0")), "#####0.00"), ".", ","): elemDETAIL.appendChild elemField ' 0 AS FPA5 STO SQL OPOTE OK
                               
                                
                             elemDETAILS.appendChild elemDETAIL
                             Add_to_Array CC9, R("aj" + Format(LLK, "0")), R!MSIGN
                            End If
                         Next
                          
                          
                          
                          
                          
                          
                         .appendChild elemDETAILS
                         
                    
                    End With
                    elemARTICLES.appendChild elemARTICLE
                    R.MoveNext
                 Loop
                    
                    
                    
                    
                .appendChild elemARTICLES
                  
              
                    
               R.Close
               
                
'                 Set elemCUSTOMERS = docStock.createElement("CUSTOMERS")
'
'                  'Exit Sub
'
'                  R.Open SQLPEL, Gdb, adOpenDynamic, adLockOptimistic
'                   Do While Not R.EOF
'                    Set elemCUSTOMER = docStock.createElement("CUSTOMER")
'                     With elemCUSTOMER
'                        Set elemField = docStock.createElement("ID"): elemField.Text = str(R!ID): .appendChild elemField
'                        Set elemField = docStock.createElement("NAME"): elemField.Text = (R!EPO): .appendChild elemField
'                        Set elemField = docStock.createElement("VAT"): elemField.Text = Trim((R!CAFM)): .appendChild elemField
'
'
'                        Set elemField = docStock.createElement("JOB"): elemField.Text = (R!CEPA): .appendChild elemField
'                        Set elemField = docStock.createElement("DOY"): elemField.Text = (R!Cdoy): .appendChild elemField
'                        Set elemField = docStock.createElement("CUSTVAT"): elemField.Text = "2": .appendChild elemField
'
'
'                        Set elemField = docStock.createElement("ADDRESS"): elemField.Text = R!cDIE: .appendChild elemField
'                        Set elemField = docStock.createElement("ZIP"): elemField.Text = R!tk: .appendChild elemField
'                        Set elemField = docStock.createElement("CITY"): elemField.Text = R!cpOL: .appendChild elemField
'
'                        Set elemField = docStock.createElement("PHONE1"): elemField.Text = R!cTHL: .appendChild elemField
'                        Set elemField = docStock.createElement("PHONE2"): elemField.Text = " ": .appendChild elemField
'                        Set elemField = docStock.createElement("PHONE3"): elemField.Text = " ": .appendChild elemField
'
'
'                        Set elemField = docStock.createElement("FAX1"): elemField.Text = " ": .appendChild elemField
'                        Set elemField = docStock.createElement("FAX2"): elemField.Text = " ": .appendChild elemField
'                        Set elemField = docStock.createElement("EMAIL"): elemField.Text = " ": .appendChild elemField
'
'                        Set elemField = docStock.createElement("ISKEPYO"): elemField.Text = "0": .appendChild elemField
'                        Set elemField = docStock.createElement("ISDHMOSIOU"): elemField.Text = "0": .appendChild elemField
'                        Set elemField = docStock.createElement("ISEA"): elemField.Text = "0": .appendChild elemField
'
'                        Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": .appendChild elemField
'                        Set elemField = docStock.createElement("ACCADDR"): elemField.Text = " ": .appendChild elemField
'                        Set elemField = docStock.createElement("STRADDR"): elemField.Text = " ": .appendChild elemField
'
'                        Set elemField = docStock.createElement("STRNAME"): elemField.Text = " ": .appendChild elemField
'                        Set elemField = docStock.createElement("BANK1"): elemField.Text = " ": .appendChild elemField
'
'                        Set elemField = docStock.createElement("BANKACC1"): elemField.Text = " ": .appendChild elemField
'                        Set elemField = docStock.createElement("BANK2"): elemField.Text = " ": .appendChild elemField
'                        Set elemField = docStock.createElement("BANKACC2"): elemField.Text = " ": .appendChild elemField
'
'                        Set elemField = docStock.createElement("EACOUNTRY"): elemField.Text = " ": .appendChild elemField
'                        Set elemField = docStock.createElement("EAPREFIX"): elemField.Text = " ": .appendChild elemField
'                        Set elemField = docStock.createElement("EAVAT"): elemField.Text = " ": .appendChild elemField
'
'
'                        Set elemField = docStock.createElement("IDTYPE"): elemField.Text = " ": .appendChild elemField
'
'
'
'                     End With
'
'                    elemCUSTOMERS.appendChild elemCUSTOMER
'                    R.MoveNext
'
'                   Loop
'
'                .appendChild elemCUSTOMERS
                 
                     
'
'
'1116             For i = 0 To UBound(varStock)
'1118                 Set elemARTICLES = docStock.createElement("CUSTOMERS")
'1120                 With elemARTICLES
'1122                     Set elemField = docStock.createElement("ProductCode")
'1124                     elemField.Text = CStr(varStock(i)(0))
'1126                     .appendChild elemField
'1128                     Set elemField = docStock.createElement("ProductPrice")
'1130                     elemField.Text = CStr(varStock(i)(1))
'1132                     .appendChild elemField
'                    End With
'1134                 .appendChild elemARTICLES
'                Next
            End With
1136         Set .documentElement = elemDATA
            On Error Resume Next
'138         Kill "C:\created.xml"
'            On Error GoTo cmdCommand5_Click_Err
'140         .save "C:\created.xml"
'
'
'            Par7MyData.FormatXmlDocument docStock
'
'
            
            
            
138                     Kill "C:\created.xml"
            On Error GoTo cmdCommand5_Click_Err
140         '.save "C:\created.xml"


            Par7MyData.FormatXmlDocument docStock

           ' cd1.FILENAME = "GEF_" + Format(Now, "yyyy-MM-dd HH-mm") + ".xml"
           ' cd1.ShowSave
            CD1.FILENAME = "GEFB_" + Format(Now, "yyyy-MM-dd HH-mm") + ".xml"
            CD1.ShowSave
            
        ' .save cd1.FILENAME '"C:\created.xml"
         .save CD1.FILENAME '"C:\created.xml"

        End With
    
    MsgBox "ok"
    
    
    
        '<EhFooter>
        Exit Sub

cmdCommand5_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.Par7MyData.cmdCommand5_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub





            
Private Sub makeXML(ByVal sql As String, ByVal sqlwh As String, ByVal SQLPEL As String)
'=============================================================================================================

Dim R As New ADODB.Recordset

            
            
            
            
            
            R.Open sqlwh + " order by HME", Gdb, adOpenDynamic, adLockOptimistic
            

 





Dim mpelP As String
Dim mpela As String
mpelP = InputBox("Κωδ.Πελατη ", "", "30-00-00-0000")

mpela = InputBox("Κωδ.Προμηθευτη ", "", "50-00-00-0000")







     Dim varStock As Variant
       'global       Dim docStock As MSXML2.DOMDocument
        Dim elemDATA As MSXML2.IXMLDOMElement
        Dim elemARTICLES As MSXML2.IXMLDOMElement
        Dim elemARTICLE As MSXML2.IXMLDOMElement
         
         
        Dim elemDETAILS As MSXML2.IXMLDOMElement
        Dim elemDETAIL As MSXML2.IXMLDOMElement
         
          Dim elemCUSTOMERS As MSXML2.IXMLDOMElement
          Dim elemCUSTOMER As MSXML2.IXMLDOMElement
          
        Dim elemField As MSXML2.IXMLDOMElement
        Dim I As Integer
    

102     Set docStock = New MSXML2.DOMDocument
104     With docStock
106         .appendChild .createProcessingInstruction("xml", _
                                                      "version=""1.0"" encoding=""utf-8""")
108         Set elemDATA = .createElement("DATA")
110         With elemDATA
112             '.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
114             '.setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
                 
                            
                            
                            
                 Set elemARTICLES = docStock.createElement("ARTICLES")
                 
                 
                 
                 Dim mType As String
                 Dim pel30 As String
                 
                 Do While Not R.EOF 'For k = 1 To 5
                    Me.Caption = Format(R!hme, "dd/MM/yyyy")
                    Set elemARTICLE = docStock.createElement("ARTICLE"): elemARTICLE.Text = CStr(k):
                    With elemARTICLE
                    
'                       If Trim(R!EIDOS) = "e" Then
'                          mType = "0"
'                          pel30 = mpelP  '"30-0000"
'                       Else
'                          mType = "1"
'                          pel30 = mpela ' "50-0000"
'                       End If
                       
                     If Trim(R!EIDOS) = "e" Then
                          mType = "0"
                          pel30 = mpelP  '"30-0000"
                       Else
                          mType = "1"
                          pel30 = mpela ' "50-0000"
                       End If
                       
                       
                       
                       
                       Set elemField = docStock.createElement("MTYPE"): elemField.Text = mType: .appendChild elemField
                      
                       Set elemField = docStock.createElement("MSIGN"): elemField.Text = str(R!MSIGN): .appendChild elemField
                       Set elemField = docStock.createElement("ISKEPYO"): elemField.Text = "1": .appendChild elemField
                       Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": .appendChild elemField
                       
                       
                       Set elemField = docStock.createElement("LCODE"): elemField.Text = pel30: .appendChild elemField
                       Set elemField = docStock.createElement("CUSTID"): elemField.Text = str(R!PELID): .appendChild elemField
                       Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: .appendChild elemField
                       Set elemField = docStock.createElement("MDATE"): elemField.Text = Format(R!hme, "dd/MM/yyyy"): .appendChild elemField
                       Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: .appendChild elemField
                       
                       Set elemField = docStock.createElement("CASHAMT"): elemField.Text = "0,00": .appendChild elemField
                       Set elemField = docStock.createElement("LCCASH"): elemField.Text = "0,00": .appendChild elemField
                       
                       Set elemField = docStock.createElement("CHEQUEAMT"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("LCCHEQUE"): elemField.Text = " ": .appendChild elemField
                        
                       
                       Set elemField = docStock.createElement("TAXAMT1"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("TAXAMT2"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("TAXAMT3"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("TAXAMT4"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("TAXAMT5"): elemField.Text = " ": .appendChild elemField
                       
                       Set elemField = docStock.createElement("LCTAX1"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("LCTAX2"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("LCTAX3"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("LCTAX4"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("LCTAX5"): elemField.Text = " ": .appendChild elemField
                       
                       
                       Set elemField = docStock.createElement("KEPYOAMT"): elemField.Text = "": .appendChild elemField
                       Set elemField = docStock.createElement("ISBUILD"): elemField.Text = "0": .appendChild elemField
                       Set elemField = docStock.createElement("INBRCODE"): elemField.Text = "": .appendChild elemField
                       
                       
                       Set elemField = docStock.createElement("SUMKEPYOYP"): elemField.Text = cVAL(R!SKAU): .appendChild elemField
                       Set elemField = docStock.createElement("SUMKEPYONOTYP"): elemField.Text = "0,00": .appendChild elemField
                       Set elemField = docStock.createElement("SUMKEPYOFPA"): elemField.Text = cVAL(R!sfpa): .appendChild elemField
                       Set elemField = docStock.createElement("OTHEREXPEND"): elemField.Text = "": .appendChild elemField
                       
                       
                       Set elemField = docStock.createElement("CASHREGISTERID"): elemField.Text = "": .appendChild elemField
                       Set elemField = docStock.createElement("HASRETAILID"): elemField.Text = "": .appendChild elemField
                       Set elemField = docStock.createElement("CANCELGROUPID"): elemField.Text = "": .appendChild elemField
                       Set elemField = docStock.createElement("CANCELED"): elemField.Text = "": .appendChild elemField
                      
                       
                       Set elemDETAILS = docStock.createElement("DETAILS")
                       
'                          Set elemDETAIL = docStock.createElement("DETAIL")
'                                Set elemField = docStock.createElement("LCODE"): elemField.Text = "70-0087": elemDETAIL.appendChild elemField
'                                Set elemField = docStock.createElement("NETAMT"): elemField.Text = Format(R!aj2, "####0.00"): elemDETAIL.appendChild elemField
'                                Set elemField = docStock.createElement("VATAMT"): elemField.Text = Format(R!FPA2, "####0.00"): elemDETAIL.appendChild elemField
'                                Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
'                                Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
'                          elemDETAILS.appendChild elemDETAIL
                          
                         Dim LLK As Integer
                         For LLK = 1 To 7
                           If R("aj" + Format(LLK, "0")) > 0 Then
                             Set elemDETAIL = docStock.createElement("DETAIL")
                                Dim CC9 As String: CC9 = FIND_KOD(R!ATIM, LLK)  'FIND_KOD(R!ATIM, LLK)
                                Set elemField = docStock.createElement("LCODE"): elemField.Text = CC9: elemDETAIL.appendChild elemField
                                Set elemField = docStock.createElement("NETAMT"): elemField.Text = Replace(Format(R("aj" + Format(LLK, "0")), "#####0.00"), ".", ","): elemDETAIL.appendChild elemField
                                Set elemField = docStock.createElement("VATAMT"): elemField.Text = Replace(Format(R("FPA" + Format(LLK, "0")), "#####0.00"), ".", ","): elemDETAIL.appendChild elemField ' 0 AS FPA5 STO SQL OPOTE OK
                                Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
                                Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
                             elemDETAILS.appendChild elemDETAIL
                             Add_to_Array CC9, R("aj" + Format(LLK, "0")), R!MSIGN
                            End If
                         Next
                          
                          
                          
                          
                          
                          
                         .appendChild elemDETAILS
                         
                    
                    End With
                    elemARTICLES.appendChild elemARTICLE
                    R.MoveNext
                 Loop
                    
                    
                    
                    
                .appendChild elemARTICLES
                  
              
                    
               R.Close
               
                
                 Set elemCUSTOMERS = docStock.createElement("CUSTOMERS")
                     
                  'Exit Sub
                  
                  R.Open SQLPEL, Gdb, adOpenDynamic, adLockOptimistic
                   Do While Not R.EOF
                    Set elemCUSTOMER = docStock.createElement("CUSTOMER")
                     With elemCUSTOMER
                        Set elemField = docStock.createElement("ID"): elemField.Text = str(R!ID): .appendChild elemField
                        Set elemField = docStock.createElement("NAME"): elemField.Text = (R!EPO): .appendChild elemField
                        Set elemField = docStock.createElement("VAT"): elemField.Text = Trim((R!CAFM)): .appendChild elemField
                     
                     
                        Set elemField = docStock.createElement("JOB"): elemField.Text = (R!CEPA): .appendChild elemField
                        Set elemField = docStock.createElement("DOY"): elemField.Text = (R!Cdoy): .appendChild elemField
                        Set elemField = docStock.createElement("CUSTVAT"): elemField.Text = "2": .appendChild elemField
                    
                    
                        Set elemField = docStock.createElement("ADDRESS"): elemField.Text = R!cDIE: .appendChild elemField
                        Set elemField = docStock.createElement("ZIP"): elemField.Text = R!tk: .appendChild elemField
                        Set elemField = docStock.createElement("CITY"): elemField.Text = R!cpOL: .appendChild elemField
                    
                        Set elemField = docStock.createElement("PHONE1"): elemField.Text = R!cTHL: .appendChild elemField
                        Set elemField = docStock.createElement("PHONE2"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("PHONE3"): elemField.Text = " ": .appendChild elemField
                    
                     
                        Set elemField = docStock.createElement("FAX1"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("FAX2"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("EMAIL"): elemField.Text = " ": .appendChild elemField
                   
                        Set elemField = docStock.createElement("ISKEPYO"): elemField.Text = "0": .appendChild elemField
                        Set elemField = docStock.createElement("ISDHMOSIOU"): elemField.Text = "0": .appendChild elemField
                        Set elemField = docStock.createElement("ISEA"): elemField.Text = "0": .appendChild elemField
                     
                        Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": .appendChild elemField
                        Set elemField = docStock.createElement("ACCADDR"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("STRADDR"): elemField.Text = " ": .appendChild elemField
                     
                        Set elemField = docStock.createElement("STRNAME"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("BANK1"): elemField.Text = " ": .appendChild elemField
                         
                        Set elemField = docStock.createElement("BANKACC1"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("BANK2"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("BANKACC2"): elemField.Text = " ": .appendChild elemField
                         
                        Set elemField = docStock.createElement("EACOUNTRY"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("EAPREFIX"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("EAVAT"): elemField.Text = " ": .appendChild elemField
 
 
                        Set elemField = docStock.createElement("IDTYPE"): elemField.Text = " ": .appendChild elemField
 
                     
                     
                     End With
                    
                    elemCUSTOMERS.appendChild elemCUSTOMER
                    R.MoveNext
                    
                   Loop
                  
                .appendChild elemCUSTOMERS
                 
                     
'
'
'1116             For i = 0 To UBound(varStock)
'1118                 Set elemARTICLES = docStock.createElement("CUSTOMERS")
'1120                 With elemARTICLES
'1122                     Set elemField = docStock.createElement("ProductCode")
'1124                     elemField.Text = CStr(varStock(i)(0))
'1126                     .appendChild elemField
'1128                     Set elemField = docStock.createElement("ProductPrice")
'1130                     elemField.Text = CStr(varStock(i)(1))
'1132                     .appendChild elemField
'                    End With
'1134                 .appendChild elemARTICLES
'                Next
            End With
1136         Set .documentElement = elemDATA
            On Error Resume Next
'138         Kill "C:\created.xml"
'            On Error GoTo cmdCommand5_Click_Err
'140         .save "C:\created.xml"
'
'
'            Par7MyData.FormatXmlDocument docStock
'
'
            
            
            
138                     Kill "C:\created.xml"
            On Error GoTo cmdCommand5_Click_Err
140         '.save "C:\created.xml"


            Par7MyData.FormatXmlDocument docStock

           ' cd1.FILENAME = "GEF_" + Format(Now, "yyyy-MM-dd HH-mm") + ".xml"
           ' cd1.ShowSave
            CD1.FILENAME = "GEFB_" + Format(Now, "yyyy-MM-dd HH-mm") + ".xml"
            CD1.ShowSave
            
        ' .save cd1.FILENAME '"C:\created.xml"
         .save CD1.FILENAME '"C:\created.xml"


            
            
            
            
            
            
            





        End With
    
    MsgBox "ok"
    
    
    
        '<EhFooter>
        Exit Sub

cmdCommand5_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.Par7MyData.cmdCommand5_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>

























End Sub




Private Sub makeXML_KATHG(ByVal sql As String, ByVal sqlwh As String, ByVal SQLPEL As String)
'=============================================================================================================

Dim R As New ADODB.Recordset

   Dim rk As New ADODB.Recordset
            
            
            
            
            R.Open sqlwh + " order by HME", Gdb, adOpenDynamic, adLockOptimistic
            

 





Dim mpelP As String
Dim mpela As String
mpelP = InputBox("Κωδ.Πελατη ", "", "30-00-00-0000")

mpela = InputBox("Κωδ.Προμηθευτη ", "", "50-00-00-0000")







     Dim varStock As Variant
       'global       Dim docStock As MSXML2.DOMDocument
        Dim elemDATA As MSXML2.IXMLDOMElement
        Dim elemARTICLES As MSXML2.IXMLDOMElement
        Dim elemARTICLE As MSXML2.IXMLDOMElement
         
         
        Dim elemDETAILS As MSXML2.IXMLDOMElement
        Dim elemDETAIL As MSXML2.IXMLDOMElement
         
          Dim elemCUSTOMERS As MSXML2.IXMLDOMElement
          Dim elemCUSTOMER As MSXML2.IXMLDOMElement
          
        Dim elemField As MSXML2.IXMLDOMElement
        Dim I As Integer
    

102     Set docStock = New MSXML2.DOMDocument
104     With docStock
106         .appendChild .createProcessingInstruction("xml", _
                                                      "version=""1.0"" encoding=""utf-8""")
108         Set elemDATA = .createElement("DATA")
110         With elemDATA
112             '.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
114             '.setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
                 
                            
                            
                            
                 Set elemARTICLES = docStock.createElement("ARTICLES")
                 
                 
                 
                 Dim mType As String
                 Dim pel30 As String
                 
                 Do While Not R.EOF 'For k = 1 To 5
                    Me.Caption = Format(R!hme, "dd/MM/yyyy")
                    Set elemARTICLE = docStock.createElement("ARTICLE"): elemARTICLE.Text = CStr(k):
                    With elemARTICLE
                    
'                       If Trim(R!EIDOS) = "e" Then
'                          mType = "0"
'                          pel30 = mpelP  '"30-0000"
'                       Else
'                          mType = "1"
'                          pel30 = mpela ' "50-0000"
'                       End If
                       
                     If Trim(R!EIDOS) = "e" Then
                          mType = "0"
                          pel30 = mpelP  '"30-0000"
                       Else
                          mType = "1"
                          pel30 = mpela ' "50-0000"
                       End If
                       
                       
                       
                       
                       Set elemField = docStock.createElement("MTYPE"): elemField.Text = mType: .appendChild elemField
                      
                       Set elemField = docStock.createElement("MSIGN"): elemField.Text = str(R!MSIGN): .appendChild elemField
                       Set elemField = docStock.createElement("ISKEPYO"): elemField.Text = "1": .appendChild elemField
                       Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": .appendChild elemField
                       
                       
                       Set elemField = docStock.createElement("LCODE"): elemField.Text = pel30: .appendChild elemField
                       Set elemField = docStock.createElement("CUSTID"): elemField.Text = str(R!PELID): .appendChild elemField
                       Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: .appendChild elemField
                       Set elemField = docStock.createElement("MDATE"): elemField.Text = Format(R!hme, "dd/MM/yyyy"): .appendChild elemField
                       Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: .appendChild elemField
                       
                       Set elemField = docStock.createElement("CASHAMT"): elemField.Text = "0,00": .appendChild elemField
                       Set elemField = docStock.createElement("LCCASH"): elemField.Text = "0,00": .appendChild elemField
                       
                       Set elemField = docStock.createElement("CHEQUEAMT"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("LCCHEQUE"): elemField.Text = " ": .appendChild elemField
                        
                       
                       Set elemField = docStock.createElement("TAXAMT1"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("TAXAMT2"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("TAXAMT3"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("TAXAMT4"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("TAXAMT5"): elemField.Text = " ": .appendChild elemField
                       
                       Set elemField = docStock.createElement("LCTAX1"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("LCTAX2"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("LCTAX3"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("LCTAX4"): elemField.Text = " ": .appendChild elemField
                       Set elemField = docStock.createElement("LCTAX5"): elemField.Text = " ": .appendChild elemField
                       
                       
                       Set elemField = docStock.createElement("KEPYOAMT"): elemField.Text = "": .appendChild elemField
                       Set elemField = docStock.createElement("ISBUILD"): elemField.Text = "0": .appendChild elemField
                       Set elemField = docStock.createElement("INBRCODE"): elemField.Text = "": .appendChild elemField
                       
                       
                       Set elemField = docStock.createElement("SUMKEPYOYP"): elemField.Text = cVAL(R!SKAU): .appendChild elemField
                       Set elemField = docStock.createElement("SUMKEPYONOTYP"): elemField.Text = "0,00": .appendChild elemField
                       Set elemField = docStock.createElement("SUMKEPYOFPA"): elemField.Text = cVAL(R!sfpa): .appendChild elemField
                       Set elemField = docStock.createElement("OTHEREXPEND"): elemField.Text = "": .appendChild elemField
                       
                       
                       Set elemField = docStock.createElement("CASHREGISTERID"): elemField.Text = "": .appendChild elemField
                       Set elemField = docStock.createElement("HASRETAILID"): elemField.Text = "": .appendChild elemField
                       Set elemField = docStock.createElement("CANCELGROUPID"): elemField.Text = "": .appendChild elemField
                       Set elemField = docStock.createElement("CANCELED"): elemField.Text = "": .appendChild elemField
                      
                       
                       Set elemDETAILS = docStock.createElement("DETAILS")
                       
                        
                          
                          
                          Dim sqlk As String
                          
                          sqlk = "SELECT ISNULL(KATHGORIA,0) AS CAT,G.FPA,ROUND(SUM(ISNULL(KAU_AJIA,0)),2) AS KAU_AJIA,ROUND(SUM(ISNULL(MIK_AJIA,0)),2) AS MIK_AJIA,"
                          sqlk = sqlk + " ID_NUM FROM EGGTIM G  INNER JOIN EID D  ON G.KODE=D.KOD"
                          sqlk = sqlk + "  WHERE  ID_NUM=" + str(R!id_num) + " GROUP BY ID_NUM, G.FPA,KATHGORIA"
                          
                          
                          
                          rk.Open sqlk, Gdb, adOpenDynamic, adLockOptimistic
                          Do While Not rk.EOF
                             If rk!kau_ajia > 0 Then
                                Set elemDETAIL = docStock.createElement("DETAIL")
                                Set elemField = docStock.createElement("LCODE"): elemField.Text = FINDkat_KOD(R!ATIM, rk!FPA, rk!CAT): elemDETAIL.appendChild elemField
                                Set elemField = docStock.createElement("NETAMT"): elemField.Text = Format(rk("KAU_AJIA"), "####0.00"): elemDETAIL.appendChild elemField
                                Set elemField = docStock.createElement("VATAMT"): elemField.Text = Format(rk("MIK_AJIA") - rk("KAU_AJIA"), "####0.00"): elemDETAIL.appendChild elemField ' 0 AS FPA5 STO SQL OPOTE OK
                                Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
                                Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
                                elemDETAILS.appendChild elemDETAIL
                             End If
                             rk.MoveNext
                          Loop
                         rk.Close
                         
                          
                          
                          
                          
                          
                         .appendChild elemDETAILS
                         
                    
                    End With
                    elemARTICLES.appendChild elemARTICLE
                    R.MoveNext
                 Loop
                    
                    
                    
                    
                .appendChild elemARTICLES
                  
              
                    
               R.Close
               
                
                 Set elemCUSTOMERS = docStock.createElement("CUSTOMERS")
                     
                  'Exit Sub
                  
                  R.Open SQLPEL, Gdb, adOpenDynamic, adLockOptimistic
                   Do While Not R.EOF
                    Set elemCUSTOMER = docStock.createElement("CUSTOMER")
                     With elemCUSTOMER
                        Set elemField = docStock.createElement("ID"): elemField.Text = str(R!ID): .appendChild elemField
                        Set elemField = docStock.createElement("NAME"): elemField.Text = (R!EPO): .appendChild elemField
                        Set elemField = docStock.createElement("VAT"): elemField.Text = Trim((R!CAFM)): .appendChild elemField
                     
                     
                        Set elemField = docStock.createElement("JOB"): elemField.Text = (R!CEPA): .appendChild elemField
                        Set elemField = docStock.createElement("DOY"): elemField.Text = (R!Cdoy): .appendChild elemField
                        Set elemField = docStock.createElement("CUSTVAT"): elemField.Text = "2": .appendChild elemField
                    
                    
                        Set elemField = docStock.createElement("ADDRESS"): elemField.Text = R!cDIE: .appendChild elemField
                        Set elemField = docStock.createElement("ZIP"): elemField.Text = R!tk: .appendChild elemField
                        Set elemField = docStock.createElement("CITY"): elemField.Text = R!cpOL: .appendChild elemField
                    
                        Set elemField = docStock.createElement("PHONE1"): elemField.Text = R!cTHL: .appendChild elemField
                        Set elemField = docStock.createElement("PHONE2"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("PHONE3"): elemField.Text = " ": .appendChild elemField
                    
                     
                        Set elemField = docStock.createElement("FAX1"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("FAX2"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("EMAIL"): elemField.Text = " ": .appendChild elemField
                   
                        Set elemField = docStock.createElement("ISKEPYO"): elemField.Text = "0": .appendChild elemField
                        Set elemField = docStock.createElement("ISDHMOSIOU"): elemField.Text = "0": .appendChild elemField
                        Set elemField = docStock.createElement("ISEA"): elemField.Text = "0": .appendChild elemField
                     
                        Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": .appendChild elemField
                        Set elemField = docStock.createElement("ACCADDR"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("STRADDR"): elemField.Text = " ": .appendChild elemField
                     
                        Set elemField = docStock.createElement("STRNAME"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("BANK1"): elemField.Text = " ": .appendChild elemField
                         
                        Set elemField = docStock.createElement("BANKACC1"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("BANK2"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("BANKACC2"): elemField.Text = " ": .appendChild elemField
                         
                        Set elemField = docStock.createElement("EACOUNTRY"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("EAPREFIX"): elemField.Text = " ": .appendChild elemField
                        Set elemField = docStock.createElement("EAVAT"): elemField.Text = " ": .appendChild elemField
 
 
                        Set elemField = docStock.createElement("IDTYPE"): elemField.Text = " ": .appendChild elemField
 
                     
                     
                     End With
                    
                    elemCUSTOMERS.appendChild elemCUSTOMER
                    R.MoveNext
                    
                   Loop
                  
                .appendChild elemCUSTOMERS
                 
                     

            End With
1136         Set .documentElement = elemDATA
            On Error Resume Next
'138         Kill "C:\created.xml"
'            On Error GoTo cmdCommand5_Click_Err
'140         '.save "C:\created.xml"
'
'
'            Par7MyData.FormatXmlDocument docStock

138              Kill "C:\created.xml"
            On Error GoTo cmdCommand5_Click_Err
140         '.save "C:\created.xml"


            Par7MyData.FormatXmlDocument docStock

           ' cd1.FILENAME = "GEF_" + Format(Now, "yyyy-MM-dd HH-mm") + ".xml"
           ' cd1.ShowSave
            CD1.FILENAME = "GEFB_" + Format(Now, "yyyy-MM-dd HH-mm") + ".xml"
            CD1.ShowSave
            
        ' .save cd1.FILENAME '"C:\created.xml"
         .save CD1.FILENAME '"C:\created.xml"

       





        End With
    
    MsgBox "ok"
    
    
    
        '<EhFooter>
        Exit Sub

cmdCommand5_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.Par7MyData.cmdCommand5_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>


End Sub






Private Sub GmakeXML(ByVal sql As String, ByVal sqlwh As String, ByVal SQLPEL As String, ByVal withcateg As Integer, ByVal ONLYXREOPIS As Integer)
  '=============================================================================================================
'On Error Resume Next
'Gdb.CommitTrans
List1.Clear
DoEvents


'Gdb.BeginTrans
    On Error GoTo cmdCommand5_Click_Err
    ' γ κατηγορια βιβλια
    Dim R As New ADODB.Recordset
          '
           
          ' Exit Sub
           
          '
           
          ' Exit Sub
          
          Dim iin As Integer
          iin = 0
           
           
           Dim mProbl As String
           Dim rr0 As New ADODB.Recordset
           'If ONLYXREOPIS = 0 Then
                ' rr0.Open sql + sqlwh + "  AND TIM.EIDOS='r' ", Gdb, adOpenDynamic, adLockOptimistic
                 'If Par7MyData.checkIntegrity(1, rr0, mProbl) = False Then
                  ' If mProbl = "" Then
                   '   mProbl = "ΠΡΟΒΛΗΜΑ ΑΚΕΡΑΙΟΤΗΤΑΣ "
                  ' End If
                 'End If
                 
                 ' If Par7MyData.checkIntegrity(1, rr0, mProbl) = False Then
                 
                  ' End If
                 
                 
                ' rr0.Close
                 
           'End If
           
           
           
           
           
           
           
           
           
           
           
           
           
           Dim Debugg As String
           Debugg = "  " ' AND Left(ATIM, 7) = 'υ090570' "
100        R.Open sql + sqlwh + Debugg + "  order by HME", Gdb, adOpenDynamic, adLockOptimistic
           
          ' Exit Sub
           Dim RRC2 As New ADODB.Recordset
           RRC2.Open "SELECT COUNT(*) " + sqlwh, Gdb, adOpenDynamic, adLockOptimistic
          ' Exit Sub
           
           If RRC2(0) = 0 Then
              MsgBox "ΔΕΝ ΒΡΕΘΗΚΑΝ ΕΓΓΡΑΦΕΣ"
              
              RRC2.Close
              Exit Sub
           End If
           RRC2.Close
           
           
           
           
           
     Dim varStock As Variant
       'global       Dim docStock As MSXML2.DOMDocument
        Dim elemDATA As MSXML2.IXMLDOMElement
        Dim elemARTICLES As MSXML2.IXMLDOMElement
        Dim elemARTICLE As MSXML2.IXMLDOMElement
         
         
        Dim elemDETAILS As MSXML2.IXMLDOMElement
        Dim elemDETAIL As MSXML2.IXMLDOMElement
         
          Dim elemCUSTOMERS As MSXML2.IXMLDOMElement
          Dim elemCUSTOMER As MSXML2.IXMLDOMElement
          
        Dim elemField As MSXML2.IXMLDOMElement
        Dim I As Integer
    

102     Set docStock = New MSXML2.DOMDocument
104     With docStock
106         .appendChild .createProcessingInstruction("xml", _
                                                      "version=""1.0"" encoding=""utf-8""")
108         Set elemDATA = .createElement("DATA")
110         With elemDATA
                '.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
                '.setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
112              Set elemARTICLES = docStock.createElement("ARTICLES")
                 Dim mType As String
                 Dim pel30 As String
                 Dim SUMX As Double, SUMP As Double
114              Do While Not R.EOF 'For k = 1 To 5
116                 Me.Caption = Format(R!hme, "dd/MM/yyyy") + str(iin)
                    Set elemARTICLE = docStock.createElement("ARTICLE"): elemARTICLE.Text = CStr(k):
118                 With elemARTICLE
'                      If Left(R!ATIM, 7) = "υ090544" Then
'                        Dim DUMM As Integer
'                        DUMM = 1
'                      End If
                    
120                   If ONLYXREOPIS = 1 Then
                           'apo thn kathgoria pelath
122                        pel30 = FINDkat_PEL(R!EIDOS, R!KPE, R!KODGAL)


                                'Dim m_kodtam2 As String
                                'm_kodtam2 = GGET_CVALUE("select TRIM(ISNULL(KODLOG,'00-00')) AS KODLOG from XREOPIS WHERE EIDOS=" + str(R!EIDXPI) + "") 'f_kod_tameio
                                If R!KODLOG = "00-00" Or R!EIDXPI = 0 Then
                                    List1.AddItem R!ATIM + "ΔΕΝ ΕΧΩ ΚΩΔ.ΛΟΓ.XREOPIS ΚΩΔ." + str(R!EIDXPI) + " HMEP:" + Format(R!hme, "DD/MM/YYYY") + "ΚΩΔ." + R!KPE + " ID=" + str(R!EGGID)
                                End If
                                If Left(R!KODLOG, 8) = "38-00-00" Then
                                       mType = "14"
                                Else
                                       mType = "11"
                                End If
                                
                      
                    
                      Else
                    
                    
                    
                        ' παιρνει τουσ κωδικους απο τις παραμετρους παραστατικου
124                     If R!EIDOS = "e" Then
126                       mType = "13"
                         ' If chk2.Value = vbChecked Then
                            ' pel30 = txt.Text + Left$(Trim(R!KPE) + "0000", 4)
                         ' Else
                            ' pel30 = txt.Text + txt22.Text ' "30-00-00-0000"
128                          pel30 = FIND_apo_Parastatiko_PELKOD(R!ATIM, 0, Left$(Trim(R!KPE) + "0000", 4))
                         ' End If
                          
                        Else
130                       mType = "12"
132                       pel30 = FIND_apo_Parastatiko_PELKOD(R!ATIM, 0, Left$(Trim(R!KPE) + "0000", 4))
                         ' pel30 = "50-00-00-0000"
                        
                        End If
                       
                        'If withcateg = 1 Then
                         '  pel30 = FIND_apo_Parastatiko_PELKOD(R!ATIM, 0, Left$(Trim(R!KPE) + "0000", 4))
                       ' End If
                       
                       
    '                       If ONLYXREOPIS = 1 Then
    '                         If R!EIDOS = "e" Then
    '                           pel30 = "30-00-00-0000"
    '                         Else
    '                           pel30 = "50-00-00-0000"
    '                         End If
    '                       End If
                       
                      End If
                     
       
                       
                       
                       
134                    Set elemField = docStock.createElement("MTYPE"): elemField.Text = mType: .appendChild elemField
                      
136                    Set elemField = docStock.createElement("MSIGN"): elemField.Text = str(R!MSIGN): .appendChild elemField

                If ONLYXREOPIS = 1 Then
                     Set elemField = docStock.createElement("ISKEPYO"): elemField.Text = "0": .appendChild elemField
                Else
                   If R!EIDOS = "e" Then
138                    Set elemField = docStock.createElement("ISKEPYO"): elemField.Text = "1": .appendChild elemField
                   Else
                       Set elemField = docStock.createElement("ISKEPYO"): elemField.Text = "2": .appendChild elemField
                   End If
                End If
140                    Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "1": .appendChild elemField ' υποκαταστημα???
                       
                       
142                    Set elemField = docStock.createElement("LCODE"): elemField.Text = pel30: .appendChild elemField
144                    Set elemField = docStock.createElement("CUSTID"): elemField.Text = str(R!PELID): .appendChild elemField
146                    Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: .appendChild elemField
148                    Set elemField = docStock.createElement("MDATE"): elemField.Text = Format(R!hme, "dd/MM/yyyy"): .appendChild elemField
150                    Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: .appendChild elemField
                       
152                    Set elemField = docStock.createElement("CASHAMT"): elemField.Text = "0,00": .appendChild elemField
154                    Set elemField = docStock.createElement("LCCASH"): elemField.Text = "0,00": .appendChild elemField
                       
156                    Set elemField = docStock.createElement("CHEQUEAMT"): elemField.Text = " ": .appendChild elemField
158                    Set elemField = docStock.createElement("LCCHEQUE"): elemField.Text = " ": .appendChild elemField
                        
                       
160                    Set elemField = docStock.createElement("TAXAMT1"): elemField.Text = " ": .appendChild elemField
162                    Set elemField = docStock.createElement("TAXAMT2"): elemField.Text = " ": .appendChild elemField
164                    Set elemField = docStock.createElement("TAXAMT3"): elemField.Text = " ": .appendChild elemField
166                    Set elemField = docStock.createElement("TAXAMT4"): elemField.Text = " ": .appendChild elemField
168                    Set elemField = docStock.createElement("TAXAMT5"): elemField.Text = " ": .appendChild elemField
                       
170                    Set elemField = docStock.createElement("LCTAX1"): elemField.Text = " ": .appendChild elemField
172                    Set elemField = docStock.createElement("LCTAX2"): elemField.Text = " ": .appendChild elemField
174                    Set elemField = docStock.createElement("LCTAX3"): elemField.Text = " ": .appendChild elemField
176                    Set elemField = docStock.createElement("LCTAX4"): elemField.Text = " ": .appendChild elemField
178                    Set elemField = docStock.createElement("LCTAX5"): elemField.Text = " ": .appendChild elemField
                       
                       
180                    Set elemField = docStock.createElement("KEPYOAMT"): elemField.Text = "": .appendChild elemField
182                    Set elemField = docStock.createElement("ISBUILD"): elemField.Text = "0": .appendChild elemField
184                    Set elemField = docStock.createElement("INBRCODE"): elemField.Text = "": .appendChild elemField
                       
                       
186                    Set elemField = docStock.createElement("SUMKEPYOYP"): elemField.Text = cVAL(R!SKAU): .appendChild elemField
188                    Set elemField = docStock.createElement("SUMKEPYONOTYP"): elemField.Text = "0,00": .appendChild elemField
190                    Set elemField = docStock.createElement("SUMKEPYOFPA"): elemField.Text = cVAL(R!sfpa): .appendChild elemField
192                    Set elemField = docStock.createElement("OTHEREXPEND"): elemField.Text = "": .appendChild elemField
                       
                       
194                    Set elemField = docStock.createElement("CASHREGISTERID"): elemField.Text = "": .appendChild elemField
196                    Set elemField = docStock.createElement("HASRETAILID"): elemField.Text = "": .appendChild elemField
198                    Set elemField = docStock.createElement("CANCELGROUPID"): elemField.Text = "": .appendChild elemField
200                    Set elemField = docStock.createElement("CANCELED"): elemField.Text = "": .appendChild elemField
                      
                       
202                    Set elemDETAILS = docStock.createElement("DETAILS")
                       
                       
                       
                       Dim CRDB3050 As String, CRDB2070 As String
204                    If R!EIDOS = "e" Then
206                      If ONLYXREOPIS = 1 Then
                           CRDB3050 = "0": CRDB2070 = "1"

                         Else
                         
                           If R!MSIGN = -1 Then
                               CRDB3050 = "1": CRDB2070 = "0"
                           Else
                               CRDB3050 = "0": CRDB2070 = "1"
                           End If
                           
                         End If
                           
                       Else
                          If ONLYXREOPIS = 1 Then
                             CRDB3050 = "1": CRDB2070 = "0"
                          Else
                             If R!MSIGN = -1 Then
                                CRDB3050 = "0": CRDB2070 = "1"
                             Else
                               CRDB3050 = "1": CRDB2070 = "0"
                             End If
                          End If
                       
                       
                       
208                        'CRDB3050 = "1": CRDB2070 = "0"
                       End If
                       
                       
                       
                       
210                   If ONLYXREOPIS = 1 Then  'B_C1
                          
                      ' ----   εισπραξη
                      '--- TON 30 TON PAIRNEI APO THN KATHGORIA PELATH KODGAL =>PINAKES C3
212                    If Len(FINDtrp_KOD(Val(Left(R!trp, 2)))) > 2 Then ' Val(Left(R!trp, 2)) = 1 Then
214                     Set elemDETAIL = docStock.createElement("DETAIL")
216                             Set elemField = docStock.createElement("LCODE"): elemField.Text = pel30: elemDETAIL.appendChild elemField
218                             Set elemField = docStock.createElement("CRDB"): elemField.Text = CRDB2070: elemDETAIL.appendChild elemField
220                             Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(R!SKAU + R!sfpa, "####0.00"): elemDETAIL.appendChild elemField
222                             Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: elemDETAIL.appendChild elemField
224                             Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: elemDETAIL.appendChild elemField
226                             Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
228                             Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
230                       elemDETAILS.appendChild elemDETAIL
                          ' tameio 38-00-*******
232                      Set elemDETAIL = docStock.createElement("DETAIL")
                         
                              ' Dim m_kodtam As String
234                           ' m_kodtam = GGET_CVALUE("select TRIM(ISNULL(KODLOG,'00-00')) AS KODLOG from XREOPIS WHERE EIDOS=" + str(R!EIDXPI) + "") 'f_kod_tameio
                               
236                           ' If m_kodtam = "00-00" Or Len(Trim(m_kodtam)) = 0 Then
                               '                    List1.AddItem "-ΔΕΝ ΕΧΩ ΚΩΔ.ΛΟΓ.XREOPIS ΚΩΔ." + str(R!EIDXPI) + " HMEP:" + Format(R!hme, "DD/MM/YYYY") + "ΚΩΔ." + R!KPE + " ID=" + str(R!EGGID)
                
238                                'List1.AddItem "ΔΕΝ ΕΧΩ ΚΩΔΙΚΟ ΛΟΓΙΣΤΙΚΗΣ XREOPIS ΚΩΔ." + str(R!EIDXPI)
                              ' End If
                               
                         
                         
240                             Set elemField = docStock.createElement("LCODE"): elemField.Text = Trim(R!KODLOG): elemDETAIL.appendChild elemField
242                             Set elemField = docStock.createElement("CRDB"): elemField.Text = CRDB3050: elemDETAIL.appendChild elemField
244                             Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(R!SKAU + R!sfpa, "####0.00"): elemDETAIL.appendChild elemField
246                             Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: elemDETAIL.appendChild elemField
248                             Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: elemDETAIL.appendChild elemField
250                             Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
252                             Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
254                       elemDETAILS.appendChild elemDETAIL
                        End If
                        ' ----   εισπραξη
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      Else
                         
                          If CRDB3050 = 1 Then SUMP = SUMP + R!SKAU + R!sfpa Else SUMX = SUMX + R!SKAU + R!sfpa
                          
                          
                          
256                       Set elemDETAIL = docStock.createElement("DETAIL")
258                             Set elemField = docStock.createElement("LCODE"): elemField.Text = pel30: elemDETAIL.appendChild elemField
260                             Set elemField = docStock.createElement("CRDB"): elemField.Text = CRDB3050: elemDETAIL.appendChild elemField
262                             Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(R!SKAU + R!sfpa, "####0.00"): elemDETAIL.appendChild elemField
264                             Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: elemDETAIL.appendChild elemField
266                             Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: elemDETAIL.appendChild elemField
268                             Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
270                             Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
272                       elemDETAILS.appendChild elemDETAIL
                          
    '                          <DETAIL>
       '                    <LCODE>50.00.00.0000</LCODE>
    '                      <CRDB>1</CRDB>
    '                       <AMOUNT>1240</AMOUNT>
    '                      <INVOICE>ΤΔΑ 12</INVOICE>
    '                        <REASON>ΠΑΠΑΔΟΠΟΥΛΟΣ</REASON>
    '                      <ISAGRYP></ISAGRYP>
    '                       <KEPYOPARTY>100</KEPYOPARTY>
    '                        </DETAIL>
                     
                     
                     
                                     
                     If Left(R!ATIM, 7) = "G001226" Or Left(R!ATIM, 7) = "G210078" Or Left(R!ATIM, 7) = "G021074" Or Left(R!ATIM, 7) = "G021101" Then
                          
'                           If ONLYXREOPIS = 0 Then
'                                   ' rr0.Open "SELECT * FROM TIM WHERE ID_NUM=" + str(R!id_num), Gdb, adOpenDynamic, adLockOptimistic
'
'                                   'rr0.Close
'
'                                   rr0.Open "SELECT * FROM TIM WHERE ID_NUM=" + str(R!id_num), Gdb, adOpenDynamic, adLockOptimistic
'                                   If Par7MyData.checkIntegrity(1, rr0, mProbl) = False Then
'                                      If mProbl = "" Then
'                                         mProbl = "ΠΡΟΒΛΗΜΑ ΑΚΕΡΑΙΟΤΗΤΑΣ "
'                                      End If
'                                   End If
'                                   rr0.Close
'                            End If
                      End If
                     
                     
                     
274                    If withcateg = 1 Then  ' αν παει κατα κατηγορία
                          Dim sqlk As String

276                        sqlk = "SELECT ISNULL(KATHGORIA,1) AS CAT,G.FPA,ROUND(SUM(ISNULL(KAU_AJIA,0)),2) AS KAU_AJIA,ROUND(SUM(ISNULL(MIK_AJIA,0)),2) AS MIK_AJIA,"
278                       sqlk = sqlk + " ID_NUM FROM EGGTIM G  INNER JOIN EID D  ON G.KODE=D.KOD"
280                       sqlk = sqlk + "  WHERE  ID_NUM=" + str(R!id_num) + " GROUP BY ID_NUM, G.FPA,KATHGORIA"



                         Dim rk As New ADODB.Recordset
282                       rk.Open sqlk, Gdb, adOpenDynamic, adLockOptimistic
                          
                          
                          
                          
                          Dim s1k As Single, s1m As Single, nnl As Integer
284
                           'nnl = rk.RecordCount ' debug
                           s1m = 0: s1k = 0: nnl = 0
286                       Do While Not rk.EOF
288                           nnl = nnl + 1
290                           s1k = s1k + rk!kau_ajia
292                           s1m = s1m + rk!MIK_AJIA
294                           rk.MoveNext
                          Loop
                          
                          Dim diff As Single
                          Dim diffk As Single
                          Dim difff As Single
296                       If Abs(s1m - R!aji) > 0.005 Then
                                If ONLYXREOPIS = 0 Then
                                   rr0.Open "SELECT * FROM TIM WHERE ID_NUM=" + str(R!id_num), Gdb, adOpenDynamic, adLockOptimistic
                                   If Par7MyData.checkIntegrity(1, rr0, mProbl) = False Then
                                      If mProbl = "" Then
                                         mProbl = "ΠΡΟΒΛΗΜΑ ΑΚΕΡΑΙΟΤΗΤΑΣ "
                                      End If
                                   End If
                                   rr0.Close
                                End If

                                s1m = 0: s1k = 0: nnl = 0
                                rk.MoveFirst
                                Do While Not rk.EOF
                                    nnl = nnl + 1
                                    s1k = s1k + rk!kau_ajia
                                    s1m = s1m + rk!MIK_AJIA
                                    rk.MoveNext
                                Loop
298                          diff = s1m - R!aji '001
300                          diffk = s1k - (R!aj1 + R!aj2 + R!aj3 + R!aj4 + R!aj5 + R!aj6 + R!AJ7) '0.01
302                          difff = diff  '0.01
304                          Debug.Print R!ATIM + " DIFF=" + str(Round(diff, 2))
                          End If
                          
                          
                          
                          
306                       rk.MoveFirst
308                       Do While Not rk.EOF
310                          If rk!kau_ajia > 0 Then
                                If CRDB2070 = 1 Then SUMP = SUMP + rk("KAU_AJIA") - diffk Else SUMX = SUMX + rk("KAU_AJIA") - diffk
312                             Set elemDETAIL = docStock.createElement("DETAIL")
314                             Set elemField = docStock.createElement("LCODE"): elemField.Text = FINDkat_KOD(R!ATIM, rk!FPA, rk!CAT): elemDETAIL.appendChild elemField
316                             Set elemField = docStock.createElement("CRDB"): elemField.Text = CRDB2070: elemDETAIL.appendChild elemField
                               ' Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(rk("KAU_AJIA"), "####0.00"): elemDETAIL.appendChild elemField
318                             Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(rk("KAU_AJIA") - diffk, "####0.00"): elemDETAIL.appendChild elemField
320                             diffk = 0
322                             Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: elemDETAIL.appendChild elemField
324                             Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: elemDETAIL.appendChild elemField
326                             Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
328                             Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
330                             elemDETAILS.appendChild elemDETAIL
                             End If
332                          If rk!MIK_AJIA - rk!kau_ajia > 0 Then
                                If CRDB2070 = 1 Then SUMP = SUMP + rk!MIK_AJIA - rk!kau_ajia - difff Else SUMX = SUMX + rk!MIK_AJIA - rk!kau_ajia - difff
334                             Set elemDETAIL = docStock.createElement("DETAIL")
336                             Set elemField = docStock.createElement("LCODE"): elemField.Text = FINDFPAkat_KOD(R!ATIM, rk!FPA, rk!CAT): elemDETAIL.appendChild elemField
338                             Set elemField = docStock.createElement("CRDB"): elemField.Text = CRDB2070: elemDETAIL.appendChild elemField
                               ' Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(rk!MIK_AJIA - rk!KAU_AJIA, "####0.00"): elemDETAIL.appendChild elemField
340                             Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(rk!MIK_AJIA - rk!kau_ajia - difff, "####0.00"): elemDETAIL.appendChild elemField
342                             difff = 0
344                             Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: elemDETAIL.appendChild elemField
346                             Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: elemDETAIL.appendChild elemField
348                             Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
350                             Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
352                             elemDETAILS.appendChild elemDETAIL
                             End If
                             
                             
                             
                             
                             
354                          rk.MoveNext
                          Loop
                          
                          If Abs(SUMX - SUMP) > 0.005 Then
                             List1.AddItem "ΑΝΙΣ/ΣΤΟ " + Format(R!hme, "dd/MM/yyyy") + " διαφ=" + R!ATIM + Format(SUMX - SUMP, "###0.00") + " Χ=" + R!ATIM + Format(SUMX, "###0.00") + " Π=" + R!ATIM + Format(SUMP, "###0.00") + " τιμ=" + Format(R!aji, "###0.00")
                             If nnl = 0 Then
                                List1.AddItem "προβλημα με κατηγ.ειδους " + sqlk
                             End If
                          End If
                          
                          
                          
356                      rk.Close
                          
                          
                          
                       Else
                          
                         Dim LLK As Integer
358                      For LLK = 1 To 7
360                        If R("aj" + Format(LLK, "0")) > 0 Then
                             ' καθαρη αξια
362                          Set elemDETAIL = docStock.createElement("DETAIL")
364                             Set elemField = docStock.createElement("LCODE"): elemField.Text = pel30: elemDETAIL.appendChild elemField
366                             Set elemField = docStock.createElement("CRDB"): elemField.Text = CRDB2070: elemDETAIL.appendChild elemField
368                             Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(R("aj" + Format(LLK, "0")), "####0.00"): elemDETAIL.appendChild elemField
370                             Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: elemDETAIL.appendChild elemField
372                             Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: elemDETAIL.appendChild elemField
374                             Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
376                             Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
378                          elemDETAILS.appendChild elemDETAIL
                            'φπα
380                         If LLK <> 5 Then
382                          Set elemDETAIL = docStock.createElement("DETAIL")
384                             Set elemField = docStock.createElement("LCODE"): elemField.Text = FIND_fpa(R!ATIM, LLK): elemDETAIL.appendChild elemField
386                             Set elemField = docStock.createElement("CRDB"): elemField.Text = CRDB2070: elemDETAIL.appendChild elemField
388                             Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(R("fpa" + Format(LLK, "0")), "####0.00"): elemDETAIL.appendChild elemField
390                             Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: elemDETAIL.appendChild elemField
392                             Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: elemDETAIL.appendChild elemField
394                             Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
396                             Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
398                          elemDETAILS.appendChild elemDETAIL
                            End If


                            End If
                         Next
                      End If
                      
                          
                      ' ----   εισπραξη
                      '--- TON 30 TON PAIRNEI APO THN KATHGORIA PELATH KODGAL =>PINAKES C3
400                    If Len(FINDtrp_KOD(Val(Left(R!trp, 2)))) > 2 Then ' Val(Left(R!trp, 2)) = 1 Then
402                     Set elemDETAIL = docStock.createElement("DETAIL")
404                             Set elemField = docStock.createElement("LCODE"): elemField.Text = pel30: elemDETAIL.appendChild elemField
406                             Set elemField = docStock.createElement("CRDB"): elemField.Text = CRDB2070: elemDETAIL.appendChild elemField
408                             Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(R!SKAU + R!sfpa, "####0.00"): elemDETAIL.appendChild elemField
410                             Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: elemDETAIL.appendChild elemField
412                             Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: elemDETAIL.appendChild elemField
414                             Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
416                             Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
418                       elemDETAILS.appendChild elemDETAIL
                          
420                      Set elemDETAIL = docStock.createElement("DETAIL")
422                             Set elemField = docStock.createElement("LCODE"): elemField.Text = FINDtrp_KOD(Val(Left(R!trp, 2))): elemDETAIL.appendChild elemField
424                             Set elemField = docStock.createElement("CRDB"): elemField.Text = CRDB3050: elemDETAIL.appendChild elemField
426                             Set elemField = docStock.createElement("AMOUNT"): elemField.Text = Format(R!SKAU + R!sfpa, "####0.00"): elemDETAIL.appendChild elemField
428                             Set elemField = docStock.createElement("INVOICE"): elemField.Text = R!ATIM: elemDETAIL.appendChild elemField
430                             Set elemField = docStock.createElement("REASON"): elemField.Text = R!EPO: elemDETAIL.appendChild elemField
432                             Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": elemDETAIL.appendChild elemField
434                             Set elemField = docStock.createElement("KEPYOPARTY"): elemField.Text = "100": elemDETAIL.appendChild elemField
436                       elemDETAILS.appendChild elemDETAIL
                        End If
                        ' ----   εισπραξη
                      
                      
                      
                     End If ' ONLYXREOPIS
                          
                          
                          
                      
                          
                          
                          
                          
438                      .appendChild elemDETAILS
                         
                    
                    End With
440                 elemARTICLES.appendChild elemARTICLE

                   If ONLYXREOPIS = 1 Then  'B_C1
                       If Abs(SUMX - SUMP) > 0.005 Then
                           Gdb.Execute "UPDATE EGG SET B_C1=' -' WHERE ID=" + str(R!EGGID)
                       Else
                            Gdb.Execute "UPDATE EGG SET B_C1='*' WHERE ID=" + str(R!EGGID)
                       End If
                   Else
                   
                    If Abs(SUMX - SUMP) > 0.005 Then
                        Gdb.Execute "UPDATE TIM SET B_C1=' -' WHERE ID_NUM=" + str(R!id_num)
                    Else
                        Gdb.Execute "UPDATE TIM SET B_C1='*' WHERE ID_NUM=" + str(R!id_num)
                    End If
                   End If
                   '

DoEvents
                    If CMDSTOP.Enabled = False Then
                       MsgBox "ΔΙΑΚΟΠΗ ΔΙΑΔΙΑΣΙΑΣ"
                       CMDSTOP.Enabled = True
                       Exit Sub
                    End If
                    

442                 R.MoveNext
                    SUMP = 0: SUMX = 0
                    iin = iin + 1
                 Loop
                    
                    
                    
                    
444             .appendChild elemARTICLES
                  
              'C:\Program Files\Power Dominus Programs\MenuNET65
              
              
              
              
              
              
              
              
              
              
              
              
              
              
              
              
              
                    
446            R.Close


If List1.ListCount > 0 Then
   Open "C:\MERCVB\LOGGEFYR.TXT" For Output As #13
        Dim KLI77 As Integer
        For KLI77 = 0 To List1.ListCount - 1
            Print #13, List1.List(KLI77)
        Next



    Close #13

End If




 MsgBox "OK  Αρθρα " + str(iin)
 'Exit Sub
               
                Dim iic As Integer
                iic = 0
448              Set elemCUSTOMERS = docStock.createElement("CUSTOMERS")
                     
450               R.Open SQLPEL, Gdb, adOpenDynamic, adLockOptimistic
452                Do While Not R.EOF
454                 Set elemCUSTOMER = docStock.createElement("CUSTOMER")
456                  With elemCUSTOMER
458                     Set elemField = docStock.createElement("ID"): elemField.Text = str(R!ID): .appendChild elemField
460                     Set elemField = docStock.createElement("NAME"): elemField.Text = (R!EPO): .appendChild elemField
462                     Set elemField = docStock.createElement("VAT"): elemField.Text = (R!CAFM): .appendChild elemField
                     
                     
464                     Set elemField = docStock.createElement("JOB"): elemField.Text = (R!CEPA): .appendChild elemField
466                     Set elemField = docStock.createElement("DOY"): elemField.Text = (R!Cdoy): .appendChild elemField
468                     Set elemField = docStock.createElement("CUSTVAT"): elemField.Text = "2": .appendChild elemField
                    
                    
470                     Set elemField = docStock.createElement("ADDRESS"): elemField.Text = R!cDIE: .appendChild elemField
472                     Set elemField = docStock.createElement("ZIP"): elemField.Text = R!tk: .appendChild elemField
474                     Set elemField = docStock.createElement("CITY"): elemField.Text = R!cpOL: .appendChild elemField
                    
476                     Set elemField = docStock.createElement("PHONE1"): elemField.Text = R!cTHL: .appendChild elemField
478                     Set elemField = docStock.createElement("PHONE2"): elemField.Text = " ": .appendChild elemField
480                     Set elemField = docStock.createElement("PHONE3"): elemField.Text = " ": .appendChild elemField
                    
                     
482                     Set elemField = docStock.createElement("FAX1"): elemField.Text = " ": .appendChild elemField
484                     Set elemField = docStock.createElement("FAX2"): elemField.Text = " ": .appendChild elemField
486                     Set elemField = docStock.createElement("EMAIL"): elemField.Text = " ": .appendChild elemField
                   
488                     Set elemField = docStock.createElement("ISKEPYO"): elemField.Text = "0": .appendChild elemField
490                     Set elemField = docStock.createElement("ISDHMOSIOU"): elemField.Text = "0": .appendChild elemField
492                     Set elemField = docStock.createElement("ISEA"): elemField.Text = "0": .appendChild elemField
                     
494                     Set elemField = docStock.createElement("ISAGRYP"): elemField.Text = "0": .appendChild elemField
496                     Set elemField = docStock.createElement("ACCADDR"): elemField.Text = " ": .appendChild elemField
498                     Set elemField = docStock.createElement("STRADDR"): elemField.Text = " ": .appendChild elemField
                     
500                     Set elemField = docStock.createElement("STRNAME"): elemField.Text = " ": .appendChild elemField
502                     Set elemField = docStock.createElement("BANK1"): elemField.Text = " ": .appendChild elemField
                         
504                     Set elemField = docStock.createElement("BANKACC1"): elemField.Text = " ": .appendChild elemField
506                     Set elemField = docStock.createElement("BANK2"): elemField.Text = " ": .appendChild elemField
508                     Set elemField = docStock.createElement("BANKACC2"): elemField.Text = " ": .appendChild elemField
                         
510                     Set elemField = docStock.createElement("EACOUNTRY"): elemField.Text = " ": .appendChild elemField
512                     Set elemField = docStock.createElement("EAPREFIX"): elemField.Text = " ": .appendChild elemField
514                     Set elemField = docStock.createElement("EAVAT"): elemField.Text = " ": .appendChild elemField
 
 
516                     Set elemField = docStock.createElement("IDTYPE"): elemField.Text = " ": .appendChild elemField
 
                     
                     
                     End With
                    
518                 elemCUSTOMERS.appendChild elemCUSTOMER
520                 R.MoveNext
                    DoEvents
                    iic = iic + 1
                    Me.Caption = "Συναλλασόμενοι :" + str(iic)
                    
                    
                    
                     If CMDSTOP.Enabled = False Then
                       MsgBox "ΔΙΑΚΟΠΗ ΔΙΑΔΙΑΣΙΑΣ"
                       CMDSTOP.Enabled = True
                       Exit Sub
                    End If
                    
                    
                    
                    
                   Loop
                  
522             .appendChild elemCUSTOMERS
                 
                     MsgBox "OK  Αρθρα " + str(iin) + " Συναλλασόμενοι:" + str(iic)
                     
    '
    '
    '1116             For i = 0 To UBound(varStock)
    '1118                 Set elemARTICLES = docStock.createElement("CUSTOMERS")
    '1120                 With elemARTICLES
    '1122                     Set elemField = docStock.createElement("ProductCode")
    '1124                     elemField.Text = CStr(varStock(i)(0))
    '1126                     .appendChild elemField
    '1128                     Set elemField = docStock.createElement("ProductPrice")
    '1130                     elemField.Text = CStr(varStock(i)(1))
    '1132                     .appendChild elemField
    '                    End With
    '1134                 .appendChild elemARTICLES
    '                Next
            End With
524         Set .documentElement = elemDATA
            On Error Resume Next
526         Kill "C:\created.xml"
            On Error GoTo cmdCommand5_Click_Err
            '.save "C:\created.xml"


528         Par7MyData.FormatXmlDocument docStock

           ' cd1.FILENAME = "GEF_" + Format(Now, "yyyy-MM-dd HH-mm") + ".xml"
           ' cd1.ShowSave
530         CD1.FILENAME = "GEF_" + Format(Now, "yyyy-MM-dd HH-mm") + ".xml"
532         CD1.ShowSave
            
        ' .save cd1.FILENAME '"C:\created.xml"
534      .save CD1.FILENAME '"C:\created.xml"





        End With
    'Gdb.CommitTrans
    
536 MsgBox "ok"
    
    
    
        '<EhFooter>
        Exit Sub

cmdCommand5_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.Par7MyData.cmdCommand5_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub


Function FIND_fpa(ATIM As String, L As Integer)
Dim R As New ADODB.Recordset
R.Open "SELECT ISNULL(LFPA" + Format(L, "0") + ",'00')  FROM PARASTAT WHERE EIDOS='" + Left$(ATIM, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
FIND_fpa = Trim(R(0))
R.Close
If FIND_fpa = "00" Or Len(FIND_fpa) = 0 Then
   List1.AddItem "(101)ΔΕΝ ΒΡΕΘΗΚΕ ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ KATHΓ.ΦΠΑ" + str(L) + " ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + ATIM + "  "
End If
End Function

Function FIND_apo_Parastatiko_PELKOD(ATIM As String, L As Integer, ByVal KODIKOSPELATH As String)
' ΒΡΙΣΚΩ ΤΟΝ ΚΩΔΙΚΟ ΛΟΓΙΣΤΙΚΗΣ apo τον τυπο του παραστατικου
' αν εχει αστερακια βαζω τον κωδικο του πελατη
Dim R As New ADODB.Recordset
R.Open "SELECT ISNULL(XREPEL,'00')  FROM PARASTAT WHERE EIDOS='" + Left$(ATIM, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
If R.EOF() Then
  FIND_apo_Parastatiko_PELKOD = "00"
Else
  FIND_apo_Parastatiko_PELKOD = Trim(R(0))
  If F_PEL_KOD = 1 Then
     FIND_apo_Parastatiko_PELKOD = Replace$(FIND_apo_Parastatiko_PELKOD, "****", KODIKOSPELATH)
  End If
End If

R.Close
If FIND_apo_Parastatiko_PELKOD = "00" Or Len(FIND_apo_Parastatiko_PELKOD) = 0 Then
   List1.AddItem "(102)ΔΕΝ ΒΡΕΘΗΚΕ ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ ΠΕΛ/ΠΡΟΜ  ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + ATIM + "  "
End If


End Function






Function FIND_KOD(ATIM As String, L As Integer)
' ΒΡΙΣΚΩ ΤΟΝ ΚΩΔΙΚΟ ΛΟΓΙΣΤΙΚΗΣ
Dim R As New ADODB.Recordset
R.Open "SELECT ISNULL(LAJ" + Format(L, "0") + ",'00')  FROM PARASTAT WHERE EIDOS='" + Left$(ATIM, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
FIND_KOD = Trim(R(0))
R.Close
If FIND_KOD = "00" Or Len(FIND_KOD) = 0 Then
   List1.AddItem "(103)ΔΕΝ ΒΡΕΘΗΚΕ ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ KATHΓ.ΦΠΑ" + str(L) + " ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + ATIM + "  "
End If
End Function


Function FINDkat_PEL(ByVal EIDOS As String, ByVal kod As String, ByVal CATEGPEL As String)
' ΒΡΙΣΚΩ ΤΟΝ ΚΩΔΙΚΟ ΛΟΓΙΣΤΙΚΗΣ ΑΠΟ ΤΗΝ ΚΑΤΗΓΟΡΙΑ ΠΕΛΑΤΗ (KODGAL)  17/3/23
Dim c As String, R As New ADODB.Recordset, PROSU As String
PROSU = GGET_CVALUE("SELECT C2 FROM PINAKES WHERE TYPOS=13 AND AYJON=" + CATEGPEL)
'ΑΝΤΙΚΑΘΙΣΤΩ ΤΑ ΑΣΤΕΡΑΚΙΑ ** ΜΕ ΤΟΝ ΚΩΔΙΚΟ ΤΗΣ ΛΟΓΙΣΤΙΚΗΣ
c = Replace$(Trim(PROSU), "****", Right(Trim$(kod), 4))
If c = "00" Or Len(c) = 0 Then
   List1.AddItem "(104)ΔΕΝ ΒΡΕΘΗΚΕ ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ ΣTHN KATHΓΟΡΙΑ ΠΕΛΑΤΗ ΚΩΔ.ΠΕΛ/ΠΡΟΜ=" + kod
End If
FINDkat_PEL = c
End Function













Function FINDkat_KOD(ATIM As String, L As Integer, CATEG As Integer)
' ΒΡΙΣΚΩ ΤΟΝ ΚΩΔΙΚΟ ΛΟΓΙΣΤΙΚΗΣ
Dim c As String
Dim R As New ADODB.Recordset
R.Open "SELECT ISNULL(LAJ" + Format(L, "0") + ",'00')  FROM PARASTAT WHERE EIDOS='" + Left$(ATIM, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
c = Trim(R(0))
R.Close

Dim PROSU As String
PROSU = GGET_CVALUE("SELECT C2 FROM PINAKES WHERE TYPOS=15 AND AYJON=" + str(CATEG))

'ΑΝΤΙΚΑΘΙΣΤΩ ΤΑ ΑΣΤΕΡΑΚΙΑ ** ΜΕ ΤΟΝ ΚΩΔΙΚΟ ΤΗΣ ΛΟΓΙΣΤΙΚΗΣ
c = Replace$(c, "**", Trim$(PROSU))


If c = "00" Or Len(c) = 0 Then
   List1.AddItem "(105)ΔΕΝ ΒΡΕΘΗΚΕ ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ KATHΓ.ΦΠΑ" + str(L) + " ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + ATIM + "  "
End If

FINDkat_KOD = c

End Function


Function FINDtrp_KOD(trp As Integer)
' ΒΡΙΣΚΩ ΤΟΝ ΚΩΔΙΚΟ ΛΟΓΙΣΤΙΚΗΣ
Dim c As String


Dim PROSU As String
PROSU = GGET_CVALUE("SELECT C2 FROM PINAKES WHERE TYPOS=12 AND AYJON=" + str(trp))

'ΑΝΤΙΚΑΘΙΣΤΩ ΤΑ ΑΣΤΕΡΑΚΙΑ ** ΜΕ ΤΟΝ ΚΩΔΙΚΟ ΤΗΣ ΛΟΓΙΣΤΙΚΗΣ
' C = Replace$(C, "**", Trim$(PROSU))


c = Trim(PROSU)
If c = "00" Or Len(c) = 0 Then

    ' ΑΝ ΕΙΝΑΙ ΠΙΣΤΩΣΗ ΜΗΝ ΒΓΑΖΕΙΣ ΜΗΝΥΜΑ (MYDATA(N1)=5
   If 5 = GGET_NVALUE("SELECT N1 FROM PINAKES WHERE TYPOS=12 AND AYJON=" + str(trp)) Then
   Else
      List1.AddItem "(106)ΔΕΝ ΒΡΕΘΗΚΕ ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ KATHΓ.ΦΠΑ" + str(L) + " ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + ATIM + "  "
   End If
End If

FINDtrp_KOD = c

End Function























Function FINDFPAkat_KOD(ATIM As String, L As Integer, CATEG As Integer)
' ΒΡΙΣΚΩ ΤΟΝ ΚΩΔΙΚΟ ΛΟΓΙΣΤΙΚΗΣ
Dim c As String
Dim R As New ADODB.Recordset
R.Open "SELECT ISNULL(LFPA" + Format(L, "0") + ",'00')  FROM PARASTAT WHERE EIDOS='" + Left$(ATIM, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
c = Trim(R(0))
R.Close

Dim PROSU As String
PROSU = GGET_CVALUE("SELECT C2 FROM PINAKES WHERE TYPOS=15 AND AYJON=" + str(CATEG))

'ΑΝΤΙΚΑΘΙΣΤΩ ΤΑ ΑΣΤΕΡΑΚΙΑ ** ΜΕ ΤΟΝ ΚΩΔΙΚΟ ΤΗΣ ΛΟΓΙΣΤΙΚΗΣ
c = Replace$(c, "**", Trim$(PROSU))


If c = "00" Or Len(c) = 0 Then
   List1.AddItem "(107)ΔΕΝ ΒΡΕΘΗΚΕ ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ KATHΓ.ΦΠΑ" + str(L) + " ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + ATIM + "  "
End If

FINDFPAkat_KOD = c

End Function

















Function cVAL(ByVal Z As Single) As String  ' STR ME KOMA , !!!!!!!!!!!!!!!!!!!!!!!!!
    Dim X As String
    X = Format(Z, "########0.00")
    cVAL = Replace(X, ".", ",")
    
   
End Function



Private Sub cmdΓΚατηγορια_Click()
'-------------------------------------------ΕΙΣΠΡΑΞΕΙς -------------------------------------------
Dim R As New ADODB.Recordset

         Dim synt As String
            If chkEPAN.Value = vbChecked Then
                synt = ""

            Else
                synt = " and (B_C1 is null or LEFT(B_C1,1)<>'*') "

            End If
            ' ExecuteSQLQuery("update TIM SET AJ7=0 WHERE AJ7 IS NULL")


        Dim pol As String: pol = " "
        Dim polepis As String: polepis = " "
        Dim ago As String: ago = " "
        Dim AGOEPIS As String: AGOEPIS = " "
        
        Dim sql As String, sqlwh As String
            sql = "SELECT ISNULL(XREOPIS.KODLOG,'00-00') AS KODLOG,EGG.ID AS EGGID,ISNULL(EIDXPI,0) AS EIDXPI,'1.' AS TRP,EGG.EIDOS AS EIDOS,PEL.ID AS PELID,0 AS ID_NUM,0 AS AJ1  , 0 AS AJ2 ,0 AS AJ3,0 AS AJ4,0 AS AJ5,0 AS AJI,"
            sql = sql + "0 AS FPA1,0 AS FPA2,0 AS FPA3,0 AS FPA4,0 AS FPA5,ISNULL(ATIM,' ') AS ATIM,"
            sql = sql + "HME,PEL.EPO,PEL.AFM,EGG.KOD AS KPE,PEL.DIE,PEL.XRVMA"    '"CONVERT(CHAR(10),HME,3) AS HMEP
            sql = sql + ",PEL.EPA,PEL.POL,0 AS AJ6,0 AS FPA6,0 AS AJ7,0 AS FPA7,EGG.XREOSI+EGG.PISTOSI AS SKAU,0 AS SFPA ,"
            sql = sql + " (CASE WHEN XREOSI>0 THEN 1 ELSE -1 END ) AS MSIGN,ISNULL(KODGAL,'1') AS KODGAL "
            sqlwh = "   FROM EGG INNER JOIN PEL ON EGG.EIDOS=PEL.EIDOS AND EGG.KOD=PEL.KOD  LEFT JOIN  XREOPIS  ON EGG.EIDXPI=XREOPIS.EIDOS "
          '  sqlwh = sqlwh + " WHERE  (IDTIM IS NULL)    and HME>='" + Format(apo, "mm/dd/yyyy") + "'  AND HME<='" + Format(eos, "mm/dd/yyyy") + "'  "
            sqlwh = sqlwh + " WHERE  (IDTIM IS NULL) and EGG.KOD<>'9990'   and HME>='" + Format(APO, "mm/dd/yyyy") + "'  AND HME<='" + Format(EOS, "mm/dd/yyyy") + "'  "
            sqlwh = sqlwh + "  AND ABS(EGG.XREOSI)+ABS(EGG.PISTOSI)>0  " + synt
            
            
            
            SQLPEL = "SELECT ID,EPO,ISNULL(AFM,'000000000') AS CAFM,ISNULL(EPA,'') AS CEPA,ISNULL(DOY,'') AS CDOY,ISNULL(DIE,'') AS CDIE,ISNULL(XRVMA,'') AS TK," _
                 & "ISNULL(POL,'') AS CPOL,ISNULL(THL,'') AS CTHL " _
                 & " FROM PEL WHERE ID IN ( SELECT DISTINCT PEL.ID " + sqlwh + ")"
            
            
            
    
    Gdb.BeginTrans
    GmakeXML sql, sqlwh, SQLPEL, 0, 1
    Gdb.CommitTrans
End Sub

Private Sub Command1_Click()
    Gdb.Execute "DROP TABLE SQLTEMP"
    
End Sub

Private Sub Form_DblClick()
 '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "BOHU10"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU10.Form_DblClick " & "at line " & Erl

        Resume Next






End Sub

Private Sub Form_Load()
f_categ = Val(FINDPARAMETROI(1, "BOHU10", "F_CATEG", "0", "1=Γ ΚΑΤΗΓ ΜΕ ΚΑΤΗΓΟΡΙΑ ΕΙΔΟΥΣ 0=ΧΩΡΙΣ ΚΑΤΗΓΟΡΙΑ"))
F_PEL_KOD = Val(FINDPARAMETROI(1, "BOHU10", "F_PEL_KOD", "0", "1=ME **** ΒΑΖΕΙ ΤΟΝ ΚΩΔΙΚΟ ΠΕΛΑΤΗ ΣΤΟΝ 30/50 0=ΟΧΙ"))
f_kod_tameio = FINDPARAMETROI(1, "BOHU10", "f_kod_tameio", "38-00-00-0000", "ΚΩΔΙΚΟΣ ΤΑΜΕΙΟ ΧΡΕΟΠΙΣΤΩΣΕΩΝ")
'apo.Value = Now()
'eos.Value = Now()
APO.Value = Now()
EOS.Value = Now()





End Sub

'detail απο Γ κατηγ
'
'    <DETAIL>
'                    <LCODE>64.00.00.0087</LCODE>
'                    <CRDB>0</CRDB>
'                    <AMOUNT>1000</AMOUNT>
'                    <INVOICE>ΤΔΑ 12</INVOICE>
'                    <REASON>ΠΑΠΑΔΟΠΟΥΛΟΣ</REASON>
'                    <ISAGRYP>0</ISAGRYP>
'                    <KEPYOPARTY>100</KEPYOPARTY>
'                </DETAIL>




'Πεδιό    Περιγραφή            Υποχρεωτικό   Έγκυρες Τιμές(Μέγεθος)
'MTYPE     Είδος Εγγραφής         Ναι         0 έσοδο, 1 έξοδο, 2 χρέωση,3 πίστωση
'MSIGN   Αντιστροφή Προσήμου     Default 1    -1 ναι, 1 όχι
'ISKEPYO       ΚΕΠΥΟ             Default 0    0 όχι, 1 ναι
'ISAGRYP Υποχρέωση Υποβολής                   0 Υπόχρεος, 1 ΜηΥπόχρεος
'LCODE Λογαριασμός                 Ναι        Αλφαριθμητικό (20)
'CUSTID ID Συναλλασσομένου                    Ακέραιος αριθμός
'INVOICE Παραστατικό                          Αλφαριθμητικό (20)
'MDATE Ημερομηνία                 Ναι         ηη/μμ/εεεε
'REASON Αιτιολογία                            Αλφαριθμητικό (60)
'CASHAMT Μετρητά                               Ποσό
'LCCASH Λογαριασμός Μετρητών                  Αλφαριθμητικό (20)
'CHEQUEAMT Αξιόγραφα Ποσό
'LCCHEQUE Λογαριασμός Αξιογράφων               Αλφαριθμητικό (20)
'TAXAMT1 Αξία Φόρου 1 Ποσό
'TAXAMT2 Αξία Φόρου 2 Ποσό
'TAXAMT3 Αξία Φόρου 3 Ποσό
'TAXAMT4 Αξία Φόρου 4 Ποσό
'TAXAMT5 Αξία Φόρου 5 Ποσό
'LCTAX1 Λογαριασμός Φόρου 1 Αλφαριθμητικό (20)
'LCTAX2 Λογαριασμός Φόρου 2 Αλφαριθμητικό (20)
'LCTAX3 Λογαριασμός Φόρου 3 Αλφαριθμητικό (20)
'LCTAX4 Λογαριασμός Φόρου 4 Αλφαριθμητικό (20)
'LCTAX5 Λογαριασμός Φόρου 5 Αλφαριθμητικό (20)
'KEPYOAMT Αξία ΚΕΠΥΟ Ποσό
'ISBUILD Εγγραφή βιβλίου κοστ. Οικοδ.                  0 όχι, 1 ναι
'INBRCODE Κωδ. Εγκατάστασης Αλφαριθμητικό (20)
'SUMKEPYOYP Αξία ΚΕΠΥΟ Υπόχρεου Ποσό
'SUMKEPYONOTYP Αξία ΚΕΠΥΟ ΜηΥπόχρεου                   Ποσό
'SUMKEPYOFPA Αξία ΦΠΑ ΚΕΠΥΟΥπόχρεου                    Ποσό
'OTHEREXPEND Λοιπές Δαπάνες                            0 όχι, 1 ναι
'CASHREGISTERID Κωδικός ΤαμειακήςΜηχανής               Αλφαριθμητικό(40)
'HASRETAILID Αν έχει συνδεδεμένηΤαυτότητα Λιανικών     0 όχι, 1 ναι
'CANCELGROUPID Πεδίο ομαδοποίησης
'σχέσης ακυρωμένου -ακυρωτικού                         Αλφαριθμητικό (40)
'CANCELED Αν είναι ακυρωμένο ή Ακυρωτικό               0 Ακυρωτικό, 1 Ακυρωμένο

'Τα πεδία του <DETAIL> αναλύονται παρακάτω:
'Πεδιό               Περιγραφή                         Έγκυρες Τιμές (Μέγεθος)
'LCODE            Λογαριασμός                          Αλφαριθμητικό (20)
'ΝΕΤΑΜΤ           Καθαρή αξία                          Ποσό
'VATAMT            Αξία Φ.Π.Α.                         Ποσό
'ISAGRYP Υποχρέωση Υποβολής                             0 Υπόχρεoς, 1 Μη Υπόχρεoς
'KEPYOPARTY        ΚΕΠΥΟ                               Αριθμητικό


'Μέσα στο <DATA> περιέχεται επίσης και το node <CUSTOMERS> που περιέχει με τη
'σειρά του πολλά node <CUSTOMER> ένα για κάθε συναλλασσόμενο.
'Τα πεδία του <CUSTOMER> είναι τα παρακάτω:

'Πεδιό                    Περιγραφή            Υποχρεωτικό            Έγκυρες Τιμές (Μέγεθος)
'ID          ID Συναλλασσομένου             Ναι                   Ακέραιος αριθμός
'NAME        Επωνυμία                       Ναι                   Αλφαριθμητικό (60)
'VAT         Α.Φ.Μ.                         Ναι                   Έγκυρο ΑΦΜ
'JOB         Επάγγελμα                                            Αλφαριθμητικό (40)6
'DOYCODE     Κωδικός ΔΟΥ                    Ναι                   Έγκυρος Κωδικός ΔΟΥ
'CUSTVAT     Καθεστώς ΦΠΑ                                          1 κανονικό, 2 τεκμαρτό,
                                                                 ' 3 Πρακ.Ταξ., 4 Καπν/κών,
                                                                 ' 5 Αθρ. 45, 6 Αθρ. 46,
                                                                 ' 7 Απαλλασσόμενη, 8 Αγροτών
'ADDRESS     Διεύθυνση                                           Αλφαριθμητικό (60)
'ZIP         ΤΚ                                                  Αλφαριθμητικό (10)
'CITY        Πόλη Αλφαριθμητικό (40)
'PHONE1      Τηλέφωνο 1 Αλφαριθμητικό (40)
'PHONE2      Τηλέφωνο 2 Αλφαριθμητικό (40)
'PHONE3      Τηλέφωνο 3 Αλφαριθμητικό (40)
'FAX1        Fax 1 Αλφαριθμητικό (40)
'FAX2        Fax 2 Αλφαριθμητικό (40)
'EMAIL       Ηλεκτρονική Διεύθυνση Αλφαριθμητικό (40)
'ISKEPYO     Υπόχρεος ΚΕΠΥΟ                 Default 1            0 όχι, 1 ναι
'ISDIMOSIOU  Συναλλασσόμενος Δημοσίου       Default 0            0 όχι, 1 ναι
'ISEA        Ενδοκοινοτικές Συναλλαγές      Default 0            0 όχι, 1 ναι
'ISAGRYP     Υποχρέωση Υποβολής             Default 0            0 Υπόχρεος, 1 Μη Υπόχρεος,2 Διπλή Ιδιότητα
'ACCADDR     Διεύθυνση Λογιστηρίου Αλφαριθμητικό (60)
'STRADDR     Διεύθυνση Αποθήκης Αλφαριθμητικό (60)
'STRNAME     Υπεύθυνος Αποθήκης Αλφαριθμητικό (60)
'BANK1       Τράπεζα 1 Αλφαριθμητικό (60)
'BANKACC1    Λογαριασμός Τραπέζης 1 Αλφαριθμητικό (60)
'BANK2       Τράπεζα 2 Αλφαριθμητικό (60)
'BANKACC2    Λογαριασμός Τραπέζης 2 Αλφαριθμητικό (60)
'EACOUNTRY   Χώρα ΕνδοκοινοτικώνΣυναλλαγών Αλφαριθμητικό(40)
'EAPREFIX    Πρόθεμα Χώρα Ενδ. Συναλ. Αλφαριθμητικό (2)
'IDTYPE      Είδος Ταυτότητας   Αλφαριθμητικό (10) Έγκυρες Τιμές:'ΑΤ, ΕΣ, ΠΝ, ΠΑ, ΕΑ, ΛΙ, ΠΣ,'ΔΙ, ΚΑ, ΧΤ, ΤΟ, ΑΚΑΗ, ΕΔΤ,'ΕΙΔ , ΑΠΠΑ, ΕΙΔΑ
'IDNUMB      Αριθμός Ταυτότητας Αλφαριθμητικό (20)
'EAVAT       Αρ. Μητρώου ΦΠΑ Αλφαριθμητικό (12)

Private Sub lblΑπο_Click()

End Sub
