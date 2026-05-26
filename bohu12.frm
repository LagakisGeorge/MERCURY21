VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form bohu12 
   BackColor       =   &H00FF0000&
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   7695
      Left            =   9000
      TabIndex        =   19
      Top             =   3120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   13573
      _Version        =   393216
      Rows            =   56
      Cols            =   4
      WordWrap        =   -1  'True
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   12000
      TabIndex        =   18
      Text            =   "ετος"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   12000
      TabIndex        =   17
      Text            =   "μηνας"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   12000
      TabIndex        =   16
      Text            =   "αφμ"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton XMLAGO 
      Caption         =   "ΑΓΟΡΕΣ XML"
      Height          =   360
      Left            =   11400
      TabIndex        =   15
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton XMLPOL 
      Caption         =   "ΠΩΛΗΣΕΙΣ ΣΕ XML"
      Height          =   360
      Left            =   11400
      TabIndex        =   14
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdAGORES 
      Caption         =   "ΑΓΟΡΕΣ ΣΕ EXCEL"
      Height          =   360
      Left            =   9000
      TabIndex        =   13
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdEXCEL 
      Caption         =   "ΠΩΛΗΣΕΙΣ ΣΕ EXCEL"
      Height          =   360
      Left            =   9000
      TabIndex        =   12
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton create2014 
      Caption         =   "Δημιουργία αρχείου 2014 και μετά"
      Height          =   735
      Left            =   9000
      TabIndex        =   11
      Top             =   1080
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker D1 
      Height          =   330
      Left            =   8340
      TabIndex        =   7
      Top             =   120
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   271515649
      CurrentDate     =   41776
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Εξοδος"
      Height          =   315
      Left            =   11160
      TabIndex        =   5
      Top             =   120
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Επικεφαλίδα Αρχείου"
      Height          =   315
      Left            =   165
      TabIndex        =   4
      Top             =   2070
      Width           =   2130
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   420
      Left            =   5175
      Top             =   8610
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   741
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
      Caption         =   "Adodc1"
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
      Bindings        =   "bohu12.frx":0000
      Height          =   3975
      Left            =   0
      TabIndex        =   3
      Top             =   3165
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
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
      Caption         =   "ΧΡΕΩΣΤΙΚΑ"
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
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   3345
      MultiSelect     =   1  'Simple
      TabIndex        =   2
      Top             =   135
      Width           =   3432
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Δημιουργία Αρχείου"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   2220
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   225
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   135
      Width           =   3075
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "bohu12.frx":0015
      Height          =   3495
      Left            =   0
      TabIndex        =   6
      Top             =   7320
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "ΠΙΣΤΩΤΙΚΑ"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   420
      Left            =   2880
      Top             =   2400
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   741
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
      Caption         =   "Adodc2"
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
   Begin MSComCtl2.DTPicker d2 
      Height          =   330
      Left            =   8355
      TabIndex        =   9
      Top             =   720
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   122355713
      CurrentDate     =   41776
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ΕΤΟΣ"
      Height          =   195
      Left            =   11160
      TabIndex        =   22
      Top             =   1680
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ΜΗΝΑΣ"
      Height          =   195
      Left            =   11160
      TabIndex        =   21
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ΑΦΜ"
      Height          =   195
      Left            =   11160
      TabIndex        =   20
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Εως"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   7500
      TabIndex        =   10
      Top             =   735
      Width           =   675
   End
   Begin VB.Label apo 
      BackStyle       =   0  'Transparent
      Caption         =   "Από"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   7485
      TabIndex        =   8
      Top             =   135
      Width           =   675
   End
End
Attribute VB_Name = "bohu12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)


Dim f_AFM         '= Val(FindParametroi(1,"PELAT2", "f_KENTRA_ADYNATISMATOS", "0", "Αν είναι κέντρο αδυνατίσματος=1"))

Dim f_tam

Private Sub cmdAGORES_Click()

    Dim N

    N = MsgBox("Σύνολο ανά κετηγορία;", vbYesNo)

    If N = vbYes Then N = 1 Else N = 0
    print3_xar "SELECT * FROM SYNOLAKEPYO WHERE POLHS=2  ", "000000111", "ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΑΠΟ " + Format(d1, "DD/MM/YYYY") + " ΕΩΣ " + Format(d2, "DD/MM/YYYY"), N
  
End Sub

Private Sub cmdEXCEL_Click()

    Dim N

    N = MsgBox("Σύνολο ανά κετηγορία;", vbYesNo)

    If N = vbYes Then N = 1 Else N = 0
    ' gia excel υπουργειου
    
    print3_xar "SELECT POLHS,AFM,REPLACE(LTRIM(convert(decimal(10,2),AJIA)),'.',',') as [ΑΞΙΑ],REPLACE(LTRIM(convert(decimal(10,2),FPA)),'.',',') AS [ΦΠΑ] ,LTRIM(convert(decimal(10,0),TEM)) AS  [TIM],PIS,SHME AS [ΤΕΛ.ΠΑΡ], EPO,DOY,EPA,DIE  FROM SYNOLAKEPYO WHERE POLHS=1 OR POLHS=3 OR POLHS=4 ", "00011100000000", "ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΑΠΟ " + Format(d1, "DD/MM/YYYY") + " ΕΩΣ " + Format(d2, "DD/MM/YYYY"), N
    
    'gia excel kanoniko (bgazei arithmoys
   ' print3_xar "SELECT POLHS,AFM,AJIA as [ΚΑΘ.ΑΞΙΑ],FPA AS [ΦΠΑ],TEM AS [TMX],PIS,SHME AS [ΤΕΛ.ΠΑΡ], EPO,DOY,EPA,DIE  FROM SYNOLAKEPYO WHERE POLHS=1 OR POLHS=3 OR POLHS=4 ", "00111000000000", "ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΑΠΟ " + Format(d1, "DD/MM/YYYY") + " ΕΩΣ " + Format(d2, "DD/MM/YYYY"), n
    
End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim R      As New ADODB.Recordset

        Dim RH     As New ADODB.Recordset

        Dim rPEL   As New ADODB.Recordset

        Dim SQLPEL As String

        Dim k      As Integer, sql As String, pol As String, ag As String

100     pol = ""

110     For k = 0 To List1.ListCount - 1

120         If List1.Selected(k) Then
130             pol = pol + "'" + Right(List1.List(k), 1) + "',"
            End If

        Next

140     ag = ""

150     For k = 0 To List2.ListCount - 1

160         If List2.Selected(k) Then
170             ag = ag + "'" + Right(List2.List(k), 1) + "',"
            End If

        Next

180     sql = "select (TIM.AJ1+TIM.AJ2+TIM.AJ3+TIM.AJ4+TIM.AJ5) as AJIA,LEFT(ATIM,1) AS TYPOS," & "1 AS AEG,PEL.AFM,PEL.EPO,PEL.DOY,PEL.EPA,PEL.DIE,PEL.POL,PLAISIO AS TK,TIM.EIDOS " & " INTO TIMKEPYO FROM TIM  INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD " & " WHERE TIM.AJ1+TIM.AJ2+TIM.AJ3+TIM.AJ4+TIM.AJ5>=300 AND LEFT(ATIM,1) IN (" + pol + ag + "'@')  "
190     sql = sql + " and HME>='" + Format(d1, "MM/DD/YYYY") + "' AND HME<='" + Format(d2, "MM/DD/YYYY") + "'"

        On Error Resume Next

200     Gdb.Execute "DROP TABLE TIMKEPYO"
210     Gdb.Execute "DROP TABLE TIMKEPYO2"
220     Gdb.Execute "DROP TABLE PEL22"

        Dim nn

230     Gdb.Execute sql, nn

240     If nn = 0 Then
250         MsgBox "ΔΕΝ ΔΗΜΙΟΥΡΓΗΘΗΚΕ ΤΟ ΑΡΧΕΙΟ"
260         MsgBox Err.Description

            Exit Sub

        End If

        Dim polhs As String, polepis As String, ago As String, AGOEPIS As String

        Dim a, AA As Long

270     a = Get_AJ(polhs, polepis, ago, AGOEPIS)

        On Error GoTo Command1_Click_Err

        Dim L As Long

280     Gdb.Execute "UPDATE TIMKEPYO SET AJIA=-AJIA WHERE TYPOS IN (" + polepis + "," + AGOEPIS + ")", L

        'SOYMAREI ANA AFM TIMOLOGIA KAI AXIES STO TIMKEPYO2
290     sql = "select SUM(AJIA) AS SAJIA , SUM(AEG) AS ARITMOS,AFM,EIDOS into TIMKEPYO2 " & "FROM TIMKEPYO GROUP BY AFM,EIDOS HAVING SUM(AEG)>0 "

300     Gdb.Execute sql
310     Gdb.Execute "DELETE FROM TIMKEPYO2 WHERE AFM='' OR AFM IS NULL", nn

320     R.Open "SELECT COUNT(*) FROM TIMKEPYO2 WHERE SAJIA<0 AND EIDOS='e' ", Gdb, adOpenDynamic, adLockOptimistic

330     MsgBox "υπαρχουν " + str(R(0)) + " πιστωτικα πελατων"

340     R.Close

350     R.Open "SELECT COUNT(*) FROM TIMKEPYO2 WHERE SAJIA<0 AND EIDOS='r' ", Gdb, adOpenDynamic, adLockOptimistic

360     MsgBox "υπαρχουν " + str(R(0)) + " πιστωτικα προμηθευτων"

370     R.Close

380     sql = "select SAJIA,ARITMOS,TIMKEPYO2.AFM,TIMKEPYO2.EIDOS " & "FROM TIMKEPYO2 " & " WHERE TIMKEPYO2.EIDOS='e' ORDER BY TIMKEPYO2.EIDOS"

        'Exit Sub    INNER JOIN PEL ON TIMKEPYO2.AFM=PEL.AFM

        '
        '
        '
        'Exit Sub

390     R.Open sql, Gdb, adOpenDynamic, adLockOptimistic

400     Adodc1.ConnectionString = gConnect
410     Adodc1.RecordSource = sql    '   "select * FROM TIMKEPYO ORDER BY TYPOS"
        'R.Open "seLect * from TIMKEPYO ", Gdb, adOpenDynamic, adLockOptimistic
420     Adodc1.Refresh

430     RH.Open "SELECT TOP 1 * FROM MKEPYO", Gdb, adOpenDynamic, adLockOptimistic

440     Open "C:\MERCVB\IL01" For Output As #1

450     Print #1, "0IL01    " + Right(Date, 8) + Space(133)
460     Print #1, "1" + RH!m_etos + RH!M_DOY2 + RH!AR_FAKEL + Left(RH!m_onoma, 18) + RH!M_AFM + RH!M_EPA + RH!M_POLH + RH!M_ODOS + RH!M_ARIT + RH!TAX_KOD + RH!M_THL + RH!M_FAX + RH!M_EMAIL + RH!m_ypok
470     AA = 0

        Dim sumTim As Long, sumAjia As Single

        Dim DUM

480     sumTim = 0: sumAjia = 0

'        On Error Resume Next
On Error GoTo Command1_Click_Err

490     Do While Not R.EOF

500         If R("EIDOS") = "e" Then
510             AA = AA + 1
520             rPEL.Open "select * from PEL WHERE EIDOS='e' and AFM='" + R!AFM + "'", Gdb, adOpenDynamic, adLockBatchOptimistic
530             pol = "2" + StrZero(AA, 5)
540             pol = pol + Left(IfNull(R!AFM) + Space(9), 9) + ""
550             pol = pol + Left(IfNull(rPEL!EPO) + Space(27), 27)
560             pol = pol + Left(IfNull(rPEL!EPA) + Space(15), 15)
570             pol = pol + Format(Val("0"), "0") + Left(IfNull(rPEL!pol) + Space(10), 10)
580             pol = pol + Left(IfNull(rPEL!DIE) + Space(16), 16) + Left(Right(IfNull(rPEL!DIE), 3) + Space(3), 3)
590             pol = pol + Left(IfNull(rPEL!plaisio) + Space(5), 5)
600             pol = pol + StrZero(R!ARITMOS, 7) + StrZero(R!SAJIA * 100, 16) + Space(35)
610             sumTim = sumTim + R!ARITMOS
620             sumAjia = sumAjia + R!SAJIA
630             DUM = check_afm(CNull(R!AFM))
640             rPEL.Close

650             If DUM = 0 Then
660                 MsgBox "Λάθος ΑΦΜ πελατης " + CNull(R!AFM)    '+ " " + R!EPO
                End If

670             Print #1, pol
            End If

680         R.MoveNext
        Loop

690     MsgBox "ΠΕΛΑΤΕΣ : " + Chr(13) + str(sumTim) + " TIM " + str(sumAjia) + " ΑΞΙΑ"

        Dim t_dhm    'ΤΙΜΟΛΟΓΙΑ ΔΗΜΟΣΙΟΥ

700     t_dhm = 0
710     Print #1, "3         " + StrZero(t_dhm * 100, 16) + "         " + StrZero(sumTim, 16) + StrZero(sumAjia * 100, 16) + Space(83)

720     AA = 0

730     R.Close

740     sql = "select SAJIA,ARITMOS,TIMKEPYO2.AFM,TIMKEPYO2.EIDOS " & "FROM TIMKEPYO2 " & " WHERE TIMKEPYO2.EIDOS='r' ORDER BY TIMKEPYO2.EIDOS"

750     R.Open sql, Gdb, adOpenDynamic, adLockOptimistic

760     sumTim = 0: sumAjia = 0
770     R.MoveFirst

780     Do While Not R.EOF

790         If R("EIDOS") = "r" Then
800             rPEL.Open "select * from PEL WHERE EIDOS='r' and AFM='" + R!AFM + "'", Gdb, adOpenDynamic, adLockBatchOptimistic
810             AA = AA + 1
820             pol = "4" + StrZero(AA, 5)
830             pol = pol + Left(IfNull(rPEL!AFM) + Space(9), 9) + ""
840             pol = pol + Left(IfNull(rPEL!EPO) + Space(27), 27)
850             pol = pol + Left(IfNull(rPEL!EPA) + Space(15), 15)
860             pol = pol + Format(Val("0"), "0") + Left(IfNull(rPEL!pol) + Space(10), 10)
870             pol = pol + Left(IfNull(rPEL!DIE) + Space(16), 16) + Left(Right(IfNull(rPEL!DIE), 3) + Space(3), 3)
                'pol = pol + Left(IfNull(rPEL!Tk) + Space(5), 5)
880             pol = pol + Left(IfNull(rPEL!plaisio) + Space(5), 5)
890             pol = pol + StrZero(R!ARITMOS, 7) + StrZero(R!SAJIA * 100, 16) + Space(35)

900             DUM = check_afm(CNull(rPEL!AFM))
910

920             If DUM = 0 Then
930                 MsgBox "Λάθος ΑΦΜ πελατης " + CNull(rPEL!AFM)    '+ " " + R!EPO
                End If

940             sumTim = sumTim + R!ARITMOS
950             sumAjia = sumAjia + R!SAJIA
960             Print #1, pol
                rPEL.Close
            End If

970         R.MoveNext
        Loop

        ' Print #1, "5...."
980     Print #1, "5         " + StrZero(t_dhm * 100, 16) + "000000000" + StrZero(sumTim, 16) + StrZero(sumAjia * 100, 16) + Space(83)
        '          '5999999011'+strzero(100*t_dhm,16)+'999999999'+strzero(s_tim,16)+strzero(100*s_aji,16)+space(83)
990     Close #1

        '-------------- IL01.PIS -----------------------------------------

1000    R.Close

1010    sql = "select SUM(AJIA) AS SAJIA , SUM(AEG) AS ARITMOS,AFM,EPO,DIE,DOY,POL,TK,EIDOS,EPA " & "FROM TIMKEPYO GROUP BY AFM,EPO,DIE,DOY,POL,TK,EIDOS,EPA HAVING SUM(AJIA)<0 ORDER BY EIDOS,EPO"

1020    R.Open sql, Gdb, adOpenDynamic, adLockOptimistic

1030    Adodc2.ConnectionString = gConnect
1040    Adodc2.RecordSource = sql    '   "select * FROM TIMKEPYO ORDER BY TYPOS"
        'R.Open "seLect * from TIMKEPYO ", Gdb, adOpenDynamic, adLockOptimistic
1050    Adodc2.Refresh

1060    RH.Close
1070    RH.Open "SELECT TOP 1 * FROM MKEPYO", Gdb, adOpenDynamic, adLockOptimistic

1080    Open "C:\IL01.PIS" For Output As #1

1090    Print #1, "0IL01.PIS" + Right(Date, 8) + Space(133)
1100    Print #1, "1" + RH!m_etos + RH!M_DOY2 + RH!AR_FAKEL + Left(RH!m_onoma, 18) + RH!M_AFM + RH!M_EPA + RH!M_POLH + RH!M_ODOS + RH!M_ARIT + RH!TAX_KOD + RH!M_THL + RH!M_FAX + RH!M_EMAIL + RH!m_ypok
1110    AA = 0

1120    sumTim = 0: sumAjia = 0

        'On Error GoTo LATOS
        On Error GoTo Command1_Click_Err


        'On Error GoTo Command1_Click_Err
1130    Do While Not R.EOF

1140        If R!EIDOS = "e" Then
1150            AA = AA + 1
1160            pol = "2" + StrZero(AA, 5)
1170            pol = pol + Left(IfNull(R!AFM) + Space(9), 9) + ""
1180            pol = pol + Left(IfNull(R!EPO) + Space(27), 27)
1190            pol = pol + Left(IfNull(R!EPA) + Space(15), 15)
1200            pol = pol + Format(Val("0"), "0") + Left(IfNull(R!pol) + Space(10), 10)
1210            pol = pol + Left(IfNull(R!DIE) + Space(16), 16) + Left(Right(IfNull(R!DIE), 3) + Space(3), 3)
1220            pol = pol + Left(IfNull(R!tk) + Space(5), 5)
1230            pol = pol + StrZero(R!ARITMOS, 7) + StrZero(-R!SAJIA * 100, 16) + Space(35)
1240            sumTim = sumTim + R!ARITMOS
1250            sumAjia = sumAjia + (-R!SAJIA)
1260            DUM = check_afm(CNull(R!AFM))

1270            If DUM = 0 Then
1280                MsgBox "Λάθος ΑΦΜ " + CNull(R!AFM) + " " + CNull(R!EPO)
                End If

1290            Print #1, pol
            End If

1300        R.MoveNext
        Loop

1310    t_dhm = 0
1320    Print #1, "3         " + StrZero(t_dhm * 100, 16) + "999999999" + StrZero(sumTim, 16) + StrZero(sumAjia * 100, 16) + Space(83)

1330    AA = 0
1340    sumTim = 0: sumAjia = 0

1350   If Not R.EOF Then
             R.MoveFirst
       End If
1360    Do While Not R.EOF

1370        If R("EIDOS") = "r" Then
1380            AA = AA + 1
1390            pol = "4" + StrZero(AA, 5)
1400            pol = pol + Left(IfNull(R!AFM) + Space(9), 9) + ""
1410            pol = pol + Left(IfNull(R!EPO) + Space(27), 27)
1420            pol = pol + Left(IfNull(R!EPA) + Space(15), 15)
1430            pol = pol + Format(Val("0"), "0") + Left(IfNull(R!pol) + Space(10), 10)
1440            pol = pol + Left(IfNull(R!DIE) + Space(16), 16) + Left(Right(IfNull(R!DIE), 3) + Space(3), 3)
1450            pol = pol + Left(IfNull(R!tk) + Space(5), 5)
1460            pol = pol + StrZero(R!ARITMOS, 7) + StrZero(-R!SAJIA * 100, 16) + Space(35)
1470            sumTim = sumTim + R!ARITMOS
1480            sumAjia = sumAjia + (-R!SAJIA)

1490            DUM = check_afm(CNull(R!AFM))

1500            If DUM = 0 Then
1510                MsgBox "Λάθος ΑΦΜ " + CNull(R!AFM) + " " + CNull(R!EPO)
                End If

1520            Print #1, pol
            End If

1530        R.MoveNext
        Loop

        ' Print #1, "5...."
1540    Print #1, "5         " + StrZero(t_dhm * 100, 16) + "000000000" + StrZero(sumTim, 16) + StrZero(sumAjia * 100, 16) + Space(83)
        '          '5999999011'+strzero(100*t_dhm,16)+'999999999'+strzero(s_tim,16)+strzero(100*s_aji,16)+space(83)
1550    Close #1

1560    MsgBox "ΔHMIOYΡΓHΘHKE TO AΡXEIO C:\MERCVB\IL01"

        Exit Sub

LATOS:

        On Error Resume Next

1570    Close #1
1580    MsgBox "ΔΕΝ ΔΗΜΙΟΥΡΓΗΘΗΚΕ"

1590    Resume Next

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu12.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu12.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim F As New DataEntry2

100     Load F
110     F.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM MKEPYO WHERE 1=1 "    '    sql
        'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
120     F.ParamGrid.TextMatrix(0, 2) = "SELECT * " & " FROM MKEPYO WHERE 1=1 "    'ParamGrid   sql2

        'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
130     F.ParamGrid.TextMatrix(0, 3) = "delete from MKEPYO where 1=1 AND "    'ParamGrid              DelSQL

        'ποιό update query θα δημιουργείται (προαιρετικά )
        'αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
140     F.ParamGrid.TextMatrix(0, 4) = ""
150     F.ParamGrid.TextMatrix(1, 1) = "ΕΤΟΣ"
160     F.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "
170     F.ParamGrid.TextMatrix(3, 1) = "Α.Φ.Μ. "
180     F.ParamGrid.TextMatrix(4, 1) = "ΠΑΤΡΩΝΥΜΟ "
190     F.ParamGrid.TextMatrix(5, 1) = "ΕΠΑΓΓΕΛΜΑ "
200     F.ParamGrid.TextMatrix(6, 1) = "ΔΟΥ"
210     F.ParamGrid.TextMatrix(7, 1) = "ΚΩΔ.ΔΟΥ "
220     F.ParamGrid.TextMatrix(8, 1) = "Υποκ. ή Ε για Κεντρικό "
230     F.ParamGrid.TextMatrix(9, 1) = "ΠΟΛΗ "
240     F.ParamGrid.TextMatrix(10, 1) = "ΟΔΟΣ "
250     F.ParamGrid.TextMatrix(11, 1) = "ΑΡΙΘΜΟΣ "
260     F.ParamGrid.TextMatrix(12, 1) = "ΤΗΛΕΦΩΝΟ"
270     F.ParamGrid.TextMatrix(13, 1) = "FAX"
280     F.ParamGrid.TextMatrix(14, 1) = "E-MAIL "
290     F.ParamGrid.TextMatrix(15, 1) = "ΑΡ.ΦΑΚΕΛ "
300     F.ParamGrid.TextMatrix(16, 1) = "ΤΑΧ.ΚΩΔ."

        'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
310     F.Labels(0).Width = F.Labels(0).Width + 1000    '* TextWidth("A")
320     F.Textboxes(0).Left = F.Textboxes(0).Left + 1000    ' * TextWidth("A")
330     F.Left = Command2.Left
340     F.Top = Command2.Top + MDIForm1.Top
350     F.SHOW 1

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu12.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu12.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu12.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu12.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub create2014_Click()

        '<EhHeader>
        On Error GoTo create2014_Click_Err

        '</EhHeader>
        Dim R      As New ADODB.Recordset

        Dim RH     As New ADODB.Recordset

        Dim rPEL   As New ADODB.Recordset

        Dim SQLPEL As String
     
        Dim k      As Integer

        Dim PolTam

        PolTam = "''"

        Dim PolTim, PolPis, PolTamMhx, PolTamTam, AgoTim, AgoPis, AGOEXO

        PolTim = "''": PolPis = "''": PolTamMhx = "''": PolTamTam = "''": AgoTim = "''": AgoPis = "''": AGOEXO = "''"
        R.Open "seLect * FROM PARASTAT ", Gdb, adOpenDynamic, adLockOptimistic
        R.MoveFirst
        k = 0

        'Dim M_TAM As String

        f_tam = ""

        Do While Not R.EOF

            If R("pol") = "1" Then  ' poliseis

                Select Case R("myf")

                    Case 1
                        PolTim = PolTim + ",'" + R("eidos") + "'"

                    Case 2
                        PolPis = PolPis + ",'" + R("eidos") + "'"

                    Case 3
                        PolTam = PolTam + ",'" + R("eidos") + "'"

                    Case 4
                        PolTamMhx = PolTamMhx + ",'" + R("eidos") + "'"
                        f_tam = Trim(R("TAMEIAKI"))
                End Select
                
                
                
                
                

            End If

            If R("pol") = "2" Then  ' agores

                Select Case R("myf")

                    Case 1
                        AgoTim = AgoTim + ",'" + R("eidos") + "'"

                    Case 2
                        AgoPis = AgoPis + ",'" + R("eidos") + "'"

                    Case 5
                        AGOEXO = AGOEXO + ",'" + R("eidos") + "'"
                End Select

            End If

            R.MoveNext
        Loop

        R.Close
     
        '     PolTim = apo2kaimeta(PolTim)
        '     PolPis = apo2kaimeta(PolPis)
        '     PolTamMhx = apo2kaimeta(PolTamMhx)
        '     PolTam = apo2kaimeta(PolTam)
        '     AgoTim = apo2kaimeta(AgoTim)
        '     AgoPis = apo2kaimeta(AgoPis)

        Dim ALL As String

        ALL = PolTim + "," + PolPis + "," + PolTamMhx + "," + PolTam + "," + AgoTim + "," + AgoPis + "," + AGOEXO

MsgBox "παραστατικά για κεπυο" + ALL

        
        
        
        Dim sql As String, pol As String, ag As String

180     sql = "select 1 AS POLHS,'normal' AS PIS,'          ' as TAM, "
        sql = sql + "(TIM.AJ1+TIM.AJ2+TIM.AJ3+TIM.AJ4+TIM.AJ5) as AJIA,LEFT(ATIM,1) AS TYPOS,(TIM.FPA1+TIM.FPA2+TIM.FPA3+TIM.FPA4) as FPA," & "1 AS AEG,PEL.AFM,PEL.EPO,PEL.DOY,PEL.EPA,PEL.DIE,PEL.POL,PLAISIO AS TK,TIM.EIDOS,HME " & " INTO TIMKEPYO FROM TIM  INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD  " & " WHERE TIM.AJ1+TIM.AJ2+TIM.AJ3+TIM.AJ4+TIM.AJ5>=0 AND LEFT(ATIM,1) IN (" + ALL + ")"
190     sql = sql + " and HME>='" + Format(d1, "MM/DD/YYYY") + "' AND HME<='" + Format(d2, "MM/DD/YYYY") + "'"

        On Error Resume Next

200     Gdb.Execute "DROP TABLE TIMKEPYO"
210     Gdb.Execute "DROP TABLE TIMKEPYO2"
212     Gdb.Execute "DROP TABLE  SYNOLAKEPYO"

220   ' Gdb.Execute "DROP TABLE PEL22"

        Dim nn
On Error GoTo 0
230     Gdb.Execute sql, nn

        Gdb.Execute "UPDATE TIMKEPYO SET POLHS=1 WHERE TYPOS IN (" + PolTim + ")"
        
        
        R.Open "select count(*) from TIMKEPYO  WHERE AFM='000000000' and TYPOS IN (" + PolTim + ")"
        If R(0) > 0 Then
            
            MsgBox "Προσοχή βρέθηκαν " + str(R(0)) + " πελατες με ΑΦΜ=000000000 και θα μεταφερθούν στις λιανικές"
            Gdb.Execute "UPDATE TIMKEPYO SET POLHS=3,AFM='000000000',EPO='',EPA='',DIE='',DOY=''   WHERE AFM LIKE '00000000%' and TYPOS IN (" + PolTim + ")", k
        End If
        
' Gdb.Execute "UPDATE TIMKEPYO SET SET POLHS=3,AFM='000000000',EPO='',EPA='',DIE='',DOY=''   WHERE AFM='000000000' and POLHS=1", K
        
        
        Gdb.Execute "UPDATE TIMKEPYO SET POLHS=2 WHERE TYPOS IN (" + AgoTim + ")"
        Gdb.Execute "UPDATE TIMKEPYO SET POLHS=1,PIS='credit'  WHERE TYPOS IN (" + PolPis + ")"
        Gdb.Execute "UPDATE TIMKEPYO SET POLHS=2,PIS='credit'  WHERE TYPOS IN (" + AgoPis + ")"

        Gdb.Execute "UPDATE TIMKEPYO SET POLHS=3,PIS='CASH'  WHERE TYPOS IN (" + PolTam + ")"
        Gdb.Execute "UPDATE TIMKEPYO SET AFM='000000000',EPO='',EPA='',DIE='',DOY=''   WHERE TYPOS IN (" + PolTam + ")"
        
        
        
        
        Gdb.Execute "UPDATE TIMKEPYO SET POLHS=4,PIS='CASH'  WHERE TYPOS IN (" + PolTamMhx + ")"
        Gdb.Execute "UPDATE TIMKEPYO SET POLHS=5,PIS='CREDIT'  WHERE TYPOS IN (" + AGOEXO + ")"

240     If nn = 0 Then
250         MsgBox "ΔΕΝ ΔΗΜΙΟΥΡΓΗΘΗΚΕ ΤΟ ΑΡΧΕΙΟ"
260         MsgBox Err.Description

            Exit Sub

        End If

        Dim a

       R.Close
       R.Open "select * from TIMKEPYO", Gdb, adOpenDynamic, adLockOptimistic

        Do While Not R.EOF
            a = check_afm(R("AFM"))

            If a = 0 Then
                MsgBox R("EPO")
            End If

            R.MoveNext
        Loop

        R.Close

        Gdb.Execute "SELECT POLHS,PIS,AFM,SUM(AJIA) AS AJIA,SUM(FPA) AS FPA,SUM(AEG) AS TEM,MAX(HME) AS SHME,EPO,EPA,DIE,DOY into SYNOLAKEPYO FROM TIMKEPYO GROUP BY POLHS,AFM,PIS,EPO,DOY,EPA,DIE ORDER BY POLHS,AFM"

        'print3_xar "SELECT POLHS,AFM,EPO,DOY,EPA,DIE,SUM(AJIA) AS [ΣΥΝ.ΑΞΙΑ],SUM(FPA) AS [ΣΥΝ.ΦΠΑ],SUM(AEG) AS [ΑΡ.ΤΙΜΟΛ]  FROM TIMKEPYO GROUP BY POLHS,AFM,EPO,DOY,EPA,DIE ORDER BY POLHS,AFM", "11111111", "ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΑΠΟ " + Format(D1, "DD/MM/YYYY") + " ΕΩΣ " + Format(d2, "DD/MM/YYYY"), 0 ' RR.RecordSource

        Adodc1.ConnectionString = gConnect
        Adodc1.RecordSource = "SELECT * FROM SYNOLAKEPYO"
        Adodc1.Refresh

        Exit Sub

create2014_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu12.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu12.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "BOHU12"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU12.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>


f_AFM = Val(FINDPARAMETROI(1, "BOHU12", "F_AFM", "0", "ΑΦΜ ΕΠΙΧΕΙΡΗΣΗΣ"))


Text1.Text = f_AFM
Text2.Text = Format(Month(Now), "00")
Text3.Text = Format(Year(Now), "00")

100     Me.KeyPreview = True

        Dim sql As String

110     mForm_Load Me, fh, fw, ft, fl

        '  'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
120     Me.Picture = LoadPicture(gPicture)

        Dim R As New ADODB.Recordset

        Dim k As Integer

        ' POL,TITLOS,EIDOS
130     R.Open "seLect * FROM PARASTAT WHERE AJIA_APOU IN ('1','2','3','4') AND XONDR=1", Gdb, adOpenDynamic, adLockOptimistic
        'seLect * FROM PARASTAT WHERE AJIA_APOU IN ('3','4') AND XONDR=1
140     R.MoveFirst
150     k = 0

160     Do While Not R.EOF

170         If (Not IsNull(R("TITLOS"))) And R("pol") = "1" Then
180             List1.AddItem R("TITLOS") + Space(1) + R("eidos")
190             List1.Selected(k) = True
200             k = k + 1
            End If

210         R.MoveNext
        Loop

220     R.Close
230     R.Open "seLect * FROM PARASTAT WHERE AJIA_APOU IN ('1','2') ", Gdb, adOpenDynamic, adLockOptimistic
240     R.MoveFirst
250     k = 0

260     Do While Not R.EOF

270         If Not IsNull(R("TITLOS")) And R("pol") = "2" Then
280             List2.AddItem R("TITLOS") + Space(1) + R("eidos")
290             List2.Selected(k) = True
300             k = k + 1
            End If

310         R.MoveNext
        Loop

320     R.Close

        On Error GoTo CREATE_MKEPYO

330     R.Open "select * FROM MKEPYO", Gdb, adOpenDynamic, adLockOptimistic

340     If R.EOF Then GoTo CREATE_MKEPYO

350     R.Close

        Exit Sub

CREATE_MKEPYO:

        '"CREATE TABLE PARAMETROI (FORMA CHAR(20),VAR CHAR(25) ,TIMH CHAR(10),SXOLIA CHAR(50)  )"
360     Gdb.Execute "create table MKEPYO (" & "M_ETOS CHAR(4)," & "M_ONOMA  char(28)," & "M_AFM char( 9)," & "M_PATR   char(19)," & "M_EPA    char(16)," & "M_DOY    char(14)," & "M_DOY2    char(4)," & "M_YPOK    char(4)," & "M_POLH   char(10)," & "M_ODOS   char(16)," & "M_ARIT   char( 3)," & "M_THL    char( 10 )," & "M_FAX   char( 10 )," & "M_EMAIL    char( 34)," & "AR_FAKEL char(6)," & "TAX_KOD  char(5)  )"

370     Gdb.Execute "INSERT INTO MKEPYO (M_ETOS) VALUES ('2008')"

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu12.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu12.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'
'
'proc sygk_tim
'******************* συγκεντρωτική τιμολογίων *********************
'PRIV ARR2[1],pistotika:=0
'priv ektypvsh, eid_tim, orio, ajia_timol, progr, DIEPICT, AYJ, arit_seirvn, Error:=0
'priv Log, m_recEGGtim, m_recTIM, Files:=0, files3:=0
'priv ddd,files2:=0,PP
'
'    priv marxeio:='EKT'+XEIRISTHS
'    strupin:={   {'LINE' , 'C', 250  , 0 }  }
'    dbcreate(MARXEIO , strupin )
'
'select 20
'IF FILE("MKEP.DBF")
'  use mkep excl
'End If
'
'  STRUPIN:={;
'& "M_ETOS'   char( 4), _
    '& "M_ONOMA'  char(28), _
    '& "M_AFM'    char( 9), _
    '& "M_PATR'   char(19), _
    '& "M_EPA'    char(16), _
    '& "M_DOY'    char(14), _
    '& "M_DOY2'    char(4), _
    '& "M_YPOK'    char(4), _
    '& "M_POLH'   char(10), _
    '& "M_ODOS'   char(16), _
    '& "M_ARIT'   char( 3), _
    '& "M_THL'    char( 10 , 0}, _
    '& "M_FAX'   ,'C' , 10 ), _
    '& "M_EMAIL'    char( 34 ,0 }, _
    '& "AR_FAKEL' char(6), _
    '& "TAX_KOD'  char(5)}
''
'     lk_check ("MKEP.DBF")
'      use MKEP EXCL
'      if reccount()=0
'        appe blank
'      End If
'
'      go 1
'    Clear
'
'    @ 0,0 say ""
'Text
'   Προσοχή !!!!!!!
'
'   Οι πελάτες/προμηθευτές του δημοσίου   πρεπει να έχουν κατηγορία  00
'
'   Οι πελάτες/προμηθευτές του εχωτερικού  πρεπει να έχουν κατηγορία -1
'
'endtext
'
'
'
'    @ 8,0 say 'Προς ΔΟΥ ' get m_doy
'    @ 9,0 say 'Kωδ.ΔΟΥ ' get m_doy2
'    @ 10,0 say 'KENT=E EΔΡΑ=0 ΑΡ.ΕΓΚ.ΜΗΤΡ.ΤΑΧΙS ΥΠΟΚ. ' get m_YPOK
'    Read
'
'    @ 11,0 say 'Eτoς     ' get m_etos
'    @ 12,0 say 'Oνομ/μο  ' get m_onoma
'    @ 13,0 say 'AΦΜ      ' get m_afm
'    @ 14,0 say 'ΠΑΤΡΩΝΥΜΟ' get m_patr
'    @ 15,0 say 'ΕΠΑΓΓΕΛΜΑ' get m_epa
'    @ 16,0 say 'ΠΟΛΗ'      get m_POLH
'    @ 17,0 say 'ΟΔΟΣ'      get m_ODOS
'    @ 18,0 say 'ΑΡΙΘΜΟΣ'   get m_ARIT
'    @ 19,0 say 'ΤΗΛΕΦΩΝΟ'  get m_thl
'    @ 20,0 say 'ΑΡ.ΦΑΚΕΛΟΥ (ΑΕ)'  get AR_FAKEL
'    @ 21,0 say 'ταχ.κωδικας'  get tax_kod
'    @ 22,0 say 'fax'  get m_fax
'    @ 23,0 say 'E-mail'  get m_email
'    rec_lock (0)
'    Read
'    unlock
'    m_die = m_odos + ' '+ m_arit+' '+ m_polh
'
'
'
'
'
'
'
'Clear
'ektypvsh = 1
'do epilogh2 with 2,ektypvsh,3,"ΣΕ ΟΘΟΝΗ","ΕΚΤΥΠΩΤΗ","APXEIO"
'if lastkey()=27
'  Return
'End If
'
'if ektypvsh <> 3
'  eid_tim = 1
'  do epilogh2 with 4,eid_tim,3,"ΠΩΛΗΣΗΣ","ΠΑΡΟΧΗΣ ΥΠΗΡΕΣΙΩΝ","ΑΓΟΡΑΣ+ΠΑΡΟΧΗΣ"
'Else
'  eid_tim = 4
'End If
'
'
'if lastkey()=27
'  Return
'End If
'
'*if eid_tim<3
'  orio = 0
'  DDD = "GgDΞΓ     "
'  apo = ctod("  /  /    ")
'  eos = ctod("  /  /    ")
'
'  priv mshow:='Ν'
'  @ 6,0 say "ΔΩΣΕ ΟΡΙΟ      " GET orio
'  @ 7,0 say "Kωδικοί Αγορών " GET ddd
'  @ 8,0 say "Aπό            " GET apo
'  @ 9 ,0 say "Εως            " GET eos
'  @ 10,0 say "Να δείχνει στην οθόνη αυτούς που δεν μπαίνουν σην συγκεντρωτική (00,-1);" get mshow
'  Read
'
'  DDD = Trim(DDD)
'
'
'*else
'*  orio=0
'*endif
'
'if lastkey()=27
'  Return
'End If
'
'do open_data with .f.,3,'EGGTIM','TIM','PEL'
'
'  SELE EGGTIM
'   m_recEGGtim = lastrec()
'
'  SELE TIM
'   m_recTIM = lastrec()
'
'   if year(apo)>1980
'      set filter to hme>=apo .and. hme<=eos .and. tim->aj1+tim->aj2+tim->aj3+tim->aj4+tim->aj5>orio
'      go top
'   End If
'
'*** οι λογιστές δουλεύουν μόνο το ΤΙΜ οπότε το ΤΙΜ θα έχει περισσότερες εγγραφές απο το EGGTIM
'log=if(   m_recTIM   >   m_recEggTIM  , 1 , 0)
'
'SELE PEL
'
'if eid_tim<3
'   do openkod with 'μικρ', 'τιμολόγια έως '+STR(ORIO,6)
'   do openkod with 'δημο', 'δημοσιου'
'End If
'
'    m_eidos = if(eid_tim=3,'r','e')
'    m_eidos = if(eid_tim=4,'',m_eidos ) && για να μηδενισει πελατες + προμηθευτές
'
'    seek m_eidos
'
'    if fil_lock(0)
'       repl rest artim with 0,sumtim with 0 while eidos=m_eidos
'       unlock
'    End If
'
'    SELE TIM
'
'*--------------------------------------------------------------------------
'
'    if eid_tim = 1   && .........  πώλησης
'
'       seires_[1]='ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΚΑΤΑΣΤΑΣΗ ΠΕΛΑΤΩΝ'
'       seires_[2]='(ΓΙΑ ΠΩΛΗΣΕΙΣ ΑΓΑΘΩΝ-ΠΑΡΟΧΗ ΥΠΗΡΕΣΙΩΝ(σε επιτηδευματίες),ΕΙΣΠΡΑΞΗ ΑΜΟΙΒΩΝ '
'       seires_[3]='ΑΠΟΖΗΜΙΩΣΕΩΝ,ΟΙΚΟΝΟΜΙΚΩΝ ΕΝΙΣΧΥΣΕΩΝ ΚΑΙ ΑΛΛΩΝ ΔΙΚΑΙΩΜΑΤΩΝ'
'       seires_[4]='(Aρθρο 20 $ 1 Π.Δ. 186/1992 Κ.Β.Σ.-Συντάσσεται σε τρία αντίτυπα)'
'       seires_[5]=''
'
'       do sygk_t2  with 'T'
'       do sygk_t2  with 't'
'       do sygk_t2  with 'P'
'       do sygk_t2  with 'Y'
'       do elegxoi_afm with m_eidos
'    elseif eid_tim = 3
'
'       seires_[1]='ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΚΑΤΑΣΤΑΣΗ ΠΡΟΜΗΘΕΥΤΩΝ'
'       seires_[2]='(ΓΙΑ ΑΓΟΡΕΣ ΕΜΠΟΡΕΥΣΙΜΩΝ ΚΑΙ ΠΑΓΙΩΝ ΑΓΑΘΩΝ,ΛΗΨΗ ΥΠΗΡΕΣΙΩΝ,'
'       seires_[3]='KATABOΛΗ ΑΜΟΙΒΩΝ,ΚΑΤΑΒΟΛΗ ΑΠΟΖΗΜΙΩΣΕΩΝ,ΟΙΚΟΝΟΜΙΚΩΝ'
'       seires_[4]='ΕΝΙΣΧΥΣΕΩΝ ΚΑΙ ΑΛΛΩΝ ΔΙΚΑΙΩΜΑΤΩΝ'
'       seires_[5]='(Aρθρο 20 $ 1 Π.Δ. 186/1992 Κ.Β.Σ.-Συντάσσεται σε τρία αντίτυπα)'
'
'     For k = 1 To Len(DDD)
'       do sygk_t2  with subs ( ddd , k , 1 )
'     Next
'       *      ddd = "GgDΞΓ     "
'       *do sygk_t2  with 'G'
'       *do sygk_t2  with 'g'
'       *do sygk_t2  with 'D'
'       *do sygk_t2  with 'Ξ'
'       *do sygk_t2  with 'Γ'
'
'       do elegxoi_afm with m_eidos
'    elseif eid_tim = 2
'       do sygk_t2  with 'Y'
'    Else
'       *** αρχείο ***
'       m_eidos = 'r'
'
'
'
'     @ 24,0 say "Οριο " + str(Orio)
'
'     For k = 1 To Len(DDD)
'       do sygk_t2  with subs ( ddd , k , 1 )
'     Next
'
'       m_eidos = 'e'
'       do sygk_t2  with 'T'
'       do sygk_t2  with 't'
'       do sygk_t2  with 'P'
'       do sygk_t2  with 'Y'
'
'       do elegxoi_afm with ''
'       Cls
'       ** δημιουργία αρχείου ΚΕΠΥΟ
'          SELECT 0
'            use mkep
'
'          select 0
'            net_use ( 'KEPYO',.T.,0)
'            zap
'            add_rec (0)
'                   ** EPIKEFALIDA ARXEIOY
'                      REPL SEIRES WITH '0'+'IL01    ' +RIGHT(DTOS(DATE()) , 8)+ SPACE(133)
'                      *   REPL SEIRES WITH '0'+'    IL01'+DTOS(DATE()) + SPACE(133)
'                   add_rec (0)
'                   REPL SEIRES WITH '1'+ mkep->m_etos+mkep->M_DOY2+mkep->AR_FAKEL+left(mkep->m_onoma,18)+mkep->m_afm+mkep->m_epa+mkep->m_polh+mkep->m_odos+mkep->m_arit+mkep->tax_kod+mkep->m_thl+mkep->m_fax+mkep->m_email+mkep->m_ypok
'
'
'                   *UNLOCK
'                *ENDIF
'       SELE PEL
'       set index to ep,ko
'       seek 'e'
'       aa = 0
'       s_tim = 0
'       s_aji = 0
'       do while .not. eof() .and. eidos='e'
'
'         if SUMTIM>0 .and. kod<>'δημο' .and. kodgal<>"00" .and. kodgal<>'-1' .and. kod<>'μικρ'
'            @ 20,0 say pel->kod+pel->epo
'            SELE kepyo
'            add_rec (0)
'            aa = aa + 1
'
'            *** στοιχεια αγοραστού ***
'            repl SEIRES WITH  '2'+strzero(aa,5)+;
'                 left(pel->afm,9)+''+;
'                 left(pel->epo+'        ',27);
'                 +left(pel->epa+' ',15)+;
'                TRANS(VAL(PEL->XRVMA),'9') +  left(pel->pol,10)+;
'                  left ( left(pel->die,16) + right(trim(pel->die),3),19) + ;
'                  left(pel->typos,5)+;
'                  strzero(pel->artim,7)+strzero(pel->sumtim*100,16)+space(35)
'            s_tim = s_tim + pel->artim
'            s_aji = s_aji + pel->sumtim
'         Else
'            if sumtim<0
'               pistotika = 1
'            End If
'            if SUMTIM>0 .and.( kodgal="00" .or. kodgal="-1") .and. mshow $ 'ΝnνΝ'
'              @ 22,0
'              wait pel->kod+pel->epo+pel->kodgal+ " >> Πάτησε ένα πλήκτρο"
'              @ 22,0 clear to 24,79
'            End If
'
'         End If
'         SELE PEL
'         Skip
'       enddo
'               SELE PEL
'               seek 'e' + 'δημο'
'               t_dhm = sumtim
'               s_tim = s_tim + ArTim
'               s_aji = s_aji + sumtim
'
'               SELE kepyo
'
'/*    il01ypoyrgeioy
'0 IL01 20051017
'120045111      ΛΑΓΑΚΗΣ ΓΕΩΡΓΙΟΣ  028783755ΕΜΠΟΡΙΟ         ΔΡΑΜΑ     Κ.ΠΑΛΑΙΟΛΟΓΟΥ   21 6610025210222252520127507glagakis@yahoo.com                   0
'200001094014298ΤΡΑΠΕΖΑ ΠΕΙΡΑΙΩΣ           ΤΡΑΠΕΖΑ        0ΔΡΑΜΑ                             00000020000000000250000
'3999999011000000000060000099999999900000000000000020000000000850000
'400001094014298ΠΕΙΡΑΙΩΣ                   ΤΡΑΠΕΖΑ        0                                  00000040000000004000000
'5999999011000000000000000099999999900000000000000040000000004000000
'
'
'2 tim ajias 2500 E  o pelatis
'6000 timologio dimosioy
'*/
'
'          if  t_dhm + s_aji > 0   && sep 2004
'               add_rec (0)
'               repl SEIRES WITH '3999999011'+strzero(t_dhm*100,16)+'999999999'+strzero(s_tim,16)+strzero(s_aji*100,16)+space(83)
'          End If
'         Print "Πωλήσεις:Τιμολόγια " + trans(s_tim, "999,999") + " Αξία " + trans(s_aji, "999,999.99")
'
'       * ως εδω διορθωμένα  8/4/2000
'
'
'
'       ************************* ΑΓΟΡΕΣ *********************************
'       SELE PEL
'       seek 'r'
'       aa = 0
'       s_tim = 0
'       s_aji = 0
'       do while .not. eof() .and. eidos='r'
'         if SUMTIM>0 .and. kod<>'δημο'
'            * ====================== .and. kodgal<>"00" .and. kodgal<>"-1"
'            SELE kepyo
'            add_rec (0)
'            aa = aa + 1
'            *** στοιχεια αγοραστού ***
'            repl  SEIRES WITH '4'+strzero(aa,5)+left(pel->afm,9)+''+;
'                  left(pel->epo+'        ',27)+left(pel->epa+' ',15)+TRANS(VAL(PEL->XRVMA),'9')+;
'                  left(pel->pol,10)+ left(pel->die,16)+;
'                  right(trim(space(5)+pel->die),3)+;
'                  LEFT(pel->TYPOS,5)+strzero(pel->artim,7)+;
'                  strzero(pel->sumtim*100,16)+space(35)
'           s_tim = s_tim + pel->artim
'           s_aji = s_aji + pel->sumtim
'         Else
'            if sumtim<0
'               pistotika = 1
'            End If
'         End If
'         SELE PEL
'         Skip
'       enddo
'               SELE PEL
'               seek 'r' + 'δημο'
'               t_dhm = sumtim
'               s_tim = s_tim + ArTim
'               s_aji = s_aji + sumtim
'
'               SELE kepyo
'
'         if s_aji +t_dhm > 0
'               add_rec (0)
'*                                 '5999999011000000000000000099999999900000000000000040000000004000000
'               repl  SEIRES WITH '5999999011'+strzero(100*t_dhm,16)+'999999999'+strzero(s_tim,16)+strzero(100*s_aji,16)+space(83)
'         End If
'         Print "Αγορές : Τιμολόγια " + trans(s_tim, "999,999") + " Αξία " + trans(s_aji, "999,999.99")
'               * repl  SEIRES WITH '5999999011'+strzero(100*t_dhm,16)+'999999999'+strzero(s_tim,16)+strzero(100*s_aji,16)+space(83)
'               COPY TO IL01 SDF
'               swpruncmd ( 'copy il01.txt a:il01' , 0 , '' ,'' )
'               Close DATA
'         Wait "ΠΑΤΗΣΤΕ ΕΝΤΕΡ ΓΙΑ ΣΥΝΕΧΕΙΑ "
'
'   if pistotika=1
'     Wait 'υπάρχουν πιστωτικά '
'     do il01_pistotika
'   End If
'
'   pp = 1
'   do epilogh2 with 24,pp,2,"ΒΛΕΠΩ ΤΑ ΤΙΜΟΛΟΓΙΑ ΔΗΜΟΣΙΟΥ","ΕΞΟΔΟΣ"
'
'        IF PP=1
'               select 0
'               use &marxeio
'               arr2 [1] = "LINE"
'               dbedit ( 1,0,24,79,arr2,"db17_udf")
'               Close DATA
'        End If
'
'
'
'
'               Return
'
'
'    endif   &&   δισκεττα
'*--------------------------------------------------------------------------
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
'     if error = 1  && λείπουν στοιχεία
'        Close DATA
'        Return
'     End If
'
'
'   SELE 0
'   use mkep
'
'
'    seires_[6] = left ( 'Πρός Δ.Ο.Υ. ' + mkep->m_doy + '    Eτος '+ mkep->m_etos + space(100) , 100 )
'    seires_[7] = left ( 'Ονοματεπώνυμο ή Επωνυμία  ' + mkep->m_onoma + ' Aρ.Φορολ.Μητρώου '+mkep->m_afm + space(100) , 100 )
'    seires_[8] = left ( 'Πατρώνυμο   '+ mkep->m_patr + space(100) , 100 )
'    seires_[9] = left ( 'Eπάγγελμα   '+ mkep->m_epa + '   Διεύθυνση  '+ mkep->m_odos + mkep->m_arit + '   Tηλ.'+ mkep->m_thl + space(100) , 100 )
'
'    arit_seirvn = 9
'
'
'    SELE PEL
'
'set index to ko,ep
'seek 'e' + 'μικρ'
'  if rec_lock(0)
'    dele
'  End If
'
'
'
'
'    set index to ep,ko
'
'progr=if(ektypvsh=1,"printbox2","print_xar")
'DIEPICT =if(ektypvsh=1,"XXXXXX","XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")
'seek m_eidos
'Clear
'sthles = 132
'*  EPIK='ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΚΑΤΑΣΤΑΣΗ ΤΙΜΟΛΟΓΙΩΝ '+IF(EID_TIM=1,'ΠΩΛΗΣΗΣ ΑΓΑΘΩΝ','ΠΑΡΟΧΗΣ ΥΠΗΡΕΣΙΩΝ')
'
'if ektypvsh=1
' do &progr with 6,'eidos=m_eidos .and. .not. eof()','SUMTIM>0 .and. kodgal<>"00"','000011', _
    '   "Α. Φ. Μ.","afm","XXXXXXXXX", _
    '   "EΠΩΝΥΜΙΑ ","epo","XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", _
    '   "ΕΠΑΓΓΕΛΜΑ","epa","XXXXXXXXXXXXXXX", _
    '   "ΔΙΕΥΘΥΝΣΗ","TRIM(die+SPACE(1)+pol)+space(50)","&DIEPICT", _
    '   "ΑΡ.ΤΜ","artim","99,999", _
    '   "ΑΞΙΑ ΠΑΡΑΣ","sumtim","#####,###.##"
'Else
' do &progr with 7,'eidos=m_eidos .and. .not. eof()','SUMTIM>0 .and. kodgal<>"00"','0000011', _
    '   "A/A","ayj","999-", _
    '   "Α. Φ. Μ.","afm","xXXXXXXXX", _
    '   "EΠΩΝΥΜΙΑ ","epo","XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", _
    '   "ΕΠΑΓΓΕΛΜΑ","epa","XXXXXXXXXXXXXXX", _
    '   "ΔΙΕΥΘΥΝΣΗ","TRIM(die+SPACE(1)+pol)","&DIEPICT", _
    '   "ΑΡ.ΤΜ","artim","99,999", _
    '   "ΑΞΙΑ ΠΑΡΑΣ","sumtim","#####,###.##"
'End If
'set index to ko,ep
'seek 'e' + 'μικρ'
'  if rec_lock(0)
'    dele
'  End If
'seek 'e' + 'δημο'
'  if rec_lock(0)
'    dele
'  End If
'Return
'
'proc openkod
'******** άνοιγμα κωδικων δημοσιου & μικρων τιμολογιων *********
'   priv kodikos, eponymia
'   Parameters kodikos, eponymia
'   SELE PEL
'   seek 'e' + kodikos
'   if eof()
'     /*
'          set dele off
'         seek 'e' + kodikos
'         if found()
'         if rec_lock(0)
'             recall
'             unlock
'           End If
'         Else
'     */
'
'         add_rec (0)
'         if rec_lock(0)
'           repl eidos with 'e',kod with kodikos,epo with eponymia
'           unlock
'         End If
'
'        *  endif
'        *  set dele on
'   End If
'Return
'
'
'proc sygk_t2
'*****************   σουμάρει τιμολόγια  **********************
'priv ajia_timol
'para kod_tim
'
'    seek kod_tim
'    Do While Left(atim, 1) = kod_tim
'       @ 20,0 say 'αριθμός τιμολογίου '+ subs(atim,2,5)
'
'       if pel5='konst'
'           ajia_timol = aj1 + aj2 + aj3
'       Else
'           ajia_timol = aj1 + aj2 + aj3 + aj4 + aj5 + aj6 + aj7 + aj8
'       End If
'
'       SELE PEL
'
'@ 24,60 say "ΟΡΙΟ " + str(orio)
'       if ajia_timol>orio && αν είναι μεγάλο τιμολόγιο
'           seek m_eidos + tim->kpe
'           if eof()
'               Print 'προσοχή ο κωδικός ',tim->kpe ,' σβήστηκε .  Βρές το τιμολόγιο ',subs(tim->atim,2,5)
'               Print 'και  ανοιξε την καρτέλλα του πελάτη . '
'               Wait 'πατησε ένα πλήκτρο για επιστροφή '
'               Error = 1
'               SELE TIM
'               Return
'           End If
'
'           if kodgal='00'   && δημοσίου
'
'              * καταγραφει το τιμολογιο του δημοσίου
'              select 0
'              use &marxeio
'              appe blank
'              repl LINE with pel->kod + pel->epo + "τιμολ." + tim->atim+" ημερ." + dtoc(TIM->hme)
'              Close
'              SELE PEL
'
'              seek 'e' + 'δημο'
'           End If
'
'
'        if left(shm2,4)<>'3001'   && ejvterikoy
'           if rec_lock(0)
'              if kod_tim $ 'PD'
'                  repl sumtim with sumtim - ajia_timol
'              Else
'                  repl sumtim with sumtim + ajia_timol
'                      * repl artim with artim + 1
'                      repl artim with artim + if(  RIGHT(tim->atim,1)="Ε" .and. tim->kerdos>0 .and. log=1,tim->kerdos,1)
'                  if pel5='logis'
'                      repl artim with artim + if(  RIGHT(tim->atim,1)="Ε"  .and. tim->kerdos>0 .and. log=1,tim->kerdos,1)
'                      *  repl artim with artim + if(tim->kerdos>0,tim->kerdos,1)
'                  End If
'              End If
'              unlock
'           End If
'        End If
'
'       End If
'
'
'           SELE TIM
'           Skip
'    enddo
'
'Return
'
'
'
'

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 12975, 11535, fh, fw, ft, fl
End Sub

'Two handy functions for Null handling
'You 're probably aware that most VB functions don't work well with Null values, which is an issue when you're working with database columns that can accept Nulls. For example, the following statement:
'
'Dim s As String
's = rs.FIELDS("AnyField")
'
'can raise error 94 "Invalid use of Null". The usual workaround is to force the conversion to string by appending an empty string, as in:
'
's = rs.FIELDS("AnyField") & ""
'
'However, this solution slightly affect code readability, especially if other people in your team aren't aware of the trick. A better and more flexible solution is to build a function that can convert a Null value to any other value:
'
'Function IfNull(value As Variant, Optional NullValue As Variant = "") As Variant
'    If IsNull(value) Then
'        IfNull = NullValue
'    Else
'        IfNull = value
'    End If
'End Function
'
'You can use the above function with just one argument to replace the trick based on the empty string, but you can even display a more meaningful string (in reports, for example):
'
'' clear the textbox if the Department is Null
'txtDept = IfNull(rs.FIELDS("Department"))
'' Display the name of the Department, or "unknown" if Null
'Print "Department is " & IfNull(rs.FIELDS("Department"), "unknown")
'
'Another handy function for Null handling mimicks the NullIf function in T-SQL, and is especially useful if you are importing data from a database or a text file that uses a special string - for example "Unknown" - in place of the Null constant, and you need to process these special strings as Null values:
'
'Function NullIf(value As Variant, NullValue As Variant) As Variant
'    If value = NullValue Then
'        NullIf = Null
'    Else
'        NullIf = value
'    End If
'End Function
'
'For example, you can easily convert all "Unknown" string to Null as follows:
'
'' read the value from a text file
'Line Input #1, dept$
'' store in the database, but convert
'' it to Null if equal to "Unknown"
'rs.FIELDS("Department") = NullIf(dept$, "Unknown")
'

Private Sub List1_Click()

End Sub

Private Sub List1_DblClick()

        '<EhHeader>
        On Error GoTo List1_DblClick_Err

        '</EhHeader>
        Dim k As Integer

100     For k = 0 To List1.ListCount - 1
110         List1.Selected(k) = False

        Next

        '<EhFooter>
        Exit Sub

List1_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu12.List1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu12.List1_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function apo2kaimeta(ByVal s As String) As String
    apo2kaimeta = mID(s, 2, Len(s) - 1)
End Function

Private Sub XMLAGO_Click()

'========================================================================
'==========================================================================================
    On Error Resume Next
    
    'DBGrid1.Clear


    Dim M_AFM As String: M_AFM = Text1.Text
    Dim m_mhnas As String: m_mhnas = Text2.Text
    Dim m_etos As String: m_etos = Text3.Text

If Len(M_AFM) <> 9 Then
   MsgBox "λαθος στο ΑΦΜ"
   Exit Sub
End If

    If Len(Dir("C:\SYGK", vbDirectory)) = 0 Then
        MkDir "C:\SYGK"
    End If

    Dim file

    Dim F_CASH

    F_CASH = "rmh123456"
    file = "C:\SYGK\agor.XML"
    Kill file

    Dim R As New ADODB.Recordset




Dim m_filename As String
m_filename = "C:\SYGK\AG" + Text1.Text + "_" + Text3.Text + Text2.Text + ".XML"
    Open m_filename For Output As #1






Open "C:\SYGK\synolaAG.txt" For Output As #5


    Open "C:\SYGK\Agor.XML" For Output As #1
    Print #1, "<?xml version=""1#"" encoding=""UTF-8"" standalone=""YES""?>"
    Print #1, "<packages>"
    Print #1, " <package actor_afm=""" + M_AFM + """Month = """ + m_mhnas + """ Year = """ + m_etos + """ > "
                                     ' Print #1, " <groupedRevenues action=""Replace"">"
    Print #1, " <groupedExpenses action=""replace; "">"
    R.Open "SELECT * FROM SYNOLAKEPYO WHERE POLHS=2", Gdb, adOpenDynamic, adLockOptimistic

Dim sxre As Single, spis As Single: sxre = 0: spis = 0

Dim sfpa(10) As Single


DBGrid1.row = 0: DBGrid1.Col = 1
    DBGrid1.Text = "Καθ.Αξία"

DBGrid1.row = 0: DBGrid1.Col = 2
    DBGrid1.Text = "Φ.Π.Α."


'        <groupedExpenses action="replace">
'            <expense>
'                <afm>090909099</afm>
'                <amount>0</amount>
'                <tax>0</tax>
'                <invoices>1</invoices>
'                <note>normal</note>
'                <nonObl>0</nonObl>
'                <date>2014-01-01</date>
'            </expense>


    Do While Not R.EOF
        Print #1, "  <expense>"
        Print #1, "     <afm>" + R!AFM + "</afm>"
        Print #1, "     <amount>" + Format(R!ajia, "####.00") + "</amount>"
        Print #1, "     <tax>" + Format(R!FPA, "####.00") + "</tax>"
        Print #1, "     <invoices>" + Format(R!TEM, "####00") + "</invoices>"
        Print #1, "     <note>" + R!PIS + "</note>"
        Print #1, "     <nonObl>0</nonObl>"
        Print #1, "     <date>" + Format(d2, "YYYY-MM-DD") + "</date>"
        Print #1, " </expense>"
        
        
        If R!PIS = "normal" Then
           sxre = sxre + R!ajia
           sfpa(1) = sfpa(1) + R!FPA
        Else
           spis = spis + R!ajia
           sfpa(2) = sfpa(2) + R!FPA
        End If
        
        
        R.MoveNext
    Loop
           
    Print " </groupedRevenues>"
    R.Close


DBGrid1.ColWidth(0) = 2000
DBGrid1.ColWidth(1) = 2000

DBGrid1.row = 1: DBGrid1.Col = 0
    DBGrid1.Text = "Σύνολο Χρ.τιμ. "
    DBGrid1.Col = 1
    DBGrid1.Text = Format(sxre, "###,###.00")
    
    
    DBGrid1.Col = 2
    DBGrid1.Text = Format(sfpa(1), "###,###.00")
    
    
    'SFPA(1) = SFPA(1) + R!FPA
    
On Error GoTo 0
    
    
    DBGrid1.row = 2: DBGrid1.Col = 0
    DBGrid1.Text = "Σύνολο πισ.τιμ. "
    DBGrid1.Col = 1
    DBGrid1.Text = Format(-spis, "####,###.00")
    
    DBGrid1.Col = 2
    DBGrid1.Text = Format(-sfpa(2), "###,###.00")
    
    
    
    Print #5, "Σύνολο Χρ.τιμ. " + Format(sxre, "########0.00")
    Print #5, " Πιστωτικά " + Format(spis, "########0.00")

    

 '   Print #1, "   <groupedCashRegisters action=""Replace"">"

    R.Open "SELECT SUM(AJIA) AS SAJIA,SUM(FPA) AS SFPA FROM SYNOLAKEPYO WHERE POLHS=3", Gdb, adOpenDynamic, adLockOptimistic


Dim slian As Single
'If Not R.EOF Then
'          Print #1, "  <cashregister>"
'          Print #1, "      <amount>" + Format(R!SAJIA, "######.00") + "</amount>"
'          Print #1, "      <tax>" + Format(R!sfpa, "######.00") + "</tax>"
'          Print #1, "      <date>" + Format(d2, "YYYY/MM/DD") + "</date>"
'          Print #1, "  </cashregister>"
'          slian = R!SAJIA
'Else
'       slian = 0
'
'End If

'
'    Print #5, "Λιαν" + Format(slian, "######,##0.00")
'    DBGrid1.row = 3: DBGrid1.Col = 0
'    DBGrid1.Text = "Σύνολο λιανικών"
'    DBGrid1.Col = 1
'    DBGrid1.Text = Format(slian, "####,##0.00")
'
'
'    sfpa(3) = R!sfpa
'    DBGrid1.Col = 2
'    DBGrid1.Text = Format(R!sfpa, "####,##0.00")





R.Close
    
'ΣΟΥΜΑ ΕΞΟΔΩΝ ================================
    

R.Open "SELECT SUM(AJIA) AS SAJIA,SUM(FPA) AS SFPA FROM SYNOLAKEPYO WHERE POLHS=3", Gdb, adOpenDynamic, adLockOptimistic
'Dim slian As Single
'If Not R.EOF Then
'          Print #1, "  <cashregister>"
'          Print #1, "      <amount>" + Format(R!SAJIA, "######.00") + "</amount>"
'          Print #1, "      <tax>" + Format(R!sfpa, "######.00") + "</tax>"
'          Print #1, "      <date>" + Format(d2, "YYYY/MM/DD") + "</date>"
'          Print #1, "  </cashregister>"
'          slian = R!SAJIA
'Else
'       slian = 0
'
'End If
'
'
'
          
          
          
          
'         <otherExpenses>
'            <amount>1019,1</amount>
'            <tax>109,1</tax>
'            <date>2014-09-09</date>
'        </otherExpenses>
'    </package>
'</packages>
          
          
          
          
Dim sTam As Single
If Not R.EOF Then
    Print #1, "   <otherExpenses>"
    Print #1, "  <cashreg_id>" + F_CASH + "</cashreg_id>"
    Print #1, "      <amount>" + Format(R!SAJIA, "######.00") + "</amount>"
    Print #1, "      <tax>" + Format(R!sfpa, "######.00") + "</tax>"
    Print #1, "      <date>" + Format(d2, "YYYY-MM-DD") + "</date>"
    Print #1, "  </otherExpenses>"
    sTam = R!SAJIA
Else
    sTam = 0
End If
           
'
'    Print #5, "Λιαν" + Format(sTam, "####,##0.00")
'    DBGRID1.row = 4: DBGRID1.Col = 0
'    DBGRID1.Text = "Σύνολο ταμειακών"
'    DBGRID1.Col = 1
'    DBGRID1.Text = Format(sTam, "####,##0.00")
'    Print #5, Format(sTam, "####,##00.00")
'    sfpa(4) = R!sfpa
'
'    DBGRID1.row = 4: DBGRID1.Col = 2
'    DBGRID1.Text = Format(R!sfpa, "####,##0.00")
'
'    Print #5, Format(R!sfpa, "####,##00.00")
'    DBGRID1.row = 5
'    DBGRID1.Text = "Σύνολο  "
'    DBGRID1.Col = 1
'    DBGRID1.Text = Format(sxre - spis + sTam + slian, "###,##0.00")
'
'
'    DBGRID1.Col = 2
'    DBGRID1.Text = Format(sfpa(1) - sfpa(2) + sfpa(3) + sfpa(4), "###,##0.00")
'
    Print "</package>"
    Print "</packages>"

    Close #1
    
    Close #5

    MsgBox "ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ C:\SYGK\aaa.XML" + Chr(3) + " Kαι c:\sygk\synola.txt το αρχείο με τα σύνολα"
    

















End Sub

Private Sub XMLPOL_Click()
'==========================================================================================
    On Error Resume Next
    
    'DBGrid1.Clear


    Dim M_AFM As String: M_AFM = Text1.Text
    Dim m_mhnas As String: m_mhnas = Text2.Text
    Dim m_etos As String: m_etos = Text3.Text

If Len(M_AFM) <> 9 Then
   MsgBox "λαθος στο ΑΦΜ"
   Exit Sub
End If

    If Len(Dir("C:\SYGK", vbDirectory)) = 0 Then
        MkDir "C:\SYGK"
    End If

    Dim file

    Dim F_CASH

    F_CASH = "rmh123456"
    file = "C:\SYGK\XML.TXT"
    Kill file

    Dim R As New ADODB.Recordset

Open "C:\SYGK\synola.txt" For Output As #5
Dim m_filename As String

m_filename = "C:\SYGK\" + Text1.Text + "_" + Text3.Text + Text2.Text + ".XML"

    Open m_filename For Output As #1
'
    Print #1, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""YES""?>"
    Print #1, "<packages>"
    Print #1, " <package actor_afm=""" + M_AFM + """ month=""" + m_mhnas + """ year=""" + m_etos + """>"
    Print #1, " <groupedRevenues action=""replace"">"
    R.Open "SELECT * FROM SYNOLAKEPYO WHERE POLHS=1", Gdb, adOpenDynamic, adLockOptimistic

Dim sxre As Single, spis As Single: sxre = 0: spis = 0

Dim sfpa(10) As Single


DBGrid1.row = 0: DBGrid1.Col = 1
    DBGrid1.Text = "Καθ.Αξία"

DBGrid1.row = 0: DBGrid1.Col = 2
    DBGrid1.Text = "Φ.Π.Α."





Dim a
    Do While Not R.EOF
        Print #1, "  <revenue>"
        
        
        Print #1, "     <afm>" + R!AFM + "</afm>"
        Print #1, "     <amount>" + Replace(Format(R!ajia, "####.00"), ".", ",") + "</amount>"
        Print #1, "     <tax>" + Replace(Format(R!FPA, "####.00"), ".", ",") + "</tax>"
        Print #1, "     <invoices>" + Format(R!TEM, "####00") + "</invoices>"
        Print #1, "     <note>" + R!PIS + "</note>"
        a = check_afm(R!AFM)
        
        Print #1, "     <date>" + Replace(Format(R!Shme, "YYYY/MM/DD"), "/", "-") + "</date>"
        Print #1, "</revenue>"
        
        
        If R!PIS = "normal" Then
           sxre = sxre + R!ajia
           sfpa(1) = sfpa(1) + R!FPA
        Else
           spis = spis + R!ajia
           sfpa(2) = sfpa(2) + R!FPA
        End If
        
        
        R.MoveNext
    Loop
           
    Print " </groupedRevenues>"
    R.Close


DBGrid1.ColWidth(0) = 2000
DBGrid1.ColWidth(1) = 2000

DBGrid1.row = 1: DBGrid1.Col = 0
    DBGrid1.Text = "Σύνολο Χρ.τιμ. "
    DBGrid1.Col = 1
    DBGrid1.Text = Format(sxre, "###,###.00")
    
    
    DBGrid1.Col = 2
    DBGrid1.Text = Format(sfpa(1), "###,###.00")
    
    
    'SFPA(1) = SFPA(1) + R!FPA
    
On Error GoTo 0
    
    
    DBGrid1.row = 2: DBGrid1.Col = 0
    DBGrid1.Text = "Σύνολο πισ.τιμ. "
    DBGrid1.Col = 1
    DBGrid1.Text = Format(-spis, "####,###.00")
    
    DBGrid1.Col = 2
    DBGrid1.Text = Format(-sfpa(2), "###,###.00")
    
    
    
    'Print #5, "Σύνολο Χρ.τιμ. " + Format(sxre, "########0.00")
    'Print #5, " Πιστωτικά " + Format(spis, "########0.00")

    

    Print #1, "   <groupedCashRegisters action=""replace"">"

    R.Open "SELECT SUM(AJIA) AS SAJIA,SUM(FPA) AS SFPA FROM SYNOLAKEPYO WHERE POLHS=3", Gdb, adOpenDynamic, adLockOptimistic


Dim slian As Single
If Not R.EOF Then
          Print #1, "  <cashregister>"
          Print #1, "      <amount>" + Replace(Format(R!SAJIA, "######.00"), ".", ",") + "</amount>"
          Print #1, "      <tax>" + Replace(Format(R!sfpa, "######.00"), ".", ",") + "</tax>"
          Print #1, "      <date>" + Replace(Format(d2, "YYYY/MM/DD"), "/", "-") + "</date>"
          Print #1, "  </cashregister>"
          slian = R!SAJIA
Else
       slian = 0
         
End If

    
    'Print #5, "Λιαν" + Format(slian, "######,##0.00")
    DBGrid1.row = 3: DBGrid1.Col = 0
    DBGrid1.Text = "Σύνολο λιανικών"
    DBGrid1.Col = 1
    DBGrid1.Text = Format(slian, "####,##0.00")


    sfpa(3) = R!sfpa
    DBGrid1.Col = 2
    DBGrid1.Text = Format(R!sfpa, "####,##0.00")





R.Close
    
    R.Open "SELECT SUM(AJIA) AS SAJIA,SUM(FPA) AS SFPA FROM SYNOLAKEPYO WHERE POLHS=4", Gdb, adOpenDynamic, adLockOptimistic
          
          
Dim sTam As Single
If Not R.EOF Then
    Print #1, "  <cashregister>"
    Print #1, "  <cashreg_id>" + f_tam + "</cashreg_id>"
    Print #1, "      <amount>" + Replace(Format(R!SAJIA, "######.00"), ".", ",") + "</amount>"
    Print #1, "      <tax>" + Replace(Format(R!sfpa, "######.00"), ".", ",") + "</tax>"
    Print #1, "      <date>" + Format(d2, "YYYY-MM-DD") + "</date>"
    Print #1, "  </cashregister>"
    Print #1, "</groupedCashRegisters>"
On Error Resume Next
sTam = R!SAJIA
Else
    sTam = 0
End If
           
           
    'Print #5, "Λιαν" + Format(sTam, "####,##0.00")
    DBGrid1.row = 4: DBGrid1.Col = 0
    DBGrid1.Text = "Σύνολο ταμειακών"
    DBGrid1.Col = 1
    DBGrid1.Text = Format(sTam, "####,##0.00")
    'Print #5, Format(sTam, "####,##00.00")
    sfpa(4) = R!sfpa
          
    DBGrid1.row = 4: DBGrid1.Col = 2
    DBGrid1.Text = Format(R!sfpa, "####,##0.00")
    
    'Print #5, Format(R!sfpa, "####,##00.00")
           
           
           
           
           
           
    DBGrid1.row = 5
    DBGrid1.Text = "Σύνολο  "
    DBGrid1.Col = 1
    DBGrid1.Text = Format(sxre - spis + sTam + slian, "###,##0.00")
           
           
    DBGrid1.Col = 2
    DBGrid1.Text = Format(sfpa(1) - sfpa(2) + sfpa(3) + sfpa(4), "###,##0.00")
           
           
           
           
           
           
           
           
           
    Print #1, "</package>"
    Print #1, "</packages>"

    Close #1
    
    Dim k As Integer
    
    For k = 1 To 5
    Print #5, Left(DBGrid1.TextMatrix(k, 0) + Space(30), 30) + Right(Space(30) + DBGrid1.TextMatrix(k, 1), 30) + Right(Space(30) + DBGrid1.TextMatrix(k, 2), 30)
    
    Next
    
    
    
    Close #5

    MsgBox "ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ " + m_filename + Chr(3) + " Kαι c:\sygk\synola.txt το αρχείο με τα σύνολα"
    



End Sub
