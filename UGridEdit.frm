VERSION 5.00
Object = "{CCA214C0-DFEB-4C91-9F0D-2665F77F6E23}#1.2#0"; "IDAutomationLinear.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form UGridEdit 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   10440
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   10080
      TabIndex        =   15
      Text            =   "1"
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EΚΤΎΠΩΣΗ ΟΛΩΝ"
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ΗΔΗ ΥΠΑΡΧΟΝ"
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Update 
      Caption         =   "ΝΕΟ Barcode"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ΕΞΟΔΟΣ"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox timText3 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   14
      Cols            =   25
      RowHeightMin    =   300
      AllowUserResizing=   2
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode1 
      Height          =   1065
      Left            =   11160
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   1500
      _cx             =   2646
      _cy             =   1879
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      SymbologyId     =   13
      DataToEncode    =   "0940"
      Orientation     =   0
      BarHeight       =   1
      NarrowBarWidth  =   0.03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   0
      Code128CharSet  =   2
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   1
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0.2
      TopMarginCM     =   0.2
      SupplementToEncode=   ""
   End
   Begin VB.Label LTI 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "συσκ"
      Height          =   255
      Left            =   9480
      TabIndex        =   16
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label hme 
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label CH1 
      Height          =   495
      Left            =   7440
      TabIndex        =   11
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label XTI 
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LTI5 
      Caption         =   " "
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode2 
      Height          =   1065
      Left            =   4440
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   1995
      _cx             =   3519
      _cy             =   1879
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      SymbologyId     =   6
      DataToEncode    =   "123456789012"
      Orientation     =   0
      BarHeight       =   1
      NarrowBarWidth  =   0.03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   0
      Code128CharSet  =   1
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   1
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0.2
      TopMarginCM     =   0.2
      SupplementToEncode=   ""
   End
   Begin VB.Label ono 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label kod 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "UGridEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Fproth As Boolean

Dim F_XBARCODE

Dim F_YBARCODE

'Private Sub MSFlexGrid1_Click()
'
'End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim R As Long, c As Long

100     gm_str(Val(Label1.Caption)) = ""

110     For R = 1 To Grid1.rows - 1
120         For c = 1 To Grid1.ColS - 1
130             Grid1.TextMatrix(R, c) = Replace(Grid1.TextMatrix(R, c), ";", " ")
140             gm_str(Val(Label1.Caption)) = gm_str(Val(Label1.Caption)) + Grid1.TextMatrix(R, c) + ";"
            Next
        Next

150     Grid1.Clear

160     Unload Me

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '' ΚΑΤΑΧΩΡΗΣΗ BARCODE
        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        On Error Resume Next

        Dim D, PROELEYSH

        Dim barc As String

100     barc = Text1.Text    ' + ChDigEAN13(Text1.Text)
110     PROELEYSH = Left(Grid1.TextMatrix(Grid1.row, 0), 3) + Left(Grid1.TextMatrix(0, Grid1.Col), 2) + "*" + Format(Grid1.row, "00") + Format(Grid1.Col, "00")
120     Gdb.Execute "insert into BARCODES (KOD,ERG,PROELEYSH) VALUES ('" + kod.Caption + "','" + barc + "','" + PROELEYSH + "')", D

130     If D = 0 Then
140         MsgBox "Δεν καταχωρήθηκε"
        End If

150     Grid1.TextMatrix(Grid1.row, Grid1.Col) = barc

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '''EKTYPVSH OLON
        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>
        Dim x

        Dim rec  As ADODB.Recordset

        Dim REC2 As New ADODB.Recordset

        Dim REC3 As New ADODB.Recordset

        Dim F_ArPerEtik

100     F_ArPerEtik = Val(FINDPARAMETROI(1, "APOT2", "F_ArPerEtik", "2", "Αριστερό περιθώριο ετικέττας")) 'posa psifia tha exei h kathe seira

110     F_XBARCODE = Val(FINDPARAMETROI(1, "APOT2", "F_XBARCODE", "348", "Χ εικόνας"))   'posa psifia tha exei h kathe seira
120     F_YBARCODE = Val(FINDPARAMETROI(1, "APOT2", "F_YBARCODE", "824", "Y εικόνας"))   'posa psifia tha exei h kathe seira

130     For Each x In Printers

140         If InStr(x.DeviceName, "2844") > 0 Then
                ' Set printer as system default.
150             Set Printer = x

                ' Stop looking for a printer
                Exit For

            End If

        Next

        ' GoSub FIND_N_NUMBER

        Dim FF, ARX

        ' ARX = Text1.Text

        '  FF = InputBox(ARX, , ARX)

        ' Text1.Text = Right("000000000000" + LTrim(Str(FF)), 12)
        ' BarCode2.DataToEncode = FF

        Dim R As Integer, c As Integer

        Dim fores

        Dim PARAM

        Dim k As Integer

        Dim D, PROELEYSH

        Dim barc As String

        Dim X1, Y1

        Dim FontSize

160     FontSize = 16
170     X1 = F_XBARCODE - 400
180     Y1 = 100 + F_YBARCODE

190     If Len(Dir("C:\DONTSIOU.TXT", vbNormal)) > 0 Then
            '   Printer.FontName = "TAHOMA" '  "COURIER NEW"
200         X1 = F_XBARCODE - 900
210         Y1 = -340 + F_YBARCODE

            '   FONTSIZE = 14
220         DONTSIOU_ETIK

            Exit Sub

        End If

        Dim sokratis As Integer

230     If Len(Dir("C:\SOKRATIS.TXT", vbNormal)) > 0 Then
240         sokratis = 1
        End If

250     If Len(Dir("C:\LAGEURO\diesel.TXT", vbNormal)) > 0 Then
260         Printer.FontName = "TAHOMA"    '  "COURIER NEW"
270         X1 = F_XBARCODE - 900 + 500 + 700
280         Y1 = -340 + F_YBARCODE - 100
290         FontSize = 14

            ' diesel_ETIK
            ' Exit Sub
        End If

        Dim YPARXEI

        Dim sysk

        Dim sql

300     For R = 1 To Grid1.rows - 1
310         For c = 1 To Grid1.ColS - 1

320             If Val(Grid1.TextMatrix(R, c)) > 0 Then
330                 sql = "select * FROM BARCODES WHERE KOD='" + kod.Caption + "' AND PROELEYSH='" + Left(Grid1.TextMatrix(R, 0), 3) + Left(Grid1.TextMatrix(0, c), 2) + "*" + Format(R, "00") + Format(c, "00") + "'"
340                 sqls sql
350                 REC2.Open sql, Gdb, adOpenDynamic, adLockOptimistic

360                 List1.AddItem Left(Grid1.TextMatrix(R, 0), 3) + Left(Grid1.TextMatrix(0, c), 2) + "*" + Format(R, "00") + Format(c, "00") + "'"
370                 List1.AddItem kod.Caption + Left(Grid1.TextMatrix(R, 0), 3) + Left(Grid1.TextMatrix(0, c), 2) + "*" + Format(R, "00") + Format(c, "00") + "'"

380                 If REC2.EOF Then
390                     REC2.Close

400                     If sokratis = 1 Then
410                         GoSub FINDsokr_N_NUMBER
                        Else
420                         GoSub FIND_N_NUMBER
                        End If

430                     YPARXEI = False
440                     List1.AddItem "ΝΕΟ BARCODE"
                    Else
450                     Text1.Text = Left(REC2("ERG"), 12)
460                     REC2.Close
470                     YPARXEI = True
480                     List1.AddItem "ΥΠΑΡΧΩΝ BARCODE"
                    End If

490                 ARX = Text1.Text
500                 FF = ARX
                    'Text1.Text = Right("000000000000" + LTrim(Str(FF)), 12)
                    'BarCode2.DataToEncode = FF

510                 If sokratis = 1 Then
                        'Text1.Text = Right("000000000000" + LTrim(Str(FF)), 12)
520                     BarCode1.DataToEncode = FF
530                     barc = Text1.Text
                    Else
540                     Text1.Text = Right("000000000000" + LTrim(str(FF)), 12)
550                     BarCode2.DataToEncode = FF
560                     barc = Text1.Text + ChDigEAN13(Text1.Text)
                    End If

570                 PROELEYSH = Left(Grid1.TextMatrix(R, 0), 3) + Left(Grid1.TextMatrix(0, c), 2) + "*" + Format(R, "00") + Format(c, "00")

580                 If YPARXEI = False Then
590                     sql = "insert into BARCODES (KOD,ERG,PROELEYSH) VALUES ('" + kod.Caption + "','" + barc + "','" + PROELEYSH + "')"
600                     sqls sql
610                     Gdb.Execute sql, D
                    End If

620                 If Val(Text2.Text) = 0 Then sysk = 1 Else sysk = Val(Text2.Text)

630                 For k = 1 To Val(Grid1.TextMatrix(R, c)) / sysk

640                     If sokratis = 1 Then
'650                         Printer.CurrentY = 400    '900 ok     '200 EKOBE TO WOW STHN MESH
'660                         Printer.FontSize = 6
'670                         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
'680                         Printer.FontBold = True
'690                         Printer.CurrentY = 12
'700                         Printer.Print Tab(7); ""
'
'710                         If Val(Me.CH1.Caption) = 1 Then
'720                             Printer.Print Tab(7); "ΠΡΟΣΦΟΡΑ"    ' Me.CH1.Caption
'                            Else
'730                             Printer.Print Tab(10); "Michelle"    ' Me.CH1.Caption
'                            End If
'
'740                         Printer.FontSize = 10
'750                         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
'760                         Printer.FontBold = True
'770                         Printer.Print Tab(6); Format(Val(lti5.Caption), "##0.00") + "€"
'
'780                         BarCode1.ShowText = YES
'
'790                         BarCode1.DataToEncode = Right(FF, 8)
'800                         BarCode1.NarrowBarWidth = 0.025  '0.03
'                            '   BarCode1.DataToEncode = DataGrid1.Text
'                            ' , X1, Y1, , 700
'810                         Printer.PaintPicture BarCode1.Picture, 0, 700, , 400    ' '-88,124   ' HTAN 500
'820                         Printer.CurrentY = 600    '400
'830                         Printer.FontSize = 6
'
'840                         Printer.Print ; kod.Caption; "--"; Grid1.TextMatrix(R, 0)    'Text1(0).Text; " "; mlabel
'850                         Printer.FontSize = 8
'860                         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
'870                         Printer.CurrentY = 600    '900 ok   '400
'880                         Printer.EndDoc

                        Else

890                         Printer.CurrentY = 0   '000000020404
900                         Printer.CurrentX = 0    ' M'600 + F_ArPerEtik * 100  -  000100007060

910                         Printer.FontSize = 8
920                         Printer.Print Tab(F_ArPerEtik); ono.Caption
930                         BarCode2.SymbologyID = EAN13

                            '---  Printer.PaintPicture BarCode2.Picture, 0, 100, , 700
940                         Printer.PaintPicture BarCode2.Picture, X1, Y1, , 700
950                         Printer.CurrentY = 900  '000000020404
960                         Printer.CurrentX = 0    ' M'600 + F_ArPerEtik * 100
970                         Printer.Print Tab(7); Trim(CH1.Caption) + Format(100 * Val(XTI), "0000") + Left(Format(hme.Caption, "DDMM"), 4) + Right(hme.Caption, 2) + "      " + Grid1.TextMatrix(0, c) + "-" + Grid1.TextMatrix(R, 0)
980                         Printer.CurrentY = 400  '000000020404
990                         Printer.CurrentX = 1700 - 500 - 1000 ' M'600 + F_ArPerEtik * 100
1000                        Printer.FontSize = 16  'HTAN 16
1010                        Printer.Print Format(Val(lti5.Caption), "##0.00") + "€"
1020                        Printer.EndDoc

                        End If

                    Next

                End If

                '   REC2.Close

1030        Next c  ''000000027366

1040    Next R

        '' ΚΑΤΑΧΩΡΗΣΗ BARCODE
        '  On Error Resume Next
        '  Dim d, PROELEYSH
        '  Dim barc As String
        '  barc = Text1.Text + ChDigEAN13(Text1.Text)
        '  PROELEYSH = Left(Grid1.TextMatrix(Grid1.row, 0), 3) + Left(Grid1.TextMatrix(0, Grid1.Col), 2) + "*" + Format(Grid1.row, "00") + Format(Grid1.Col, "00")
        '  Gdb.Execute "insert into BARCODES (KOD,ERG,PROELEYSH) VALUES ('" + kod.Caption + "','" + barc + "','" + PROELEYSH + "')", d
        '  If d = 0 Then
        '     MsgBox "Δεν καταχωρήθηκε"
        '  End If
        '   Grid1.TextMatrix(Grid1.row, Grid1.Col) = barc
        '
        ' timText3.Text = barc

        '  If FF <> ARX Then
        '      Dim ANS
        '      ANS = MsgBox("Να αποθηκευτεί ο αριθμός " + FF, vbYesNo)
        '      If ANS = vbYes Then
        '         GoSub update_N_NUMBER
        '      End If
        '   End If

        Exit Sub

FINDsokr_N_NUMBER:
1050    MN = 0
1060    REC3.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
1070    REC3.MoveFirst
1080    REC3.Move 10

1090    If IsNull(REC3("epel")) Then
1100        REC3("epel") = 1
        Else
1110        REC3("epel") = REC3("epel") + 1
1120        Gdb.Execute "UPDATE MEM SET EPEL=EPEL+1 WHERE    BASTIM_GID=22", MN

        End If

        On Error Resume Next

1130    If MN = 0 Then
1140        REC3.Update
        Else
1150        List1.AddItem "ME INSERT ΕΝΗΜΕΡΩΘΗΚΕ"
            'ΕΝΗΜΕΡΩΘΗΚΕ ΜΕ ΤΟ UPDATE MEM
        End If

1160    Text1.Text = Right("0000000000" + LTrim(str(REC3("EPEL"))), 8)
1170    REC3.Close
1180    Return

FIND_N_NUMBER:
        '************************
1190    MN = 0

        'rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

        'REC2.Close
1200    REC3.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
1210    REC3.MoveFirst
1220    REC3.Move 10
        'rec.Edit

1230    If IsNull(REC3("epel")) Then
1240        REC3("epel") = 1
        Else
1250        REC3("epel") = REC3("epel") + 1
1260        Gdb.Execute "UPDATE MEM SET EPEL=EPEL+1 WHERE    BASTIM_GID=22", MN

        End If

        On Error Resume Next

1270    If MN = 0 Then
1280        REC3.Update
        Else
1290        List1.AddItem "ME INSERT ΕΝΗΜΕΡΩΘΗΚΕ"
            'ΕΝΗΜΕΡΩΘΗΚΕ ΜΕ ΤΟ UPDATE MEM
        End If

1300    Text1.Text = Right("000000000000" + LTrim(str(REC3("EPEL"))), 12)
1310    REC3.Close
1320    Return

update_N_NUMBER:
        '************************

1330    REC2.Close
1340    REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
1350    REC2.MoveFirst
1360    REC2.Move 10
1370    REC2("epel") = Val(FF)
1380    REC2.Update

1390    Text4.Text = Right("000000000000" + LTrim(str(REC2("EPEL"))), 12)
1400    REC2.Close

1410    Return

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DONTSIOU_ETIK()

        '<EhHeader>
        On Error GoTo DONTSIOU_ETIK_Err

        '</EhHeader>
        Dim YPARXEI

        Dim sysk

        Dim x

        Dim rec  As ADODB.Recordset

        Dim REC2 As New ADODB.Recordset

        Dim FF, ARX

        Dim R As Integer, c As Integer

        Dim fores

        Dim PARAM

        Dim k As Integer

        Dim D, PROELEYSH

        Dim barc As String

        Dim X1, Y1

        Dim FontSize

100     FontSize = 16

110     X1 = F_XBARCODE - 400
120     Y1 = 100 + F_YBARCODE

130     For Each x In Printers

140         If InStr(x.DeviceName, "2844") > 0 Then
                ' Set printer as system default.
150             Set Printer = x

                ' Stop looking for a printer
                Exit For

            End If

        Next

160     F_ArPerEtik = Val(FINDPARAMETROI(1, "APOT2", "F_ArPerEtik", "2", "Αριστερό περιθώριο ετικέττας")) 'posa psifia tha exei h kathe seira

170     For R = 1 To Grid1.rows - 1
180         For c = 1 To Grid1.ColS - 1

190             If Val(Grid1.TextMatrix(R, c)) > 0 Then

200                 REC2.Open "select * FROM BARCODES WHERE KOD='" + kod.Caption + "' AND PROELEYSH='" + Left(Grid1.TextMatrix(R, 0), 3) + Left(Grid1.TextMatrix(0, c), 2) + "*" + Format(R, "00") + Format(c, "00") + "'", Gdb, adOpenDynamic, adLockOptimistic

210                 If REC2.EOF Then
220                     REC2.Close
230                     GoSub FIND_N_NUMBER
240                     YPARXEI = False
                    Else
250                     Text1.Text = Left(REC2("ERG"), 12)
260                     REC2.Close
270                     YPARXEI = True
                    End If

280                 ARX = Text1.Text

290                 FF = ARX

300                 Text1.Text = Right("000000000000" + LTrim(str(FF)), 12)
310                 BarCode2.DataToEncode = FF

320                 barc = Text1.Text + ChDigEAN13(Text1.Text)
330                 PROELEYSH = Left(Grid1.TextMatrix(R, 0), 3) + Left(Grid1.TextMatrix(0, c), 2) + "*" + Format(R, "00") + Format(c, "00")

340                 If YPARXEI = False Then
350                     Gdb.Execute "insert into BARCODES (KOD,ERG,PROELEYSH) VALUES ('" + kod.Caption + "','" + barc + "','" + PROELEYSH + "')", D
                    End If

360                 If Val(Text2.Text) = 0 Then sysk = 1 Else sysk = Val(Text2.Text)

370                 For k = 1 To Val(Grid1.TextMatrix(R, c)) / sysk

380                     Printer.CurrentY = 0   '000000020404
390                     Printer.CurrentX = 0    ' M'600 + F_ArPerEtik * 100  -  000100007060

400                     Printer.FontSize = 8
410                     Printer.FontName = "Arial Unicode MS"
420                     Printer.Print Tab(F_ArPerEtik); ono.Caption
430                     BarCode2.SymbologyID = EAN13

                        '---  Printer.PaintPicture BarCode2.Picture, 0, 100, , 700
440                     Printer.PaintPicture BarCode2.Picture, X1, Y1, , 700
450                     Printer.CurrentY = 900  '000000020404
460                     Printer.CurrentX = 0    ' M'600 + F_ArPerEtik * 100
470                     Printer.Print Tab(F_ArPerEtik); Trim(CH1.Caption) + Format(100 * Val(lti), "0000") + Left(Format(hme.Caption, "DDMM"), 4) + Right(hme.Caption, 2) + "      " + Grid1.TextMatrix(0, c) + "-" + Grid1.TextMatrix(R, 0)
480                     Printer.CurrentY = 400  '000000020404
490                     Printer.CurrentX = 1700    ' M'600 + F_ArPerEtik * 100
500                     Printer.FontSize = 16  'HTAN 16
                        '    Printer.Print Format(Val(LTI5.Caption), "##0.00") + "€"
510                     Printer.EndDoc
                    Next

                End If

520         Next c  ''000000027366

530     Next R

        '' ΚΑΤΑΧΩΡΗΣΗ BARCODE
        '  On Error Resume Next
        '  Dim d, PROELEYSH
        '  Dim barc As String
        '  barc = Text1.Text + ChDigEAN13(Text1.Text)
        '  PROELEYSH = Left(Grid1.TextMatrix(Grid1.row, 0), 3) + Left(Grid1.TextMatrix(0, Grid1.Col), 2) + "*" + Format(Grid1.row, "00") + Format(Grid1.Col, "00")
        '  Gdb.Execute "insert into BARCODES (KOD,ERG,PROELEYSH) VALUES ('" + kod.Caption + "','" + barc + "','" + PROELEYSH + "')", d
        '  If d = 0 Then
        '     MsgBox "Δεν καταχωρήθηκε"
        '  End If
        '   Grid1.TextMatrix(Grid1.row, Grid1.Col) = barc
        '
        ' timText3.Text = barc

        '  If FF <> ARX Then
        '      Dim ANS
        '      ANS = MsgBox("Να αποθηκευτεί ο αριθμός " + FF, vbYesNo)
        '      If ANS = vbYes Then
        '         GoSub update_N_NUMBER
        '      End If
        '   End If

        Exit Sub

FIND_N_NUMBER:
        '************************

        'rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

        'REC2.Close
540     REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
550     REC2.MoveFirst
560     REC2.Move 10

        'rec.Edit
570     If IsNull(REC2("epel")) Then
580         REC2("epel") = 1
        Else
590         REC2("epel") = REC2("epel") + 1
        End If

600     REC2.Update

610     Text1.Text = Right("000000000000" + LTrim(str(REC2("EPEL"))), 12)
620     REC2.Close
630     Return

update_N_NUMBER:
        '************************

640     REC2.Close
650     REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
660     REC2.MoveFirst
670     REC2.Move 10
680     REC2("epel") = Val(FF)
690     REC2.Update

700     Text4.Text = Right("000000000000" + LTrim(str(REC2("EPEL"))), 12)
710     REC2.Close

720     Return

        '<EhFooter>
        Exit Sub

DONTSIOU_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.DONTSIOU_ETIK " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.DONTSIOU_ETIK " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>

100     If KeyCode = 27 Then

110         Command1.SetFocus
            '  SendKeys ""
            '  DoEvents
            '   SendKeys "{ENTER}"

        End If

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     Me.KeyPreview = True
110     Fproth = True

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Paint()

        '==============================================================================
        '<EhHeader>
        On Error GoTo Form_Paint_Err

        '</EhHeader>
        Dim k  As Long, R As Long, c As Long

        Dim K2 As Long

100     K2 = 1

110     If Label1.Caption = "NoPaint" Then
            ' ok
        Else

120         If Fproth Then
130             Fproth = False
            Else

                Exit Sub

            End If

140         For R = 1 To Grid1.rows - 1
                'Grid1.ColWidth(R) = 300

150             For c = 1 To Grid1.ColS - 1
160                 k = InStr(K2, gm_str(Val(Label1.Caption)), ";")

170                 If k = 0 Then
180                     Grid1.TextMatrix(R, c) = ""
                    Else

190                     If k - K2 <= 0 Then
200                         Grid1.TextMatrix(R, c) = ""    ' mID(gm_str(Val(Label1.Caption)), k2, 1)
                        Else
210                         Grid1.TextMatrix(R, c) = mID(gm_str(Val(Label1.Caption)), K2, k - K2)
                        End If
                    End If

220                 K2 = k + 1
                Next
            Next

        End If

        Dim platos As Single

230     platos = (Grid1.Width + 80) / Grid1.ColS

240     For R = 1 To Grid1.ColS - 1
250         Grid1.ColWidth(R) = platos
260         Grid1.ColAlignment(R) = 3

        Next

270     Grid1.Height = Grid1.RowHeight(0) * (Grid1.rows + 1)

280     TIMText3.Text = Grid1.Text
290     TIMText3.Visible = False
300     TIMText3.Top = Grid1.Top + Grid1.CellTop
310     TIMText3.Left = Grid1.Left + Grid1.CellLeft
320     TIMText3.Width = Grid1.CellWidth - 15
330     TIMText3.Height = Grid1.CellHeight - 25

340     TIMText3.SelStart = 0
350     TIMText3.SelLength = Len(TIMText3.Text)

360     TIMText3.Visible = True
370     TIMText3.SetFocus

        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Form_Paint " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Form_Paint " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_DblClick()

        '<EhHeader>
        On Error GoTo Grid1_DblClick_Err

        '</EhHeader>

100     TIMText3.Visible = True

        '<EhFooter>
        Exit Sub

Grid1_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Grid1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Grid1_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_LeaveCell()

        '<EhHeader>
        On Error GoTo Grid1_LeaveCell_Err

        '</EhHeader>
100     If Label1.Caption = "NoPaint" Then
            'ok
        Else
110         Grid1.Text = TIMText3.Text
        End If

        '<EhFooter>
        Exit Sub

Grid1_LeaveCell_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Grid1_LeaveCell " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Grid1_LeaveCell " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_Scroll()

        '<EhHeader>
        On Error GoTo Grid1_Scroll_Err

        '</EhHeader>

100     TIMText3.Visible = False

        '<EhFooter>
        Exit Sub

Grid1_Scroll_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Grid1_Scroll " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Grid1_Scroll " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_EnterCell()

        '<EhHeader>
        On Error GoTo Grid1_EnterCell_Err

        '</EhHeader>

100     If Grid1.row > 12 Then    'f_SeiresTimologioy Then
110         Grid1.row = f_SeiresTimologioy
        End If

120     TIMText3.Text = Grid1.Text
130     TIMText3.Visible = False
140     TIMText3.Top = Grid1.Top + Grid1.CellTop
150     TIMText3.Left = Grid1.Left + Grid1.CellLeft
160     TIMText3.Width = Grid1.CellWidth - 15
170     TIMText3.Height = Grid1.CellHeight - 25

180     TIMText3.SelStart = 0
190     TIMText3.SelLength = Len(TIMText3.Text)

200     TIMText3.Visible = True
210     TIMText3.SetFocus

        '<EhFooter>
        Exit Sub

Grid1_EnterCell_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Grid1_EnterCell " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Grid1_EnterCell " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TIMText3_GotFocus()

        '<EhHeader>
        On Error GoTo TIMText3_GotFocus_Err

        '</EhHeader>

100     TIMText3.BackColor = vbYellow
110     TIMText3.SelStart = 0
120     TIMText3.SelLength = Len(TIMText3.Text)

        '<EhFooter>
        Exit Sub

TIMText3_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.TIMText3_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.TIMText3_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub timText3_KeyDown(KeyCode As Integer, Shift As Integer)

        'ΕΝΤΡΥ ΤΙΜΟΛΟΓΙΟΥ
        '<EhHeader>
        On Error GoTo timText3_KeyDown_Err

        '</EhHeader>
        Dim DUM

100     If KeyCode = 13 Then
110         KeyCode = 0
120         DUM = 0
        End If

130     If KeyCode > 95 Then
140         Me.Caption = KeyCode
150         DUM = 0
        End If

        '<EhFooter>
        Exit Sub

timText3_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.timText3_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.timText3_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub timText3_KeyPress(KeyAscii As Integer)

        'ΕΝΤΡΥ ΤΙΜΟΛΟΓΙΟΥ
        '<EhHeader>
        On Error GoTo timText3_KeyPress_Err

        '</EhHeader>
        Dim DUM

100     If KeyAscii = 13 Then
110         KeyAscii = 0
120         DUM = 0
        End If

130     If KeyAscii = 27 Then
140         KeyAscii = 0
150         DUM = 0
        End If

        '<EhFooter>
        Exit Sub

timText3_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.timText3_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.timText3_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub timText3_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo timText3_KeyUp_Err

        '</EhHeader>

100     If KeyCode = 13 Then   'enter ------------------------------------------------
110         KeyCode = 0

120         If Label1.Caption = "NoPaint" Then
                'ok
            Else
130             Grid1.Text = TIMText3.Text    ' κρατάει το grid το κείμενο
            End If

            '  Grid1.Text = timText3.Text
            ' μιά σειρά παρακάτω
140         If Grid1.Col < Grid1.ColS - 1 Then  ' πάει στην επόμενη κολόνα
150             Grid1.Col = Grid1.Col + 1
160             Grid1_EnterCell

                Exit Sub

            Else

170             If Grid1.row < Grid1.rows - 1 Then  ' πάει στην επόμενη SEIRA
180                 Grid1.Col = 1
190                 Grid1.row = Grid1.row + 1
200                 Grid1_EnterCell

                    Exit Sub

                End If
            End If

        End If

210     If KeyCode = 38 Then   ' ΠΑΝΩ ΒΕΛΟΣ

            '      Grid1.Text = timText3.Text ' κρατάει το grid το κείμενο

220         If Label1.Caption = "NoPaint" Then
                'ok
            Else
230             Grid1.Text = TIMText3.Text    ' κρατάει το grid το κείμενο
            End If

            'π'αει μία σειρά επάνω
240         If Grid1.row > 1 Then
250             Grid1.row = Grid1.row - 1
                'Grid1.Col = 1
            End If

260         Grid1_EnterCell
        End If

270     If KeyCode = 37 Then   ' aristera ΒΕΛΟΣ

280         If Grid1.Col >= 2 Then
                ' Grid1.Text = timText3.Text ' κρατάει το grid το κείμενο

290             If Label1.Caption = "NoPaint" Then
                    'ok
                Else
300                 Grid1.Text = TIMText3.Text    ' κρατάει το grid το κείμενο
                End If

                ' πάει μία κολόνα αριστερά
310             Grid1.Col = Grid1.Col - 1
320             Grid1_EnterCell
            End If

        End If

330     If KeyCode = 40 Then   ' KATΩ ΒΕΛΟΣ

340         If Label1.Caption = "NoPaint" Then
                'ok
            Else
350             Grid1.Text = TIMText3.Text    ' κρατάει το grid το κείμενο
            End If

            ' μιά σειρά παρακάτω
360         If Grid1.row < Grid1.rows - 1 Then
370             Grid1.row = Grid1.row + 1
380             Grid1_EnterCell
            End If

        End If

        Exit Sub

MHNYMA:
390     HandleError "Par1:timtext3_KeyUp"

400     Resume Next

        '<EhFooter>
        Exit Sub

timText3_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.timText3_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.timText3_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Update_Click()

        '<EhHeader>
        On Error GoTo Update_Click_Err

        '</EhHeader>

        Dim x

        Dim rec  As ADODB.Recordset

        Dim REC2 As New ADODB.Recordset

100     F_XBARCODE = Val(FINDPARAMETROI(1, "APOT2", "F_XBARCODE", "348", "Χ εικόνας"))   'posa psifia tha exei h kathe seira
110     F_YBARCODE = Val(FINDPARAMETROI(1, "APOT2", "F_YBARCODE", "824", "Y εικόνας"))   'posa psifia tha exei h kathe seira

120     For Each x In Printers

130         If InStr(x.DeviceName, "2844") > 0 Then
                ' Set printer as system default.
140             Set Printer = x

                ' Stop looking for a printer
                Exit For

            End If

        Next

150     GoSub FIND_N_NUMBER

        Dim FF, ARX

160     ARX = Text1.Text

170     FF = InputBox(ARX, , ARX)

180     Text1.Text = Right("000000000000" + LTrim(str(FF)), 12)
190     BarCode2.DataToEncode = FF

        Dim k

        Dim fores

        Dim PARAM

200     fores = Val(InputBox("Πόσα κομμάτια από το " + Text1.Text, , 1))

210     For k = 1 To fores
220         Printer.CurrentY = 0   '000000020404
230         Printer.CurrentX = 0    ' M'600 + F_ArPerEtik * 100
240         Printer.FontSize = 8
250         Printer.Print ono.Caption
260         BarCode2.SymbologyID = EAN13
270         Printer.PaintPicture BarCode2.Picture, F_XBARCODE - 400, 100 + F_YBARCODE, , 700

280         Printer.CurrentY = 900  '000000020404
290         Printer.CurrentX = 0    ' M'600 + F_ArPerEtik * 100
300         Printer.Print Trim(CH1.Caption) + Format(100 * Val(XTI), "0000") + Left(Format(Now, "DDMMYYYY"), 4) + Right(Format(Now, "DDMMYYYY"), 2) + "      " + Grid1.TextMatrix(0, Grid1.Col) + Grid1.TextMatrix(Grid1.row, 0)

310         Printer.CurrentY = 400  '000000020404
320         Printer.CurrentX = 1700    ' M'600 + F_ArPerEtik * 100

330         Printer.FontSize = 16
340         Printer.Print Format(Val(lti5.Caption), "##0.00") + "€"

350         Printer.EndDoc
        Next

        '' ΚΑΤΑΧΩΡΗΣΗ BARCODE
        On Error Resume Next

        Dim D, PROELEYSH

        Dim barc As String

360     barc = Text1.Text + ChDigEAN13(Text1.Text)
370     PROELEYSH = Left(Grid1.TextMatrix(Grid1.row, 0), 3) + Left(Grid1.TextMatrix(0, Grid1.Col), 2) + "*" + Format(Grid1.row, "00") + Format(Grid1.Col, "00")
380     Gdb.Execute "insert into BARCODES (KOD,ERG,PROELEYSH) VALUES ('" + kod.Caption + "','" + barc + "','" + PROELEYSH + "')", D

390     If D = 0 Then
400         MsgBox "Δεν καταχωρήθηκε"
        End If

410     Grid1.TextMatrix(Grid1.row, Grid1.Col) = barc

420     TIMText3.Text = barc

430     If FF <> ARX Then

            Dim ANS

440         ANS = MsgBox("Να αποθηκευτεί ο αριθμός " + FF, vbYesNo)

450         If ANS = vbYes Then
460             GoSub update_N_NUMBER
            End If
        End If

        Exit Sub

FIND_N_NUMBER:
        '************************

        'rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

470     REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
480     REC2.MoveFirst
490     REC2.Move 10

        'rec.Edit
500     If IsNull(REC2("epel")) Then
510         REC2("epel") = 1
        Else
520         REC2("epel") = REC2("epel") + 1
        End If

530     REC2.Update

540     Text1.Text = Right("000000000000" + LTrim(str(REC2("EPEL"))), 12)

550     Return

update_N_NUMBER:
        '************************

560     REC2.Close
570     REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
580     REC2.MoveFirst
590     REC2.Move 10
600     REC2("epel") = Val(FF)
610     REC2.Update

620     Text4.Text = Right("000000000000" + LTrim(str(REC2("EPEL"))), 12)

630     Return

        '<EhFooter>
        Exit Sub

Update_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.UGridEdit.Update_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.UGridEdit.Update_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Sub SOKRATIS_etik()
''========================etiketes ========================================================
'    Dim d1 As Date, d2 As Date, R As Integer, C As Integer, k As Integer, S As Single
'    Dim db As Database, DUM, rec As New ADODB.Recordset
'    Dim mlabel As String
'    Dim MPER As String
'    Dim mtimes As Single
'    Dim TIMM As Single
'    Dim timologio As String
'    Dim CC As String
'
'    'On Error Resume NextN  N
'
'    Dim nn As Long
'    Dim X As Printer
'
'    For Each X In Printers
'        If InStr(X.DeviceName, "SATO") > 0 Then
'            ' Set printer as system default.
'            Set Printer = X
'            ' Stop looking for a printer.
'            Exit For
'        End If
'    Next
'
'
'    On Error GoTo 0
'    rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
'
'
'    Dim R1 As New ADODB.Recordset
'    R1.Open "SELECT * FROM BARCODES WHERE KOD='" + data2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'
'
'
'
'    If R1.EOF Then
'        GoSub FIND_N_NUMBER  ' EPISTREFEI TO mlabel
'    Else
'        If Val(R1("ERG")) = 0 Then
'            GoSub FIND_N_NUMBER
'        Else
'            mlabel = R1("ERG")
'        End If
'    End If
'
'
'    MPER = data2.Recordset("ono")
'    mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)
'
'    TIMM = Format(data2.Recordset("lti5"), "#####.00")
'
'    'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")
'
'
'
'
'
'
'
'    For k = 1 To Val(mtimes)
'
'
'        Printer.FONTSIZE = 6
'        Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
'        Printer.FontBold = True
'
'
'        Printer.CurrentY = 12
'
'        Printer.Print Tab(0); MPER
'
'
'        'If Len(DataGrid1.Text) = 12 Then
'        ' BarCode1.SymbologyID = CODE128
'        'Else
'        '     BarCode1.SymbologyID = EAN13
'        'End If
'        BarCode1.ShowText = NO
'        BarCode1.NarrowBarWidth = 0.025  '0.03
'        BarCode1.DataToEncode = DataGrid1.Text
'        Printer.PaintPicture BarCode1.Picture, -18, F_YBARCODE  ' '-88,124
'
'
'
'
'        'Printer.FontName = "EAN8"
'        'Printer.FontName = "EAN8"
'        'Printer.FontSize = 8
'        Printer.CurrentY = 900
'        CC = Text1.Text
'
'        Printer.Print ; CC; " "; mlabel
'        Printer.FONTSIZE = 8
'        Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
'        'Printer.Print
'        'Printer.Print
'        'Printer.Print
'        'Printer.Print
'
'        Printer.CurrentY = 900
'        'Printer.Print Tab(5); Format(Val(TIMM), "###0.00") + " ΕΥΡΩ "
'        Printer.EndDoc
'
'    Next
'
'
'
'    Exit Sub
'
'
'
'FIND_N_NUMBER:
'    '************************
'    rec.MoveFirst
'    rec.Move 10
'    'rec.Edit
'    If IsNull(rec("epel")) Then
'        rec("epel") = 1
'    Else
'        rec("epel") = rec("epel") + 1
'    End If
'    rec.Update
'    mlabel = Right("00000000" + LTrim(Str(rec("EPEL"))), 8)
'
'    Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + data2.Recordset("KOD") + "','" + mlabel + "')", nn
'
'
'    If nn = 0 Then
'        MsgBox "ΔΕΝ ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ BARCODE"
'    End If
'
'
'
'    '  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Text1(0).Text + "'"
'    Return
'
'End Sub

