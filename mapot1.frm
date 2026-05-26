VERSION 5.00
Object = "{CCA214C0-DFEB-4C91-9F0D-2665F77F6E23}#1.2#0"; "IDAutomationLinear.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form par2 
   BackColor       =   &H00FF8080&
   Caption         =   "Εττικέτα"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   12225
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ekkrem 
      Caption         =   "Μόνο τα εκκρεμή"
      Height          =   375
      Left            =   9600
      TabIndex        =   27
      Top             =   165
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "mapot1.frx":0000
      Left            =   120
      List            =   "mapot1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   1200
      Width           =   2415
   End
   Begin MSComctlLib.ImageCombo Combo2 
      Height          =   330
      Left            =   75
      TabIndex        =   24
      Top             =   105
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
      ImageList       =   "ImageList1"
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2760
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   8220
      Width           =   3270
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Επαναφέρω σε εκκρεμότητα"
      Height          =   270
      Left            =   90
      TabIndex        =   22
      ToolTipText     =   "Το μετατρέπω σε ατιμολόγητο"
      Top             =   7920
      Width           =   2400
   End
   Begin MSFlexGridLib.MSFlexGrid timGrid 
      Height          =   1575
      Left            =   15
      TabIndex        =   21
      Top             =   1695
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   8
      FixedRows       =   7
      BackColorFixed  =   15581332
      TextStyle       =   1
      TextStyleFixed  =   1
      GridLinesFixed  =   1
      MergeCells      =   1
   End
   Begin VB.CheckBox noask 
      Caption         =   "Ολα ετικέτες"
      Height          =   210
      Left            =   3270
      TabIndex        =   20
      Top             =   7905
      Width           =   1485
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Βγάζω από εκκρεμότητα"
      Height          =   270
      Left            =   75
      TabIndex        =   18
      ToolTipText     =   "Το μετατρέπω σε τιμολογημένο"
      Top             =   7635
      Width           =   2400
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   9795
      Top             =   2730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2370
      Left            =   8010
      ScaleHeight     =   115.5
      ScaleMode       =   2  'Point
      ScaleWidth      =   153.75
      TabIndex        =   17
      Top             =   8595
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ΜΕΤΟΝΟΜΑΣΙΑ ΦΩΤΟΓΡΑΦΙΩΝ"
      Height          =   450
      Left            =   5805
      TabIndex        =   16
      Top             =   9120
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1485
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.CheckBox Check5 
      Caption         =   "CODE 128"
      Height          =   255
      Left            =   3285
      TabIndex        =   9
      Top             =   7635
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc EGGTIM 
      Height          =   375
      Left            =   360
      Top             =   9660
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin MSDataGridLib.DataGrid DBGrid1 
      Bindings        =   "mapot1.frx":002E
      Height          =   1545
      Left            =   4500
      TabIndex        =   8
      Top             =   165
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   2725
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Bindings        =   "mapot1.frx":0040
      Height          =   3615
      Left            =   30
      TabIndex        =   7
      Top             =   3300
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   16744576
      Cols            =   10
      FixedCols       =   0
      BackColorBkg    =   16744576
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSAdodcLib.Adodc TIM 
      Height          =   375
      Left            =   2940
      Top             =   9630
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      Caption         =   "tim"
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
   Begin VB.CommandButton Command5 
      Caption         =   "Ετικέτες"
      Height          =   420
      Left            =   3285
      TabIndex        =   6
      Top             =   7155
      Width           =   1365
   End
   Begin VB.CommandButton Command4 
      Caption         =   "¨Εξοδος"
      Height          =   420
      Left            =   7890
      TabIndex        =   5
      Top             =   7155
      Width           =   1365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Διαγραφή"
      Height          =   420
      Left            =   90
      TabIndex        =   4
      Top             =   7155
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Εκτύπωση"
      Height          =   420
      Left            =   5730
      TabIndex        =   3
      Top             =   7155
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Υπολογισμός"
      Height          =   360
      Left            =   2820
      TabIndex        =   1
      Top             =   1200
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1485
      TabIndex        =   0
      Top             =   495
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   2835
      TabIndex        =   14
      Top             =   450
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   112918529
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   300
      Left            =   2835
      TabIndex        =   15
      Top             =   825
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   112918529
      CurrentDate     =   38814
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6975
      Top             =   8115
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":0055
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":05A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":0AE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":102B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":18BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":1E0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":1F21
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":2473
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":29C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":2F17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":34A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mapot1.frx":3A3B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox bohuitiko 
      Height          =   285
      Left            =   9375
      TabIndex        =   29
      Top             =   6555
      Width           =   135
   End
   Begin VB.Label CallFromOther 
      BackStyle       =   0  'Transparent
      Caption         =   "CallFromOther"
      Height          =   240
      Left            =   9615
      TabIndex        =   28
      Top             =   705
      Visible         =   0   'False
      Width           =   2160
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode1 
      Height          =   1065
      Left            =   120
      TabIndex        =   25
      Top             =   8640
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
   Begin ATLCONTROLLibCtl.BarCode BarCode2 
      Height          =   1065
      Left            =   8250
      TabIndex        =   19
      Top             =   7635
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
      SymbologyId     =   13
      DataToEncode    =   "123456789012"
      Orientation     =   0
      BarHeight       =   1
      NarrowBarWidth  =   0.03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   0
      Code128CharSet  =   2
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   0
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0.2
      TopMarginCM     =   0.2
      SupplementToEncode=   ""
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Επωνυμία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Εως"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2385
      TabIndex        =   11
      Top             =   825
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Από"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2385
      TabIndex        =   10
      Top             =   465
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Αρ.Παρ/κού"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   495
      Width           =   1335
   End
End
Attribute VB_Name = "par2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim F_STADIO    ' STADIO PARASTATIKOY
'Dim F_PSIFIA_ATIM ' = Val(FindParametroi("PAR1", "F_PSIFIA_ATIM", "5", "ΨΗΦΙΑ ΤΙΜΟΛΟΓΙΟΥ (5)"))
Dim F_1ST_CHOICE    ' ΠΟΙΟ ΠΑΡΑΣΤΑΤΙΚΟ ΘΑ ΔΕΙΧΝΕΙ ΠΡΏΤΟ
Dim f_psifia_atim
Dim F_ArPerEtik    ' αριστερο περιθωριο ετικεττας
Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub bohuitiko_LostFocus()
    Combo2.Text = "*"
    Command1_Click

End Sub

Private Sub Form_Paint()

'  If CallFromOther.Caption = "1" Then
'    Command1_Click
'  End If
'



End Sub

Private Sub Form_Resize()
    mForm_Resize Me, 12465, 10365, fh, fw, ft, fl
End Sub

'Private Sub Form_Load()
'End Sub

Private Sub Form_Unload(Cancel As Integer)

End Sub











Private Sub Command1_Click()
    Dim d1 As Date, d2 As Date, synt As String





    'eggtim.DatabaseName = gDir---



    d1 = DTPicker1.Value

    d2 = DateAdd("d", 1, DTPicker2.Value)

    TIM.ConnectionString = gConnect
    ' TIM.DatabaseName = gDir


    If Val(Text1.Text) = 0 Then    ' ΒΓΆΖΕΙ ΟΛΑ ΤΑ ΠΑΡΑΣΤΑΤΙΚΑ ΚΑΙ ΔΙΑΛΕΓΕΙΣ
        If Right(Combo2.Text, 1) = "*" Then
            synt = " left(ATIM,1)>' ' "

            If Len(Text1.Text) > 0 Then
                synt = " ATIM='" + Text1.Text + "'"
            End If

        Else
            synt = " left(ATIM,1)='" + Right(Combo2.Text, 1) + "'"
        End If

    Else
        synt = " ATIM LIKE '%" + Text1.Text + "%'"
        'If f_psifia_atim = 5 Then
        '    synt = " left(ATIM,6)='" + Right(Combo2.Text, 1) + Right("00000" + Text1.Text, 5) + "'"
        'Else
        '    synt = " left(ATIM,7)='" + Right(Combo2.Text, 1) + Right("000000" + Text1.Text, 6) + "'"
        'End If
    End If

    If Len(Text2.Text) > 0 Then
        synt = synt + " AND PEL.EPO LIKE '" + SameLetters(Text2.Text) + "%'"
    End If

    If ekkrem.Value = vbChecked Then

        synt = synt + " AND ART IS NULL "

    End If




    Dim XX As String
    On Error GoTo 0

    If Combo3.Text = Combo3.List(1) Then
        XX = "select '-'+PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.KOD,PEL.EIDOS,TRP,SKOPOS,FORTOSH,ART,PROOR,PARAT,ART,B_N1,B_N2,SXETIKO  " _
           & " from TIM LEFT join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' order by PEL.EPO,HME;"
    Else
        XX = "select '-'+PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.KOD,PEL.EIDOS,TRP,SKOPOS,FORTOSH,ART,PROOR,PARAT,ART,B_N1,B_N2 " _
           & " from TIM LEFT join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' order by HME,PEL.EPO;"
    End If



    TIM.RecordSource = XX
    TIM.Refresh

End Sub


Private Sub Command2_Click()
'------------------------------------------------------------------------
    Dim sql As String, fDB As Database, d1 As Date, d2 As Date
    Dim synt As String, db As Database
    Dim DUM
    Dim X As String

    On Error GoTo LATOS

    d1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
    d2 = DateAdd("d", 1, d1)


    X = "select KODE,ATIM,left(ONOMA,35) as [Περιγραφή],POSO as [Ποσότητα],TIMM as [Tιμ_Μον]," _
      & " EKPT as [Εκπτωση],str(KAU_AJIA,10,2) as [K_Αξία],str(MIK_AJIA,10,2) as [ΜΕΦΠΑ_Αξία],PROELEYSH AS [ΜΕΓ] from EGGTIM where" _
      & " left(ATIM," + Str(f_psifia_atim + 1) + ")='" + Left(TIM.Recordset("ATIM"), f_psifia_atim + 1) + "'"
    X = X + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'"


    X = X + " AND APOT=" + Format(TIM.Recordset("B_N1"), "0") + " order by PROOD_AJ;"

    Dim TIT As String
    TIT = TIM.Recordset(0) + " Αριθ." + TIM.Recordset(1) + "  " + Format(TIM.Recordset(2), "dd/mm/yyyy")



    DUM = print3_xar(X, "111111111", TIT, 0)

    Exit Sub

    ' προσοχη δεν πρέπει να σβήνω τους πίνακες tim,eggtim

    '-----------  next lines are OK --------------------------

    On Error Resume Next

    d1 = d1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
    d2 = DateAdd("d", 1, d1)


    Set db = OpenDatabase("c:\mercvb\reports\reports.mdb")
    If gConnect = "Access" Then
        Set fDB = OpenDatabase(gDir, False, False)

    Else
        Set fDB = OpenDatabase(gDir, False, False, gConnect)
    End If
    'arn_enhm_apot
    'db.Execute "drop table eggtim"

    sql = "select fpa,atim,format(hme,'dd/mm/yy') as [Ημερομ],kode as [Kωδ_Είδ],onoma as [Περιγραφή],poso as [Ποσότητα],timm as [Tιμ_Μον],timm*poso as [Αξία]  into eggtim in 'c:\mercvb\reports\reports.mdb' "
    sql = sql + " from eggtim where left(atim,7)='" + TIM.Recordset("atim") + "' and hme>=#" + Format(d1, "mm/dd/yyyy") + "# and hme<#" + Format(d2, "mm/dd/yyyy") + "#;"

    fDB.Execute sql


    On Error GoTo 0
    synt = " left(atim,7)='" + Left(EGGTIM.Recordset("atim"), 7) + "'"

    sql = "select pel.epo,atim,format(hme,'dd/mm/yy') as shme,aji as ajia,pel.epa,pel.die,pel.afm into tim in 'c:\mercvb\reports\reports.mdb' "
    sql = sql + "from tim inner join pel on tim.eidos+tim.kpe=pel.eidos+pel.kod where " + synt + " and hme>=#" + Format(d1, "mm/dd/yyyy") + "# and hme<#" + Format(d2, "mm/dd/yyyy") + "#  order by prood_aj;"

    On Error Resume Next
    'db.Execute "drop table tim"

    fDB.Execute sql

    'CrystalReport1.Action = 1



    Exit Sub
LATOS:




End Sub

Private Sub Command3_Click()
'-----------------------διαγραφή παραστατικού ---------------
    Dim d1 As Date, d2 As Date
    Dim fDB As Database
    Dim MATIM As String
    Dim ANS As Integer, sql As String


    On Error GoTo LATOS
    d1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
    d2 = DateAdd("d", 1, d1)


    'Set db = OpenDatabase("c:\mercvb\reports\reports.mdb")
    'If gConnect = "Access" Then
    '  Set fDB = OpenDatabase(gDir, False, False)

    'Else
    '  Set fDB = OpenDatabase(gDir, False, False, gConnect)
    'End If
    '------------------------
    Dim g_hme, g_atim, DUM


    g_atim = TIM.Recordset("ATIM")

    g_hme = TIM.Recordset("sHME")


    ANS = MsgBox("Είσαι σίγουρος για την διαγραφή του παραστατικού " + g_atim, vbYesNo)
    If ANS = vbNo Then Exit Sub
    'On Error Resume Next




    If Len(gdirlog) > 0 Then
        DUM = del_Gefyroma(g_hme, g_atim, 1)
        If TIM.Recordset("trp") = "ΜΕ" Then
            DUM = del_Gefyroma(g_hme, g_atim, 2)
        End If
    End If
    MATIM = TIM.Recordset("atim")

    ' DEL GEFYROSI
    ' ENHM_APOT
    ' ENHM_PEL
    DoEvents
    On Error Resume Next
    
    arn_enhm_pel
    'Exit Sub

    DoEvents
    'arn_enhm_apot

    'Exit Sub

    sql = "delete  "
    sql = sql + " from EGGTIM where left(ATIM,7)='" + Left(MATIM, 7) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"

    Gdb.Execute sql
    DoEvents

    sql = "delete  "
    sql = sql + " from TIM where left(ATIM,7)='" + Left(MATIM, 7) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"

    Gdb.Execute sql
    DoEvents

    TIM.Refresh

    Exit Sub
LATOS:





End Sub

Private Sub Command4_Click()
'F_1ST_CHOICE = Val(FindParametroi("PAR2", "F_1ST_CHOICE", "2", "Πρoεπιλεγμένο παραστατικό")) 'posa psifia tha exei h kathe seira

    Dim DUM
    DUM = UpdateParametroi("PAR2", "F_1ST_CHOICE", Format(1 + Val(Left(Combo2.Text, 2)), "##"), "Πρoεπιλεγμένο παραστατικό")    'posa psifia tha exei h kathe seira



    'αν το καλουν απο αλλού αστο
    If CallFromOther.Caption = "1" Then
        PAR1.LASTTIMOL.Refresh
        Unload Me
    Else
        Unload Me
        SendKeys "%"
        SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}"
    End If

    'Combo2.ListIndex
End Sub

Private Sub Command5_Click()
'========================etiketes ========================================================
    Dim d1 As Date, d2 As Date, R As Integer, C As Integer, k As Integer, S As Single
    Dim db As Database, DUM, rec As New ADODB.Recordset

    'On Error Resume Next
    Dim ar_per
    ar_per = F_ArPerEtik    ' ρυθμιζομενο







    Dim X As Printer

    For Each X In Printers
        If InStr(X.DeviceName, "SATO") > 0 Then
            ' Set printer as system default.
            Set Printer = X
            ' Stop looking for a printer.
            Exit For
        End If
    Next

    'eggtim.DatabaseName = gDir
    On Error GoTo LATOS
    d1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
    d2 = DateAdd("d", 1, d1)

    TIM.ConnectionString = gConnect

    EGGTIM.ConnectionString = gConnect

    EGGTIM.RecordSource = "select KODE,PELKOD,ATIM,HME,POSO,TIMM,EID.ONO,EID.LTI5,EID.LTI,EID.XTI,EID.ERG from EGGTIM inner join EID on EGGTIM.KODE=EID.KOD where left(ATIM," + Str(1 + f_psifia_atim) + ")='" + Left(TIM.Recordset("atim"), 1 + f_psifia_atim) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'  order by PROOD_AJ ;"

    EGGTIM.Refresh
    'Label5.Caption = TIM.Recordset("epo")
    'Label2.Caption = TIM.Recordset("die")
    'Label3.Caption = TIM.Recordset("epa")
    'Label4.Caption = TIM.Recordset("afm")


    'DUM = OpenD(db)
    'Set rec = db.OpenRecordset("select *from mem")

    On Error GoTo 0
    rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic


    If EGGTIM.Recordset.EOF Then
        MsgBox "ΔΕΝ ΕΥΡΕΘΗΣΑΝ ΕΓΓΡΑΦΕΣ "
        Exit Sub
    End If
    Dim nkod


    EGGTIM.Recordset.MoveFirst
    Dim mlabel, MPER, mtimes, TIMM, timologio
    Do While Not EGGTIM.Recordset.EOF

        If IsNull(EGGTIM.Recordset("ERG")) Then
            GoSub FIND_N_NUMBER
        Else
            If Val(EGGTIM.Recordset("ERG")) = 0 Then
                GoSub FIND_N_NUMBER
            Else
                mlabel = EGGTIM.Recordset("erg")
            End If
        End If

        MPER = EGGTIM.Recordset("ono")
        nkod = EGGTIM.Recordset("kode")
        If noask.Value = vbChecked Then
            mtimes = EGGTIM.Recordset("poso")
        Else
            mtimes = InputBox("Πόσες ετικέτες;", MPER, EGGTIM.Recordset("poso"))
        End If

        TIMM = Format(EGGTIM.Recordset("LTI5"), "#####.00")
        timologio = EGGTIM.Recordset("kode") + Left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + mID(Format(EGGTIM.Recordset("hme"), "yyyyMMDD"), 3, 6)


        Dim EPIPLEON As Long
        EPIPLEON = 0

        For k = 1 To Val(mtimes)



            If Check5 Then  ' lag

                If Len(Dir("C:\LAGEURO\LAGAKIS2.TXT")) Then
                    ar_per = 5
                    ' BarCode2.SymbologyID = CODE128
                    BarCode2.DataToEncode = mlabel
                    Printer.PaintPicture BarCode2.Picture, 100, -80, 2000, 400    ' ORIZ,KATHETA
                    EPIPLEON = 3


                Else
                    Printer.FontName = "128"
                    Printer.FONTSIZE = 5
                    Printer.CurrentY = 0
                    Printer.Print Tab(10); mlabel
                End If




                Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName 12305019

                Printer.FONTSIZE = 7
                '  Printer.Print
                '  Printer.Print
                ' Printer.Print
                Printer.CurrentY = 110
                Printer.Print

                Printer.Print Tab(ar_per + EPIPLEON); mlabel
                Printer.Print Tab(ar_per + EPIPLEON); MPER
                Printer.Print Tab(ar_per + EPIPLEON); timologio
                Printer.FONTSIZE = 16
                Printer.FontBold = True
                Printer.Print Tab(ar_per); Trim(LTrim(TIMM)) + " €"

                Printer.EndDoc

            ElseIf Len(Dir("C:\LAGEURO\SOKRATIS.TXT", vbNormal)) > 0 Then

                Printer.FONTSIZE = 6
                Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
                Printer.FontBold = True
                Printer.CurrentY = 12
                Printer.Print Tab(0); nkod + MPER
                BarCode1.ShowText = NO
                BarCode1.NarrowBarWidth = 0.025  '0.03
                BarCode1.DataToEncode = mlabel
                Printer.PaintPicture BarCode1.Picture, -18, 124   '-88,124
                Printer.Print ; " "; mlabel
                Printer.FONTSIZE = 8
                Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
                Printer.CurrentY = 900
                Printer.Print Tab(5); Format(Val(TIMM), "###0.00") + " ΕΥΡΩ "
                Printer.EndDoc

            ElseIf Len(Dir("C:\LAGEURO\LOISIR.TXT")) Then

                ' 12345
                Printer.FONTSIZE = 6   '10000011  10000012


                Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName00000002
                Printer.FontBold = True
                ' Printer.CurrentY = 120
                BarCode1.ShowText = NO
                BarCode1.NarrowBarWidth = 0.025  '0.03
                BarCode1.DataToEncode = mlabel    ' "12345678" '  EGGTIM.Recordset("kode")
                '  Printer.PaintPicture BarCode1.Picture, -100, 30, , 470 '-18,124
                Printer.PaintPicture BarCode1.Picture, -130, 30, , 470    '-18,124

                Printer.FONTSIZE = 6    ' 10000020

                Printer.CurrentY = 400    '546  ΠΑΝΟ ΣΤΗΝ ΕΤΙΚΕΤΑ   '
                Printer.Print Tab(7); EGGTIM.Recordset("kode")


                Printer.FONTSIZE = 7
                Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
                Printer.CurrentY = 480    '420 OK ΠΑΝΟ ΣΤΗΝ ΕΤΙΚΕΤΑ
                ' Printer.Print

                Printer.Print Tab(6); Format(Val(TIMM), "###0.00") + "€"



                Printer.EndDoc

                '  Printer.FONTSIZE = 6
                '  Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName00000002
                '  Printer.FontBold = True
                '  Printer.CurrentY = 120
                '  BarCode1.ShowText = YES
                '  BarCode1.NarrowBarWidth = 0.025  '0.03
                '  BarCode1.DataToEncode = "12345678" '  EGGTIM.Recordset("kode")
                '  Printer.PaintPicture BarCode1.Picture, -100, 0 '-18,124
                ''  Printer.Print ; " "; mlabel
                '  Printer.FONTSIZE = 8
                '  Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
                '  Printer.CurrentY = 100
                '  Printer.Print Tab(5); Format(Val(TIMM), "###0.00") + " ΕΥΡΩ "
                '  Printer.EndDoc


















            Else

                Printer.Print Tab(20); MPER
                Printer.Print Tab(30); mlabel
                Printer.FontName = "ean13"    ' "128"
                Printer.FONTSIZE = 6
                Printer.Print Tab(18); EGGTIM.Recordset("kode")    ' mlabel
                Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
                Printer.FONTSIZE = 8
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.EndDoc
            End If


        Next

        EGGTIM.Recordset.MoveNext

    Loop

    Exit Sub

FIND_N_NUMBER:
    '************************
    rec.MoveFirst
    rec.Move 10
    'rec.Edit
    If IsNull(rec("epel")) Then
        rec("epel") = 1
    Else
        rec("epel") = rec("epel") + 1
    End If
    rec.Update
    mlabel = Right("00000000" + LTrim(Str(rec("EPEL"))), 8)
    Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + EGGTIM.Recordset("KODE") + "'"
    Return


LATOS:





End Sub

Private Sub Command6_Click()
    EGGTIM.Recordset.MoveFirst
    Dim a, b
    a = InputBox("φακελος με τις φωτογραφιες π.χ. c:\mercvb ")
    b = InputBox("Αρχικός αριθμός")

    '
    '
    Dim PI
    'PI = "C:\MERCVB\IMAGES\" + Text1(0).Text
    'If Len(Dir(PI)) > 0 Then
    '   Picture1.Picture = LoadPicture(PI)
    'Else
    '   Picture1.Picture = LoadPicture()
    'End If

    Dim ANS, F

    Do While Not EGGTIM.Recordset.EOF

        PI = a + "\" + Format(b, "0000") + ".jpg"
        If Len(Dir(PI)) > 0 Then
            Picture1.Picture = LoadPicture(PI)
        Else

            Me.Caption = "Αναζήτηση φωτογραφίας του είδους :" + EGGTIM.Recordset(3) + EGGTIM.Recordset(0)
            CD1.InitDir = a
            CD1.ShowSave
            Picture1.Picture = LoadPicture(CD1.FILENAME)
            PI = CD1.FILENAME
        End If

        ANS = MsgBox(Left(EGGTIM.Recordset(3), 25) + Chr(13) + mID$(EGGTIM.Recordset(3), 26, 15), vbYesNo, "Να αποθηκευτεί η φωτογραφία")

        If Len(Dir("C:\MERCVB\IMAGES\" + EGGTIM.Recordset("kode"))) > 0 Then
            Kill "C:\MERCVB\IMAGES\" + EGGTIM.Recordset("kode")
        End If

        If ANS = vbYes Then
            ' A + "\" + Format(B, "0000") + ".jpg"
            Name PI As "C:\MERCVB\IMAGES\" + EGGTIM.Recordset("kode")

        Else
            CD1.InitDir = a
            CD1.ShowSave
            '       Picture1.Picture = LoadPicture(CD1.FILENAME)
        End If

        b = b + 1
        EGGTIM.Recordset.MoveNext
    Loop




End Sub

Private Sub Command7_Click()
'----------------------βγαζο απο ekkremothta ---------------
    Dim d1 As Date, d2 As Date
    Dim fDB As Database
    Dim MATIM As String
    Dim ANS As Integer, sql As String


    On Error GoTo LATOS
    d1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
    d2 = DateAdd("d", 1, d1)

    Dim g_hme, g_atim, DUM


    g_atim = TIM.Recordset("ATIM")

    g_hme = TIM.Recordset("sHME")



    MATIM = TIM.Recordset("atim")

    sql = "UPDATE TIM SET ART='1' where left(ATIM," + Format(f_psifia_atim + 1, "0") + ")='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"
    Dim k As Integer

    Gdb.Execute sql, k
    If k = 1 Then timGrid.TextMatrix(5, 0) = "ΤΙΜΟΛΟΓΗΜΕΝΟ"

    DoEvents

    Dim X
    X = TIM.Recordset.Bookmark
    TIM.Refresh
    TIM.Recordset.Bookmark = X

    Exit Sub
LATOS:


End Sub

Private Sub Command8_Click()

'----------------------epanafero se ekkremothta ---------------
    Dim d1 As Date, d2 As Date
    Dim fDB As Database
    Dim MATIM As String
    Dim ANS As Integer, sql As String


    On Error GoTo LATOS
    d1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
    d2 = DateAdd("d", 1, d1)

    Dim g_hme, g_atim, DUM


    g_atim = TIM.Recordset("ATIM")

    g_hme = TIM.Recordset("sHME")



    MATIM = TIM.Recordset("atim")

    sql = "UPDATE TIM SET ART=NULL where left(ATIM," + Format(f_psifia_atim + 1, "0") + ")='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"
    Dim k As Integer

    Gdb.Execute sql, k
    If k = 1 Then timGrid.TextMatrix(5, 0) = "ΜΗ ΤΙΜΟΛΟΓΗΜΕΝΟ"

    DoEvents


    Dim X
    X = TIM.Recordset.Bookmark
    TIM.Refresh
    TIM.Recordset.Bookmark = X



    Exit Sub
LATOS:





End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DBGrid1_RowColChange 0, 0
    End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim d1 As Date, d2 As Date, R As Integer, C As Integer, k As Integer, S As Single
    On Error Resume Next
    'eggtim.DatabaseName = gDir
    d1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))  ' DTPicker1.Value
    d2 = DateAdd("d", 1, d1)
    'TIM.DatabaseName = gDir
    TIM.ConnectionString = gConnect
    EGGTIM.ConnectionString = gConnect
    On Error GoTo 0
    Dim sql As String


    Dim RDB As New ADODB.Recordset
    RDB.Open "select * from PARASTAT where EIDOS='" + Left(TIM.Recordset("atim"), 1) + "';", Gdb, adOpenForwardOnly, adLockReadOnly









    If TIM.Recordset.EOF Then Exit Sub


    sql = "select KODE,ATIM,HME as [Ημερομ],ONOMA as [Περιγραφή],POSO as [Ποσότητα]," _
        & " STR(ROUND(TIMM,2),10,2) as [Tιμ_Μον],str(TIMM*POSO,10,2) as [Αξία],EKPT as [Εκπτωση],PROELEYSH AS [ΜΕΓ] from EGGTIM" _
        & " where left(ATIM," + Str(f_psifia_atim + 1) + " )='" + Left(TIM.Recordset("ATIM"), 1 + f_psifia_atim)
    sql = sql + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy")
    If IsNull(TIM.Recordset("B_N1")) Then
        sql = sql + "'  AND APOT=1 ORDER BY PROOD_AJ;"
    Else
        sql = sql + "'  AND APOT=" + Format(TIM.Recordset("B_N1"), "0") + "  order by PROOD_AJ;"
    End If

    EGGTIM.RecordSource = sql
    EGGTIM.Refresh

    On Error Resume Next
    'Label13.Caption = TIM.Recordset("B_N1")
    'Label14.Caption = TIM.Recordset("B_N2")
    'Label5.Caption = TIM.Recordset("epo")
    'Label2.Caption = TIM.Recordset("die")
    'Label3.Caption = TIM.Recordset("epa")
    'Label4.Caption = TIM.Recordset("afm")
    'If IsNull(TIM.Recordset("SKOPOS")) Then Label10.Caption = "" Else Label10.Caption = TIM.Recordset("SKOPOS")
    'If IsNull(TIM.Recordset("PROOR")) Then label11.Caption = "" Else label11.Caption = TIM.Recordset("PROOR")
    'If IsNull(TIM.Recordset("ART")) Then Label12.Caption = "- " Else Label12.Caption = "ΤΙΜΟΛΟΓΗΘΗΚΕ"
    'Label11.Caption = TIM.Recordset("PROOR")

    'If IsNull(TIM.Recordset("trp")) Then Text3.Text = "" Else Text3.Text = TIM.Recordset("trp")
    'If IsNull(TIM.Recordset("SKOPOS")) Then Text4.Text = "" Else Text4.Text = TIM.Recordset("SKOPOS")
    'If IsNull(TIM.Recordset("fortosh")) Then Text5.Text = "" Else Text5.Text = TIM.Recordset("fortosh")
    'If IsNull(TIM.Recordset("PROOR")) Then Text6.Text = "" Else Text6.Text = TIM.Recordset("PROOR")

    timGrid.ColWidth(0) = 4000
    timGrid.ColAlignment(0) = 1


    timGrid.TextMatrix(0, 0) = TIM.Recordset("epo")
    timGrid.TextMatrix(1, 0) = TIM.Recordset("die")
    timGrid.TextMatrix(2, 0) = TIM.Recordset("epa")
    timGrid.TextMatrix(3, 0) = TIM.Recordset("afm")
    timGrid.TextMatrix(4, 0) = Format(TIM.Recordset("SHME"), "DD/MM/YYYY") + "    " + TIM.Recordset("ATIM")
    If RDB("STADIO") <= 3 Then
        Command7.Enabled = True
        Command8.Enabled = True
        If Left(TIM.Recordset("ART"), 1) = "1" Then
            timGrid.TextMatrix(5, 0) = "ΤΙΜΟΛΟΓΗΜΕΝΟ "
            If IsNull(TIM.Recordset("sxetiko")) Then
            Else
                timGrid.TextMatrix(5, 0) = "ΤΙΜΟΛΟΓΗΜΕΝΟ " + TIM.Recordset("sxetiko")
            End If

        Else
            timGrid.TextMatrix(5, 0) = "ΜΗ ΤΙΜΟΛΟΓΗΜΕΝΟ"
        End If
    Else
        Command7.Enabled = False
        Command8.Enabled = False
        timGrid.TextMatrix(5, 0) = ""
    End If

    timGrid.ColWidth(1) = 6000
    timGrid.ColAlignment(1) = 1


    timGrid.TextMatrix(0, 1) = "Πληρωμή : " + TIM.Recordset("trp")
    timGrid.TextMatrix(1, 1) = TIM.Recordset("SKOPOS")
    timGrid.TextMatrix(2, 1) = TIM.Recordset("fortosh")
    timGrid.TextMatrix(3, 1) = TIM.Recordset("PROOR")
    timGrid.TextMatrix(4, 1) = "Από Αποθήκη " + Format(TIM.Recordset("B_N1"), "##") + " " + "Σε  Αποθήκη " + Format(TIM.Recordset("B_N2"), "##")
    timGrid.TextMatrix(5, 1) = TIM.Recordset("parat")






    ' Label8.Caption = TIM.Recordset("ATIM") + " " + Format(TIM.Recordset("SHME"), "DD/MM/YYYY")

    MSFlexGrid1.ColWidth(0) = 1500
    MSFlexGrid1.ColWidth(3) = 3000
    MSFlexGrid1.ColAlignment(3) = 1
    SYNOLA

End Sub
Sub SYNOLA()
' μορφοποίηση του MSFLEXGRID1

    Dim d1 As Date, d2 As Date, R As Integer, C As Integer, k As Integer, S As Single
    'On Error Resume Next
    MSFlexGrid1.TextMatrix(0, 1) = "Αρ.Τιμ"
    MSFlexGrid1.AddItem ("" & Chr(9))    ' & "Σύνολο" & Chr(9) & eis & Chr(9) & pl & Chr(9) & Okeis & Chr(9) + Okpl)
    On Error GoTo 0
    R = MSFlexGrid1.ROWS
    C = MSFlexGrid1.ColS
    S = 0
    For k = 1 To R - 1
        MSFlexGrid1.TextMatrix(k, 2) = Format(MSFlexGrid1.TextMatrix(k, 2), "dd/mm/yyyy")
        ' MSFlexGrid1.TextMatrix(k, 5) = Format(MSFlexGrid1.TextMatrix(k, 5), "######.00")
        MSFlexGrid1.TextMatrix(k, C - 2) = Format(Val(MSFlexGrid1.TextMatrix(k, C - 1)), "#####.00")
        S = S + Val(MSFlexGrid1.TextMatrix(k, C - 2))
    Next
    MSFlexGrid1.TextMatrix(R - 1, C - 2) = Format(S, "#####.00")
    MSFlexGrid1.ColAlignment = 6
    MSFlexGrid1.AddItem ("" & Chr(9))
    'MSFlexGrid1.TextMatrix(R, c - 1) = Format(TIM.Recordset("ajia") - s, "#####.00")
    ' MSFlexGrid1.TextMatrix(R, c - 2) = "ΦΠΑ"
    MSFlexGrid1.AddItem ("" & Chr(9))
    MSFlexGrid1.TextMatrix(R + 1, C - 2) = Format(TIM.Recordset("ajia"), "#####.00")

    MSFlexGrid1.ColAlignment(3) = 1

End Sub

Private Sub DTPicker1_LostFocus()
    If DTPicker2.Value < DTPicker1.Value Then
        DTPicker2.Value = DTPicker1.Value
    End If
End Sub

Private Sub Form_Activate()

    Dim X As String, k As Integer
    Dim db As Database
    Dim R As New ADODB.Recordset

    f_psifia_atim = Val(FINDPARAMETROI("PAR1", "F_PSIFIA_ATIM", "5", "ΨΗΦΙΑ ΤΙΜΟΛΟΓΙΟΥ (5)"))
    ' db.Open gDir





    'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
    'SkinFramework1.ApplyWindow Me.hWnd
    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics


    timGrid.ColWidth(0) = 4000
    timGrid.ColWidth(1) = 6000



    'If gConnect = "Access" Then
    '   Set db = OpenDatabase(gDir, False, False)
    'Else
    '   Set db = OpenDatabase(gDir, False, False, gConnect)
    'End If
    F_1ST_CHOICE = Val(FINDPARAMETROI("PAR2", "F_1ST_CHOICE", "2", "Πρoεπιλεγμένο παραστατικό"))    'posa psifia tha exei h kathe seira
    F_ArPerEtik = Val(FINDPARAMETROI("PAR2", "F_ArPerEtik", "2", "Αριστερό περιθώριο ετικέττας"))  'posa psifia tha exei h kathe seira

    R.Open "select POL,TITLOS,EIDOS,STADIO from PARASTAT ORDER BY POL,TITLOS;", Gdb, adOpenDynamic, adLockOptimistic
    'Combo2.Clear
    If R.RecordCount = 0 Then
        Exit Sub
    End If

    Dim L As Integer, FirstTime As Boolean
    Dim o As ComboItem

    Combo3.Text = Combo3.List(1)

    L = 0
    FirstTime = False


    Set o = Combo2.ComboItems.Add(1, "", "ΟΛΑ ΤΑ ΠΑΡΑΣΤΑΤΙΚΑ     *")
    ' Combo2.ComboItems(K).Image = 2


    R.MoveFirst
    k = 1
    Do While Not R.EOF
        If Not IsNull(R("TITLOS")) And Left(R("TITLOS"), 1) <> "-" Then
            If R("POL") = 2 Then  ' αγορες
                If FirstTime = False Then L = k: FirstTime = True

                ' Combo2.AddItem Format(K, "##") + "---" + " " + R("TITLOS") + Space(1) + R("eidos")
                ' Combo2.ItemBackColor(K) = vbRed



                Set o = Combo2.ComboItems.Add(k + 1, "", Format(k, "##") + "---" + " " + R("TITLOS") + Space(1) + R("eidos"))
                Combo2.ComboItems(k).Image = 2
                ' Combo2. vbRed
            Else

                ' Combo2.AddItem Format(K, "##") + " " + R("TITLOS") + Space(1) + R("eidos")

                Set o = Combo2.ComboItems.Add(k + 1, "", Format(k, "##") + "---" + " " + R("TITLOS") + Space(1) + R("eidos"))
                Combo2.ComboItems(k + 1).Image = 3



            End If
            k = k + 1
            'mEidos_Par(K) = R("eidos")
        End If
        R.MoveNext
    Loop
    Combo2.Text = Combo2.ComboItems(F_1ST_CHOICE)
    ' Combo2.ListIndex = F_1ST_CHOICE
    ' Combo2.Text = Combo2.List(F_1ST_CHOICE)

    'DTPicker1.value = Now
    'DTPicker2.value = Now
    'If CallFromOther.Caption = "1" Then
    '    Command1_Click
    'End If





End Sub

'ενημέρωση αποθήκης αρνητικα
Sub arn_enhm_apot()

    Dim G, m, GA, MA, P
    Dim R As New ADODB.Recordset, k, L As Integer
    Dim r2
    Dim e As New ADODB.Recordset

    Dim gm_f_gfkod As String
    Dim m2, gm_f_gfFPA
    Dim m_m As Boolean
    Dim f_k, f_p
    Dim f_dat
    Dim F_POS_APOU, F_AJIA_APOU
    Dim meid As New ADODB.Recordset


    f_k = 0    ' kodikoi
    f_p = 3    ' posotites
    'If gConnect = "Access" Then
    'Set db = OpenDatabase(gDir, False, False)
    'Else
    'Set db = OpenDatabase(gDir, False, False, gConnect)
    'End If
    f_dat = EGGTIM.Recordset("Ημερομ")
    G = "G" + Format(Month(f_dat), "00")
    m = "M" + Format(Month(f_dat), "00")
    GA = "GA" + Format(Month(f_dat), "00")
    MA = "MA" + Format(Month(f_dat), "00")

    For k = 1 To MSFlexGrid1.ROWS - 1    'Do While Not R.EOF
        If Len(MSFlexGrid1.TextMatrix(k, f_k)) > 0 Then

            meid.Open "SELECT *FROM EID WHERE KOD='" + MSFlexGrid1.TextMatrix(k, f_k) + "'", Gdb, adOpenKeyset, adLockOptimistic

            If meid.RecordCount = 1 Then
                P = Val(MSFlexGrid1.TextMatrix(k, f_p))    '("POSO")

                ' a = P * Val(MSFlexGrid1.TextMatrix(K, f_t)) * (100 - Val(MSFlexGrid1.TextMatrix(K, f_e))) / 100

                R.Open "select * from PARASTAT where EIDOS='" + Left(TIM.Recordset("atim"), 1) + "';", Gdb, adOpenForwardOnly, adLockReadOnly

                'f_PEL = R("pel")
                Dim f_matim
                F_POS_APOU = R("POS_APOU")
                F_AJIA_APOU = R("ajia_APOU")
                R.Close



                Dim mYP As Single, mG As Single, MM As Single
                mYP = meid!pos


                Select Case F_POS_APOU
                Case "1"    ' EISAGOGES=1
                    mYP = meid("POS") + P
                    mG = meid(G) + P
                Case "2"    ' -EISAGOGES
                    mYP = meid("POS") - P
                    '         mG = meid(G) - P
                Case "3"    ' EXAGOGES=1
                    mYP = meid("POS") - P
                    MM = meid(m) + P
                Case "4"    ' -EXAGOGES
                    mYP = meid("POS") + P
                    MM = meid(m) - P
                End Select

                Select Case F_AJIA_APOU
                Case "1"    ' EISAGOGES=1
                    ' meid(GA) = meid(GA) + a
                Case "2"    ' -EISAGOGES
                    '         meid(GA) = meid(GA) - a
                Case "3"    ' EXAGOGES=1
                    '         meid(MA) = meid(MA) + a
                Case "4"    ' -EXAGOGES
                    '  meid(MA) = meid(MA) - a
                End Select
                Gdb.Execute "UPDATE EID SET POS=" + Str(mYP) + " ," + G + "=" + Str(mG) + "," + m + "=" + Str(MM)
            End If  ' If meid.RecordCount = 1 Then


            meid.Close
        End If    'len( msflexgrid1.TextMatrix(K, f_k) )>0
    Next



End Sub

Function del_Gefyroma(g_hme, g_atim, pist)
'-------- pist=1 pistosei   pist=2 metrhtois
'*********************** γεφυρωμα ****************************
    Dim xre_pis(20), xre_Pis_rec(20), xreosis(20), GetList(20)
    Dim DUM

    Dim JJ, JJK, m_ait As String
    Dim kodiklog(9), TYP(9), xp(9), kau(9), a As String, mnq As String, eid_par As String
    Dim mgre As String
    Dim k As Integer

    Dim atr_xre, atr_pis, LW, DIAFORETIKA
    atr_xre = 0: atr_pis = 0
    Dim ok_gefyr
    Dim db As Database, dbL As Database
    Dim abat As Recordset, temp As Recordset, r2 As Recordset, gefyres As Recordset
    Dim bbs As String
    Dim anoigma As Integer
    Dim kod_pel As String
    Dim mbat_pel As String
    Dim abat2 As Recordset
    Dim lo As Integer
    Dim JJ_ROBLEM
    Dim mmhkos, MHKOS(5) As Integer

    Dim MHNAS
    Dim sql, K2 As Integer, logar(5) As Integer

    If gConnect = "Access" Then
        Set db = OpenDatabase(gDir, False, False)
        Set dbL = OpenDatabase(gdirlog, False, False)
    Else

        On Error Resume Next
        Set db = OpenDatabase(gDir, False, False, gConnect)
        Set dbL = OpenDatabase(gdirlog, False, False, gConnect)
    End If

    Dim par As Recordset
    'ΒΡΙΣΚΩ ΑΠΟ ΤΙΣ GEFYRES.DBF ΤΗΝ ΚΑΤΑΛΛΗΛΗ ΕΓΓΡΑΦΗ
    Set par = db.OpenRecordset("SELECT *FROM PARASTAT WHERE EIDOS='" + Left(g_atim, 1) + "';")
    Set gefyres = db.OpenRecordset("GEFYRES")
    If pist = 1 Then
        gefyres.Move par("GEF_P") - 1  ' RECORD ΓΕΦΥΡΏΣΕΩΝ ΕΠΙ ΠΙΣΤΩΣΕΙ
    Else
        gefyres.Move par("GEF_m") - 1  ' RECORD ΓΕΦΥΡΏΣΕΩΝ metritois
    End If
    ' Set gTim = db.OpenRecordset("select *from tim where atim='" + g_atim + "' and hme>=#" + Format(g_hme, "mm/dd/yyyy") + "#;")
    ok_gefyr = 1
    '
    If g_hme <= ghme_orist Then

        MsgBox " ΔΕΝ ΕΠΙΤΡΕΠΟΝΤΑΙ ΚΙΝΗΣΕΙΣ  ΣΕ ΟΡΙΣΤΙΚΟΠΟΙΗΜΕΝΕΣ ΕΓΓΡΑΦΕΣ"
        Exit Function
    End If

    If InStr("oOοΟ", gefyres("gefyrono")) Then
        ok_gefyr = 0
        Exit Function
    End If
    Dim EIDPAR



    eid_par = "  "
    If gefyres("mEID_PAR") <> "  " Then
        eid_par = Left(gefyres("mEID_PAR"), 5) + "-" + g_atim    ' Τ.πωλ-Τ00012Γ
    End If
    If gefyres("mmGRE") <> " " Then
        mgre = gefyres("mmGRE")    ' Σ
    End If
    m_ait = gefyres("parathrhs")    'π.χ.  ΠΩΛΗΣΕΙΣ
    '
    ''
    ''
    '''
    '''
    '''
    '  For K = 1 To 9
    '      a = LTrim(Str(K))
    '      kodiklog(K) = Gefyres("kod" + a)  ' 3000****
    '      typ(K) = Gefyres("typos" + a)
    '      xp(K) = Gefyres("xpi" + a)
    '  Next
    '
    '  m_ait = Gefyres("parathrhs") 'π.χ.  ΠΩΛΗΣΕΙΣ

    '*** ψαχνω για διπλοεγγραφές












    gvar = "Select *from temp where eidhme='1' and hme=#"
    gvar = gvar + Format(g_hme, "mm/dd/yyyy") + "#  and mid(apa,7,7)='"
    gvar = gvar + mID(eid_par, 7, 7) + "'"

    Set temp = dbL.OpenRecordset(gvar)

    If temp.RecordCount > 0 Then
        ' OK
    End If
    temp.MoveFirst



    ' ενημερώνω τα αρχείο λογαριασμών abat
    Do While Not temp.EOF    ' and  eidhme+dtos(hme)+left(gre,1)+apa =  '1' + dtos(g_hme) + mgre  + EID_PAR
        mmhkos = Len(Trim(temp("kod")))
        If temp("eidHME") = "3" Then    ' απογραφή
            MHNAS = "00"
        Else
            MHNAS = find_mhnas(temp("hme"))
        End If
        sql = "select *from abat WHERE "

        If mmhkos = gDIG_4 Then sql = sql + "(bat='4' and left(kod," + Str(gDIG_4) + ")='" + temp("kod") + "') or"
        If mmhkos >= gDIG_3 Then sql = sql + "(bat='3' and left(kod," + Str(gDIG_3) + ")='" + Left(temp("kod"), gDIG_3) + "') or "
        If mmhkos >= 4 Then sql = sql + "(bat='2' and left(kod,4)='" + Left(temp("kod"), 4) + "') or "
        If mmhkos >= 2 Then sql = sql + "(bat='1' and left(kod,2)='" + Left(temp("kod"), 2) + "') "

        'συγκεντρώνω τους λογαριασμους που με ενδιαφέρουν
        Set abat2 = dbL.OpenRecordset(sql)
        ' εαν λειπουν τους ανοιγω
        MHKOS(1) = 2: MHKOS(2) = 4: MHKOS(3) = gDIG_3: MHKOS(4) = gDIG_4
        If mmhkos = 2 Then lo = 1
        If mmhkos = 4 Then lo = 2
        If mmhkos = gDIG_3 Then lo = 3
        If mmhkos = gDIG_4 Then lo = 4
        If abat2.RecordCount < lo Then
            For K2 = 1 To lo: logar(K2) = 0: Next
            On Error Resume Next
            abat2.MoveFirst
            Do While Not abat2.EOF
                logar(Val(abat2("bat"))) = 1
                abat2.MoveNext
            Loop
            For K2 = 1 To lo
                If logar(K2) = 0 Then
                    abat2.AddNew
                    abat2("bat") = LTrim(Str(K2))
                    abat2("kod") = Left(temp("kod"), MHKOS(K2))
                    abat2.Update
                End If
            Next
        End If
        abat2.MoveFirst
        Do While Not abat2.EOF
            abat2.Edit
            If temp("xpi") = "Π" Then
                If IsNull(abat2("ppros" + MHNAS)) Then abat2("ppros" + MHNAS) = temp("xre") Else abat2("ppros" + MHNAS) = abat2("ppros" + MHNAS) + temp("xre")
            Else
                If IsNull(abat2("ppros" + MHNAS)) Then abat2("xpros" + MHNAS) = temp("xre") Else abat2("xpros" + MHNAS) = abat2("xpros" + MHNAS) + temp("xre")
            End If
            abat2.Update
            abat2.MoveNext
        Loop
        abat2.Close
        temp.MoveNext
    Loop



    temp.Close
    abat.Close


    ' ΔΙΑΓΡΑΦΩ ΤΙς ΕΓΓΡΑΦΕΣ
    gvar = "DELETE *from temp where eidhme='1' and hme=#"
    gvar = gvar + Format(g_hme, "mm/dd/yyyy") + "# and left(apa,13)='"
    gvar = gvar + eid_par + "'"
    dbL.Execute gvar
    'DELETE *from temp where eidhme='1' and hme=#12/27/2004# and left(apa,13)='.ΑΠ.Λ-L00145Α'
    dbL.Close
    '
    '
    '
    '
    '
End Function

Sub arn_enhm_pel()
    On Error GoTo 0
    '----------------------------------------------------------------------

    Dim X As String

    'Dim DB As New ADODB.Connection
    'Dim E As New ADODB.Recordset
    'Dim R As New ADODB.Recordset, K
    'DB.Open gDir

    Dim e As New ADODB.Recordset, R As New ADODB.Recordset, k
    Dim PEL As New ADODB.Recordset
    Dim f_pel
    'If gConnect = "Access" Then
    '  Set db = OpenDatabase(gDir, False, False)
    'Else
    '   Set db = OpenDatabase(gDir, False, False, gConnect)
    'End If

    '

    R.Open "select * from PARASTAT where EIDOS='" + Left(TIM.Recordset("atim"), 1) + "';", Gdb, adOpenForwardOnly, adLockReadOnly

    f_pel = R("pel")
    Dim f_matim
    f_matim = TIM.Recordset("atim")

    PEL.Open "SELECT *FROM PEL WHERE EIDOS='" + TIM.Recordset("EIDOS") + "' AND KOD='" + TIM.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic


    'Set R = db.OpenRecordset("egg") ' , DB, adOpenKeyset, adLockOptimistic

    Dim f_suma    '= Val(SynoloG.Caption)
    f_suma = TIM.Recordset("ajia")




    'PEL.Edit


    '**********************************************************************

    Dim TELYP As Single

    If TIM.Recordset("TRP") = "ΜΕ" Then
    Else
        If PEL("eidos") = "e" Then
            If f_pel = "1" Then
                TELYP = PEL("typ") - f_suma
            End If
            If f_pel = "2" Then
                TELYP = PEL("typ") + f_suma
            End If
        End If
        If PEL("eidos") = "r" Then
            If f_pel = "1" Then
                TELYP = PEL("typ") - f_suma
            End If
            If f_pel = "2" Then
                TELYP = PEL("typ") + f_suma
            End If
        End If

        Gdb.Execute "UPDATE PEL SET TYP=" + Str(TELYP) + " WHERE  EIDOS='" + TIM.Recordset("EIDOS") + "' AND KOD='" + TIM.Recordset("KOD") + "'"
    End If




    Dim sql

    sql = "delete from EGG where EIDOS='" + TIM.Recordset("EIDOS") + "' AND KOD='" + TIM.Recordset("KOD") + "' and ATIM='" + TIM.Recordset("atim") + "'"
    Gdb.Execute sql

    R.Close

End Sub

Private Sub Form_DblClick()
    PARAMETROI.PARAM.Caption = "PAR2"
    PARAMETROI.SHOW 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Command4_Click


End Sub

Private Sub Form_Load()

    mForm_Load Me, fh, fw, ft, fl

    Me.Picture = LoadPicture(gPicture)
    Me.KeyPreview = True
    DTPicker1.Value = Now
    DTPicker2.Value = Now


End Sub

Private Sub Label4_Click()

End Sub

Private Sub MSFlexGrid1_Click()
    Dim F
    '  f = MSFlexGrid1.Text
    '
    '
    '
    '' kartella
    '  apot21.Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)
    '
    '  apot21.Text2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)
    '
    '  apot21.Label1.Caption = "1" ' Left(APOTH1.Text, 1)
    '
    '  apot21.Label2.Caption = "0"
    '
    '
    '
    '  apot21.SHOW
    '



















End Sub
Sub SOKRATIS_etik()

End Sub

Private Sub MSFlexGrid1_DblClick()
    Dim F
    F = MSFlexGrid1.Text



    ' kartella
    apot21.Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)

    apot21.Text2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)

    apot21.Label1.Caption = "1"    ' Left(APOTH1.Text, 1)

    apot21.Label2.Caption = "0"



    apot21.SHOW




End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then

        MSFlexGrid1.Col = 0
        ' APOT2.kod.Text = MSFlexGrid1.Text   ' "10"


        APOT2.kod.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)

        APOT2.SHOW
        APOT2.kod.SetFocus
        SendKeys "{ENTER}"
    End If





End Sub
