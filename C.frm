VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form pelat5 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   10305
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5520
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
      Caption         =   "ÁðïèÞêåõóç Report"
      Height          =   435
      Left            =   2910
      TabIndex        =   14
      Top             =   8025
      Width           =   2310
   End
   Begin VB.CheckBox rCheck 
      Caption         =   "×ñçóéìïðïéþ ôï åðéëåãìåíï report"
      Height          =   435
      Left            =   2910
      TabIndex        =   13
      Top             =   7515
      Width           =   2310
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   135
      TabIndex        =   12
      Top             =   7500
      Width           =   2715
   End
   Begin VB.TextBox Text2 
      Height          =   1425
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   6045
      Width           =   5085
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7365
      TabIndex        =   9
      Top             =   6060
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "åêôýðùóç"
      Height          =   285
      Left            =   3435
      TabIndex        =   8
      Top             =   810
      Width           =   2235
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Bindings        =   "C.frx":0000
      Height          =   4815
      Left            =   45
      TabIndex        =   6
      Top             =   1200
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSAdodcLib.Adodc EGG 
      Height          =   375
      Left            =   -45
      Top             =   10740
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   375
      Left            =   120
      Top             =   7680
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
         Size            =   8,25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   117702657
      CurrentDate     =   38358
   End
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   -120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Õðïëïãéóìüò"
      Height          =   615
      Left            =   6405
      TabIndex        =   1
      Top             =   7845
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "C.frx":0014
      Left            =   240
      List            =   "C.frx":001E
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   117702657
      CurrentDate     =   38358
   End
   Begin lvButton.lvButtons_H command2 
      Height          =   615
      Left            =   8040
      TabIndex        =   15
      Top             =   7800
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1085
      Caption         =   "Åîïäïò"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   9069372
      cBhover         =   13016952
      LockHover       =   3
      cGradient       =   -2147483628
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483629
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Aéôéïëïãßá"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   6045
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Åùò"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Áðü"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "pelat5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim F_RUN As Integer
'
Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub Command1_Click()
'
    Dim X As String, order As String
    Dim k As Integer, R As New ADODB.Recordset, sum As Single


    'ÂÑÉÓÊÙ ÔÏÕÓ ËÏÃÁÑÉÁÓÌÏÕÓ ÐÏÕ ÁÍÁÖÅÑÏÍÔÁÉ ÓÔÏ ÔÁÌÅÉÏ,ÐÉÓÔÙÔÉÊÅÓ ÃÉÁ ÍÁ ÔÏÕÓ ÅÎÁÉÑÅÓÙ ÁÐÏ ÔÉÓ ÊÉÍÇÓÅÉÓ
    Dim TAMEIA As String
    TAMEIA = ""
    R.Open "SELECT *FROM PINAKES WHERE TYPOS=12 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
    Do While Not R.EOF
        If R("typos") = 12 Then
            TAMEIA = TAMEIA + "'" + Format(R("TIMH"), "0000") + "',"
        End If
        R.MoveNext
    Loop
    TAMEIA = Left(TAMEIA, Len(TAMEIA) - 1)
    R.Close







    F_RUN = 1


    Dim TT
    TT = GetCurrentTime()

    Dim FILTRO As String


    If Len(Trim(Text1.Text)) > 0 Then
        FILTRO = " and EGG.AIT LIKE '" + Text1.Text + "%' "
    Else
        FILTRO = ""
    End If


    'If gXEIRISTHS = 2 Then
    '
    '
    '  FILTRO = " and  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%' " + FILTRO
    '
    '
    '
    'End If

    'CONVERT(CHAR(8),HME,3)
    X = "select CONVERT(CHAR(8),HME,3),EGG.ATIM,EGG.AIT,ROUND(EGG.XREOSI,2) as [×ñÝùóç],ROUND(EGG.PISTOSI,2) as [Ðßóôùóç],PEL.EPO,EGG.USERID   from EGG "
    'x = "select egg.hme,egg.atim,egg.ait,iif(egg.xre>0,format(egg.XRE,'###,###.00'),'') as [×ñÝùóç],iif(egg.xre<0,format(-egg.XRE,'###,###.00'),'') as [Ðßóôùóç],kod,eidos  into dok5  in 'c:\mercvb\reports\reports.mdb' from egg "
    X = X + " inner join PEL on EGG.EIDOS=PEL.EIDOS and EGG.KOD=PEL.KOD "

    'x = x + " inner join pel on egg.eidos+egg.kod=pel.eidos+pel.kod  "
    X = X + " where EGG.KOD NOT IN (" + TAMEIA + ") AND EGG.HME>='" + Format(DTPicker1.value, "mm/dd/yyyy") + _
        "'  and EGG.HME<'" + Format(DateAdd("d", 1, DTPicker2.value), "mm/dd/yyyy") + _
        "' and EGG.EIDOS='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "' " + FILTRO + " order by ID"




    Dim DUM, db As Database, dbacc As Database
    Dim NX, NP
    NX = 3: NP = 4


    If rCheck.value = vbChecked Then
        X = PutDates2(Text2.Text)
    End If


    'dum = OpenD(db)

    DATA1.ConnectionString = gConnect
    'Data1.DatabaseName = gDir
    DATA1.RecordSource = X    ' "egg"

    DATA1.Refresh
    Text2.Text = X


    'On Error Resume Next
    '
    'Set dbacc = OpenDatabase("c:\mercvb\reports\reports.mdb")
    'dbacc.Execute "drop table dok5"
    'On Error GoTo 0
    '
    ''Kill gDir + "\temp.dbf"
    '
    'db.Execute x
    '
    'dbacc.Execute "drop table pel5"
    'db.Execute "select eidos,kod,epo into pel5 in 'c:\mercvb\reports\reports.mdb' from pel where eidos='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "'"
    '
    'Data1.DatabaseName = "c:\mercvb\reports\reports.mdb"
    'Data1.Connect = "Access"
    'Data1.RecordSource = "select dok5.*,pel5.epo from dok5 inner join pel5  on dok5.eidos+dok5.kod=pel5.eidos+pel5.kod"
    'Data1.Refresh
    '
    MDIForm1.Caption = GetCurrentTime() - TT


    'Me.Caption = "--"
    'Set R = Data1.Recordset
    On Error Resume Next
    'R.MoveFirst

    Dim mc, sum2
    mc = 0

    sum = 0
    sum2 = 0
    For mc = 1 To MSFlexGrid1.ROWS - 1
        sum = sum + mval(MSFlexGrid1.TextMatrix(mc, 3))
        sum2 = sum2 + mval(MSFlexGrid1.TextMatrix(mc, 4))
        If mc Mod 2 = 0 Then
            MSFlexGrid1.row = mc
            For k = 0 To MSFlexGrid1.ColS - 1: MSFlexGrid1.Col = k: MSFlexGrid1.CellBackColor = &HC0E0FF: Next
        End If
        MSFlexGrid1.TextMatrix(mc, 0) = Format(MSFlexGrid1.TextMatrix(mc, 0), "dd/mm/yyyy")


        MSFlexGrid1.TextMatrix(mc, NX) = Format(Val(MSFlexGrid1.TextMatrix(mc, NX)), "######.##")
        MSFlexGrid1.TextMatrix(mc, NP) = Format(Val(MSFlexGrid1.TextMatrix(mc, NP)), "######.##")



        MSFlexGrid1.TextMatrix(mc, 0) = Format(MSFlexGrid1.TextMatrix(mc, 0), "dd/mm/yyyy")

    Next

    MSFlexGrid1.ColWidth(5) = 2000
    MSFlexGrid1.ColWidth(4) = 1200    'ÁÐÁ
    MSFlexGrid1.ColWidth(3) = 1200    'ÁÐÁ
    MSFlexGrid1.ColWidth(2) = 1600    'ÅÉÄ



    MSFlexGrid1.AddItem ("Óýíïëá" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(sum, "###,###.00")) & Chr(9) & Format(sum2, "###,###.00")



    MSFlexGrid1.TopRow = MSFlexGrid1.ROWS - 10
    MSFlexGrid1.ColAlignment(4) = 6    ' RIGHT

    MSFlexGrid1.ColAlignment(3) = 6    ' RIGHT
    'MSFlexGrid1.ColAlignment(7) = 6 ' RIGHT





    'DBGrid3.Columns(2).Caption = Format(sum, "###,###.00")
    'DBGrid3.ReBind
    'DBGrid3.Columns(0).Caption = Format(mc, "###,###")


    'If Check1 Then typos2








    If Check1 Then

        Dim PELA
        PELA = IIf(Combo1.Text = Combo1.List(0), "ÐÅËÁÔÅÓ", "ÐÑÏÌÇÈÅÕÔÅÓ")
        On Error GoTo 0    ' Resume Next
        print3_xar X, "000110", PELA, 0    'typos2
    End If








End Sub

Function PutDates2(ByVal X As String)

    Dim X1 As String, X2 As String

    X1 = "'" + Format(DTPicker1.value, "MM/DD/YYYY") + "'"
    X2 = "'" + Format(DateAdd("H", 23.99, DTPicker2.value), "MM/DD/YYYY") + "'"

    Dim Sql2 As String
    Sql2 = X    ' TEXT2.Text

    Sql2 = Replace(Sql2, "@X1", X1)
    Sql2 = Replace(Sql2, "@X2", X2)

    PutDates2 = Sql2

End Function




Private Sub Typose()
End Sub
Sub typos2()

End Sub



Private Sub Command11_Click()

    If Len(Dir("c:\mercvb\queries\pelKIN", vbDirectory)) = 0 Then
        MkDir "c:\mercvb\queries\pelKIN"
    End If



    CD1.InitDir = "c:\mercvb\queries\pel"
    CD1.ShowSave

    Dim F
    F = CD1.FILENAME

    Open F For Output As #5
    Print #5, Text2.Text
    Close #5


End Sub

Private Sub Command2_Click()
    Unload Me
    If gApoMenu = True Then
        SendKeys "%"
        SendKeys "{ENTER}+{DOWN}+{DOWN}+{DOWN}+{DOWN} "
    End If

End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

End Sub

Private Sub File1_Click()
    Dim a
    a = 0

    Dim F As String

    F = File1.FILENAME


    Dim ss As String
    Dim b As String

    ss = ""
    Open File1.Path + "\" + F For Input As #1
    Do While Not EOF(1)

        Line Input #1, b
        ss = ss + b + " "
    Loop
    Close #1


    Text2.Text = ss


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Command2_Click



End Sub

Private Sub Form_Load()
    mForm_Load Me, fh, fw, ft, fl

    Me.KeyPreview = True

    'If gUserId = 2 Then
    ' Command1.Enabled = False
    '
    'End If


    'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
    ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
    'SkinFramework1.ApplyWindow Me.hWnd
    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics



    Combo1.Text = Combo1.List(0)
    'Combo2.Text = Combo2.List(1)
    F_RUN = 0


    Me.Picture = LoadPicture(gPicture)

    DTPicker1.value = Now
    DTPicker2.value = Now

    If Len(Dir("c:\mercvb\queries\pel", vbDirectory)) = 0 Then
        MkDir "c:\mercvb\queries\pel"
    End If

    File1.Path = "c:\mercvb\queries\pel"


    CD1.InitDir = "c:\mercvb\queries\pel"






End Sub

Function mval(S As String)
    Dim k, L
    For k = 1 To Len(S)
        L = InStr(S, ",")
        If L > 0 Then
            Mid$(S, L, 1) = " "
        End If
    Next
    mval = Val(S)

End Function

Private Sub Form_Resize()
    mForm_Resize Me, 10545, 9585, fh, fw, ft, fl
End Sub
