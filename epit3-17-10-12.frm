VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form epit3 
   BackColor       =   &H00404080&
   ClientHeight    =   8832
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   15144
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8832
   ScaleWidth      =   15144
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "epit3-17-10-12.frx":0000
      Left            =   7980
      List            =   "epit3-17-10-12.frx":0002
      TabIndex        =   24
      Top             =   5985
      Width           =   2805
   End
   Begin VB.ComboBox PLAFON 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "epit3-17-10-12.frx":0004
      Left            =   1590
      List            =   "epit3-17-10-12.frx":0006
      TabIndex        =   22
      Top             =   7080
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "epit3-17-10-12.frx":0008
      Left            =   5760
      List            =   "epit3-17-10-12.frx":0012
      TabIndex        =   21
      Text            =   "1.«Ã.À«Œ≈Ÿ”"
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1590
      TabIndex        =   16
      Top             =   6540
      Width           =   1635
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "epit3-17-10-12.frx":0033
      Left            =   1590
      List            =   "epit3-17-10-12.frx":0035
      TabIndex        =   11
      Top             =   6015
      Width           =   1695
   End
   Begin VB.CheckBox katamhna 
      BackColor       =   &H00FF0000&
      Caption         =   "”ı„Í.Í·Ù· ÏﬁÌ·"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8055
      TabIndex        =   10
      Top             =   5220
      Width           =   180
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1590
      TabIndex        =   8
      Top             =   5610
      Width           =   1635
   End
   Begin VB.CheckBox ekt 
      BackColor       =   &H00FF0000&
      Caption         =   "≈ÍÙ˝˘ÛÁ"
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8055
      TabIndex        =   5
      Top             =   5520
      Width           =   195
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Bindings        =   "epit3-17-10-12.frx":0037
      Height          =   4635
      Left            =   225
      TabIndex        =   4
      Top             =   120
      Width           =   14415
      _ExtentX        =   25421
      _ExtentY        =   8170
      _Version        =   393216
      BackColor       =   14088709
      Cols            =   10
      FixedCols       =   0
      RowHeightMin    =   400
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   1080
      Top             =   8160
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5525
      _ExtentY        =   656
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
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
      Caption         =   "”ı„ÍÂÌÙÒ˘ÙÈÍ‹"
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8055
      TabIndex        =   3
      Top             =   4935
      Width           =   195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   615
      Left            =   9465
      TabIndex        =   2
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "–ÒÔ‚ÔÎﬁ"
      Height          =   615
      Left            =   4530
      TabIndex        =   1
      Top             =   6960
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "epit3-17-10-12.frx":004B
      Left            =   4485
      List            =   "epit3-17-10-12.frx":0067
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5505
      Width           =   3180
   End
   Begin MSComCtl2.DTPicker d1 
      Height          =   345
      Left            =   1185
      TabIndex        =   6
      Top             =   4950
      Width           =   1605
      _ExtentX        =   2836
      _ExtentY        =   614
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   117309441
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker d2 
      Height          =   345
      Left            =   3435
      TabIndex        =   7
      Top             =   4905
      Width           =   1605
      _ExtentX        =   2836
      _ExtentY        =   614
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   117309441
      CurrentDate     =   38814
   End
   Begin VB.Label PlafonText 
      BackStyle       =   0  'Transparent
      Caption         =   "ƒ·ÌÂÈ·Í.ÀÔ„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   7110
      Width           =   1050
   End
   Begin VB.Label Label9 
      Caption         =   "≈–ŸÕ.–—œÃ«»≈’‘«"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   7800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000004&
      Height          =   1080
      Left            =   7980
      Top             =   4860
      Width           =   3435
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000004&
      Height          =   2115
      Left            =   225
      Top             =   5490
      Width           =   4170
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   585
      Left            =   225
      Top             =   4875
      Width           =   5520
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "≈˘Ú"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2955
      TabIndex        =   19
      Top             =   4935
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "¡¸"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   630
      TabIndex        =   18
      Top             =   4965
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "¡ÒÈË.≈ÈÙ·„ﬁÚ"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   435
      TabIndex        =   17
      Top             =   6555
      Width           =   1170
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "≈ÍÙ˝˘ÛÁ"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   8280
      TabIndex        =   15
      Top             =   5490
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "”ı„Í.Í·Ù· ÏﬁÌ·"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   8310
      TabIndex        =   14
      Top             =   5205
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "”ı„ÍÂÌÙÒ˘ÙÈÍ‹"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   8295
      TabIndex        =   13
      Top             =   4905
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "‘Ò‹ÂÊ·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   6045
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "≈˘ÌıÏﬂ·"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   495
      TabIndex        =   9
      Top             =   5625
      Width           =   930
   End
End
Attribute VB_Name = "epit3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim F_PINAKIO As Integer
Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub Combo1_Click()
    Command1_Click

End Sub

Private Sub Command1_Click()

    Dim synt As String
    Dim eis As String
    Dim pl As String
    Dim Okeis As String
    Dim Okpl As String
    Dim platos, k
    Dim NEIS, NPL, nOKeIS, NOKPL, NARX
    Dim L As Integer
    Dim sql




For k = 1 To MSFlexGrid1.ROWS - 1
    MSFlexGrid1.RowHeight(k) = 900
Next
 MSFlexGrid1.RowHeightMin = 400

    MSFlexGrid1.ColWidth(0) = 2500  'hme
    MSFlexGrid1.ColWidth(1) = 3800  'epo

    MSFlexGrid1.ColWidth(2) = 1200  'epo
    MSFlexGrid1.ColWidth(3) = 1200  'epo




    If Check1 Or katamhna Then
        sygk
        Exit Sub
    End If
    ' Data1.DatabaseName = gDir
    DATA1.ConnectionString = gConnect

    MSFlexGrid1.Visible = False
    Dim SUgm_str


    If Combo1.ListIndex = 2 Then    '-----------  OLES MAZI

        MSFlexGrid1.ColWidth(0) = 2100  'hme
        MSFlexGrid1.ColWidth(3) = 1500  'hme
        MSFlexGrid1.ColWidth(4) = 1500  'hme

        ' MSFlexGrid1.Font.Name = "Arial"

        If Left(Combo3.Text, 1) = "1" Then    '«Ã≈—.À«Œ≈Ÿ”
            synt = "HMEL>='" + Format(d1, "mm/dd/yy") + "' and HMEL<='" + Format(d2, "mm/dd/yy") + "' AND ARX<3"
        Else
            synt = "HMEE>='" + Format(d1, "mm/dd/yy") + "' and HMEE<='" + Format(d2, "mm/dd/yy") + "' AND ARX<3"
        End If




        NEIS = 2: eis = "STR((CASE ARX WHEN 1 THEN POS ELSE 0 END),10,2) as [≈ﬂÛÒ·ÓÁ],"

        NPL = 3: pl = "STR((CASE ARX WHEN 2 THEN POS ELSE 0 END),10,2)  as [–ÎÁÒ˘Ïﬁ],"
        nOKeIS = 4: Okeis = "STR((CASE ARX WHEN 5 THEN POS ELSE 0 END),10,2) as [≈ÈÛÒ·˜ËÂﬂÛÂÚ],"
        NOKPL = 5: Okpl = "STR((CASE ARX WHEN 6 THEN POS ELSE 0 END),10,2)  as [–ÎÁÒ˘Ï›ÌÂÚ]"
        NARX = 6
        On Error GoTo 0
        sql = "select HMEL AS [«Ï_ÀﬁÓÂ˘Ú],EPO as [≈˘ÌıÏﬂ·]," + eis + pl + Okeis + Okpl
        sql = sql + ", CONVERT(CHAR(12),HMEE,3) as [HÏ.≈ÈÛ·„],ARX,ONO AS [Ã≈‘¡¬…¬] from GRA where " + synt + " order by HMEL;"
        DATA1.RecordSource = sql
        DATA1.Refresh
        SUgm_str = "001111"

        MSFlexGrid1.ColS = 8
        For k = 1 To MSFlexGrid1.ROWS - 1
            MSFlexGrid1.TextMatrix(k, 0) = Format(k, "00") + "." + Format(MSFlexGrid1.TextMatrix(k, 0), "DDD DD/MM/YYYY")

            '      If Val(MSFlexGrid1.TextMatrix(k, NARX)) = 1 Then
            '          MSFlexGrid1.TextMatrix(k, NPL) = ""
            '          MSFlexGrid1.TextMatrix(k, NOKPL) = ""
            '          MSFlexGrid1.TextMatrix(k, nOKeIS) = ""
            '       End If
            '      If Val(MSFlexGrid1.TextMatrix(k, NARX)) = 2 Then
            '          MSFlexGrid1.TextMatrix(k, NEIS) = ""
            '          MSFlexGrid1.TextMatrix(k, NOKPL) = ""
            '          MSFlexGrid1.TextMatrix(k, nOKeIS) = ""
            '       End If
            '
            '      If Val(MSFlexGrid1.TextMatrix(k, NARX)) = 4 Then
            '          MSFlexGrid1.TextMatrix(k, NEIS) = ""
            '          MSFlexGrid1.TextMatrix(k, NOKPL) = ""
            '          MSFlexGrid1.TextMatrix(k, nOKeIS) = ""
            '       End If
            '
            '      If Val(MSFlexGrid1.TextMatrix(k, NARX)) = 3 Then
            '          MSFlexGrid1.TextMatrix(k, NEIS) = ""
            '          MSFlexGrid1.TextMatrix(k, NOKPL) = ""
            '          MSFlexGrid1.TextMatrix(k, nOKeIS) = ""
            '       End If




            '      If Val(MSFlexGrid1.TextMatrix(k, NARX)) = 5 Then
            '          MSFlexGrid1.TextMatrix(k, NEIS) = ""
            '          MSFlexGrid1.TextMatrix(k, NOKPL) = ""
            '          MSFlexGrid1.TextMatrix(k, NPL) = ""
            '       End If
            '      If Val(MSFlexGrid1.TextMatrix(k, NARX)) = 6 Then
            '          MSFlexGrid1.TextMatrix(k, NEIS) = ""
            '          MSFlexGrid1.TextMatrix(k, NOKPL) = ""
            '          MSFlexGrid1.TextMatrix(k, NPL) = ""
            '       End If


            If k Mod 2 = 0 Then
                For L = 0 To MSFlexGrid1.ColS - 1
                    MSFlexGrid1.row = k: MSFlexGrid1.Col = L: MSFlexGrid1.CellBackColor = vbYellow    '&HC0E0FF
                Next
            End If


        Next



        MSFlexGrid1.ColAlignment(4) = 7    'right
        MSFlexGrid1.ColAlignment(2) = 7    'right
        MSFlexGrid1.ColAlignment(3) = 7    'right
        MSFlexGrid1.ColAlignment(5) = 7    'right

        eis = Format(Calc_Suma(NEIS), "###,###.00")
        pl = Format(Calc_Suma(NPL), "###,###.00")

        Okeis = Format(Calc_Suma(nOKeIS), "###,###.00")
        Okpl = Format(Calc_Suma(NOKPL), "###,###.00")

        ' ÒÔÛË›Ù˘ ÛÂÈÒ‹ „È· Ù· Û˝ÌÔÎ·

        MSFlexGrid1.AddItem ("" & Chr(9) & "”˝ÌÔÎÔ" & Chr(9) & eis & Chr(9) & pl & Chr(9) & Okeis & Chr(9) + Okpl)

        MSFlexGrid1.AddItem ("" & Chr(9) & "ƒÈ·ˆÔÒ‹" & Chr(9) & Format(Calc_Suma(NEIS) - Calc_Suma(NPL), "###,###.00"))


        ' ÍÔÍÍÈÌÈÊ˘ Ù· Û˝ÌÔÎ·

        MSFlexGrid1.row = MSFlexGrid1.ROWS - 2
        MSFlexGrid1.Col = 0
        MSFlexGrid1.RowSel = MSFlexGrid1.ROWS - 1
        MSFlexGrid1.Col = 2
        MSFlexGrid1.CellBackColor = vbRed
        MSFlexGrid1.Col = 3
        MSFlexGrid1.CellBackColor = vbRed
        MSFlexGrid1.ColAlignment(3) = 7



    Else  '----------------------- ÔÎÂÛ Ï·ÊÈ

        Dim DD

        DD = Str(Combo1.ListIndex + 1)
        If Combo1.ListIndex = 7 Then    ' diagramenes epitages
            DD = "0"
        End If
        If Val(DD) = 5 Then
            DD = "3"
        End If




        ' ÂÎ·Ù˛Ì-ÒÔÏÁËÂıÙ˛Ì

        If Left(Combo3.Text, 1) = "2" Then    ' .«Ã.≈…”¡√Ÿ√«”
            sql = "select HMEL as [«Ï_ÀﬁÓÁÚ],EPO as [E˘ÌıÏﬂ·]," _
                & "str(POS,10,2) as [¡Óﬂ·],PAR AS [A— ≈–…‘],HMEE AS [«Ã.≈…”],(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=6 AND AYJON=KOD_BANK) AS [‘—¡–≈∆¡] ,PINAKIO AS [–…Õ¡ …œ],ONO+'  ' AS [Ã≈‘¡¬…¬] from GRA where HMEE>='" + Format(d1, "Mm/dd/yy") _
                + "' and HMEE<'" + Format(DateAdd("D", 1, d2), "mm/dd/yy") + _
                  "' and ARX=" + DD


        Else

            sql = "select HMEL as [«Ï_ÀﬁÓÁÚ],EPO as [E˘ÌıÏﬂ·]," _
                & "str(POS,10,2) as [¡Óﬂ·],PAR AS [A— ≈–…‘],HMEE AS [«Ã.≈…”],(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=6 AND AYJON=KOD_BANK) AS [‘—¡–≈∆¡] ,PINAKIO AS [–…Õ¡ …œ],ONO+'  ' AS [Ã≈‘¡¬…¬] from GRA where HMEL>='" + Format(d1, "Mm/dd/yy") _
                + "' and HMEL<='" + Format(DateAdd("D", 1, d2), "mm/dd/yy") + _
                  "' and ARX=" + DD
        End If











        If Len(Text1.Text) > 0 Then
            sql = sql + " and EPO LIKE '" + Text1.Text + "%' "
        End If



        If Val(Combo1.ListIndex) < 2 Then  '


            If Val(Combo2.Text) > 0 Then    'TRAPEZA
                sql = sql + " AND KOD_BANK LIKE '%" + Format(Val(Combo2.Text), "#0") + "%'"
            End If

        End If


        '  If Val(Combo2.Text) = 2 Then  'TRAPEZA


        If Combo1.ListIndex = 3 Then
            '    If Val(Left(Combo2.Text, 2)) > 0 Then
            sql = sql + " AND PINAKIO=" + Format(Val(Left(PLAFON.Text, 2)), "0")
            '   End If

        End If



        If Val(Text2.Text) > 0 Then  ' ariumow epitaghw
            sql = sql + " AND left(PAR," + Str(Len(Trim(Text2.Text))) + ") ='" + Trim(Text2.Text) + "'"
        End If


        If Left(Combo4.Text, 1) = "2" Then
            sql = sql + "  ORDER BY PAR"
        ElseIf Left(Combo4.Text, 1) = "3" Then
            sql = sql + "  ORDER BY ID"


        Else
            sql = sql + "  ORDER BY HMEL"
        End If




        SUgm_str = "0010"

        Dim nn
        Gdb.Execute "UPDATE GRA SET KOD_BANK='0' WHERE LEFT(KOD_BANK,1) NOT IN ('0','1','2','3','4','5','6','7','8','9')", nn



        DATA1.RecordSource = sql

        'MSFlexGrid1.ColAlignment(2) = 7
        DATA1.Refresh

        Dim X As Single
        X = 0
        For k = 1 To MSFlexGrid1.ROWS - 1
            MSFlexGrid1.TextMatrix(k, 0) = Format(k, "##") + ". " + Format(MSFlexGrid1.TextMatrix(k, 0), "DDD DD/MM/YYYY")
            MSFlexGrid1.TextMatrix(k, 4) = Format(MSFlexGrid1.TextMatrix(k, 4), "DD/MM/YYYY")
            X = X + Val(MSFlexGrid1.TextMatrix(k, 2))
            MSFlexGrid1.TextMatrix(k, 2) = Format(MSFlexGrid1.TextMatrix(k, 2), "###,###.00")

            If k Mod 2 = 0 Then
                For L = 0 To MSFlexGrid1.ColS - 1
                    MSFlexGrid1.row = k: MSFlexGrid1.Col = L: MSFlexGrid1.CellBackColor = &HC0E0FF    '&HC0E0FF
                Next
            End If
            MSFlexGrid1.TextMatrix(k, 7) = k

        Next

        MSFlexGrid1.AddItem ("”˝ÌÔÎ·" & Chr(9) & "" & Chr(9) & Format(X, "###,###.00"))
        MSFlexGrid1.Col = 2
        MSFlexGrid1.row = MSFlexGrid1.ROWS - 1
        MSFlexGrid1.CellBackColor = vbRed

        MSFlexGrid1.ColAlignment(2) = 7    'de??? st????s?

    End If

    platos = 1000

    platos = 0
    For k = 0 To MSFlexGrid1.ColS - 1
        platos = platos + MSFlexGrid1.ColWidth(k)
    Next

    '  MSFlexGrid1.width = K * 460 + platos

    MSFlexGrid1.width = 11175
    For k = 0 To MSFlexGrid1.ColS - 1
        MSFlexGrid1.ColWidth(k) = MSFlexGrid1.ColWidth(k) * (MSFlexGrid1.width / platos) * 0.75 + 100
    Next



    If MSFlexGrid1.ROWS > 7 Then
        MSFlexGrid1.TopRow = MSFlexGrid1.ROWS - 7
    End If



    MSFlexGrid1.ColWidth(0) = 2100  '«Ã≈—
    MSFlexGrid1.ColWidth(2) = 2000  'epo
    MSFlexGrid1.ColWidth(3) = 2000  '¡—.≈–…‘
    MSFlexGrid1.ColWidth(4) = 1400  '‘—¡–≈∆¡
    MSFlexGrid1.ColWidth(5) = 1400  '
    MSFlexGrid1.ColWidth(6) = 2000  'epo
    MSFlexGrid1.ColWidth(7) = 1400  '








    MSFlexGrid1.Visible = True

    'MSFlexGrid1.Refresh

    If ekt Then
        'typos2
        print3_xar UCase(sql), SUgm_str, "≈ÈÙ·„›Ú " + Combo1.Text + " ·¸ " + Format(d1, "dd/mm/yyyy") + " ›˘Ú " + Format(d2, "dd/mm/yyyy"), 0
    End If


    'If Excel Then
    '  Print4_Excel UCase(sql), "111111111111111"
    'End If







End Sub

Sub sygk()
'********************************  ”’√ ≈Õ‘—Ÿ‘… ¡ *********************************
    Dim synt As String
    Dim eis As String
    Dim pl As String
    Dim Okeis As String
    Dim Okpl As String
    Dim platos, k
    Dim db As Database, db2 As Database
    Dim X As Single
    Dim X2 As Single
    Dim sql As String

    MSFlexGrid1.ColWidth(0) = 1500  'hme
    MSFlexGrid1.ColWidth(1) = 1000  'epo

    If Combo1.ListIndex = 2 Then


        synt = " HMEL>='" + Format(d1, "mm/dd/yy") + "' and HMEL<='" + Format(d2, "mm/dd/yy") + "'"

        eis = " ( CASE ARX WHEN 1 THEN POS ELSE 0  END) AS EIS , "  '"iif(arx=1,pos,0) as eis ,"
        pl = " ( CASE  WHEN ARX=2 THEN POS ELSE 0  END) AS PL , "  ' "POS as PL,"   ' "iif(arx=2,pos,0) as pl,"
        Okeis = " ( CASE ARX WHEN 6 THEN POS ELSE 0  END) AS HDHE , "    ' "POS as HDHE,"   '"iif(arx=6,pos,0) as hdhe,"
        Okpl = " ( CASE ARX WHEN 7 THEN POS ELSE 0  END) AS HDHP  "    ' "POS as HDHP" ' "iif(arx=7,pos,0) as hdhp"

        'Set db = OpenDatabase(gDir, False, False, gConnect)

        Set db2 = OpenDatabase("c:\mercvb\reports\reports.mdb", False, False)
        On Error Resume Next
        Gdb.Execute "drop table EPIT"



        On Error GoTo 0
        Gdb.Execute "select HMEL," + eis + pl + Okeis + Okpl + ",ARX into EPIT from GRA WHERE " + synt + " order by HMEL;"

        ' Gdb.Execute "insert into epit in 'c:\mercvb\reports\reports.mdb' select hmel," + eis + pl + Okeis + Okpl + " from gra where " + synt + " order by hmel;"

        'Data1.DatabaseName = "c:\mercvb\reports\reports.mdb"
        DATA1.ConnectionString = gConnect


        If katamhna Then
            sql = "select month(HMEL) as [ÏÁÌ_ÀﬁÓÁÚ] ,STR(sum(EIS),10,2) as [≈ÈÛÒ‹ÓÂÈÚ],STR(sum(PL),10,2) as [–ÎÁÒ˘Ï›Ú],ARX from EPIT where " + synt + " group by ARX,month(HMEL) HAVING SUM(EIS)>0 OR SUM(PL)>0 order by month(HMEL);"
        Else
            sql = "select HMEL as [«ÏÂÒ_ÀﬁÓÁÚ] ,STR(sum(EIS),10,2) as [≈ÈÛÒ‹ÓÂÈÚ],STR(sum(PL),10,2) as [–ÎÁÒ˘Ï›Ú],ARX from EPIT where " + synt + " group by HMEL,ARX order by HMEL;"
        End If

        DATA1.RecordSource = sql
        DATA1.Refresh
        'MSFlexGrid1.ColAlignment(4) = 7 'right
        MSFlexGrid1.ColAlignment(1) = 7    'right
        MSFlexGrid1.ColAlignment(2) = 7    'right




        X = 0
        For k = 1 To MSFlexGrid1.ROWS - 1
            If MSFlexGrid1.TextMatrix(k, 3) = 2 Then MSFlexGrid1.TextMatrix(k, 1) = ""
            If MSFlexGrid1.TextMatrix(k, 3) = 1 Then MSFlexGrid1.TextMatrix(k, 2) = ""


            If katamhna Then
            Else

                '  MSFlexGrid1.TextMatrix(K, 0) = Format(MSFlexGrid1.TextMatrix(K, 0), "dd/mm/yyyy")

                MSFlexGrid1.TextMatrix(k, 0) = Format(MSFlexGrid1.TextMatrix(k, 0), "DDD DD/MM/YYYY")

            End If

            X = X + Val(MSFlexGrid1.TextMatrix(k, 1))
            X2 = X2 + Val(MSFlexGrid1.TextMatrix(k, 2))
            MSFlexGrid1.ColWidth(3) = 0



        Next

        MSFlexGrid1.AddItem "”˝ÌÔÎ·" & Chr(9) & Format(X, "###,###.00") & Chr(9) & Format(X2, "###,###.00")
        MSFlexGrid1.Col = 2
        MSFlexGrid1.row = MSFlexGrid1.ROWS - 1
        MSFlexGrid1.CellBackColor = vbRed











        'MSFlexGrid1.ColAlignment(5) = 7 'right

        'eis = Format(Calc_Suma(Data1.Recordset, "??sp???e??"), "###,###.00")
        'pl = Format(Calc_Suma(Data1.Recordset, "?????µ??"), "###,###.00")

        'Okeis = Format(Calc_Suma(Data1.Recordset, "?d???sp?a??e?se?"), "###,###.00")
        'Okpl = Format(Calc_Suma(Data1.Recordset, "?d??????µ??e?"), "###,###.00")

        ' p??s??t? se??? ??a ta s????a

        'MSFlexGrid1.AddItem ("" & Chr(9) & "S????a" & Chr(9) & eis & Chr(9) & pl & Chr(9) & Okeis & Chr(9) + Okpl)

        'MSFlexGrid1.AddItem ("" & Chr(9) & "??af???" & Chr(9) & Format(Calc_Suma(Data1.Recordset, "?????µ??") - Calc_Suma(Data1.Recordset, "??sp???e??"), "###,###.00"))


        ' ????????? ta s????a

        'MSFlexGrid1.Row = MSFlexGrid1.Rows - 2
        'MSFlexGrid1.Col = 0
        'MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
        'MSFlexGrid1.Col = 2
        'MSFlexGrid1.CellBackColor = vbRed
        'MSFlexGrid1.Col = 3
        'MSFlexGrid1.CellBackColor = vbRed
        'MSFlexGrid1.ColAlignment(3) = 7



    Else
        MSFlexGrid1.ColWidth(1) = 2000
        If katamhna Then
            sql = "select MONTH(HMEL) as [MHN_ÀﬁÓÁÚ],str(sum(POS),10,2) as [”’ÕœÀ¡] from GRA where HMEL>='" + Format(d1, "mm/dd/yy") + "' and HMEL<='" + Format(d2, "mm/dd/yy") + "' and ARX=" + Str(Combo1.ListIndex + 1) + " group by MONTH(HMEL) order by MONTH(HMEL)"
        Else
            sql = "select CONVERT(DATETIME,HMEL,3) as [«Ï_ÀﬁÓÁÚ],str(sum(POS),10,2) as [”’ÕœÀ¡] from GRA where HMEL>='" + Format(d1, "mm/dd/yy") + "' and HMEL<='" + Format(d2, "mm/dd/yy") + "' and ARX=" + Str(Combo1.ListIndex + 1) + " group by HMEL order by HMEL"
        End If


        DATA1.RecordSource = sql
        'MSFlexGrid1.ColAlignment(2) = 7
        DATA1.Refresh


        X = 0
        For k = 1 To MSFlexGrid1.ROWS - 1
            If katamhna Then
            Else
                MSFlexGrid1.TextMatrix(k, 0) = Format(MSFlexGrid1.TextMatrix(k, 0), "dd/mm/yyyy")
            End If

            X = X + Val(MSFlexGrid1.TextMatrix(k, 1))
        Next

        MSFlexGrid1.AddItem ("”˝ÌÔÎ·" & Chr(9) & Format(X, "###,###.00"))
        '  MSFlexGrid1.Col = 2
        MSFlexGrid1.row = MSFlexGrid1.ROWS - 1
        MSFlexGrid1.CellBackColor = vbRed

        MSFlexGrid1.ColAlignment(1) = 7    'de??? st????s?

    End If

    platos = 0
    For k = 0 To MSFlexGrid1.ColS - 1
        platos = platos + MSFlexGrid1.ColWidth(k)
    Next

    MSFlexGrid1.width = k * 150 + platos + 100    ' + 2720





    If MSFlexGrid1.ROWS > 7 Then
        MSFlexGrid1.TopRow = MSFlexGrid1.ROWS - 7
    End If



    If ekt Then
        'typos2
        print3_xar UCase(sql), "1111111111111111", "", 0
    End If


    'If Excel Then
    '  Print4_Excel UCase(sql), "111111111111111"
    'End If






    'MSFlexGrid1.Refresh

End Sub


Private Sub Command2_Click()
    Unload Me



    If gApoMenu Then
        SendKeys "%"
        SendKeys "{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}"
    End If


End Sub

Private Sub EPIT31_Click(Index As Integer)
    Dim a
    a = 0



End Sub

Sub epit311_Click(Index As Integer)

    Dim a As String, mc As Integer
    Dim k As Integer
    Dim R As New ADODB.Recordset
    Dim r2 As New ADODB.Recordset


    Dim kod As String
    Dim sql As String
    Dim b
    Dim M_TRAPEZA
    Dim KOD_TRAPEZAS As String
    Dim C As String
    Dim hmee As Date




    MSFlexGrid1.Col = 1
    b = MSFlexGrid1.Text




    MSFlexGrid1.Col = 3
    a = MSFlexGrid1.Text

    Dim hmel As String
    hmel = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)

    On Error GoTo MHNYMA

    If Index = 3 Then    ' Ã≈‘¡÷œ—¡ ≈–…‘¡√«”
        epit3200.SHOW 1
        k = MsgBox("Õ· ÏÂÙ·ˆÂÒËÂﬂ Á ÂÈÙ·„ﬁ " + Chr(13) + b + Chr(13) + a, vbYesNo)
        'Me.Caption = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        If k = vbYes Then
            Gdb.Execute "UPDATE GRA SET ONO='" + Left(Label9.Caption, 6) + Me.Caption + "',ARX=3 WHERE LEFT(PAR,10)='" + Left(a, 10) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", mc
        Else
            Exit Sub
        End If

        R.Open "SELECT * FROM GRA WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly




        '≈Õ«Ã≈—ŸÕŸ ‘œÕ ÒÔÏÁËÂıÙﬁ

        sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,XREOSI,PISTOSI,EID,USERID ) VALUES "
        sql = sql + "('" + Format(Now, "MM/dd/YYYY")
        sql = sql + "',-" + LTrim(R("pos")) + ",'"
        sql = sql + Left("≈–…‘." + R("EPO"), 30) + "','"
        sql = sql + Me.Caption + "','"
        sql = sql + "r" + "',"
        sql = sql + Format(Val(Left(R("PAR"), 9)), "##0000000") + ","

        sql = sql + Format(R("POS"), "000000.00") + ","
        sql = sql + "0,"

        sql = sql + "'E–',"
        sql = sql + Str(gUserId) + " )"
        k = 0
        Gdb.Execute sql, k
        '       Gdb.Execute sql, k
        '             Gdb.Execute sql, k

        '                   Gdb.Execute sql, k
        If k > 0 Then
            MsgBox "≈Õ«Ã≈—Ÿ»« ≈ «  ¡—‘≈À¡ ‘œ’ –—œÃ«»≈’‘«"
        Else
            MsgBox "–—œ”œ◊« !!! ƒ≈Õ ≈Õ«Ã≈—Ÿ»« ≈ «  ¡—‘≈À¡ ‘œ’ –—œÃ«»≈’‘«"
        End If
    End If




    If Index = 0 Then    ' –À«—ŸÃ« ≈–…‘¡√«”

        k = MsgBox("Õ· –À«—Ÿ»≈… Á ÂÈÙ·„ﬁ " + Chr(13) + b + Chr(13) + a, vbYesNo)
        Me.Caption = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        If k = vbYes Then
            Gdb.Execute "UPDATE GRA SET ARX=ARX+5 WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", mc
        End If

        R.Open "SELECT * FROM GRA WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly




        '≈Õ«Ã≈—ŸÕŸ  ¡… ‘œÕ ¡ÀÀ«Àœ◊—≈œ –…”‘ŸÕœÕ‘¡” ‘œÕ
        If Len(Dir("C:\KAGI.TXT")) > 0 And R("ARX") = 9 Then  ' –À«—ŸÃ≈Õ≈” ‘œ’ –À¡÷œÕ  ********************************
            '–À«—ŸÃ« ≈–…‘¡√«” ¡–O –À¡÷œÕ
            sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,XREOSI,PISTOSI,EID,USERID ) VALUES "
            sql = sql + "('" + Format(Now, "MM/dd/YYYY")
            'BRISKO ‘œÕ  Ÿƒ… œ ‘œ’ –…Õ¡ …œ’



            r2.Open "SELECT * FROM PINAKES WHERE TYPOS=18 AND AYJON=" + Str(R("PINAKIO")), Gdb, adOpenForwardOnly, adLockReadOnly

            KOD_TRAPEZAS = Format(r2("TIMH"), "0000")

            sql = sql + "',-" + LTrim(R("pos")) + ",'"
            sql = sql + Left("≈–…‘." + R("EPO"), 30) + "','"
            sql = sql + KOD_TRAPEZAS + "','"
            sql = sql + "e" + "',"
            'sql = sql + Left(R("PAR"), 9) + ","
            sql = sql + Format(Val(Left(R("PAR"), 9)), "##0000000") + ","

            sql = sql + "0,"
            sql = sql + Format(R("POS"), "000000.00") + ","
            sql = sql + "'E–…',"
            sql = sql + Str(gUserId) + " )"
            k = 0
            If Val(KOD_TRAPEZAS) > 0 Then
                Gdb.Execute sql, k
            End If
            If k > 0 Then
                MsgBox "≈Õ«Ã≈—Ÿ»« ≈ œ AÀÀ«Àœ◊—≈œ”"
            End If
        End If




        '≈Õ«Ã≈—ŸÕŸ  ¡… ‘œ œÿ≈Ÿ” –…”‘ŸÕœÕ‘¡” ‘œ
        If Len(Dir("C:\KAGI.TXT")) > 0 And R("ARX") = 7 Then  ' –À«—ŸÃ≈Õ≈” –—œÃ«»≈’‘ŸÕ  ********************************
            '–À«—ŸÃ« ≈–…‘¡√«” ¡–O œÿ≈Ÿ”
            '      R.Open "SELECT * FROM GRA WHERE LEFT(PAR,8)='" + Left(A, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
            sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,XREOSI,PISTOSI,EID,USERID ) VALUES "
            sql = sql + "('" + Format(Now, "MM/dd/YYYY")
            For k = 0 To Combo2.ListCount - 1
                If Val(Left(Combo2.List(k), 2)) = R("KOD_BANK") Then
                    KOD_TRAPEZAS = Right(Combo2.List(k), 4)
                End If
            Next
            sql = sql + "',-" + LTrim(R("pos")) + ",'"
            sql = sql + Left("≈–…‘." + R("EPO"), 30) + "','"
            sql = sql + KOD_TRAPEZAS + "','"
            sql = sql + "e" + "',"

            sql = sql + Format(Val(Left(R("PAR"), 9)), "##0000000") + ","
            ' sql = sql + Left(R("PAR"), 9) + ","
            sql = sql + "0,"
            sql = sql + Format(R("POS"), "000000.00") + ","
            sql = sql + "'E–…',"
            sql = sql + Str(gUserId) + " )"
            k = 0
            If Val(KOD_TRAPEZAS) > 0 Then
                Gdb.Execute sql, k
            End If
            If k > 0 Then
                MsgBox "≈Õ«Ã≈—Ÿ»« ≈ œ œÿ≈Ÿ”"
            End If
        End If


    End If

    If Index = 1 Then    ' ”≈ ‘—¡–≈∆¡
        ' On Error GoTo 0
        F_PINAKIO = 1

        epit3100.SHOW 1




        Gdb.Execute "UPDATE GRA SET ARX=4,PINAKIO=" + Left(Me.Caption, 2) + " WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", k


        If k > 0 Then
            MsgBox "ÃÂÙÂˆ›ÒËÁ"
        End If






        If Len(Dir("C:\KAGI.TXT")) > 0 Then    ' ********************************

            'ENHMERVSH ¡ÀÀ«Àœ◊—≈œ’ (–À¡÷œÕ) ‘—¡–≈∆¡”
            R.Open "SELECT * FROM GRA WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly


            sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,XREOSI,PISTOSI,EID,USERID ) VALUES "
            sql = sql + "('" + Format(Now, "MM/dd/YYYY")

            KOD_TRAPEZAS = Right(Me.Caption, 4)

            sql = sql + "',-" + LTrim(R("pos")) + ",'"
            sql = sql + Left("≈–…‘." + R("EPO"), 30) + "','"

            sql = sql + KOD_TRAPEZAS + "','"
            sql = sql + "e" + "',"
            sql = sql + Format(Val(Left(R("PAR"), 9)), "##0000000") + ","

            sql = sql + Format(R("POS"), "000000.00") + ","

            sql = sql + "0,"

            sql = sql + "'E–',"
            sql = sql + Str(gUserId) + " )"
            k = 0
            If Val(KOD_TRAPEZAS) > 0 And R("ARX") = 4 Then    ' SE PLAFON
                Gdb.Execute sql, k
            End If


            If k > 0 Then
                MsgBox "≈Õ«Ã≈—Ÿ»« ≈ œ ¡ÀÀ«Àœ◊—≈œ”"
            End If


        End If




        ' On Error GoTo 0




    End If






    If Index = 2 Then    ' ++++++++++++++ ƒ…¡√—¡÷« ≈–…‘¡√«”

        If Len(Dir("C:\KAGI.TXT")) > 0 Then    ' ********************************
            k = MsgBox("Õ· ‰È·„Ò·ˆÂﬂ Á ÂÈÙ·„ﬁ " + Chr(13) + b + Chr(13) + a, vbYesNo)
            Me.Caption = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
            If k = vbYes Then
                Gdb.Execute "UPDATE GRA SET ARX=0 WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", mc

                If mc > 0 Then    ' ------------
                    MsgBox "ƒ…≈√—¡÷« « ≈–…‘¡√«." + Chr(13) + "–—œ”œ◊« !!!!!  ‘¡ ‘œ–œ…«”‘≈ ‘«Õ  ¡—‘≈À¡ ‘œ’ –≈À¡‘«/–—œÃ«»≈’‘«"
                Else
                    MsgBox "ƒ≈Õ ƒ…≈√—¡÷« « ≈–…‘¡√«. "
                End If    '------------------

            End If
        Else    '' **************************************************************

            k = MsgBox("Õ· ‰È·„Ò·ˆÂﬂ Á ÂÈÙ·„ﬁ " + a, vbYesNo)
            If k = vbYes Then  ' \\\\\\\\\\\\\\\\\\\\\\\
                R.Open "SELECT *FROM GRA WHERE PAR='" + a + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                If Not R.EOF Then    ' =================
                    kod = R("KOD")
                    hmee = R("HMEE")

                    Dim RR As New ADODB.Recordset
                    RR.Open "SELECT count(*) FROM GRA WHERE PAR='" + a + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                    If RR(0) > 1 Then
                        MsgBox "ƒ≈Õ ƒ…≈√—¡÷« « ≈–…‘¡√«. "
                        Exit Sub
                    End If
                    RR.Close


                    '  Gdb.BeginTrans
                    Gdb.Execute "DELETE FROM GRA WHERE PAR='" + a + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", mc
                    If mc > 0 Then    '---------
                        sql = "DELETE FROM EGG WHERE KOD='" + kod + "' AND LEFT(AIT,7)='" + Left(a, 7) + "' AND HME ='" + Format(Right(hmee, 10), "MM/DD/YYYY") + "'"
                        Gdb.Execute sql, mc
                        If mc = 1 Then
                            '         Gdb.CommitTrans
                            MsgBox "ƒ…≈√—¡÷« « ≈–…‘¡√«  ¡… «  …Õ«”« ‘œ’ –≈À¡‘«"
                        Else
                            MsgBox "ƒ≈Õ ƒ…≈√—¡÷« «  …Õ«”« ‘œ’ Àœ√¡—…¡”Ãœ’ " + Chr(13) + " ƒ…¡√—¡ÿ‘≈ ‘«Õ ¡–œ ‘«Õ  …Õ«”«"
                            '       Gdb.RollbackTrans
                        End If
                    End If    ' -----------
                End If    ' ================
            End If    ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\

        End If    ' **********************************************************

    End If    ' ++++++++++++++ ƒ…¡√—¡÷« ≈–…‘¡√«”



    If Index = 5 Then    ' ++++++++++++++ EPANAFORA STA XERIA MAS



        If Combo1.ListIndex < 3 Then    ' –—œÃ«»≈’‘«Ú / PELATHS

            Dim mHMEL As Date
            Dim mPoso As Single
            Dim mParat As String

            'ENHMERVSH ¡ÀÀ«Àœ◊—≈œ’ (–À¡÷œÕ) ‘—¡–≈∆¡”
            R.Open "SELECT * FROM GRA WHERE LEFT(PAR,10)='" + Left(a, 10) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
            mHMEL = R("HMEL")
            mPoso = R("POS")
            mParat = R("PAR")


            Dim mID As Long
            mID = R("id")

            mHMEL = InputBox("Õ≈¡ «Ã≈—.À«Œ≈Ÿ” ", "≈–…‘¡√« " + a, CDate(mHMEL))
            mPoso = InputBox("Õ≈O –œ”œ ", "≈–…‘¡√« " + a, Str(mPoso))
            mParat = InputBox("Õ≈œ” ¡—…»Ãœ” ", "≈–…‘¡√« " + a, mParat)

            '      Gdb.Execute "UPDATE GRA SET HMEL='" + Format(Right(mHMEL, 10), "MM/DD/YYYY") + "'," _
                   '      & "POS=" + Str(mPoso) + ",PAR='" + mParat + "' WHERE ID=" + Str(mID)



            '         R("HMEL") = Format(mHMEL, "MM/DD/YYYY")
            '         R("POSO") = mPoso
            '         R("PARAT") = mParat
            '         R.Update




            If IsNull(R("id")) Then
                Gdb.Execute "UPDATE GRA SET HMEL='" + Format(Right(mHMEL, 10), "MM/DD/YYYY") + "'," _
                          & "POS=" + Str(mPoso) + ",PAR='" + mParat + "'   WHERE LEFT(PAR,10)='" + Left(a, 10) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'"
            Else
                Gdb.Execute "UPDATE GRA SET HMEL='" + Format(Right(mHMEL, 10), "MM/DD/YYYY") + "'," _
                          & "POS=" + Str(mPoso) + ",PAR='" + mParat + "' WHERE ID=" + Str(mID)
            End If


            R.Close


        Else    ' SE XERIA PROMHUEYTH

            k = MsgBox("Õ· ≈–¡Õ≈—»≈… ”‘¡ ◊≈—…¡ Ã¡”  Á ÂÈÙ·„ﬁ " + a, vbYesNo)
            If k = vbYes Then  ' \\\\\\\\\\\\\\\\\\\\\\\
                Gdb.Execute "UPDATE GRA  SET ARX=1 WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", mc
                If mc = 1 Then
                    MsgBox "ok Â·ÌﬁÒËÂ"
                Else
                    MsgBox "–ÒÔÛÔ˜ﬁ ‰ÂÌ  Â·ÌﬁÒËÂ"
                End If
            End If    ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\
        End If


    End If    ' ++++++++++++++ ƒ…¡√—¡÷« ≈–…‘¡√«”





    Command1_Click

    Exit Sub



MHNYMA:
    HandleError "EPIT3:EPIT311_click"
    Resume Next

End Sub

Private Sub d1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Command2_Click

End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    mForm_Load Me, fh, fw, ft, fl

    'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
    ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
    'SkinFramework1.ApplyWindow Me.hWnd
    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics


    Combo4.AddItem "1.KATA HMEP"
    Combo4.AddItem "2.KATA API—»MO"
    Combo4.AddItem "2.KATA ”≈…—¡  ¡‘¡◊Ÿ—«”«”"




    Me.Picture = LoadPicture(gPicture)
    Combo1.ListIndex = 0
    MSFlexGrid1.ColWidth(0) = 30
    'Data1.DatabaseName = gDir
    DATA1.ConnectionString = gConnect
    d1 = Now
    d2 = DateAdd("m", 6, Now)
    Check1.BackColor = epit3.BackColor


    Dim R As New ADODB.Recordset

    R.Open "SELECT *FROM PINAKES WHERE TYPOS=6 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
    ' trapezes
    Do While Not R.EOF
        If R("typos") = 6 Then
            ' Combo2.AddItem Str(R("AYJON")) + "  " + R("PERIGRAFH")
            Combo2.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH") + Right(Space(30) + Format(R("TIMH"), "0000"), 30)
        End If
        R.MoveNext
    Loop
    R.Close


    R.Open "SELECT *FROM PINAKES WHERE TYPOS=18 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
    ' trapezes
    Do While Not R.EOF
        If R("typos") = 18 Then
            ' Combo2.AddItem Str(R("AYJON")) + "  " + R("PERIGRAFH")
            PLAFON.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH") + Right(Space(30) + Format(R("TIMH"), "0000"), 30)
        End If
        R.MoveNext
    Loop
    R.Close




    'pelatew promitheytes

    Combo1.Text = Combo1.List(2)
    Combo1.ListIndex = 2









End Sub
Function Calc_Suma(ByVal CO As Long) As Single
    Dim k As Long
    On Error Resume Next
    Calc_Suma = 0
    For k = 0 To MSFlexGrid1.ROWS - 1
        Calc_Suma = Calc_Suma + Val(MSFlexGrid1.TextMatrix(k, CO))
    Next

End Function

Private Sub Form_Resize()
    mForm_Resize Me, 11865, 9405, fh, fw, ft, fl

End Sub

Private Sub Label2_Click()


    UPDATE_PINAKES 6, Label2, Combo2



    ''----------- OIKOGENEIES ---------------------------
    '  Dim f As New DataEntry2, DD
    '  Dim R As New ADODB.Recordset
    ''-----------------------------------------------------------------------------------------
    ''ÃœÕ¡ƒ≈” Ã≈‘—«”«”
    '  'sql ÔÔı Ë· „ﬂÌÂÙ·È ÙÔ update ﬁ ÙÔ addnew
    ' '  f.SHOW
    '  Load f
    '
    '  Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=6"
    '  f.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=6 " '    sql
    ''Ù· Â‰ﬂ· Ôı Ë· Ï·ﬂÌÔıÌ sto ParamGrid Í·È ÛÙ· textboxes
    'f.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " _
     '& " FROM PINAKES WHERE TYPOS=6 ORDER BY AYJON " 'ParamGrid   sql2
    '
    ''ÏÂ ÔÈ‹ ÂÌÙÔÎﬁ Ë· Í‹ÌÂÈ ÙÔ delete (ÙÔ where ÂÈÌ·È ÙÔ kod=...
    'f.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=6 AND " 'ParamGrid              DelSQL
    '
    ''ÔÈ¸ update query Ë· ‰ÁÏÈÔıÒ„ÂﬂÙ·È (ÒÔ·ÈÒÂÙÈÍ‹ )
    ''·Ì ‰ÂÌ Ë›Î˘ update Ù¸ÙÂ ‚‹Ê˘ f.ParamGrid.TextMatrix(0, 4) = ""
    ''                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
    ' DD = " UPDATE PINAKES SET TYPOS=6 where TYPOS IS NULL AND AYJON="" + t1   "
    '' AddNewScript
    '  DD = "Sub Main()" & vbCrLf & _
       '         " DIM mCon  " & vbCrLf & _
       '         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
       '         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
       '         " mCon.execute """ + DD + " " & vbCrLf & _
       '        "Set r=Nothing" & vbCrLf & _
       '        "Set mCon=Nothing" & vbCrLf & _
       '        "End Sub"
    'f.ParamGrid.TextMatrix(0, 4) = DD
    'f.ParamGrid.TextMatrix(1, 1) = "A/A"
    'f.ParamGrid.TextMatrix(2, 1) = "‘…‘Àœ” "
    ''Ã≈√¡ÀŸÕŸ À…√œ ‘¡ LABELS
    'f.Labels(0).width = f.Labels(0).width + 1000 '* TextWidth("A")
    'f.Textboxes(0).Left = f.Textboxes(0).Left + 1000 ' * TextWidth("A")
    'f.Left = Label2.Left
    'f.Top = Label2.Top + MDIForm1.Top
    '
    '
    '
    'f.SHOW 1
    'Combo2.Clear
    ' R.Open "SELECT *FROM PINAKES WHERE TYPOS=6 ORDER BY AYJON", Gdb, adOpenDynamic, _
      '     adLockOptimistic
    'Do While Not R.EOF
    '  Combo2.AddItem R("PERIGRAFH")
    '  R.MoveNext
    'Loop
    'R.Close
    'Combo2.Refresh









End Sub

Private Sub MSFlexGrid1_Click()
    Dim a As String, mc As Integer
    Dim k As Integer

    Dim L As Integer



    ' ƒ≈Õ ¬√¡∆≈… Ã≈Õœ’ ¡Õ ≈◊Ÿ –≈À¡‘ŸÕ+–—œÃ«»≈’‘ŸÕ Ã¡∆…
    If Combo1.ListIndex = 2 Or Combo1.ListIndex >= 4 Then
        Exit Sub
    End If


    ' ”‘¡ ”’√ ≈Õ‘—Ÿ‘… ¡ Õ¡ Ã«Õ ¬√¡∆≈… Ã≈Õœ’
    If Check1 Then
        Exit Sub
    End If


    ' MSFlexGrid1.row = MSFlexGrid1.row - 1




    On Error Resume Next
    MSFlexGrid1.Col = 2



    a = MSFlexGrid1.Text
    On Error Resume Next

    MSFlexGrid1.Col = 0
    If IsDate(Right(MSFlexGrid1.Text, 10)) Then
        ' OK
    Else
        Exit Sub
    End If



    Dim CurRow As Integer
    CurRow = MSFlexGrid1.row
    '  Command1_Click




    For k = 1 To MSFlexGrid1.ROWS - 1
        If MSFlexGrid1.RowIsVisible(k) And MSFlexGrid1.CellBackColor = vbRed Then

            If k Mod 2 = 0 Then
                For L = 0 To MSFlexGrid1.ColS - 1
                    MSFlexGrid1.row = k: MSFlexGrid1.Col = L: MSFlexGrid1.CellBackColor = &HC0E0FF
                Next
            End If
            If k Mod 2 = 1 Then
                For L = 0 To MSFlexGrid1.ColS - 1
                    MSFlexGrid1.row = k: MSFlexGrid1.Col = L: MSFlexGrid1.CellBackColor = vbYellow
                Next
            End If
        End If

    Next



    MSFlexGrid1.row = CurRow

    For k = 0 To MSFlexGrid1.ColS - 1
        MSFlexGrid1.Col = k
        MSFlexGrid1.CellBackColor = vbRed
    Next





    ' On Error GoTo 0

    mc = 0

    If Combo1.Text = Combo1.List(0) Then
        MDIForm1.epit311(0).Caption = "≈…”–—¡Œ«"
    Else
        MDIForm1.epit311(0).Caption = "–À«—ŸÃ«"
    End If


    mc = mc + 1
    Load MDIForm1.epit311(mc)
    MDIForm1.epit311(mc).Caption = "”≈ –À¡÷œÕ"

    If Combo1.ListIndex = 1 Then    ' –—œÃ«»≈’‘«Ú
        MDIForm1.epit311(mc).Enabled = False
    End If

    If Combo1.ListIndex = 3 Then    ' EINAI HDH STO PLAFON
        MDIForm1.epit311(mc).Enabled = False
    End If








    mc = mc + 1
    Load MDIForm1.epit311(mc)
    MDIForm1.epit311(mc).Caption = "ƒ…¡√—¡÷«"

    mc = mc + 1
    Load MDIForm1.epit311(mc)
    MDIForm1.epit311(mc).Caption = "¡–œ”‘œÀ« ”≈ –—œÃ«»≈’‘«"
    If Combo1.ListIndex = 1 Then    ' –—œÃ«»≈’‘«Ú
        MDIForm1.epit311(mc).Enabled = False
    End If

    mc = mc + 1
    Load MDIForm1.epit311(mc)
    MDIForm1.epit311(mc).Caption = "≈Œœƒœ”"

    '—… œ 1130L  MARIA SINTOY 25220-21100



    If Combo1.ListIndex = 3 Then    ' EINAI HDH STO PLAFON

        mc = mc + 1
        Load MDIForm1.epit311(mc)
        MDIForm1.epit311(mc).Caption = "≈–¡Õ¡÷œ—¡ ”‘¡ ◊≈—…¡ Ã¡”"

        MDIForm1.epit311(1).Enabled = False
        MDIForm1.epit311(2).Enabled = False
        MDIForm1.epit311(3).Enabled = False
        MDIForm1.epit311(0).Enabled = False

    End If


    If Combo1.ListIndex < 3 Then    ' –—œÃ«»≈’‘«Ú / PELATHS

        mc = mc + 1
        Load MDIForm1.epit311(mc)
        MDIForm1.epit311(mc).Caption = "ƒ…œ—»Ÿ”«"




    End If




    '  Call Me.PopupMenu(MDIForm1.par1menu)

    Call Me.PopupMenu(MDIForm1.EPITAG311)







End Sub

Private Sub plafontext_Click()

    UPDATE_PINAKES 18, PlafonText, PLAFON

End Sub
