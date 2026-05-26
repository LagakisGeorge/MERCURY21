VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form PARAMETROI 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   12675
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   16470
   LinkTopic       =   "Form2"
   ScaleHeight     =   12675
   ScaleWidth      =   16470
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Bindings        =   "PARAMETROI.frx":0000
      Height          =   6135
      Left            =   480
      TabIndex        =   10
      Top             =   840
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   10821
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "VAR"
      Columns(0).DataField=   "VAR"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "SXOLIA"
      Columns(1).DataField=   "SXOLIA"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TIMH"
      Columns(2).DataField=   "TIMH"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   979
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2778"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=15266"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=15134"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2778"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Named:id=33:Normal"
      _StyleDefs(49)  =   ":id=33,.parent=0,.bgcolor=&H80FF&"
      _StyleDefs(50)  =   "Named:id=34:Heading"
      _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=34,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=37:Caption"
      _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(59)  =   "Named:id=38:HighlightRow"
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdFOCUS 
      Caption         =   "FOCUS"
      Height          =   360
      Left            =   4440
      TabIndex        =   9
      Top             =   240
      Width           =   990
   End
   Begin VB.TextBox txtpRINTER 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   8040
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   3360
      TabIndex        =   7
      Top             =   7080
      Width           =   6375
   End
   Begin VB.CommandButton cmdPRINTERLIST 
      Caption         =   "À…”‘¡ ≈ ‘’–Ÿ‘ŸÕ"
      Height          =   360
      Left            =   360
      TabIndex        =   6
      Top             =   7080
      Width           =   2655
   End
   Begin VB.ComboBox category 
      Height          =   288
      Left            =   7320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   2892
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   450
      Left            =   12120
      TabIndex        =   0
      Top             =   240
      Width           =   1425
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   3720
      Top             =   9840
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   609
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
      Bindings        =   "PARAMETROI.frx":0015
      Height          =   2400
      Left            =   10680
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4233
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   33023
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "–¡—¡Ã≈‘—œ…"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "VAR"
         Caption         =   " Ÿƒ"
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
         DataField       =   "SXOLIA"
         Caption         =   "”˜¸ÎÈ·"
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
      BeginProperty Column02 
         DataField       =   "TIMH"
         Caption         =   "TÈÏﬁ"
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
         MarqueeStyle    =   4
         Size            =   349
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "¡Ì·ÊﬁÙÁÛÁ"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label PARAM 
      Height          =   270
      Left            =   180
      TabIndex        =   2
      Top             =   9600
      Visible         =   0   'False
      Width           =   3090
   End
End
Attribute VB_Name = "PARAMETROI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   PARAMETROI.PARAM.Caption = "APOT2"
'   PARAMETROI.SHOW 1
'    F_TAB = Val(FindParametroi(1,"APOT2", "F_TAB", "3", "”Â ÔÈ¸ TAB ÂﬂÌ·È ÛÙ·Ï·ÙÁÏ›ÌÔ"))
'Function FindParametroi(FORMA As String, parametros As String, default As String, sxolia As String)
'   Dim R As New ADODB.Recordset
'   Dim sql As String
'
'   If default = "DELETE" Then
'     Gdb.Execute "DELETE FROM PARAMETROI WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
'     FindParametroi = 0
'     Exit Function
'   End If
'   sql = "select * from PARAMETROI where FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
'   R.Open sql, Gdb, adOpenDynamic, adLockBatchOptimistic
'
'   If R.EOF Then
'      sql = "insert into PARAMETROI (FORMA,VAR,TIMH,SXOLIA) VALUES ('" + FORMA + "','" + parametros + "','" + default + "','" + sxolia + "')"
'      Gdb.Execute sql
'
'      FindParametroi = default
'   Else
'
'      FindParametroi = R("TIMH")
'
'
'   End If
'
'
'
'
'
'End Function

Dim F_DOK

Private Sub category_Change()

           ' Adodc1.RecordSource = "select * FROM PARAMETROI WHERE KATEG=" + Left$(category.Text, 1) + " AND  FORMA='" + PARAM.Caption + "' ORDER BY SXOLIA"
          '  Adodc1.Refresh

End Sub

Private Sub category_Click()

            Adodc1.RecordSource = "select * FROM PARAMETROI WHERE KATEG=" + Left$(category.Text, 1) + " AND  FORMA='" + PARAM.Caption + "' ORDER BY SXOLIA"
            Adodc1.Refresh

End Sub

Private Sub cmdFOCUS_Click()
 'DataGrid1.SetFocus
End Sub

Private Sub cmdPRINTERLIST_Click()
 
 Dim i As Integer
 
  For i = 0 To Printers.Count - 1
       List1.AddItem Printers(i).DeviceName
  Next i
  
  
  
End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PARAMETROI.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PARAMETROI.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()
'
'        ' DATAGRID1.SelEndCol=0
'        '<EhHeader>
'        On Error GoTo DataGrid1_Click_Err
'
'        '</EhHeader>
'
'100     DataGrid1.Col = 2
'110     DataGrid1.SelStart = 0
'120     DataGrid1.SelLength = Len(DataGrid1.Text)
'
'        '    TIMText3.SelStart = 0
'        '    TIMText3.SelLength = Len(TIMText3.Text)
'
'        '<EhFooter>
'        Exit Sub
'
'DataGrid1_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.PARAMETROI.DataGrid1_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.PARAMETROI.DataGrid1_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>

End Sub

Private Sub DataGrid1_DblClick()
'      If DataGrid1.Col = 0 Then
'            Adodc1.RecordSource = "select * FROM PARAMETROI WHERE FORMA='" + PARAM.Caption + "' ORDER BY VAR"
'            Adodc1.Refresh
'      End If
'
'      If DataGrid1.Col = 1 Then
'            Adodc1.RecordSource = "select * FROM PARAMETROI WHERE FORMA='" + PARAM.Caption + "' ORDER BY SXOLIA"
'            Adodc1.Refresh
'      End If
'
      
      
End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     Adodc1.ConnectionString = gConnect

        ' DataGrid1.C= False
        'DataGrid1.Splits(1).Columns(0).Visible = False

'         If Len(PARAM.Caption) > 0 Then
'            Adodc1.RecordSource = "select * FROM PARAMETROI WHERE FORMA='" + PARAM.Caption + "'"
'         Else
'            Adodc1.RecordSource = "select * FROM PARAMETROI"
'         End If
'            Adodc1.Refresh
'        '    Adodc1.Recordset.MoveFirst
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PARAMETROI.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PARAMETROI.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Paint()

        '<EhHeader>
        On Error GoTo Form_Paint_Err

        '</EhHeader>
        On Error Resume Next

100     If Len(PARAM.Caption) > 0 Then
110         Adodc1.RecordSource = "select * FROM PARAMETROI WHERE FORMA='" + PARAM.Caption + "' ORDER BY KATEG,SXOLIA"
        Else
120         Adodc1.RecordSource = "select * FROM PARAMETROI"
        End If

130     Adodc1.Refresh
140     Adodc1.Recordset.MoveFirst
150  '   DataGrid1.columns(1).Width = 7000

        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PARAMETROI.Form_Paint " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PARAMETROI.Form_Paint " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Unload(Cancel As Integer)

   PARAMETROI_TO_ARRAY

End Sub

Private Sub List1_Click()
'txtpRINTER
Dim a As String
    a = List1.List(List1.ListIndex)
txtpRINTER.Text = a

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
   ' DataGrid1.SetFocus
    
  
  End If
  
End Sub

Private Sub Text1_LostFocus()
Dim mO As String
mO = SameLetters(Text1.Text)
         Adodc1.RecordSource = "select * FROM PARAMETROI WHERE FORMA='" + PARAM.Caption + "' AND (VAR LIKE '%" + mO + "%'  OR SXOLIA LIKE '%" + mO + "%' )   ORDER BY KATEG,SXOLIA"

         Adodc1.Refresh
         On Error Resume Next
         Adodc1.Recordset.MoveFirst




End Sub
