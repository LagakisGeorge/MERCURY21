VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form TAMEIAKES 
   Caption         =   "TAMEIAKES"
   ClientHeight    =   6645
   ClientLeft      =   1110
   ClientTop       =   390
   ClientWidth     =   17880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6645
   ScaleWidth      =   17880
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Bindings        =   "TAMEIAKES.frx":0000
      Height          =   5988
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   17628
      _ExtentX        =   31089
      _ExtentY        =   10557
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2778"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2778"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DataView        =   2
      GroupByCaption  =   "”˝ÒÂÙÂ ÙÔÌ ÙﬂÙÎÔ ÙÁÚ ÛÙﬁÎÁÚ „È· ÔÏ·‰ÔÔﬂÁÛÁ"
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=1200,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0,.bold=0,.fontsize=1200,.italic=0,.underline=0,.strikethrough=0"
      _StyleDefs(46)  =   ":id=33,.charset=161"
      _StyleDefs(47)  =   ":id=33,.fontname=MS Sans Serif"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   17880
      TabIndex        =   0
      Top             =   6015
      Width           =   17880
      Begin VB.CommandButton cmdCommand5 
         Caption         =   "œÀ≈”"
         Height          =   360
         Left            =   11880
         TabIndex        =   11
         Top             =   0
         Width           =   990
      End
      Begin VB.CommandButton cmdCommand4 
         Caption         =   "Command2"
         Height          =   360
         Left            =   10680
         TabIndex        =   10
         Top             =   0
         Width           =   990
      End
      Begin VB.CommandButton cmdCommand3 
         Caption         =   "Ã« ¡Õ¡¬ ‘¡Ã"
         Height          =   360
         Left            =   9120
         TabIndex        =   9
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdCommand2 
         Caption         =   "Ã« ”’Õƒ"
         Height          =   360
         Left            =   7920
         TabIndex        =   8
         Top             =   0
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "”≈ EXCEL"
         Height          =   372
         Left            =   6120
         TabIndex        =   7
         Top             =   0
         Width           =   1572
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1213
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc ADODC2 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6315
      Width           =   17880
      _ExtentX        =   31538
      _ExtentY        =   582
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
      Caption         =   " "
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
End
Attribute VB_Name = "TAMEIAKES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Col         As TrueOleDBGrid80.Column
Dim ColS As TrueOleDBGrid80.columns

Private Sub cmdCommand2_Click()
   ' ADODC2.Recordset.Filter =
    
    
    
    
    
  '  On Error GoTo ErrHandler

100     Set ColS = TDBGrid.columns

        Dim c As Integer

110     c = TDBGrid.Col

120     TDBGrid.HoldFields

130     On Error Resume Next
        ADODC2.Recordset.Filter = " AND SyndMePOS<=2 and TamAdhmeForol=1 "

140     TDBGrid.Col = c

150     TDBGrid.EditActive = True

Dim R As New ADODB.Recordset
R.Open "SELECT COUNT(*) AS MH  FROM TAMEIAKES WHERE ISNULL(AnabTameiakis,0)=0", Gdb, adOpenDynamic, adLockOptimistic
MsgBox "OXI " + str(R(0))
R.Close
R.Open "SELECT COUNT(*) AS MH  FROM TAMEIAKES WHERE   ISNULL(AnabTameiakis,0)=2  ", Gdb, adOpenDynamic, adLockOptimistic
MsgBox "NAI " + str(R(0))




End Sub

Private Sub cmdCommand3_Click()

100     Set ColS = TDBGrid.columns

        Dim c As Integer

110     c = TDBGrid.Col

120     TDBGrid.HoldFields

130     ADODC2.Recordset.Filter = "  AnabTameiakis<>1    "
140     TDBGrid.Col = c
150     TDBGrid.EditActive = True
End Sub

Private Sub cmdCommand5_Click()


150              ADODC2.ConnectionString = gConnect ' " PROVIDER=MSDASQL;dsn=dell;uid=sa;pwd=p@ssw0rd;" 'gConnect
Dim sql As String
sql = "SELECT SyndMePOS as [·/· excel] ,EPO,[AFM]    ,[SN],[DIE]     ,[THL]   ,[PARATHRHSEIS]    ,[PERIOXH],KOD    ,[AnabTameiakis]      ,[HmerAnabTam]     ,[Timologio]   ,[Voucher]      ,[YpobSePlatforma]    ,[XROSTA]   ,[TyposTamiakis]      ,[TamAdhmeForol]      ,[AritmosPos]      ,[PerigPos]      ,[ID],[ID_NUM] ,[LagKontGlab] ,AADESTATUS,[AADERANTEBOY] from TAMEIAKES  Order by SyndMePOS"



                 ADODC2.RecordSource = sql '  "SELECT * FROM TAMEIAKES"
170        ADODC2.Refresh



End Sub

Private Sub Command1_Click()
Dim sql As String
sql = "select [DIE],EPO,[THL],[SyndMePOS],[RanteboySyndesis],[AFM],[SN],[PERIOXH],[AnabTameiakis],KOD,[HmerAnabTam],[Timologio],[Voucher],[YpobSePlatforma],[XROSTA],[TyposTamiakis],[TamAdhmeForol],[AritmosPos],[PerigPos]"
  Dim synt As String
  synt = ""
  If Len(ADODC2.Recordset.Filter) > 2 Then
      synt = " WHERE " + Replace(ADODC2.Recordset.Filter, "*", "%")
  End If
  On Error Resume Next
    
' sql = sql + ",[PARATHRHSEIS],[ID],[ID_NUM],[LagKontGlab]  from TAMEIAKES " + synt + "   Order by DIE,EPO"

sql = "select * from TAMEIAKES " ' + synt + "   Order by DIE,EPO"




 print3_xar sql, "01111111", "", 0    'sugm_str
End Sub

Private Sub Form_Load()

150              ADODC2.ConnectionString = gConnect ' " PROVIDER=MSDASQL;dsn=dell;uid=sa;pwd=p@ssw0rd;" 'gConnect
Dim sql As String
sql = "SELECT SyndMePOS as [·/· excel] ,EPO,[AFM]    ,[SN],[DIE]     ,[THL]   ,[PARATHRHSEIS]    ,[PERIOXH],KOD    ,[AnabTameiakis]      ,[HmerAnabTam]     ,[Timologio]   ,[Voucher]      ,[YpobSePlatforma]    ,[XROSTA]   ,[TyposTamiakis]      ,[TamAdhmeForol]      ,[AritmosPos]      ,[PerigPos]      ,[ID],[ID_NUM] ,[LagKontGlab] ,AADESTATUS,[AADERANTEBOY] from TAMEIAKES where SyndMePOS>999 Order by SyndMePOS"



                 ADODC2.RecordSource = sql '  "SELECT * FROM TAMEIAKES"
170        ADODC2.Refresh


220     TDBGrid.AlternatingRowStyle = True

230     TDBGrid.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
240     TDBGrid.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI

        'Dim S As TrueDBGrid80.Split

        'Set S = TDBGrid.Splits.ADD(0) ' Create a split with index 0

        'TDBGrid.Splits.Remove 1       ' Remove the split with index 1

        TDBGrid.columns(0).Width = 500 'AA
        Dim k As Integer

       ' For K = 1 To 4
            TDBGrid.columns(1).Width = 4500  'EPO
            TDBGrid.columns(2).Width = 2000
               TDBGrid.columns(3).Width = 2500
            TDBGrid.columns(4).Width = 1000
             TDBGrid.columns(5).Width = 1900 'DIE
            TDBGrid.columns(6).Width = 3100 'par
             TDBGrid.columns(7).Width = 500
            TDBGrid.columns(8).Width = 500 'ANAB
             TDBGrid.columns(9).Width = 500 'KOD
            TDBGrid.columns(10).Width = 600 'PARAT
             TDBGrid.columns(11).Width = 3000 'PARAT
            
            
       ' Next





'TDBGrid.columns(6).Width = 2500 select  EPO   ,[SyndMePOS]
'     SELECT EPO,SN,SyndMePOS,[AADERANTEBOY],AADESTATUS ,[AFM]    ,[SN],[DIE]     ,[THL]     ,[PERIOXH],KOD    ,[AnabTameiakis]   ,[PARATHRHSEIS]     ,[HmerAnabTam]     ,[Timologio]   ,[Voucher]      ,[YpobSePlatforma]    ,[XROSTA]   ,[TyposTamiakis]      ,[TamAdhmeForol]      ,[AritmosPos]      ,[PerigPos]      ,[ID],[ID_NUM] ,[LagKontGlab]  from TAMEIAKES Order by SN






End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
 TDBGrid.Height = Me.ScaleHeight - ADODC2.Height - 30 - picButtons.Height
 TDBGrid.Width = Me.ScaleWidth - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub ADODC2_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub ADODC2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  ADODC2.Caption = "Record: " & CStr(ADODC2.Recordset.AbsolutePosition)
End Sub

Private Sub ADODC2_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  ADODC2.Recordset.MoveLast
  grdDataGrid.SetFocus
  SendKeys "{down}"

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  
    Dim ANS As Integer
    ANS = MsgBox("≈…”¡… ”…√œ’—œ”", vbYesNo)

    If ANS = vbYes Then
        Dim AA As String
        AA = InputBox("ƒŸ”≈  Ÿƒ… œ")

        If AA = "12345" Then
        Else
            MsgBox "¡ ’—œ"
            Exit Sub
    
        End If
    End If

    On Error GoTo DeleteErr

    With ADODC2.Recordset
        .Delete
        .MoveNext

        If .EOF Then .MoveLast
    End With

    Exit Sub
DeleteErr:
    MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  ADODC2.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  ADODC2.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub TDBGrid_DblClick()

On Error Resume Next

    Dim fpelat4 As New pelat4

        'If Index = 0 Then
100     fpelat4.SHOW
TDBGrid.Col = 9
110     fpelat4.Text2(0).Text = Trim(TDBGrid.Text)

fpelat4.PELPROM.ListIndex = 0

150     fpelat4.data1.ConnectionString = gConnect
160     fpelat4.Refresh
170     gApoMenu = False

        'pelat4.Text2(0).SetFocus

        'fpelat4.SHOW
180     fpelat4.Text2(0).SetFocus
190    ' SendKeys "+{ENTER}"

        Dim DUM

200     DUM = fpelat4.KARTELLA(gDir)
        'End If
210     Set fpelat4 = Nothing




End Sub

Private Sub TDBGrid_FilterChange()

        '<EhHeader>
        On Error Resume Next  ' GoTo TDBGrid_FilterChange_Err

        '</EhHeader>

        'Gets called when an action is performed on the filter bar

        On Error GoTo errHandler

100     Set ColS = TDBGrid.columns

        Dim c As Integer

110     c = TDBGrid.Col

120     TDBGrid.HoldFields
Dim CC As String

CC = getFilter()

If Len(Trim(CC)) < 1 Then
   Exit Sub
    'CC = "  EPO LIKE '*A*'   "

End If

130     ADODC2.Recordset.Filter = CC

140     TDBGrid.Col = c

150     TDBGrid.EditActive = True

        Exit Sub

errHandler:

160     MsgBox Err.Source & ":" & vbCrLf & Err.Description

170     Call cmdClearFilter_Click

        '<EhFooter>
        Exit Sub

TDBGrid_FilterChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.TDBGrid_FilterChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.TDBGrid_FilterChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub



Private Function getFilter() As String

        'Creates the SQL statement in adodc1.recordset.filter
        'and only filters text currently. It must be modified to filter other data types.
        '<EhHeader>
        On Error Resume Next ' GoTo getFilter_Err

        '</EhHeader>

        Dim tmp As String

        Dim N   As Integer

100     For Each Col In ColS

110         If Trim(Col.FILTERTEXT) <> "" Then
120             N = N + 1

130             If N > 1 Then
140                 tmp = tmp & " AND "
                End If

150             If ADODC2.Recordset(Col.ColIndex).Type = 5 Then    '
160                 If InStr(">< >= <=  = ", Left(Col.FILTERTEXT, 1)) > 0 And Len(Col.FILTERTEXT) > 1 And IsNumeric(Right(Col.FILTERTEXT, 1)) Then
170                     tmp = tmp & Col.DataField & Col.FILTERTEXT
                    Else
180                     tmp = tmp & Col.DataField & " >  " & Val(Col.FILTERTEXT)
                    End If
                ElseIf ADODC2.Recordset(Col.ColIndex).Type = 3 Then     '
                     tmp = tmp & Col.DataField & " = " & Col.FILTERTEXT & " "
                Else
                If Len(Col.FILTERTEXT) < 1 Then
                        Col.FILTERTEXT = " "
                   End If
                   
190                 tmp = tmp & Col.DataField & " LIKE '*" & Col.FILTERTEXT & "*'"
                End If

            End If

200     Next Col


     
210     getFilter = tmp

        '<EhFooter>
        Exit Function

getFilter_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.getFilter " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.getFilter " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub TDBGrid_HeadClick(ByVal ColIndex As Integer)

        '--------------------------------------------------------------------------------------=
        '<EhHeader>
        On Error GoTo TDBGrid_HeadClick_Err

        '</EhHeader>
        Dim sumes(100) As String

100     For k = 0 To ADODC2.Recordset.FIELDS.Count - 1
110         sumes(k) = TDBGrid.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
        Next

120     If ADODC2.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc" Then   ' strSort
130         ADODC2.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
140         ADODC2.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

150     For k = 0 To ADODC2.Recordset.FIELDS.Count - 1
160         TDBGrid.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
        Next

        '<EhFooter>
        Exit Sub

TDBGrid_HeadClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.TDBGrid_HeadClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.TDBGrid_HeadClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdClearFilter_Click()

        '<EhHeader>
        On Error Resume Next  ' NEXT '  GoTo cmdClearFilter_Click_Err

        '</EhHeader>

        'Clears filter from grid

100     For Each Col In TDBGrid.columns

110         Col.FILTERTEXT = ""

120     Next Col

130     ADODC2.Recordset.Filter = adFilterNone

        '<EhFooter>
        Exit Sub

cmdClearFilter_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.cmdClearFilter_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.cmdClearFilter_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


