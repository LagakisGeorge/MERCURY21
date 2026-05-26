VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form apot4PERP 
   BackColor       =   &H00FF0000&
   Caption         =   "Τιμοκατάλογοι"
   ClientHeight    =   12315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20835
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12315
   ScaleWidth      =   20835
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   5895
      Left            =   12360
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "apot4PERP.frx":0000
      Top             =   4920
      Width           =   7215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Νέα εγγραφή"
      Height          =   285
      Left            =   4200
      TabIndex        =   12
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ολος ο τιμοκατάλογος κατά Είδος"
      Height          =   360
      Left            =   7560
      TabIndex        =   11
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ολος ο τιμοκατάλογος κατά Πελατη"
      Height          =   360
      Left            =   7560
      TabIndex        =   10
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox timokEID 
      Height          =   315
      Left            =   5400
      TabIndex        =   8
      Top             =   510
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Διαγραφή τρέχουσας σειράς"
      Height          =   285
      Left            =   7560
      TabIndex        =   7
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox timokPEL 
      Height          =   315
      Left            =   660
      TabIndex        =   5
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Κριτήρια"
      Height          =   360
      Left            =   15330
      TabIndex        =   4
      Top             =   2250
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Διαγραφή Τιμοκαταλόγου"
      Height          =   360
      Left            =   15330
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   2655
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nέα είδη στον τιμοκατάλογο"
      Height          =   360
      Left            =   15330
      TabIndex        =   0
      Top             =   1860
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000010&
      Caption         =   "Εξοδος"
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   3375
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Bindings        =   "apot4PERP.frx":0223
      Height          =   9105
      Left            =   210
      TabIndex        =   1
      Top             =   1710
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   16060
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "TIMOKPEL"
      Columns(0).DataField=   "TIMOKPEL"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "TIMOKEID"
      Columns(1).DataField=   "TIMOKEID"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "EKPT"
      Columns(2).DataField=   "EKPT"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "TIMH"
      Columns(3).DataField=   "TIMH"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "TIMOKID"
      Columns(4).DataField=   "TIMOKID"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=136,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11400
      Top             =   6840
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Τιμοκατάλογος Είδους"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   5070
      TabIndex        =   9
      Top             =   90
      Width           =   1665
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Τιμοκατάλογος Πελατη"
      ForeColor       =   &H8000000F&
      Height          =   345
      Left            =   540
      TabIndex        =   6
      Top             =   120
      Width           =   1845
   End
End
Attribute VB_Name = "apot4PERP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fsql      As String

Dim FIsHidden As Boolean

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Dim Xval As String

Private Sub Command2_Click()
  Unload Me
  
End Sub

Private Sub Command5_Click()
 Dim m_ID As Long
 m_ID = Adodc1.Recordset("timokid")
 Dim ANS As Integer
 ANS = MsgBox("να διαγραφεί η εγγραφή " + Adodc1.Recordset("timokpel") + " " + Adodc1.Recordset("timokeid"), vbYesNo)
Dim k As Integer

 
 If ANS = vbYes Then
    Gdb.Execute "delete from TIMOKAT WHERE TIMOKID=" + str(m_ID), k
    If k > 0 Then
       MsgBox "ΔΙΕΓΡ’ΦΗ"
    
    End If
    
 End If
 Adodc1.Refresh




End Sub

Private Sub Command6_Click()

    Adodc1.RecordSource = "SELECT * FROM TIMOKAT  order by  TIMOKPEL"
     Adodc1.Refresh

TDBGrid1.SetFocus








End Sub

Private Sub Command7_Click()


    Adodc1.RecordSource = "SELECT * FROM TIMOKAT  order by  TIMOKEID"
     Adodc1.Refresh

TDBGrid1.SetFocus





End Sub

Private Sub Command8_Click()

  Gdb.Execute "insert into TIMOKAT (TIMOK,KOD,TIMOKEID,TIMOKPEL,EKPT )VALUES (0,'','','" + timokPEL.Text + "',0)"
  Adodc1.Refresh
  



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot4.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl

110     Me.KeyPreview = True

120     FIsHidden = False

130     Me.Picture = LoadPicture(gPicture)

140     Adodc1.ConnectionString = gConnect

150     Adodc1.RecordSource = "select TOP 10 * FROM TIMOKAT"

        On Error GoTo create

160     Adodc1.Refresh

        Exit Sub

create:
        Resume Next
End Sub



Private Sub timokEID_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If
End Sub

Private Sub timokEID_LostFocus()

On Error Resume Next

100     If Len(timokEID.Text) < 1 Then Exit Sub
110     Adodc1.RecordSource = "SELECT * FROM TIMOKAT  WHERE TIMOKEID='" + timokEID.Text + "'"
120     Adodc1.Refresh

TDBGrid1.SetFocus








End Sub

Private Sub timokPEL_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If
End Sub

Private Sub timokPEL_LostFocus()
On Error Resume Next

100     If Len(timokPEL.Text) < 1 Then Exit Sub
110     Adodc1.RecordSource = "SELECT * FROM TIMOKAT  WHERE TIMOKPEL='" + timokPEL.Text + "'"
120     Adodc1.Refresh

'Adodc1.SetFocus

        '<EhFooter>
        Exit Sub


'



End Sub



