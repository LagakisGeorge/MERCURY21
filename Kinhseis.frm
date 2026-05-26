VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form APOT5 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   9045
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   13275
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000010&
      Caption         =   "Εξοδος"
      Height          =   735
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5400
      Width           =   2055
   End
   Begin VB.ListBox List2 
      Height          =   3180
      Left            =   9960
      MultiSelect     =   1  'Simple
      TabIndex        =   36
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SQL"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   8760
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1215
      Left            =   120
      TabIndex        =   34
      Top             =   7560
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2143
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Kinhseis.frx":0000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Εκτύπωση"
      Height          =   375
      Left            =   2520
      TabIndex        =   33
      Top             =   6720
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc RR 
      Height          =   330
      Left            =   240
      Top             =   8040
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "RR"
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
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1320
      TabIndex        =   30
      Top             =   7200
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   1080
      TabIndex        =   29
      Top             =   4320
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd-M-yyyy  HH:mm"
      Format          =   292683779
      CurrentDate     =   37771
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   3960
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   450
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      CustomFormat    =   "dd-M-yyyy  HH:mm"
      Format          =   292683779
      CurrentDate     =   37771
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H8000000D&
      Caption         =   "Check3"
      Height          =   255
      Left            =   4920
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Kinhseis.frx":0084
      Left            =   8760
      List            =   "Kinhseis.frx":0086
      TabIndex        =   26
      Text            =   "Combo2"
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H8000000D&
      Caption         =   "Αξίες αντι τιμή μονάδος"
      Height          =   375
      Left            =   4920
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1320
      TabIndex        =   23
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1320
      TabIndex        =   20
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Αποστολή σε αρχ.dbase"
      Height          =   615
      Left            =   7560
      TabIndex        =   18
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   2520
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Text            =   "c:\kinhseis.txt"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Αποστολή σε αρχείο ASCII"
      Height          =   615
      Left            =   4920
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000D&
      Caption         =   "Barcode"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   13
      Top             =   4440
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000D&
      Caption         =   "Φορολογικός"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   12
      Top             =   3960
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Αναζήτηση"
      Height          =   615
      Left            =   3105
      TabIndex        =   9
      Top             =   3960
      Width           =   1470
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   4680
      Width           =   975
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   375
      Left            =   240
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   1215
      Left            =   4980
      TabIndex        =   14
      Top             =   3900
      Width           =   2535
   End
   Begin TrueOleDBGrid80.TDBGrid DATAGRID1 
      Bindings        =   "Kinhseis.frx":0088
      Height          =   3825
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   6747
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
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
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
      ColumnFooters   =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DataView        =   2
      GroupByCaption  =   "Σύρετε τον τίτλο της στήλης για ομαδοποίηση"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Προμηθευτής"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Κατηγορία ΦΠΑ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Απο Απόδειξη"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Εως Απόδειξη"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Εως κωδικό"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ΑΠΟΘΗΚΗ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Από κωδικό"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Εως"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Από Ημερ/νία"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "APOT5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Col As TrueOleDBGrid80.Column

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Dim ColS As TrueOleDBGrid80.columns

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim synt, a, B, c, d, HM2, mydb, FF, Start, R, k

        Dim DB, ee

        Dim s1, S2, DUM

        Dim C50(1 To 31) As Integer, C100(1 To 31) As Integer, L

        Dim d1 As New ADODB.Recordset

        Dim CC As String

100     DataGrid1.AlternatingRowStyle = True

110     DataGrid1.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
120     DataGrid1.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI

        Dim polepis As String, pol As String, ago As String, AGOEPIS As String

130     pol = ""

        ' A = Get_AJ(pol, polepis, ago, AGOEPIS)
140     For k = 0 To List2.ListCount - 1

150         If List2.Selected(k) Then
160             pol = pol + "'" + Right(List2.List(k), 1) + "',"
            End If

        Next

170     If Len(pol) > 1 Then
180         pol = Left(pol, Len(pol) - 1)
        Else

190         MsgBox "ΔΕΝ ΕΠΙΛΕΞΑΤΕ ΠΑΡΑΣΤΑΤΙΚΑ"

            Exit Sub

        End If

        'Set db = OpenDatabase(mDir + "\kinhseis.mdb")
        'db.EID.ConnectionString = ";DATABASE=D:\TAMEIA\EIDH.MDB"
        'db.Close

        On Error Resume Next

        'Printer.FontSize = 14
        'Printer.FontName = "Courier"

        ' Text1.Text = DTPicker1.Value

        '  Text2.Text = DTPicker2.Value

200     Printer.FontSize = 10
210     Printer.FontName = "Courier New"
220     Printer.Font.Charset = 161
230     Printer.FontBold = True

240     RR.ConnectionString = gConnect    '  = mDir + "\kinhseis.mdb"
250     RR.RecordSource = "select top 10 * from EGGTIM"
260     RR.Refresh

        ' Exit Sub

270     synt = ""
        ' ---------- ελεγχος ημερομηνιών -----------------------------------------
        '                If IsDate(Text1.Text) And IsDate(Text2.Text) Then'
        '             DTPicker2.Value = DateAdd("d", 1, DTPicker2.Value)
        ' If Combo1.List(3) = Combo1.Text Then   'παρουσιαση κινήσεων
        '      synt = "LEFT(ATIM,1) IN (" + pol + ") AND  EGGTIM.HME>='" + Format(DTPicker1.value, "mm/dd/yyyy") + "' and EGGTIM.HME<= '" + Format(DTPicker2.value, "mm/dd/yyyy") + "'"
        '  Else
280     synt = "LEFT(ATIM,1) IN (" + pol + ") AND    HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy hh:mm") + "' and HME<= '" + Format(DTPicker2.Value, "mm/dd/yyyy hh:mm") + "'"
        '  End If

        ' Else
        '   MsgBox "Οχι σωστές ημερομηνίες"
        '   Exit Sub
        ' End If

290     Me.MousePointer = vbHourglass

        ' ---------- ελεγχος κωδικών -----------------------------------------
300     If Option1(0).Value = True Then

310         If IsNumeric(Text3.Text) And IsNumeric(Text4.Text) Then
320             synt = IIf(synt = "", "", synt + " and ")
330             synt = synt + " EGGTIM.KOD >='" + Trim(LTrim(Text3.Text)) + "' and EGGTIM.KOD<='" + Trim(LTrim(Text4.Text)) + "'"
            End If

        Else

340         If IsNumeric(Text3.Text) And IsNumeric(Text4.Text) Then
350             synt = IIf(synt = "", "", synt + " and ")
360             synt = synt + " EGGTIM.ERG >='" + Trim(LTrim(str(Text3.Text))) + "' and EGGTIM.ERG<='" + Trim(LTrim(Text4.Text)) + "'"
            End If
        End If

        ' ---------- ελεγχος ΑΠΟΘΗΚΩΝ -----------------------------------------
370     If IsNumeric(Text5.Text) Then
380         synt = IIf(synt = "", "", synt + " and ")
390         synt = synt + " APOT=" + Text5.Text   '  + " and SEIRIAKH<=" + Text6.Text
        End If

        ' ---------- ελεγχος αποδείξεων -----------------------------------------
400     If IsNumeric(Text9.Text) And IsNumeric(Text10.Text) Then
410         synt = IIf(synt = "", "", synt + " and ")
420         synt = synt + " ATIM>='" + Text9.Text + "' and ATIM<='" + Text10.Text + "'"
        End If

        ' ---------- ελεγχος ΦΠΑ -----------------------------------------
430     If IsNumeric(Text8.Text) Then
440         synt = IIf(synt = "", "", synt + " and ")
450         synt = synt + " FPA=" + Text8.Text
        End If

460     synt = IIf(synt = "", "true", synt)

        '-------------------------------------------------------------------------
470     If Combo1.List(0) = Combo1.Text Then    ' Z mhxanhs
            ' rr.RecordSource = "select sum(ajia) as sajia,SUM(POSO) AS SPOSO from eggtim where " + synt

480         RR.RecordSource = "select STR(sum(POSO*TIMM),10,2) as [ΑΞΙΑ],SUM(POSO) AS [ΠΟΣΟΤΗΤΕΣ] from EGGTIM where  " + synt

490         RR.Refresh
500         DataGrid1.columns(0).Width = 1000

            '     createGrid (2)
            '     DataGrid1.Columns(0).DataField = "sajia"
            '     DataGrid1.Columns(0).Caption = "ΑΞΙΑ"
            '     DataGrid1.Columns(0).NumberFormat = "#####.00"
            '
            '     DataGrid1.Columns(1).DataField = "sPOSO"
            '     DataGrid1.Columns(1).Caption = "ΣΥΝ.ΠΟΣΟΤΗΤΑΣ"
            '     DataGrid1.Columns(1).NumberFormat = ""
            '

            '     If Check1 Then
            '       RR.Recordset.MoveFirst
            '       Printer.Print Space(10) + "Από " + Text1.Text + " έως " + Text2.Text + Space(10) + " Αξία "
            '       Printer.Print Space(10) + Space(30) + Right(Space(12) + Format(RR.Recordset("sajia"), "###,###,###.00"), 12)
            '       Printer.Print Space(30) + "------------"
            '       Printer.EndDoc
            '     End If
            '

            '-------------------------------------------------------------------------
510     ElseIf Combo1.List(3) = Combo1.Text Then   'παρουσιαση κινήσεων

            ' ---------- ελεγχος PROMHU -----------------------------------------

            '     RR.RecordSource = "select CONVERT(CHAR(8),HME,3) as [Ημερ/νία],EGGTIM.*,EID.ONO,EID.AEG from EGGTIM  INNER  join  EID ON EGGTIM.KODE=EID.KOD where EGGTIM.PIS>0 AND " + synt + " order by HME,ATIM;"
520         RR.RecordSource = " select datename(dw,HME) AS WEEKDAY ,day(HME) AS DAY ,MONTH(HME) AS MONTH ,YEAR(HME) AS YEAR, DATEPART(hh,HME) AS HOU,DATEPART(MI,HME ) AS MINUTE ,POSO,TIMM,ROUND(POSO*TIMM*(1-EKPT/100),2) AS AXIA, KODE,APOT  FROM EGGTIM where EGGTIM.PIS>0 AND " + synt + " order by HME,ATIM;"
            ' INNER  join  EID ON EGGTIM.KODE=EID.KOD where EGGTIM.PIS>0 AND " + synt + " order by HME,ATIM;"

            '     createGrid (7)
            '        DATAGRID1.Columns(0).DataField = "name"
            '        DATAGRID1.Columns(1).DataField = "POSO"
            '        DATAGRID1.Columns(2).DataField = "TIMM"
            '        DATAGRID1.Columns(3).DataField = "AJIA"
            '        DATAGRID1.Columns(4).DataField = "EGGTIM.HME"
            '        DATAGRID1.Columns(5).DataField = "ATIM"
            '        DATAGRID1.Columns(6).DataField = "eggtim.ERG"
            '
            '        DATAGRID1.Columns(0).Caption = "Oνομα"
            '        DATAGRID1.Columns(1).Caption = "Ποσότητα"
            '        DATAGRID1.Columns(2).Caption = "Τιμ.μον"
            '        DATAGRID1.Columns(3).Caption = "Αξία "
            '        DATAGRID1.Columns(4).Caption = "Ημερ/νία"
            '        DATAGRID1.Columns(5).Caption = "ΠΑΡ/ΚΟ"
            '        DATAGRID1.Columns(6).Caption = "ERG"
            '
            '        DATAGRID1.Columns(0).width = 3000
            '        DATAGRID1.Columns(1).width = 800
            '        DATAGRID1.Columns(2).width = 800
            '        DATAGRID1.Columns(3).width = 800
            '
            '        DATAGRID1.Columns(1).Alignment = dbgRight
            '        DATAGRID1.Columns(2).Alignment = dbgRight
            '        DATAGRID1.Columns(3).Alignment = dbgRight
            '        DATAGRID1.Columns(2).NumberFormat = "#####.00"
            '        DATAGRID1.Columns(3).NumberFormat = "#####.00"
            '

530         RR.Refresh

            '--------------------------'παρουσιαση κινήσεων ανα (σουμα hmeras)-----------------------------------------------
540     ElseIf Combo1.List(6) = Combo1.Text Then
            'Set mydb = OpenDatabase(mDir + "\kinhseis.mdb")

            '
            On Error Resume Next

550         Gdb.Execute "drop table TEMP2"
            '     milsec 1000

            On Error GoTo Command1_Click_Err

560         RR.RecordSource = "select STR(sum(PIS*TIMM),10,2) as [ΑΞΙΑ] ,CONVERT(CHAR(12),HME,3) as [ΗΜΕΡ/ΝΙΑ] from EGGTIM where PIS>0 AND " + synt + " group by CONVERT(CHAR(12),HME,3) ORDER BY CONVERT(CHAR(12),HME,3);"
570         RR.Refresh

580         d1.Open "SELECT SUM(PIS*TIMM) FROM EGGTIM where PIS>0 AND " + synt, Gdb, adOpenForwardOnly, adLockReadOnly
590         DataGrid1.columns(0).FooterText = Format(d1(0), "######0.00")
600         DataGrid1.columns(0).alignment = dbgRight

            '   createGrid (2)
            '     DataGrid1.Columns(0).DataField = "sajia"
            '     DataGrid1.Columns(0).Caption = "ΑΞΙΑ"
            '     DataGrid1.Columns(0).NumberFormat = "######.00"
            '
            '     DataGrid1.Columns(1).DataField = "mhme" 'Format(HME, "dd/mm/yyyy")
            '     DataGrid1.Columns(1).Caption = "mhme"
            '     'DataGrid1.Columns(1).NumberFormat = ""

            '   If Check1 Then
            '       RR.Recordset.MoveFirst
            '       Printer.Print Space(10) + "Ημερομηνία " + Space(10) + "Αξία "
            '       s1 = 0
            '       Do While Not RR.Recordset.EOF
            '           Printer.Print Space(10) + RR.Recordset("mHME") + Space(10) + Right(Space(12) + Format(RR.Recordset("sajia"), "###,###,###.00"), 12)
            '           If Not IsNull(RR.Recordset("sajia")) Then s1 = s1 + RR.Recordset("sajia")
            '           RR.Recordset.MoveNext
            '       Loop
            '       Printer.Print Space(30) + "------------"
            '
            '       Printer.Print Space(30) + Right(Space(12) + Format(s1, "###,###,###.00"), 12)
            '       Printer.EndDoc
            '
            '   End If

            'σουμα ανα κατηγορία
            '----------------------------------------------------------------------------------
610     ElseIf Combo1.List(7) = Combo1.Text Then   'παρουσιαση κινήσεων ανα (σουμα hmeras)

            'Set mydb = OpenDatabase(mDir + "\kinhseis.mdb")
            On Error Resume Next

620         Gdb.Execute "drop table TEMP2"

            On Error GoTo Command1_Click_Err

630         RR.RecordSource = "select sum(PIS*TIMM) as [ΑΞΙΑ],EID.KODLOG AS [ΚΑΤΗΓΟΡΙΑ] from EGGTIM LEFT JOIN EID ON EGGTIM.KODE=EID.KOD where EGGTIM.PIS>0 AND " + synt + " group by EID.KODLOG ORDER BY EID.KODLOG;"
640         RR.Refresh

650         d1.Open "SELECT SUM(PIS*TIMM) FROM EGGTIM where PIS>0 AND " + synt, Gdb, adOpenForwardOnly, adLockReadOnly
660         DataGrid1.columns(0).FooterText = Format(d1(0), "######0.00")
670         DataGrid1.columns(0).alignment = dbgRight

            '   createGrid (2)
            '     DATAGRID1.Columns(0).DataField = "sajia"
            '     DATAGRID1.Columns(0).Caption = "ΑΞΙΑ"
            '     DATAGRID1.Columns(0).NumberFormat = "###,###.00"
            '
            '     DATAGRID1.Columns(1).DataField = "mhme" 'Format(HME, "dd/mm/yyyy")
            '     DATAGRID1.Columns(1).Caption = "mhme"

            '   RR.Refresh

            '   If Check1 Then
            '       RR.Recordset.MoveFirst
            '       Printer.Print Text1.Text + " - " + Text2.Text + Space(10) + "Αξία "
            '       s1 = 0
            '       Do While Not RR.Recordset.EOF
            '           Printer.Print Space(10) + RR.Recordset("mHME") + Space(10) + Right(Space(12) + Format(RR.Recordset("sajia"), "###,###,###.00"), 12)
            '           If Not IsNull(RR.Recordset("sajia")) Then s1 = s1 + RR.Recordset("sajia")
            '           RR.Recordset.MoveNext
            '       Loop
            '       Printer.Print Space(22) + "------------"
            '
            '       Printer.Print Space(22) + Right(Space(12) + Format(s1, "###,###,###.00"), 12)
            '       Printer.EndDoc
            '
            '   End If
            '
            'σουμα ανα φπα
            '----------------------------------------------------------------------------------
680     ElseIf Combo1.List(8) = Combo1.Text Then   'παρουσιαση κινήσεων ανα (σουμα hmeras)

            'Set mydb = OpenDatabase(mDir + "\kinhseis.mdb")
            On Error Resume Next

            '     Gdb.Execute "drop table TEMP2"
            On Error GoTo Command1_Click_Err

690         RR.RecordSource = "select STR(sum(PIS*TIMM),10,2) as [ΑΞΙΑ] ,FPA as [ΦΠΑ] from EGGTIM where " + synt + " group by FPA;"
700         RR.Refresh

710         d1.Open "SELECT SUM(PIS*TIMM) FROM EGGTIM where PIS>0 AND " + synt, Gdb, adOpenForwardOnly, adLockReadOnly
720         DataGrid1.columns(0).FooterText = Format(d1(0), "######0.00")
730         DataGrid1.columns(0).alignment = dbgRight

            '   createGrid (2)
            '     DATAGRID1.Columns(0).DataField = "sajia"
            '     DATAGRID1.Columns(0).Caption = "ΑΞΙΑ"
            '     DATAGRID1.Columns(0).NumberFormat = "###,###.00"
            '
            '     DATAGRID1.Columns(1).DataField = "mhme" 'Format(HME, "dd/mm/yyyy")
            '     DATAGRID1.Columns(1).Caption = "ΤΜΗΜΑ(ΦΠΑ)"

            '   If Check1 Then
            '       RR.Recordset.MoveFirst
            '       Printer.Print Text1.Text + " - " + Text2.Text + Space(10) + "Αξία "
            '       s1 = 0
            '       Do While Not RR.Recordset.EOF
            '           Printer.Print Space(10) + Str(RR.Recordset("mHME")) + Space(10) + Right(Space(12) + Format(RR.Recordset("sajia"), "###,###,###.00"), 12)
            '           If Not IsNull(RR.Recordset("sajia")) Then s1 = s1 + RR.Recordset("sajia")
            '           RR.Recordset.MoveNext
            '       Loop
            '       Printer.Print Space(22) + "------------"
            '
            '       Printer.Print Space(22) + Right(Space(12) + Format(s1, "###,###,###.00"), 12)
            '       Printer.EndDoc
            '
            '   End If

            'σουμα ανα MHXANH
            '----------------------------------------------------------------------------------
740     ElseIf Combo1.List(9) = Combo1.Text Then   'παρουσιαση κινήσεων ανα (σουμα hmeras)

            'Set mydb = OpenDatabase(mDir + "\kinhseis.mdb")
            On Error Resume Next

750         Gdb.Execute "drop table TEMP2"

            On Error GoTo Command1_Click_Err

760         RR.RecordSource = "select sum(PIS*TIMM) as [ΑΞΙΑ] ,APOT AS [ΑΠΟΘΗΚΗ] from EGGTIM where " + synt + " group by APOT;"
770         RR.Refresh
            '  createGrid (2)
            '     DATAGRID1.Columns(0).DataField = "SAJIA"
            '     DATAGRID1.Columns(0).Caption = "ΑΞΙΑ"
            '     DATAGRID1.Columns(0).NumberFormat = "###,###.00"
            '
            '     DATAGRID1.Columns(1).DataField = "mhme" 'Format(HME, "dd/mm/yyyy")
            '     DATAGRID1.Columns(1).Caption = "ΤΜΗΜΑ(ΦΠΑ)"

            '   If Check1 Then
            '       RR.Recordset.MoveFirst
            '       Printer.Print Text1.Text + " - " + Text2.Text + Space(10) + "Αξία "
            '       s1 = 0
            '       Do While Not RR.Recordset.EOF
            '           Printer.Print Space(10) + Str(RR.Recordset("mHME")) + Space(10) + Right(Space(12) + Format(RR.Recordset("sajia"), "###,###,###.00"), 12)
            '           If Not IsNull(RR.Recordset("sajia")) Then s1 = s1 + RR.Recordset("sajia")
            '           RR.Recordset.MoveNext
            '       Loop
            '       Printer.Print Space(22) + "------------"
            '
            '       Printer.Print Space(22) + Right(Space(12) + Format(s1, "###,###,###.00"), 12)
            '       Printer.EndDoc
            '
            '   End If
            '

            '---------------------------------'παρουσιαση κινήσεων (σουμα apodeiξης)----------------------------------------
780     ElseIf Combo1.List(4) = Combo1.Text Then

            On Error Resume Next

            'Set mydb = OpenDatabase(mDir + "\kinhseis.mdb")
            'Gdb.Execute "DROP TABLE TEMP2"

            On Error GoTo Command1_Click_Err

            'Gdb.Execute "select sum(POSO*TIMM) as SAJIA ,ATIM,HME  INTO TEMP2  from EGGTIM where POSO>0 AND " + synt + " group by ATIM,HME;"
            'RR.RecordSource = "select SAJIA,ATIM,HME from TEMP2 order by HME,ATIM ;"

790         RR.RecordSource = "select sum(PIS*TIMM) as SAJIA ,ATIM,HME  from EGGTIM where PIS>0 AND " + synt + " group by ATIM,HME;"
800         RR.Refresh

810         d1.Open "SELECT SUM(PIS*TIMM) FROM EGGTIM where PIS>0 AND " + synt, Gdb, adOpenForwardOnly, adLockReadOnly
820         DataGrid1.columns(0).FooterText = Format(d1(0), "######0.00")
            'DataGrid1.Columns(1).FooterText = Format(d1(1), "######0.00")
830         DataGrid1.columns(0).alignment = dbgRight
            'DataGrid1.Columns(1).Alignment = dbgRight
            'DataGrid1.Columns(3).width = 4000

            '   If Check1 Then
            '
            '      For K = 1 To 31
            '          Printer.Print K, C50(K), C100(K)
            '      Next
            '      Printer.EndDoc
            '
            '
            '
            '       RR.Recordset.MoveFirst
            '       Printer.Print Space(10) + "Παραστατικό" + Space(10) + "Αξία Παρ/κού"
            '       s1 = 0
            '       Do While Not RR.Recordset.EOF
            '           Printer.Print Space(10) + RR.Recordset("atim") + Space(10) + Right(Space(12) + Format(RR.Recordset("sajia"), "#########.00"), 12)
            '           If Not IsNull(RR.Recordset("sajia")) Then s1 = s1 + RR.Recordset("sajia")
            '           RR.Recordset.MoveNext
            '       Loop
            '
            '       Printer.Print Space(26) + "------------"
            '
            '       Printer.Print Space(26) + Right(Space(12) + Format(s1, "###,###,###.00"), 12)
            '
            '
            '
            '       Printer.EndDoc
            '
            '   End If

            ' τα καλύτερα κατά ποσότητα ,αξια , ονομα
            '-------------------------------------------------------------------------
840     ElseIf Combo1.List(5) = Combo1.Text Or Combo1.List(1) = Combo1.Text Or Combo1.List(2) = Combo1.Text Then  'kata posothta
850         RR.RecordSource = "select TOP 10 * FROM EGGTIM"
860         RR.Refresh

            On Error Resume Next

870         Gdb.Execute "drop table TEMP2"
880         Start = GetCurrentTime()

            Do
890         Loop Until GetCurrentTime() - Start > 3000

            On Error GoTo Command1_Click_Err

900         d = "select sum(PIS*TIMM) as SAJIA,KODE,sum(PIS) as SPOSO,SUM(XRE*TIMM) AS [ΑΞ ΑΓΟΡΩΝ],SUM(XRE) AS [ΠΟΣ ΑΓΟΡΩΝ]  into TEMP2 from EGGTIM where " + synt + " group by KODE;"
910         Gdb.Execute d

            ' start = GetCurrentTime(): Do: Loop Until GetCurrentTime() - start > 2000
            'RR.RecordSource = "select count(*) as mcount from TEMP2 where SPOSO>0.1;"
            'RR.Refresh
            'Me.Caption = "Πλήθος ειδών " + Str(RR.Recordset("mcount"))
            'RR.RecordSource = "select count(*) as mcount from TEMP2 where KODE LIKE '91%' ;"
            'RR.Refresh

            'Me.Caption = Me.Caption + "   Πλήθος καρτών " + Str(RR.Recordset("mcount"))

            On Error GoTo Command1_Click_Err

920         If IsNumeric(Text11.Text) Then   'ELEGXOS PROMITHEYTI

930             If Combo1.List(1) = Combo1.Text Then  'kata posothta
940                 RR.RecordSource = "select TEMP2.*,EID.PROM,EID.ONO from TEMP2 inner join EID  on TEMP2.KODE=EID.KOD WHERE EID.PROM='" + Text11.Text + "' order by SPOSO desc "
950             ElseIf Combo1.List(5) = Combo1.Text Then  'kata onoma
960                 RR.RecordSource = "select TEMP2.*,EID.PROM,EID.name from TEMP2 inner join EID  on TEMP2.KOD=EID.KOD WHERE EID.PROM='" + Text11.Text + "' order by EID.name;"
                Else
970                 RR.RecordSource = "select TEMP2.*,EID.PROM,EID.name from TEMP2 inner join EID  on TEMP2.KOD=EID.KOD  WHERE EID.PROM='" + Text11.Text + "' order by sajia desc"
                End If

            Else

980             If Combo1.List(1) = Combo1.Text Then  'kata posothta
990                 CC = "[ΠΟΣΟΤΗΤΑ] desc "
1000            ElseIf Combo1.List(5) = Combo1.Text Then  'kata onoma
1010                CC = "EID.ONO  "
                Else
1020                CC = "[ΑΞΙΑ] desc "
                End If

                ' RR.RecordSource = "SELECT * FROM TEMP2"
1030            RR.RecordSource = "select STR(TEMP2.SPOSO,10,2) AS [ΠΟΣΟΤΗΤΑ],STR(TEMP2.SAJIA,10,2) AS [ΑΞΙΑ],TEMP2.KODE,EID.ERG AS [BARCODE],EID.ONO,EID.KODLOG from TEMP2 INNER join EID  on TEMP2.KODE=EID.KOD  order by " + CC
1040            RR.Refresh
                'RR.C
                '  Dim d1 As New ADODB.Recordset
1050            d1.Open "SELECT SUM(SPOSO),SUM(SAJIA) FROM TEMP2", Gdb, adOpenForwardOnly, adLockReadOnly
1060            DataGrid1.columns(0).FooterText = Format(d1(0), "######0.00")
1070            DataGrid1.columns(1).FooterText = Format(d1(1), "######0.00")
1080            DataGrid1.columns(0).alignment = dbgRight
1090            DataGrid1.columns(1).alignment = dbgRight
                'Exit Sub
1100            DataGrid1.columns(3).Width = 4000
                ' Exit Sub
                'DATAGRID1.Columns(0).FooterForeColor = vbBlue
                'DATAGRID1.Columns(0).FooterBackColor = vbYellow
                'RR.RecordSource = "select TEMP2.*,EID.ONO from TEMP2 LEFT join EID  on TEMP2.KODE=EID.ERG  order by EID.ONO;"
                'RR.Refresh
                'Else
                '   RR.RecordSource = "select TEMP2.*,EID.ONO from TEMP2 inner LEFT EID  on TEMP2.KODE=EID.KOD  order by SAJIA desc"
                '  RR.Refresh
                'End If
            End If

            '   If Check1 Then
            '       RR.Recordset.MoveFirst
            '       Printer.Print Space(10) + Text1.Text + " - " + Text2.Text + Space(10)
            '       s1 = 0
            '       Do While Not RR.Recordset.EOF
            '           Printer.Print Tab(10); left(RR.Recordset("name"), 30) _
            '           ; Tab(42); Right(Space(14) + Format(RR.Recordset("sposo"), "###,###,###.00"), 14) _
            '           ; Tab(57); Right(Space(14) + Format(RR.Recordset("sajia"), "###,###,###.00"), 14) _
            '           ; Tab(73); RR.Recordset("ERG")
            '
            '           If Not IsNull(RR.Recordset("sajia")) Then s1 = s1 + RR.Recordset("sajia")
            '           RR.Recordset.MoveNext
            '       Loop
            '       Printer.Print Tab(57); "-------------"
            '
            '       Printer.Print Tab(57); Right(Space(14) + Format(s1, "###,###,###.00"), 14)
            '       Printer.EndDoc
            '
            '   End If

        End If

        'RR.Refresh
        'Dim L As Long
1110    L = 1000

1120    For k = 0 To DataGrid1.columns.Count - 1
1130        L = L + DataGrid1.columns(0).Width
        Next

1140    If L > 10000 Then
1150        L = 10000
        End If

        ' DataGrid1.width = L

1160    RichTextBox1.Text = RR.RecordSource

1170    Me.MousePointer = vbNormal

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Sub createGrid(sthles)
'    Dim k
'    ' sbinei oles tis sthles kai dimioyrgei oses ueleis
'    For k = DataGrid1.Columns.Count To 1 Step -1
'        DataGrid1.Columns.Remove (k - 1)
'    Next
'    For k = 1 To sthles
'        '         dataGrid1..Columns.Add (K - 1)
'    Next
'
'    For k = 1 To sthles
'        '         dataGrid1.Columns(K - 1).Visible = True
'    Next
'
'End Sub

'Private Sub Command2_Click()
'' ΕΝΣΩΜΑΤΩΘΗΚΕ Η ΑΛΛΑΓΗ ΠΟΥ ΕΓΙΝΕ ΓΙΑ ΚΑΙΣΑ ΛΕΜΟΝΙΑ ΤΟΝ ΟΚΤΩΒΡΙΟ
'
''Private Sub Command2_Click()
'Dim file
'Dim synt, A, b, c, d, hm2, mydb, ff, start
'Dim r0, D0, st, counter, ff2, r2, ff3
''GoTo 10
'Dim TMHMa(1 To 10), Fpa(1 To 10), K
'Dim Pos_Fpa, ANS
'
'
'
'
'
'On Error Resume Next
'
'
'
'
'
'
'
'
'
'synt = ""
''Text1.Text = DTPicker1.Value
''Text2.Text = DTPicker2.Value
'
'
'
'
'
'
''If IsDate(Text1.Text) And IsDate(Text2.Text) Then
''   hm2 = DateAdd("d", 1, Text2.Text)
''   synt = "hme>=#" + Format(Text1.Text, "mm/dd/yyyy") + "# and hme<= #" + Format(hm2, "mm/dd/yyyy") + "#"
''Else
''   MsgBox "Οχι σωστές ημερομηνίες"
''   Exit Sub
''End If
'Me.Caption = "1"
'If Option1(0).Value = True Then
'
'    If IsNumeric(Text3.Text) And IsNumeric(Text4.Text) Then
'       synt = IIf(synt = "", "", synt + " and ")
'       synt = synt + " forol >='" + Trim(LTrim(Str(Text3.Text))) + "' and forol<='" + Trim(LTrim(Text4.Text)) + "'"
'    End If
'Else
'    If IsNumeric(Text3.Text) And IsNumeric(Text4.Text) Then
'       synt = IIf(synt = "", "", synt + " and ")
'       synt = synt + " barcode >='" + Trim(LTrim(Str(Text3.Text))) + "' and barcode<='" + Trim(LTrim(Text4.Text)) + "'"
'    End If
'End If
'
'If IsNumeric(Text5.Text) And IsNumeric(Text6.Text) Then
'   synt = IIf(synt = "", "", synt + " and ")
'   synt = synt + " seiriakh>=" + Text5.Text + " and seiriakh<=" + Text6.Text
'End If
'
'synt = IIf(synt = "", "true", synt)
'
'
'
'
'Set ff2 = OpenDatabase(mDir + "\eidh.mdb")
'Me.Caption = "4"
'Set D0 = ff2.OpenRecordset("param")
'D0.MoveFirst
'
'
'
'
'If CDate(Text1.Text) <= D0("d8") Then
'  ANS = MsgBox("Τελευταία ενημέρωση " + Format(D0("d8"), "dd/mm/yyyy") + "Eίστε σίγουρος οτι θέλετε να συνεχίσετε;", vbYesNo)
'  If ANS = vbNo Then Exit Sub
'End If
'D0.Close
'ff2.Close
'
'
'
'
'
'
'
'
'     ' rr.RecordSource = "select *from eggtim where " + synt
'
' file = FreeFile
'
'If Len(Text7.Text) = 0 Then
'   MsgBox "Δεν δηλώσατε όνομα αρχείου"
'End If
'
'Open Text7.Text For Output As #file
'
'
'Me.Caption = "2"
''Set r0 = rr.Recordset
'' Set d0 = Data1.Recordset
'Me.Caption = "3"
'
'Set ff3 = OpenDatabase(mDir + "\kinhseis.mdb")
'
'
'Set r0 = ff3.OpenRecordset("select *from eggtim where " + synt + " ORDER BY HME,ATIM")
'
'
'' το αντιγραφει σσε dbaseIII; format
''Kill "c:\tameia\stat5.dbf"
''ff3.Execute "select  eggtim.* into STAT5  in  'c:\tameia' 'dBase III;'  from eggtim where " + synt
'
'
'Set ff2 = OpenDatabase(mDir + "\eidh.mdb")
'Me.Caption = "4"
'Set D0 = ff2.OpenRecordset("param")
'
'GoSub load_antisToix
'
'
'D0.MoveFirst
'D0.MoveNext
'
'Set r2 = ff2.OpenRecordset("eid")
'Me.Caption = "5"
'r2.Index = "barcode"
'Me.Caption = "6"
'counter = 0
'
'r0.MoveFirst
'Me.Caption = "7"
'Do While Not r0.EOF
'   Me.Caption = r0("barcode")
'   'Me.Caption = "8"
'   r2.Seek "=", r0("barcode")
'   Me.Caption = "9"
'   If r2.NoMatch Then
'     counter = counter + 1
'     If counter = 1 Then
'        Me.Caption = "10"
'        MsgBox "Στην διπλανή λίστα εμφανίζονται οι κωδικοί που δεν υπάρχουν"
'        List1.Visible = True
'      End If
'      List1.AddItem r0("barcode")
'   End If
'   st = Space(78)
'
'   Me.Caption = "11"
'
'
'   'BARCODE
'   Mid$(st, Val(D0("c1")), D0("n1")) = Right(Space(13) + r0("barcode"), D0("n1"))
'
'   ' ΗΜΕΡΟΜΗΝΙΑ
'   Mid$(st, Val(D0("c2")), D0("n2")) = left(Format(r0("HME"), "dd/mm/yyyy") + Space(40), D0("n2"))
'
'   'ΑΡΙΘΜ.ΠΑΡΑΣΤΑΤΙΚΟΥ
'   Mid$(st, Val(D0("c3")), D0("n3")) = Right(Space(10) + Format(Val(r0("atim")), "########"), D0("n3"))
'
'   'ΠΟΣΟΤΗΤΑ
'   If r0("timm") < 0 Then
'      Mid$(st, Val(D0("c4")), D0("n4")) = Right(Space(10) + Format(-Abs(r0("poso")), "#####0.000"), D0("n4"))
'   Else
'      Mid$(st, Val(D0("c4")), D0("n4")) = Right(Space(10) + Format(r0("poso"), "#####0.000"), D0("n4"))
'   End If
'
'
'   ' ΤΙΜΗ ΜΟΝΑΔΟΣ
'       Mid$(st, Val(D0("c5")), D0("n5")) = Right(Space(10) + Format(Abs(r0("timm")), "#####0.00"), D0("n5"))
'
'    ' κατηγορια ΦΠΑ
'   Mid$(st, Val(D0("c6")), D0("n6")) = Right(Space(10) + Format(r0("n1"), "#####0.00"), D0("n6"))
'
'
'   If Not IsNull(D0("c7")) Then  ' ekptosi ΔΡΧ
'      Mid$(st, Val(D0("c7")), D0("n7")) = Right(Space(10) + Format(Abs(r0("n2")), "#####0.00"), D0("n7"))
'   End If
'
'   '------------------------------------------
'   GoSub bres_antistoix
'   If Not IsNull(D0("c8")) Then  ' ποσοστό ΦΠΑ
'      Mid$(st, Val(D0("c8")), D0("n8")) = Right(Space(2) + Format(Pos_Fpa, "00"), D0("n8"))
'   End If
'
'
'   If Not IsNull(D0("c9")) Then  ' αξια ΦΠΑ
'      Mid$(st, Val(D0("c9")), D0("n9")) = Right(Space(10) + Format(Round(Abs(r0("ajia")) * Pos_Fpa / 100, 2), "#####0.00"), D0("n9"))
'   End If
'
'   If Not IsNull(D0("c10")) Then  'ΚΑΘΑΡΗ ΑΞΙΑ ΜΕΤΑ ΤΗΝ ΕΚΠΤΩΣΗ
'      Mid$(st, Val(D0("c10")), D0("n10")) = Right(Space(10) + Format(r0("ajia"), "#####0.00"), D0("n10"))
'   End If
'
'
'   Print #file, st
'   Me.Caption = "12"
'   r0.MoveNext
'
'Loop
'Close #file
'
'
'D0.MoveFirst
'D0.Edit
'D0("d8") = Text2.Text
'D0.Update
'
'
'
'
'
'Exit Sub
'
'
''========================
'bres_antistoix:
'
'For K = 1 To 10
'   If TMHMa(K) = r0("n1") Then
'      Pos_Fpa = Fpa(K)
'      Exit For
'   End If
'Next
'Return
'
'
'
'
'
''===========================
'load_antisToix:
'D0.MoveFirst
'D0.Move 3
'
'For K = 1 To 10
'
'  If IsNull(D0("c" + Format(K, "##"))) Or IsNull(D0("n" + Format(K, "##"))) Then
'     TMHMa(K) = 0: Fpa(K) = 0
'  Else
'     TMHMa(K) = Val(D0("c" + Format(K, "##")))
'     Fpa(K) = D0("n" + Format(K, "##"))
'  End If
'
'Next
'
'
'
'
'Return
'
'End Sub

'End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"

130         SendKeys "{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Command4_Click()
'' αποστολή σε αρχεία DBASE
''-------------------------
'Dim file, DUM
'Dim synt, A, b, c, d, hm2, mydb, ff, start
'Dim r0, D0, st, counter, ff2, r2, ff3
'Dim db As Database, R As Recordset, SEIRA As String
'
'Text1.Text = DTPicker1.Value
'
'  Text2.Text = DTPicker2.Value
'
'
'   Set db = OpenDatabase(mDir + "\eidh")
'        Set R = db.OpenRecordset("param")
'        SEIRA = Trim(R("c3"))
'        R.Close
'
'
'
'
'
'
'
'On Error Resume Next
'
'synt = ""
'
'If IsDate(Text1.Text) And IsDate(Text2.Text) Then
'   hm2 = DateAdd("d", 1, Text2.Text)
'   synt = "hme>=#" + Format(Text1.Text, "mm/dd/yyyy") + "# and hme<= #" + Format(hm2, "mm/dd/yyyy") + "#"
'Else
'   MsgBox "Οχι σωστές ημερομηνίες"
'   Exit Sub
'End If
'Me.Caption = "1"
'If Option1(0).Value = True Then
'
'    If IsNumeric(Text3.Text) And IsNumeric(Text4.Text) Then
'       synt = IIf(synt = "", "", synt + " and ")
'       synt = synt + " forol >='" + Trim(LTrim(Str(Text3.Text))) + "' and forol<='" + Trim(LTrim(Text4.Text)) + "'"
'    End If
'Else
'    If IsNumeric(Text3.Text) And IsNumeric(Text4.Text) Then
'       synt = IIf(synt = "", "", synt + " and ")
'       synt = synt + " barcode >='" + Trim(LTrim(Str(Text3.Text))) + "' and barcode<='" + Trim(LTrim(Text4.Text)) + "'"
'    End If
'End If
'
'If IsNumeric(Text5.Text) And IsNumeric(Text6.Text) Then
'   synt = IIf(synt = "", "", synt + " and ")
'   synt = synt + " seiriakh>=" + Text5.Text + " and seiriakh<=" + Text6.Text
'End If
'
'synt = IIf(synt = "", "true", synt)
'
'
'     ' rr.RecordSource = "select *from eggtim where " + synt
'
' file = FreeFile
'
'If Len(Text7.Text) = 0 Then
'   MsgBox "Δεν δηλώσατε όνομα αρχείου"
'End If
'
'Open Text7.Text For Output As #file
'
'
'Me.Caption = "2"
''Set r0 = rr.Recordset
'' Set d0 = Data1.Recordset
'Me.Caption = "3"
'
'Set ff3 = OpenDatabase(mDir + "\kinhseis.mdb")
'
'
'Set r0 = ff3.OpenRecordset("select *from eggtim where " + synt + " ORDER BY HME,ATIM")
'
'
'' το αντιγραφει σσε dbaseIII; format
'Kill "c:\tameia\stat5.dbf"
'ff3.Execute "select  eggtim.* into STAT5  in  'c:\tameia' 'dBase III;'  from eggtim where " + synt
'
'
'If Combo2.Text = Combo2.List(0) Then ' Internet
'   DUM = Shell("c:\falcon\pkzip.exe " + SEIRA + ".  c:\tameia\stat5.dbf", 6)
'
'   file = FreeFile
'   Open "c:\tameia\send_egg" For Output As #file
'   Print #file, "open www.lagakis.gr"
'   Print #file, "lagakis"
'   Print #file, "392127223"
'   Print #file, "binary"
'   Print #file, "cd xry"
'   Print #file, "send " + SEIRA
'   Print #file, "quit"
'   Close #file
'
'   A = GetCurrentTime()
'   Do While GetCurrentTime() - A < 5000
'   Loop
'
'   DUM = Shell("c:\falcon\ftp.exe -s:c:\tameia\send_egg", 6)
'   MsgBox "Ολοκληρώθηκε η διαδικασία"
'   ff3.Close
'Else
'   DUM = Shell("c:\falcon\pkzip.exe a:5  c:\tameia\stat5.dbf", 6)
'End If
'
'End Sub

Private Sub Command5_Click()

        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>

100     print3_xar UCase(RR.RecordSource), "11000000000", "", 0   ' RR.RecordSource

        Exit Sub

110     With DataGrid1.PrintInfo
            ' Set the page header
120         .PageHeaderFont.Italic = True
130         .PageHeader = "Composers table"

            'Column headers will be on every page
140         .RepeatColumnHeaders = True

            ' Display page numbers (centered)
150         .PageFooter = "\tσελίδα: \p"
            ' Invoke Print Preview
160         .PrintPreview
        End With

        'End If

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>

100     RR.RecordSource = RichTextBox1.Text
110     RR.Refresh

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DATAGRID1_FilterChange()

        'Private Sub DATAGRID1_FilterChange()
        '<EhHeader>
        On Error GoTo DATAGRID1_FilterChange_Err

        '</EhHeader>

        'Gets called when an action is performed on the filter bar

        On Error GoTo errHandler

100     Set ColS = DataGrid1.columns

        Dim c As Integer

110     c = DataGrid1.Col

120     DataGrid1.HoldFields

130     RR.Recordset.Filter = getFilter()

140     DataGrid1.Col = c

150     DataGrid1.EditActive = True

        Exit Sub

errHandler:

160     MsgBox Err.Source & ":" & vbCrLf & Err.Description

        '    Call cmdClearFilter_Click

170     Resume Next

        '<EhFooter>
        Exit Sub

DATAGRID1_FilterChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.DATAGRID1_FilterChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.DATAGRID1_FilterChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DATAGRID1_GroupColMove(ByVal Position As Integer, _
                                   ByVal ColIndex As Integer, _
                                   Cancel As Integer)

        '<EhHeader>
        On Error GoTo DATAGRID1_GroupColMove_Err

        '</EhHeader>

        Dim strSort As String

        Dim Col     As TrueOleDBGrid80.Column

        ' Loop through GroupColumns collection and construct

        ' the sort string for the Sort property of the Recordset

100     For Each Col In DataGrid1.GroupColumns

110         If strSort <> vbNullString Then

120             strSort = strSort & ", "

            End If

130         strSort = strSort & "[" & Col.DataField & "]"

140     Next Col

150     DataGrid1.HoldFields

160     RR.Recordset.sort = strSort

        '<EhFooter>
        Exit Sub

DATAGRID1_GroupColMove_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.DATAGRID1_GroupColMove " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.DATAGRID1_GroupColMove " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Initialize()

        '<EhHeader>
        On Error GoTo Form_Initialize_Err

        '</EhHeader>
        Dim D0, ff2

100     Data1.ConnectionString = gConnect    ' = mDir + "\kinhseis.mdb"

110     DTPicker1.Value = Now
120     DTPicker2.Value = Now

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        '<EhFooter>
        Exit Sub

Form_Initialize_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.Form_Initialize " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.Form_Initialize " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Function getFilter() As String

        'Creates the SQL statement in adodc1.recordset.filter
        'and only filters text currently. It must be modified to filter other data types.
        '<EhHeader>
        On Error GoTo getFilter_Err

        '</EhHeader>

        Dim tmp As String

        Dim N   As Integer

100     For Each Col In ColS

110         If Trim(Col.FILTERTEXT) <> "" Then
120             N = N + 1

130             If N > 1 Then
140                 tmp = tmp & " AND "
                End If

150             If RR.Recordset(Col.ColIndex).Type = 5 Then    '
160                 If InStr(">< >= <=  = ", Left(Col.FILTERTEXT, 1)) > 0 And Len(Col.FILTERTEXT) > 1 And IsNumeric(Right(Col.FILTERTEXT, 1)) Then
170                     tmp = tmp & Col.DataField & Col.FILTERTEXT
                    Else
180                     tmp = tmp & Col.DataField & " >  " & Val(Col.FILTERTEXT)
                    End If

                Else
190                 tmp = tmp & Col.DataField & " LIKE '" & Col.FILTERTEXT & "*'"
                End If

            End If

200     Next Col

210     getFilter = tmp

        '<EhFooter>
        Exit Function

getFilter_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.getFilter " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.getFilter " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command3_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        ' ΠΡΟΣΟΧΗ ΠΡΕΠΕΙ ΝΑ ΜΕΤΑΦΕΡΘΕΙ param apo eidh.mdb -> SQL SERVER
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl

        Dim FF, ff2, D0

        'On Error Resume Next

110     Me.Picture = LoadPicture(gPicture)

120     Me.KeyPreview = True

        'Text1.Text = Format(Now, "dd/mm/yyyy")
        'Text2.Text = Format(Now, "dd/mm/yyyy")

130     Combo1.AddItem "Συνολο αξίας"    '0
140     Combo1.AddItem "Καλύτερα κατά ποσότητα"    '1
150     Combo1.AddItem "Καλύτερα κατά αξία"    '2
160     Combo1.AddItem "Αναλυτικές κινήσεις"    '3
170     Combo1.AddItem "Αθροισμα ανα απόδειξη"    '4
180     Combo1.AddItem "Σύνολα Αλφαβητικά"    '5
190     Combo1.AddItem "Σύνολα ανά ημέρα"    '6
200     Combo1.AddItem "Σύνολα ανά κατηγορία"    '7
210     Combo1.AddItem "Σύνολα ανά ΤΜΗΜΑ(ΦΠΑ)"    '8
220     Combo1.AddItem "Σύνολα ανά MHXANH"    '9
230     Combo1.Text = Combo1.List(0)    ' "Καλύτερα κατά αξία"

240     Combo2.AddItem "Internet"
250     Combo2.AddItem "Δισκέττα"
260     Combo2.Text = Combo2.List(0)

        On Error Resume Next

        Dim d1 As New ADODB.Recordset

270     d1.Open "SELECT * FROM param", Gdb, adOpenForwardOnly, adLockReadOnly

280     Me.Caption = "4"
290     d1.MoveFirst
300     d1.MoveNext

310     If IsNull(d1("filename")) Then
320         Text7.Text = " "
        Else
330         Text7.Text = d1("filename")
        End If

340     d1.Close

350     DTPicker1.Value = Now
360     DTPicker2.Value = Now

        'GEMISMA MULTI SELECT LISTBOX
        Dim R As New ADODB.Recordset

370     R.Open "select POL,TITLOS,EIDOS,AJIA_APOU from PARASTAT where POL=1 ORDER BY POL,TITLOS;", Gdb, adOpenDynamic, adLockOptimistic
380     List2.Clear
390     R.MoveFirst

        Dim k As Integer

400     k = 0

410     Do While Not R.EOF

420         If R("pol") = "1" And (R("AJIA_APOU") = "3" Or R("AJIA_APOU") = "4") Then
430             List2.AddItem R("TITLOS") + Space(1) + R("eidos")
440             List2.Selected(k) = True
450             k = k + 1
            End If

460         R.MoveNext
        Loop

470     R.Close
        'For k = 0 To List2.ListCount - 1
        '           List2.Selected(k) = True
        'Next

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Paint()

    ' rr.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + mDir + "\kinhseis.mdb"
End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 13395, 9600, fh, fw, ft, fl
End Sub

Private Sub Text3_LostFocus()

        '<EhHeader>
        On Error GoTo Text3_LostFocus_Err

        '</EhHeader>

100     Text4.Text = Text3.Text

        '<EhFooter>
        Exit Sub

Text3_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT5.Text3_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT5.Text3_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
