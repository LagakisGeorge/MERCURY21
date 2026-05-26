VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form apot4 
   BackColor       =   &H00FF0000&
   Caption         =   "Τιμοκατάλογοι"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Height          =   4335
      Left            =   -240
      TabIndex        =   13
      Top             =   6720
      Width           =   13095
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4680
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   360
         Width           =   6495
      End
      Begin VB.TextBox txtText3 
         Height          =   375
         Left            =   4680
         TabIndex        =   17
         ToolTipText     =   "Ορίζουμε στον πελατη ποιό σετ  θέλει (num2 πεδίο) στην τιμολόγηση"
         Top             =   720
         Width           =   6495
      End
      Begin VB.CommandButton cmdΚΑΤΑΧΩΡΗΣΗ 
         Caption         =   "ΚΑΤΑΧΩΡΗΣΗ"
         Height          =   360
         Left            =   10320
         TabIndex        =   16
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdΕξοδος 
         BackColor       =   &H80000010&
         Caption         =   "Εξοδος"
         Height          =   375
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid mfgDetail 
         Height          =   2280
         Left            =   360
         TabIndex        =   21
         Top             =   1800
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4022
         _Version        =   393216
         Rows            =   8
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   360
         BackColor       =   16777215
         BackColorFixed  =   -2147483648
         ForeColorFixed  =   8388608
         BackColorSel    =   16777088
         BackColorBkg    =   16777215
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblΑριθμόςΣετ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αριθμός Σετ (Προτιμώμενων) Ειδών"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   900
         Width           =   2655
      End
      Begin VB.Label lblΠροεπιλεγμέναΕίδη 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Προεπιλεγμένα είδη που θα βγάινουν στα παραστατικά για τους πελάτες που έχουν δηλωμένο (πεδίο num2)  αριθμό set"
         Height          =   555
         Left            =   600
         TabIndex        =   18
         Top             =   1440
         Width           =   8970
      End
      Begin VB.Label lblΚριτήριαΕιδών 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κριτήρια ειδών που θα προστεθούν"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   450
         Width           =   2730
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton cmdΔιαγραφήΕίδους 
         Caption         =   "Διαγραφή Είδους από Τιμοκάλογο"
         Height          =   360
         Left            =   5280
         TabIndex        =   22
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CommandButton cmdΑντιγραφήΤιμοκαταλόγου 
         BackColor       =   &H000080FF&
         Caption         =   "Αντιγραφή Τιμοκαταλόγου"
         Height          =   360
         Left            =   9600
         MousePointer    =   2  'Cross
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox TIMOK 
         Height          =   285
         Left            =   2265
         TabIndex        =   8
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox EKPT 
         Height          =   285
         Left            =   5280
         TabIndex        =   7
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5280
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000080FF&
         Caption         =   "Διαγραφή Τιμοκαταλόγου"
         Height          =   360
         Left            =   9600
         MousePointer    =   2  'Cross
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   2532
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Κριτήρια"
         Height          =   360
         Left            =   9600
         TabIndex        =   4
         Top             =   720
         Width           =   2532
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nέo είδoς/Διορθωση στον τιμοκατάλογο"
         Height          =   480
         Left            =   5280
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdΤιμοκατάλογοιΕιδικοί 
         Caption         =   "Τιμοκατάλογοι Ειδικοί"
         Height          =   360
         Left            =   9600
         TabIndex        =   2
         Top             =   240
         Width           =   2532
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
         Bindings        =   "apot4.frx":0000
         Height          =   3960
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   6985
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
         Splits(0).DividerColor=   14215660
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
         AllowUpdate     =   0   'False
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Τιμοκατάλογος Νο"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   300
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Εκπτωση % / Τιμή"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   10
         Top             =   840
         Width           =   1800
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδικός"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Top             =   1440
         Width           =   1800
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1320
      Top             =   7440
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
End
Attribute VB_Name = "apot4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fsql      As String

Dim f_ww  As Long, f_hh As Long '= 12005, hh=15060,


Dim FIsHidden As Boolean

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Dim Xval As String

Private Sub cmdΑντιγραφήΤιμοκαταλόγου_Click()


'insert into your_table (c1, c2, ...)
'select c1, c2, ...
'From your_table
'where ID = 1


If Len(TIMOK.Text) = 0 Then
    MsgBox "Διαλέξτε τον αρχικό κατάλογο"
    TIMOK.SetFocus
    


End If

Dim c As String

c = InputBox("Να αντιγραφεί ο κατάλογος στον ", "")
Dim sql As String
sql = "insert into TIMOKAT ( [KOD],[EKPT],[TIMOK],[ONO] ) SELECT [KOD],[EKPT]," + c + ",[ONO] FROM TIMOKAT WHERE TIMOK=" + TIMOK.Text
Gdb.Execute sql






End Sub

Private Sub cmdΔιαγραφήΕίδους_Click()

Dim ANS As Integer

ANS = MsgBox("Να διαγραφεί το είδος " + Text2.Text + " από τον τιμοκατάλογο " + TIMOK.Text + " ;", vbYesNo)

If ANS = vbYes Then
  Gdb.Execute "delete FROM TIMOKAT  WHERE TIMOK=" + TIMOK.Text + " AND  KOD = '" + Text2.Text + "'"
End If


  Adodc1.RecordSource = "SELECT * FROM TIMOKAT  WHERE TIMOK=" + TIMOK.Text + " AND  KOD LIKE '" + Text2.Text + "%'"
    Adodc1.Refresh


End Sub

Private Sub cmdΕξοδος_Click()
   Unload Me
   
End Sub

Private Sub cmdΤιμοκατάλογοιΕιδικοί_Click()
  apot4PERP.SHOW
  
End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim a, B, c, R As New ADODB.Recordset

        Dim n1 As Long, n2 As Long

100     If Len(Text1.Text) = 0 Then
110         MsgBox "διαλεξτε κριτήρια για τον τιμοκατάλογο"
        Else
            
            R.Open " SELECT * FROM EID WHERE " + Text1.Text, Gdb, adOpenDynamic, adLockOptimistic
            If R.EOF Then
                MsgBox "Δεν βρέθηκαν τέτοια είδη"
                R.Close
                
                Exit Sub
            End If
            R.Close
            
            
            
            
            'On Error Resume Next

120         ' Gdb.Execute " DROP TABLE DOKTIMOKAT"
130         ' Gdb.Execute "SELECT KOD," + EKPT.Text + " AS EKPT, " + TIMOK.Text + " AS TIMOK,ONO  INTO DOKTIMOKAT FROM EID WHERE " + Text1.Text
140         Gdb.Execute "DELETE FROM TIMOKAT WHERE TIMOK=" + TIMOK.Text + " AND KOD IN ( SELECT KOD FROM EID WHERE " + Text1.Text + ") ", n2
150         Gdb.Execute "INSERT INTO TIMOKAT(KOD,EKPT,TIMOK,ONO) SELECT KOD," + Replace(EKPT.Text, ",", ".") + " AS EKPT, " + TIMOK.Text + " AS TIMOK,ONO FROM EID WHERE " + Text1.Text, n1

            If n2 > 0 Then
155             MsgBox "Δημιουργήθηκαν " + str(n1) + " νέα είδη στον τιμοκατάλογο " + Chr(13) + " Διαγράφηκαν " + str(n2) + " είδη που υπήρχαν ήδη"
            Else
158             MsgBox "Δημιουργήθηκαν " + str(n1) + " νέα είδη στον τιμοκατάλογο "
            End If
        End If



160     Adodc1.RecordSource = "SELECT * FROM TIMOKAT  WHERE TIMOK=" + TIMOK.Text + " AND  KOD LIKE '" + Text2.Text + "%'"   '"SELECT * FROM TIMOKAT  WHERE TIMOK=" + TIMOK.Text
170     Adodc1.Refresh
        TIMOK.SetFocus

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.apot4.Command1_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot4.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then

120         SendKeys "%"

130         SendKeys "{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot4.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot4.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>
        Dim ANS As Integer

100     If Val(TIMOK.Text) = 0 Then
110         MsgBox "διαλέξτε τιμοκατάλογο"

            Exit Sub

        End If

120     ANS = MsgBox("Είσαι σίγουρος για την διαγραφή του τιμοκαταλόγου Νο " + TIMOK.Text + " ;", vbYesNo)

130     If ANS = vbYes Then
140         Gdb.Execute "DELETE FROM TIMOKAT WHERE TIMOK=" + TIMOK.Text + ""
150         MsgBox "διεγράφη"
        End If

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot4.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot4.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>
        Dim N As Integer

100     N = 0
    
        Dim M_num1

110     M_num1 = FINDPARAMETROI(1, "APOT2", "F_num1", "Αριθ1", "Ετικέτα κειμ.1")
    
        On Error Resume Next

120     If Not FIsHidden Then

130         With KRITHRIA.Grid1

140             N = N + 1: .TextMatrix(N, 0) = "Περιγραφή": .TextMatrix(N, 3) = "EID.ONO": .TextMatrix(N, 4) = "c"
150             N = N + 1: .TextMatrix(N, 0) = "Οικογένεια": .TextMatrix(N, 3) = "EID.AEG": .TextMatrix(N, 4) = "n"
160             N = N + 1: .TextMatrix(N, 0) = "Υποοικογένεια": .TextMatrix(N, 3) = "EID.KODLOG": .TextMatrix(N, 4) = "n"
170             N = N + 1: .TextMatrix(N, 0) = "Υπόλοιπο": .TextMatrix(N, 3) = "EID.POS": .TextMatrix(N, 4) = "n"
180             N = N + 1: .TextMatrix(N, 0) = "Υπόλ.Αποθ1": .TextMatrix(N, 3) = "EID.POS01": .TextMatrix(N, 4) = "n"
190             N = N + 1: .TextMatrix(N, 0) = "Υπόλ.Αποθ2": .TextMatrix(N, 3) = "EID.POS02": .TextMatrix(N, 4) = "n"
200             N = N + 1: .TextMatrix(N, 0) = "Προμηθευτής": .TextMatrix(N, 3) = "EID.PROM": .TextMatrix(N, 4) = "c"
210             N = N + 1: .TextMatrix(N, 0) = "BARCODE": .TextMatrix(N, 3) = "EID.ERG": .TextMatrix(N, 4) = "c"

220             N = N + 1: .TextMatrix(N, 0) = "Τιμή Αγοράς": .TextMatrix(N, 3) = "EID.XTI": .TextMatrix(N, 4) = "n"
230             N = N + 1: .TextMatrix(N, 0) = "Τιμή Λιανικής": .TextMatrix(N, 3) = "EID.LT5": .TextMatrix(N, 4) = "n"

240             N = N + 1: .TextMatrix(N, 0) = "Τιμή Χονδρικής": .TextMatrix(N, 3) = "EID.LTI": .TextMatrix(N, 4) = "n"

250             N = N + 1: .TextMatrix(N, 0) = "Κατ.ΦΠΑ": .TextMatrix(N, 3) = "EID.FPA": .TextMatrix(N, 4) = "n"

260             N = N + 1: .TextMatrix(N, 0) = "Πόντοι": .TextMatrix(N, 3) = "EID.PONTOI": .TextMatrix(N, 4) = "n"
            
270             N = N + 1: .TextMatrix(N, 0) = M_num1: .TextMatrix(N, 3) = "EID.NUM1": .TextMatrix(N, 4) = "n"

                '.AddItem ("περιγραφή" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "EID.ONO" & Chr(9) & "C" & Chr(9)), 1
                ' .AddItem ("Κατηγορία" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "EID.KODLOG" & Chr(9) & "N" & Chr(9)), 1

                '.AddItem ("" & Chr(9) & "Σύνολα" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(125, "###,###.00") & Chr(9) & Format(132, "###,###.00") & Chr(9))

            End With

280         FIsHidden = True

        End If

290     KRITHRIA.SHOW 1
300     Fsql = MDIForm1.StatusBar1.Panels(6).Text

310     If Len(Text1.Text) > 0 Then
320         Text1.Text = Text1.Text + " or " + Fsql
        Else
330         Text1.Text = Text1.Text + Fsql
        End If

340     MDIForm1.StatusBar1.Panels(6).Text = ""

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot4.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot4.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub EKPT_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub Form_DblClick()

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "APOT4"
110     PARAMETROI.SHOW 1


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command2_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot4.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot4.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
       
        
        f_ww = Val(FINDPARAMETROI(1, "APOT4", "F_WW", "12005", "ΣΥΝΤΕΛ.ΜΕΓΕΘΥΝΣΗΣ ΠΛΑΤΟΥΣ ΟΘΟΝΗΣ(12005)"))
        f_hh = Val(FINDPARAMETROI(1, "APOT4", "F_HH", "15060", "ΣΥΝΤΕΛ.ΜΕΓΕΘΥΝΣΗΣ ΥΨΟΥΣ ΟΘΟΝΗΣ(15060)"))
        
        'Dim f_ww  As Long, f_hh As Long '= 12005, hh=15060,
        
        'ww= 12005, hh=15060,

100     mForm_Load Me, fh, fw, ft, fl

110     Me.KeyPreview = True

120     FIsHidden = False

        '
        '    Dim R As New ADODB.Recordset
        '    '================================ EIDH_DEFAULT ================================
        '    R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'EIDH_DEFAULT'", Gdb, adOpenDynamic, adLockOptimistic
        '    If R(0) = 0 Then
        '       Gdb.Execute "CREATE TABLE EIDH_DEFAULT ( AA INT NOT NULL,KOD VARCHAR(16) NOT NULL,TIMM FLOAT NULL,EKPT FLOAT NULL,ID INT IDENTITY(1,1) NOT NULL ) "
        '    End If
        '  R.Close

        mfgDetail.ColWidth(0) = 2000
        mfgDetail.ColWidth(1) = 4000
        mfgDetail.ColWidth(2) = 2000
        mfgDetail.ColWidth(3) = 1000
        mfgDetail.ColWidth(4) = 1000

        mfgDetail.TextMatrix(0, 0) = "Κωδικός"
        mfgDetail.TextMatrix(0, 1) = "Περιγραφή"
        mfgDetail.TextMatrix(0, 2) = "Τιμή Μονάδος"
        mfgDetail.TextMatrix(0, 3) = "Εκπτωση"
        'On Error GoTo 0
        mfgDetail.TextMatrix(0, 4) = "A/A εμφάνισης"

        ' Allocate space for 100 rows, 4 columns
        '  XX.ReDim 0, 19, 0, 3
        ' Dim row As Long, Col As Integer
        ' The LowerBound and UpperBound properties correspond
        ' to the LBound and UBound functions in Visual Basic.
        ' Hard-coded dimensions can be used instead, if known.
        'For row = XX.LowerBound(1) To XX.UpperBound(1)
        '   For Col = XX.LowerBound(2) To XX.UpperBound(2)
        '      XX(row, Col) = ""   '    " " & row & " " & col
        ' Next Col
        ' Next row
        ' Bind True DBGrid Control to this XArray instance
        ' Set TDBGrid2.Array = XX

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

130     Me.Picture = LoadPicture(gPicture)

140     Adodc1.ConnectionString = gConnect

150     Adodc1.RecordSource = "select TOP 10 * FROM TIMOKAT"

        On Error GoTo create

160     Adodc1.Refresh

        Exit Sub

create:
170     Gdb.Execute "SELECT TOP 1 KOD,120.12 AS EKPT,9 AS TIMOK,ONO INTO TIMOKAT FROM EID"

        '  APO = Now
        '  EOS = Now
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot4.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot4.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Function sumCol(ByRef gr As Control, ByRef S)
'' σουμαρει τις κολόνες του flexgrid και
''τις τυπώνει στην τελευταία σειρά kai tis στοιχιζει δεξιά
''s array με 0 , 1   0=οχι αθροισμα  1=ναι στο αθροισμα
'' επιστρέφουν τα αθροίσματα στο S
''π.χ. το έστειλα με (0,0,1) και επιστρέφει (0,0,2344.23) 2344.23 είναι το άθροισμα της 3ης στήλης
'
'' το καλούμε ως εξής
''Dim sumes, dum
''sumes = Array(0, 0, 1, 1, 1, 1)
''dum = sumCol(Grid1, sumes)
'
'
'
''Dim K As Long
''Dim L As Integer
''
''For L = 0 To UBound(s)
''  If s(L) > 0 Then
''     s(L) = 0
''     For K = 1 To gr.rows - 1
''        s(L) = s(L) + Val(gr.TextMatrix(K, L))
''     Next
''  End If
''Next
''
''Dim syn
''syn = ""  ' & "" & Chr(9) & "Σύνολα" & Chr(9) & Format(sum, "###,###.00") & Chr(9))
''For L = 0 To UBound(s)
''   If s(L) <> 0 Then
''      syn = syn + Format(s(L), "###,###.00") + Chr(9)
''   Else
''      syn = syn + "" + Chr(9)
''   End If
''   gr.ColAlignment(L) = 7 'Right
''Next
''
''gr.AddItem syn
''
''If gr.rows > 10 Then
''  Grid1.TopRow = Grid1.rows - 10
''End If
'
'
'
'
'
'
'End Function

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>  'ww= 12005, hh=15060,

    mForm_Resize Me, 8, f_ww, f_hh, fh, fw, ft, fl   ' h=9060   5-1-14

End Sub

Private Sub KATAXVRHSH_Click()

    Dim N     As Integer

    Dim k     As Integer

    Dim MKOD  As String

    Dim MTIMM As String

    Dim MEKPT As String

    Gdb.Execute "delete from EIDH_DEFAULT where AA=" + str(Val(txtText3.Text))

    Dim AATIM As String

    For k = 0 To 6   ' TDBGrid2.ROWS

        MTIMM = "0"
        MEKPT = "0"

        MKOD = mfgDetail.TextMatrix(k, 0)
        MTIMM = (Replace(mfgDetail.TextMatrix(k, 2), ",", "."))
        AATIM = mfgDetail.TextMatrix(k, 4)
  
        MEKPT = (Replace(mfgDetail.TextMatrix(k, 3), ",", "."))
  
        If Len(Trim(MKOD)) > 0 And Val(MTIMM) > 0 Then
            If Len(MTIMM) = 0 Then MTIMM = "0"
            If Len(MEKPT) = 0 Then MEKPT = "0"
            Gdb.Execute "insert into EIDH_DEFAULT(AA,KOD,TIMM,EKPT,AA_TIMOL) VALUES (" + txtText3.Text + ",'" + MKOD + "'," + str(Val(MTIMM)) + "," + str(Val(MEKPT)) + "," + AATIM + ")"
        End If
  
    Next

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then

        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If
    
    
    
    
    

End Sub

Private Sub Text2_Change()

        '<EhHeader>
        On Error GoTo Text2_Change_Err

        '</EhHeader>

100     If Len(Text2.Text) < 1 Then Exit Sub
110     Adodc1.RecordSource = "SELECT * FROM TIMOKAT  WHERE TIMOK=" + TIMOK.Text + " AND  KOD LIKE '" + Text2.Text + "%'"
120     Adodc1.Refresh

130     Text1.Text = " KOD LIKE '" + Text2.Text + "%'"

        '<EhFooter>
        Exit Sub

Text2_Change_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot4.Text2_Change " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot4.Text2_Change " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then

        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub TIMOK_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub TIMOK_LostFocus()
  TIMOK_LostFocus1
End Sub


Public Sub TIMOK_LostFocus1()
        '<EhHeader>
        On Error GoTo TIMOK_LostFocus_Err

        '</EhHeader>
        On Error Resume Next

100     If Len(TIMOK.Text) < 1 Then Exit Sub
110     Adodc1.RecordSource = "SELECT KOD,ONO,EKPT,TIMOKID FROM TIMOKAT  WHERE TIMOK=" + TIMOK.Text + " ORDER BY KOD"
120     Adodc1.Refresh

        '<EhFooter>
        Exit Sub

TIMOK_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot4.TIMOK_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot4.TIMOK_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtText3_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim N     As Integer

    Dim k     As Integer

    Dim MKOD  As String

    Dim MTIMM As String

    Dim MEKPT As String

    If KeyCode <> 13 Then

        Exit Sub

    End If

    On Error Resume Next

    mfgDetail.Clear
    mfgDetail.TextMatrix(0, 0) = "Κωδικός"
    mfgDetail.TextMatrix(0, 1) = "Περιγραφή"
    mfgDetail.TextMatrix(0, 2) = "Τιμή Μονάδος"
    mfgDetail.TextMatrix(0, 3) = "Εκπτωση"
    mfgDetail.TextMatrix(0, 4) = "A/A ΕΜΦΑΝΙΣΗΣ"

    Dim R As New ADODB.Recordset

    R.Open "SELECT *,EID.ONO FROM EIDH_DEFAULT INNER JOIN EID ON EIDH_DEFAULT.KOD=EID.KOD WHERE AA=" + txtText3.Text, Gdb, adOpenDynamic, adLockOptimistic

    'TDBGrid2.Refresh

    k = 1

    Do While Not R.EOF
        mfgDetail.TextMatrix(k, 0) = R("KOD")
        mfgDetail.TextMatrix(k, 1) = R("ONO")
        mfgDetail.TextMatrix(k, 2) = R("timm")
        mfgDetail.TextMatrix(k, 3) = R("ekpt")
        mfgDetail.TextMatrix(k, 4) = R("AA_TIMOL")
        R.MoveNext
        k = k + 1
  
    Loop

End Sub

Private Sub mfgDetail_KeyPress(KeyAscii As Integer)

    Dim R As New ADODB.Recordset

    '// for shifting cursor to column and row when enter pressed
    '///////////////////////////////////////////////////////////
    If KeyAscii = vbKeyReturn Then
        If mfgDetail.Col + 1 = mfgDetail.ColS Then
            If mfgDetail.row + 1 = mfgDetail.rows Then mfgDetail.row = 0: mfgDetail.Col = 0
            mfgDetail.row = mfgDetail.row + 1
            mfgDetail.Col = 0
        ElseIf mfgDetail.Col = 0 Then
            R.Open "SELECT * FROM EID WHERE KOD='" + mfgDetail.TextMatrix(mfgDetail.row, 0) + "'", Gdb, adOpenDynamic, adLockOptimistic
        
            If Not R.EOF Then
                mfgDetail.TextMatrix(mfgDetail.row, 1) = R("ONO")
            End If

            mfgDetail.Col = mfgDetail.Col + 2
        
        Else
            mfgDetail.Col = mfgDetail.Col + 1
        End If
    End If
    
    '// 8 = backspace . for deleting characters
    '//////////////////////////////////////////
    If KeyAscii = 8 Then
        If Len(Xval) = 0 Then Exit Sub
        Xval = Left$(Xval, Len(Xval) - 1)

        Exit Sub

    End If
    
    '// for storing texts in Variable
    
    '////////////////////////////////
    
    If KeyAscii <> 13 Then Xval = Xval & Chr(KeyAscii)

End Sub

Private Sub mfgDetail_KeyUp(KeyCode As Integer, Shift As Integer)
    
    '// for storing texts in grid
    '////////////////////////////
    mfgDetail.Text = LTrim(Xval)
    
    '// for deleting whole text
    '//////////////////////////
    If KeyCode = vbKeyDelete Then
        mfgDetail.Text = ""
        Xval = ""
    End If

End Sub

Private Sub mfgDetail_RowColChange()
    '// for clearing variable when row or column changes
    '///////////////////////////////////////////////////
    
    '  mfgDetail.TextMatrix(mfgDetail.row, mfgDetail.Col) = a
    Xval = mfgDetail.Text
    
End Sub

