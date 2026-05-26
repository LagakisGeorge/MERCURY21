VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form par7 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   14310
   WindowState     =   2  'Maximized
   Begin VB.TextBox MACText 
      Height          =   315
      Left            =   4050
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   7125
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSAdodcLib.Adodc CTIM 
      Height          =   360
      Left            =   510
      Top             =   6780
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   635
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   960
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Mercvb\reports\timsql1.rpt"
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Μετασχηματισμός"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ΕΞΟΔΟΣ"
      Height          =   615
      Left            =   9000
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc LASTTIMOL 
      Height          =   330
      Left            =   10785
      Top             =   5055
      Width           =   1200
      _ExtentX        =   2117
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Bindings        =   "par7dok.frx":0000
      Height          =   3585
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   6324
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
      Splits(0)._SavedRecordSelectors=   0   'False
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
   Begin MSAdodcLib.Adodc CEGGTIM 
      Height          =   360
      Left            =   480
      Top             =   7155
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   635
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
   Begin MSAdodcLib.Adodc CMEM 
      Height          =   360
      Left            =   2100
      Top             =   7170
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   635
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
Attribute VB_Name = "par7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_TITLOS       '= R("TITLOS")
Dim F_POS_APOU        '= R("POS_APOU")
Dim F_AJIA_APOU        '= R("AJIA_APOU")
Dim f_pel  '= R("pel")
Dim f_pol   '= R("pol") ' 1=POLISEIS 2=AGORES
Dim F_STADIO     '= R("STADIO") ' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
Dim F_rec_p    ' = 0
Dim F_XONDR    '= 0
Dim f_dat
Dim F_ATIM
Dim f_SynSeiresTimologioy
Dim f_psifia_atim
Dim fSCR As Object ' SCRIPT CONTROL
Dim F928 As Integer

Dim F_File_Timologioy As String









 Sub PAR2SHOWING()

 
  Dim FF As PAR2
   
  TDBGrid1.Col = 1
Set FF = New PAR2

FF.Text1 = TDBGrid1.Text

  TDBGrid1.Col = 0
FF.DTPicker1 = CDate(TDBGrid1.Text)
FF.DTPicker2 = CDate(TDBGrid1.Text)

FF.CallFromOther.Caption = "2" ' ΓΙΑ ΝΑ ΚΑΝΕΙ REFRESH STO PAR1.LASTTIMOL

'GIA OLA TA PARASTATIKA
FF.Combo2.Text = "*"



  FF.Show
  

FF.bohuitiko.SetFocus
DoEvents

FF.Command1.SetFocus
 par7.LASTTIMOL.Refresh
  
  
  

  
  
'  SendKeys "~" 'ENTER
'  SendKeys "~" 'enter  {DOWN}"
' SendKeys "~" 'ENTER
 SendKeys "~" 'ENTER
''FF.Command1_Click






End Sub

Function find_parastat(recno As Integer, par As String, ReadWrite As Integer) As Long
'ReadWrite=0   Βρίσκει τον αριθμο παραστατικού          -------------------------
'ReadWrite=1   Aποθηκεύει τον αριθμο παραστατικού + 1   -------------------------
'Βρίσκει τις κατηγορίες φπα

Dim k As Integer
Dim X, Bohuhtiko As String
On Error GoTo MHNYMA


'Dim db As Database
 Dim R As New ADODB.Recordset
  R.Open "SELECT * FROM PARASTAT WHERE EIDOS='" + par + "'", Gdb, adOpenDynamic, adLockOptimistic

If IsNull(R("ARITMISI")) Then
   X = 0
Else
  X = R("ARITMISI")
End If
R.Close

'MEM..MoveFirst
'For k = 1 To 9: f_mfpa(k) = R("fpa" + LTrim(Str(k))): Next


If X = 0 Then
   find_parastat = 0
   Exit Function
End If

If ReadWrite = 0 Then
    R.Open "SELECT * FROM ARITMISI WHERE ID=" + Str(X), Gdb, adOpenDynamic, adLockOptimistic
    If IsNull(R("ARITMISI")) Then
       find_parastat = 0
    Else
       find_parastat = R("ARITMISI") + 1
    End If
    R.Close
    
   
Else
   Gdb.Execute "UPDATE ARITMISI SET ARITMISI= ARITMISI+1  WHERE  ID=" + Str(X)
   
    R.Open "SELECT * FROM ARITMISI WHERE ID=" + Str(X), Gdb, adOpenDynamic, adLockOptimistic
    If IsNull(R("ARITMISI")) Then
       find_parastat = 0
    Else
       find_parastat = R("ARITMISI")
    End If
    R.Close
   
   
   
   
   
   
End If

Exit Function

MHNYMA:
HandleError "Par1:FIND_Parastat"
Resume Next
            



End Function



Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
'CrystalReport1.SelectionFormula = "{eggtim.atim}='a000001'"
'CrystalReport1.Connect = "merc"
'CrystalReport1.Connect = "mercs"
'CrystalReport1.Connect = gConnect
'CrystalReport1.Action = 1


End Sub

Sub Command5_Click()
'ΜΕΤΑΣΧΗΜΑΤΙΖΩ ΤΗΝ ΠΑΡΑΓΓΕΛΙΑ ΣΕ ΑΠΟΔ.ΠΑΡ.ΥΠΗΡΕΣΙΩΝ
'INSERT INTO DOKGILANTZISTIM"
'INSERT INTO DOKGILANTZISEGGTIM"
Dim R As New ADODB.Recordset




Dim DUM

Dim par As New ADODB.Recordset


'ΜΕΤΑΣΧΗΜΑΤΙΖΩ ΤΗΝ ΠΑΡΑΓΓΕΛΙΑ ΣΕ ΑΠΟΔ.ΠΑΡ.ΥΠΗΡΕΣΙΩΝ
'===============================================================================
'ΤΣΙΜΠΑΩ ΑΤΙΜ ΚΑΙ ΗΜΕ ΑΠΟ ΤΟ GRID


Dim mKPE

TDBGrid1.Col = 4
mKPE = TDBGrid1.Text



TDBGrid1.Col = 1
f_matim = TDBGrid1.Text
TDBGrid1.Col = 0
f_dat = CDate(TDBGrid1.Text)
'ΑΝΤΙΓΡΑΦΩ ΤΗΝ ΠΑΡΑΓΓΕΛΙΑ ΣΕ ΠΡΟΣΩΡΙΝΑ TABLES

Dim n
Dim k
Dim SQLFields2



R.Open "select * FROM TIM WHERE left(ATIM,1)='y' and MONTH(HME)=" + Str(Month(Now)) + " AND KPE='" + mKPE + "'", Gdb, adOpenDynamic, adLockBatchOptimistic
If R.EOF Then
Else
   MsgBox "ΕΧΕΙ ΚΟΠΕΙ ΑΠΟΔΕΙΞΗ "
   R.Close
   Exit Sub
End If

mONOMA = "  " '    r(0)
R.Close

R.Open "SELECT TOP 0 *  FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
SQLFields = ""
For k = 0 To R.FIELDS.Count - 1
   If R.FIELDS(k).Name = "ID_NUM" Then
   Else
      SQLFields = SQLFields + R.FIELDS(k).Name + ","
   End If
Next
SQLFields = Left(SQLFields, Len(SQLFields) - 1)

'Exit Sub

' SELECT  TOP 0* INTO DOKGILANTZISTIM         FROM   TIM
' SELECT  TOP 0 *INTO DOKGILANTZISEGGTIM  FROM    EGGTIM


On Error GoTo gilantzis
'On Error GoTo 0


On Error GoTo gilantzis

Gdb.Execute "DELETE FROM DOKGILANTZISTIM"
Gdb.Execute "DELETE FROM DOKGILANTZISEGGTIM"

On Error Resume Next



'Gdb.Execute "SET IDENTITY_INSERT DOKGILANTZISTIM OFF;"

Gdb.Execute "INSERT  INTO DOKGILANTZISTIM(" + SQLFields + ")   SELECT " + SQLFields + " FROM TIM    WHERE ATIM='" + f_matim + " ' AND HME='" + Format(f_dat, "MM/DD/YYYY") + "'", k
Gdb.Execute "INSERT INTO DOKGILANTZISEGGTIM SELECT  *  FROM EGGTIM WHERE ATIM='" + f_matim + " ' AND HME='" + Format(f_dat, "MM/DD/YYYY") + "'", k

MILSEC 500

Gdb.Execute "UPDATE DOKGILANTZISEGGTIM SET FPA=(CASE WHEN FPA=1 THEN " + Str(g_Fpa(1)) + " ELSE " + Str(g_Fpa(2)) + " END )"




'ΑΛΛΑΖΩ ΤΗΝ ΠΑΡΑΓΓΕΛΙΑ ΣΕ ΑΠΥ

 'ΓΙΑ ΝΑ ΠΙΑΝΕΙ ΤΟΝ ΣΩΣΤΟ ΑΡΙΘΜΟ
 ' ΑΛΛΑ ΔΗΜΙΟΥΡΓΗΣΕ ΠΡΟΒΛΗΜΑ  5-10-2009 ΒΑΖΟΝΤΑΣ ΤΟΝ ΙΔΙΟ ΑΡΙΘΜΟ ΣΕ ΠΟΛΛΑ ΠΑΡΑΣΤΑΤΙΚΑ
Dim MATIM
Dim C_DAT

  MATIM = "y" + Format(find_parastat(1, "y", 1), Left("00000000", f_psifia_atim))
  C_DAT = Format(Now, "MM/DD/YYYY")

'Exit Sub

'βρίσκω το είδος που αντιστοιχεί στον μήνα τον τρέχοντα


mKode = Right(Format(Now, "dd/mm/yyyy"), 2) + Mid(Format(Now, "dd/mm/yyyy"), 4, 2)

R.Close

R.Open "select ONO FROM EID WHERE KOD='" + mKode + "'", Gdb, adOpenDynamic, adLockBatchOptimistic
If R.EOF Then
   MsgBox "Ανοίξτε τον κωδικό " + mKode
   R.Close
   Exit Sub
End If

mONOMA = R(0)
R.Close







Gdb.Execute "update  DOKGILANTZISTIM    set ATIM='" + MATIM + "' ,HME='" + C_DAT + "'"
Gdb.Execute "update  DOKGILANTZISEGGTIM set ATIM='" + MATIM + "' ,HME='" + C_DAT + "',KODE='" + mKode + "',ONOMA='" + mONOMA + "';"
  
Gdb.Execute "INSERT  INTO TIM(" + SQLFields + ")   SELECT " + SQLFields + " FROM DOKGILANTZISTIM " '   WHERE ATIM='" + f_matim + " ' AND HME='" + Format(f_dat, "MM/DD/YYYY") + "'", K
Gdb.Execute "INSERT INTO EGGTIM SELECT * FROM DOKGILANTZISEGGTIM"


'ΒΓΑΖΕΙ ΑΠΟ ΤΗΝ ΕΚΚΡΕΜΟΤΗΤΑ ΤΗΝ ΠΑΡΑΓΓΕΛΙΑ
Gdb.Execute "update  TIM    set ART='1' WHERE ATIM='" + f_matim + "' AND HME='" + Format(f_dat, "MM/DD/YYYY") + "'"

'===============================================================================

f_matim = MATIM
f_dat = Now 'CDate(C_DAT)


'Bρισκει το είδος του παραστατικού
'γιατί πολλές φορές το πουλάει σαν λιανική ενώ είναι χονδρική πώληση
' find_eid_parastat

DUM = gfind_eid_parastat(Left(f_matim, 1), F_TITLOS, F_POS_APOU, F_AJIA_APOU, f_pel, f_pol, F_STADIO, F_rec_p, F_XONDR, f_FORM_EKTY)
F_PARAS = F_TITLOS

'εκτύπωση
Dim KN
'r.Close

R.Open "SELECT COUNT(*) FROM DOKGILANTZISEGGTIM", Gdb, adOpenDynamic, adLockOptimistic

Dim M_PEL
Dim M_POSO

KN = f_SynSeiresTimologioy - R(0)  'f_SeiresTimologioy - R.RecordCount
R.Close
R.Open "SELECT * FROM DOKGILANTZISEGGTIM", Gdb, adOpenDynamic, adLockOptimistic
R.MoveFirst
M_PEL = R("PELKOD")
M_POSO = R("MIK_AJIA")

For k = 1 To KN
   R.AddNew: R.Update
Next
R.Close






'=======================ΕΝΗΜΕΡΩΝΩ ΤΟ MEM ΜΕ ΤΑ ΣΤΑΘΕΡΑ ΣΤΟΙΧΕΙΑ ========
Dim fmem As New ADODB.Recordset
On Error Resume Next
Gdb.Execute "DROP TABLE MEMREP"
Gdb.Execute "SELECT TOP 1  C1,C2,C3,C4,BANKS,PICT,PAR1,PAR2,PAR3,PAR4,PAR5,MEMO  INTO MEMREP FROM MEM"

 fmem.Open "SELECT COUNT(*) FROM MEMREP", Gdb, adOpenDynamic, adLockOptimistic
If fmem(0) = 0 Then
    fmem.AddNew
    fmem.Update
End If
fmem.Close

fmem.Open "SELECT * FROM MEMREP", Gdb, adOpenDynamic, adLockOptimistic
fmem("c1") = "ΑΠΟΔΕΙΞΗ ΠΑΡΟΧΗΣ ΥΠΗΡΕΣΙΩΝ" ' Left(F_PARAS, Len(F_PARAS) - 1) 'parastatiko
fmem("c2") = "" ' SKOPOS.Text 'skopos
fmem("c3") = "" 'FORTOSH.Text 'fortosi
fmem("c4") = "" ' PROORISMOS.Text 'paradosi
fmem("banks") = "" ' Combo4.Text 'τροποσ πληρωμης
fmem("pict") = Time$
fmem("par1") = "" 'Text1.Text
fmem("par2") = "" 'SEIRA.Text
fmem("par3") = "" ' AYTOKINHTO.Text
fmem("par5") = "" ' MATIM2
fmem("par6") = Format(DTPicker1.Value, "dd/mm/yyyy")

fmem("MEMO") = "" ' parat.Text
'If Combo4.Text = Combo4.List(0) Then ' METRITOIS
'   fmem("par4") = Format(PEL.Recordset("TYP"), "#####.00")
'Else
'   fmem("par4") = Format(PEL.Recordset("TYP") - Val(SynoloG.Caption), "#####.00")
'End If
fmem.Update


























'========= ΠΡΟΣΟΧΗ ΤΥΠΩΝΕΙ ΜΟΝΟ ΜΕ PRINTSQLCRYSTAL =======================

f_FORM_EKTY = 90   '  ME SQL SERVER H EKTYPOSI
If f_FORM_EKTY > 0 Then ' ---------------------------------------------


If f_FORM_EKTY >= 90 Then ' φορμα τύπου f90.txt  '====================================
        ' DUM = PAR1.printCrystal(f_matim, f_dat) 'δημιουργει τα προσωρινά αρχεία
        DUM = print_timol(f_matim, f_dat)
Else ' crystal reports timol1.rpt(σφραγίδα)   timol1an.rpt (αντίγραφο)
   
   If f_FORM_EKTY > 0 Then '      PEL.Recordset("eidos") = "e" Then  ' EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE


        ' On Error GoTo 0 ' na sbhstei 15-6-2006
'      If f_FORM_EKTY = 15 Then     '  ME SQL SERVER H EKTYPOSI
         On Error GoTo 0
'         DUM = PrintSqlCrystal(f_matim, CDate(f_dat))    'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
'      Else  ' ME MDB
'          DUM = PAR1.printCrystal(f_matim, CDate(f_dat)) 'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
'      End If
'      If DUM < -900 Then
'         MsgBox "αδυνατη η εκτύπωση" + Chr(13) + "Διορθώστε το παραστατικό για να ξανατυπωθεί"
'         Exit Sub
'      End If
'
'    Dim xxx As Printer
'    For Each xxx In Printers
'       If InStr(UCase(xxx.DeviceName), "DESKJET") > 0 Then
'           ' Set printer as system default.
'           Set Printer = xxx
'           ' Stop looking for a printer.
'           Exit For
'        End If
'     Next
'
'    Dim ANS As Integer
'    ANS = MsgBox("Σε οθόνη το τιμολόγιο;", vbYesNo)
'    If ANS = vbYes Then
'        PAR1.CrystalReport1.Destination = crptToWindow
'    End If
'
'
'    If f_FORM_EKTY = 12 Then     '     λιανικη
'
'       PAR1.CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"
'
'      ElseIf f_FORM_EKTY = 13 Then     '     λιανικη
'       PAR1.CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol3.rpt"
'
'
'
'    ElseIf f_FORM_EKTY = 15 Then     '     λιανικη
        CrystalReport1.ReportFileName = "C:\Mercvb\reports\timsql1.rpt"
        CrystalReport1.Connect = gConnect

       ' CrystalReport1.Action = 1
'
'    Else
'       PAR1.CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
'    End If
'
'    If f_FORM_EKTY = 10 Then
'       PAR1.CrystalReport1.Action = 1
'    ElseIf f_FORM_EKTY = 20 Then
'
'       PAR1.CrystalReport1.Destination = crptToWindow
'       PAR1.CrystalReport1.Action = 1
'    ElseIf f_FORM_EKTY = 18 Then
'       PAR1.CrystalReport1.Action = 1
'       PAR1.CrystalReport1.Action = 1
'       PAR1.CrystalReport1.Action = 1
'    Else
'       PAR1.CrystalReport1.Action = 1
'       PAR1.CrystalReport1.Action = 1
'    End If
 
 
 End If   ' If f_FORM_EKTY > 0 Then '  EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
   
End If  ' If ektyp.Value = 1 Then  '========================================



End If  '-------------------------------------------------------------























LASTTIMOL.Refresh




R.Close
R.Open "SELECT TOP 0 * FROM EGG", Gdb, adOpenDynamic, adLockBatchOptimistic

On Error GoTo 0


'R.AddNew
'
'   R("UserID") = gUserId
'   R("hme") = Format(Now, "DD/MM/YYYY") ' CDate(Format(DTPicker1.value, "DD/MM/YYYY")) ' DTPicker1.Value
'   R("apa") = Val(mID$(f_matim, 2, 6))
'   R("atim") = f_matim ' Left(f_matim, 1 + f_psifia_atim)
'   R("ait") = Left("ΑΠΟΔ.ΠΑΡ.ΥΠΗΡ", 20)
'   R("XREOSI") = 0: R("PISTOSI") = 0 ' ΑΠΟΦΥΓΗ ΝULL
'   R("xre") = M_POSO '  f_suma
'   If f_pel = "1" Then R("XREOSI") = M_POSO ' f_suma
'   R("eidos") = "e" '  PEL.Recordset("eidos")
'   R("kod") = M_PEL ' PEL.Recordset("kod")
'On Error GoTo 0
'R.Update

Dim sql As String
'sql = "insert INTO TIM (USERID,HME,APA,ATIM,AIT,XREOSI,PISTOSI,XRE,EIDOS,KOD) VALUES ("
'sql = sql + "" + Str(gUserId) + ","
'sql = sql + "'" + Format(Now, "MM/DD/YYYY") + "',"
'sql = sql + "" + mID$(f_matim, 2, 6) + ","
'sql = sql + "'" + Left("ΑΠΟΔ.ΠΑΡ.ΥΠΗΡ", 20) + "',"
'sql = sql + "" + Str(M_POSO) + ",0,"  'XREOSI   PISTOSI
'sql = sql + "" + Str(M_POSO) + ","  'XRE
'sql = sql + "'e',"
'sql = sql + "'" + M_PEL + "')"

'Gdb.Execute sql, n

'sql = "insert INTO TIM (USERID,HME,APA,ATIM,AIT,XREOSI,PISTOSI,XRE,EIDOS,KOD) VALUES ("
'sql = sql + "" + Str(gUserId) + ","
'sql = sql + "'" + Format(Now, "MM/DD/YYYY") + "',"
'sql = sql + "" + mID$(f_matim, 2, 6) + "," 'apa
'sql = sql + "'" + f_matim + "'," 'atim
'sql = sql + "'" + Left("ΑΠΟΔ.ΠΑΡ.ΥΠΗΡ", 20) + "',"
'sql = sql + "" + Str(M_POSO) + ",0,"  'XREOSI   PISTOSI
'sql = sql + "" + Str(M_POSO) + ","  'XRE
'sql = sql + "'e',"
'sql = sql + "'" + M_PEL + "')"
'
'Gdb.Execute sql, n
'


sql = "insert INTO EGG (USERID,HME,APA,ATIM,AIT,XREOSI,PISTOSI,XRE,EIDOS,KOD) VALUES ("
sql = sql + "" + Str(gUserId) + ","
sql = sql + "'" + Format(Now, "MM/DD/YYYY") + "',"
sql = sql + "" + Mid$(f_matim, 2, 6) + "," 'apa
sql = sql + "'" + f_matim + "'," 'atim
sql = sql + "'" + Left("ΑΠΟΔ.ΠΑΡ.ΥΠΗΡ", 20) + "',"
sql = sql + "" + Str(M_POSO) + ",0,"  'XREOSI   PISTOSI
sql = sql + "" + Str(M_POSO) + ","  'XRE
sql = sql + "'e',"
sql = sql + "'" + M_PEL + "')"

Gdb.Execute sql, n

















Exit Sub

gilantzis:
Gdb.Execute "SELECT  TOP 0* INTO DOKGILANTZISTIM         FROM   TIM"
Gdb.Execute "SELECT  TOP 0 *INTO DOKGILANTZISEGGTIM  FROM    EGGTIM"

Resume Next












End Sub

Private Sub Form_Load()

F_File_Timologioy = FindParametroi("PAR1", "F_File_Timologioy", "c:\print", "Α.12 Αρχείο εκτύπωσης παραστατικού")

f_psifia_atim = Val(FindParametroi("PAR1", "F_PSIFIA_ATIM", "5", "Γ. 9. ΨΗΦΙΑ ΑΡΙΘΜΟΥ ΤΙΜΟΛΟΓΙΟΥ (5)"))
f_SynSeiresTimologioy = Val(FindParametroi("PAR1", "f_SynSeiresTimologioy", "25", "Β. 3 Σειρές ΤΙΜ.ΠΩΛ ΚΕΝΕΣ+ΓΡΑΜ"))
    
    
 
F928 = Val(FindParametroi("PAR1", "F928", "0", "Α.11 Εκτύπωση φόρμας 1=928 0=437"))

F_File_Timologioy = FindParametroi("PAR1", "F_File_Timologioy", "c:\print", "Α.12 Αρχείο εκτύπωσης παραστατικού")
F_File2_Timologioy = FindParametroi("PAR1", "F_File2_Timologioy", "c:\print", "Α.12b Αρχείο εκτύπωσης παραστ.Υποκ")

F_Type_File_Timologioy = Val(FindParametroi("PAR1", "F_Type_File_Timologioy", "1", "Α.13 1=εκτύπωση αρχείου παρ/κού 0=Οχι"))
F_Type2_File_Timologioy = Val(FindParametroi("PAR1", "F_Type2_File_Timologioy", "1", "Α.13b 1=εκτύπ.αρχείου παρ/κού υποκ. 0=Οχι"))

   
    
    
    
    
    par7.KeyPreview = True
    
    par7.Picture = LoadPicture(gPicture)
    
    
'ΔΕΙΧΝΕΙ ΤΙΣ ΠΑΡΑΓΓΕΛΙΕΣ ΠΕΛΑΤΟΥ ΜΕ ΗΜΕΡΟΜΗΝΙΑ 1/1
'ΟΙ ΟΠΟΙΕΣ ΔΕΝ ΕΧΟΥΝ ΤΙΜΟΛΟΓΗΘΕΙ
LASTTIMOL.ConnectionString = gConnect
Dim sql

sql = "SELECT  HME,ATIM,AJI AS [AΞΙΑ],PEL.EPO AS [ΕΠΩΝΥΜΙΑ],PEL.KOD AS [ΚΩΔΙΚΟΣ] "
sql = sql + " FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
sql = sql + " WHERE LEFT(ATIM,1)='a' and day(HME)=1 AND MONTH(HME)=1 AND (ART IS NULL OR ART='') " ' AND (LEFT(ART,1)<>'1'"
sql = sql + " ORDER BY PEL.EPO "
' On Error Resume Next

LASTTIMOL.RecordSource = sql

LASTTIMOL.Refresh
TDBGrid1.MarqueeStyle = 3  'SEIRA MARKARISMENH
TDBGrid1.Splits(0).AlternatingRowStyle = True 'ENALAX XROMATA SEIRON
TDBGrid1.Splits(0).ExtendRightColumn = True
TDBGrid1.Columns(1).Width = 750


End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then

'MDIForm1.runt.Item(0).Caption = "Παρουσίαση"
'MDIForm1.runt.Item(1).Caption = "Διόρθωση"
On Error Resume Next

   Call Me.PopupMenu(MDIForm1.par1menu)
  ' LASTTIMOL.Refresh
   
   
' μεσω drop down menu
'MdiForm1
   
   
   
'ΕΚΤΥΠΩΣΗ
'   par1.diortosis
'  DoEvents
'  par1.Command5_Click
   
   
'ΔΙΟΡΘΩΣΗ
 'par1.diortosis
   
'ΠΑΡΟΥΣΙΑΣΗ
'par1.PAR2SHOWING
   
'ΜΕΤΑΣΧΗΜΑΤΙΣΜΟΣ
'par1.metasx.value = vbChecked
'   par1.diortosis
   
   
   
   
End If


End Sub

Function gfind_eid_parastat(ByVal EIDOS, F_TITLOS, F_POS_APOU, F_AJIA_APOU, f_pel, f_pol, F_STADIO, F_rec_p, F_XONDR, f_FORM_EKTY)

  Dim R As New ADODB.Recordset

  R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
  R.MoveFirst
  Do While Not R.EOF
     If EIDOS = R("EIDOS") Then
       F_TITLOS = R("TITLOS")
       F_POS_APOU = R("POS_APOU")
       F_AJIA_APOU = R("AJIA_APOU")
       f_pel = R("pel")
        f_FORM_EKTY = R("FORM_EKTYP")
       f_pol = R("pol") ' 1=POLISEIS 2=AGORES
       F_STADIO = R("STADIO") ' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
       If IsNull(R("GEF_P")) Then
           F_rec_p = 0
       Else
           F_rec_p = R("GEF_P") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
       End If
       If IsNull(R("XONDR")) Then
           F_XONDR = 0
       Else
           F_XONDR = R("XONDR") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
       End If
     End If
     R.MoveNext
  Loop
R.Close

  
End Function




Function PrintSqlCrystal(MATIM, char_date)
'------------------------------------------------------------------------
'δημιουργεί τα αρχεία TIM  , EGGTIM , MEM    ΣΤΟ  c:\mercvb\reports\reports.mdb"
'ΓΙΑ ΤΗΝ ΕΚΤΥΠΩΣΗ ΤΟΥ ΤΙΜΟΛΟΓΙΟΥ

Dim sql As String, fDB As Database, d1 As Date, d2 As Date
Dim synt As String, db As Database
Dim DUM
Dim REGGTIM As New ADODB.Recordset
Dim rtim As New ADODB.Recordset
Dim R As New ADODB.Recordset
Dim k As Integer
Dim KN As Integer

'-----------  next lines are OK --------------------------

On Error GoTo MHNYMA

d1 = CDate(char_date)
d2 = DateAdd("d", 1, d1)

'Set db = OpenDatabase("c:\mercvb\reports\reports.mdb")
'If gConnect = "Access" Then
'  Set fDB = OpenDatabase(gDir, False, False)
'
'Else
'  Set fDB = OpenDatabase(gDir, False, False, gConnect)
'End If
'                  arn_enhm_apot

'=====   EGGTIM  ======


'On Error Resume Next
'db.Execute "delete *from eggtim"
10




'TO   round(TIMM*POSO*(100-EKPT)/100,2)  ANTIKATASTAUHKE ME TO KAU_AJIA  4-1-2006


'sql = sql + "MONA,EKPT "
'sql = sql + " from EGGTIM where APOT=" + Str(Val(left(APOTH1.Text, 2))) + " and  left(ATIM," + Str(F_PSIFIA_ATIM + 1) + ")='" + left(mAtim, F_PSIFIA_ATIM + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' ORDER BY PROOD_AJ;"
'REGGTIM.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

'copy to access table eggtim
'Set R = db.OpenRecordset("eggtim")
'REGGTIM.MoveFirst
'Do While Not REGGTIM.EOF
  
 '  R.AddNew
  ' For K = 0 To REGGTIM.FIELDS.Count - 1
   '  R(K) = REGGTIM(K)
   'Next
   'R.Update
   'REGGTIM.MoveNext
'Loop



233
sql = "select FPA,ATIM,HME as [Ημερομ],KODE as [Kωδ_Είδ],ONOMA as [Περιγραφή],"

If F_XONDR = 0 Then 'LIANIKH
   sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],MIK_AJIA as [Αξία],"
Else
   sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],KAU_AJIA as [Αξία],"
End If
sql = sql + "MONA,EKPT "

sql = sql + " INTO EGGTIMREP from EGGTIM where APOT=1 and  left(ATIM," + Str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' ORDER BY PROOD_AJ;"

On Error Resume Next

' GoTo 233

Gdb.Execute "DROP TABLE TIMREP", k
Gdb.Execute "DROP TABLE EGGTIMREP", k
Gdb.Execute sql, k

Exit Function
'GoTo 10
On Error GoTo 0


On Error Resume Next
R.Open "SELECT * FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly
'
''copy to access table eggtim
'Set R = db.OpenRecordset("eggtim")
'REGGTIM.MoveFirst
'Do While Not REGGTIM.EOF
'
'   R.AddNew
'   For K = 0 To REGGTIM.FIELDS.Count - 1
'     R(K) = REGGTIM(K)
'   Next
'   R.Update
'   REGGTIM.MoveNext
'Loop
'
'
'
'REGGTIM.Close

Gdb.Execute "update EGGTIMREP set FPA=" + Str(g_Fpa(1)) + " where FPA=1"
Gdb.Execute "update EGGTIMREP set FPA=" + Str(g_Fpa(2)) + " where FPA=2"
Gdb.Execute "update EGGTIMREP set FPA=" + Str(g_Fpa(3)) + " where FPA=3"
Gdb.Execute "update EGGTIMREP set FPA=" + Str(g_Fpa(4)) + " where FPA=4"
Gdb.Execute "update EGGTIMREP set FPA=0 where FPA=5 "
R.Open "SELECT COUNT(*) FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly
KN = f_SeiresTimologioy - R(0)  'f_SeiresTimologioy - R.RecordCount
R.Close





'On Error GoTo 0

'R.Close
R.Open "SELECT * FROM EGGTIMREP", Gdb, adOpenDynamic, adLockOptimistic

'R.Open "SELECT * FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly



' On Error GoTo 0
For k = 1 To KN
   R.AddNew: R.Update
Next
R.Close








20
synt = " left(ATIM," + Str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "'"


' αυτο το κομματι δουλευei με odbc  ================================
'Dim db9 As Database
'db.Execute "drop table tim"
'Set db9 = OpenDatabase("", False, False, gConnect)
'SQL = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL "
'SQL = SQL + " into tim in 'c:\mercvb\reports\reports.mdb'  from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>=#" + Format(d1, "mm/dd/yyyy") + "# and HME<#" + Format(d2, "mm/dd/yyyy") + "#;"
'db9.Execute SQL
'GoTo 22
' αυτο το κομματι δουλευei με odbc   ==================================

Gdb.Execute "DROP TABLE TIMREP2"
'========  TIM ==========
sql = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL "
sql = sql + " INTO TIMREP2 from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"

Gdb.Execute sql


'=============  KENES SEIRES NA GEMISEI EGGTIM  ===================

'Set meggtim = db.OpenRecordset("eggtim")
Dim SAJIA As Single


R.Open "SELECT SUM([Αξία]) FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly


SAJIA = R(0)
R.Close

'Set R = db.OpenRecordset("SELECT * FROM TIM")
R.Open "SELECT * FROM TIMREP2", Gdb, adOpenForwardOnly, adLockReadOnly
   
Dim MATIM2
MATIM2 = R("ATIM")

   
   
If F_XONDR = 0 Then 'LIANIKH
    db.Execute "UPDATE TIMREP2  SET AJI=" + Str(SAJIA) + " , AJIA=" + Str(SAJIA)
   
    db.Execute "UPDATE EGGTIMREP SET [Tιμ_Μον]=[Tιμ_Μον]*(1+fpa/100)"
'  db.Execute "UPDATE EGGTIM SET [Αξία]=round([Αξία],2)"
  
    db.Execute "UPDATE TIMREP2 Set aj1=0,aj2=0,aj3=0,aj4=0,aj5=0,fpa1=0,fpa2=0,fpa3=0,fpa4=0"
   
Else
     If Abs(SAJIA - (R("AJ1") + R("AJ2") + R("AJ3") + R("AJ4") + R("AJ5"))) > 0.005 Then
       MsgBox "Aσυμφωνία παραστατικού. Διαγράψτε το παραστατικό και ξαναπεράστε το"
        End
     End If
   
End If



'=======================ΕΝΗΜΕΡΩΝΩ ΤΟ MEM ΜΕ ΤΑ ΣΤΑΘΕΡΑ ΣΤΟΙΧΕΙΑ ========

On Error GoTo 0
Dim fmem As New ADODB.Recordset
Gdb.Execute "DROP TABLE MEMREP"
Gdb.Execute "SELECT TOP 1  C1,C2,C3,C4,BANKS,PICT,PAR1,PAR2,PAR3,PAR4,PAR5,MEMO  INTO MEMREP FROM MEM"

 fmem.Open "SELECT COUNT(*) FROM MEMREP", Gdb, adOpenDynamic, adLockOptimistic
If fmem(0) = 0 Then
    fmem.AddNew
    fmem.Update
End If

fmem.Close


fmem.Open "SELECT * FROM MEMREP", Gdb, adOpenDynamic, adLockOptimistic
'fmem.Edit
fmem("c1") = F_TITLOS ' Left(F_PARAS, Len(F_PARAS) - 1) 'parastatiko
On Error Resume Next
fmem("c2") = R("SKOPOS") ' .Text 'skopos
fmem("c3") = R("FORTOSH")  '.Text 'fortosi
fmem("c4") = R("PROORISMOS")  '.Text 'paradosi
fmem("banks") = R("TRP") ' Combo4.Text 'τροποσ πληρωμης
fmem("pict") = Time$
fmem("par1") = Text1.Text
fmem("par2") = SEIRA.Text
fmem("par3") = AYTOKINHTO.Text
fmem("par5") = MATIM2
fmem("par6") = Format(DTPicker1.Value, "dd/mm/yyyy")

fmem("MEMO") = parat.Text
If Combo4.Text = Combo4.List(0) Then ' METRITOIS
   fmem("par4") = Format(PEL.Recordset("TYP"), "#####.00")
Else
   fmem("par4") = Format(PEL.Recordset("TYP") - Val(SynoloG.Caption), "#####.00")
End If
fmem.Update




Gdb.Execute "SELECT TIMREP2.* , MEMREP.*  INTO TIMREP FROM TIMREP2,MEMREP"

Exit Function

MHNYMA:
HandleError "Par1:PrintCrystal"
Resume Next
            


End Function



Function ektyp_forma(ByVal arxeio As String, ByVal atim As String, ByVal hme As Date)
' arxeio π.χ. f90.txt  η φόρμα της εκτύπωσης

' atim αριθμός παραστατικού
' hme  ημερομηνία
 
 


Dim DUM
 Dim eField(1 To 50), ePict(1 To 50), eProw(1 To 50), ePcol(1 To 50), eFonts(1 To 50)
 Dim mField(1 To 50), mPict(1 To 50), mProw(1 To 50), mPcol(1 To 50), mFonts(1 To 50)
 Dim sField(1 To 50), sPict(1 To 50), sProw(1 To 50), sPcol(1 To 50), sFonts(1 To 50)
 Dim m_No_of_seir, I, k, m, C
 Dim m_seir_synol, m_seir_eid, sf
 Dim A(1 To 250), ar_ped(1 To 250)
 Dim xa(1 To 250), m_m, npic
 Dim m1
' dim gm_str(1 To 250)
Dim db As Database
Dim TIM As Recordset
Dim PEL As Recordset
Dim EGGTIM As Recordset
Dim m_entol_ektyp
Dim xart11, kod_tim
Dim ejodos
Dim e, mM, S
Dim ektypoths
Dim SELIDA
Dim m_syn_row, syn_row

On Error GoTo MHNYMA



'cEGGTIM.RecordSource = "eggtim"
Set fSCR = CreateObject("MSScriptControl.ScriptControl")


CEGGTIM.ConnectionString = gConnect
CEGGTIM.RecordSource = "select *from DOKGILANTZISEGGTIM"
CEGGTIM.Refresh

CTIM.ConnectionString = gConnect
CTIM.RecordSource = "select *from DOKGILANTZISTIM inner JOIN PEL ON DOKGILANTZISTIM.EIDOS=PEL.EIDOS AND DOKGILANTZISTIM.KPE=PEL.KOD"
CTIM.Refresh

CMEM.ConnectionString = gConnect
CMEM.RecordSource = "select *from MEMREP "
CMEM.Refresh





fSCR.language = "vbscript"
fSCR.AddObject "cTIM", CTIM

fSCR.AddObject "cEGGTIM", CEGGTIM

'fSCR.AddObject "cMEM", CMEM
fSCR.AddObject "mactext", MACText

CEGGTIM.Recordset.MoveFirst

CMEM.Recordset.MoveFirst

CTIM.Recordset.MoveFirst

If Len(Dir("c:\PETREL.TXT")) > 1 Then
      DUM = load_forma(arxeio, 0)
 Else
     DUM = load_forma(arxeio, 1)
End If
 'Me.Caption = "load_forma=OK"
      
  For I = 1 To 250
    If Left$(gm_str(I), 10) = "**********" Then
       Exit For
    End If
  Next


    '* ΣΥΝΟΛΙΚΕΣ ΣΕΙΡΕΣ
         m_No_of_seir = Val(gm_str(I + 1))

    '* ΣΕΙΡΑ ΠΟΥ ΑΡΧΙΖΟΥΝ ΤΑ ΕΙΔΗ
         m_seir_eid = Val(gm_str(I + 2))

    '* ΣΕΙΡΑ ΠΟΥ ΑΡΧΙΖΟΥΝ ΤΑ ΣΥΝΟΛΑ
         m_seir_synol = Val(gm_str(I + 3))
    
    '* entolh gia ton εκτυπωτή
         m_entol_ektyp = gm_str(I + 4)

    '**************
    ' PRINTER.
    Dim cSelida
    
    
    If F928 = 1 Then
       cSelida = "Σελίδα "
    Else
       cSelida = to437("Σελίδα ")
    End If
    
    
    
    
    
    
    
    
    SELIDA = 1


Dim mfil As String

'για να σωζει σε διαφορετικα σημεια
'If f_YPOK = 0 Then
    mfil = F_File_Timologioy
'Else
 '   mfil = F_File2_Timologioy
'End If


10 Open mfil For Output As #1
    
    '---------------- επικεφαλίδες
20  ' On Error GoTo errorl
30      e = 0
40      Do While Len(Trim(gm_f(e + 1))) > 0
         e = e + 1
         'MM = Trim(gm_f(e))
'
         'm_m = ExecuteLine("gvar=" + MM)
         'eField(e) = gvar
50         ePict(e) = Trim(gpic(e))
60         eProw(e) = gm_r(e)
70         ePcol(e) = gm_c(e)
80         eFonts(e) = ""  '     && trim( Fonts )
        ' Printer.Print gm_f(e) + Format(eProw(e), "0000") + "=" + Format(ePcol(e), "0000") + "="
90      Loop
      '**************** ΣΕΙΡΕΣ ΜΕ ΕΙΔΗ ****************************************
100      m = 0
110    Do While Len(Trim(gm_f(e + m + 2))) > 0
120         m = m + 1

         'MM = Trim(gm_f(M + e + 1))

         'm_m = ExecuteLine("gvar=" + MM)

         'mField(M) = gvar
130         mPict(m) = Trim(gpic(m + e + 1))
140         mProw(m) = gm_r(m + e + 1)
150         mPcol(m) = gm_c(m + e + 1)
         'mFonts(M) = "" '     && trim( Fonts )
         'MM = mPict(M + e + 1)
160      Loop


Dim fcheckonly As Boolean


      '****************  ΣΥΝΟΛΑ  ****************************************
      S = 0
170      Do While Len(Trim(gm_f(e + m + S + 3))) > 0
         S = S + 1
'         MM = Trim(gm_f(M + e + s + 2))
'         MM = "gvar=" + MM
'         m_m = ExecuteLine("gvar=" + MM)
'
'         sField(s) = gvar
180         sPict(S) = Trim(gpic(m + e + S + 2))
190         sProw(S) = gm_r(m + e + S + 2)
200         sPcol(S) = gm_c(m + e + S + 2)
'         sFonts(s) = ""  '   && trim( Fonts )
'         MM = sPict(M + e + s + 2)
'         'Debug.Print Format(sField(s), MM)
210      Loop


 'Printer.FontSize = 10
' Printer.FontName = "Courier New"
' Printer.Font.Charset = 161
 





       If Len(Trim(m_entol_ektyp)) > 0 Then
           'Printer.Print  ' &_entol_ektyp
       End If



'===================================== ΕΠΙΚΕΦΑΛΙΔΕΣ ===================
       syn_row = 0
       ' ΕΚΤΥΠΩΣΗ ΕΠΙΚΕΦΑΛΙΔΩΝ
       
' On Error GoTo errorl
       m1 = 0
208       For k = 1 To e
             
             mM = ePict(k)
             'For i = 1 To eProw(K): Printer.Print: Next
             For I = 1 To eProw(k)
                 If k = 1 And m_No_of_seir = 1000 Then
                    ' ΜΗΝ ΒΑΖΕΙΣ ΚΕΝΗ ΣΕΙΡΑ ΣΤΗΝ ΑΡΧΗ
                 Else
                     Print #1, " "
                 End If
                 
             Next
             syn_row = syn_row + eProw(k) 'συγκεντρώνω τις σειρές
             
220           mM = mac(Trim(Trim(gm_f(k)))) ': m_m = ExecuteLine("gvar=" + MM)
           'Printer.Print Tab(ePcol(K)); Format(gvar, ePict(K))
           
              If eProw(k + 1) > eProw(k) Then 'αν πρόκειται να τυπώσει παρακάτω μην τυπώνεις δίπλα μου
                '   Print #1, Tab(ePcol(K)); Format(mM, ePict(K))
                Print #1, Tab(ePcol(k)); Format(Left(mM, Len(ePict(k))), ePict(k));
              Else
                  Print #1, Tab(ePcol(k)); Format(Left(mM, Len(ePict(k))), ePict(k));
              End If
             
                m1 = eProw(k)
             
230        Next
      
      
      
    If m_seir_eid >= m_No_of_seir Then
        GoTo 400
    End If
        
      
      
      
      
240   For k = 1 To m_seir_eid - syn_row
          Print #1, Space(80) 'Printer.Print:
      Next 'βρισκω την σειρά  των ειδών
       
       
   
       
       
       
       
       
250       syn_row = m_seir_eid
     '  Printer.Print
       'On Error GoTo 0
260       CEGGTIM.Recordset.MoveFirst
'=========================================== ΕΙΔΗ ====================
270       Do While Not CEGGTIM.Recordset.EOF
         

         DUM = "select *from eid where kod='" + CEGGTIM.Recordset("kode") + "'"
         If IsNull(CEGGTIM.Recordset("kode")) Then Exit Do
         
         'Set gEid = db.OpenRecordset(dum)
         For k = 1 To m
           
280           mM = mac(Trim(Trim(gm_f(e + k + 1))))                 ': m_m = ExecuteLine("gvar=" + MM)
           '   Printer.Print Format(gvar, gpic(e + K + 1));
290           Print #1, Tab(mPcol(k)); GFORMAT(mM, gpic(e + k + 1));
         Next
310         CEGGTIM.Recordset.MoveNext
            syn_row = syn_row + 1

            
            ' ΑΝ ΔΕΝ ΧΩΡΑΝΕ ΟΙ ΣΕΙΡΕΣ ΣΤΟ ΤΙΜΟΛΟΓΙΟ ( Η ΣΕΙΡΑ ΠΟΥ ΤΥΠΏΝΕΤΑΙ ΞΕΠΕΡΝΑ ΤΗΝ ΣΕΙΡΑ ΤΩΝ ΣΥΝΟΛΩΝ)
            If syn_row > m_seir_synol Then
                Print #1, ""
                Print #1, cSelida + Format(SELIDA, "##") ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                SELIDA = SELIDA + 1
                Print #1, Chr(12) ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                'If SELIDA > 1 Then
                 '   Print #1, "Σελίδα " + Format(SELIDA, "##"); ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                'End If
                GoSub TYPOSE_HEADER

            End If




       Loop

       
  npic = Space(120)
  
330    For k = 1 To m_seir_synol - syn_row: Print #1,: Next  'βρισκω την σειρά  των synolon
       
       For k = 1 To S
340         mM = Trim(gm_f(m + e + k + 2))
350        ' m_m = ExecuteLine("gvar=" + MM)
         
                 mM = mac(Trim(gm_f(m + e + k + 2)))
         
         
360         For I = 1 To sProw(k): Print #1,: Next
         'Printer.Print Format(gvar, gpic(M + e + K + 2));
370         Print #1, Tab(sPcol(k)); GFORMAT(mM, gpic(m + e + k + 2));
       Next
       
'       If SELIDA > 1 Then
'                Print #1, "Σελίδα " + Format(SELIDA, "##") ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
'       End If
       
       
       
       
       
       
400 Close #1

380 CEGGTIM.Recordset.Close

Set fSCR = Nothing


If f_YPOK = 0 Then
 F_Type_File_Timologioy = 1
 If F_Type_File_Timologioy = 1 Then
   If F928 = 1 Then
      If Len(Dir("c:\marietos.txt")) > 1 Then
         ' kanei "type c:\print"  meso dos
         DUM = Shell("c:\mercvb\notedos.exe /p " + mfil, vbMaximizedFocus)
       ElseIf Len(Dir("c:\AGIOS.txt")) > 1 Then
         ' kanei "type c:\print"  meso dos
         DUM = Shell("c:\mercvb\ff.exe", vbMaximizedFocus)
         
       ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then
         ' kanei "type c:\print"  meso dos
          '  DUM = EKTYP_ME_PRINT(mfil)
            ' DUM = Shell("c:\mercvb\ff.exe", vbMaximizedFocus)
       Else
            DUM = Shell("c:\mercvb\notepad.exe /p " + mfil, vbMaximizedFocus)
       End If
   Else  ' <>928
        If Len(Dir("c:\FAKKAV.TXT")) > 1 Then
         '  DUM = EKTYP_ME_PRINT(mfil)
        ElseIf Len(Dir("c:\PETREL.TXT")) > 1 Then
           DUM = Shell("C:\T.BAT ", vbMinimizedFocus)
        Else
           DUM = Shell("C:\MERCVB\PRINT.EXE " + Trim(mfil), vbMinimizedFocus)
      End If
   End If
 Else

      If Len(Dir("c:\DION2.EXE")) > 1 Then
        ' DUM = EKTYP_ME_PRINT(mfil)
         MILSEC 5000
         FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
      ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then
          'MILSEC 8000
          Do While Len(Dir("c:\FOROL\FOROL.txt")) > 1
            MILSEC 500
          Loop
         FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
         ' kanei "type c:\print"  meso dos
'            DUM = EKTYP_ME_PRINT(mfil)

      ElseIf Len(Dir("c:\KENO1.txt")) > 1 Then
          'MILSEC 8000
          MILSEC 1000
          Do While Len(Dir(mfil)) > 1
            MILSEC 1000
          Loop
          FileCopy "C:\KENO1.TXT", mfil
          FileCopy "C:\KENO1.TXT", "C:\PP2"
      End If
 End If

Else
'-----------------------------------------------
 If F_Type2_File_Timologioy = 1 Then
   If F928 = 1 Then
   Else  ' <>928
         DUM = Shell("C:\MERCVB\PRINT.EXE " + mfil, vbMinimizedFocus)
   End If
 Else
      If Len(Dir("c:\DION2.EXE")) > 1 Then
      ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then
      End If
 End If
End If







Exit Function



create_index:

'db.Execute "CREATE INDEX eidko ON eid(kod);"
Resume Next



errorl:
    MsgBox Str(Erl) + "MM=" + mM 'RECOVER USING objErrorInfo
      'sfalma = 1
      'MsgBox "Ά¦ª ©«¦ §›ε¦ " + Field + "  £.§›ε¦¬ " + Str(i)
'    Resume Next
 Close #1
 'Resume Next
 
 
 

 
 
 Exit Function
     

'
'User-Defined String Formats (Format Function)
'
'
'You can use any of the following characters to create a formatexpression for strings:
'
'Character Description
'@ Character placeholder. Display a character or a space. If the string has a character in the position where the at symbol (@) appears in the format string, display it; otherwise, display a space in that position. Placeholders are filled from right to left unless there is an exclamation point character (!) in the format string.
'& Character placeholder. Display a character or nothing. If the string has a character in the position where the ampersand (&) appears, display it; otherwise, display nothing. Placeholders are filled from right to left unless there is an exclamation point character (!) in the format string.
'< Force lowercase. Display all characters in lowercase format.
'> Force uppercase. Display all characters in uppercase format.
'! Force left to right fill of placeholders. The default is to fill placeholders from right to left.





TYPOSE_HEADER:
       syn_row = 0
       ' ΕΚΤΥΠΩΣΗ ΕΠΙΚΕΦΑΛΙΔΩΝ
       
       m1 = 0
       For k = 1 To e
             mM = ePict(k)
             For I = 1 To eProw(k)
                 If k = 1 And m_No_of_seir = 1000 Then
                    ' ΜΗΝ ΒΑΖΕΙΣ ΚΕΝΗ ΣΕΙΡΑ ΣΤΗΝ ΑΡΧΗ
                 Else
                     Print #1, " "
                 End If
             Next
             syn_row = syn_row + eProw(k) 'συγκεντρώνω τις σειρές
             mM = mac(Trim(Trim(gm_f(k)))) ': m_m = ExecuteLine("gvar=" + MM)
      If eProw(k + 1) > eProw(k) Then 'αν πρόκειται να τυπώσει παρακάτω μην τυπώνεις δίπλα μου
           Print #1, Tab(ePcol(k)); Format(Left(mM, Len(ePict(k))), ePict(k));
      Else
           Print #1, Tab(ePcol(k)); Format(Left(mM, Len(ePict(k))), ePict(k));
      End If
      m1 = eProw(k)
    Next
                If SELIDA > 1 Then
                    Print #1, cSelida + Format(SELIDA, "##") ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                End If
    
    For k = 1 To m_seir_eid - syn_row - 1
          Print #1, Space(80) 'Printer.Print:
    Next 'βρισκω την σειρά  των ειδών
    syn_row = m_seir_eid
Return

MHNYMA:
HandleError "Par1:ektypforma"
Resume Next
        
End Function

'εκτύπωση
Function print_timol(ByVal atim As String, ByVal hme As Date)
Dim DUM

    Me.Caption = "ektyp_forma"
    
'If f_FORM_EKTY = 99 Then
    DUM = ektyp_forma("c:\mercvb\f95.txt", atim, hme)
'Else
 '   DUM = ektyp_forma("c:\mercvb\f" + Format(f_FORM_EKTY, "00") + ".txt", atim, hme)
'End If

'
'GoTo 10

End Function

Function mac(ByVal S As String) As String
'δινω ενα στρινγκ και μου επιστρέφει μία τιμή
Dim X
'On Error GoTo 0
MACText.Text = "   "
On Error GoTo err2
X = "MACtEXT.Text = " + S
fSCR.ExecuteStatement X

If F928 = 1 Then
    mac = MACText.Text
Else
    mac = to437(MACText.Text)
End If

Exit Function

err2:
Me.Caption = X
mac = "  "
Resume Next
End Function



