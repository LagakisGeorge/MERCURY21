VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form pelat43 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   10380
   ClientLeft      =   8085
   ClientTop       =   450
   ClientWidth     =   12090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   12090
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton zoomXreopistoseis 
      Caption         =   "Ρυθμιση"
      Height          =   360
      Left            =   5040
      TabIndex        =   29
      Top             =   1200
      Width           =   990
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   6495
      Left            =   4800
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox EPOsympsif 
         Height          =   375
         Left            =   2400
         TabIndex        =   22
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "pelat43.frx":0000
         Height          =   4455
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   7858
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
      Begin VB.Label LabelEPO 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label LBLEPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label lblΕπωνυμίαΣυμψηφιζόμενου 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Επωνυμία Συμψηφιζόμενου "
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label lblΣυμψηφισμόςΜε 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Συμψηφισμός με Κωδικό"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.TextBox m38 
      Height          =   405
      Left            =   2880
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox m30 
      Height          =   405
      Left            =   2880
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000010&
      Caption         =   "Εξοδος"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"pelat43.frx":0014
   End
   Begin VB.ComboBox Aitiologies 
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Width           =   7575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "genarate ARITMISI"
      Height          =   735
      Left            =   1320
      TabIndex        =   11
      Top             =   8640
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Εκτύπωση Απόδειξης Είσπραξης"
      Enabled         =   0   'False
      Height          =   825
      Left            =   2880
      TabIndex        =   10
      Top             =   5160
      Width           =   1770
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6600
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      Destination     =   1
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4320
      Top             =   8880
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.CommandButton Command1 
      Caption         =   "Καταχώρηση"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   199426049
      CurrentDate     =   37743
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   5055
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   5400
      Top             =   9120
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin MSAdodcLib.Adodc PELDATA1 
      Height          =   330
      Left            =   4440
      Top             =   8880
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      Destination     =   1
   End
   Begin VB.Label LABEL4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".."
      Height          =   195
      Left            =   600
      TabIndex        =   26
      Top             =   7920
      Width           =   90
   End
   Begin VB.Label lLabel4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   840
      TabIndex        =   25
      Top             =   8160
      Width           =   45
   End
   Begin VB.Label Label7 
      Caption         =   "Λογ.Ταμείου/Τράπεζας"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Πελάτης/Προμηθ"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Παράμετροι εκτύπωσης"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ποσό"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Αιτιολογία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Αριθμός Παραστατικού"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "pelat43"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim F_EKTYPOTHS As String '= FINDPARAMETROI(1, "PAR1", "F_EKTYPOTHS", "..", "Εκτυπωτής παραστατικών(με .. προεπιλεγμένος)") '
Dim f_er_symcif As String

Dim mLock As Integer

Dim f_tameio    '= FindParametroi(1,"PELAT43", "F_TAMEIO", "9990", "ΛΟΓΑΡΙΑΣΜΟΣ ΤΑΜΕΙΟ")

Dim F_CanDelete            As Integer

Dim f_tam                  As Boolean

Dim F_PEL_ANTITYPA         As Integer

Dim f_PROM_ANTITYPA        As Integer

Dim F_AA                   As Integer  ' AA ΕΓΓΡΑΦΗς eggtim

Dim F_FORMA1 As String
Dim F_FORMA2 As String

Dim f_LOGISTIKH            As Integer

Dim f_symcif               As Integer ' αν είναι Συμψηφισμός =1

Dim F_EISP_SQL As String

Dim F_ID_EGG As Long ' TO ID THS EGGRAFHS EGG ΠΟΥ ΔΗΜΙΟΥΡΓΗΘΗΚΕ (ΜΑΝΑ)

Dim f_KENTRA_ADYNATISMATOS As Integer

Private Sub Combo1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Combo1_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Combo1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Combo1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Combo1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo1_LostFocus()

        '<EhHeader>
        On Error GoTo Combo1_LostFocus_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

100     R.Open "select * from XREOPIS where TITLOS='" + Trim(Left(Combo1.Text, 50)) + "' AND PEL ='" + gBuff + "'", Gdb, adOpenForwardOnly, adLockReadOnly

        On Error Resume Next

110     Text1.Text = R("ARITMISI")
120     Aitiologies.Text = Left(Combo1.Text, 50)
130     F_AA = R("EIDOS")

        F_FORMA1 = CNull(R!FORMA1)
        F_FORMA2 = CNull(R!FORMA2)
        '---------------------------------------------------------------------------------
        If IsNull(R!KODLOG) Then
            m38.Text = "38-00-00-0000"
        Else
            m38.Text = R!KODLOG '"38-00-00-0000"
        End If
      
        If gBuff = "e" Then
          'If R!epit = 4 Then ' ΠΙΣΤΩΤΙΚΕ
           ' m30.Text = "30-02-00-0000"
          'Else
            m30.Text = "30-00-00-0000"
          'End If
        Else
            m30.Text = "50-00-00-0000"
        End If

        f_symcif = 0
        f_er_symcif = "e"
        ' αν είναι Συμψηφισμός τότε epit=3 taytpel="r" ή "e" kai taytxpi="X/Π" ανάλογα αν ο αντισυμβαλλόμενος θα χρεωθεί ή θα πιστωθεί
        If Not IsNull(R!epit) Then
            f_symcif = 0
            
            If R!epit = 2 Then
                Frame1.Visible = True
                If IsNull(R!TaytPel) Then
                    MsgBox " δεν συμπληρώσατε Είδος(r/e) -> Παράμετροι Κινήσεων Πελατών"
                    f_symcif = 0

                    Exit Sub
          
                End If
          
            End If

            If R!epit = 2 Then
                Frame1.Visible = True
                If IsNull(R!TaytXpi) Then
                    MsgBox " δεν συμπληρώσατε Χρέωση/Πίστωση -> Παράμετροι Κινήσεων Πελατών"
                    f_symcif = 0

                    Exit Sub
          
                End If
          
            End If
             
              
            If R!epit = 2 Then
              f_er_symcif = R!TaytPel
              f_symcif = 1
              Text4.Visible = True
              Label4.Caption = "κωδικός " + IIf(R!TaytPel = "r", "Προμηθ.", "Πελάτη")
              Label4.Visible = True
            End If
            
            
             If R!epit = 5 Then
                 f_symcif = 5
             End If
             
            
        End If

        '<EhFooter>
        Exit Sub

Combo1_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Combo1_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Combo1_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command1_Click()

        '----------------------------------------------------------------------
        'Δημιουργείται μία εγγραφή στο egg και ενημερώνεται το pel
        '
        '
        '
        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim DB As Database

        Dim R As New ADODB.Recordset, x As String, D, tam As Boolean

        Dim xpi, PEL, mXre

        Dim mXreosi, mPistosi
 Dim M_TAYT_KOD


         If Val(Text1.Text) = 0 Then
            If Len(Trim(Text1.Text)) <= 1 Then
                MsgBox "ΔΕΝ ΔΩΣΑΤΕ ΑΡΙΘΜΟ"
                Exit Sub
            
            End If
         End If
         
         If DTPicker1.Value > Now Then
                MsgBox "ΛΑΘΟΣ ΗΜΕΡΟΜΗΝΙΑ"
                Exit Sub
         
         End If
         
         
' ΕΑΝ ΕΧΩ ΣΥΜΨΗΦΙΣΜΟ ΝΑ ΜΕ ΖΗΤΑΕΙ ΝΤΕ ΚΑΙ ΚΑΛΑ ΑΝΤΙΚΡΥΖΟΜΕΝΟ ΛΟΓΑΡΙΑΣΜΟ
'If f_symcif = 1 Then
'
'   If Len(Text4.Text) = 0 Then
'        MsgBox "ΔΕΝ ΔΩΣΑΤΕ ΚΩΔΙΚΟ ΣΥΜΨΗΦΙΖΟΜΕΝΟΥ ΛΟΓΑΡΙΑΣΜΟΥ"
'                EPOsympsif.SetFocus
'
'                Exit Sub
'
'   End If
'
'
'End If


If Val(Text3.Text) = 0 Then
   MsgBox "ΔΕΝ ΒΑΛΑΤΕ ΠΟΣΟ"
   Exit Sub
End If









        'βρίσκω τις παραμέτρους της κίνησης
        '----------------------------------

100     If gXEIRISTHS = 8 Then
110         'Adodc1.RecordSource = "select * from XREOPIS where PEL ='" + gBuff + "'"
        Else
120         'Adodc1.RecordSource = "select * from XREOPIS where RIGHT(TITLOS,1)<>'.' AND PEL ='" + gBuff + "'"
        End If
Adodc1.RecordSource = "select * from XREOPIS where PEL ='" + gBuff + "'"
130     Adodc1.Refresh

140     Set D = Adodc1.Recordset
150     Adodc1.Recordset.MoveFirst

160     Do While Not D.EOF

170         If D("titlos") = Trim(Left(Combo1.Text, 50)) Then
180             xpi = D("xreopis")

190             If IsNull(D("tameiakh")) Then
200                 tam = False
                Else
210                 tam = D("tameiakh")
                End If

220             f_tam = tam
                M_TAYT_KOD = D!KODLOG
230             PEL = IIf(IsNull(D("taytpel")), "", D("taytpel"))
240             x = IIf(IsNull(D("taytxpi")), "", D("taytxpi"))
            End If

250         D.MoveNext
        Loop

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If
260     R.Open "SELECT TOP 10 *FROM EGG", Gdb, adOpenDynamic, adLockOptimistic

        Dim TaytPel, TaytXpi

270     TaytPel = Trim(PEL)
280     TaytXpi = x

        Dim nSize

290     nSize = R("ait").DefinedSize

        'Set pel = db.OpenRecordset("pel")
        'pel.Index = "ko"
        'pel.Seek "=", gbuff, PELDATA1.Recordset("kod")

        'Set PEL = PELDATA1.Recordset

        'ενημέρωση του πελάτη
        '--------------------
        'PEL.Edit
300     mXreosi = 0: mPistosi = 0
310     Text3.Text = DOT(Text3.Text)

320     If PELDATA1.Recordset("eidos") = "e" Then
        
330         If xpi = "Χ" Then
340             mXre = Val(Text3.Text)
350             mXreosi = Val(Text3.Text)
            End If

360         If xpi = "Π" Then
370             mXre = -Val(Text3.Text)
380             mPistosi = Val(Text3.Text)
            End If

'390         If gXEIRISTHS = 8 Then
'400             If Right(Combo1.Text, 1) = "." Then
'
'410                 If xpi = "Χ" Then
'420                     mXre = Val(Text3.Text)
'430                     mXreosi = Val(Text3.Text)
'440                     xpi = "Χ"
'                    Else
'450                     mXre = -Val(Text3.Text)
'460                     mPistosi = Val(Text3.Text)
'470                     xpi = "Π"
'                    End If
'
'                End If
'
'            End If
        End If

480     If PELDATA1.Recordset("eidos") = "r" Then
490     '    If gXEIRISTHS = 8 Then
500      '       If Right(Combo1.Text, 1) = "." Then
                    ' If left(Combo1.Text, 2) = "ΧΡ" Then
                    '    xpi = "Χ"
                    ' Else
                    '    xpi = "Π"
                    ' End If
          '      End If
          '  End If    '

510         If xpi = "Π" Then
520             mXre = Val(Text3.Text)
530             mPistosi = Val(Text3.Text)
            End If

540         If xpi = "Χ" Then
550             mXre = -Val(Text3.Text)
560             mXreosi = Val(Text3.Text)
            End If
        End If

        ' PELDATA1.Recordset("typ") = PELDATA1.Recordset("typ") + mXre
570     PELDATA1.Recordset.Update

        'προσθέτει εγγραφή στο EGG
        '-------------------------
        'R.AddNew
        '   R("hme") = DTPicker1.Value
        '   R("atim") = Text1.Text
        '   R("apa") = Val(Text1.Text)
        '   R("ait") = aitiologies.Text
        '   R("xre") = mXre
        '   R("eidos") = PEL("eidos")
        '   R("kod") = PEL("kod")
        '   R("xpi") = xpi
        'R.Update
        Dim ssa As String

        Dim meidos

580     meidos = PELDATA1.Recordset("eidos")

590    ' If gXEIRISTHS = 8 Then
600     '    If Right(Combo1.Text, 2) = ".." Then
610      '       If meidos = "e" Then meidos = "E"
620       '      If meidos = "r" Then meidos = "R"
           ' End If

       'End If

        Dim M_ATIM As String
        
        
        Dim mait As String
         mait = Trim(Left(Aitiologies.Text, 50))
        If f_symcif = 1 Then
           mait = Trim(Left(Aitiologies.Text, 50)) + " (με κωδ." + Text4.Text
        End If
        
630     ssa = "insert into EGG (EIDXPI,HME,ATIM,APA,AIT,EIDOS,KOD,XPI,XREOSI,PISTOSI,AAXREOPIS,USERID"

        If f_LOGISTIKH = 1 Then
         
            ssa = ssa + ",M30,M38 "
         
        End If

        ssa = ssa + ") values (" + mID$(Combo1.Text, 51, 10) + ","

        '640     If Len(Dir("C:\KAGI.TXT")) > 0 Then
        '650         ssa = ssa + "'" + Format(Now, "mm/dd/yyyy") + "',"
        '660         M_ATIM = Label5.Caption
        'ssa = ssa + "'" + Label5.Caption + "'," ' ATIM
        '        Else
670     ssa = ssa + "'" + Format(DTPicker1, "mm/dd/yyyy") + "',"
680     M_ATIM = Text1.Text

        '        End If
690     ssa = ssa + "'" + M_ATIM + "',"    ' ATIM

700     ssa = ssa + str(Val(Text1.Text)) + ","    'APA
710     ssa = ssa + "'" + Left(mait, nSize) + "',"    'AIT
720     ssa = ssa + "'" + meidos + "',"    ' EIDOS='E'
730     ssa = ssa + "'" + PELDATA1.Recordset("kod") + "',"    ' KOD
740     ssa = ssa + "'" + xpi + "',"    ' XPI
750     ssa = ssa + Replace(str(mXreosi), ",", ".") + ","  'XREOSI
760     ssa = ssa + Replace(str(mPistosi), ",", ".") + "," + str(F_AA) + "," + str(gUserId)  ' PISTOSI
        
        If f_LOGISTIKH = 1 Then
            ssa = ssa + ",'" + m30.Text + "','" + m38.Text + "' "
        End If
 
        ssa = ssa + ")"
        
        Dim k

770     Gdb.Execute ssa, k

        '============================================  ΣΥΜΨΗΦΙΣΜΟΣ ======================================


 'ΠΙΑΝΩ ΤΟ ΙΔ ΤΗΣ MANAS EGGRAFHS ( ΠΟΥ ΔΗΜΙΟΥΡΓΗΘΗΚΕ ΠΑΡΑΠΑΝΩ)
            Dim IDN As Long, RR7 As New ADODB.Recordset
            RR7.Open "SELECT MAX(ID) FROM EGG", Gdb, adOpenDynamic, adLockOptimistic
            IDN = RR7(0)
            F_ID_EGG = IDN
            RR7.Close


780     If (f_symcif = 1 Or f_symcif = 5) And k > 0 Then





'            'ΠΙΑΝΩ ΤΟ ΙΔ ΤΗΣ MANAS EGGRAFHS ( ΠΟΥ ΔΗΜΙΟΥΡΓΗΘΗΚΕ ΠΑΡΑΠΑΝΩ)
'            Dim IDN As Long, RR7 As New ADODB.Recordset
'            RR7.Open "SELECT MAX(ID) FROM EGG", Gdb, adOpenDynamic, adLockOptimistic
'            IDN = RR7(0)
'            RR7.Close
            
            






           

            Dim MM_XREOSI

            Dim MM_PISTOSI
            
            'ΘΑ ΧΡΕΩΘΕΙ/ΠΙΣΤΩΘΕΙ Ο αντισυμβαλλόμενος
            
            If Len(TaytPel) = 1 Then  'πληκτρολογώ τον κωδικό στο κουτάκι text4
               If Len(Trim(CNull(M_TAYT_KOD))) = 0 Or Len(Trim(CNull(M_TAYT_KOD))) >= 10 Then ' ΑΝ ΕΧΩ ΚΩΔΙΚΟ ΛΟΓΙΣΤΙΚΗς
                  M_TAYT_KOD = Text4.Text
               End If
            Else

                If mID(TaytPel, 2, 4) = "****" Then  ' θα χρεώθεί ο ίδιος λογαριασμός αν έχω δηλώσει e****  ή r****
                    M_TAYT_KOD = PELDATA1.Recordset("kod")
                Else ' 'ΘΑ ΧΡΕΩΘΕΙ/ΠΙΣΤΩΘΕΙ Ο  ΛΟΓΑΡΙΑΣΜΟΣ που ειναι δηλωμένος π.χ. e1020 θα χρεωθεί/πιστωθεί(αναλογα το taytxpi) ο 1020 πελάτης
                    M_TAYT_KOD = Trim(mID(TaytPel, 2, 10))
                End If
            End If

820         If Trim(TaytXpi) = "Π" Then
830             MM_PISTOSI = mPistosi + mXreosi
840             MM_XREOSI = 0
            Else
850             MM_XREOSI = mPistosi + mXreosi
860             MM_PISTOSI = 0
            End If

870         ssa = "insert into EGG (IDTIM,HME,ATIM,APA,AIT,EIDOS,KOD,XPI,XREOSI,PISTOSI,AAXREOPIS,USERID,IDEGGSYND "

            If f_LOGISTIKH = 1 Then
                ssa = ssa + ",M30,M38 "
            End If
            If f_symcif = 5 Then
               mait = Trim(PELDATA1.Recordset("kod")) + "-" + PELDATA1.Recordset("EPO") + ")"
            Else
               mait = "Συμψ." + PELDATA1.Recordset("kod") + "-" + PELDATA1.Recordset("EPO") + ")"
            End If
            
            ssa = ssa + ") values (1," 'ΒΑΖΩ IDTIM=1 ΓΙΑ ΝΑ ΜΗΝ ΤΟ ΒΛΕΠΕΙ ΣΤΗΝ ΓΕΦΥΡΑ EPSILON (NA BLEPEI MONO THN KYRIA EGGRAFH TOY SYMCHFISMOY)

880         ssa = ssa + "'" + Format(DTPicker1, "mm/dd/yyyy") + "',"
890         ssa = ssa + "'" + M_ATIM + "',"    ' ATIM
900         ssa = ssa + Text1.Text + ","    'APA
910         ssa = ssa + "'" + Left(mait, 25) + "',"    'AIT
920         ssa = ssa + "'" + Left(TaytPel, 1) + "',"    ' EIDOS='E'
930         ssa = ssa + "'" + M_TAYT_KOD + "',"    ' KOD
940         ssa = ssa + "'" + TaytXpi + "',"    ' XPI
950         ssa = ssa + str(MM_XREOSI) + ","    'XREOSI

            '  960         ssa = ssa + str(MM_PISTOSI) + "," + str(gUserId)     ' PISTOSI
            ssa = ssa + str(MM_PISTOSI) + "," + str(F_AA) + "," + str(gUserId)  ' PISTOSI
        
            If f_LOGISTIKH = 1 Then
                ssa = ssa + ",'00-00-00-0000','00-00-00-0000' "
            End If
 
            ssa = ssa + "," + str(IDN) + ")"

970         Gdb.Execute ssa, k
            Gdb.Execute "UPDATE EGG SET IDEGGSYND=" + str(IDN + 1) + " WHERE ID=" + str(IDN)
        
        End If

        'ΕΝΗΜΕΡΩΣΗ ΤΑΜΕΙΟΥ
980     If tam Then
990         If mPistosi > 0 Then
1000            mXreosi = mPistosi
1010            mPistosi = 0
            Else
1020            mPistosi = mXreosi
1030            mXreosi = 0

            End If

1040        GoSub UPDATE_TAMEIO

        End If

        ''ΕΝΗΜΕΡΩΣΗ ΤΑΜΕΙΟΥ
        'If UCase(meidos) = "E" And tam Then
        '   If MPISTOSI > 0 Then ''εισπραξη από πελάτη
        '       MXREOSI = MPISTOSI
        '      MPISTOSI = 0
        '      GoSub UPDATE_TAMEIO
        '   End If
        'Else
        '   If MXREOSI > 0 And tam Then 'πληρωμη σε προμηθευτή
        '      MPISTOSI = MXREOSI
        '      MXREOSI = 0
        '      GoSub UPDATE_TAMEIO
        '
        '   End If
        'End If

1050    R.Close

        Dim SQLT As String

        Dim TYP

        'ENHMERVSH ME PROHGOYMENO YPOLOIPO
        '----------------------------------
1060    SQLT = "SELECT KOD,"
1070    SQLT = SQLT + "SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END )  -"
1080    SQLT = SQLT + " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END )   AS YP"
1090    SQLT = SQLT + " FROM EGG  where HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND EIDOS='" + meidos + "' AND KOD='" + PELDATA1.Recordset("kod") + "' GROUP BY KOD"
1100    R.Open SQLT, Gdb, adOpenDynamic, adLockOptimistic

1110    If IsNull(R("YP")) Then
1120        TYP = 0
        Else
1130        TYP = R("YP")
        End If

1140    If IsNull(PELDATA1.Recordset("AYP")) Then
1150        TYP = TYP
        Else
1160        TYP = TYP + PELDATA1.Recordset("AYP")
        End If

1170    Gdb.Execute "UPDATE PEL SET ENERGOS=1, TYP=" + DOT(str(TYP)) + " WHERE  EIDOS='" + meidos + "' AND KOD='" + PELDATA1.Recordset("kod") + "'"

        'If f_KENTRA_ADYNATISMATOS = 1 And mXreosi = 0 And mPistosi = 0 Then
        '   Gdb.Execute "UPDATE PEL SET NUM5=(CASE WHEN NUM5 IS NULL THEN 0 ELSE NUM5 END) + 1  WHERE  EIDOS='" + meidos + "' AND KOD='" + PELDATA1.Recordset("kod") + "'"
        'End If

        'Dim R As New ADODB.Recordset
1180    Gdb.Execute "UPDATE XREOPIS SET ARITMISI=ARITMISI+1 where EIDOS='" + Trim(mID$(Combo1.Text, 51, 10)) + "'"
        ' Gdb.Execute "UPDATE ARITMISI SET ARITMISI=" + Text1.Text + "  WHERE  ID=55"

        Dim RR As Recordset

1190    Set DB = OpenDatabase("C:\MERCVB\REPORTS\REPORTS.MDB")
1200    Set RR = DB.OpenRecordset("mem")



         DB.Execute "update MEM SET N1=" + str(gVal(Text3.Text))
         DB.Execute "update MEM SET PAR10=" + str(gVal(Text1.Text))
         DB.Execute "update MEM SET HME=#" + Format(DTPicker1.Value, "dd/MM/yyyy") + "#"
         
DB.Execute "update MEM SET par2='" + PELDATA1.Recordset("kod") + "'"

DB.Execute "update MEM SET C1='" + PELDATA1.Recordset("epO") + "'"
DB.Execute "update MEM SET C2='" + PELDATA1.Recordset("epA") + "'"
DB.Execute "update MEM SET C3='" + Left(PELDATA1.Recordset("die"), 25) + "'"
DB.Execute "update MEM SET C4='" + PELDATA1.Recordset("afm") + "'"

DB.Execute "update MEM SET C5='" + PELDATA1.Recordset("afm") + "'"
DB.Execute "update MEM SET PAR9='" + Format(DTPicker1.Value, "dd/MM/yyyy") + "'"


DB.Execute "update MEM SET par12='" + Combo1.Text + "'"











1210   ' RR.Edit
1220   ' RR!n1 = Val(Text3.Text)    ' poso
    
        'για να παρει το νεο υπολοιπο
1230    PELDATA1.Refresh
    
1240    'RR!par10 = Text1.Text    ' ar.parastatikoy
1250   ' RR!hme = DTPicker1.Value
1260   ' RR!PAR2 = PELDATA1.Recordset("kod")
'1270    RR!C1 = PELDATA1.Recordset("epo")
'1280    RR!c2 = PELDATA1.Recordset("epa")
'1290    RR!c3 = Left(PELDATA1.Recordset("die"), 25)
'1300    RR!c4 = PELDATA1.Recordset("afm")

'1310    RR!par12 = Combo1.Text

        Dim sql

1320    sql = "select sum(XREOSI) AS XX0,sum(PISTOSI) AS PP0 from EGG WHERE  HME>='" + "01/01/" + LTrim(str(Year(Now))) + "' AND HME<='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND EIDOS ='" + PELDATA1.Recordset("eidos") + "' and KOD='" + PELDATA1.Recordset("KOD") + "'"
1330    R.Close
1340    R.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly



           DB.Execute "update MEM SET par3='" + Format(PELDATA1.Recordset("typ"), "###,000.00") + "'"

1350   ' RR!par3 = Format(PELDATA1.Recordset("typ"), "###,000.00")
1360    R.Close



DB.Execute "update MEM SET par1='" + PELDATA1.Recordset("DOY") + "'"
DB.Execute "update MEM SET PAR11='" + Left(Aitiologies.Text, RR("PAR11").Size) + "'"

DB.Execute "update MEM SET MEMO ='" + Olografos(Val(DOT(Text3.Text))) + "'"




1370 '   RR!PAR1 = PELDATA1.Recordset("doy")
1380  '  RR!PAR11 = Left(Aitiologies.Text, RR("PAR11").Size)

1390  '  RR!memo = Olografos(Val(DOT(Text3.Text)))





1400   ' RR.Update

        'If Len(Dir("C:\AGIOS.TXT")) > 2 Then
        ' ' RR.Refresh
        '
        '  Dim K
        '  For K = 1 To 100: DoEvents: Next
        '
        '  If PELDATA1.Recordset("eidos") = "e" Then
        '     CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\EISPR.RPT"
        '  Else
        '     CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\PLIR.RPT"
        '  End If
        '  CrystalReport1.Destination = crptToPrinter
        '  CrystalReport1.Action = 1
        '  CrystalReport1.Action = 1
        'End If

1410    Command1.Enabled = False

1420    Command2.Enabled = True

        Exit Sub

UPDATE_TAMEIO:

1430    ssa = "insert into EGG (HME,ATIM,APA,AIT,EIDOS,KOD,XPI,XREOSI,PISTOSI,USERID) values ("
1440    ssa = ssa + "'" + Format(Now, "mm/dd/yyyy") + "',"
1450    ssa = ssa + "'" + M_ATIM + "',"    ' ATIM
1460    ssa = ssa + Text1.Text + ","    'APA
1470    ssa = ssa + "'" + Left(Me.Caption, 25) + "',"    'AIT
1480    ssa = ssa + "'e',"    ' EIDOS='E'
1490    ssa = ssa + "'" + f_tameio + "',"    ' KOD
1500    ssa = ssa + "'" + xpi + "',"    ' XPI
1510    ssa = ssa + str(mXreosi) + ","    'XREOSI
1520    ssa = ssa + str(mPistosi) + "," + str(gUserId) + ")"    ' PISTOSI

        On Error Resume Next

1530    Gdb.Execute ssa

1540    Return

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()


Dim c, D As String

Dim CPOSO As String
CPOSO = Format(gVal(Text3.Text), "###0.00")



c = "<%PS1;3;" + LTrim(Trim(Replace(CPOSO, ",", "."))) + ">"

D = "<%SL;" + LTrim(Trim(PELDATA1.Recordset("afm"))) + ";;;;;;;355;;" + Trim(Text1.Text) + ";0.00;0.00;0.00;0.00;0.00;0.00;0.00;0.00;0.00;0.00;EUR;1;" + LTrim(Trim(Replace(CPOSO, ",", "."))) + ";;;;>"

Dim c2, d2 As String
' den paizei me eteroxronismenh
c2 = "" '"<%PS1;2;" + Replace(Text3.Text, ",", ".") + ">"




    Dim cnn           As ADODB.Connection

                Dim RDB           As ADODB.Recordset

                Dim strConnection As String, strProvider As String, strSource As String

                strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
                strSource = "Data Source=C:\MERCVB\REPORTS\REPORTS.mdb;" '& App.Path &

                Set cnn = New ADODB.Connection
                strConnection = strProvider & strSource & "Persist Security Info=False"
                cnn.Open strConnection
                
               On Error Resume Next
               cnn.Execute "update MEM SET PAR18='" + D + "'"
               cnn.Execute "update MEM SET PAR19='" + c + "'"
            '   cnn.Execute "update MEM SET PAR20='" + c2 + "'"
         
                
'                     'On Error GoTo 0
'                   Dim crxApp As CRAXdDRT.Application
'                   Dim crxRpt As CRAXdDRT.Report
'
'
'
'
'                   Set crxApp = New CRAXdDRT.Application
'                   Set crxRpt = crxApp.OpenReport("c:\MERCVB\reports\PROEISPR.rpt")
'                    'Set CRXReport = CRXApplication.OpenReport("c:\MERCVB\reports\timol1.rpt", 1)
'
'
'              ' ExportReportToPDF crxRpt, "c:\mercvb\reports\timol1.pdf", "titlos"
'
'
'                   crxRpt.SelectPrinter "", printerName, ""
'                  ' crxRpt. = "Bullzip PDF Printer"  'Printer.DeviceName ' "Bullzip PDF Printer"
'                  crxRpt.PrintOut False, 1  ' NoCopies
                    
                   
  '              Ucr9print.printingCR9 F_EKTYPOTHS, "c:\MERCVB\reports\PROEISPR.rpt" ' CrystalReport1.ReportFileName
                
                


'<%PS3;3;15.00>
'
'Μαζί λοιπόν με το παραπάνω script, το e.line που θα το συνοδεύει με τα δεδομένα προς τον ΦΗΜ για να γίνει η προείσπραξη είναι το παρακάτω:
'
'<%SL;144476123;;;;;;;355;A;201;0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00;EUR;1;15.00;;;;>







        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     If PELDATA1.Recordset("eidos") = "e" Then
110         If F_AA = 1 Then
120             CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\plir.RPT"
            Else

               If Len(Trim$(F_EISP_SQL)) > 1 Then
                'EISPSQL.rpt  F_EISP_SQL
                 CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\" + F_EISP_SQL
                 
                 ' CrystalReport1.Q
                  CrystalReport1.SelectionFormula = "  {EGG.ID}=" + str(F_ID_EGG)
                  CrystalReport1.Connect = gConnect
                  
                  CrystalReport1.Destination = crptToWindow
             
                 Ucr9print.printingCR9 F_EKTYPOTHS, (CrystalReport1.ReportFileName)
                 'CrystalReport1.Action = 1
                ' CrystalReport1.R
               Else
             

130               If f_tam Then
                   
                    If Len(F_FORMA1) < 2 Then
140                      CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\EISPR.RPT"
                    Else
                        CrystalReport1.ReportFileName = F_FORMA1
                    End If
                    
                     If Len(F_FORMA2) < 2 Then
                        CrystalReport2.ReportFileName = "" 'C:\MERCVB\REPORTS\EISPR.RPT"
                    Else
                        CrystalReport2.ReportFileName = F_FORMA2
                    End If
                    
                    
                    
                    
                    
                    
                  Else
                  
                      If Len(F_FORMA1) < 2 Then
150                      CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\eispr.RPT"
                      Else
                        CrystalReport1.ReportFileName = F_FORMA1
                      End If
                      
                      If Len(F_FORMA2) < 2 Then
152                      CrystalReport2.ReportFileName = "" ' C:\MERCVB\REPORTS\eispr.RPT"
                      Else
                        CrystalReport2.ReportFileName = F_FORMA2
                      End If
                      
                      
                      
                      
                      
                      
                  End If
                
               End If
              
            End If
       
        Else

160         If f_tam Then

                    If Len(F_FORMA1) < 2 Then
162                      CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\plir.RPT"
                    Else
                        CrystalReport1.ReportFileName = F_FORMA1
                    End If

                   If Len(F_FORMA2) < 2 Then
164                      CrystalReport2.ReportFileName = "" 'C:\MERCVB\REPORTS\plir.RPT"
                    Else
                        CrystalReport2.ReportFileName = F_FORMA2
                    End If



170            ' CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\plir.RPT"
            Else
            
                    If Len(F_FORMA1) < 2 Then
172                      CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\sympsif2.RPT"
                    Else
                        CrystalReport1.ReportFileName = F_FORMA1
                    End If
            
                    If Len(F_FORMA2) < 2 Then
174                      CrystalReport2.ReportFileName = "" 'C:\MERCVB\REPORTS\sympsif2.RPT"
                    Else
                        CrystalReport2.ReportFileName = F_FORMA2
                    End If
            
            
            
180             ' CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\sympsif2.RPT"
            End If

        End If

190     CrystalReport1.Destination = crptToPrinter

        Dim N As Integer

        If PELDATA1.Recordset("eidos") = "e" Then

          '  For N = 1 To F_PEL_ANTITYPA
               ' CrystalReport1.Action = 1
                Ucr9print.printingCR9 F_EKTYPOTHS, (CrystalReport1.ReportFileName)
           ' Next
           
           If Len(F_FORMA2) > 2 Then
                  Ucr9print.printingCR9 F_EKTYPOTHS, (CrystalReport2.ReportFileName)
           End If
           
           
        Else

           ' For N = 1 To f_PROM_ANTITYPA
               ' CrystalReport1.Action = 1
                Ucr9print.printingCR9 F_EKTYPOTHS, (CrystalReport1.ReportFileName)
           ' Next
           
            If Len(F_FORMA2) > 2 Then
                  Ucr9print.printingCR9 F_EKTYPOTHS, (CrystalReport2.ReportFileName)
           End If
           
           
           
           

        End If
        
        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>
        Dim DUM

        Dim ANS

100     If Command1.Enabled = False Then
110         ANS = vbYes
        Else
120         ANS = MsgBox("Είσαι σίγουρος;", vbYesNo)
        End If

130     If ANS = vbYes Then
            '
        Else

            Exit Sub

        End If

        'DUM = pelat4.KARTELLA(gDir)
140     Unload Me

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset, x As String, D, tam As Boolean

100     R.Open "SELECT * FROM ARITMISI", Gdb, adOpenDynamic, adLockOptimistic
110     R.MoveFirst

        Dim k As Integer

120     k = 0

130     Do While Not R.EOF
140         k = k + 1
150         R(0) = k
160         R(1) = k
170         R.Update
180         R.MoveNext
        Loop

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err
        

        '</EhHeader>

100     Label5.Caption = DataGrid1.columns(2).Text + Left(DataGrid1.columns(0).Text, 7)

        '<EhFooter>
        Exit Sub

DataGrid1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.DataGrid1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.DataGrid1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
         Text4.Text = Data1.Recordset("KOD")
         LabelEPO.Caption = Data1.Recordset("EPO")
         
         
End Sub

Private Sub DTPicker1_GotFocus()

        '<EhHeader>
        On Error GoTo DTPicker1_GotFocus_Err

        '</EhHeader>

100     DTPicker1.ToolTipText = "."

        '<EhFooter>
        Exit Sub

DTPicker1_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.DTPicker1_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.DTPicker1_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo DTPicker1_KeyUp_Err

        '</EhHeader>
100     If KeyCode = vbKeyReturn Then
110         KeyCode = 0    'suppress the beep

120         If DTPicker1.ToolTipText = "." Then
130             DTPicker1.ToolTipText = ""
            Else
140             keybd_event VK_TAB, 0, 0, 0    'send a tab
            End If
        End If

        '<EhFooter>
        Exit Sub

DTPicker1_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.DTPicker1_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.DTPicker1_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub EPOsympsif_Change()
       '</EhHeader>

100
110       Data1.RecordSource = "select EPO,DIE,KOD from PEL where  EIDOS='" + f_er_symcif + "'  and EPO like '" + (SameLetters(EPOsympsif.Text)) + "%';"
120     Data1.Refresh

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "PELAT43"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyUp_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command3_Click

        '<EhFooter>
        Exit Sub

Form_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Form_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Form_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()
F_EKTYPOTHS = FINDPARAMETROI(1, "PAR1", "F_EKTYPOTHS", "..", "Εκτυπωτής παραστατικών(με .. προεπιλεγμένος)") '
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        Dim R    'As Recordset

100     f_tam = False

110     Me.KeyPreview = True

 Data1.ConnectionString = gConnect

'EISPSQL.rpt  F_EISP_SQL
F_EISP_SQL = FINDPARAMETROI(1, "PELAT43", "F_EISP_SQL", "", "Εναλλακτική φόρμα sql π.χ. EISPSQL.rpt")

        f_KENTRA_ADYNATISMATOS = Val(FINDPARAMETROI(1, "PELAT2", "f_KENTRA_ADYNATISMATOS", "0", "Αν είναι κέντρο αδυνατίσματος=1"))

        Dim r2 As New ADODB.Recordset

120     F_CanDelete = Val(FINDPARAMETROI(1, "BOHU11", "F_CanDelete", "1", "μπορω να σβήνω κινήσεις"))

        'If Len(Dir("C:\KAGI.TXT")) > 0 And PELDATA1.Recordset("EIDOS") = "r" Then
        '   Text4.Visible = True
        '   DataGrid1.Visible = True
        '   Label4.Visible = True
        '   Label5.Visible = True
        'End If

        f_LOGISTIKH = Val(FINDPARAMETROI(1, "PELAT43", "f_LOGISTIKH", "0", "ΓΕΦΥΡΑ ΣΕ ΛΟΓΙΣΤΙΚΗ=1"))

        If f_LOGISTIKH = 1 Then

            Dim DUM

            m30.Visible = True
            m38.Visible = True
            Label6.Visible = True
            Label7.Visible = True
            DUM = ADD_FIELD("EGG", "M30", "CHAR(14)")
            DUM = ADD_FIELD("EGG", "M38", "CHAR(14)")
        End If

122     f_PROM_ANTITYPA = Val(FINDPARAMETROI(1, "PELAT43", "F_PROM_ANTITYPA", "2", "ΑΡΙΘΜΟΣ ΑΝΤΙΤΥΠΩΝ (ΠΩΛΗΣΕΩΝ)"))
124     F_PEL_ANTITYPA = Val(FINDPARAMETROI(1, "PELAT43", "F_PEL_ANTITYPA", "2", "ΑΡΙΘΜΟΣ ΑΝΤΙΤΥΠΩΝ (ΑΓΟΡΩΝ)"))

        'δεν αφήνει ημερομηνίες μικρότερες της τελευταίας καταχωρηθείσας
130     If (gUserId <> 4 And Len(Dir("c:\mercvb\ll.txt")) > 0) Or F_CanDelete = 0 Then
140         r2.Open "SELECT max(HME)  FROM EGG", Gdb, adOpenDynamic, adLockOptimistic

150         If CDate(Format(Now, "dd/mm/yyyy")) < CDate(Format(r2(0), "dd/mm/yyyy")) Then
160             MsgBox "ΥΠΑΡΧΕΙ ΗΜΕΡΟΜΗΝΙΑ ΜΕΓΑΛΥΤΕΡΗ ΤΗΣ ΣΗΜΕΡΙΝΗΣ "
                ' Unload Me
            End If

170         r2.Close
180         mLock = 1
190         DTPicker1.Enabled = False
        Else
200         mLock = 0
        End If

210     f_tameio = FINDPARAMETROI(1, "PELAT43", "F_TAMEIO", "9990", "ΛΟΓΑΡΙΑΣΜΟΣ ΤΑΜΕΙΟ")
220     r2.Open "SELECT * FROM PEL WHERE KOD='" + f_tameio + "' AND EIDOS='e'", Gdb, adOpenForwardOnly, adLockReadOnly

230     If r2.EOF Then
240         Gdb.Execute "insert into PEL (EIDOS,KOD,EPO) VALUES ('e','" + f_tameio + "','TAMEIO')"
        End If

250     r2.Close
      
260     DTPicker1.Value = gWorkDay

270     Me.Picture = LoadPicture(gPicture)
280     Me.WindowState = 0
290     Adodc1.ConnectionString = gConnect

        On Error GoTo create

        ' If gXEIRISTHS = 8 Then
300     Adodc1.RecordSource = "select * from XREOPIS where PEL ='" + gBuff + "'"
        '  Else
        '     Adodc1.RecordSource = "select * from XREOPIS where RIGHT(TITLOS,1)<>'.' AND PEL ='" + gBuff + "'"
        ' End If

310     Adodc1.Refresh

        On Error GoTo Form_Load_Err    'Resume Next

        Dim DDD

        On Error GoTo UPDATE_STR

320     Set R = Adodc1.Recordset
330     R.MoveFirst

340     Do While Not R.EOF
350         Combo1.AddItem Left(R("titlos") + Space$(50), 50) + Format(R!EIDOS, "000")
360         R.MoveNext
        Loop

        On Error Resume Next

370     R.MoveFirst
380     DDD = R("ARITMISI")

        On Error Resume Next

        'If gXEIRISTHS = 8 Then
        '     Combo1.AddItem "ΧΡΕΩΣΗ ΚΑΡΤΕΛΑΣ.."
        '     Combo1.AddItem "ΠΙΣΤΩΣΗ ΚΑΡΤΕΛΑΣ.."
        'End If

390     Combo1.Text = Combo1.List(2)    'metrhta

400     Text1.Text = DDD

410     R.Close

420     R.Open "SELECT *FROM PINAKES WHERE TYPOS=19 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

430     Do While Not R.EOF

440         If R("typos") = 19 Then
450             Aitiologies.AddItem R("PERIGRAFH")
            Else

                Exit Do

            End If

460         R.MoveNext
        Loop

        Exit Sub

UPDATE_STR:

        On Error Resume Next

470     Resume Next

create:

480     Gdb.Execute "CREATE TABLE XREOPIS (PEL CHAR(1),TITLOS CHAR(30),XREOPIS CHAR(1)  )"

490     Gdb.Execute "insert into  XREOPIS (PEL,TITLOS,XREOPIS ) VALUES ('e','ΧΡΕΩΣΗ ΠΕΛΑΤΗ','Χ')"
500     Gdb.Execute "insert into  XREOPIS (PEL,TITLOS,XREOPIS ) VALUES ('e','ΕΙΣΠΡΑΞΗ ΑΠΟ ΠΕΛΑΤΗ','Π')"
510     Gdb.Execute "insert into  XREOPIS (PEL,TITLOS,XREOPIS ) VALUES ('r','ΠΛΗΡΩΜΗ ΠΡΟΜΗΘΕΥΤΗ','Χ')"
520     Gdb.Execute "insert into  XREOPIS (PEL,TITLOS,XREOPIS ) VALUES ('r','ΧΡΕΩΣΗ ΠΡΟΜΗΘΕΥΤΗ','Χ')"

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label8_Click()

        '<EhHeader>
        On Error GoTo Label8_Click_Err

        '</EhHeader>

100     RichTextBox1.Height = 1455
110     RichTextBox1.Width = 3015

        '<EhFooter>
        Exit Sub

Label8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Label8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Label8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Text1_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Text1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Text1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Text1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub aitiologies_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo aitiologies_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

aitiologies_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.aitiologies_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.aitiologies_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Text3_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Text3_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Text3_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Text3_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text3_LostFocus()

        '<EhHeader>
        On Error GoTo Text3_LostFocus_Err

        '</EhHeader>

100     Text3.Text = Replace(Text3.Text, ",", ".")

        '<EhFooter>
        Exit Sub

Text3_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Text3_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Text3_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text4_GotFocus()
   DataGrid1.Visible = True
   
End Sub

Private Sub Text4_LostFocus()

        '<EhHeader>
        On Error GoTo Text4_LostFocus_Err

 
        '<EhFooter>
        Exit Sub

Text4_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat43.Text4_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat43.Text4_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub zoomXreopistoseis_Click()


Dim FF As New bohu4
'bohu4.MDIChild = False
bohu4.SHOW 1 'vbModal









End Sub
