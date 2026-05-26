VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form apot2ROULISSEARCH 
   Caption         =   "Form1"
   ClientHeight    =   11424
   ClientLeft      =   228
   ClientTop       =   -2388
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11424
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd≈ÓÔ‰ÔÚ 
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   360
      Left            =   14640
      TabIndex        =   7
      Top             =   8040
      Width           =   990
   End
   Begin VB.PictureBox listes 
      BackColor       =   &H00FFFF80&
      Height          =   15000
      Left            =   0
      ScaleHeight     =   14952
      ScaleWidth      =   16368
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   16415
      Begin VB.CommandButton mnext 
         Caption         =   ">>"
         Height          =   255
         Left            =   4840
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3120
         Width           =   495
      End
      Begin VB.CommandButton prev 
         Caption         =   "<<"
         Height          =   255
         Left            =   3880
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox lti 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   9000
         TabIndex        =   42
         Top             =   9240
         Width           =   1095
      End
      Begin VB.TextBox poso 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   7320
         TabIndex        =   41
         Top             =   9240
         Width           =   1215
      End
      Begin VB.TextBox kod 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   1800
         TabIndex        =   40
         Top             =   9240
         Width           =   1575
      End
      Begin VB.TextBox ono 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   3840
         TabIndex        =   39
         Top             =   9240
         Width           =   3015
      End
      Begin VB.TextBox lti 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   9000
         TabIndex        =   38
         Top             =   8160
         Width           =   1095
      End
      Begin VB.TextBox poso 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   7320
         TabIndex        =   37
         Top             =   8160
         Width           =   1215
      End
      Begin VB.TextBox kod 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   1800
         TabIndex        =   36
         Top             =   8160
         Width           =   1575
      End
      Begin VB.TextBox ono 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   3840
         TabIndex        =   35
         Top             =   8160
         Width           =   3015
      End
      Begin VB.TextBox lti 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   9000
         TabIndex        =   34
         Top             =   7080
         Width           =   1095
      End
      Begin VB.TextBox poso 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   7320
         TabIndex        =   33
         Top             =   7080
         Width           =   1215
      End
      Begin VB.TextBox kod 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   1800
         TabIndex        =   32
         Top             =   7080
         Width           =   1575
      End
      Begin VB.TextBox ono 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   3840
         TabIndex        =   31
         Top             =   7080
         Width           =   3015
      End
      Begin VB.TextBox lti 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   9000
         TabIndex        =   30
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox poso 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   7320
         TabIndex        =   29
         Top             =   6000
         Width           =   1215
      End
      Begin VB.TextBox kod 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   1800
         TabIndex        =   28
         Top             =   6000
         Width           =   1575
      End
      Begin VB.TextBox ono 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   3840
         TabIndex        =   27
         Top             =   6000
         Width           =   3015
      End
      Begin VB.TextBox lti 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   9000
         TabIndex        =   26
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox poso 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   7320
         TabIndex        =   25
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox kod 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   1800
         TabIndex        =   24
         Top             =   4920
         Width           =   1575
      End
      Begin VB.TextBox ono 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   3840
         TabIndex        =   23
         Top             =   4920
         Width           =   3015
      End
      Begin VB.TextBox lti 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   9000
         TabIndex        =   22
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox poso 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   7320
         TabIndex        =   21
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox kod 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   1824
         TabIndex        =   20
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox ono 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   3840
         TabIndex        =   19
         Top             =   3840
         Width           =   3015
      End
      Begin VB.PictureBox foto2 
         Height          =   975
         Index           =   5
         Left            =   120
         ScaleHeight     =   924
         ScaleWidth      =   1404
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   9240
         Width           =   1455
      End
      Begin VB.PictureBox foto2 
         Height          =   975
         Index           =   4
         Left            =   120
         ScaleHeight     =   924
         ScaleWidth      =   1404
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   8160
         Width           =   1455
      End
      Begin VB.PictureBox foto2 
         Height          =   975
         Index           =   3
         Left            =   120
         ScaleHeight     =   924
         ScaleWidth      =   1404
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   7080
         Width           =   1455
      End
      Begin VB.PictureBox foto2 
         Height          =   975
         Index           =   2
         Left            =   120
         ScaleHeight     =   924
         ScaleWidth      =   1404
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1455
      End
      Begin VB.PictureBox foto2 
         Height          =   975
         Index           =   1
         Left            =   120
         ScaleHeight     =   924
         ScaleWidth      =   1404
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   " ·Ë·ÒÈÛÏ¸Ú ÂÈÎÔ„ﬁÚ"
         Height          =   255
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3120
         Width           =   2775
      End
      Begin VB.PictureBox foto2 
         Height          =   975
         Index           =   0
         Left            =   120
         ScaleHeight     =   924
         ScaleWidth      =   1404
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ListBox Listnames 
         Height          =   432
         Left            =   11640
         TabIndex        =   8
         Top             =   120
         Width           =   3015
      End
      Begin VB.ListBox List 
         Height          =   3120
         Index           =   1
         Left            =   2760
         TabIndex        =   6
         Top             =   0
         Width           =   2775
      End
      Begin VB.ListBox List 
         Height          =   3120
         Index           =   2
         Left            =   5520
         TabIndex        =   5
         Top             =   0
         Width           =   2775
      End
      Begin VB.ListBox List_KAT 
         Height          =   3120
         Left            =   8640
         TabIndex        =   4
         Top             =   0
         Width           =   2775
      End
      Begin VB.ListBox List 
         Height          =   3120
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label eggrafes 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label4"
         Height          =   255
         Left            =   200
         TabIndex        =   47
         Top             =   3180
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "‘ÈÏﬁ"
         Height          =   255
         Left            =   9000
         TabIndex        =   46
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label LnMax 
         Caption         =   "–ÔÛ¸ÙÁÙ·"
         Height          =   255
         Left            =   7320
         TabIndex        =   43
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "–ÂÒÈ„Ò·ˆﬁ"
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   " ˘‰ÈÍ¸Ú"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "÷˘ÙÔ„Ò·ˆﬂ·"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3480
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc EGGTIM2 
      Height          =   495
      Left            =   15000
      Top             =   9960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   868
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
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Rich1 
      DataField       =   "MEMO"
      DataSource      =   "EID"
      Height          =   375
      Left            =   16920
      TabIndex        =   1
      Top             =   8880
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   656
      _Version        =   393217
      TextRTF         =   $"APOT2ROULIS.frx":0000
   End
   Begin MSAdodcLib.Adodc EID 
      Height          =   492
      Left            =   14760
      Top             =   2400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   868
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
      ConnectStringType=   3
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
      Caption         =   "EID"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7812
      Left            =   15180
      TabIndex        =   0
      Top             =   -60
      Width           =   252
   End
End
Attribute VB_Name = "apot2ROULISSEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private NewVersion As String

Private OldVersion As String

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long
                                       
'========================================================================================================
Dim F_id_LOCAL As Long

Dim gdblag     As New ADODB.Connection

Dim f_exit     As Boolean

Dim f_pel      As String

Dim f_xaraktiras(500)

Dim f_fylo(10)

Dim f_kataskeyasths(500)

Dim f_hlikies(30)

Dim F_FAKEL_KOINOS      As String

Dim f_site              As Integer

Dim f_arxikh_eikona(10) As String

Dim f_vasi              As String

'Private AutoClick As Boolean
'Private Declare Function GetActiveWindow Lib "user32" () As Long

Dim f_error             As Integer
                
'**************************************
' Name: Fun with MouseWheel
' Description:Just intercepting MouseWheel event with API. Make an empty project (standard exe) and paste code.
' By: vViktor
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=56768&lngWId=1'for details.'**************************************

Private Const PM_REMOVE = &H1

Private Type POINTAPI

    x As Long
    y As Long

End Type

Private Type Msg

    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    PT As POINTAPI

End Type

Private Declare Function PeekMessage _
                Lib "user32" _
                Alias "PeekMessageA" (lpMsg As Msg, _
                                      ByVal hWnd As Long, _
                                      ByVal wMsgFilterMin As Long, _
                                      ByVal wMsgFilterMax As Long, _
                                      ByVal wRemoveMsg As Long) As Long

Private Declare Function WaitMessage Lib "user32" () As Long

Private bCancel As Boolean

Private Const WM_MOUSEWHEEL = 522

'·ÔËÁÍÂıÂÈ Ù· ÛÙÔÈ˜ÂÈ· Ù˘Ì Í·ÙÁ„ÔÒÈ˘Ì
Dim fc_max          As Long

Dim FC_id(5000)     As Long

Dim FC_parenT(5000) As Long

Dim fc_names(5000)  As String

Dim fc_aa(5000)     As Integer

Dim names(5000)     As String

Private r0          As New ADODB.Recordset

Private nmax        As Long

Private Sub cmd≈ÓÔ‰ÔÚ_Click()
 
    Unload Me

End Sub

Private Sub Command1_Click()

    List_KAT.ListIndex = -1

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        cmd≈ÓÔ‰ÔÚ_Click
    End If

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     Me.AutoRedraw = True
        'Me.keyprev = True
        KeyPreview = True
102     f_pel = "mpoyg" ' "lagak"
 
        'syndesh xori dsn me connection string  ' ip toys 188.40.85.131
        ' gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=db5.papaki.gr;PORT =3306;DATABASE=vasi;UID=lagakis;Password=3898339214231"
104     On ERRROR GoTo NOCONNECT
106     f_site = Val(FINDPARAMETROI(1, "MDIFORM1", "F_SITE", "0", " Ÿƒ… œ” SITE 1=VM 2=TOYS"))
        'f_site = 222

108     If f_site = 222 Then
110         f_connect = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
112     ElseIf f_site = 333 Then
114         f_connect = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=db9.papaki.gr;PORT =3306;DATABASE=vasi3;UID=autoroulis;Password=autoroulis2013"
        Else
116         f_connect = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=db9.papaki.gr;PORT =3306;DATABASE=lagakis_new;UID=lagakis;Password=39213921"
        End If

        '104 gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=toys-shop.gr;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"

        On Error GoTo 0 ' Form_Load_Err

120     EID.ConnectionString = gConnect

        Dim DUM

        'GoTo PARAKAMCH

122     DUM = ADD_FIELD("EID", "MEMOENG", "TEXT")
124     DUM = ADD_FIELD("EID", "PIC0", "VARCHAR(100)")
126     DUM = ADD_FIELD("EID", "PIC1", "VARCHAR(100)")
128     DUM = ADD_FIELD("EID", "PIC2", "VARCHAR(100)")
130     DUM = ADD_FIELD("EID", "PIC3", "VARCHAR(100)")
132     DUM = ADD_FIELD("EID", "PIC4", "VARCHAR(100)")
134     DUM = ADD_FIELD("EID", "PIC5", "VARCHAR(100)")
136     DUM = ADD_FIELD("EID", "PIC6", "VARCHAR(100)")
    
138     DUM = ADD_FIELD("EID", "DIATHESIMO", "INT")

140     DUM = ADD_FIELD("EID", "ID_WEB", "INT")  ' TO ID TOY SITE
142     DUM = ADD_FIELD("EID", "HME_DIATH", "DATETIME")

144     DUM = ADD_FIELD("EID", "MANUFACTURERS_ID", "INT")

145     DUM = ADD_FIELD("EID", "TEXT_PRODUCTS_ZIPCODE", "VARCHAR(30)")

146     DUM = ADD_FIELD("EID", "ERGOSTASIO", "VARCHAR(30)")
148     DUM = ADD_FIELD("EID", "YOU_TUBE", "VARCHAR(100)")

150     DUM = ADD_FIELD("EID", "ADDITIONALPRICE", "DECIMAL(10,2)")
152     DUM = ADD_FIELD("EID", "SHIPPINGPRICE", "DECIMAL(10,2)")

154     DUM = ADD_FIELD("EID", "WEIGHT", "DECIMAL(10,2)")
156     DUM = ADD_FIELD("EID", "WEBPRICE", "DECIMAL(10,2)")

158     DUM = ADD_FIELD("EID", "FYLO", "INT")
160     DUM = ADD_FIELD("EID", "XARAKTIRAS", "INT")
162     DUM = ADD_FIELD("EID", "USED", "BIT") ' 1=METAX  0=NEW

        '' 139 DUM = ADD_FIELD("EID", "CATEGORY", "INT")

        'DIATHESIMO   INTEGER
        'HME_DIATH Date
        'KATASKEYASTHS VarChar(50)
        'ERGOSTASIO
        'Indv.Shipping Price:
        'Each Additional Price:
    
164     F_FAKEL_KOINOS = Trim(FINDPARAMETROI(1, "APOT2", "F_FAKEL_KOINOS", "C:\MERCVB", "KOINOXPH”TO” ÷¡ ≈Àœ” ≈… œÕŸÕ –.◊. \\PC\MERCVB"))
    
165     gdblag.Open f_connect
    
166     If f_site = 111 Then
168         f_vasi = "" + f_vasi + ""
170     ElseIf f_site = 333 Then
172         f_vasi = "abie3"
        Else
174         f_vasi = ""
        End If
    
        '
        '  category_parent_id  category_child_id   category_list   category_id vendor_id   category_name
        '0   1   NULL    1   1   ”Ù·ËÂÒÔﬂ HY
        '0   2   NULL    2   1   ÷ÔÒÁÙÔﬂ HY
        '0   3   NULL    3   1   –ÂÒÈˆÂÒÂÈ·Í‹
        '0   6   NULL    6   1   ¡Ì·Î˛ÛÈÏ·
        '0   19  NULL    19  1   ”ÙÈÍ‹ÍÈ· USB
        '0   14  NULL    14  1    ·Î˛‰È·
        '2   10  NULL    10  1   NET BOOKS
        '3   22  NULL    22  1   –ÔÌÙﬂÍÈ·
        '3   21  NULL    21  1   ≈ÍÙı˘Ù›Ú
        '6   11  NULL    11  1   Ã≈À¡Õ…¡-‘œÕ≈—
        '6   9   NULL    9   1   ˜·ÒÙÈ·
        '9   12  NULL    12  1   ¡’‘œ œÀ«‘≈” ≈‘… ≈‘≈”
        '9   13  NULL    13  1   PHOTO PAPER
        '11  15  NULL    15  1   LEXMARK
        '11  16  NULL    16  1   OKI
        '11  17  NULL    17  1   CANON
        '
        Dim R   As New ADODB.Recordset

        Dim sql As String

        'On Error Resume Next

        'lagakis
        'sql = "SELECT category_child_id AS ID,category_parent_id AS PARENT,category_name AS NAME   FROM jos_vm_category_xref a inner join jos_vm_category b on a.category_child_id=b.category_id  ORDER BY category_parent_id"
    
        'mpoygoudis

        '=============================  ¡‘«√œ—…≈” =====================================================

176     If f_site = 222 Then
178         sql = "SELECT a.categories_id as ID,a.parent_id as PARENT,b.CATEGORIES_NAME AS NAME from categories a inner join categories_description b on a.categories_id=b.categories_id order by parent_id,a.categories_id"
        Else  'vm
180         sql = "SELECT a.virtuemart_category_id as ID,b.category_parent_id as PARENT,a.category_name as NAME FROM " + f_vasi + "_virtuemart_categories_el_gr a inner join " + f_vasi + "_virtuemart_category_categories b on a.virtuemart_category_id=b.category_child_id"
        End If
     
        Dim N As Long

182     N = 0

        Dim error As Integer

        'products_id,options_id,options_values_id

        '================================ PRODUCT_ATTRIBUTES ================================
184     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'PRODUCT_ATTRIBUTES'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

186     If R(0) = 0 Then
188         Gdb.Execute "CREATE TABLE PRODUCT_ATTRIBUTES ( KOD VARCHAR(16) NOT NULL,OPTIONS_VALUES_ID INT,PRODUCT_ID int NOT NULL,OPTIONS_ID  INT  NULL ) "
        End If

190     R.Close

        '================================ PRODUCT_TO_CATEGORIES ================================
192     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'PRODUCT_TO_CATEGORIES'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

194     If R(0) = 0 Then
196         Gdb.Execute "CREATE TABLE PRODUCT_TO_CATEGORIES ( KOD VARCHAR(16) NOT NULL,PRODUCT_ID int NOT NULL,CATEGORIES_ID INT NOT NULL ) "
        End If

198     R.Close

        '================================ GNISIOI_KODIKOI ================================
200     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'GNISIOI_KODIKOI'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

202     If R(0) = 0 Then
204         Gdb.Execute "CREATE TABLE GNISIOI_KODIKOI ( KOD VARCHAR(16) NOT NULL,GNISIOS VARCHAR(30) NOT NULL) "
        End If

206     R.Close

208     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'CATEGORIESANT'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

210     If R(0) = 0 Then
212         Gdb.Execute "CREATE TABLE CATEGORIESANT " & "([ID] [int] NOT NULL,[PARENT] [INT] NOT NULL,[NAME] [varCHAR](80) NULL,CONSTRAINT [PK_CATEGORIESANT] PRIMARY KEY CLUSTERED " & "([ID] Asc )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY] "
        End If

214     R.Close

216     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'CATEGTEL'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

218     If R(0) = 0 Then
            'Gdb.Execute "CREATE TABLE CATEGTEL ( [ID] [int] NOT NULL,[PARENT] [INT] NOT NULL,[NAME] [varCHAR](50) NULL,[AA] [INT] NOT NULL )ON PRIMARY  "
220         Gdb.Execute "CREATE TABLE CATEGTEL ( [ID] [int] NOT NULL,[PARENT] [INT] NOT NULL,[NAME] [varCHAR](80) NULL,[AA] [INT] NOT NULL )  "
        End If

222     R.Close

        '≈À≈√◊Ÿ ¡Õ ≈Õ«Ã≈—Ÿ»« ≈ ”«Ã≈—¡
224     R.Open "SELECT OBJECT_NAME(OBJECT_ID) AS DatabaseName, last_user_update AS HME  FROM sys.dm_db_index_usage_stats  WHERE database_id = DB_ID( 'MERCURY')  AND OBJECT_ID=OBJECT_ID('CATEGORIESANT')", Gdb, adOpenDynamic, adLockOptimistic

226     If Not R.EOF Then
228         If Not IsNull(R!hme) Then
230             If Format(R!hme, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
232                 ' r.Close
                    ' GoTo PARAKAMCH
                End If
            End If
        End If

234     R.Close
        '
        '
        '

236     f_err = 0
238     Gdb.Execute "DELETE FROM CATEGORIESANT"
240     Gdb.Execute "DELETE FROM CATEGTEL"

242     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic

244     Do While Not R.EOF
246         Gdb.Execute "INSERT INTO CATEGORIESANT (ID,PARENT,NAME) VALUES (" + str(R!ID) + "," + str(R!Parent) + ",'" + Replace(Left(R!Name, 80), "'", "''") + "')"
248         R.MoveNext

250         If f_err = 1 Then Exit Do
        Loop

252     R.Close

254     Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,1 FROM CATEGORIESANT WHERE PARENT=0"
256     Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,2  FROM CATEGORIESANT WHERE PARENT IN (SELECT ID FROM CATEGTEL)"

        ' GIA NA FYGEI  TO 1&2  EPIPEDO APO TO CATEGORIESANT
258     Gdb.Execute "DELETE FROM CATEGORIESANT WHERE ID IN (SELECT ID FROM CATEGTEL)"

260     Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,3  FROM CATEGORIESANT WHERE PARENT IN (SELECT ID FROM CATEGTEL)"

        ' GIA NA FYGEI  TO 3  EPIPEDO APO TO CATEGORIESANT
262     Gdb.Execute "DELETE FROM CATEGORIESANT WHERE ID IN (SELECT ID FROM CATEGTEL)"

264     Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,4  FROM CATEGORIESANT WHERE PARENT IN (SELECT ID FROM CATEGTEL)"

        ' GIA NA FYGEI  TO 4  EPIPEDO APO TO CATEGORIESANT
266     Gdb.Execute "DELETE FROM CATEGORIESANT WHERE ID IN (SELECT ID FROM CATEGTEL)"

268     Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,5  FROM CATEGORIESANT"

PARAKAMCH:

        '---------------------- „ÂÏÈÛÏ· ÙÔı treeview control ----------------------------

270     R.Open "SELECT * FROM CATEGTEL ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic
        ' 192 r.MoveFirst

        'r.Open "SELECT * FROM CATEG1", Gdb
300     R.Close
302     fc_max = N

        '-----------------------------  ÍÔÏÏ·ÙÈ ÒÔıÎÁ ------------------------------------------------------------------
        ' gia ta listbox mono
304     R.Open "SELECT * FROM CATEGTEL where AA=1 ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic

        'Dim roulis(5000, 4)
        Dim i As Long

306     i = 0

308     Do While Not R.EOF
310         List(0).AddItem R!Name, i
312         List(0).ItemData(i) = R!ID
314         i = i + 1
316         R.MoveNext
        Loop

318     R.Close

        GoTo parakamch2
     
414     sql = "SELECT * FROM products_options_values where language_id=4  order by products_options_values_id "
416     n7 = 0
418     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic

420     Do While Not R.EOF
422         hlikies.AddItem R!products_options_values_name
424         f_hlikies(n7) = R!products_options_values_id
426         n7 = n7 + 1
 
428         R.MoveNext

430         If f_err = 1 Then Exit Do
        Loop

432     R.Close
434     kataskeyasths.ListIndex = 0
     
parakamch2:
     
436     R.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
        'FPA
        'R.Seek "=", 1, 1
        '438 Do While Not R.EOF
        '440     If R("typos") = 1 Then
        '442        Exit Do ' Combo2.AddItem str(R("AYJON")) + " -> " + str(R("TIMH"))
        '         End If
        '444     R.MoveNext
        '446     If f_err = 1 Then Exit Do
        '     Loop
        '     ' mon.metrhshs
        '448 R.Close
     
450     'Combo2.ListIndex = 1
   
        ' sql = "SELECT * FROM abie3_virtuemart_customs where virtuemart_custom_id=22  "
 
        Dim f As String

        Dim gg
 
452     sql = "SELECT * FROM " + f_vasi + "_virtuemart_customs where virtuemart_custom_id=23  "

        R.Close

        ' DEBUG NEXT 3 LINES
        '  List(2).AddItem "«À≈ ‘—… ¡"
        '   List(2).AddItem "÷—≈Õ¡"
    
        '  Exit Sub

454     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic
 
456     f = R!custom_value
458     gg = Split(f, ";")

460     For k = LBound(gg) To UBound(gg)
462         List_KAT.AddItem gg(k)
            'KATHGORIA.ItemData(k) = gg(k)
        Next

464     ' List_KAT.ListIndex = 0
466     R.Close

        'refresh_typos
        'List_KAT.ListIndex = 0
     
        '    Dim gg2
        '468 sql = "SELECT * FROM " + f_vasi + "_virtuemart_customs where virtuemart_custom_id=23"
        '470 r.Open sql, gdblag, adOpenDynamic, adLockOptimistic
        '
        '
        '472  f = r!custom_value
        '474  gg2 = Split(f, ";")
        '476  For k = LBound(gg2) To UBound(gg2)
        '           If Len(Trim(gg2(k))) > 0 Then
        '478           EIDOS.AddItem gg2(k)
        '           End If
        '
        '     Next
        '480  EIDOS.ListIndex = 0
        '482  r.Close
        '
     
484     If f_err = 1 Then
486         MsgBox "ƒÂÌ ˆÔÒÙ˛ËÁÍ·Ì Û˘ÛÙ‹ ÔÈ ﬂÌ·ÍÂÚ."
        End If
     
        'MARKA.ListIndex = 0
        '488 ftp.ListIndex = 0
     
        ' Me.Print "Please use now mouse wheel to move this form."
490     Me.SHOW
        ' ProcessMessages

        '<EhFooter>
        Exit Sub
        
NOCONNECT:
        MsgBox "¡ƒ’Õ¡‘« « ”’Õƒ≈”« Ã≈ ‘œ SITE. ƒœ …Ã¡”‘≈ ¡—√œ‘≈—¡"
        Unload Me

        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.APOT2ROULIS.Form_Load " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        f_err = 1

        Resume Next

        '</EhFooter>

End Sub

Private Sub list_Click(Index As Integer)
   
    If Index < 3 Then
   
        MOVELIST Index
   
    End If
   
    '124    If list(INDEX + 1).ListCount > 0 Then
    '126      list(INDEX + 1).SetFocus
    '128      list(INDEX + 1).ListIndex = 0
    '    End If
End Sub

Private Sub List_KAT_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 13 Then

        Exit Sub

    End If

    SHOW_ITEMS
End Sub

Private Sub list_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

100     If KeyCode <> 13 Then

            Exit Sub

        End If

        MOVELIST Index
        SHOW_ITEMS
End Sub
    
Sub SHOW_ITEMS()
    
    Dim R     As New ADODB.Recordset

    Dim i     As Long

    Dim lItem As ListItem
    
    Listnames.Clear
    
    'MOVELIST INDEX
    
    Dim D As Long
    
    D = Index + 1

    If D >= 3 Then D = 2
    
    'If List(D).ListCount > 0 Then
    '   List(D).SetFocus
    '   List(D).ListIndex = 0
    'Else
    
    Dim sql
    
    Dim ID_OPEL As String

    Dim synt    As String
    
    synt = ""
    
    If List_KAT.ListIndex > -1 Then
    
        synt = "YOU_TUBE='" + List_KAT.Text + "' AND "
    
    End If
    
    '     ' ƒ…¡À≈Œ≈ ÃœÕœ   "OPEL"  O–OTE  List(1).ListIndex = -1
    '  If List(1).ListIndex = -1 Then
    '    ID_OPEL = str(List(0).ItemData(List(0).ListIndex))
    '    sql = "select COUNT(*) from EID where " + synt + " KOD in (select KOD from PRODUCT_TO_CATEGORIES where CATEGORIES_ID IN ("
    '    sql = sql + "SELECT ID  From CATEGTEL  WHERE PARENT  IN (SELECT ID  From CATEGTEL WHERE PARENT=" + ID_OPEL + " )"
    '    sql = sql + " Union  SELECT ID From CATEGTEL WHERE PARENT = " + ID_OPEL + " ))"
    '  ElseIf List(2).ListIndex = -1 Then ' ƒ…¡À≈Œ≈ ÃœÕœ   "OPEL->VECTRA"  O–OTE  List(2).ListIndex = -1
    '    ID_OPEL = str(List(1).ItemData(List(1).ListIndex))
    '    sql = "select  COUNT(*)  from EID where  " + synt + " KOD in (select KOD from PRODUCT_TO_CATEGORIES where CATEGORIES_ID IN ("
    '    sql = sql + "SELECT ID  From CATEGTEL  WHERE  PARENT=" + ID_OPEL + " "
    '    sql = sql + " Union  SELECT ID From CATEGTEL WHERE ID = " + ID_OPEL + " ))"
    '   ElseIf List(2).ListIndex > -1 Then ' ƒ…¡À≈Œ≈   "OPEL->VECTRA->1200CC"  O–OTE  List(2).ListIndex = -1
    '    ID_OPEL = str(List(2).ItemData(List(2).ListIndex))
    '    sql = "select  COUNT(*)  from EID where  " + synt + "  KOD in (select KOD from PRODUCT_TO_CATEGORIES where CATEGORIES_ID IN ("
    '    sql = sql + " SELECT ID From CATEGTEL WHERE ID = " + ID_OPEL + " ))"
    '  End If
    '
    '  R.Open sql, Gdb, adOpenDynamic, adLockOptimistic
    '
    '  Dim nMax As Long
    '
    '  nMax = R(0)
    '  Label4.Caption = "≈„„Ò·ˆ›Ú " + str(nMax)
    '  R.Close
    
    ' ƒ…¡À≈Œ≈ ÃœÕœ   "OPEL"  O–OTE  List(1).ListIndex = -1
    If List(1).ListIndex = -1 Then
        ID_OPEL = str(List(0).ItemData(List(0).ListIndex))
        sql = "select  ROW_NUMBER() OVER( ORDER BY KOD DESC) AS Row,* from EID where " + synt + " KOD in (select KOD from PRODUCT_TO_CATEGORIES where CATEGORIES_ID IN ("
        sql = sql + "SELECT ID  From CATEGTEL  WHERE PARENT  IN (SELECT ID  From CATEGTEL WHERE PARENT=" + ID_OPEL + " )"
        sql = sql + " Union  SELECT ID From CATEGTEL WHERE PARENT = " + ID_OPEL + " ))"
    ElseIf List(2).ListIndex = -1 Then ' ƒ…¡À≈Œ≈ ÃœÕœ   "OPEL->VECTRA"  O–OTE  List(2).ListIndex = -1
        ID_OPEL = str(List(1).ItemData(List(1).ListIndex))
        sql = "select  ROW_NUMBER() OVER( ORDER BY KOD DESC) AS Row,* from EID where  " + synt + " KOD in (select KOD from PRODUCT_TO_CATEGORIES where CATEGORIES_ID IN ("
        sql = sql + "SELECT ID  From CATEGTEL  WHERE  PARENT=" + ID_OPEL + " "
        sql = sql + " Union  SELECT ID From CATEGTEL WHERE ID = " + ID_OPEL + " ))"
    ElseIf List(2).ListIndex > -1 Then ' ƒ…¡À≈Œ≈   "OPEL->VECTRA->1200CC"  O–OTE  List(2).ListIndex = -1
        ID_OPEL = str(List(2).ItemData(List(2).ListIndex))
        sql = "select  ROW_NUMBER() OVER( ORDER BY KOD DESC) AS Row,* from EID where  " + synt + "  KOD in (select KOD from PRODUCT_TO_CATEGORIES where CATEGORIES_ID IN ("
        sql = sql + " SELECT ID From CATEGTEL WHERE ID = " + ID_OPEL + " ))"
    End If

    On Error Resume Next

    r0.Close
    r0.Open sql, Gdb, adOpenDynamic, adLockOptimistic
    
    ' Dim nmax As Long
    r0.MoveLast
    nmax = r0("row")
    eggrafes.Caption = "≈„„Ò·ˆ›Ú " + str(nmax)
    r0.MoveFirst
  
    prev.Enabled = True
    mnext.Enabled = True
    
    foto_grid
    
    '  R.Open "select * from EID where KOD in (select KOD from PRODUCT_TO_CATEGORIES where CATEGORIES_ID=" + str(List(Index).ItemData(List(Index).ListIndex)) + ")", Gdb, adOpenDynamic, adLockOptimistic
    'PRODUCT_TO_CATEGORIES ( KOD VARCHAR(16) NOT NULL,PRODUCT_ID int NOT NULL,CATEGORIES_ID INT NOT NULL )
    'ListView1.ListItems.Clear
    i = 0
        
    '        F_FAKEL_KOINOS = "c:\mercvb"
    '
    'Do While Not R.EOF
    '           Listnames.AddItem r0!KOD + " " + r0!ONO + " " + Format(r0!LTI5, "0000.00")
    ''           Set lItem = ListView1.ListItems.ADD(, , R!KOD)  ' Rnd(100))
    ''           lItem.SubItems(1) = R!ONO
    ''           lItem.SubItems(2) = str(R!LTI5)
    '
    '
    '
    '
    '     KOD(i).Text = r0!KOD
    '     ONO(i).Text = r0!ONO
    '     POSO(i).Text = IIf(IsNull(r0!POS), 0, r0!POS)
    '    LTI(i).Text = r0!LTI5
    '
    '    If IsNull(r0!pic0) Then
    '       foto2(i).Picture = LoadPicture()
    '    Else
    '       If Len(r0!pic0) > 0 Then
    '          foto2(i).Picture = LoadPicture()
    '          foto2(i).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + r0!pic0)
    '          StretchSourcePictureFromPicture apot2ROULISSEARCH.foto2(i).Picture, apot2ROULISSEARCH.foto2(i)
    '       Else
    '          foto2(i).Picture = LoadPicture()
    '       End If
    '    End If
    '    i = i + 1
    '    If i > 5 Then Exit Do
    '    r0.MoveNext
    'Loop

    '
    '      If Len(Dir(F_FAKEL_KOINOS + "\images\" + PIC(n).Caption)) < 2 Then
    '          PIC(n).Caption = ""
    '      End If
    '
    '
    '    If Len(PIC(n).Caption) > 0 Then
    '       foto2(n).Picture = LoadPicture()
    '       foto2(n).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + PIC(n).Caption)
    '       StretchSourcePictureFromPicture apot2ROULIS.foto2(n).Picture, apot2ROULIS.foto2(n)
    '    Else
    '       foto2(n).Picture = LoadPicture()
    '    End If

End Sub

Sub foto_grid()

    ' F_FAKEL_KOINOS = "c:\mercvb"
    Dim i As Integer

    i = 0

    Dim first As Long

    first = r0(0)

    On Error Resume Next

    Do While Not r0.EOF
        Listnames.AddItem r0!kod + " " + r0!ono + " " + Format(r0!lti5, "0000.00")
        '           Set lItem = ListView1.ListItems.ADD(, , R!KOD)  ' Rnd(100))
        '           lItem.SubItems(1) = R!ONO
        '           lItem.SubItems(2) = str(R!LTI5)
           
        kod(i).Text = r0!kod
        ono(i).Text = r0!ono
        poso(i).Text = IIf(IsNull(r0!POS), 0, r0!POS)
   
        If IsNull(r0!lti5) Then
            lti(i).Text = 0
        Else
            lti(i).Text = r0!lti5
        End If
   
        If IsNull(r0!pic0) Then
            foto2(i).Picture = LoadPicture()
        Else

            If Len(r0!pic0) > 0 Then
                foto2(i).Picture = LoadPicture()

                On Error Resume Next

                foto2(i).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + r0!pic0)
                StretchSourcePictureFromPicture apot2ROULISSEARCH.foto2(i).Picture, apot2ROULISSEARCH.foto2(i)
            Else
                foto2(i).Picture = LoadPicture()
            End If
        End If

        i = i + 1

        If i > 5 Then Exit Do
        r0.MoveNext
    Loop

    If r0.EOF Then r0.MoveLast

    eggrafes.Caption = "≈„„Ò·ˆ›Ú " + Format(first, "####") + "-" + Format(r0(0), "####") + " ”˝ÌÔÎÔ " + Format(nmax, "####")

    If i > 5 Then Exit Sub

    Dim k As Integer

    For k = i To 5
        foto2(k).Picture = LoadPicture()
        kod(k).Text = ""
        ono(k).Text = ""
        poso(k).Text = ""
        lti(k).Text = ""
    Next

End Sub

Sub MOVELIST(Index As Integer)

        Dim i As Long

        Dim R As New ADODB.Recordset

        On Error GoTo MOVELIST

102     If Index <= 3 And List(Index).ListCount > 0 Then
104         R.Open "SELECT * FROM CATEGTEL where AA=" + str(Index + 2) + " AND PARENT=" + str(List(Index).ItemData(List(Index).ListIndex)) + " ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic
            'Dim roulis(5000, 4)
            'Dim I As Long
106         i = 0

108         If (Index + 1) <= 2 Then List(Index + 1).Clear
110         If (Index + 2) <= 2 Then List(Index + 2).Clear
112         If (Index + 3) <= 2 Then List(Index + 3).Clear
 
114         Do While Not R.EOF
116             List(Index + 1).AddItem R!Name, i
118             List(Index + 1).ItemData(i) = R!ID
120             R.MoveNext
            Loop

122         R.Close
   
124         If List(Index).ListCount > 0 Then
126             ' list(INDEX + 1).SetFocus
128             ' list(INDEX + 1).ListIndex = 0
            End If

        End If

        Exit Sub '

MOVELIST:
130     MsgBox "MOVELIST" + str(Erl)

        Resume Next

End Sub

Private Sub List1_Click()
    '   Dim a
    '   a = Left(List1.List(List1.ListIndex), 15)
    '
    '   kod.Text = a
    '   EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + a + "'"
    '   EID.Refresh
    '
    '    If Not EID.Recordset.EOF Then
    '      ' UPDATE_CONTROLS
    '    Else
    '       Exit Sub
    '    End If
    
End Sub

Private Sub ProcessMessages()

    Dim Message As Msg
 
    Dim s

    s = 200

    Do While Not bCancel

        If f_exit = True Then
            Unload Me
        End If
 
        WaitMessage 'Wait For message and...

        If PeekMessage(Message, Me.hWnd, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then '...when the mousewheel is used...
            If Message.wParam < 0 Then '...scroll up...

                If VScroll1.Value + s >= VScroll1.Max Then
                    VScroll1.Value = VScroll1.Max
                Else
                    VScroll1.Value = VScroll1.Value + s
                End If
 
                ' Me.Top = Me.Top + 240
            Else '... or scroll down
 
                If VScroll1.Value - s <= VScroll1.Min Then
                    VScroll1.Value = VScroll1.Min
                Else
                    VScroll1.Value = VScroll1.Value - s
                End If

                Me.Caption = time$
                '  Me.Top = Me.Top - 240
            End If
        End If

        DoEvents
 
        ' Unload Me
 
    Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' bCancel = True
    ' Me.Hide
    On Error Resume Next

    gdblag.Close
 
End Sub

Private Sub mnext_Click()
    r0.MoveNext

    If r0.EOF Then
        r0.MoveLast
        mnext.Enabled = False
    End If

    prev.Enabled = True
    foto_grid
End Sub

Private Sub prev_Click()

    Dim k As Integer

    For k = 0 To 5
        r0.MovePrevious

        If r0.BOF Then
            r0.MoveFirst
            prev.Enabled = False
        End If

    Next
 
    mnext.Enabled = True
 
    foto_grid
End Sub
