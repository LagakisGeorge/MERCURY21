VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "oe"
   ClientHeight    =   10470
   ClientLeft      =   165
   ClientTop       =   615
   ClientWidth     =   15090
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0442
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   9735
      Left            =   0
      ScaleHeight     =   9735
      ScaleWidth      =   3045
      TabIndex        =   2
      Top             =   420
      Width           =   3045
      Begin VB.PictureBox picSet 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   34
         ScaleHeight     =   3495
         ScaleWidth      =   2940
         TabIndex        =   8
         Top             =   6015
         Visible         =   0   'False
         Width           =   2940
         Begin MSComctlLib.ListView Listview2 
            Height          =   3228
            Left            =   36
            TabIndex        =   9
            Top             =   0
            Width           =   2988
            _ExtentX        =   5265
            _ExtentY        =   5689
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDragMode     =   1
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483641
            BackColor       =   -2147483634
            Appearance      =   0
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "MDIForm1.frx":1F6FC8
            OLEDragMode     =   1
            NumItems        =   0
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   33
            TabIndex        =   10
            Top             =   180
            Width           =   2055
         End
      End
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   34
         ScaleHeight     =   4575
         ScaleWidth      =   2940
         TabIndex        =   3
         Top             =   1080
         Width           =   2940
         Begin MSComctlLib.ListView Listview1 
            Height          =   4500
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2988
            _ExtentX        =   5265
            _ExtentY        =   7938
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDragMode     =   1
            _Version        =   393217
            Icons           =   "i32x32"
            SmallIcons      =   "i32x32"
            ForeColor       =   -2147483641
            BackColor       =   -2147483634
            Appearance      =   0
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "MDIForm1.frx":1F78A2
            OLEDragMode     =   1
            NumItems        =   0
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   33
            TabIndex        =   5
            Top             =   180
            Width           =   2055
         End
      End
      Begin lvButton.lvButtons_H cmdSet 
         Height          =   375
         Left            =   30
         TabIndex        =   6
         Top             =   5655
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   661
         Caption         =   "System Setings"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
      Begin lvButton.lvButtons_H cmdFile 
         Height          =   375
         Left            =   30
         TabIndex        =   7
         Top             =   675
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   661
         Caption         =   "Quick Launch"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
      Begin VB.Image Image1 
         Height          =   23145
         Index           =   0
         Left            =   3000
         Picture         =   "MDIForm1.frx":1F817C
         Stretch         =   -1  'True
         Top             =   -1560
         Width           =   570
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   0
         Picture         =   "MDIForm1.frx":1F8374
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today is :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   45
         TabIndex        =   14
         Top             =   465
         Width           =   630
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   750
         TabIndex        =   13
         Top             =   450
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome ,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   45
         TabIndex        =   12
         Top             =   255
         Width           =   645
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   750
         TabIndex        =   11
         Top             =   240
         Width           =   180
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   10155
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2893
            MinWidth        =   2893
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "GDIR"
            TextSave        =   "GDIR"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7937
            MinWidth        =   7937
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      HelpContextID   =   65278
      ImageList       =   "imlToolbarPics"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Table"
            Object.ToolTipText     =   "Table type Recordset"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dynaset"
            Object.ToolTipText     =   "Dynaset type Recordset"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Schedule"
            Object.ToolTipText     =   "Ημερολόγιο"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BIBLIO"
            ImageIndex      =   8
         EndProperty
      EndProperty
      Begin MSAdodcLib.Adodc TELEFON 
         Height          =   330
         Left            =   2880
         Top             =   30
         Visible         =   0   'False
         Width           =   2640
         _ExtentX        =   4657
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
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   7620
         Top             =   -30
      End
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   -210
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FilterIndex     =   1144
   End
   Begin MSComctlLib.ImageList imlToolbarPics 
      Left            =   1605
      Top             =   2295
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F856C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F867E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F8790
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F88A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F89B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F8AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F8BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F8CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F8DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F8F0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu pelat 
      Caption         =   "&Πελάτες/Προμηθευτές"
      Begin VB.Menu mpelat1 
         Caption         =   "&1.Εισαγωγή νέου"
         Shortcut        =   ^P
      End
      Begin VB.Menu mpelat2 
         Caption         =   "&2.Διόρθωση/Διαγραφή"
      End
      Begin VB.Menu mpelat3 
         Caption         =   "&3.Παρουσίαση"
      End
      Begin VB.Menu mpelat4 
         Caption         =   "&4.Καρτέλλα"
         Shortcut        =   ^K
      End
      Begin VB.Menu mpelat5 
         Caption         =   "&5.Παρουσίαση Κινήσεων"
      End
      Begin VB.Menu exit 
         Caption         =   "&6.Εξοδος"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu apot 
      Caption         =   "&Αποθήκη"
      Begin VB.Menu mapot1 
         Caption         =   "&1.Εισαγωγή νέου είδους"
         Index           =   21
         Shortcut        =   ^E
      End
      Begin VB.Menu mapot2 
         Caption         =   "&2.Διόρθωση είδους"
         Index           =   22
         Shortcut        =   ^D
      End
      Begin VB.Menu pareid 
         Caption         =   "&3.Παρουσίαση ειδών"
         Index           =   23
         Begin VB.Menu mapot31 
            Caption         =   "Είδη Αποθήκης"
            Index           =   310
         End
      End
      Begin VB.Menu apo4 
         Caption         =   "&4.Τιμοκατάλογος Ειδών"
         Index           =   24
      End
      Begin VB.Menu apo5 
         Caption         =   "&5.Κίνηση Ειδών"
         Index           =   25
      End
      Begin VB.Menu mapot6 
         Caption         =   "&6.Οικογένειες-Υποοικογένειες"
         Index           =   26
      End
      Begin VB.Menu mapot7 
         Caption         =   "&7.Ρυθμιζόμενες εκτυπώσεις"
         Index           =   27
      End
   End
   Begin VB.Menu epit 
      Caption         =   "&Επιταγές"
      Begin VB.Menu eisepit 
         Caption         =   "&1.Eισαγωγή Επιταγής"
      End
      Begin VB.Menu diorepit 
         Caption         =   "&2.Στατιστικά"
      End
      Begin VB.Menu parepit 
         Caption         =   "&3.Παρουσίαση Επιταγών"
      End
   End
   Begin VB.Menu par 
      Caption         =   "Παρασ&τατικά"
      Begin VB.Menu mpar1 
         Caption         =   "&1.Εισαγωγή Παρασ/κού"
         Index           =   41
         Shortcut        =   ^T
      End
      Begin VB.Menu diagrpar 
         Caption         =   "&2.Προβολή Παραστατικών κατά ημέρα"
         Index           =   42
      End
      Begin VB.Menu parenos 
         Caption         =   "&3.Παρουσίαση ενός Παραστατικού"
         Index           =   43
      End
      Begin VB.Menu timspar 
         Caption         =   "&4.Παρουσίαση Παραστατικών"
         Index           =   44
      End
      Begin VB.Menu mpar5 
         Caption         =   "&5.Εικόνα Επιχείρησης"
      End
      Begin VB.Menu mPar6 
         Caption         =   "&6.Ταμείο ημέρας"
      End
      Begin VB.Menu mpar7 
         Caption         =   "&7.Μαζικός μετασχηματισμός"
      End
      Begin VB.Menu mpar8 
         Caption         =   "&8.Εισαγωγή Δελτίων Ποσοτικής Παραλαβής"
      End
      Begin VB.Menu mpar9 
         Caption         =   "9.Σύνολα φόρτωσης"
      End
   End
   Begin VB.Menu bohu 
      Caption         =   "&Βοηθητικά"
      Begin VB.Menu mBohu1 
         Caption         =   "&1.Πίνακες"
         Index           =   51
      End
      Begin VB.Menu mbohu2 
         Caption         =   "&2.Παραμετροι Προγράμματος"
         Index           =   52
      End
      Begin VB.Menu par3 
         Caption         =   "&3.Παράμετροι Παραστατικών"
         Index           =   53
      End
      Begin VB.Menu parkin 
         Caption         =   "&4.Παράμετροι κινήσεων πελατών"
         Index           =   54
      End
      Begin VB.Menu pros 
         Caption         =   "&5.Import-Export"
         Index           =   55
      End
      Begin VB.Menu genlog 
         Caption         =   "&6.Ημερολόγιο Γεν.Λογιστικής"
         Index           =   56
      End
      Begin VB.Menu DBF 
         Caption         =   "&7.Αποστολή σε Φορητά Τερματικά"
         Index           =   57
      End
      Begin VB.Menu pontoif 
         Caption         =   "&8.Μεταφορές Αρχείων-Πόντοι"
         Index           =   58
      End
      Begin VB.Menu apoemporiko 
         Caption         =   "&9.Aντίγραφα / Επαναφορά"
         Index           =   59
      End
      Begin VB.Menu calculator 
         Caption         =   "10.Αριθμομηχανή"
         Index           =   510
         Shortcut        =   {F2}
      End
      Begin VB.Menu mbohu11 
         Caption         =   "11.Αλλαγή κωδικών χρήστη"
      End
      Begin VB.Menu mbohu12 
         Caption         =   "12.Συγκεντρωτική Τιμολογίων"
      End
      Begin VB.Menu mBohu13 
         Caption         =   "13.Ελεγχοι πληρότητας αρχείων"
      End
   End
   Begin VB.Menu helpform 
      Caption         =   "Βοήθεια"
      Begin VB.Menu man 
         Caption         =   "Βιβλίο οδηγιών"
      End
      Begin VB.Menu ekdoseis 
         Caption         =   "Εκδόσεις"
      End
   End
   Begin VB.Menu RUNTIME 
      Caption         =   "RUNTIMEMENU"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu runt 
         Caption         =   "RUNT"
         Index           =   0
      End
      Begin VB.Menu runt2 
         Caption         =   "runt2"
      End
   End
   Begin VB.Menu par1menu 
      Caption         =   "Par1"
      Visible         =   0   'False
      Begin VB.Menu par2show 
         Caption         =   "Παρουσίαση Παρ/κού"
      End
      Begin VB.Menu diortosis 
         Caption         =   "Διόρθωση"
      End
      Begin VB.Menu ektyp_par 
         Caption         =   "Εκτύπωση Παρ/κού"
      End
      Begin VB.Menu metasx_par 
         Caption         =   "Μετασχηματισμός Παρ/κού"
      End
   End
   Begin VB.Menu EPITAG311 
      Caption         =   "EPITAG311"
      Visible         =   0   'False
      Begin VB.Menu epit311 
         Caption         =   "epit311"
         Index           =   0
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'c5450
Option Explicit
Dim ETAIR(20), mc As Integer    ' POPUP MENU
Attribute mc.VB_VarUserMemId = 1073938432
Dim F_TOOLBAR As Long
Attribute F_TOOLBAR.VB_VarUserMemId = 1073938434
Dim f_Ypenth As Long
Attribute f_Ypenth.VB_VarUserMemId = 1073938435
Dim f_supervisor_code
Attribute f_supervisor_code.VB_VarUserMemId = 1073938436
Dim f_telephone As Integer
Attribute f_telephone.VB_VarUserMemId = 1073938437
Dim F_BACKGROUND As Integer
Attribute F_BACKGROUND.VB_VarUserMemId = 1073938438


' Dim f_loadpar1 As Integer




Private Sub antigrafa_Click(Index As Integer)
'   BOHU8.Show

End Sub

Private Sub Adodc1_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub apo4_Click(Index As Integer)
    gApoMenu = True
    apot4.SHOW

End Sub



Private Sub apo5_Click(Index As Integer)
    gApoMenu = True
    APOT5.SHOW
End Sub

Private Sub apoemporiko_Click(Index As Integer)
    gApoMenu = True

    bohu8.SHOW


End Sub



Private Sub calculator_Click(Index As Integer)
    Dim DUM
    DUM = Shell("calc.exe", vbMaximizedFocus)

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub DBF_Click(Index As Integer)
    gApoMenu = True
    bohu6.SHOW

End Sub

Private Sub diagrpar_Click(Index As Integer)


    gApoMenu = True
    par2frmOrder.SHOW

End Sub



Private Sub diortosis_Click()
    par1.diortosis
    
    If Len(par1.Grid1.TextMatrix(1, 1)) = 0 Then
        
        MsgBox "OK"
        par1.diortosis
    
    End If
    
    
    
    
End Sub

Private Sub eisepit_Click()
    gApoMenu = True
    EPIT1.SHOW

End Sub



Private Sub ekdoseis_Click()


    help1.WindowState = 2   'maximized



    help1.SHOW


















End Sub

Private Sub ektyp_par_Click()

    par1.diortosis
    DoEvents

    par1.Command5_Click

End Sub

Private Sub epit311_Click(Index As Integer)
    epit3.epit311_Click Index

End Sub

Private Sub exit_Click()

    gApoMenu = True
    Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub genlog_Click(Index As Integer)
    gApoMenu = True
    bohu5.SHOW
End Sub

Private Sub Image1_Click()

    gApoMenu = True
    HMEROL2.SHOW
End Sub

Private Sub Image2_Click()


    gApoMenu = True
    hmerol.SHOW

End Sub

Private Sub helpform_Click()
''  helpforma.SHOW
'
'
'  Dim sFilename As String
'Dim appWord As Word.Application
'Dim WRDdoc As Word.Document
'
'sFilename = "c:\mercvb\help\help.doc"
'Set appWord = New Word.Application
''Set WRDdoc = appWord.Documents.Open(sFilename)
''appWord.Visible = True












End Sub

Private Sub man_Click()



    Dim SFILENAME As String
    Dim appWord As Word.Application
    Dim WRDdoc As Word.Document

    SFILENAME = "c:\mercvb\help\help.doc"
    Set appWord = New Word.Application
    Set WRDdoc = appWord.Documents.Open(SFILENAME)
    appWord.Visible = True



End Sub

Private Sub mapot1_Click(Index As Integer)
    gApoMenu = True
    APOT1.SHOW

End Sub

Private Sub mapot2_Click(Index As Integer)

    gApoMenu = True

    '
    ' MDIForm1.'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", "NormalOffice2007.ini"
    ' MDIForm1.'SkinFramework1.ApplyWindow APOT2.hWnd
    ' MDIForm1.'SkinFramework1.ApplyOptions = MDIForm1.'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
    '
    APOT2.WindowState = 2
    APOT2.SHOW


    '  'SkinFramework1.LoadSkin "", ""

End Sub

Private Sub mapot31_Click(Index As Integer)
    apot3.SHOW

End Sub

Private Sub mapot32_Click(Index As Integer)
'apot32.SHOW

End Sub

Private Sub mapot6_Click(Index As Integer)

    gApoMenu = True
    APOT6.SHOW
End Sub


Private Sub mapot7_Click(Index As Integer)
    If gXEIRISTHS <= 2 Then
        apot7.Dir1.Visible = False
        apot7.Text2.Visible = False
        apot7.apo.Visible = False
        apot7.eos.Visible = False
        apot7.Label1.Visible = False
        apot7.Label2.Visible = False
    Else
        apot7.SHOW
    End If

End Sub

Private Sub mbohu1_Click(Index As Integer)

    gApoMenu = True
    BOHU1.SHOW


End Sub

Private Sub mbohu11_Click()


    gApoMenu = True
    bohu11.SHOW

End Sub

Private Sub mbohu12_Click()
    bohu12.SHOW

End Sub

Private Sub mBohu13_Click()
    BOHU13.SHOW

End Sub

Private Sub mbohu2_Click(Index As Integer)

    gApoMenu = True
    BOHU2.SHOW
End Sub

Private Sub MDIForm_DblClick()
    Dim a As Long, X
    If gXEIRISTHS = 9 Then
        f_supervisor_code = FINDPARAMETROI("MDIFORM1", "f_supervisor_code", "0", "Κωδικός")
        X = f_supervisor_code
        PARAMETROI.PARAM.Caption = "MDIFORM1"
        PARAMETROI.SHOW 1
        f_supervisor_code = FINDPARAMETROI("MDIFORM1", "f_supervisor_code", "0", "Κωδικός")


        If f_supervisor_code = X Then
        Else
            If IsNull(f_supervisor_code) Then f_supervisor_code = ""
            Gdb.Execute "UPDATE DIK SET CODE='" + f_supervisor_code + "' WHERE  LEFT(KOD,6)='070109' ", a
            If a > 0 Then
                MsgBox "ενημερώθηκε ο νέος κωδικός"
            End If
        End If
        '   If F_BACKGROUND > 0 Then
        '      On Error GoTo ERRORS
        '      Kill "C:\MERCVB\BACKGROUND.BMP"
        '      FileCopy "C:\MERCVB\BACKGROUND" + LTrim(Str(F_BACKGROUND)) + ".BMP", "C:\MERCVB\BACKGROUND.BMP"
        '      MILSEC 2000
        '      gPicture = "C:\MERCVB\BACKGROUND.BMP"
        '      MDIForm1.Picture = LoadPicture(gPicture)
        '
        '   End If

    End If

    Exit Sub

ERRORS:

    MsgBox "Δεν έγινε εγκατάσταση " + Chr(13) + Err.Description
    Resume Next





End Sub

Private Sub MDIForm_Load()
' Global Gdb As New ADODB.Connection
' UGridEdit.SHOW 1
' End



    Dim R As New ADODB.Recordset
    g_Stop = 0


Dim DUM




    ''πως ανοιγω ενα exel αρχείο
    ' Dim xl As New excel.Application
    ' Dim xx As Worksheet
    ' openXLS "c:\apog.xls", xx, xl
    ' Print xx.cells(1, 1)
    'Set xx = Nothing
    'xl.Quit
    'Set xl = Nothing






    gDIG_3 = 6
    gDIG_4 = 8



'DUM = ShellExecute(0&, "Print", "C:\PDF\PDFTEMP\*.PDF", "", "", vbNormalFocus)



    Dim wrkDefault, mydb, minima

    '   '   'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
    '    'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
    '    'SkinFramework1.ApplyWindow Me.hWnd
    '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
    'SkinFramework1.LoadSkin "", ""


    gPicture = "C:\MERCVB\background.BMP"
    MDIForm1.Picture = LoadPicture(gPicture)

    ' Dim dum: dum = printing_etik("354454534534", 2, "tem", 1.31, 1.34, 1.23, 9, "lagakis", 67)







    MDIForm1.SHOW
    On Error GoTo SEELINE






    minima = "ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΑΡΧΕΙΟ C:\MERCPATH.TXT"
10  Open "C:\MERCPATH.TXT" For Input As #1
20  Line Input #1, gDir
30  Line Input #1, gConnect
40  Close #1


    ' ok !!!!!  gConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=38983;Initial Catalog=MERCURY"

    '50 Close #1

    '
    'Dim dsoServer As DSO.Server
    '    Dim enuClassType As DSO.ClassTypes
    '
    '    ' Create instance of server and connect.
    '    ' "LocalHost" will default to the
    '    ' local Windows NT Server 4.0 where the
    '    ' Analysis server is installed.
    '    Set dsoServer = New DSO.Server
    '    dsoServer.Connect "LocalHost"
    '
    '
    '
    '
    '
    'Set oDatabase = dsoServer.Databases("Northwind", "stevenb")




50  MDIForm1.StatusBar1.Panels(1).Text = gDir
60  MDIForm1.StatusBar1.Panels(3).Text = "Mercury ver.11.12"














    'Exit Sub

    pelat.Enabled = False
    APOT.Enabled = False
    EPIT.Enabled = False
    par.Enabled = False
    bohu.Enabled = False



    ' gConnect = "Provider=SQLOLEDB.1;Password=38983;Persist Security Info=True;User ID=sa;Data Source=GEORGE;Initial Catalog=MERCURY;"
    frmLogin.SHOW 1





    Timer1.Enabled = True

    If gConnect = "end" Then
        Unload Me
        End
    End If


    F_BACKGROUND = Val(FINDPARAMETROI("MDIFORM1", "F_BACKGROUND", "1", "BACKGROUND ΑΡ.ΕΙΚΟΝΑΣ=1,2,3,.. ΟΧΙ=0"))






    pelat.Enabled = True
    APOT.Enabled = True
    EPIT.Enabled = True
    par.Enabled = True
    bohu.Enabled = True


    If F_BACKGROUND = 1 Then
        ' OK ΑΦΗΝΩ ΤΟ ΑΡΧΕΙΟ ΩΣ ΕΧΕΙ
    Else
        gPicture = ""
        MDIForm1.Picture = LoadPicture(gPicture)
    End If











70  ' Gdb.Open gConnect




    ' f_loadpar1 = Val(FindParametroi("MDIFORM1", "F_LOADPAR1", "1", "ΦΟΡΤΩΝΩ ΚΑΘΕ ΦΟΡΑ ΤΗΝ ΦΟΡΜΑ ΠΑΡΑΣΤΑΤΙΚΩΝ=1"))

80  f_telephone = Val(FINDPARAMETROI("MDIFORM1", "F_TELEPHONE", "0", "ΕΜΦΑΝΙΣΗ ΚΑΛΟΥΝΤΩΝ ΤΗΛΕΦΩΝΩΝ=1"))

90  F_TOOLBAR = Val(FINDPARAMETROI("MDIFORM1", "F_TOOLBAR", "550", "ΥΨΟΣ TOOLBAR"))
100 f_Ypenth = Val(FINDPARAMETROI("MDIFORM1", "F_YPENTH", "1", "ΥΠΕΝΘΥΜΙΣΗ=1 ΟΧΙ=0"))
    'f_supervisor_code = Val(FindParametroi("MDIFORM1", "f_supervisor_code", "1", "Κωδικός"))


110 tlbToolBar.height = F_TOOLBAR


    On Error GoTo stopthlefona

120 If f_telephone = 1 Then
        TELEFON.ConnectionString = gConnect
        TELEFON.RecordSource = "SELECT TOP 1 * FROM THLEFNOW"
        TELEFON.Refresh
130 End If



140 If F_TOOLBAR = 0 Then
        tlbToolBar.Visible = False
        StatusBar1.Visible = False
145 End If


    On Error Resume Next




    '150 Update_Structures

    If gXEIRISTHS >= 8 Then
160     MakeBackup ""
    End If

    On Error GoTo update_structure



    R.Open "select TOP 1 KATHGORIA from EID", Gdb, adOpenForwardOnly, adLockReadOnly
    On Error Resume Next
    R.Close

    ' Gdb.Execute "ALTER TABLE EID ADD KATHGORIA INT"








    'On Error GoTo UPDATE2_STRUCTURE
    Dim dum2
    '  DUM2 = ADD_FIELD("PARASTAT", "ID", "INT IDENTITY(1,1)")
    'R.Open "select TOP 1 ARITMISI,TAYTPEL,TAYTXPI from XREOPIS", Gdb, adOpenForwardOnly, adLockReadOnly
    'On Error Resume Next
    'R.Close











    MDIForm1.StatusBar1.Panels(2).Text = gConnect
    If gConnect = "dBase III;" Then
        On Error GoTo 0
        update_tables
    End If


    'frmLogin.SHOW 1




    If gXEIRISTHS < 8 Then
        'bohu.Visible = False
        mapot1(21).Visible = False
        pareid(23).Visible = False
        apo4(24).Visible = False
        apo5(25).Visible = False
        mapot6(26).Visible = False
        EPIT.Visible = False

        mBohu1(51).Visible = False
        mbohu2(52).Visible = False

        par3(53).Visible = False
        parkin(54).Visible = False
        pros(55).Visible = False
        genlog(56).Visible = False

        If gXEIRISTHS <= 2 Then



            mapot2(22).Visible = False
            apoemporiko(1).Visible = False    'True
            DBF(57).Visible = False
            ' pontoif(58).Visible = False
            apoemporiko(59).Visible = False
            calculator(510).Visible = False
            mpelat1.Visible = False
            mpelat2.Visible = False
            mpelat3.Visible = False
            mpelat4.Visible = False
            mpelat5.Visible = False
            APOT.Visible = False

            mapot7(27).Visible = False
            mpar7.Visible = False
            mBohu13.Visible = False

            diagrpar(42).Visible = False
            parenos(43).Visible = False
            mpar5.Visible = False


            If gXEIRISTHS = 2 Then
            Else
                mPar6.Visible = False
            End If


            mbohu12.Visible = False
            timspar(44).Visible = False

            mapot2(22).Visible = False


            mapot7(27).Visible = False





        End If


    Else
        If f_Ypenth = 1 Then utPROEID.SHOW



    End If






    ' gDir = "\\DCCSERVER\data\lageuro" '  "\\R1r5q6\d\lageuro"
    'gDir = "d:\lageuro3" '  "\\R1r5q6\d\lageuro"
    '  gConnect = "dBase IV;"

    'gDir = ektyp_forma("ss")

    '   gDir = "c:\MERCVB\EMP.mdb"
    '  gConnect = "" 'Access
    ' gDek_pos = "###,###.00"
    ' gDek_ajia = "###,###.00"


    'ektyp_forma ("c:\mercvb\f90")


    ' Get default Workspace.
45  Set wrkDefault = DBEngine.Workspaces(0)

    'MYDB = wrkDefault.OpenDatabase("c:\MERCVB\EMP.mdb")
    'If MYDB.CollatingOrder <> dbLangGreek Then
    '   wrkDefault.MYDB.Close
    '   wrkDefault.CompactDatabase "emp.mdb", "emp2.mdb", dbLangGreek
    'End If

    Me.Caption = "LoadParam"
    On Error GoTo 0

    On Error Resume Next
62  LOADPARAMETERS

570 Me.Caption = "Menu"

    On Error GoTo CREATEPARAMETROI

    R.Open "select * from PARAMETROI", Gdb, adOpenForwardOnly, adLockReadOnly

    R.Close

    Dim Sql2

    On Error GoTo CREATEXRHSEIS
    Sql2 = "select * from XRHSEIS where '" + Format(gWorkDay, "MM/DD/YYYY") + "'>=ENARXH AND '" + Format(gWorkDay, "MM/DD/YYYY") + "'<=LHXH"
    R.Open Sql2, Gdb, adOpenForwardOnly, adLockReadOnly


    If R.EOF Then
        'BAZΩ 1/1/2000 ΓΙΑ ΝΑ ΜΗΝ ΒΓΑΖΕΙ ΛΑΘΟΣ ΥΠΟΛΟΙΠΑ
        Sql2 = "INSERT INTO XRHSEIS (ETOS,ENARXH,LHXH) VALUES (" + LTrim(Str(Year(gWorkDay))) + ",'01/01/2000','12/31/" + LTrim(Str(Year(gWorkDay))) + "')"
        Gdb.Execute Sql2
        R.Close
        R.Open Sql2, Gdb, adOpenForwardOnly, adLockReadOnly
    End If


    gEnarjh = R(1)
    gLhjh = R(2)





    R.Close








    Exit Sub















    'UPDATE2_STRUCTURE:
    'On Error Resume Next
    '
    'On Error Resume Next
    '   Gdb.Execute "ALTER TABLE XREOPIS ADD ARITMISI INT "
    '   Gdb.Execute "ALTER TABLE XREOPIS ADD TAYTPEL  CHAR(15) "
    '   Gdb.Execute "ALTER TABLE XREOPIS ADD TAYTXPI  CHAR(1) "
    '
    '    Resume Next

stopthlefona:
    f_telephone = 0
    Timer1.Enabled = False

    Resume Next




update_structure:
    Gdb.Execute "ALTER TABLE EID ADD KATHGORIA INT"
    Gdb.Execute "UPDATE EID SET KATHGORIA=1"
    Resume Next




CREATEPARAMETROI:

    Gdb.Execute "CREATE TABLE PARAMETROI (FORMA CHAR(20),VAR CHAR(25) ,TIMH CHAR(10),SXOLIA CHAR(50)  )"
    Exit Sub



CREATEXRHSEIS:

    'Resume Next
    Gdb.Execute "CREATE TABLE XRHSEIS (ETOS SMALLINT,ENARXH DATETIME ,LHXH DATETIME,N1 SMALLINT,N2 SMALLINT,C1 CHAR(20),C2 CHAR(20),ID INT IDENTITY(1,1) )  "


    Exit Sub





SEELINE:
    HandleError "MdiForm-load"
    MsgBox Str(Erl), vbOKOnly, minima + " " + Err.Description

    Exit Sub    'Resume Next

End Sub

Private Sub MDIForm_Resize()

    On Error Resume Next

    If MDIForm1.width > 15225 Then
        '  MDIForm1.width = 15225  ' MDIForm1.Picture.width
    End If


    'If MDIForm1.WindowState = 2 Then MDIForm1.WindowState = 0
End Sub

Private Sub par1_Click()
    par1.SHOW
End Sub

Private Sub metasx_par_Click()

    par1.metasx.value = vbChecked


    par1.diortosis

End Sub

Private Sub mpar1_Click(Index As Integer)
    gApoMenu = True

    par1.SHOW
End Sub

Private Sub mpar5_Click()

    If gXEIRISTHS = 9 Then
    Else
        par5.SHOW
    End If
End Sub

Private Sub mPar6_Click()
    par6.SHOW

End Sub

Private Sub mpar7_Click()
    par7.SHOW

End Sub

Private Sub mpar8_Click()
' par8.SHOW

End Sub

Private Sub mpar9_Click()
' par9.SHOW
End Sub

'Private Sub mpar1_Click()

'End Sub

Private Sub mpelat1_Click()
10  pelat1.SHOW
End Sub

Private Sub mpelat2_Click()
    pelat2.SHOW
End Sub

Private Sub mpelat3_Click()
    pelat3.SHOW
End Sub

Private Sub mpelat4_Click()
    gApoMenu = True
    pelat4.resizing = 1
    pelat4.SHOW
End Sub


'Private Sub par1_Click()

'End Sub

Private Sub mpelat5_Click()
    gApoMenu = True
    pelat5.SHOW
End Sub

'Private Sub par1menu_Click()

'
'End Sub

Private Sub par2show_Click()
    par1.PAR2SHOWING




End Sub

Private Sub par3_Click(Index As Integer)
    gApoMenu = True
    bohu3.SHOW
End Sub



Private Sub pareid_Click(Index As Integer)
    gApoMenu = True
    '  apot3.SHOW

End Sub

Private Sub parenos_Click(Index As Integer)
    gApoMenu = True
    par2.SHOW

End Sub

Private Sub parepit_Click()
    gApoMenu = True
    '   epit3.Top = 1450
    '  epit3.Left = 90
    ' epit3.width = MDIForm1.width - 100
    ' epit3.height = MDIForm1.height - 1500

    epit3.SHOW
End Sub

Private Sub parkin_Click(Index As Integer)
    gApoMenu = True
    bohu4.SHOW
End Sub

Private Sub pontoi_Click(Index As Integer)
    gApoMenu = True
    bohu6.SHOW

End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    Dim a

    If KeyAscii = 27 Then
        a = 1

        a = MsgBox("Εξοδος ;", vbYesNo)
        If a = vbYes Then End


    End If
End Sub

Private Sub pontoif_Click(Index As Integer)
    bohu7.SHOW

End Sub

Private Sub pros_Click(Index As Integer)
    bohu5.SHOW


    'Dim db, DUM, aa As String, pol As String, polepis As String, ago As String, agoepis As String
    '
    ''Set db = OpenDatabase(gDir, False, False, gConnect)
    'Gdb.Execute UCase("update eid set m01=0,m02=0,m03=0,m04=0,m05=0,m06=0,m07=0,m08=0,m09=0,m10=0,m11=0,m12=0,g01=0,g02=0,g03=0,g04=0,g05=0,g06=0,g07=0,g08=0,g09=0,g10=0,g11=0,g12=0;")
    '
    ''db.Execute "CREATE INDEX eidko ON eid(kod);"
    'pol = "": polepis = "": ago = "": agoepis = ""
    'DUM = Get_AJ(pol, polepis, ago, agoepis)
    '
    '
    '' Kill gDir + "\eidko.ndx"
    ''Gdb.Execute "CREATE  INDEX  eidko  ON   eid(kod);"
    '
    '
    'Dim R As New ADODB.Recordset, EID As New ADODB.Recordset
    '
    '
    '
    '
    ''eid.Index = "eidko"
    '
    'R.Open UCase("select kode,sum(poso) as sposo ,left(atim,1) as satim,month(hme) as shme from eggtim group by kode,left(atim,1),month(hme);"), Gdb, adOpenDynamic, adLockOptimistic
    '
    '
    ''Set r = db.OpenRecordset("eggtim")
    'Do While Not R.EOF
    '
    '
    '
    '   EID.Open "SELECT *FROM EID WHERE KOD='" + R("KODE") + "'", Gdb, adOpenDynamic, adLockOptimistic
    '   'eid.Seek "=", r("kode")
    '
    '   If Not EID.EOF Then
    '      ' eid.Edit
    '      aa = Format(R("shme"), "00")
    '
    '     If InStr(pol, R("satim")) > 0 Then EID("m" + aa) = EID("m" + aa) + R("sposo")
    '     If InStr(polepis, R("satim")) > 0 Then EID("m" + aa) = EID("m" + aa) - R("sposo")
    '     If InStr(ago, R("satim")) > 0 Then EID("g" + aa) = EID("g" + aa) + R("sposo")
    '     If InStr(agoepis, R("satim")) > 0 Then EID("g" + aa) = EID("g" + aa) - R("sposo")
    '     EID.Update
    '   End If
    '   R.MoveNext
    '
    'Loop
    '
    '
    'db.Close
    '






End Sub

Private Sub RUNT_Click(Index As Integer)
    Dim k  ' INDEX=0 TO PROTO MENOY   INDEX=1 TO DEYTERO

    gDir = ETAIR(Index)


    MDIForm1.StatusBar1.Panels(1).Text = gDir





    If InStr(UCase(gDir), "DSN") > 0 Then
        gConnect = gDir
    Else
        gConnect = gConnect + ";DATABASE=" + gDir

    End If



    If Left(gDir, 1) = "D" Then
        gPicture = "C:\MERCVB\PALIA.BMP"
    Else
        gPicture = "C:\MERCVB\BACKGROUND.BMP"
    End If



    For k = 1 To mc
        Unload runt(k)
    Next

    'For K = 1 To 1000: DoEvents: Next
    'MDIForm1.Picture = LoadPicture()
    'MILSEC 1000

    MsgBox "Aλλαγή εταιρείας"
    MDIForm1.Picture = LoadPicture(gPicture)

    'For K = 1 To 1000: DoEvents: Next


End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
    Dim a

    Dim db As Database, R As New ADODB.Recordset, k As Integer

    On Error Resume Next



    R.Open "select * from ETAIREIES", Gdb, adOpenForwardOnly, adLockReadOnly

    'MDIForm1.Picture = LoadPicture()
    mc = 0
    runt(0).Caption = R(0)
    ETAIR(0) = R(0)

    Do While Not R.EOF
        mc = mc + 1
        Load runt(mc)
        runt(mc).Caption = R!TITLOS
        ETAIR(mc) = Trim(R(0))
        R.MoveNext
    Loop

    PopupMenu RUNTIME(0)





End Sub

'Private Sub TELEFON_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

'  TELEFON.Refresh
'
'  If IsNull(TELEFON.Recordset("EPO")) Then
'     MDIForm1.Caption = TELEFON.Recordset("THL") + " "
'  Else
'     MDIForm1.Caption = TELEFON.Recordset("THL") + " " + TELEFON.Recordset("EPO")
'  End If
'
'
'
'End Sub

Private Sub Timer1_Timer()

    On Error GoTo 555


    If f_telephone = 1 Then
        Dim M_EPO
        If IsNull(TELEFON.Recordset("EPO")) Then M_EPO = Left(TELEFON.Recordset("EPO"), 30) Else M_EPO = Space(30)

        TELEFON.Refresh


        'If IsNull(TELEFON.Recordset("EPO")) Or IsNull(TELEFON.Recordset("DIE")) Then
        '  MDIForm1.Caption = TELEFON.Recordset("THL") + " "
        'MDIForm1.Caption = Left(TELEFON.Recordset("THL") + " " + Space(40), 40) + Left(MDIForm1.Caption, 40)
        'Else

        If Left(MDIForm1.Caption, 10) = Left(TELEFON.Recordset("THL") + " " + TELEFON.Recordset("EPO"), 10) Then
        Else
            MDIForm1.StatusBar1.Panels(5).Text = Left(MDIForm1.Caption, 42)
            MDIForm1.Caption = Left(TELEFON.Recordset("THL"), 10) + "      " + TELEFON.Recordset("EPO") + "  προηγούμενο : " + Left(MDIForm1.Caption, 40)    '+ TELEFON.Recordset("DIE")
        End If

        'MDIForm1.Caption = Left(TELEFON.Recordset("THL") + " " + TELEFON.Recordset("EPO") + Space(40), 40) + Left(MDIForm1.Caption, 40)

        'End If

        Exit Sub

    End If



    Dim R As New ADODB.Recordset
    R.Open "SELECT * FROM CalendarEvents where ABS(DATEDIFF(mi,StartDateTime,Getdate()))< ReminderMinutesBeforeStart AND IsReminder=1 ", Gdb, adOpenForwardOnly, adLockReadOnly
    '   Exit Sub
    If R.EOF Then
        R.Close
        Exit Sub
    End If



    Dim DELT As String
    DELT = "---- Yπενθύμιση --- " + Chr(13)



    ' Exit Sub
    Do While Not R.EOF
        ' + "Προθεσμία min:" + Format(r("DIFF"), "#####") + "/"
        DELT = DELT + R("Subject") + "/" + R("Location") + "/" + R("Body") + "/" + Format(R("StartDateTime"), "dd/mm/yyyy hh:mm") + Chr(13)
        R.MoveNext
    Loop

    MsgBox DELT
    R.Close


555:



End Sub

'Private Sub timspar_Click()

'End Sub
Private Sub timspar_Click(Index As Integer)
    par4.SHOW
End Sub

Sub CREATEDATABASE(db)

    Dim DefaultWorkspace As Workspace
    Dim CurrentDatabase As Database, MyDatabase As Database
    Set DefaultWorkspace = DBEngine.Workspaces(0)
    'Set CurrentDatabase = DefaultWorkspace.Databases(0)
    ' Create new, encrypted database.
    Set MyDatabase = DefaultWorkspace.CREATEDATABASE(db, dbLangGreek)
    ' Fill in new database.
    ' Set property to that of current database.
    'MyDatabase.QueryTimeout = CurrentDatabase.QueryTimeout
    '...
    ' Complete structure of MyDatabase, possibly using structure of
    ' CurrentDatabase.
    '...
    MyDatabase.Close

End Sub


Sub MakeBackup(THESI As String)
'-------------------------------------------------------------

    Dim k As Integer
    Dim R As New ADODB.Recordset
    Dim mkod As String
    Dim DUM
    Dim DD
    Dim M_BACKUPTYPE
    M_BACKUPTYPE = FINDPARAMETROI("BOHU8", "F_BACKUPTYPE", 0, "ΤΥΠΟΣ BACKUP 0=ACCESS 1=SQL DATABASE")    'posa psifia tha exei h kathe seira

    Dim M_PATH

    M_PATH = Trim(FINDPARAMETROI("BOHU8", "F_PATH", "C:\MERCVB", "ΔΙΑΔΡΟΜΗ ΑΝΤΙΓΡΑΦΩΝ"))    'posa psifia tha exei h kathe seira)

    If Len(THESI) > 0 Then
        M_PATH = THESI
    End If


    Dim DDTOP



    If M_BACKUPTYPE = 1 Then
        DD = M_PATH + "\BACK" + Format(Day(Now), "00") + ".BAK"
        If Len(Dir(DD, vbNormal)) > 0 Then
            If Format(FileDateTime(DD), "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then
                Exit Sub
            Else
                Kill DD
                'yparxei_palio = True
            End If
        End If
        Gdb.Execute "BACKUP DATABASE [" + Gdb.DefaultDatabase + "] TO  DISK ='" + DD + "' WITH NOFORMAT, NOINIT, SKIP, NOREWIND, NOUNLOAD,  STATS = 10 ", k
        Exit Sub
    End If








    DD = M_PATH + "\BACK" + Format(Day(Now), "00") + ".MDB"
    DDTOP = "C:\MERCVB\BACK" + Format(Day(Now), "00") + ".MDB"

    Dim OK As Boolean
    OK = False

    'Dim yparxei_palio As Boolean
    'yparxei_palio = False
    If Len(Dir(DD, vbNormal)) > 0 Then
        If Format(FileDateTime(DD), "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then
            Exit Sub
        Else
            Kill DD
            'yparxei_palio = True
        End If

    End If
    Dim LISTA(20)
    LISTA(1) = "PEL"
    LISTA(2) = "EID"
    LISTA(3) = "EGG"
    LISTA(4) = "EGGTIM"
    LISTA(5) = "TIM"
    LISTA(6) = "GRA"
    LISTA(7) = "MEM"
    LISTA(8) = "PINAKES"
    LISTA(9) = "XREOPIS"
    LISTA(10) = "ARITMISI"
    LISTA(11) = "PARASTAT"
    LISTA(12) = "BARCODES"
    LISTA(13) = "HMEROL"
    LISTA(14) = "HMEROL2"
    LISTA(15) = "PARAMETROI"
    LISTA(16) = "CalendarEvents"
    LISTA(17) = "CalendarRecurrencePatterns"

    On Error Resume Next
    'Kill DD

    On Error GoTo LATOS2
    CREATEDATABASE DD
    On Error Resume Next

    Dim db As Database
    Set db = OpenDatabase("", False, False, gConnect)
    Me.MousePointer = vbHourglass
    Dim LATOS

    Me.Caption = "παρακαλώ περιμένετε - διαδικασία αντιγράφων"

    For k = 1 To 17    ' List1.ListCount - 1
        LATOS = 0
        On Error GoTo LATOS
        DoEvents
        Me.Caption = "παρακαλώ περιμένετε - διαδικασία αντιγράφων " + LISTA(k) + " " + LTrim(Str(k)) + "/17)"
        db.Execute "SELECT * INTO " + LISTA(k) + "  IN '" + DD + "'  FROM  " + LISTA(k)
    Next








    Me.MousePointer = vbNormal


    If M_PATH = "C:\MERCVB" Then
        ' OK
    Else
        ' FileCopy DD, DDTOP

    End If



    Exit Sub

LATOS:
    'List2.AddItem Err.Description
    LATOS = 1
    Resume Next
    Return

LATOS2:
    MsgBox "ΛΑΘΟΣ ΟΝΟΜΑ ΒΑΣΗΣ ACCESS  ΔΩΣΤΕ Π.Χ.  C:\BACK.MDB"


End Sub

Private Sub tlbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim a
    a = 1
    Dim DUM

    Select Case Button.Key
    Case "Dynaset"
        HMEROL2.SHOW
    Case "Table"
        hmerol.SHOW
    Case "BIBLIO"
        paralabh.SHOW
    Case "Schedule"
        DUM = Shell("C:\MERCVB\CALENDAR.EXE", vbMaximizedFocus)
    End Select

    tlbToolBar.buttons(1).style = tbrDefault
    tlbToolBar.buttons(2).style = tbrDefault






End Sub
