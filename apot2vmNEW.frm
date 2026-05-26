VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form apot2vmn 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   15825
   ClientTop       =   3285
   ClientWidth     =   15150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc EGGTIM2 
      Height          =   495
      Left            =   15000
      Top             =   9960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
         Name            =   "Tahoma"
         Size            =   7.5
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
      Height          =   372
      Left            =   7200
      TabIndex        =   54
      Top             =   3720
      Visible         =   0   'False
      Width           =   1812
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"apot2vmNEW.frx":0000
   End
   Begin MSAdodcLib.Adodc EID 
      Height          =   492
      Left            =   14760
      Top             =   2400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
         Size            =   7.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   14715
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   14775
      Begin VB.CommandButton cmdDELETE 
         Caption         =   "ΔΙΑΓΡΑΦΗ"
         Height          =   450
         Left            =   8100
         TabIndex        =   71
         Top             =   0
         Width           =   900
      End
      Begin VB.ListBox List1 
         Height          =   240
         Left            =   2640
         TabIndex        =   70
         Top             =   0
         Width           =   2655
      End
      Begin VB.CommandButton CMDmovenext 
         Caption         =   ">>"
         Height          =   250
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton CMDmoveprevious 
         Caption         =   "<<"
         Height          =   250
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   200
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Αναζήτηση με αρ.τιμολογίου"
         Height          =   180
         Left            =   10440
         TabIndex        =   66
         Top             =   0
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   900
         TabIndex        =   64
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton neo 
         BackColor       =   &H008080FF&
         Caption         =   "Nέο Είδος"
         Height          =   225
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton upload1 
         Caption         =   "Upload"
         Height          =   465
         Left            =   6720
         TabIndex        =   84
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdΕΝΗΜΕΡΩΣΗ 
         Caption         =   "ΕΝΗΜΕΡΩΣΗ"
         Height          =   465
         Left            =   5400
         TabIndex        =   13
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdANAZHTHSH 
         Caption         =   "ΑΚΥΡΩΣΗ"
         Height          =   225
         Left            =   9000
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdΕξοδος 
         Caption         =   "Εξοδος"
         Height          =   240
         Left            =   13440
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox ONO 
         Height          =   270
         Left            =   900
         TabIndex        =   6
         Top             =   490
         Width           =   1575
      End
      Begin VB.TextBox KOD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   900
         TabIndex        =   69
         Top             =   0
         Width           =   1605
      End
      Begin VB.Label mERG 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Left            =   12380
         TabIndex        =   72
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label7 
         Caption         =   "Barcode"
         Height          =   255
         Left            =   0
         TabIndex        =   65
         Top             =   240
         Width           =   735
      End
      Begin VB.Label NEA 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "ΝΕΑ ΕΓΓΡΑΦΗ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13080
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label kodlabel 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         DataField       =   "KOD"
         DataSource      =   "EID"
         Height          =   300
         Left            =   5400
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label ONOLABEL 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         DataField       =   "ONO"
         DataSource      =   "EID"
         Height          =   300
         Left            =   6720
         TabIndex        =   14
         Top             =   480
         Width           =   5600
      End
      Begin VB.Label lblΠεριγραφή 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Περιγραφή"
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   780
      End
      Begin VB.Label lblΚωδικός 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδικός"
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7812
      Left            =   15180
      TabIndex        =   1
      Top             =   -60
      Width           =   252
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   13935
      Left            =   0
      ScaleHeight     =   13875
      ScaleWidth      =   14835
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   825
      Width           =   14895
      Begin VB.CommandButton Command6 
         Caption         =   "ΕΝΗΜΕΡΩΣΗ SITE ΜΕ ΝΕΟΥΣ ΚΩΔΙΚΟΥΣ"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6600
         TabIndex        =   97
         Top             =   9120
         Width           =   2655
      End
      Begin VB.TextBox num3 
         BackColor       =   &H000080FF&
         DataField       =   "NUM3"
         DataSource      =   "EID"
         Height          =   270
         Left            =   11760
         TabIndex        =   95
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox Text5 
         DataField       =   "PROM"
         DataSource      =   "EID"
         Height          =   288
         Left            =   5625
         TabIndex        =   93
         Top             =   5550
         Width           =   1755
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   375
         Left            =   14760
         TabIndex        =   92
         Top             =   6480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox HAS_ASSORTMENT 
         Alignment       =   1  'Right Justify
         Caption         =   "HAS_ASSORTMENT"
         Height          =   255
         Left            =   13560
         TabIndex        =   91
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton oristiki 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Οριστικοποίηση"
         Enabled         =   0   'False
         Height          =   300
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CheckBox xarakt 
         Caption         =   "συμμετέχει στην περιγραφή ο χαρακτήρας"
         Height          =   210
         Left            =   1680
         TabIndex        =   89
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox Text4 
         DataField       =   "POS03"
         DataSource      =   "EID"
         Height          =   270
         Left            =   8280
         TabIndex        =   86
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text3 
         DataField       =   "POS01"
         DataSource      =   "EID"
         Height          =   270
         Left            =   6480
         TabIndex        =   85
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton download_eidh 
         Caption         =   "Downoad ειδη"
         Height          =   495
         Left            =   13440
         TabIndex        =   83
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CheckBox KATASK 
         Caption         =   "συμμετέχει στην περιγραφή ο κατασκευαστής"
         Height          =   210
         Left            =   1680
         TabIndex        =   82
         Top             =   405
         Width           =   3735
      End
      Begin VB.TextBox d1 
         Height          =   270
         Left            =   13560
         TabIndex        =   80
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox c1 
         Height          =   270
         Left            =   13560
         TabIndex        =   79
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox b1 
         Height          =   270
         Left            =   13560
         TabIndex        =   78
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   456
         Left            =   13440
         TabIndex        =   77
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Eνημέρωση site από εμπορικό"
         Height          =   615
         Left            =   13680
         TabIndex        =   76
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "preview"
         Height          =   360
         Left            =   13680
         TabIndex        =   75
         Top             =   5640
         Width           =   990
      End
      Begin VB.ComboBox ftp 
         Height          =   300
         ItemData        =   "apot2vmNEW.frx":0086
         Left            =   13470
         List            =   "apot2vmNEW.frx":0090
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   1185
         Width           =   1350
      End
      Begin VB.CheckBox reload 
         Caption         =   "Επαναφόρτωση"
         Height          =   300
         Left            =   13455
         TabIndex        =   73
         Top             =   585
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   10800
         TabIndex        =   63
         Top             =   8040
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "DOWNLOAD TABLE PRODUCTS"
         Height          =   375
         Left            =   13200
         TabIndex        =   62
         Top             =   8040
         Width           =   1575
      End
      Begin VB.TextBox WEBPRICE 
         DataField       =   "WEBPRICE"
         DataSource      =   "EID"
         Height          =   288
         Left            =   8760
         TabIndex        =   20
         Top             =   120
         Width           =   1332
      End
      Begin VB.TextBox YOU_TUBE 
         DataField       =   "YOU_TUBE"
         DataSource      =   "EID"
         Height          =   288
         Left            =   7200
         TabIndex        =   31
         Top             =   5880
         Width           =   1692
      End
      Begin VB.TextBox ZIP_CODE 
         DataField       =   "TEXT_PRODUCTS_ZIPCODE"
         DataSource      =   "EID"
         Height          =   288
         Left            =   2280
         TabIndex        =   27
         Top             =   5160
         Width           =   1572
      End
      Begin VB.TextBox INDV_PRICE 
         DataField       =   "SHIPPINGPRICE"
         DataSource      =   "EID"
         Height          =   288
         Left            =   2280
         TabIndex        =   28
         Top             =   5520
         Width           =   1572
      End
      Begin VB.TextBox ADDIT_PRICE 
         DataField       =   "ADDITIONALPRICE"
         DataSource      =   "EID"
         Height          =   288
         Left            =   2280
         TabIndex        =   29
         Top             =   5880
         Width           =   1572
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "apot2vmNEW.frx":00A3
         Left            =   2280
         List            =   "apot2vmNEW.frx":00A5
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   -240
         MultiLine       =   -1  'True
         TabIndex        =   53
         Text            =   "apot2vmNEW.frx":00A7
         Top             =   9240
         Width           =   1572
      End
      Begin VB.ListBox hlikies 
         Height          =   780
         Left            =   5640
         MultiSelect     =   2  'Extended
         TabIndex        =   30
         Top             =   3840
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "Διαθέσιμο"
         Height          =   255
         Left            =   12120
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox kataskeyasths 
         Height          =   300
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   4080
         Width           =   3255
      End
      Begin VB.ComboBox xaraktiras 
         Height          =   300
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   4800
         Width           =   3255
      End
      Begin VB.ComboBox fylo 
         Height          =   300
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   4440
         Width           =   3255
      End
      Begin VB.TextBox WEIGHT 
         DataField       =   "WEIGHT"
         DataSource      =   "EID"
         Height          =   270
         Left            =   10920
         TabIndex        =   21
         Top             =   120
         Width           =   1095
      End
      Begin MSComctlLib.TreeView T1 
         Height          =   5412
         Left            =   9000
         TabIndex        =   46
         Top             =   600
         Width           =   4452
         _ExtentX        =   7858
         _ExtentY        =   9551
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.PictureBox foto2 
         Height          =   1575
         Index           =   6
         Left            =   12840
         ScaleHeight     =   1515
         ScaleWidth      =   1875
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1935
      End
      Begin VB.PictureBox foto2 
         Height          =   1575
         Index           =   5
         Left            =   10800
         ScaleHeight     =   1515
         ScaleWidth      =   1875
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1935
      End
      Begin VB.PictureBox foto2 
         Height          =   1575
         Index           =   4
         Left            =   9000
         ScaleHeight     =   1515
         ScaleWidth      =   1875
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1935
      End
      Begin VB.PictureBox foto2 
         Height          =   1575
         Index           =   3
         Left            =   6960
         ScaleHeight     =   1515
         ScaleWidth      =   1875
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1935
      End
      Begin VB.PictureBox foto2 
         Height          =   1575
         Index           =   2
         Left            =   4920
         ScaleHeight     =   1515
         ScaleWidth      =   1875
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1935
      End
      Begin VB.PictureBox foto2 
         Height          =   1575
         Index           =   1
         Left            =   2880
         ScaleHeight     =   1515
         ScaleWidth      =   1875
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   6345
         Width           =   1935
      End
      Begin VB.TextBox LTI5 
         DataField       =   "LTI5"
         DataSource      =   "EID"
         Height          =   288
         Left            =   6480
         TabIndex        =   19
         Top             =   120
         Width           =   972
      End
      Begin VB.TextBox FONO 
         DataField       =   "ono"
         DataSource      =   "EID"
         Height          =   288
         Left            =   1560
         TabIndex        =   18
         Top             =   120
         Width           =   3612
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   11160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox foto2 
         Height          =   2052
         Index           =   0
         Left            =   0
         ScaleHeight     =   1995
         ScaleWidth      =   2715
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   6360
         Width           =   2772
      End
      Begin ADOMERCNEW.msEditor2 msEditor21 
         Height          =   2655
         Left            =   120
         TabIndex        =   7
         Top             =   855
         Width           =   8775
         _extentx        =   15478
         _extenty        =   4683
         backcolor       =   12648447
         font            =   "apot2vmNEW.frx":00AD
         backcolor       =   12648447
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   360
         Left            =   0
         TabIndex        =   2
         Top             =   13200
         Width           =   990
      End
      Begin ADOMERCNEW.msEditor2 msEditor22 
         Height          =   975
         Left            =   5640
         TabIndex        =   8
         Top             =   13680
         Visible         =   0   'False
         Width           =   2055
         _extentx        =   3625
         _extenty        =   1720
         backcolor       =   12648447
         font            =   "apot2vmNEW.frx":00D5
         backcolor       =   12648447
      End
      Begin VB.Label Label10 
         Caption         =   "Ενημέρωση Site=1   Οχι=2"
         Height          =   255
         Left            =   9000
         TabIndex        =   96
         Top             =   6120
         Width           =   2655
      End
      Begin VB.Label Προμηθευτής 
         BackStyle       =   0  'Transparent
         Caption         =   "Προμηθευτής"
         Height          =   225
         Left            =   4260
         TabIndex        =   94
         Top             =   5580
         Width           =   1350
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ξάνθη υπολ"
         Height          =   255
         Left            =   7320
         TabIndex        =   88
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Δράμα υπολ"
         Height          =   255
         Left            =   5520
         TabIndex        =   87
         Top             =   480
         Width           =   855
      End
      Begin VB.Label e1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   13560
         TabIndex        =   81
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label lblΤιμήWEB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Τιμή WEB"
         Height          =   192
         Left            =   7800
         TabIndex        =   61
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblD_Video_Youtube 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID_Video_Youtube"
         Height          =   195
         Left            =   5760
         TabIndex        =   60
         Top             =   6000
         Width           =   1245
      End
      Begin VB.Label lblEachAdditional 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Each Additional Price"
         Height          =   195
         Left            =   360
         TabIndex        =   59
         Top             =   6000
         Width           =   1425
      End
      Begin VB.Label lblIndvShipping 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Indv. Shipping Price:"
         Height          =   195
         Left            =   360
         TabIndex        =   58
         Top             =   5640
         Width           =   1410
      End
      Begin VB.Label lblZIP_CODE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZIP_CODE"
         Height          =   195
         Left            =   600
         TabIndex        =   57
         Top             =   5280
         Width           =   1050
      End
      Begin VB.Label lblΚατηγΦΠΑ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κατηγ.ΦΠΑ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   55
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label cΗλικίες 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Ηλικίες"
         Height          =   330
         Left            =   5640
         TabIndex        =   51
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Χαρακτήρας/Θέμα"
         Height          =   180
         Left            =   480
         TabIndex        =   50
         Top             =   4920
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αγόρι/Κορίτσι"
         Height          =   180
         Left            =   480
         TabIndex        =   49
         Top             =   4560
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κατασκευαστής"
         Height          =   180
         Left            =   480
         TabIndex        =   48
         Top             =   4080
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Βάρος"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   10200
         TabIndex        =   47
         Top             =   120
         Width           =   528
      End
      Begin VB.Label PIC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         DataField       =   "PIC6"
         DataSource      =   "EID"
         Height          =   180
         Index           =   6
         Left            =   10680
         TabIndex        =   45
         Top             =   10305
         Width           =   45
      End
      Begin VB.Label PIC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         DataField       =   "PIC5"
         DataSource      =   "EID"
         Height          =   180
         Index           =   5
         Left            =   8520
         TabIndex        =   44
         Top             =   10305
         Width           =   45
      End
      Begin VB.Label PIC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         DataField       =   "PIC4"
         DataSource      =   "EID"
         Height          =   180
         Index           =   4
         Left            =   6600
         TabIndex        =   43
         Top             =   10305
         Width           =   45
      End
      Begin VB.Label PIC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         DataField       =   "PIC3"
         DataSource      =   "EID"
         Height          =   180
         Index           =   3
         Left            =   4440
         TabIndex        =   42
         Top             =   10305
         Width           =   45
      End
      Begin VB.Label PIC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         DataField       =   "PIC2"
         DataSource      =   "EID"
         Height          =   180
         Index           =   2
         Left            =   2400
         TabIndex        =   41
         Top             =   10305
         Width           =   45
      End
      Begin VB.Label PIC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         DataField       =   "PIC1"
         DataSource      =   "EID"
         Height          =   240
         Index           =   1
         Left            =   3555
         TabIndex        =   40
         Top             =   7905
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Τιμή καταστηματος"
         Height          =   384
         Left            =   5400
         TabIndex        =   33
         Top             =   120
         Width           =   960
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Σύντομη περιγραφή"
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label PIC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         DataField       =   "PIC0"
         DataSource      =   "EID"
         Height          =   240
         Index           =   0
         Left            =   2820
         TabIndex        =   17
         Top             =   8175
         Width           =   1365
      End
      Begin VB.Label lblΑναλυτικήΠεριγραφή 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αναλυτική Περιγραφή (Ξενόγλωσση)"
         Height          =   180
         Index           =   1
         Left            =   5160
         TabIndex        =   10
         Top             =   13440
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label lblΑναλυτικήΠεριγραφή 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αναλυτική Περιγραφή"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   615
         Width           =   1560
      End
   End
End
Attribute VB_Name = "apot2vmn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================   ftp  ================================================================

Private WithEvents mFTP As cFTP
Attribute mFTP.VB_VarHelpID = -1

Private BeginTransfer   As Single

Private TransferRate    As Single

Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

Private FilePathName As String

Private FILENAME     As String

Private FormName     As String



Dim fmaxRec As Long

Private Declare Function OSGetPrivateProfileInt _
                Lib "kernel32" _
                Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
                                               ByVal lpKeyName As String, _
                                               ByVal nDefault As Long, _
                                               ByVal lpFileName As String) As Long

Private Declare Function OSGetPrivateProfileSection _
                Lib "kernel32" _
                Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, _
                                                   ByVal lpReturnedString As String, _
                                                   ByVal nSize As Long, _
                                                   ByVal lpFileName As String) As Long

Private Declare Function OSGetPrivateProfileString _
                Lib "kernel32" _
                Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                  ByVal lpKeyName As Any, _
                                                  ByVal lpDefault As String, _
                                                  ByVal lpReturnedString As String, _
                                                  ByVal nSize As Long, _
                                                  ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileSection _
                Lib "kernel32" _
                Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, _
                                                     ByVal lpString As String, _
                                                     ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileString _
                Lib "kernel32" _
                Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                    ByVal lpKeyName As Any, _
                                                    ByVal lpString As Any, _
                                                    ByVal lpFileName As String) As Long

Private Declare Function OSGetProfileInt _
                Lib "kernel32" _
                Alias "GetProfileIntA" (ByVal lpAppName As String, _
                                        ByVal lpKeyName As String, _
                                        ByVal nDefault As Long) As Long

Private Declare Function OSGetProfileSection _
                Lib "kernel32" _
                Alias "GetProfileSectionA" (ByVal lpAppName As String, _
                                            ByVal lpReturnedString As String, _
                                            ByVal nSize As Long) As Long

Private Declare Function OSGetProfileString _
                Lib "kernel32" _
                Alias "GetProfileStringA" (ByVal lpAppName As String, _
                                           ByVal lpKeyName As String, _
                                           ByVal lpDefault As String, _
                                           ByVal lpReturnedString As String, _
                                           ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection _
                Lib "kernel32" _
                Alias "WriteProfileSectionA" (ByVal lpAppName As String, _
                                              ByVal lpString As String) As Long

Private Declare Function OSWriteProfileString _
                Lib "kernel32" _
                Alias "WriteProfileStringA" (ByVal lpszSection As String, _
                                             ByVal lpszKeyName As String, _
                                             ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024

Private Const nBUFSIZEINIALL = 4096

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
Dim F_id_LOCAL  As Long

Dim gdblag      As New ADODB.Connection

Dim names(1000) As String

Dim f_exit      As Boolean

Dim f_pel       As String

Dim f_xaraktiras(500)

Dim f_fylo(10)

Dim f_kataskeyasths(500)

Dim f_hlikies(30)

Dim F_FAKEL_KOINOS      As String


Dim f_Limit_Posothta As Integer
        
Dim f_Limit_Timh As Integer
        






Dim f_site              As Integer

Dim f_arxikh_eikona(10) As String

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

'αποθηκευει τα στοιχεια των κατηγοριων
Dim fc_max          As Long

Dim FC_id(1000)     As Long

Dim FC_parenT(1000) As Long

Dim fc_names(1000)  As String

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

Private Sub BARCODE_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then

        KeyAscii = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If
 
End Sub

Private Sub BARCODE_LostFocus()

    If Len(barcode) > 0 Then

        If Len(EID.RecordSource) > 5 Then
            If Not EID.Recordset.EOF Then
                EID.Recordset.Move 0
            End If
        End If

        EID.RecordSource = "SELECT * FROM EID WHERE ERG='" + barcode.Text + "')"
        EID.Refresh
   
        UPDATE_CONTROLS
   
    End If

End Sub

Sub UPDATE_CONTROLS()
   
        For k = 0 To hlikies.ListCount - 1
            'If hlikies.Selected(k) Then
            hlikies.Selected(k) = False
        Next
   
        On Error Resume Next

        ONOLABEL.Caption = EID.Recordset("ono")
        WEBPRICE.Text = ""

        If IsNull(EID.Recordset("WEBPRICE")) Then
            WEBPRICE.Text = EID.Recordset("lti5")
        Else
            WEBPRICE.Text = EID.Recordset("WEBPRICE")
        End If

        mERG.Caption = EID.Recordset("ERG")
    
    If IsNull(EID.Recordset("HAS_ASSORTMENT")) Then
        HAS_ASSORTMENT.Value = vbUnchecked
    Else
       If EID.Recordset("HAS_ASSORTMENT") = 0 Then
          HAS_ASSORTMENT.Value = vbUnchecked
       Else
          HAS_ASSORTMENT.Value = vbChecked
       End If
    End If
    
        kodlabel.Caption = EID.Recordset("kod")
        Set msEditor21.mDataSource = EID  ' for code you hav mum length.

        msEditor21.mDataField = "MEMO" '  "nNotes"     ' The field of your text
        msEditor21.Enabled = True

        Set msEditor22.mDataSource = EID  ' for code you have to use the 'adoNotes' instead of the Adodc1
        msEditor22.MaxLength = 8000          ' Optional maximum length.

        msEditor22.mDataField = "MEMOENG" '  "nNotes"     ' The field of your text
        msEditor22.Enabled = True
   
        Dim ss, N As Integer, kor As String, b As String, manuf As String
 
        'fylo.ListIndex=
        'manuf = str(f_kataskeyasths(kataskeyasths.ListIndex))
        'B = str(f_xaraktiras(xaraktiras.ListIndex))

        'Gdb.Execute "UPDATE EID SET XARAKTIRAS=" + B + ",MANUFACTURERS_ID=" + manuf + ",FYLO=" + kor + " WHERE KOD='" + kod + "'"
   
        For N = 0 To 6
            f_arxikh_eikona(N) = PIC(N).Caption

            If Len(Dir(F_FAKEL_KOINOS + "\images\" + PIC(N).Caption)) < 2 Then
                PIC(N).Caption = ""
            End If
    
            If Len(PIC(N).Caption) > 0 Then
                foto2(N).Picture = LoadPicture()
                foto2(N).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + PIC(N).Caption)
                StretchSourcePictureFromPicture apot2vmn.foto2(N).Picture, apot2vmn.foto2(N)
            Else
                foto2(N).Picture = LoadPicture()
            End If

        Next
   
        Combo2.ListIndex = EID.Recordset("FPA") - 1
        kataskeyasths.ListIndex = EID.Recordset("MANUFACTURERS_ID")
   
        ' fylo.ListIndex = EID.Recordset("FYLO")

        For k = 0 To xaraktiras.ListCount - 1

            If f_xaraktiras(k) = EID.Recordset("XARAKTIRAS") Then
                xaraktiras.ListIndex = k

                Exit For

            End If

        Next

        For k = 0 To kataskeyasths.ListCount - 1

            If f_kataskeyasths(k) = EID.Recordset("MANUFACTURERS_ID") Then
                kataskeyasths.ListIndex = k

                Exit For

            End If

        Next

        For k = 0 To fylo.ListCount - 1

            If f_fylo(k) = EID.Recordset("FYLO") Then
                fylo.ListIndex = k

                Exit For

            End If

        Next

        '    Gdb.Execute "CREATE TABLE PRODUCT_ATTRIBUTES ( KOD VARCHAR(16) NOT NULL,OPTIONS_VALUES_ID INT,PRODUCT_ID int NOT NULL,OPTIONS_ID  INT  NULL ) "
        '    End If
        ' r.Close
        '
        '
        Dim R As New ADODB.Recordset

        ' r.Open "SELECT * FROM PRODUCT_TO_CATEGORIES WHERE KOD='" + KOD + "'", Gdb, adOpenDynamic, adLockOptimistic
        '
        ' Do While Not r.EOF
        '
        '   r.MoveNext
        ' Loop
        '
        ' r.Close
        '
        R.Open "SELECT * FROM PRODUCT_ATTRIBUTES WHERE KOD='" + KOD + "'", Gdb, adOpenDynamic, adLockOptimistic

        Do While Not R.EOF

            For k = 0 To hlikies.ListCount - 1

                If f_hlikies(k) = R("OPTIONS_VALUES_ID") Then
                    hlikies.Selected(k) = True

                    Exit For

                End If

            Next

            R.MoveNext
        Loop

        R.Close

        'Loop
 
        ' R.Close
  
        If EID.Recordset("diathesimo") = 1 Then
            Check1.Value = vbChecked
        Else
            Check1.Value = vbUnchecked
        End If

        Combo2.ListIndex = EID.Recordset("fpa") - 1

        T1.Nodes.Clear

        For N = 1 To fc_max
196         error = 0

            'paterades
            If FC_parenT(N) = 0 Then
202             T1.Nodes.ADD , , "r" + LTrim(str(FC_id(N))), fc_names(N)
            Else

204             T1.Nodes.ADD "r" + LTrim(str(FC_parenT(N))), tvwChild, "r" + LTrim(str(FC_id(N))), (fc_names(N))
            End If

        Next

        R.Open "SELECT * FROM PRODUCT_TO_CATEGORIES WHERE KOD='" + KOD + "'", Gdb, adOpenDynamic, adLockOptimistic

        Do While Not R.EOF

            ' T1.Nodes.Item(20).Expanded = True
            ' T1.Nodes.Item(7).Selected = True
            For k = 1 To UBound(names) - 1

                If Val(mID(names(k), 2, 4)) = R!categories_id Then
                    T1.Nodes.Item(k).Selected = True
                    T1.Nodes.Item(k).BackColor = vbCyan
                    T1.SetFocus
                    ' T1.ShowWhatsThis
        
                    Exit For

                End If

            Next

            Exit Do  ' GIA MIA KATHGORIA

            R.MoveNext
        Loop

        R.Close

        Dim surl As String

        Dim m_ID As String

        If f_site = 111 Then

            'gdblag.Close
            'gdblag.Open f_connect  '"DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131 ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
            If Not R.EOF Then
                R.Open "select * from jsnrw_virtuemart_products where product_sku='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
                m_ID = str(R("virtuemart_product_id"))
                surl = "http://www.lagakis.gr/administrator/index.php?option=com_virtuemart&view=product&task=edit&virtuemart_product_id=" + m_ID + "&product_parent_id=0"
                WebBrowser1.navigate surl
            End If
        End If

        If f_site = 222 Then
            R.Open "select * from products where products_model='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

            If Not R.EOF Then
                m_ID = str(R("products_id"))
                surl = "http://www.toys-shop.gr/admin/categories.php?cPath=103_326&pID=" + m_ID + "&action=new_product"
                WebBrowser1.navigate surl
                'While Not WebBrowser1.readyState = READYSTATE_COMPLETE
                '         DoEvents
                'Wend
            End If
  
        End If
 
        '    Dim n7 As Long   10+18+15=43
        '     n7 = 0
        ' r.Open sql, gdblag, adOpenDynamic, adLockOptimistic
        '     Do While Not r.EOF
        '        xaraktiras.AddItem r!epf_value
        '        f_xaraktiras(n7) = r!Value_ID
        '        r.MoveNext
        '        n7 = n7 + 1
        '     Loop
        '     r.Close
        '     xaraktiras.ListIndex = 0

End Sub

Private Sub ADDIT_PRICE_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub b1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then c1.SetFocus
   
End Sub

Private Sub c1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then d1.SetFocus
End Sub

Private Sub cmdANAZHTHSH_Click()
    NEA.Visible = False
    KOD.Enabled = True
    ONO.Enabled = True
    neo.Enabled = True
 
End Sub

Private Sub cmdDelete_Click()

    Dim ANS As Integer

    ANS = MsgBox("Να διαγραφεί το είδος " + KOD.Text + " " + FONO.Text, vbYesNo)

    If ANS = vbYes Then
    Else

        Exit Sub

    End If

    Dim k As Long

    'σβηνω τις εικόνες
    For k = 0 To 6

        If Len(f_arxikh_eikona(k)) > 1 Then
            FileSystem.Kill F_FAKEL_KOINOS + "\images\" + f_arxikh_eikona(k)
        End If

    Next

    Gdb.Execute "delete FROM EID WHERE KOD='" + KOD.Text + "'", k

    Gdb.Execute "delete from  BARCODES WHERE KOD='" + KOD.Text + "'", k

    If k > 0 Then
  
        Gdb.Execute "delete PRODUCT_TO_CATEGORIES where KOD='" + KOD + "'", k
        Gdb.Execute "DELETE FROM PRODUCT_ATTRIBUTES where KOD='" + KOD + "' ", k
    End If

    If f_site = 111 Then
        dele_site_lagakis
        MsgBox "ΣΒΗΣΤΗΚΕ ΜΟΝΟ ΤΟΠΙΚΑ"

        Exit Sub

    End If

    '=======================================================   σιτε =========================================

    Dim m_ID As String

    Dim R    As New ADODB.Recordset

    gdblag.Close
    gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=www.toys-shop.gr;PORT =3306;DATABASE=toyshop;UID=dtek;Password=D35m80)!"
    R.Open "select * from products where products_model='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

    Dim site_foto(6) As String

    site_foto(0) = CNull(R!products_image)
    site_foto(1) = CNull(R!products_subimage1)
    site_foto(2) = CNull(R!products_subimage2)
    site_foto(3) = CNull(R!products_subimage3)
    site_foto(4) = CNull(R!products_subimage4)
    site_foto(5) = CNull(R!products_subimage5)
    site_foto(6) = CNull(R!products_subimage6)

    m_ID = str(R("products_id"))
    R.Close

    gdblag.Execute "delete from products_shipping      where products_id=" + m_ID + " limit 1", k
    gdblag.Execute "delete from products_description   where products_id=" + m_ID + " limit 1", k

    gdblag.Execute "delete from products_to_categories where products_id=" + m_ID + " limit 5", k

    gdblag.Execute "DELETE FROM products_attributes    where products_id=" + m_ID + "  limit 9", k

    gdblag.Execute "DELETE FROM products               where products_id=" + m_ID + "  limit 1", k

    file = 3
    Open "c:\mercvb\del_foto" For Output As #file
    Print #file, "open www.toys-shop.gr"
    Print #file, "toysftp"
    Print #file, "%t6Ixj43"
    'Print #file, "kP8#H7o2fD"
    Print #file, "binary"
    Print #file, "cd httpdocs/images"

    Dim nc As Integer

    nc = 0

    For k = 0 To 6
        Me.Caption = "delete foto" + str(k)

        If Len(site_foto(k)) > 0 Then
            Print #file, "del " + Replace(site_foto(k), "*", "-")  'βγαζω το αστερακι μην κανει ζημια μεγαλη
            nc = nc + 1
        End If

    Next

    Print #file, "quit"
    Close #file

    If nc > 0 Then
        Shell ("ftp -s:c:\mercvb\del_foto")
    End If

    MsgBox "ΔΙΕΓΡΑΦΗ ΤΟΠΙΚΑ ΚΑΙ ΑΠΟ ΤΟ SITE"

End Sub

Function dele_site_lagakis() As Integer  '1=ok  0=not ok

    Dim R    As New ADODB.Recordset

    Dim m_ID As String

    R.Open "select * from jsnrw_virtuemart_products where product_sku='" + KOD.Text + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

    If R.EOF Then
        MsgBox "δεν βρέθηκε στο site ο κωδικός " + KOD
        dele_site_lagakis = 0

        Exit Function

    End If

    m_ID = str(R("virtuemart_product_id"))
    R.Close

    gdblag.Execute "delete from jsnrw_virtuemart_products_el_gr        where virtuemart_product_id=" + m_ID + " limit 1"
    gdblag.Execute "delete from jsnrw_virtuemart_product_categories    where virtuemart_product_id=" + m_ID + " limit 20"
   
    Dim sql           As String

    Dim site_foto(16) As String

    Dim N             As Integer

    N = 0
    sql = "select * from jsnrw_virtuemart_medias   where  virtuemart_media_id in "
    sql = sql + " (select virtuemart_media_id  FROM jsnrw_virtuemart_product_medias where virtuemart_product_id=" + m_ID + ") "

    R.Open sql, gdblag, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        N = N + 1
        site_foto(N) = CNull(R!file_title)
        R.MoveNext
    Loop

    R.Close

    'σβησιμο απο sql τις αναφορές εικόνων
    sql = "delete from jsnrw_virtuemart_medias   where  virtuemart_media_id in "
    sql = sql + " (select virtuemart_media_id  FROM jsnrw_virtuemart_product_medias where virtuemart_product_id=" + m_ID + ") "
    gdblag.Execute sql

    gdblag.Execute "delete from jsnrw_virtuemart_product_medias where virtuemart_product_id=" + m_ID

    gdblag.Execute "delete from jsnrw_virtuemart_products    where virtuemart_product_id=" + m_ID + " limit 1"

    file = 3
    Open "c:\mercvb\del_foto" For Output As #file
    Print #file, "open www.lagakis.gr"
    Print #file, "lagakis"
    Print #file, "39213921"
    Print #file, "binary"
    Print #file, "cd images/stories/virtuemart/product"

    Dim nc As Integer

    nc = 0

    For k = 1 To N

        If Len(site_foto(k)) > 0 Then
            Print #file, "del " + Replace(site_foto(k), "*", "-")  'βγαζω το αστερακι μην κανει ζημια μεγαλη
            nc = nc + 1
        End If

    Next

    Print #file, "quit"
    Close #file
  
    If nc > 0 Then
        Shell ("ftp -s:c:\mercvb\del_foto")
    End If

    dele_site_lagakis = 1

End Function

Private Sub cmdPreview_Click()
    WebBrowser1.SetFocus

    SendKeys ("admin")
    SendKeys ("{TAB}")
    SendKeys ("lagakis")
    SendKeys ("{Enter}")

End Sub

Private Sub cmdΕΝΗΜΕΡΩΣΗ_Click()

    Dim a       As String

    Dim m_categ As String

    Dim y       As Long

    Dim m_new   As Integer
 
    'If neo.Enabled = True Then
    
    If FONO = "" Then Exit Sub
 
    'Else
    
    KOD.Enabled = True
 
    On Error Resume Next
  
    If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty

        On Error Resume Next

        y = T1.SelectedItem.Index
   
        y = T1.SelectedItem.Index
        m_categ = mID(names(y), 2, 4)
        Set T1.DropHighlight = T1.SelectedItem
    Else
        MsgBox "Δεν επιλέξατε κατηγορία "

        Exit Sub

    End If

    If Val(Left(Combo2.Text, 2)) < 1 Then
        MsgBox "Δεν επιλέξατε κατηγορία ΦΠΑ"

        Exit Sub

    End If
 
    Dim energo As String

    energo = "0"

    If Check1.Value = vbChecked Then
        energo = "1"
    End If
 
    Dim kor   As String, b As String

    Dim manuf As String

    kor = str(f_fylo(fylo.ListIndex))
    manuf = str(f_kataskeyasths(kataskeyasths.ListIndex))
 
    If xaraktiras.ListIndex = -1 Then
        b = 0
    Else
        b = str(f_xaraktiras(xaraktiras.ListIndex))
    End If
 
    'End If
    Dim ss

    Gdb.Execute "UPDATE EID SET FPA=" + Left(Combo2.Text, 2) + ",XARAKTIRAS=" + b + ",MANUFACTURERS_ID=" + manuf + ",FYLO=" + kor + " WHERE KOD='" + KOD + "'", ss
 
    If Check1.Value = vbChecked Then DIATH = "1" Else DIATH = "0"
 
    Gdb.Execute "UPDATE EID SET DIATHESIMO=" + DIATH + " WHERE KOD='" + KOD + "'", ss
'HAS_ASSORTMENT

Gdb.Execute "UPDATE EID SET HAS_ASSORTMENT=" + IIf(HAS_ASSORTMENT.Value = vbChecked, "1", "0") + " WHERE KOD='" + KOD + "'", ss

    Dim R As New ADODB.Recordset



    Gdb.Execute "delete from PRODUCT_TO_CATEGORIES where KOD='" + KOD + "'"

    R.Open "select * from PRODUCT_TO_CATEGORIES where KOD='" + KOD + "' and CATEGORIES_ID =" + m_categ, Gdb, adOpenDynamic, adLockOptimistic

    

    If R.EOF Then
        Gdb.Execute "INSERT INTO PRODUCT_TO_CATEGORIES (KOD,PRODUCT_ID,CATEGORIES_ID) VALUES ('" + KOD + "'," + str(F_id_LOCAL) + "," + m_categ + ");"
    End If

    R.Close
    'Gdb.Execute "INSERT INTO PRODUCT_ATTRIBUTES " & "(KOD,OPTIONS_ID,OPTIONS_VALUES_ID,PRODUCT_ID) " & " VALUES ('" + kod + "',1," + str(f_hlikies(K)) + "," + str(EID.Recordset("ID")) + ");"
    'PRODUCT_TO_CATEGORIES ( KOD VARCHAR(16) NOT NULL,PRODUCT_ID int NOT NULL,CATEGORIES_ID INT NOT NULL )

    '201    DUM = ADD_FIELD("EID", "DIATHESIMO", "BOOLEAN")
    '202 DUM = ADD_FIELD("EID", "HME_DIATH", "DATE")
    '203 DUM = ADD_FIELD("EID", "MANUFACTURERS_ID", "INT")
    '204 DUM = ADD_FIELD("EID", "TEXT_PRODUCTS_ZIPCODE", "VARCHAR(30)")
    '205 DUM = ADD_FIELD("EID", "ADDITIONALPRICE", "DECIMAL(10,2)")
    '206 DUM = ADD_FIELD("EID", "WEIGHT", "DECIMAL(10,2)")
    '
    '207 DUM = ADD_FIELD("EID", "FYLO", "INT")
    '208 DUM = ADD_FIELD("EID", "XARAKTIRAS", "INT")
    
    Gdb.Execute "DELETE FROM PRODUCT_ATTRIBUTES where KOD='" + KOD + "' ", k

    Dim sort As Integer

    For k = 0 To hlikies.ListCount - 1

        If hlikies.Selected(k) Then
            Gdb.Execute "INSERT INTO PRODUCT_ATTRIBUTES " & "(KOD,OPTIONS_ID,OPTIONS_VALUES_ID,PRODUCT_ID) " & " VALUES ('" + KOD + "',1," + str(f_hlikies(k)) + "," + str(EID.Recordset("ID")) + ");"
        End If

    Next k
    
    'Gdb.Execute "INSERT INTO BARCODES (KOD,ERG) VALUES ('" + KOD + "','" + KOD + "')"
 
    Dim ANS As Integer

    Dim k0  As Integer, mc As Integer
 
    For k = 0 To 6

        If f_arxikh_eikona(k) = PIC(k).Caption Then    'ok
        Else

            If Len(f_arxikh_eikona(k)) > 1 Then
                ' ελεγχω μηπως αυτην την φωτογραφια την εχω και άλλη φορα
                mc = 0

                For k0 = 0 To 6

                    If f_arxikh_eikona(k0) = PIC(k).Caption Then
                        mc = mc + 1
                    End If

                Next

                ' αν υπαρχει μονο μια φορα σημαίνει οτι αν την σβησω δεν θα κανω ζημιά
                If mc = 1 Then
                    'ans = MsgBox("Να διαγραφεί η παλιά εικόνα " + f_arxikh_eikona(k), vbYesNo)
                    'If ans = vbYes Then
                    FileSystem.Kill F_FAKEL_KOINOS + "\images\" + f_arxikh_eikona(k)
                    'End If
                End If
     
            End If
        End If
 
    Next

    EID.Recordset.Move 0
    
    NEA.Visible = False
    
    KOD.Enabled = True
    ONO.Enabled = True
    upload1.Enabled = True
     
    neo.Enabled = True

End Sub

Private Sub cmdΕξοδος_Click()
    'bCancel = True
    'DoEvents
    'gdblag.Close
 
    Unload Me
    'f_exit = True
    'Me.Hide
 
End Sub

Private Sub Command14_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo Command14_Click_Err

        '</EhHeader>
        Dim F_EIKONA As Long

100     F_EIKONA = 11
110     CommonDialog1.FILENAME = "*.jpg"
120     CommonDialog1.ShowOpen

        Dim m As String

        On Error Resume Next

130     m = "a" + kodlabel.Caption + "-" + Format(Index, "0") ' Text1(0).Text

140     If F_EIKONA = 11 Then
150         m = m + ".JPG"
        End If

160     FileCopy CommonDialog1.FILENAME, F_FAKEL_KOINOS + "\images\" + m
    
170     foto2(Index).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + m)
            
180     StretchSourcePictureFromPicture apot2vmn.foto2(Index).Picture, apot2vmn.foto2(Index)
        PIC(Index).Caption = m
    
        '<EhFooter>
        Exit Sub

Command14_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command14_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot2vmn.Command14_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub Command2_Click()

    Dim R    As New ADODB.Recordset

    Dim rLOC As New ADODB.Recordset

    'gdblag.Close  'SERVER=188.40.85.131
    'gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=www.toys-shop.gr;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
    R.Open "select * from products WHERE products_model<>'' order by products_model", gdblag, adOpenDynamic, adLockOptimistic

    Dim N    As Long

    Dim KODE As String

    Dim m_ID As String

    Do While Not R.EOF()
        KODE = Replace(R!PRODUCTS_MODEL, "'", "~")
        m_ID = str(R!PRODUCTS_ID)
        rLOC.Open "select * from EID WHERE KOD='" + KODE + "'", Gdb, adOpenDynamic, adLockOptimistic

        If rLOC.EOF Then
            Gdb.Execute "INSERT INTO EID (KOD,ID_WEB) VALUES ('" + KODE + "'," + m_ID + ")" ', n
        End If

        rLOC.Close
   
        '  If Not IsNull(r!products_youtube) Then
        '    Gdb.Execute "UPDATE EID SET YOU_TUBE='" + Replace(r!products_youtube, "'", "''") + "' WHERE KOD='" + KODE + "'"
        ' End If

        '        If Not IsNull(r!products_subimage1) Then
        '            Gdb.Execute "UPDATE EID SET PIC1='" + Replace(r!products_subimage1, "'", "''") + "' WHERE KOD='" + KODE + "'"
        '        End If
        '        If Not IsNull(r!products_subimage2) Then
        '           Gdb.Execute "UPDATE EID SET PIC2='" + r!products_subimage2 + "' WHERE KOD='" + KODE + "'"
        '        End If
        '        If Not IsNull(r!products_subimage3) Then
        '            Gdb.Execute "UPDATE EID SET PIC3='" + r!products_subimage3 + "' WHERE KOD='" + KODE + "'"
        '        End If
        '
        '        If Not IsNull(r!products_subimage4) Then
        '          Gdb.Execute "UPDATE EID SET PIC4='" + r!products_subimage4 + "' WHERE KOD='" + KODE + "'"
        '        End If
        '
        '        If Not IsNull(r!products_subimage5) Then
        '              Gdb.Execute "UPDATE EID SET PIC5='" + r!products_subimage5 + "' WHERE KOD='" + KODE + "'"
        '        End If
        '
        '       If Not IsNull(r!products_subimage6) Then
        '          Gdb.Execute "UPDATE EID SET PIC6='" + r!products_subimage6 + "' WHERE KOD='" + KODE + "'"
        '       End If
        '       If Not IsNull(r!products_price) Then
        '          Gdb.Execute "UPDATE EID SET WEBPRICE=" + str(r!products_price) + " WHERE KOD='" + KODE + "'"
        '       End If

        If Not IsNull(R!products_weight) Then Gdb.Execute "UPDATE EID SET WEIGHT=" + str(R!products_weight) + " WHERE KOD='" + KODE + "'"
        If Not IsNull(R!products_status) Then Gdb.Execute "UPDATE EID SET DIATHESIMO=" + str(R!products_status) + " WHERE KOD='" + KODE + "'"
       
        Dim FPA As String

        FPA = "2"

        If R!products_tax_class_id = 3 Then
            FPA = "1"
        ElseIf R!products_tax_class_id = 1 Then
            FPA = "2"
        ElseIf R!products_tax_class_id = 2 Then
            FPA = "4"
        End If
    
        Gdb.Execute "UPDATE EID SET FPA=" + FPA + " WHERE KOD='" + KODE + "'"
        Gdb.Execute "UPDATE EID SET ID_WEB=" + str(R!PRODUCTS_ID) + " WHERE KOD='" + KODE + "'"
       
        If Not IsNull(nufacturers_id) Then Gdb.Execute "UPDATE EID SET  MANUFACTURERS_ID=" + str(R!manufacturers_id) + " WHERE KOD='" + KODE + "'"
        If Not IsNull(R!manufacturers_id) Then Gdb.Execute "UPDATE EID SET  MANUFACTURERS_ID=" + str(R!manufacturers_id) + " WHERE KOD='" + KODE + "'"
       
        R.MoveNext
        N = N + 1

        DoEvents
        Me.Caption = str(N)
    Loop

    R.Close
          
    'INSERT INTO `PRODUCTS`
    '(products_id,products_youtube,products_quantity,products_model,products_image,products_subimage1,
    ' products_subimage2,products_subimage3,products_subimage4,products_subimage5,products_subimage6,
    '  products_price,
    ' products_date_added,products_last_modified,products_date_available,
    ' products_weight,products_status,products_tax_class_id,
    ' manufacturers_id,products_ordered) VALUES (30917,'nTnnW6YlrTk',12,'demostration','773-12510-GIM.jpg',
    ' NULL,NULL,NULL,NULL,NULL,NULL,'8.1220','2013-10-21 08:40:09','2013-10-21 08:53:59',NULL,'0.20',0,1,109,0);

    'FPA

End Sub

Private Sub Command3_Click()

    '=============================  products_description ================================
    Dim R    As New ADODB.Recordset

    Dim rLOC As New ADODB.Recordset

    GoTo 5000

    R.Open "select * from products_description order by products_id", gdblag, adOpenDynamic, adLockOptimistic

    Dim N    As Long

    Dim KODE As String

    Do While Not R.EOF()
        Gdb.Execute "UPDATE EID SET ONO='" + greek(Replace(R!PRODUCTS_NAME, "'", "''")) + "' WHERE ID_WEB=" + str(R!PRODUCTS_ID)
      
        ' Gdb.Execute "UPDATE EID SET MEMO='" + r!products_description + "' WHERE ID_WEB=" + str(r!PRODUCTS_ID)
       
        'Gdb.Execute "UPDATE EID SET XARAKTIRAS=" + str(r!extra_value_id1) + " WHERE ID_WEB=" + str(r!PRODUCTS_ID)
        'Gdb.Execute "UPDATE EID SET FYLO=" + str(r!extra_value_id3) + " WHERE ID_WEB=" + str(r!PRODUCTS_ID)

        R.MoveNext
        N = N + 1

        DoEvents
        Me.Caption = str(N)
        '    Exit Do
    Loop

    R.Close

    Exit Sub

5000

    R.Open "select * from products_to_categories", gdblag, adOpenDynamic, adLockOptimistic

    Dim rl As New ADODB.Recordset

    Do While Not R.EOF
        ' Gdb.Execute ""
        rl.Open "select * from EID WHERE ID_WEB=" + str(R!PRODUCTS_ID), Gdb, adOpenDynamic, adLockOptimistic

        If Not rl.EOF Then
            Gdb.Execute "INSERT INTO PRODUCT_TO_CATEGORIES (KOD,CATEGORIES_ID,PRODUCT_ID) VALUES ('" + Left(rl!KOD, 15) + "'," + str(R!categories_id) + ",0 );"
        End If

        rl.Close
        N = N + 1

        DoEvents
        Me.Caption = str(N)

        R.MoveNext
    Loop

    R.Close

    Exit Sub

    'ATTRIBUTES

6000

    'gdblag.Close
    'gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=toys-shop.gr ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"

    R.Open "select * from products_attributes ", gdblag, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        rLOC.Open "select * from EID WHERE ID_WEB=" + str(R!PRODUCTS_ID), Gdb, adOpenDynamic, adLockOptimistic

        If Not rLOC.EOF Then
            Gdb.Execute "INSERT INTO PRODUCT_ATTRIBUTES " & "(KOD,OPTIONS_ID,OPTIONS_VALUES_ID,PRODUCT_ID) " & " VALUES ('" + rLOC!KOD + "',1," + str(R!options_values_id) + ",0);"
        End If

        rLOC.Close
        N = N + 1

        DoEvents

        Exit Do

        Me.Caption = str(N)
        R.MoveNext
    Loop

    R.Close

    'gdblag.Execute "DELETE FROM products_attributes where products_id=" + m_ID + "  limit 9", k
    'Dim sort As Integer
    'sort = 0
    'For k = 0 To hlikies.ListCount - 1
    '
    '  If hlikies.Selected(k) Then
    '
    '      gdblag.Execute "INSERT INTO products_attributes " _
    '      & "(products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) " _
    '      & " VALUES (" + m_ID + ",1," + str(f_hlikies(k)) + ",'0.0000',''," + str(sort) + ");"
    '      sort = sort + 1
    '  End If
    '

    ' Gdb.Execute "CREATE TABLE PRODUCT_ATTRIBUTES ( KOD VARCHAR(16) NOT NULL,OPTIONS_VALUES_ID INT,PRODUCT_ID int NOT NULL,OPTIONS_ID  INT  NULL ) "

    'r.Open "SELECT * FROM PRODUCT_TO_CATEGORIES WHERE KOD='" + KOD + "'", Gdb, adOpenDynamic, adLockOptimistic
    ' Do While Not r.EOF
    '   ' T1.Nodes.Item(20).Expanded = True
    '   ' T1.Nodes.Item(7).Selected = True
    '   For k = 1 To UBound(names) - 1
    '        If Val(mID(names(k), 2, 4)) = r!CATEGORIES_ID Then
    '           T1.Nodes.Item(k).Selected = True
    '           T1.Nodes.Item(k).BackColor = vbCyan
    '           T1.SetFocus
    '          ' T1.ShowWhatsThis
    '
    '           Exit For
    '        End If
    '   Next
    '   Exit Do  ' GIA MIA KATHGORIA
    '   r.MoveNext
    ' Loop
    ' r.Close
 
    ' r.Open "SELECT * FROM PRODUCT_ATTRIBUTES WHERE KOD='" + KOD + "'", Gdb, adOpenDynamic, adLockOptimistic
    ' Do While Not r.EOF
    '    For k = 0 To hlikies.ListCount - 1
    '      If f_hlikies(k) = r("OPTIONS_VALUES_ID") Then

    'ΑΠΟΘΗΚΕΥΟΝΤΑΙ ΣΤΟ PRODUCTS_DESCRIPTION ΩΣ ΕΞΗΣ
    ' extra_value_id2,extra_value_id3)
    ' extra_value_id1=χαρακτηρας/θεμα  extra_value_id2=ηλικια  extra_value_id3=φυλο  13=ΑΓΟΡΙ&ΚΟΡΙΤΣΙ

    'INSERT INTO `products_description` (products_id,language_id,products_name,products_description,products_url,
    ' products_viewed,          extra_value_id1,extra_value_id2,extra_value_id3) extra_value_id1=χαρακτηρας/θεμα  extra_value_id2=ηλικια  extra_value_id3=φυλο  13=ΑΓΟΡΙ&ΚΟΡΙΤΣΙ
    ' VALUES (29306,4,'CITY LINE ? OVAL ORANGE #90698','<p></p>','',   22,0,0,13);

    '----------  ΟΡΙΣΜΟΣ   extra_field_labels  -------------
    ' extra_field_labels
    '1', '4', 'ΧΑΡΑΚΤΗΡΑΣ-ΘΕΜΑ', '1'
    '1', '1', 'Character/Theme', '1'
    '2', '4', 'ΗΛΙΚΙΑ', '1'    ' αυτο δεν χρησιμοποιείται πιά (δουλευει το products_attributes
    '2', '1', 'Age', '1'
    '3', '4', 'ΦΥΛΟ', '1'
    '3', '1', 'Gender', '1'

End Sub

Private Sub Command4_Click()

    Dim a As String, b As String

    a = InputBox("ΑΡΙΘΜΟΣ ΤΙΜΟΛ.ΑΓΟΡΑΣ (ME * ΕΠΙΛΕΓΩ ΗΜΕΡΟΜΗΝΙΕΣ ")
    
    If a = "*" Then
       
       
       b = InputBox("ΑΠΟ ΗΜΕΡ/ΝΙΑ (ΕΩΣ ΣΗΜΕΡΑ) ", , Format(Now, "DD/MM/YYYY"))
       
       
       
       Dim R2 As New ADODB.Recordset
       'ΝΑ ΜΗΝ ΕΜΦΑΝΙΖΕΙ ΑΥΤΑ ΠΟΥ ΕΧΟΥΝ 0 ΥΠΟΛΟΙΠΟ Ή ΑΥΤΑ ΠΟΥ ΕΧΟΥΝ ΗΔΗ ΦΩΤΟΓΡΑΦΙΕΣ (ΓΙΑΤΙ ΕΙΝΑΙ ΗΔΗ ΑΝΑΡΤΗΜΕΝΑ ΣΤΟ SITE)
       R2.Open "SELECT DISTINCT KODE FROM EGGTIM G inner JOIN EID E ON G.KODE=E.KOD WHERE   LEFT(ATIM,1) IN ('G','g') AND (E.NUM3 is null or  E.NUM3='2')  AND E.POS01+E.POS03>0 AND (LEN(E.PIC0)<2  OR E.PIC0 IS NULL) AND  HME>='" + Format(CDate(b), "MM/DD/YYYY") + "' ORDER BY KODE", Gdb, adOpenDynamic, adLockOptimistic
       
       ' r2.Open "SELECT DISTINCT KODE FROM EGGTIM inner JOIN EID ON EGGTIM.KODE=EID.KOD WHERE   LEFT(ATIM,1) IN ('G','g') AND EID.NUM3<>2 AND  KODE NOT IN (SELECT KOD FROM EID WHERE POS01+POS03<=0 OR LEN(PIC0)>2) AND  HME>='" + Format(CDate(b), "MM/DD/YYYY") + "' ORDER BY KODE", Gdb, adOpenDynamic, adLockOptimistic
       
       
       Open "C:\MERCVB\NEW_PRODUCTS.TXT" For Output As #1
       Do While Not R2.EOF
              Print #1, R2!KODE
              R2.MoveNext
       Loop
       Close #1
       
       R2.Close
       MsgBox "ΑΠΟΘΗΚΕΥΤΗΚΑΝ ΣΤΟ ΑΡΧΕΙΟ C:\MERCVB\NEW_PRODUCTS.TXT"
       EGGTIM2.ConnectionString = gConnect
       
       'EGGTIM2.RecordSource = "SELECT DISTINCT KODE FROM EGGTIM WHERE   LEFT(ATIM,1) IN ('G','g') AND  KODE NOT IN (SELECT KOD FROM EID WHERE POS01+POS03<=0 OR LEN(PIC0)>2) AND  HME>='" + Format(CDate(b), "MM/DD/YYYY") + "' ORDER BY KODE" '
       EGGTIM2.RecordSource = "SELECT DISTINCT KODE FROM EGGTIM G inner JOIN EID E ON G.KODE=E.KOD WHERE   LEFT(ATIM,1) IN ('G','g') AND E.NUM3<>'2' AND E.POS01+E.POS03>0 AND (LEN(E.PIC0)<2  OR E.PIC0 IS NULL)  AND  HME>='" + Format(CDate(b), "MM/DD/YYYY") + "' ORDER BY KODE"
       '"SELECT DISTINCT KODE FROM EGGTIM WHERE  LEFT(ATIM,1) IN ('G','g')  AND KODE NOT IN (SELECT KOD FROM EID WHERE LEN(PIC0)>2) AND  HME>='" + Format(CDate(b), "MM/DD/YYYY") + "' ORDER BY KODE"
       EGGTIM2.Refresh
    Else
       b = InputBox("ΗΜΕΡ/ΝΙΑ ΤΙΜΟΛ.ΑΓΟΡΑΣ", , Format(Now, "DD/MM/YYYY"))
       EGGTIM2.ConnectionString = gConnect
       EGGTIM2.RecordSource = "SELECT KODE FROM EGGTIM WHERE ATIM LIKE 'G%" + a + "%' AND CONVERT(CHAR(10),HME,103)='" + b + "' ORDER BY PROOD_AJ"
       EGGTIM2.Refresh
    End If
    If Not EGGTIM2.Recordset.EOF Then
        CMDmovenext.BackColor = vbGreen
        CMDmoveprevious.BackColor = vbGreen
        KOD.Text = EGGTIM2.Recordset("KODE")
        KOD.SetFocus
        FONO.SetFocus
    
    Else
    
        CMDmovenext.BackColor = vbRed
        CMDmoveprevious.BackColor = vbRed
    End If

End Sub

Private Sub Command5_Click()

    Dim rs As New ADODB.Recordset

    Dim R  As New ADODB.Recordset

    Dim r0  As New ADODB.Recordset


    Dim energo, mq, mTimh

    Dim k As Integer

    Dim m_extra As String
    Dim m_man As String
    Dim m_fylo As String
    
    Dim m_Weight As String
    Dim m_you As String


    Dim ANS

Dim FPA As String




Dim Details As Integer


    ANS = MsgBox("Να γινει ενημέρωση του site;", vbYesNo)

    If ANS = vbNo Then Exit Sub


    ANS = MsgBox("Μόνο τιμές & ποσότητες;", vbYesNo)

    If ANS = vbYes Then
         Details = 0
    Else
         Details = 1
    End If

    Dim cc As String
    
    cc = InputBox("ο κωδικός να αρχιζει από ", , "%")
    



'ΥΠΟΛΟΙΠΑ ΑΠΟΘΗΚΩΝ
Dim FF As New apot3param
 FF.Label_titlos = "Υπόλοιπα Ειδών"

FF.load_sthles_EIDON
'FF.LOAD_KINHSEIS_EIDON
FF.Rep_Fields.Text = "select KOD as [Κωδικός],LEFT(ONO,35) as [Περιγραφή],POS01 AS [Υπόλοιπο]"
FF.gemisma_Krithrion
MDIForm1.fkrithria.Grid1.TextMatrix(4, 1) = 0.1
MDIForm1.fkrithria.Grid1.TextMatrix(4, 2) = 999999999
MDIForm1.fkrithria.ΕΧΙΤ_Click
FF.sumes.Text = "001"
FF.WHERE_QUERY.Text = "POS>=0.1 AND POS<=999999999" ' StatusBar1.Panels(6).Text"
'YPOLOGISMOS_Click
'FF.SHOW
FF.YPOLOGISMOS_Click
FF.CMDeXIT_Click
Set FF = Nothing









Dim R9 As New ADODB.Recordset

    Dim TIM

    TIM = GetCurrentTime()

    On Error Resume Next 'GoTo command5_out

    R.Open "select  * from EID WHERE len(KOD)>1 AND LEN(PIC0)>2 and KOD LIKE '" + cc + "'", Gdb, adOpenDynamic, adLockOptimistic

    gdblag.Close
    gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=176.9.148.246;PORT =3306;DATABASE=toyshop;UID=dtek;Password=D35m80)!"


Dim mproductsID As String


    Do While Not R.EOF
   
        If IsNull(R!WEBPRICE) Then
            If IsNull(R!LTI5) Then
                mTimh = "0"
            Else
       
                mTimh = Format(R!LTI5 / ((100 + g_Fpa(R!FPA)) / 100), "####.00")
       
            End If

        Else
            mTimh = Format(R!WEBPRICE / ((100 + g_Fpa(R!FPA)) / 100), "####.00")
        End If
   
        If IsNull(R!POS) Then
            mq = "0"
            energo = "0"
        Else
     
            mq = Format(R!POS, "####.00")

            If R!POS > 0 Then
                energo = "1"
            Else
                energo = "0"
            End If
    
            'If Val(mq) < 5 Then
            '      energo = "0"
            'End If
                 
            If Val(mTimh) < f_Limit_Timh Then
                If Val(mq) < f_Limit_Posothta Then
                    energo = "0"
                End If
            End If
            
            
        End If
        
        
        If Details = 0 Then
           gdblag.Execute "update products set web_updated=0, products_price=" + mTimh + ",products_quantity=" + mq + ",products_status=" + energo + " where products_model='" + R!KOD + "' limit 1"
        End If
         
        
        '============ TA YPOLOIPA PEDIA ==========================================
        If Details = 1 Then
        
        
       m_man = "0"
       If IsNull(R!manufacturers_id) Then
        
       Else
            m_man = str(R!manufacturers_id)
       End If
        
       m_extra = ""
       If IsNull(R!xaraktiras) Then  ' extra_value_id1)
       
       Else
         m_extra = str(R!xaraktiras)
       End If
        
              
        
        
       m_fylo = ""
       If IsNull(R!fylo) Then ' Then   extra_value_id3
       
       Else
         m_fylo = str(R!fylo)
       End If
        
          
       m_Weight = ""
       If IsNull(R!WEIGHT) Then  ' products_weight' products_weight
       
       Else
         m_Weight = str(R!WEIGHT)
       End If
        
        
'        "YOU_TUBE", R!products_youtube,
        
        
       m_you = ""
       If IsNull(R!YOU_TUBE) Then   ' products_youtube
       Else
         m_you = R!YOU_TUBE
       End If
   
        
        List2.AddItem R!KOD, 0
        DoEvents
   
        gdblag.Execute "update products set products_youtube='" + m_you + "',products_weight=" + m_Weight + ",manufacturers_id=" + m_man + ",web_updated=0, products_price=" + mTimh + ",products_quantity=" + mq + ",products_status=" + energo + " where products_model='" + R!KOD + "' limit 1"
        
        
        
        
          'dblag.Execute "update products_description set  extra_value_id3=" + kor + ",extra_value_id1=" + b + " where products_id=" + m_ID, n
         
         r0.Open "select products_id from products where products_model='" + R!KOD + "' limit 1", gdblag, adOpenDynamic, adLockOptimistic
         
         If Not r0.EOF Then
              mproductsID = str(r0(0))
              gdblag.Execute "update products_description  set extra_value_id1=" + m_extra + " , extra_value_id3=" + m_fylo + "  where products_id=" + str(r0(0)) + " limit 1"
         End If
         r0.Close
    
    
    
    Dim aa  As String
    
    Dim m_erg As String
    
    
    m_erg = Replace(R!ONO, "'", "`")
    
    If IsNull(R("ERG")) Then
    Else
        If Len(R("ERG")) > 0 Then
            m_erg = m_erg + " #" + R("ERG")
        End If
    End If
    
    
'
'    If IsNull(R!memo) Then
'        aa = ""
'    Else
'       aa = RichToHTML(R!memo, 0&, Len(R!memo))  '(.TextRTF)
'       aa = Replace(a, "'", "`")
'    End If
'
'
'
'    m_erg = Replace(R!ONO, "'", "`")
'
'    If IsNull(R("ERG")) Then
'    Else
'        If Len(R("ERG")) > 0 Then
'            m_erg = m_erg + " #" + R("ERG")
'        End If
'    End If
'
'
    
  
    
          KOD.Text = R!KOD
          EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + R!KOD + "'"
          EID.Refresh
          UPDATE_CONTROLS
    
    
         EID.Recordset.Move 0  ' gia na ginei to rich1 opos to mseditor21
 
    'aa = RichToHTML(Rich1, 0&, Len(Rich1.Text))  '(.TextRTF)
    'aa = Replace(a, "'", "`")
    

       gdblag.Execute "update products_description set  products_name='" + m_erg + "',products_description='" + RichToHTML(Rich1, 0&, Len(Rich1.Text)) + "' where products_id=" + mproductsID, N
        
        
        
        
        
        If R!HAS_ASSORTMENT = 1 Then
              gdblag.Execute "update products set  has_assortment = 1   where products_model='" + R!KOD + "' limit 1"
        Else
              gdblag.Execute "update products set  has_assortment = 0   where products_model='" + R!KOD + "' limit 1"
        End If
        
        If Not IsNull(R!erg) Then
           gdblag.Execute "update products set erg_code = '" + R!erg + "'  where products_model='" + R!KOD + "' limit 1"
        End If
        
        
        
'---------------------------
    FPA = "1"
    If R!FPA = 1 Then
        FPA = "3"
    ElseIf R!FPA = 2 Then
        FPA = "1"
    ElseIf R!FPA = 4 Then
        FPA = "2"
    End If
    gdblag.Execute "update products set  products_tax_class_id=" + FPA + "  where products_model='" + R!KOD + "' limit 1"


msh = ""
'SHIPPINGPRICE
 If IsNull(R!SHIPPINGPRICE) Then
     msh = "0"
 Else
     msh = str(R!SHIPPINGPRICE)
     R9.Open "select * from products_shipping where products_id=" + mproductsID, gdblag, adOpenDynamic, adLockOptimistic
    If R9.EOF Then
        gdblag.Execute "INSERT INTO products_shipping (products_id,products_ship_methods_id,products_ship_zip,products_ship_price,products_ship_price_two) VALUES (" + mproductsID + ",0,'','','');", N
    End If
    R9.Close
    gdblag.Execute "update products_shipping set products_ship_price='" + msh + "' where products_id=" + mproductsID, N
End If






 
 'gdblag.Execute "update products set  set products_tax_class_id=" + fpa + "  where products_model='" + R!KOD + "' limit 1"


'-----------------------------


'-----------------------------
        
        
        
        
        
        'gdblag.Execute
'        sql = "update products_shipping set  products_ship_price_two='" + ADDIT_PRICE.Text
'        sql = sql + "',products_ship_price='" + INDV_PRICE.Text + "',products_ship_zip='"
'        sql = sql + ZIP_CODE.Text + "' where products_id=" + m_ID
'
        
        ', n
        
        
        
        '============ TA YPOLOIPA PEDIA ==========================================
        
        
        
        
        
        End If ' DETAILS = 1 Then
        
        
        
        
        
        R.MoveNext
        k = k + 1
   
        If k Mod 10 = 0 Then

            DoEvents
            Me.Caption = str(k)
        End If

    Loop

    MsgBox "OK ΕΝΗΜΕΡΩΘΗΚΕ " + Format((GetCurrentTime() - TIM) / 1000, "#######.00")

    Exit Sub

command5_out:

    On Error Resume Next

    MsgBox "λαθος στο ειδος " + R!KOD

End Sub

Private Sub Command6_Click()

'Open "C:\MERCVB\UPDIDWEB.TXT" For Output As #3
'Open "C:\MERCVB\EIDHLIP.TXT" For Input As #2
'
'Dim AA As String
'Dim BB() As String
'Dim KN As Integer
'
'Do While Not EOF(2)
' Input #2, AA
' BB = Split(AA, ",")
'
' BB = Split(AA, ";")
' Gdb.Execute "UPDATE EID SET KODSYNOD='3004',ID_WEB=" + BB(0) + " WHERE KOD='" + BB(1) + "'", KN
' If KN = 1 Then
'    List2.AddItem "OK " + BB(1)
'    Print #3, "OK " + BB(1)
'
'
' Else
'
'    List2.AddItem "not OK " + BB(1)
'    Print #3, "not OK " + BB(1)
' End If
'
'
'Loop
'
'Close #2
'Close #3
'
'Exit Sub

'List2.Clear
'
'
'Dim R300 As New ADODB.Recordset
''Open "C:\MERCVB\TELEYTAIA.TXT" For Output As #1
'Dim KN As Integer
'Dim r2 As New ADODB.Recordset
'r2.Open "select * from EID WHERE   LEN(PIC0)>2 ", Gdb, adOpenDynamic, adLockOptimistic
''EID.Refresh
'   gdblag.Close
'   gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=176.9.148.246;PORT =3306;DATABASE=toyshop;UID=dtek;Password=D35m80)!"   ' www.toys-shop.gr
'
'
'Do While Not r2.EOF
'
'
'
'    List2.AddItem r2!kod, 0
'    DoEvents
'
'    R300.Open "select * from products order by products_id", gdblag, adOpenDynamic, adLockOptimistic
'R300.MoveLast
'
'For ll = 1 To 300
'             Print #1, R300!PRODUCTS_MODEL + " ," + Format(R300!PRODUCTS_ID, "#####00")
'             R300.MovePrevious
'Next
'R300.Close
'Close #1
'
'
 'R300.Open "select * from products order by products_id desc", gdblag, adOpenDynamic, adLockOptimistic





'R300.Open "select MAX(product_id) from products  ", gdblag, adOpenDynamic, adLockOptimistic
    
  On Error GoTo 0
    
    
    Dim synol As Long
Dim R2 As New ADODB.Recordset
Dim r300 As New ADODB.Recordset

R2.Open "select count(*) from EID WHERE   LEN(PIC0)>2 and (ID_WEB IS NULL OR ID_WEB=0)", Gdb, adOpenDynamic, adLockOptimistic
    synol = R2(0)
    R2.Close
    
    Dim nn As Long
    nn = 0
    Dim nn2 As Long
    nn2 = 0
    
R2.Open "select * from EID WHERE   LEN(PIC0)>2 and (ID_WEB IS NULL OR ID_WEB=0) ", Gdb, adOpenDynamic, adLockOptimistic

Do While Not R2.EOF
    DoEvents
    r300.Open "select * from products where products_model='" + R2!KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
    If r300.EOF Then
          r300.Close
          KOD.Text = R2!KOD
          EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + R2!KOD + "'"
          EID.Refresh
          UPDATE_CONTROLS
          mpoyg_upload
          nn2 = nn2 + 1
          'Me.Caption = r2!kod
    Else
          Gdb.Execute "UPDATE EID SET ID_WEB=" + str(r300!PRODUCTS_ID) + "   WHERE KOD='" + r300!PRODUCTS_MODEL + "'"
          r300.Close
    End If
    R2.MoveNext
    nn = nn + 1
    If synol > 0 Then
       Me.Caption = Format(nn / synol, "##0.00") + "%   Νέα είδη " + Format(nn2, "#####0")
    End If
Loop
    

R2.Close



Close #1



End Sub



Private Sub mpoyg_upload2()












End Sub







Private Sub Command7_Click()
'ΠΑΙΡΝΕΙ ΜΙΑ ΛΙΣΤΑ ΙΔ ΚΑΙ ΚΩΔΙΚΟ ΚΑΙ ΑΝΟΙΓΕΙ ΣΤΟ SITE ΤΟΥΣ ΑΝΤΙΣΤΟΙΧΟΥΣ ΚΩΔΙΚΟΥΣ



'Open "C:\MERCVB\UPDIDWEB.TXT" For Output As #3
'Open "C:\MERCVB\EIDHLIP.TXT" For Input As #2
'
'Dim AA As String
'Dim BB() As String
'Dim KN As Integer
'
'Do While Not EOF(2)
' Input #2, AA
' BB = Split(AA, ",")
'
' BB = Split(AA, ";")
' Gdb.Execute "UPDATE EID SET KODSYNOD='3004',ID_WEB=" + BB(0) + " WHERE KOD='" + BB(1) + "'", KN
' If KN = 1 Then
'    List2.AddItem "OK " + BB(1)
'    Print #3, "OK " + BB(1)
'
'
' Else
'
'    List2.AddItem "not OK " + BB(1)
'    Print #3, "not OK " + BB(1)
' End If
'
'
'Loop
'
'Close #2
'Close #3
'
'Exit Sub

List2.Clear


Dim r300 As New ADODB.Recordset
'Open "C:\MERCVB\TELEYTAIA.TXT" For Output As #1


Dim KN As Long

Dim R2 As New ADODB.Recordset



   gdblag.Close
   gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=176.9.148.246;PORT =3306;DATABASE=toyshop;UID=dtek;Password=D35m80)!"   ' www.toys-shop.gr

  ' R300.Open "select * from products order by products_id", gdblag, adOpenDynamic, adLockOptimistic



'Do While Not r2.EOF
' R300.Close


 '  R300.Open "select *  from products where products_id>40670 order by products_id", gdblag, adOpenDynamic, adLockOptimistic





Dim sql As String


'KN = R300!products_id

For KN = 40683 To 41007 ' ID
    
    List2.AddItem str(KN)
    
    DoEvents
    
   
    
    
    
  R2.Open "select top 1 * from EID WHERE  KODSYNOD='3004' AND   LEN(PIC0)>2 and ID_WEB=" + str(KN), Gdb, adOpenDynamic, adLockOptimistic
    If R2.EOF Then
    
        'sql = "INSERT INTO products (products_youtube,products_quantity,products_model,products_image,products_subimage1,"
        'sql = sql + " products_subimage2,products_subimage3,products_subimage4,products_subimage5,products_subimage6,products_price,"
        'sql = sql + " products_date_added,products_last_modified,products_date_available,products_weight,products_status,products_tax_class_id,"
        'sql = sql + " manufacturers_id,products_ordered) VALUES ('',6,'order','', "
        'sql = sql + " NULL,NULL,NULL,NULL,NULL,NULL,'','" + M_DATE + "','" + M_DATE + "',NULL,'0.20',1,1,109,0); "
        gdblag.Execute "update products set products_model='order' where products_id=" + str(KN), N
    Else
    
'       R300.Close
       r300.Open "select * from products  where products_id=" + str(KN), gdblag, adOpenDynamic, adLockOptimistic
       sql = r300!PRODUCTS_MODEL
       r300.Close
       
       If sql = R2!KOD Then
             sql = ""
       Else
         gdblag.Execute "update products set products_model='" + R2!KOD + "' where products_id=" + str(KN), N
          KOD.Text = R2!KOD
          EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + R2!KOD + "'"
          EID.Refresh
          UPDATE_CONTROLS
          mpoyg_upload
       End If
        
    End If
    R2.Close
    
    
    
 Next
    
Exit Sub
    
    
    
    
    
    
    
 'R300.Open "select * from products order by products_id", gdblag, adOpenDynamic, adLockOptimistic
'R300.MoveLast

'For ll = 1 To 300
 '            Print #1, R300!PRODUCTS_MODEL + " ," + Format(R300!PRODUCTS_ID, "#####00")
  '           R300.MovePrevious
'Next
'R300.Close
'Close #1


 'R300.Open "select * from products order by products_id desc", gdblag, adOpenDynamic, adLockOptimistic





r300.Open "select MAX(product_id) from products  ", gdblag, adOpenDynamic, adLockOptimistic
    
    r300.Open "select * from products where products_model='" + R2!KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
   

    If r300.EOF Then
          
          r300.Close
          
          '//  DEBUG
          
          If IsNull(R2!ID_WEB) Then
            FF = "0"
            Else
            FF = Format(R2!ID_WEB, "#######0")
          End If
          
          Print #1, R2!KOD + "," + R2!erg + "," + R2!ONO + "," + FF
          GoTo 9999
          '//  DEBUG
          
          KOD.Text = R2!KOD
          EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + R2!KOD + "'"
          EID.Refresh
          UPDATE_CONTROLS
          mpoyg_upload
          
          '//  DEBUG
9999
          '//  DEBUG
          
          
          
          Debug.Print R2!KOD
          Me.Caption = R2!KOD
       
        
       
       Else
       
       r300.Close
       

    End If
'    R300.Close
    
   ' r2.MoveNext
    
'Loop
    

R2.Close



Close #1


End Sub

Private Sub d1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim ss As Single

    If KeyCode = 13 Then
        ss = Val(b1.Text) * Val(c1.Text) * Val(d1.Text) / 5000
        ss = ss / 2
        e1.Caption = Format(ss, "####0.00")
        b1.SetFocus
    End If

End Sub

Private Sub download_eidh_Click()

    Dim R  As New ADODB.Recordset

    Dim rl As New ADODB.Recordset


Dim eikona(10)          As String



    Dim ANS

    ANS = MsgBox("Να γινει ενημέρωση του εμπορικού από το site;", vbYesNo)
    
    
    
    

    If ANS = vbNo Then Exit Sub



fmaxRec = InputBox("Πόσες εγγραφές το μέγιστο; ", , 10)






    Dim site_foto(100)

    List2.Clear
    gdblag.Close
    gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=176.9.148.246;PORT =3306;DATABASE=toyshop;UID=dtek;Password=D35m80)!"



 R.Open "select count(*) from products where web_updated=1", gdblag, adOpenDynamic, adLockOptimistic
  MsgBox "υπάρχουν συνολικά " + str(R(0)) + " είδη που ενημερώθηκαν απο το site"
  R.Close
  





    gdblag.Execute "update products  set web_updated=0 where PRODUCTS_MODEL like '%order%'  "



    Dim R3 As New ADODB.Recordset

    R.Open "select *,products_description.* from products  inner join products_description  on products.products_id=products_description.products_id where web_updated=1    order by products.products_id desc  limit 0," + str(fmaxRec), gdblag, adOpenDynamic, adLockOptimistic

    
    
    
   file = 22: Open "c:\mercvb\get_foto" For Output As #file
        Print #file, "open 176.9.148.246"
        Print #file, "toyshopftp"
        Print #file, "T35m80)!@"
        Print #file, "binary"
        Print #file, "cd httpdocs/images"
  
    
    
    
    
    
    
    
    
    Do While Not R.EOF
   
        m_ID = R!PRODUCTS_ID
        R3.Open "select * from products_shipping where products_id=" + str(m_ID), gdblag, adOpenDynamic, adLockOptimistic

        rl.Open "select * from EID WHERE KOD='" + R!PRODUCTS_MODEL + "'", Gdb, adOpenDynamic, adLockOptimistic

        If rl.EOF Then
            List2.AddItem html2(R!PRODUCTS_DESCRIPTION) '     r!PRODUCTS_MODEL + r("products_name")
            Gdb.Execute "insert into EID (KOD,ONO,MEMO) VALUES ('" + Replace(R!PRODUCTS_MODEL, "'", "''") + "','" + Replace(R!PRODUCTS_NAME, "'", "''") + "','" + html2(R!PRODUCTS_DESCRIPTION) + "')"
        Else
            List2.AddItem R!PRODUCTS_MODEL + " Υπάρχει"
        End If




       gdblag.Execute "update products set web_updated=3 where products_id=" + str(m_ID)



        Dim DUM As Integer
   
        'html2(r!PRODUCTS_DESCRIPTION)
        DUM = UPD_EID("MEMO", "'" + html2(R!PRODUCTS_DESCRIPTION) + "'", R!PRODUCTS_MODEL)
        
        DUM = UPD_EID("ONO", "'" + CNull(R!PRODUCTS_NAME) + "'", R!PRODUCTS_MODEL)
  
        DUM = UPD_EID("PIC0", "'" + CNull(R!products_image) + "'", R!PRODUCTS_MODEL)
        DUM = UPD_EID("PIC1", "'" + CNull(R!products_subimage1) + "'", R!PRODUCTS_MODEL)
        DUM = UPD_EID("PIC2", "'" + CNull(R!products_subimage2) + "'", R!PRODUCTS_MODEL)
        DUM = UPD_EID("PIC3", "'" + CNull(R!products_subimage3) + "'", R!PRODUCTS_MODEL)
        DUM = UPD_EID("PIC4", "'" + CNull(R!products_subimage4) + "'", R!PRODUCTS_MODEL)
        DUM = UPD_EID("PIC5", "'" + CNull(R!products_subimage5) + "'", R!PRODUCTS_MODEL)
        DUM = UPD_EID("PIC6", "'" + CNull(R!products_subimage6) + "'", R!PRODUCTS_MODEL)
  
      
  
        DUM = UPD_EID("WEIGHT", Replace(str(R!products_weight), ",", ".") + "", R!PRODUCTS_MODEL)
  
        If Not R3.EOF Then
            DUM = UPD_EID("TEXT_PRODUCTS_ZIPCODE", "'" + CNull(R3!products_ship_zip) + "'" + "", R!PRODUCTS_MODEL)
            DUM = UPD_EID("ADDITIONALPRICE", str(Val(nNull(R3!products_ship_price_two))) + "", R!PRODUCTS_MODEL)
            DUM = UPD_EID("SHIPPINGPRICE", "'" + str(Val(nNull(R3!products_ship_price))) + "'" + "", R!PRODUCTS_MODEL)
        End If
  
        DUM = UPD_EID("DIATHESIMO", str(IIf(IsNull(R!products_status), 0, R!products_status)), R!PRODUCTS_MODEL)
  
        DUM = UPD_EID("MANUFACTURERS_ID", str(R!manufacturers_id), R!PRODUCTS_MODEL)
  
        'DIATHESIMO
        'products_status=" + energo + ",manufacturers_id=" + manuf + ",
        Dim FPA As String: FPA = "2"

        '**********  fpa tax class ******************************
        If R("PRODUCTS_tax_class_ID") = 1 Then FPA = "2"
        If R("PRODUCTS_tax_class_ID") = 3 Then FPA = "1"
        If R("PRODUCTS_tax_class_ID") = 2 Then FPA = "4"
  
  
    DUM = UPD_EID("WEBPRICE", Replace(str(R!products_PRICE * (100 + g_Fpa(Val(FPA))) / 100), ",", ".") + "", R!PRODUCTS_MODEL)
  
  
        DUM = UPD_EID("FPA", FPA, R!PRODUCTS_MODEL)
  
       
  
  
   'gdblag.Execute "update products  set web_updated=0 where web_updated=3  "
  
  
        DUM = UPD_EID("YOU_TUBE", R!products_youtube, R!PRODUCTS_MODEL)
        R3.Close

        'PRODUCTS_tax_class_ID=1 =>  23%
        'PRODUCTS_tax_class_ID=2 => 6.5%
        'PRODUCTS_tax_class_ID=3 => 13%

        R3.Open "SELECT * from products_to_categories where products_id=" + str(m_ID) + " limit 2", gdblag, adOpenDynamic, adLockOptimistic
 
        'If r.EOF Then
        Gdb.Execute "DELETE FROM PRODUCT_TO_CATEGORIES WHERE KOD='" + R!PRODUCTS_MODEL + "'"

        Gdb.Execute "INSERT INTO PRODUCT_TO_CATEGORIES (KOD,PRODUCT_ID,CATEGORIES_ID) VALUES ('" + R!PRODUCTS_MODEL + "'," + str(m_ID) + "," + str(R3!categories_id) + ");"
        R3.Close
        'End If
        'PRODUCT_TO_CATEGORIES ( KOD VARCHAR(16) NOT NULL,PRODUCT_ID int NOT NULL,CATEGORIES_ID INT NOT NULL )
  
        '-----hlikies ------------------
        R3.Open "SELECT * FROM products_attributes where options_id=1 and products_id=" + str(m_ID) + " oRder by products_options_sort_order limit 9", gdblag, adOpenDynamic, adLockOptimistic

        Do While Not R3.EOF
            '  gdblag.Execute "INSERT INTO products_attributes " _
            '  & "(products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) " _
            '  & " VALUES (" + m_ID + ",1," + str(f_hlikies(k)) + ",'0.0000',''," + str(sort) + ");", n
            '  sort = sort + 1
            Gdb.Execute "INSERT INTO PRODUCT_ATTRIBUTES " & "(KOD,OPTIONS_ID,OPTIONS_VALUES_ID,PRODUCT_ID) " & " VALUES ('" + R!PRODUCTS_MODEL + "',1," + str(R3!options_values_id) + "," + str(m_ID) + ");"
            R3.MoveNext
        Loop

        R3.Close
  
        'insert into products_description(language_id,products_id,extra_value_id3,
        '   extra_value_id1,products_name,products_description) values
        '   (4," + m_ID + "," + kor + "," + b +
  
        DUM = UPD_EID("FYLO", str(R!extra_value_id3), R!PRODUCTS_MODEL)
        DUM = UPD_EID("XARAKTIRAS", str(R!extra_value_id1), R!PRODUCTS_MODEL)
  
  
  
  ' Gdb.Execute "UPDATE EID SET FPA=" + Left(Combo2.Text, 2) + ",XARAKTIRAS=" + b + ",MANUFACTURERS_ID=" + manuf + ",FYLO=" + kor + " WHERE KOD='" + KOD + "'", ss
  
        
        
        
        
        
  
  
   eikona(0) = CNull(R!products_image)
   eikona(1) = CNull(R!products_subimage1)
   eikona(2) = CNull(R!products_subimage2)
   eikona(3) = CNull(R!products_subimage3)
   eikona(4) = CNull(R!products_subimage4)
   eikona(5) = CNull(R!products_subimage5)
   eikona(6) = CNull(R!products_subimage6)
   'eikona(1) = CNull(R!PRODUCTS_subimage1)
   
   
   
  
        For k = 0 To 6
            Me.Caption = "DOWNload foto" + str(k)

            If Len(eikona(k)) > 0 Then
     
                'την εχει ηδη την φωτογραφια  den kanv tipota
                'If site_foto(k) = PIC(k).Caption And reload.Value = vbUnchecked Then
                'Else

                   ' If Len(site_foto(k)) > 1 Then
                   '     Print #file, "del " + Replace(site_foto(k), "*", "-")  'βγαζω το αστερακι μην κανει ζημια μεγαλη
                   ' End If

                    'If ftp.Text = ftp.List(2) Then
                     '   ftp_Eikones F_FAKEL_KOINOS + "\images\" + PIC(k).Caption, "httpdocs/images/" + PIC(k).Caption
                    'Else
                        Print #file, "get " + eikona(k) + "  " + F_FAKEL_KOINOS + "\images\" + eikona(k)
                   ' End If
        
                'End If
      
                nc = nc + 1
            End If

        Next

     

   
  
  
  
  
        rl.Close
   
        R.MoveNext
        List2.Width = 2800
    Loop

    'rl.Close
    R.Close


   Print #file, "quit"
        Close #file




     '   If nc > 0 Then
            Shell ("ftp -s:c:\mercvb\get_foto")
      '  End If
  







    '=================================================================================
   file = 13
   Open "c:\mercvb\send_foto" For Output As #file
    Print #file, "open 176.9.148.246"
    Print #file, "toyshopftp"
    Print #file, "T35m80)!@"
    Print #file, "binary"
    Print #file, "cd httpdocs/images"
  
    For k = 0 To 6
        Me.Caption = "upload foto" + str(k)

        If Len(PIC(k).Caption) > 0 Then
     
            'την εχει ηδη την φωτογραφια  den kanv tipota
            If site_foto(k) = PIC(k).Caption Then
            Else

                If Len(site_foto(k)) > 1 Then
                    'Print #file, "del " + Replace(site_foto(k), "*", "-")  'βγαζω το αστερακι μην κανει ζημια μεγαλη
                End If

                'If ftp.Text = ftp.List(2) Then
                '    ftp_Eikones F_FAKEL_KOINOS + "\images\" + PIC(k).Caption, "httpdocs/images/" + PIC(k).Caption
                'Else
                'Print #file, "get " + F_FAKEL_KOINOS + "\images\" + PIC(k).Caption
                Print #file, "get " + PIC(k).Caption
                'End If
        
            End If
      
            nc = nc + 1
        End If

    Next

    Print #file, "quit"
    Close #file

    If nc > 0 Then
        Shell ("ftp -s:c:\mercvb\get_foto")
    End If

    'End If  '=================================================================================

    R.Open "select count(*) from products   where web_updated=1 ", gdblag, adOpenDynamic, adLockOptimistic
    Me.Caption = R(0)
    R.Close



oristiki.Enabled = True




End Sub

Function UPD_EID(ByVal F_eid As String, _
                 ByVal ValField As String, _
                 ByVal mkod As String) As Integer
    'If IsEmpty(ValField) Then ValField = "''"

    If Len(ValField) = 0 Or IsEmpty(ValField) Then
        If F_eid = "ADDITIONALPRICE" Or F_eid = "SHIPPINGPRICE" Then
            ValField = "0"
        Else
            ValField = "''"
        End If
    End If
On Error Resume Next
     Gdb.Execute "UPDATE EID SET  " + F_eid + "=" + ValField + " WHERE  KOD='" + mkod + "'"
End Function

Function html2(ByVal a As String)

    Dim k, N, n2

    Dim l

    l = Len(a)

    For k = 1 To 200
        n2 = InStrRev(a, ">")
        N = InStrRev(a, "<")

        If n2 * N = 0 Then

            Exit For

        End If

        a = mID(a, 1, N - 1) + IIf(Len(a) - n2 > 0, mID(a, n2 + 1, Len(a) - n2), "")
   
    Next

    html2 = Replace(a, "'", "''")
    html2 = Replace(a, ".", "." + Chr(13))
End Function

Private Sub FONO_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub Form_DblClick()

     PARAMETROI.PARAM.Caption = "APOT2VMN"
     PARAMETROI.SHOW 1


End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     Me.AutoRedraw = True
 
102     f_pel = "mpoyg" ' "lagak"
 
        'syndesh xori dsn me connection string  ' ip toys 188.40.85.131
        ' gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=db5.papaki.gr;PORT =3306;DATABASE=vasi;UID=lagakis;Password=3898339214231"
        On ERRROR GoTo NOCONNECT
        
        
        
        f_Limit_Posothta = Val(FINDPARAMETROI(1, "APOT2VMN", "f_Limit_Posothta", "5", "ΟΡΙΟ ΠΟΣΟΤΗΤΑΣ ΓΙΑ ΑΝΕΝΕΡΓΟ"))
        
        f_Limit_Timh = Val(FINDPARAMETROI(1, "APOT2VMN", "f_Limit_Timh", "8", "ΟΡΙΟ ΤΙΜΗΣ ΓΙΑ ΑΝΕΝΕΡΓΟ"))
        
        
        
        
        
        
        f_site = Val(FINDPARAMETROI(1, "MDIFORM1", "F_SITE", "0", "ΚΩΔΙΚΟΣ SITE 1=VM 2=TOYS"))

        If f_site = 222 Then
            f_connect = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=176.9.148.246;PORT =3306;DATABASE=toyshop;UID=dtek;Password=D35m80)!"  'www.toys-shop.gr
        ElseIf f_site = 444 Then ' montecri
            f_connect = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=vps.montecristo-shop.gr;PORT =3306;DATABASE=montecri_dimos;UID=montecri_lagakis;Password=123abc!@#"
        Else
            f_connect = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=db9.papaki.gr;PORT =3306;DATABASE=lagakis_new;UID=lagakis;Password=39213921"
        End If

104     gdblag.Open f_connect

        '104 gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=toys-shop.gr;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"

        On Error GoTo Form_Load_Err

'Exit Sub

106     EID.ConnectionString = gConnect

        Dim DUM

108     DUM = ADD_FIELD("EID", "MEMOENG", "TEXT")
110     DUM = ADD_FIELD("EID", "PIC0", "VARCHAR(100)")
112     DUM = ADD_FIELD("EID", "PIC1", "VARCHAR(100)")
114     DUM = ADD_FIELD("EID", "PIC2", "VARCHAR(100)")
116     DUM = ADD_FIELD("EID", "PIC3", "VARCHAR(100)")
118     DUM = ADD_FIELD("EID", "PIC4", "VARCHAR(100)")
120     DUM = ADD_FIELD("EID", "PIC5", "VARCHAR(100)")
122     DUM = ADD_FIELD("EID", "PIC6", "VARCHAR(100)")
    
124     DUM = ADD_FIELD("EID", "DIATHESIMO", "INT")

125     DUM = ADD_FIELD("EID", "ID_WEB", "INT")  ' TO ID TOY SITE
126     DUM = ADD_FIELD("EID", "HME_DIATH", "DATETIME")

128     DUM = ADD_FIELD("EID", "MANUFACTURERS_ID", "INT")
130     DUM = ADD_FIELD("EID", "TEXT_PRODUCTS_ZIPCODE", "VARCHAR(30)")

        DUM = ADD_FIELD("EID", "YOU_TUBE", "VARCHAR(100)")

132     DUM = ADD_FIELD("EID", "ADDITIONALPRICE", "DECIMAL(10,2)")
133     DUM = ADD_FIELD("EID", "SHIPPINGPRICE", "DECIMAL(10,2)")

134     DUM = ADD_FIELD("EID", "WEIGHT", "DECIMAL(10,2)")
135     DUM = ADD_FIELD("EID", "WEBPRICE", "DECIMAL(10,2)")

136     DUM = ADD_FIELD("EID", "FYLO", "INT")
138     DUM = ADD_FIELD("EID", "XARAKTIRAS", "INT")
        '' 139 DUM = ADD_FIELD("EID", "CATEGORY", "INT")

        'DIATHESIMO   INTEGER
        'HME_DIATH Date
        'KATASKEYASTHS VarChar(50)
        'TEXT_PRODUCTS_ZIPCODE
        'Indv.Shipping Price:
        'Each Additional Price:
    
        F_FAKEL_KOINOS = FINDPARAMETROI(1, "APOT2", "F_FAKEL_KOINOS", "C:\MERCVB", "KOINOXPHΣTOΣ ΦΑΚΕΛΟΣ ΕΙΚΟΝΩΝ Π.Χ. \\PC\MERCVB")
    
        '
        '  category_parent_id  category_child_id   category_list   category_id vendor_id   category_name
        '0   1   NULL    1   1   Σταθεροί HY
        '0   2   NULL    2   1   Φορητοί HY
        '0   3   NULL    3   1   Περιφερειακά
        '0   6   NULL    6   1   Αναλώσιμα
        '0   19  NULL    19  1   Στικάκια USB
        '0   14  NULL    14  1   Καλώδια
        '2   10  NULL    10  1   NET BOOKS
        '3   22  NULL    22  1   Ποντίκια
        '3   21  NULL    21  1   Εκτυπωτές
        '6   11  NULL    11  1   ΜΕΛΑΝΙΑ-ΤΟΝΕΡ
        '6   9   NULL    9   1   χαρτια
        '9   12  NULL    12  1   ΑΥΤΟΚΟΛΗΤΕΣ ΕΤΙΚΕΤΕΣ
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

        '============================= ΚΑΤΗΓΟΡΙΕΣ =====================================================

        If f_site = 222 Or f_site = 444 Then
            sql = "SELECT a.categories_id as ID,a.parent_id as PARENT,b.CATEGORIES_NAME AS NAME from categories a inner join categories_description b on a.categories_id=b.categories_id order by parent_id,a.categories_id"
        Else  'vm
            sql = "SELECT a.virtuemart_category_id as ID,b.category_parent_id as PARENT,a.category_name as NAME FROM jsnrw_virtuemart_categories_el_gr a inner join jsnrw_virtuemart_category_categories b on a.virtuemart_category_id=b.category_child_id"
        End If
     
        Dim N As Long

142     N = 0

        Dim error As Integer

        'products_id,options_id,options_values_id

        '================================ PRODUCT_ATTRIBUTES ================================
        R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'PRODUCT_ATTRIBUTES'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

        If R(0) = 0 Then
            Gdb.Execute "CREATE TABLE PRODUCT_ATTRIBUTES ( KOD VARCHAR(16) NOT NULL,OPTIONS_VALUES_ID INT,PRODUCT_ID int NOT NULL,OPTIONS_ID  INT  NULL ) "
        End If

        R.Close

        '================================ PRODUCT_TO_CATEGORIES ================================
        R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'PRODUCT_TO_CATEGORIES'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

        If R(0) = 0 Then
            Gdb.Execute "CREATE TABLE PRODUCT_TO_CATEGORIES ( KOD VARCHAR(16) NOT NULL,PRODUCT_ID int NOT NULL,CATEGORIES_ID INT NOT NULL ) "
        End If

        R.Close

144     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'CATEGORIESANT'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

146     If R(0) = 0 Then
148         Gdb.Execute "CREATE TABLE CATEGORIESANT " & "([ID] [int] NOT NULL,[PARENT] [INT] NOT NULL,[NAME] [varCHAR](50) NULL,CONSTRAINT [PK_CATEGORIESANT] PRIMARY KEY CLUSTERED " & "([ID] Asc )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY] "
        End If

150     R.Close

152     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'CATEGTEL'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

154     If R(0) = 0 Then
            'Gdb.Execute "CREATE TABLE CATEGTEL ( [ID] [int] NOT NULL,[PARENT] [INT] NOT NULL,[NAME] [varCHAR](50) NULL,[AA] [INT] NOT NULL )ON PRIMARY  "
156         Gdb.Execute "CREATE TABLE CATEGTEL ( [ID] [int] NOT NULL,[PARENT] [INT] NOT NULL,[NAME] [varCHAR](50) NULL,[AA] [INT] NOT NULL )  "
        End If

157     R.Close

        ''ΕΛΕΓΧΩ ΑΝ ΕΝΗΜΕΡΩΘΗΚΕ ΣΗΜΕΡΑ
        '158 r.Open "SELECT OBJECT_NAME(OBJECT_ID) AS DatabaseName, last_user_update AS HME  FROM sys.dm_db_index_usage_stats  WHERE database_id = DB_ID( 'MERCURY')  AND OBJECT_ID=OBJECT_ID('CATEGORIESANT')", Gdb, adOpenDynamic, adLockOptimistic
        ' If Not r.EOF Then
        '159 If Not IsNull(r!hme) Then
        '160    If Format(r!hme, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
        '         r.Close
        '         GoTo PARAKAMCH
        '    End If
        ' End If
        'End If
        'r.Close
        '
        '
        '

        f_err = 0
161     Gdb.Execute "DELETE FROM CATEGORIESANT"
162     Gdb.Execute "DELETE FROM CATEGTEL"

        Dim cName As String

164     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic

166     Do While Not R.EOF

            If f_site = 444 Then
                cName = Left(greek(R!Name), 50)
                cName = Replace(cName, "'", "''")
            Else
                cName = Replace(R!Name, "'", "''")
            End If
          
168         Gdb.Execute "INSERT INTO CATEGORIESANT (ID,PARENT,NAME) VALUES (" + str(R!ID) + "," + str(R!Parent) + ",'" + cName + "')"
170         R.MoveNext

            If f_err = 1 Then Exit Do
        Loop

172     R.Close

174     Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,1 FROM CATEGORIESANT WHERE PARENT=0"
176     Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,2  FROM CATEGORIESANT WHERE PARENT IN (SELECT ID FROM CATEGTEL)"

        ' GIA NA FYGEI  TO 1&2  EPIPEDO APO TO CATEGORIESANT
178     Gdb.Execute "DELETE FROM CATEGORIESANT WHERE ID IN (SELECT ID FROM CATEGTEL)"

180     Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,3  FROM CATEGORIESANT WHERE PARENT IN (SELECT ID FROM CATEGTEL)"

        ' GIA NA FYGEI  TO 3  EPIPEDO APO TO CATEGORIESANT
182     Gdb.Execute "DELETE FROM CATEGORIESANT WHERE ID IN (SELECT ID FROM CATEGTEL)"

184     Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,4  FROM CATEGORIESANT WHERE PARENT IN (SELECT ID FROM CATEGTEL)"

        ' GIA NA FYGEI  TO 4  EPIPEDO APO TO CATEGORIESANT
186     Gdb.Execute "DELETE FROM CATEGORIESANT WHERE ID IN (SELECT ID FROM CATEGTEL)"

188     Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,5  FROM CATEGORIESANT"

PARAKAMCH:

190     R.Open "SELECT * FROM CATEGTEL ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic
        ' 192 r.MoveFirst

        'r.Open "SELECT * FROM CATEG1", Gdb
194     Do While Not R.EOF
       
196         error = 0
            'paterades
198         mess = "id=" + LTrim(str(R!ID)) + "Parent:" + LTrim(str(R!Parent)) + " " + R!Name

200         If R!Parent = 0 Then
202             T1.Nodes.ADD , , "r" + LTrim(str(R!ID)), R!Name
            Else
    
204             T1.Nodes.ADD "r" + LTrim(str(R!Parent)), tvwChild, "r" + LTrim(str(R!ID)), (R!Name)
            End If
       
206         If error = 0 Then
208             N = N + 1
210             names(N) = "r" + Format(R!ID, "0000") + "        " + R!Name
                FC_id(N) = R!ID
                FC_parenT(N) = R!Parent
                fc_names(N) = R!Name
            End If

212         R.MoveNext
        Loop
     
214     R.Close
        fc_max = N

        '============================= ΚΑΤΗΓΟΡΙΕΣ =====================================================
        ' an einai VM  ΝΑ ΠΑΡΑΚΑΜΠΤΕΙ ΧΑΡΑΚΤΗΡΕΣ,ΦΥΛΟ,ΗΛΙΚΙΕΣ
        If f_site = 111 Or f_site = 444 Then
            kataskeyasths.Visible = False
            xaraktiras.Visible = False
            fylo.Visible = False
            b1.Visible = False
            c1.Visible = False
            d1.Visible = False
            e1.Visible = False
            GoTo 296
        End If
     
        '=========================================χαρακτηρας ============================================
     
'216     sql = "select *from extra_field_values where epf_id=1 and languages_id=4 order by  epf_value"
'        '  (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (7,3,4,0,1,'?????');
'
'        Dim n7 As Long
'
'218     n7 = 0
'220     r.Open sql, gdblag, adOpenDynamic, adLockOptimistic
'
'222     Do While Not r.EOF
'224         xaraktiras.AddItem r!epf_value
'226         f_xaraktiras(n7) = r!Value_ID
'
'
'          '  Debug.Print R(0), R(1), R(2), R(3), R(4), R(5), R(6)
'
'
'228         r.MoveNext
'230         n7 = n7 + 1
'
'            If f_err = 1 Then Exit Do
'        Loop
'
'        xaraktiras.AddItem "                "
'        f_xaraktiras(n7) = 0
'
'232     r.Close
'234     xaraktiras.ListIndex = -1
'
     
     
     XaraktiresTree
     
216          sql = "select * from XARAKTIRESTEL ORDER BY PARENTNAME,AA,NAME"
        '  (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (7,3,4,0,1,'?????');
    
'
'     XaraktiresTree
'      XaraktiresTree
'       XaraktiresTree
       
    
    
    
    
        Dim n7 As Long

218     n7 = 0



'     XaraktiresTree
'      XaraktiresTree
'       XaraktiresTree
'XaraktiresTree
'      XaraktiresTree


220     R.Open sql, Gdb, adOpenDynamic, adLockOptimistic

222     Do While Not R.EOF
224         xaraktiras.AddItem R!Name
226         f_xaraktiras(n7) = R!ID
228         R.MoveNext
230         n7 = n7 + 1
            If f_err = 1 Then
               Exit Do
            End If
               
        Loop

        xaraktiras.AddItem "                "
        f_xaraktiras(n7) = 0
        
232     R.Close
234     xaraktiras.ListIndex = -1
     
     
     
     
     
     
     
     
     
     
     
     

     
     
     
     
        '========================================= FYLO ============================================
236     sql = "select *from extra_field_values where epf_id=3 and languages_id=4 order by epf_value"
        '  (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (7,3,4,0,1,'?????');
238     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic
240     n7 = 0

242     Do While Not R.EOF
244         fylo.AddItem R!epf_value
246         f_fylo(n7) = R!Value_ID
248         n7 = n7 + 1
250         R.MoveNext

            If f_err = 1 Then Exit Do
        Loop

252     R.Close
254     fylo.ListIndex = 0
     
        '========================================= KATASKEYASTHS ============================================
256     sql = "SELECT * FROM manufacturers order by manufacturers_name"
     
258     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic
260     n7 = 0

262     Do While Not R.EOF
264         kataskeyasths.AddItem R!manufacturers_name
266         f_kataskeyasths(n7) = R!manufacturers_id
268         n7 = n7 + 1
270         R.MoveNext

            If f_err = 1 Then Exit Do
        Loop

272     R.Close
274     kataskeyasths.ListIndex = 0
     
        '================================= hlikies ===============================================================================
        'απο εδω διαλέγω τα διατήματα των ηλικιών τα αποια τα αποθηκεύω στα
        '  products_attributes
        ' SELECT * FROM `toyshop`.`products_options_values`;
        'products_options_values_id  language_id products_options_values_name
        '1   4   "0 - 1 ετών"
        '2   4   "1 - 2 ετών"
        '3   4   "2 - 3 ετών"
        '4   4   "3 - 5 ετών"
        '5   4   "5 - 8 ετών"
        '6   4   "8 - 12 ετών"
        '7   4   "12 - 14 ετών"
        '8   4   "14 - 17 ετών"
        '9   4   "18 - 99 ετών"/*
     
276     sql = "SELECT * FROM products_options_values where language_id=4  order by sort_order "   '  15-3-2016  products_options_values_id "
278     n7 = 0
280     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic

282     Do While Not R.EOF
284         hlikies.AddItem R!products_options_values_name
286         f_hlikies(n7) = R!products_options_values_id
288         n7 = n7 + 1
 
290         R.MoveNext

            If f_err = 1 Then Exit Do
        Loop

292     R.Close
294     kataskeyasths.ListIndex = 0
     
296     R.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

        'R.Index = "pinakes"

        'FPA
        'R.Seek "=", 1, 1
298     Do While Not R.EOF

300         If R("typos") = 1 Then
302             Combo2.AddItem str(R("AYJON")) + " -> " + str(R("TIMH"))
            End If

304         R.MoveNext

            If f_err = 1 Then Exit Do
        Loop

        ' mon.metrhshs
306     R.Close
     
        Combo2.ListIndex = 0
     
        If f_err = 1 Then
            MsgBox "Δεν φορτώθηκαν σωστά οι πίνακες."
        End If
     
        ftp.ListIndex = 0
     
        ' Me.Print "Please use now mouse wheel to move this form."
308     Me.SHOW
310     ' ProcessMessages

        Text2.SetFocus

        '<EhFooter>
        Exit Sub
        
NOCONNECT:
        MsgBox "ΑΔΥΝΑΤΗ Η ΣΥΝΔΕΣΗ ΜΕ ΤΟ SITE. ΔΟΚΙΜΑΣΤΕ ΑΡΓΟΤΕΡΑ"
        Unload Me

        Exit Sub

Form_Load_Err:
        '        MsgBox Err.Description & vbCrLf & _
                 "in ADOMERCNEW.apot2vmn.Form_Load " & _
                 "at line " & Erl, _
                 vbExclamation + vbOKOnly, "Application Error"
        f_err = 1

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' bCancel = True
    ' Me.Hide
    On Error Resume Next

    gdblag.Close
 
End Sub

Private Sub Form_Resize()

    Picture2.BorderStyle = 0
    Picture2.Move 0, Picture1.Height + Picture1.Top, Me.ScaleWidth - 250, Me.ScaleHeight * 3
    'Picture2.Move 0, 2000, Me.ScaleWidth - 250, Me.ScaleHeight * 3
    'WebBrowser1.Width = Picture2.Width
 
    Picture1.Width = Picture2.Width
    VScroll1.Left = Picture2.Width

    On Error Resume Next

    VScroll1.Height = Me.Height - 580
 
    VScroll1.Max = Abs(Me.ScaleHeight - Picture2.ScaleHeight)
    VScroll1.LargeChange = VScroll1.Max / 10
    VScroll1.SmallChange = VScroll1.Max / 100
    VScroll1.Min = 0 'Picture1.height

    Command1.Top = Picture2.Height - 700
    Command1.Caption = "Picture2.Height-700"

    Picture2.Move 0, Picture1.Height + Picture1.Top, Me.ScaleWidth - 250, Me.ScaleHeight * 3
End Sub

Private Sub foto2_DblClick(Index As Integer)

        '<EhHeader>
        On Error GoTo FOTO2_DBLClick_Err

        '</EhHeader>
        Dim F_EIKONA As Long

100     F_EIKONA = 11
110     CommonDialog1.FILENAME = "*.jpg"
120     CommonDialog1.ShowOpen

        Dim m As String

        On Error Resume Next
        
130    ' m = Trim(mID(CommonDialog1.FILENAME, InStrRev(CommonDialog1.FILENAME, "\") + 1, 30))
       ' m = "a" + kodlabel.Caption + "-" + m  ' CommonDialog1.FILENAME ' Text1(0).Text
        
        
        Dim cc0 As String
        If Index = 0 Then
           cc0 = ""
        Else
           cc0 = "-" + Format(Index, "0")
        End If
        
        m = toGreekLish(FONO.Text) + "-" + EID.Recordset("ERG") + cc0 + ".JPG"
        
        
        
        
        

140     If F_EIKONA = 11 Then
150         ' m = m + ".JPG"
        End If
        
        If f_site = 111 Then
            '  m = Replace(UCase(m), ".JPG", ".jpeg")
        End If
        
160     FileCopy CommonDialog1.FILENAME, F_FAKEL_KOINOS + "\images\" + m
    
170     foto2(Index).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + m)
            
180     StretchSourcePictureFromPicture apot2vmn.foto2(Index).Picture, apot2vmn.foto2(Index)
        PIC(Index).Caption = m
    
        '<EhFooter>
        Exit Sub

FOTO2_DBLClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command14_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot2vmn.FOTO2_DBLClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub foto2_MouseUp(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          x As Single, _
                          y As Single)

    Dim ANS As Integer

    If Button = 2 Then
        ANS = MsgBox("Να καταργηθεί η εικόνα;", vbYesNo)

        If ANS = vbYes Then
            foto2(Index).Picture = LoadPicture()
            PIC(Index).Caption = ""
        End If
    End If
   
End Sub

Private Sub fylo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub INDV_PRICE_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub kataskeyasths_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub KOD_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then

        KeyAscii = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If
 
End Sub

Private Sub KOD_LostFocus()

    If Len(KOD) > 0 Then

        If Len(EID.RecordSource) > 5 Then
       
            If Not EID.Recordset.EOF Then
                EID.Recordset.Move 0
            End If
    
        End If

        EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + KOD + "'"
        EID.Refresh
   
        If Not EID.Recordset.EOF Then
            UPDATE_CONTROLS
        Else

            Exit Sub

        End If
   
        F_id_LOCAL = EID.Recordset("ID")
    End If

End Sub

Private Sub List1_Click()

    Dim a

    a = Left(List1.List(List1.ListIndex), 15)
   
    KOD.Text = a
    EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + a + "'"
    EID.Refresh
   
    If Not EID.Recordset.EOF Then
        UPDATE_CONTROLS
    Else

        Exit Sub

    End If
    
End Sub

Private Sub List2_Click()

    KOD.Text = Left(List2.List(List2.ListIndex), 13)
  
End Sub

Private Sub LTI5_Change()
    LTI5.Text = Replace$(LTI5.Text, ",", ".")
End Sub

Private Sub LTI5_KeyUp(KeyCode As Integer, Shift As Integer)

100     If KeyCode = vbKeyReturn Then
110         KeyCode = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

End Sub

Private Sub LTI5_LostFocus()
    LTI5.Text = Replace$(LTI5.Text, ",", ".")
    'If Val(WEBPRICE.Text) = 0 Then
    WEBPRICE.Text = LTI5.Text
    'End If

End Sub

Private Sub cmdmovenext_Click()

    On Error Resume Next
  
    If Not EGGTIM2.Recordset.EOF Then
        EGGTIM2.Recordset.MoveNext

        If Not EGGTIM2.Recordset.EOF Then
            CMDmovenext.BackColor = vbGreen
            CMDmoveprevious.BackColor = vbGreen
            KOD.Text = EGGTIM2.Recordset("KODE")
            KOD.SetFocus
            FONO.SetFocus
        Else
            CMDmovenext.BackColor = vbYellow
            'CMDmoveprevious.BackColor = vbgree
        End If
    End If

End Sub

Private Sub cmdmoveprevious_Click()

    On Error Resume Next

    If Not EGGTIM2.Recordset.BOF Then
        EGGTIM2.Recordset.MovePrevious

        If Not EGGTIM2.Recordset.BOF Then
            CMDmovenext.BackColor = vbGreen
            CMDmoveprevious.BackColor = vbGreen
            KOD.Text = EGGTIM2.Recordset("KODE")
            KOD.SetFocus
            FONO.SetFocus
        Else
            CMDmoveprevious.BackColor = vbYellow  ' vbgree
        End If
    
    End If

End Sub

Private Sub neo_Click()
    NEA.Visible = True

    a = "insert into EID (KOD,WEIGHT) VALUES ('" + KOD.Text + "',1)"
 
    ' neo.Enabled = True
    ' Exit Sub
    Dim R As New ADODB.Recordset

    '"   στα νέα είδη αυτόματα να βάζει default βάρος=1κιλό & status=διαθέσιμο

    R.Open "SELECT * FROM EID WHERE KOD='" + KOD.Text + "'", Gdb, adOpenDynamic, adLockOptimistic

    If Not R.EOF() Then
        MsgBox "Υπάρχει ο κωδικός με το όνομα " + "" & R!ONO
        neo.Enabled = True
        KOD.SetFocus
    Else
        Gdb.Execute a
        Gdb.Execute "INSERT INTO BARCODES (KOD,ERG) VALUES ('" + KOD.Text + "','" + KOD.Text + "')"
        'WEIGHT.Text = "1"
        Check1.Value = vbChecked

    End If

    EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + KOD.Text + "'"
    EID.Refresh
   
    Dim k As Integer

    For k = 0 To 6
        foto2(k).Picture = LoadPicture()
    Next
   
    For k = 0 To hlikies.ListCount - 1
        'If hlikies.Selected(k) Then
        hlikies.Selected(k) = False
    Next
   
    F_id_LOCAL = EID.Recordset("ID")
   
    neo.Enabled = False
    KOD.Enabled = False
    ONO.Enabled = False
    upload1.Enabled = False
 
End Sub

Private Sub ONO_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyAscii = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub ONO_LostFocus()

    If Len(ONO.Text) = 0 Then

        Exit Sub

    End If

    Dim R As New ADODB.Recordset

    R.Open "select TOP 100 * from EID WHERE ONO LIKE '" + ONO.Text + "%'", Gdb, adOpenDynamic, adLockOptimistic
    List1.Clear

    Do While Not R.EOF
        List1.AddItem Left(R("KOD") + Space(15), 15) + R("ONO")
        R.MoveNext
        ' Exit Do
    Loop

    R.Close
  
End Sub

Private Sub oristiki_Click()

ANS = MsgBox("Nα οριστικοποιηθούν οι " + str(fmaxRec) + " εγγραφές;", vbYesNo)



 gdblag.Close
    gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=176.9.148.246;PORT =3306;DATABASE=toyshop;UID=dtek;Password=D35m80)!"


If ANS = vbNo Then
    gdblag.Execute "update products  set web_updated=1 where web_updated=3  "
    MsgBox "Ακυρώθηκε η οριστικοποίηση"
Else
    gdblag.Execute "update products  set web_updated=0 where web_updated=3  "
    MsgBox "ok"
End If





End Sub

Private Sub Picture2_DblClick()
     PARAMETROI.PARAM.Caption = "APOT2VMN"
     PARAMETROI.SHOW 1
End Sub

Private Sub T1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim y

    Dim nam

    If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty
        y = T1.SelectedItem.Index
        'Label1.Caption = Str(y)
        nam = T1.SelectedItem.Text
        Me.Caption = nam & " - " & str(y) + " names=" + names(y)
        Set T1.DropHighlight = T1.SelectedItem
    End If

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then

        KeyAscii = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub Text2_LostFocus()

    Dim R As New ADODB.Recordset

    If Len(Text2.Text) = 0 Then

        Exit Sub

    End If

    If Left(Text2.Text, 1) = "/" Then
        Text2.Text = Trim(mID(Text2.Text, 2, 30))
        R.Open "select TOP 100 * from EID WHERE ERG LIKE '" + Text2.Text + "%'", Gdb, adOpenDynamic, adLockOptimistic
        List1.Clear

        Do While Not R.EOF
            List1.AddItem Left(R("KOD") + Space(15), 15) + R("ONO")
            R.MoveNext
            ' Exit Do
        Loop

        R.Close

        Exit Sub

    Else
        R.Open "select * from EID WHERE KOD IN ( SELECT TOP 1 KOD FROM BARCODES WHERE ERG='" + Text2.Text + "')", Gdb, adOpenDynamic, adLockOptimistic
    End If
 
    If Not R.EOF Then
        KOD.Text = R(0)
        KOD.SetFocus
        FONO.SetFocus
        'UPDATE_CONTROLS
    Else

        Exit Sub

    End If
   
    F_id_LOCAL = R("ID")

End Sub

Private Sub upload1_Click()

    If f_site = 222 Then
        mpoyg_upload
   
    Else
        VM_upload
    End If

    On Error Resume Next

    If CMDmovenext.BackColor = vbGreen Then
        If Not EGGTIM2.Recordset.EOF Then
            EGGTIM2.Recordset.MoveNext

            If Not EGGTIM2.Recordset.EOF Then
                CMDmovenext.BackColor = vbGreen
                CMDmoveprevious.BackColor = vbGreen
                KOD.Text = EGGTIM2.Recordset("KODE")
                KOD.SetFocus
                
                FONO.SetFocus
            Else
                CMDmovenext.BackColor = vbYellow
                'CMDmoveprevious.BackColor = vbgree
            End If
        End If
    End If

End Sub

Sub mpoyg_upload()

    '=====================================================================================================================
    '=====================================================================================================================
    Dim a As String

    Dim R As New ADODB.Recordset

    Dim N As Long

    Dim m_ID

    Dim m_categ

    Dim m_new As Integer

    m_new = 0

    '             mpoyg_upload_Err:
    On Error GoTo mpoyg_upload_Err
 
    If FONO = "" Then
        MsgBox "Δεν επιλέξατε είδος "

        Exit Sub

    End If
 
    Dim k   As Integer

    Dim sql As String

    m_categ = "942"
 
    If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty
        y = T1.SelectedItem.Index
        nam = T1.SelectedItem.Text
        m_categ = mID(names(y), 2, 4)
        Set T1.DropHighlight = T1.SelectedItem
    Else
        MsgBox "Δεν επιλέξατε κατηγορία "

        Exit Sub

    End If
 
    EID.Recordset.Move 0  ' gia na ginei to rich1 opos to mseditor21
 
    a = RichToHTML(Rich1, 0&, Len(Rich1.Text))  '(.TextRTF)
    a = Replace(a, "'", "`")

    gdblag.Close
    gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=176.9.148.246;PORT =3306;DATABASE=toyshop;UID=dtek;Password=D35m80)!"   ' www.toys-shop.gr
    R.Open "select * from products where products_model='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

    M_DATE = Format(Now, "YYYY-MM-DD")

    If R.EOF Then
        m_new = 1
        sql = "INSERT INTO products (products_youtube,products_quantity,products_model,products_image,products_subimage1,"
        sql = sql + " products_subimage2,products_subimage3,products_subimage4,products_subimage5,products_subimage6,products_price,"
        sql = sql + " products_date_added,products_last_modified,products_date_available,products_weight,products_status,products_tax_class_id,"
        sql = sql + " manufacturers_id,products_ordered) VALUES ('',6,'" + KOD + "','', "
        sql = sql + " NULL,NULL,NULL,NULL,NULL,NULL,'','" + M_DATE + "','" + M_DATE + "',NULL,'0.20',1,1,109,0); "

        ' Exit Sub
        gdblag.Execute sql, N
        R.Close
        R.Open "select * from products where products_model='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
    End If

    Dim site_foto(6) As String

    site_foto(0) = CNull(R!products_image)
    site_foto(1) = CNull(R!products_subimage1)
    site_foto(2) = CNull(R!products_subimage2)
    site_foto(3) = CNull(R!products_subimage3)
    site_foto(4) = CNull(R!products_subimage4)
    site_foto(5) = CNull(R!products_subimage5)
    site_foto(6) = CNull(R!products_subimage6)

    'Exit Sub

    m_ID = str(R("products_id"))
    R.Close

    Dim b As String

    ' m_ID = "30917"

    ' extra_value_id1 χαρακτηρας θεμα
    'b = xaraktiras.List(xaraktiras.ListIndex)

    If xaraktiras.ListIndex = -1 Then
        b = "0"
    Else
        b = str(f_xaraktiras(xaraktiras.ListIndex))
    End If
 
    ' extra_value_id3 φυλο
 
    Dim mTimh

    Dim h     As Single
  
    Dim kor   As String

    Dim manuf As String
 
    kor = str(f_fylo(fylo.ListIndex))
  
    manuf = str(f_kataskeyasths(kataskeyasths.ListIndex))
  
    h = (100 + g_Fpa(Val(Left(Combo2.Text, 2)))) / 100
 
    mTimh = Format(Val(WEBPRICE.Text) / h, "####0.0000")

    'ΚΑΤΑΣΚΕΥΑΣΤΗΣ+ΧΑΡΑΚΤΗΡΑΣ+ ONO+' #'+ERG

    If Len(Trim(kataskeyasths.Text)) = 0 Then
        m_erg = ""
    Else

        If KATASK.Value = vbChecked Then
            m_erg = kataskeyasths.Text + " "
        Else
            m_erg = ""
        End If
  
    End If

    If Len(Trim(xaraktiras.Text)) = 0 Then
  
    Else

        If xarakt.Value = vbChecked Then
            m_erg = m_erg + xaraktiras.Text + " "
        End If
    End If

    m_erg = m_erg + Replace(FONO.Text, "'", "`")

    If IsNull(EID.Recordset("ERG")) Then
    Else

        If Len(EID.Recordset("ERG")) > 0 Then
            m_erg = m_erg + " #" + EID.Recordset("ERG")
        End If
  
    End If

    'm_new = 1
    If m_new = 1 Then
        gdblag.Execute "insert into products_description(language_id,products_id,extra_value_id3,extra_value_id1,products_name,products_description) values (4," + m_ID + "," + kor + "," + b + ",'" + m_erg + "','" + Replace(a, "'", "`") + "')", N
    Else

        R.Open "select * from  products_description  where products_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic

        If R.EOF Then
            gdblag.Execute "insert into products_description(language_id,products_id,extra_value_id3,extra_value_id1,products_name,products_description) values (4," + m_ID + "," + kor + "," + b + ",'" + m_erg + "','" + Replace(a, "'", "`") + "')", N
        Else
            gdblag.Execute "update products_description set  extra_value_id3=" + kor + ",extra_value_id1=" + b + " where products_id=" + m_ID, N
            gdblag.Execute "update products_description set  products_name='" + m_erg + "',products_description='" + Replace(a, "'", "`") + "' where products_id=" + m_ID, N
        End If

        R.Close
   
        ' ", products_name='" + FONO.Text + "', products_description='" + a + "'
    End If

    If N = 0 Then
        MsgBox "ΠΡΟΣΟΧΗ ΠΡΟΒΛΗΜΑ ΕΝΗΜΕΡΩΣΗΣ." + Chr(13) + " ΑΠΑΙΤΕΙΤΑΙ ΕΛΕΓΧΟΣ"
        ' Exit Sub

    End If

    Dim energo As String

    energo = "0"

    If Check1.Value = vbChecked Then
        energo = "1"
    End If


    gdblag.Execute "update products set erg_code = '" + EID.Recordset("ERG") + "'  where products_id=" + m_ID, N
    
   
    gdblag.Execute "update products set  web_updated = 0  WHERE products_id = " + m_ID, N
    
    If EID.Recordset("has_assortment") = 1 Then
       gdblag.Execute "update products set  has_assortment = 1  WHERE products_id = " + m_ID, N
    Else
       gdblag.Execute "update products set  has_assortment = 0  WHERE products_id = " + m_ID, N
    End If
    
    
    gdblag.Execute "update products set products_status=" + energo + ",manufacturers_id=" + manuf + ",products_price=" + Replace(mTimh, ",", ".") + ",products_weight=" + str(Val(WEIGHT)) + "  where products_id=" + m_ID, N
    gdblag.Execute "update products set products_youtube='" + YOU_TUBE.Text + "',products_image='" + PIC(0).Caption + "',products_subimage1='" + PIC(1).Caption + "',products_subimage2='" + PIC(2).Caption + "' where products_id=" + m_ID, N
    gdblag.Execute "update products set products_subimage3='" + PIC(3).Caption + "',products_subimage4='" + PIC(4).Caption + "',products_subimage5='" + PIC(5).Caption + "',products_subimage6='" + PIC(6).Caption + "' where products_id=" + m_ID, N

    'products_ship_price_two     products_ship_price                 products_ship_zip
    'ADDITIONALPRICE              SHIPPINGPRICE                         TEXT_PRODUCTS_ZIPCODE      YOU_TUBE  WEBPRICE
    'ADDIT_PRICE.text              INDV_PRICE.text                      ZIP_CODE.text
    R.Open "select * from products_shipping where products_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic

    If R.EOF Then
        gdblag.Execute "INSERT INTO products_shipping (products_id,products_ship_methods_id,products_ship_zip,products_ship_price,products_ship_price_two) VALUES (" + m_ID + ",0,'" + ZIP_CODE.Text + "','" + INDV_PRICE.Text + "','" + ADDIT_PRICE.Text + "');", N
    Else
        gdblag.Execute "update products_shipping set  products_ship_price_two='" + ADDIT_PRICE.Text + "',products_ship_price='" + INDV_PRICE.Text + "',products_ship_zip='" + ZIP_CODE.Text + "' where products_id=" + m_ID, N
    End If

    R.Close

    gdblag.Execute "delete from products_to_categories where products_id=" + m_ID + " limit 2"

    If R.EOF Then
        gdblag.Execute "INSERT INTO products_to_categories (products_id,categories_id) VALUES (" + m_ID + "," + m_categ + ");", N
    End If

    ' R.Close

    Dim FPA As String

    '**********  fpa tax class ******************************
    'PRODUCTS_tax_class_ID=1 =>  23%
    'PRODUCTS_tax_class_ID=2 => 6.5%
    'PRODUCTS_tax_class_ID=3 => 13%
 
    FPA = "1"

    If Val(Left(Combo2.Text, 2)) = 1 Then
        FPA = "3"
    ElseIf Val(Left(Combo2.Text, 2)) = 2 Then
        FPA = "1"
    ElseIf Val(Left(Combo2.Text, 2)) = 4 Then
        FPA = "2"
    End If
 
    '  gdblag.Close
    'gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131 ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
    gdblag.Execute "update products set products_tax_class_id=" + FPA + " where products_id=" + m_ID, N

    '----------------------------------  CATEGORIES --------------------------------------------------------------------------------------------
    R.Open "select count(*) from products_to_categories WHERE products_id=" + m_ID + " AND categories_id=" + m_categ + ";", gdblag, adOpenDynamic, adLockOptimistic

    If R(0) > 0 Then
        ' ok yparxei den kano tipota
    Else
        gdblag.Execute "insert into  products_to_categories ( products_id,categories_id) values (" + m_ID + "," + m_categ + ") ;", N
    End If

    '----------------------------------  CATEGORIES --------------------------------------------------------------------------------------------

    gdblag.Execute "DELETE FROM products_attributes where products_id=" + m_ID + "  limit 9", k, N

    Dim sort As Integer

    sort = 0

    For k = 0 To hlikies.ListCount - 1

        If hlikies.Selected(k) Then
            gdblag.Execute "INSERT INTO products_attributes " & "(products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) " & " VALUES (" + m_ID + ",1," + str(f_hlikies(k)) + ",'0.0000',''," + str(sort) + ");", N
            sort = sort + 1
        End If

    Next k

    'Server Access type
    ' INET1.AccessType = icUseDefault
    'Protocol to be used
    ' INET1.protocol = icFTP
    'Remote host name
    ' INET1.RemoteHost = "www.toys-shop.gr"
    'Server port number, usually its 21
    ' INET1.RemotePort = "21"
    'Server password
    ' INET1.Password = "T35m80)!@"
    'Server Username
    ' INET1.UserName = "toyshopftp"
    'Server session timeout
    ' INET1.RequestTimeout = "60"

    file = 3

    Dim nc As Integer

    nc = 0

    '==============================================================================
    If ftp.Text = ftp.List(1) Then   ' inet

        '    For k = 0 To 6
        '       Me.Caption = "upload foto" + str(k)
        '       If Len(PIC(k).Caption) > 0 Then
        '
        '          If Len(site_foto(k)) > 1 Then
        '              Inet1.Execute , "DELETE " + Replace(site_foto(k), "*", "-") 'βγαζω το αστερακι μην κανει ζημια μεγαλη
        '          End If
        '
        '             Inet1.Execute , "PUT  """ + F_FAKEL_KOINOS + PIC(k).Caption + """" & "images/stories/virtuemart/product/" + PIC(k).Caption
        '
        '          milsecs = GetCurrentTime()
        '          Do While Inet1.StillExecuting = True
        '              If GetCurrentTime() - milsecs > 15000 Then Exit Do
        '              DoEvents
        '          Loop
        '          nc = nc + 1
        '       End If
        '    Next
    Else 'dos & api  ========================================================================

      If Len(Dir("C:\MERCVB\MAC.TXT", vbNormal)) > 0 Then  ' macintosh
      
          '===================================
          For k = 0 To 6
            Me.Caption = "upload foto" + str(k)
            If Len(PIC(k).Caption) > 0 Then
                'την εχει ηδη την φωτογραφια  den kanv tipota
                If site_foto(k) = PIC(k).Caption And reload.Value = vbUnchecked Then
                Else
                    Shell ("c:\windows\NCFTPPUT.exe  -u toyshopftp -p  %t6Ixj43  176.9.148.246 /httpdocs/images/ " + F_FAKEL_KOINOS + "\images\" + PIC(k).Caption)
                End If
                nc = nc + 1
            End If
          Next
          '===================================
       Else

        Open "c:\mercvb\send_foto" For Output As #file
        Print #file, "open 176.9.148.246"
        Print #file, "toyshopftp"
        Print #file, "%t6Ixj43"                    ' "T35m80)!@"
        Print #file, "binary"
        Print #file, "cd httpdocs/images"

        For k = 0 To 6
            Me.Caption = "upload foto" + str(k)

            If Len(PIC(k).Caption) > 0 Then
     
                'την εχει ηδη την φωτογραφια  den kanv tipota
                If site_foto(k) = PIC(k).Caption And reload.Value = vbUnchecked Then
                Else

                    If Len(site_foto(k)) > 1 Then
                        Print #file, "del " + Replace(site_foto(k), "*", "-")  'βγαζω το αστερακι μην κανει ζημια μεγαλη
                    End If

                    If ftp.Text = ftp.List(2) Then
                        ftp_Eikones F_FAKEL_KOINOS + "\images\" + PIC(k).Caption, "httpdocs/images/" + PIC(k).Caption
                    Else
                        Print #file, "send " + F_FAKEL_KOINOS + "\images\" + PIC(k).Caption
                    End If
        
                End If
      
                nc = nc + 1
            End If

        Next

        Print #file, "quit"
        Close #file

        If nc > 0 Then
            Shell ("ftp -s:c:\mercvb\send_foto")
        End If
        
      End If
        

    End If  '=================================================================================

    Dim surl As String

'    surl = "http://176.9.148.246/admin/categories.php?cPath=103_326&pID=" + m_ID + "&action=new_product"
'
'
'
'    WebBrowser1.navigate surl
'
'    Dim msec As Long
'
'    msec = GetCurrentTime()
'
'    Do While Not WebBrowser1.readyState = READYSTATE_COMPLETE
'        DoEvents
'
'        If GetCurrentTime() - msec > 3000 Then
'
'            Exit Do
'
'        End If
'
'    Loop

    
    Exit Sub

mpoyg_upload_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.APOT2.Command2_Click " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command2_Click " & "at line " & Erl

    Resume Next

    '</EhFooter>

    Exit Sub

End Sub 'P@ssw0rd#!

Sub mpoyg_download()

    '=====================================================================================================================
    'κατεβασμα των ειδών  μπουγουδη
    '=====================================================================================================================
    Dim a As String

    Dim R As New ADODB.Recordset

    Dim N As Long

    Dim m_ID

    Dim m_categ

    Dim m_new As Integer

    m_new = 0

    '             mpoyg_upload_Err:
    On Error GoTo mpoyg_upload_Err
 
    If FONO = "" Then
        MsgBox "Δεν επιλέξατε είδος "

        Exit Sub

    End If
 
    Dim k   As Integer

    Dim sql As String

    m_categ = "942"
 
    If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty
        y = T1.SelectedItem.Index
        nam = T1.SelectedItem.Text
        m_categ = mID(names(y), 2, 4)
        Set T1.DropHighlight = T1.SelectedItem
    Else
        MsgBox "Δεν επιλέξατε κατηγορία "

        Exit Sub

    End If
 
    EID.Recordset.Move 0  ' gia na ginei to rich1 opos to mseditor21
 
    a = RichToHTML(Rich1, 0&, Len(Rich1.Text))  '(.TextRTF)
    a = Replace(a, "'", "''")

    gdblag.Close
    gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=176.9.148.246;PORT =3306;DATABASE=toyshop;UID=dtek;Password=D35m80)!"  'www.toys-shop.gr
    R.Open "select * from products where products_model='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

    M_DATE = Format(Now, "YYYY-MM-DD")

    If R.EOF Then
        m_new = 1
        sql = "INSERT INTO products (products_youtube,products_quantity,products_model,products_image,products_subimage1,"
        sql = sql + " products_subimage2,products_subimage3,products_subimage4,products_subimage5,products_subimage6,products_price,"
        sql = sql + " products_date_added,products_last_modified,products_date_available,products_weight,products_status,products_tax_class_id,"
        sql = sql + " manufacturers_id,products_ordered) VALUES ('',6,'" + KOD + "','', "
        sql = sql + " NULL,NULL,NULL,NULL,NULL,NULL,'','" + M_DATE + "','" + M_DATE + "',NULL,'0.20',1,1,109,0); "

        ' Exit Sub
        gdblag.Execute sql, N
        R.Close
        R.Open "select * from products where products_model='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
    End If

    Dim site_foto(6) As String

    site_foto(0) = CNull(R!products_image)
    site_foto(1) = CNull(R!products_subimage1)
    site_foto(2) = CNull(R!products_subimage2)
    site_foto(3) = CNull(R!products_subimage3)
    site_foto(4) = CNull(R!products_subimage4)
    site_foto(5) = CNull(R!products_subimage5)
    site_foto(6) = CNull(R!products_subimage6)

    'Exit Sub

    m_ID = str(R("products_id"))
    R.Close

    Dim b As String

    ' m_ID = "30917"

    ' extra_value_id1 χαρακτηρας θεμα
    'b = xaraktiras.List(xaraktiras.ListIndex)

    If xaraktiras.ListIndex = -1 Then
        b = "0"
    Else
        b = str(f_xaraktiras(xaraktiras.ListIndex))
    End If
 
    ' extra_value_id3 φυλο
 
    Dim mTimh

    Dim h     As Single
  
    Dim kor   As String

    Dim manuf As String
 
    kor = str(f_fylo(fylo.ListIndex))
  
    manuf = str(f_kataskeyasths(kataskeyasths.ListIndex))
  
    h = (100 + g_Fpa(Val(Left(Combo2.Text, 2)))) / 100
 
    mTimh = Format(Val(WEBPRICE.Text) / h, "####0.0000")

    'ΚΑΤΑΣΚΕΥΑΣΤΗΣ+ΧΑΡΑΚΤΗΡΑΣ+ ONO+' #'+ERG

    If Len(Trim(kataskeyasths.Text)) = 0 Then
        m_erg = ""
    Else

        If KATASK.Value = vbChecked Then
            m_erg = kataskeyasths.Text + " "
        Else
            m_erg = ""
        End If
  
    End If

    If Len(Trim(xaraktiras.Text)) = 0 Then
  
    Else
        m_erg = m_erg + xaraktiras.Text + " "
    End If

    m_erg = m_erg + Replace(FONO.Text, "'", "''")

    If IsNull(EID.Recordset("ERG")) Then
    Else
        m_erg = m_erg + " #" + EID.Recordset("ERG")
    End If

    'm_new = 1
    If m_new = 1 Then

        gdblag.Execute "insert into products_description(language_id,products_id,extra_value_id3,extra_value_id1,products_name,products_description) values (4," + m_ID + "," + kor + "," + b + ",'" + m_erg + "','" + Replace(a, "'", "''") + "')", N
    Else

        R.Open "select * from  products_description  where products_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic

        If R.EOF Then
            gdblag.Execute "insert into products_description(language_id,products_id,extra_value_id3,extra_value_id1,products_name,products_description) values (4," + m_ID + "," + kor + "," + b + ",'" + m_erg + "','" + Replace(a, "'", "''") + "')", N
        Else
            gdblag.Execute "update products_description set  extra_value_id3=" + kor + ",extra_value_id1=" + b + ",products_name='" + m_erg + "',products_description='" + Replace(a, "'", "''") + "' where products_id=" + m_ID, N
        End If

        R.Close
   
        ' ", products_name='" + FONO.Text + "', products_description='" + a + "'
    End If

    Dim energo As String

    energo = "0"

    If Check1.Value = vbChecked Then
        energo = "1"
    End If

    gdblag.Execute "update products set products_status=" + energo + ",manufacturers_id=" + manuf + ",products_price=" + Replace(mTimh, ",", ".") + ",products_weight=" + str(Val(WEIGHT)) + "  where products_id=" + m_ID, N
    gdblag.Execute "update products set products_youtube='" + YOU_TUBE.Text + "',products_image='" + PIC(0).Caption + "',products_subimage1='" + PIC(1).Caption + "',products_subimage2='" + PIC(2).Caption + "' where products_id=" + m_ID, N
    gdblag.Execute "update products set products_subimage3='" + PIC(3).Caption + "',products_subimage4='" + PIC(4).Caption + "',products_subimage5='" + PIC(5).Caption + "',products_subimage6='" + PIC(6).Caption + "' where products_id=" + m_ID, N

    'products_ship_price_two     products_ship_price                 products_ship_zip
    'ADDITIONALPRICE              SHIPPINGPRICE                         TEXT_PRODUCTS_ZIPCODE      YOU_TUBE  WEBPRICE
    'ADDIT_PRICE.text              INDV_PRICE.text                      ZIP_CODE.text
    R.Open "select * from products_shipping where products_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic

    If R.EOF Then
        gdblag.Execute "INSERT INTO products_shipping (products_id,products_ship_methods_id,products_ship_zip,products_ship_price,products_ship_price_two) VALUES (" + m_ID + ",0,'" + ZIP_CODE.Text + "','" + INDV_PRICE.Text + "','" + ADDIT_PRICE.Text + "');", N
    Else
        gdblag.Execute "update products_shipping set  products_ship_price_two='" + ADDIT_PRICE.Text + "',products_ship_price='" + INDV_PRICE.Text + "',products_ship_zip='" + ZIP_CODE.Text + "' where products_id=" + m_ID, N
    End If

    R.Close

    gdblag.Execute "delete from products_to_categories where products_id=" + m_ID + " limit 2"

    If R.EOF Then
        gdblag.Execute "INSERT INTO products_to_categories (products_id,categories_id) VALUES (" + m_ID + "," + m_categ + ");", N
    End If

    ' R.Close

    Dim FPA As String

    '**********  fpa tax class ******************************
    'PRODUCTS_tax_class_ID=1 =>  23%
    'PRODUCTS_tax_class_ID=2 => 6.5%
    'PRODUCTS_tax_class_ID=3 => 13%
 
    FPA = "1"

    If Val(Left(Combo2.Text, 2)) = 1 Then
        FPA = "3"
    ElseIf Val(Left(Combo2.Text, 2)) = 2 Then
        FPA = "1"
    ElseIf Val(Left(Combo2.Text, 2)) = 4 Then
        FPA = "2"
    End If
 
    '  gdblag.Close
    'gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131 ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
    gdblag.Execute "update products set products_tax_class_id=" + FPA + " where products_id=" + m_ID, N

    '----------------------------------  CATEGORIES --------------------------------------------------------------------------------------------
    R.Open "select count(*) from products_to_categories WHERE products_id=" + m_ID + " AND categories_id=" + m_categ + ";", gdblag, adOpenDynamic, adLockOptimistic

    If R(0) > 0 Then
        ' ok yparxei den kano tipota
    Else
        gdblag.Execute "insert into  products_to_categories ( products_id,categories_id) values (" + m_ID + "," + m_categ + ") ;", N
    End If

    '----------------------------------  CATEGORIES --------------------------------------------------------------------------------------------

    gdblag.Execute "DELETE FROM products_attributes where products_id=" + m_ID + "  limit 9", k, N

    Dim sort As Integer

    sort = 0

    For k = 0 To hlikies.ListCount - 1
  
        If hlikies.Selected(k) Then
      
            gdblag.Execute "INSERT INTO products_attributes " & "(products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) " & " VALUES (" + m_ID + ",1," + str(f_hlikies(k)) + ",'0.0000',''," + str(sort) + ");", N
            sort = sort + 1
        End If

    Next k

    'Server Access type
    ' INET1.AccessType = icUseDefault
    'Protocol to be used
    ' INET1.protocol = icFTP
    'Remote host name
    ' INET1.RemoteHost = "www.toys-shop.gr"
    'Server port number, usually its 21
    ' INET1.RemotePort = "21"
    'Server password
    ' INET1.Password = "T35m80)!@"
    'Server Username
    ' INET1.UserName = "toyshopftp"
    'Server session timeout
    ' INET1.RequestTimeout = "60"

    file = 3

    Dim nc As Integer

    nc = 0

    '==============================================================================
    If ftp.Text = ftp.List(1) Then   ' inet

        '    For k = 0 To 6
        '       Me.Caption = "upload foto" + str(k)
        '       If Len(PIC(k).Caption) > 0 Then
        '
        '          If Len(site_foto(k)) > 1 Then
        '              Inet1.Execute , "DELETE " + Replace(site_foto(k), "*", "-") 'βγαζω το αστερακι μην κανει ζημια μεγαλη
        '          End If
        '
        '             Inet1.Execute , "PUT  """ + F_FAKEL_KOINOS + PIC(k).Caption + """" & "images/stories/virtuemart/product/" + PIC(k).Caption
        '
        '          milsecs = GetCurrentTime()
        '          Do While Inet1.StillExecuting = True
        '              If GetCurrentTime() - milsecs > 15000 Then Exit Do
        '              DoEvents
        '          Loop
        '          nc = nc + 1
        '       End If
        '    Next
    Else 'dos & api  ========================================================================

        Open "c:\mercvb\send_foto" For Output As #file
        Print #file, "open www.toys-shop.gr"
        Print #file, "toyshopftp"
        Print #file, "T35m80)!@"
        Print #file, "binary"
        Print #file, "cd httpdocs/images"

        For k = 0 To 6
            Me.Caption = "upload foto" + str(k)

            If Len(PIC(k).Caption) > 0 Then
     
                'την εχει ηδη την φωτογραφια  den kanv tipota
                If site_foto(k) = PIC(k).Caption And reload.Value = vbUnchecked Then
                Else

                    If Len(site_foto(k)) > 1 Then
                        Print #file, "del " + Replace(site_foto(k), "*", "-")  'βγαζω το αστερακι μην κανει ζημια μεγαλη
                    End If

                    If ftp.Text = ftp.List(2) Then
                        ftp_Eikones F_FAKEL_KOINOS + "\images\" + PIC(k).Caption, "httpdocs/images/" + PIC(k).Caption
                    Else
                        Print #file, "send " + F_FAKEL_KOINOS + "\images\" + PIC(k).Caption
                    End If
        
                End If
      
                nc = nc + 1
            End If

        Next

        Print #file, "quit"
        Close #file

        If nc > 0 Then
            Shell ("ftp -s:c:\mercvb\send_foto")
        End If

    End If  '=================================================================================

    Dim surl As String

    surl = "http://www.toys-shop.gr/admin/categories.php?cPath=103_326&pID=" + m_ID + "&action=new_product"

    'BARMAIN.Value = 100

    WebBrowser1.navigate surl

    Dim msec As Long

    msec = GetCurrentTime()

    Do While Not WebBrowser1.readyState = READYSTATE_COMPLETE
        DoEvents

        If GetCurrentTime() - msec > 3000 Then

            Exit Do

        End If

    Loop

    'http://www.lagakis.gr/administrator/index.php?option=com_virtuemart&view=product&task=edit&virtuemart_product_id=149&product_parent_id=0

    MsgBox "OK"

    Exit Sub

mpoyg_upload_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.APOT2.Command2_Click " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command2_Click " & "at line " & Erl

    Resume Next

    '</EhFooter>

    Exit Sub

End Sub

Sub ftp_Eikones(strLocal As String, strRemote As String)

    Dim lTimer As Long

    'Dim strRemote As String

    'Dim strLocal  As String
    ' strRemote = "/httpdocs/images"  '             Text4.Text & "/" & Text5.Text
    ' strLocal = "c:\mercvb\images"   ' App.Path & "\Updates\" & Text5.Text

    Dim m2ftp  As New cFTP
    
    If m2ftp.OpenConnection("toys-shop.gr", "toysftp", "%t6Ixj43") Then
   
        'If m2ftp.OpenConnection("users.otenet.gr", "lagakis", "a8417!") Then
        m2ftp.SetFTPDirectory "/"

        If Not m2ftp.FTPUploadFile(strLocal, strRemote) Then
            Me.Caption = "Status: Error"
            MsgBox m2ftp.GetLastErrorMessage
        Else
            Me.Caption = "Status: Upload Complete"

            DoEvents

            ' RunUpdate App.Path & "\Updates\" & Text5.Text
            DoEvents
        End If

        DoEvents
        m2ftp.CloseConnection
    End If

End Sub

Sub VM_upload()

    '=====================================================================================================================
    '=====================================================================================================================
    Dim a As String

    Dim R As New ADODB.Recordset
 
    Dim m_ID

    Dim m_categ

    Dim m_new As Integer

    m_new = 0
 
    If FONO = "" Then
        MsgBox "Δεν επιλέξατε είδος "

        Exit Sub

    End If
 
    Dim k   As Integer

    Dim sql As String

    m_categ = "942"
 
    If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty
        y = T1.SelectedItem.Index
        nam = T1.SelectedItem.Text
        m_categ = mID(names(y), 2, 4)
        Set T1.DropHighlight = T1.SelectedItem
    Else
        MsgBox "Δεν επιλέξατε κατηγορία "

        Exit Sub

    End If

    'id=7  category=9
    'INSERT INTO `jsnrw_virtuemart_products` (virtuemart_vendor_id,product_parent_id,product_sku,product_weight,  ' (1,0,'12346',NULL,
    'product_weight_uom,product_length,product_width,product_height,product_lwh_uom,product_url,product_in_stock, '  'TEM' , NULL , NULL , NULL , 'M','',1
    ' product_ordered,  low_stock_notification,
    ' product_available_date , product_availability,          '2013-11-03 00:00:00'
    ' product_special,
    'product_sales,  product_unit,  product_packaging,   product_params,  hits,  intnotes,  '0,'KG',NULL, min_order_level=""|max_order_level=""|step_order_level=""|product_box=""|',NULL,''
    ' metarobot,  metaauthor,  layout, published, pordering,   '','','0',1,0,
    'created_on,  created_by,  modified_on,  modified_by,  locked_on,  locked_by) VALUES  ' '2013-11-03 16:59:59',42,'2013-11-03 16:59:59',42,'0000-00-00 00:00:00',0);
    '(1,0,'12346',NULL,
    ''TEM',NULL,NULL,NULL,'M','',
    '0,0,0,'2013-11-03 00:00:00','',
    '   0, '','','0',1,0,
    '  '','','0',1,0,
    '  '2013-11-03 16:59:59',42,'2013-11-03 16:59:59',42,'0000-00-00 00:00:00',0);
 
    ' virtuemart_product_id   virtuemart_vendor_id    product_parent_id   product_sku product_weight  product_weight_uom  product_length  product_width   product_height  product_lwh_uom   product_url    product_in_stock    product_ordered  low_stock_notification       product_available_date  product_availability    product_special   product_sales   product_unit    product_packaging    product_params                                                                         hits                     intnotes    metarobot   metaauthor  layout      published          pordering       created_on         created_by    modified_on                  modified_by         locked_on                 locked_by
    '       1                      1                        0               12345           NULL              KG              NULL              NULL            NULL              M             ""                1              0                   0                      "2013-11-03 00:00:00"            1-2m.gif                1          0                    KG                NULL          "min_order_level=""""|max_order_level=""""|step_order_level=""""|product_box=""""|"    NULL                         ""         ""          ""         0            1                   0       "2013-11-03 16:35:41"      42       "2013-11-03 16:44:17"        42             "0000-00-00 00:00:00"              0

    ' virtuemart_product_id` = {virtuemart_product_id: INT UNSIGNED},   1 aytomato
    '`virtuemart_vendor_id` = {virtuemart_vendor_id: SMALLINT UNSIGNED},  1
    '`product_parent_id` = {product_parent_id: INT UNSIGNED},   0
    '`product_sku` = {product_sku: CHAR},     '12345'
    '`product_weight` = {product_weight: DECIMAL},           null
    '`product_weight_uom` = {product_weight_uom: CHAR},     kg
    '`product_length` = {product_length: DECIMAL},        NULL
    '`product_width` = {product_width: DECIMAL},         NULL
    '`product_height` = {product_height: DECIMAL},      NULL
    '`product_lwh_uom` = {product_lwh_uom: CHAR},      'M'
    '`product_url` = {product_url: CHAR},                ''
    '`product_in_stock` = {product_in_stock: INT},       1
    '`product_ordered` = {product_ordered: INT},                                0
    '`low_stock_notification` = {low_stock_notification: INT UNSIGNED},         0
    '`product_available_date` = {product_available_date: DATETIME},             "2013-11-03 00:00:00"
    '`product_availability` = {product_availability: CHAR},                      '123.gif'
    '`product_special` = {product_special: TINYINT},                 1
    '`product_sales` = {product_sales: INT UNSIGNED},                 0
    '`product_unit` = {product_unit: VARCHAR},                           TEM
    '`product_packaging` = {product_packaging: DECIMAL UNSIGNED},
    '`product_params` = {product_params: VARCHAR},                      "min_order_level=""""|max_order_level=""""|step_order_level=""""|product_box=""""|"
    '`hits` = {hits: INT UNSIGNED},            NULL
    '`intnotes` = {intnotes: VARCHAR},       ''
    '`metarobot` = {metarobot: VARCHAR},      ''
    '`metaauthor` = {metaauthor: VARCHAR},    ''
    '`layout` = {layout: CHAR},                        0
    '`published` = {published: TINYINT},                 1
    '`pordering` = {pordering: MEDIUMINT UNSIGNED},       0
    '`created_on` = {created_on: DATETIME},              "2013-11-03 16:35:41"
    '`created_by` = {created_by: INT},                  42
    '`modified_on` = {modified_on: DATETIME},           "2013-11-03 16:35:41"
    '`modified_by` = {modified_by: INT},                 42
    '`locked_on` = {locked_on: DATETIME},               "0000-00-00 00:00:00"
    '`locked_by` = {locked_by: INT}                       0
 
    'INSERT INTO `jsnrw_virtuemart_product_categories` (virtuemart_product_id,virtuemart_category_id,ordering) VALUES (7,9,0);
    'προσοχη στο πεδιο slug πρεπει να ειναι μοναδικό π.χ. str(product_id)
 
    'INSERT INTO `jsnrw_virtuemart_products_el_gr` (virtuemart_product_id,product_s_desc,product_desc,product_name,metadesc,metakey,customtitle,slug)
    '                                                     VALUES (9,        'HP No 21',    '',           'DASSDAS','','','',"9");
 
    EID.Recordset.Move 0  ' gia na ginei to rich1 opos to mseditor21
 
    a = RichToHTML(Rich1, 0&, Len(Rich1.Text))  '(.TextRTF)
    a = Replace(a, "'", "''")

    gdblag.Close

    gdblag.Open f_connect  '"DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131 ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
    R.Open "select * from jsnrw_virtuemart_products where product_sku='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

    M_DATE = Format(Now, "YYYY-MM-DD")

    If R.EOF Then
        m_new = 1
   
        sql = "INSERT INTO `jsnrw_virtuemart_products` (virtuemart_vendor_id,product_parent_id,product_sku,product_weight,product_weight_uom,product_length,product_width,product_height,product_lwh_uom,product_url,product_in_stock,product_ordered,low_stock_notification,product_available_date,product_availability,product_special,product_sales,product_unit,product_packaging,product_params,hits,intnotes,metarobot,metaauthor,layout,published,pordering,created_on,created_by,modified_on,modified_by,locked_on,locked_by) "
        sql = sql + " VALUES (1,0,'" + KOD + "',NULL,'TEM',NULL,NULL,NULL,'M','',0,0,0,'2013-11-03 00:00:00','',0,0,'KG',NULL,"
        sql = sql + " '',NULL,'','','','0',1,0,'2013-11-03 16:59:59',42,'2013-11-03 16:59:59',42,'0000-00-00 00:00:00',0);"
        gdblag.Execute sql
        R.Close
        R.Open "select * from jsnrw_virtuemart_products where product_sku='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
    End If

    m_ID = str(R("virtuemart_product_id"))
    R.Close

    Dim b As String
 
    Dim mTimh

    Dim h     As Single
  
    Dim kor   As String

    Dim manuf As String
 
    h = (100 + g_Fpa(Val(Left(Combo2.Text, 2)))) / 100
    mTimh = Format(Val(WEBPRICE.Text) / h, "####0.0000")

    If IsNull(EID.Recordset("ERG")) Then
        m_erg = " "
    Else
        m_erg = " #" + EID.Recordset("ERG")
    End If

    Dim FPA As String

    '**********  fpa tax class ******************************
    'PRODUCTS_tax_class_ID=1 =>  23%
    'PRODUCTS_tax_class_ID=2 => 6.5%
    'PRODUCTS_tax_class_ID=3 => 13%
 
    FPA = "1"

    If Val(Left(Combo2.Text, 2)) = 1 Then
        FPA = "2"
    ElseIf Val(Left(Combo2.Text, 2)) = 2 Then
        FPA = "2"
    ElseIf Val(Left(Combo2.Text, 2)) = 4 Then
        FPA = "3"
    End If

    'm_new = 1
    If m_new = 1 Then
        ' Exit Sub
        gdblag.Execute "INSERT INTO jsnrw_virtuemart_products_el_gr (virtuemart_product_id,product_s_desc,product_desc,product_name,metadesc,metakey,customtitle,slug) VALUES (" + m_ID + ",'" + FONO.Text + "','" + a + "','" + FONO.Text + "','','','','" + m_ID + "');"
        gdblag.Execute "INSERT INTO jsnrw_virtuemart_product_categories (virtuemart_product_id,virtuemart_category_id,ordering) VALUES (" + m_ID + "," + m_categ + ",0);"
        gdblag.Execute "INSERT INTO jsnrw_virtuemart_product_prices (virtuemart_product_id,virtuemart_shoppergroup_id,product_price,override,product_override_price,product_tax_id,product_discount_id,product_currency,product_price_publish_up,product_price_publish_down,price_quantity_start,price_quantity_end,created_on,created_by,modified_on,modified_by,locked_on,locked_by) VALUES " & "(" + m_ID + ",0,'" + Format(Val(WEBPRICE.Text), "####.00000") + "',0,'0.00000'," + FPA + ",0,47,'2013-11-01 00:00:00','0000-00-00 00:00:00',0,0,'2013-11-03 16:35:41',42,'2013-11-03 16:44:17',42,'0000-00-00 00:00:00',0);"
    Else
        gdblag.Execute "update jsnrw_virtuemart_products_el_gr set product_name='" + FONO.Text + "' ,product_s_desc='" + FONO.Text + "',product_desc='" + a + "'  where virtuemart_product_id=" + m_ID
        gdblag.Execute "delete from  jsnrw_virtuemart_product_categories where virtuemart_product_id=" + m_ID + " and   virtuemart_category_id=" + m_categ
        gdblag.Execute "INSERT INTO jsnrw_virtuemart_product_categories (virtuemart_product_id,virtuemart_category_id,ordering) VALUES (" + m_ID + "," + m_categ + ",0);"
    End If

    'INSERT INTO `jsnrw_virtuemart_product_prices`
    '(virtuemart_product_price_id   ,virtuemart_product_id,          virtuemart_shoppergroup_id,     product_price,
    ' override   ,product_override_price,         product_tax_id,           product_discount_id,
    ' product_currency ,product_price_publish_up
    ',product_price_publish_down,   price_quantity_start,    price_quantity_end,   created_on, created_by,
    '  modified_on,modified_by,locked_on,locked_by) VALUES
    ' (1,1,0,'100.00000',
    '  0,'0.00000',0,0,
    '  47,'2013-11-01 00:00:00',
    ' '0000-00-00 00:00:00'  ,0,0,'2013-11-03 16:35:41',42,
    '  '2013-11-03 16:44:17',42,'0000-00-00 00:00:00',0);

    Dim energo As String

    energo = "0"

    If Check1.Value = vbChecked Then
        energo = "1"
    End If

    gdblag.Execute "update jsnrw_virtuemart_products  set published = " + energo + "   where virtuemart_product_id=" + m_ID

    'r.Open "select * from products_shipping where products_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic
    'If r.EOF Then
    '  gdblag.Execute "INSERT INTO products_shipping (products_id,products_ship_methods_id,products_ship_zip,products_ship_price,products_ship_price_two) VALUES (" + m_ID + ",0,'" + ZIP_CODE.Text + "','" + INDV_PRICE.Text + "','" + ADDIT_PRICE.Text + "');"
    'Else
    '  gdblag.Execute "update products_shipping set  products_ship_price_two='" + ADDIT_PRICE.Text + "',products_ship_price='" + INDV_PRICE.Text + "',products_ship_zip='" + ZIP_CODE.Text + "' where products_id=" + m_ID
    'End If
    '
    '
    '
    '
    'r.Close
    'r.Open "select * from products_to_categories where products_id=" + m_ID + " and  categories_id=" + m_categ, gdblag, adOpenDynamic, adLockOptimistic
    'If r.EOF Then
    '  gdblag.Execute "INSERT INTO products_to_categories (products_id,categories_id) VALUES (" + m_ID + "," + m_categ + ");"
    'End If
    'r.Close
 
    'gdblag.Execute "update products set products_tax_class_id=" + FPA + " where products_id=" + m_ID
    ''----------------------------------  CATEGORIES --------------------------------------------------------------------------------------------
    'r.Open "select count(*) from products_to_categories WHERE products_id=" + m_ID + " AND categories_id=" + m_categ + ";", gdblag, adOpenDynamic, adLockOptimistic
    'If r(0) > 0 Then
    '   ' ok yparxei den kano tipota
    'Else
    '   gdblag.Execute "insert into  products_to_categories ( products_id,categories_id) values (" + m_ID + "," + m_categ + ") ;"
    'End If
    ''----------------------------------  CATEGORIES --------------------------------------------------------------------------------------------
    '
    '
    '
    'gdblag.Execute "DELETE FROM products_attributes where products_id=" + m_ID + "  limit 9", k
    'Dim sort As Integer
    'sort = 0
    'For k = 0 To hlikies.ListCount - 1
    '
    '  If hlikies.Selected(k) Then
    '
    '      gdblag.Execute "INSERT INTO products_attributes " _
    '      & "(products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) " _
    '      & " VALUES (" + m_ID + ",1," + str(f_hlikies(k)) + ",'0.0000',''," + str(sort) + ");"
    '      sort = sort + 1
    '  End If
    '
    'Next k
   
    'Uploading file to server
    'PUT method is a command in FTP to upload file to a server
    'Inet1.Execute , "PUT  ""c:\111""" & "images/stories/virtuemart/product/111"
    'MsgBox Inet1.ResponseInfo

    'Downloading a file from server
    'GET method is a command in FTP to download file from a server
    'Inet1.Execute , "GET  ""images/stories/virtuemart/product/111""" & "c:\111"

    'Server Access type
    ' INET1.AccessType = icUseDefault
    'Protocol to be used
    ' INET1.protocol = icFTP
    'Remote host name
    ' INET1.RemoteHost = "www.lagakis.gr"
    'Server port number, usually its 21
    ' INET1.RemotePort = "21"
    'Server password
    ' INET1.Password = "39213921"
    'Server Username
    ' INET1.UserName = "lagakis"
    'Server session timeout
    ' INET1.RequestTimeout = "60"
   
    Dim site_foto(6) As String

    Dim n_fotos      As Integer

    'ΒΛΕΠΩ ΤΙ ΦΩΤΟΓΡΑΦΙΕς ΕΧΕΙ
    R.Open "SELECT * FROM jsnrw_virtuemart_product_medias  inner join  jsnrw_virtuemart_medias  on jsnrw_virtuemart_medias.virtuemart_media_id=jsnrw_virtuemart_product_medias.virtuemart_media_id where virtuemart_product_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic

    'gdblag.Execute "delete from  jsnrw_virtuemart_product_medias where virtuemart_product_id = " + m_ID
    n_fotos = 0

    Do While Not R.EOF
        site_foto(n_fotos) = CNull(R!file_title)
        n_fotos = n_fotos + 1
        R.MoveNext
    Loop
  
    ' ΠΩΣ ΑΠΟΘΗΚΕΥΟΝΤΑΙ ΟΙ ΦΩΤΟΓΡΑΦΙΕΣ
    ' 1. FTP THN FOTOGRAFIA
    ' 2. INSERT INTO `jsnrw_virtuemart_medias` (ΑΥΤΟΜΑΤΟ virtuemart_media_id      ,virtuemart_vendor_id,file_title,file_description,file_meta,file_mimetype,file_type,file_url,file_url_thumb,file_is_product_image,file_is_downloadable,file_is_forSale,file_params,file_lang,shared,published,created_on,created_by,modified_on,modified_by,locked_on,locked_by) VALUES (2,1,'background1.bmp','','','','product','images/stories/virtuemart/product/background1.bmp','',0,0,0,'','',0,1,'2013-11-05 18:55:35',42,'2013-11-06 05:34:29',42,'0000-00-00 00:00:00',0);
    ' ΠΑΙΡΝΩ ΤΟ ID TOY MEDIA ( virtuemart_media_id) ΚΑΙ ΤΟ ΒΑΖΩ ΣΤΟ 3
    ' 3. INSERT INTO `jsnrw_virtuemart_product_medias` (id,virtuemart_product_id,virtuemart_media_id,ordering) VALUES (1,7,1,1);
 
    nc = 0
    '   For k = 0 To 6
    '       Me.Caption = "upload foto" + str(k)
    '       If Len(PIC(k).Caption) > 0 Then
    '          ftp_Eikones "c:\mercvb\images\" + PIC(k).Caption, "new_site/images/stories/virtuemart/product/" + PIC(k).Caption
    '          nc = nc + 1
    '       End If
    '  Next
    '/new_site/images/stories/virtuemart/product
    '  ENALLAKTIKH APOSTOLH ME FTP DOS
   
    Dim m_inet

    m_inet = 1
     
    ' Dim nc As Integer
    Dim milsecs As Long

    If ftp.ListIndex = 1 Then ' m_inet = 1 Then
 
        '   nc = 0
        '   For k = 0 To n_fotos
        '       Me.Caption = "upload foto" + str(k)
        '       If Len(PIC(k).Caption) > 0 Then
        '
        '          If Len(site_foto(k)) > 1 Then
        '             ' INET1.Execute , "DELETE " + Replace(site_foto(k), "*", "-") 'βγαζω το αστερακι μην κανει ζημια μεγαλη
        '          End If
        '
        '          'I can use i wait loop
        '          milsecs = GetCurrentTime()
        '          Do While ' INET1.StillExecuting = True
        '              If GetCurrentTime() - milsecs > 15000 Then Exit Do
        '              DoEvents
        '          Loop
        '
        '
        '          ' INET1.Execute , "PUT  """ + F_FAKEL_KOINOS + PIC(k).Caption + """" & "images/stories/virtuemart/product/" + PIC(k).Caption
        '          'I can use i wait loop
        '          milsecs = GetCurrentTime()
        '          Do While  INET1.StillExecuting = True
        '              If GetCurrentTime() - milsecs > 15000 Then Exit Do
        '              DoEvents
        '          Loop
        '          '   Print #file, "send c:\mercvb\images\" + PIC(k).Caption
        '          nc = nc + 1
        '       End If
        '  Next
    Else   'dos + api
        file = 3
        Open "c:\mercvb\send_foto" For Output As #file
        Print #file, "open www.lagakis.gr"
        Print #file, "lagakis"
        Print #file, "39213921"
        Print #file, "binary"
        Print #file, "cd images/stories/virtuemart/product"
  
        nc = 0

        For k = 0 To 6
            Me.Caption = "upload foto" + str(k)

            If Len(PIC(k).Caption) > 0 Then
          
                If ftp.ListIndex = 2 Then ' api
                    ftp_Eikones "c:\mercvb\images\" + PIC(k).Caption, "new_site/images/stories/virtuemart/product/" + Replace(PIC(k).Caption, " ", "")
                Else
                    Print #file, "send c:\mercvb\images\" + PIC(k).Caption
                End If
           
                nc = nc + 1
            End If

        Next

        Print #file, "quit"
        Close #file

        If nc > 0 Then
            Shell ("ftp -s:c:\mercvb\send_foto")
        End If

    End If

    ' ΠΩΣ ΑΠΟΘΗΚΕΥΟΝΤΑΙ ΟΙ ΦΩΤΟΓΡΑΦΙΕΣ
    ' 1. FTP THN FOTOGRAFIA
    ' 2. INSERT INTO `jsnrw_virtuemart_medias` (ΑΥΤΟΜΑΤΟ virtuemart_media_id      ,virtuemart_vendor_id,file_title,file_description,file_meta,file_mimetype,file_type,file_url,file_url_thumb,file_is_product_image,file_is_downloadable,file_is_forSale,file_params,file_lang,shared,published,created_on,created_by,modified_on,modified_by,locked_on,locked_by) VALUES (2,1,'background1.bmp','','','','product','images/stories/virtuemart/product/background1.bmp','',0,0,0,'','',0,1,'2013-11-05 18:55:35',42,'2013-11-06 05:34:29',42,'0000-00-00 00:00:00',0);
    ' ΠΑΙΡΝΩ ΤΟ ID TOY MEDIA ( virtuemart_media_id) ΚΑΙ ΤΟ ΒΑΖΩ ΣΤΟ 3
    ' 3. INSERT INTO `jsnrw_virtuemart_product_medias` (id,virtuemart_product_id,virtuemart_media_id,ordering) VALUES (1,7,1,1);
    nc = 0
 
    '   For k = 0 To 6
    '       Me.Caption = "upload foto" + str(k)
    '       If Len(PIC(k).Caption) > 0 Then
    '          ftp_Eikones "c:\mercvb\images\" + PIC(k).Caption, "new_site/images/stories/virtuemart/product/" + Replace(PIC(k).Caption, " ", "")
    '          nc = nc + 1
    '       End If
    '  Next
    '
 
    If m_new = 0 Then
        gdblag.Execute "delete from  jsnrw_virtuemart_product_medias where virtuemart_product_id = " + m_ID
    End If
 
    For k = 0 To 6

        If Len(PIC(k).Caption) > 0 Then
            gdblag.Execute "INSERT INTO jsnrw_virtuemart_medias (virtuemart_vendor_id,file_title,file_description,file_meta,file_mimetype,file_type,file_url,file_url_thumb,file_is_product_image,file_is_downloadable,file_is_forSale,file_params,file_lang,shared,published,created_on,created_by,modified_on,modified_by,locked_on,locked_by) VALUES (1,'" + PIC(k).Caption + "','','','image/jpeg','product','images/stories/virtuemart/product/" + PIC(k).Caption + "','',0,0,0,'','',0,1,'" + Format(Now, "YYYY-MM-DD") + "',42,'" + Format(Now, "YYYY-MM-DD") + "',42,'0000-00-00 00:00:00',0);          "
            nc = nc + 1
            
            gdblag.Execute "INSERT INTO jsnrw_virtuemart_product_medias (virtuemart_product_id,virtuemart_media_id,ordering) VALUES (" + m_ID + ",(SELECT MAX(virtuemart_media_id) FROM jsnrw_virtuemart_medias )," + str(nc) + ");"
        End If

    Next

    Dim surl As String

    surl = "http://www.lagakis.gr/administrator/index.php?option=com_virtuemart&view=product&task=edit&virtuemart_product_id=" + m_ID + "&product_parent_id=0"

    'BARMAIN.Value = 100
    WebBrowser1.SetFocus
    WebBrowser1.navigate surl

    'SendKeys ("admin")
    'SendKeys ("{TAB}")
    'SendKeys ("lagakis")
    'SendKeys ("{Enter}")
    '

    While Not WebBrowser1.readyState = READYSTATE_COMPLETE

        DoEvents

    Wend

    WebBrowser1.SetFocus
    SendKeys ("admin")
    SendKeys ("{TAB}")

    SendKeys ("lagakis")
    'SendKeys ("{Enter}")

    Dim f

    f = WebBrowser1.Document.Forms(0)

End Sub

Private Sub VScroll1_Change()
    VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
    Picture2.Top = -VScroll1.Value + Picture1.Height + Picture1.Top
End Sub

'Private Sub cmbMouseWheel_Click()
'
'    Dim s As Long
'    s = 200  'step of scrolling
'
'    Select Case cmbMouseWheel.Text 'Up or Down
'        Case "Up"
'          If VScroll1.Value - s <= VScroll1.Min Then
'             VScroll1.Value = VScroll1.Min
'          Else
'              VScroll1.Value = VScroll1.Value - s
'          End If
'
'        Case "Down"
'
'
'         If VScroll1.Value + s >= VScroll1.Max Then
'             VScroll1.Value = VScroll1.Max
'          Else
'              VScroll1.Value = VScroll1.Value + s
'          End If
'
'    End Select
'
'    AutoClick = True
'        Me.cmbMouseWheel.ListIndex = 1 'Make Idle
'    AutoClick = False
'    '<EXTRA: PB(1) Don't be invisible >
'       ' If PB(1).Top < 0 Then PB(1).Top = 0
'        'If PB(1).Top > (PB(0).Height - PB(1).Height) Then PB(1).Top = (PB(0).Height - PB(1).Height)
'    '</EXTRA: PB(1) Don't be invisible >>
'End Sub

'Private Sub Form_Load()
'    cmbMouseWheel.AddItem "Up" 'Add references
'    cmbMouseWheel.AddItem "Idle"
'    cmbMouseWheel.AddItem "Down"
'
'    cmbMouseWheel.ListIndex = 1 'select Idle
'End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If GetActiveWindow = Me.hWnd And Me.ActiveControl <> tbDeactivate Then tbDeactivate.SetFocus
'    '<Info>
'        If GetActiveWindow <> Me.hWnd Then
'            Label2.Caption = "Info : Form is inactive. So you can't Scroll."
'        Else
'            Label2.Caption = "Info : You can't Scroll."
'        End If
'    '</Info>
'End Sub

'Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
''           Picture2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If GetActiveWindow = Me.hWnd And Me.ActiveControl <> cmbMouseWheel Then
'           cmbMouseWheel.SetFocus
'     End If
'    '<Info>
'        If GetActiveWindow <> Me.hWnd Then
'           Me.Caption = "Info : Form is inactive. So you can't Scroll."
'        Else
'
'           Me.Caption = "Info : You can Scroll."
'        End If
'    '</Info>
'End Sub

Private Sub Command140_Click()
End Sub

'DIATHESIMO   INTEGER
'HME_DIATH Date
'KATASKEYASTHS VarChar(50)
'TEXT_PRODUCTS_ZIPCODE
'Indv.Shipping Price:
'Each Additional Price:

'**  ONO
'**  FPA 23
'*** LTI5 (XVRIS FPA)
'AGORI_KORITSI
'THEMA
'ID_YOUTUBE
'** memo ELLHNIKO
'** IMAGE0 -6
'** pos
'** MODELO (KOD)
'** BAROS

' INSERT INTO `products` (products_id, (aytomato)
'  products_youtube,products_quantity,products_model,'
' products_image,products_subimage1,products_subimage2,
' products_subimage3,products_subimage4,products_subimage5,products_subimage6,
' products_price,
' products_date_added,products_last_modified,products_date_available,products_weight,products_status,
' products_tax_class_id,
' manufacturers_id,products_ordered)
'    VALUES (29306,'',2,'5203592906988','CITY-LINE-90698.jpg','','','','','','','5.6829',
'   '2013-08-01 08:28:44','2013-08-01 08:30:40',NULL,'1.00',1,1,176,0);

'/*
'-- Query:
'select * from products where  products_model='demostration'
'Limit 0, 1000

'-- Date: 2013-10-23 08:25
'*/
'INSERT INTO `PRODUCTS`
'(products_id,products_youtube,products_quantity,products_model,products_image,products_subimage1,
' products_subimage2,products_subimage3,products_subimage4,products_subimage5,products_subimage6,
'  products_price,
' products_date_added,products_last_modified,products_date_available,
' products_weight,products_status,products_tax_class_id,
' manufacturers_id,products_ordered) VALUES (30917,'nTnnW6YlrTk',12,'demostration','773-12510-GIM.jpg',
' NULL,NULL,NULL,NULL,NULL,NULL,'8.1220','2013-10-21 08:40:09','2013-10-21 08:53:59',NULL,'0.20',0,1,109,0);

'FPA = "1"

'================================= hlikies ===============================================================================
'απο εδω διαλέγω τα διατήματα των ηλικιών τα αποια τα αποθηκεύω στα
'  products_attributes
' SELECT * FROM `toyshop`.`products_options_values`;
'products_options_values_id  language_id products_options_values_name
'1   4   "0 - 1 ετών"
'2   4   "1 - 2 ετών"
'3   4   "2 - 3 ετών"
'4   4   "3 - 5 ετών"
'5   4   "5 - 8 ετών"
'6   4   "8 - 12 ετών"
'7   4   "12 - 14 ετών"
'8   4   "14 - 17 ετών"
'9   4   "18 - 99 ετών"/*

'οι ηλικιες που αφορουν  αποθηκεύονται ΕΔΩ
'to parakato shmainei oti ayto einai gia polles ilikies
'iNSERT INTO `products_attributes` (products_attributes_id,products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) VALUES (55926,29306,1,5,'0.0000','',0);
'INSERT INTO `products_attributes` (products_attributes_id,products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) VALUES (55927,29306,1,7,'0.0000','',1);
'INSERT INTO `products_attributes` (products_attributes_id,products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) VALUES (55928,29306,1,8,'0.0000','',2);
'INSERT INTO `products_attributes` (products_attributes_id,products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) VALUES (55929,29306,1,9,'0.0000','',3);
'INSERT INTO `products_attributes` (products_attributes_id,products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) VALUES (55930,29306,1,6,'0.0000','',4);
'=========================== ΗΛΙΚΙΕς - ΤΕΛΟς ==============================================================================================
'/*
'-- Query: SELECT * FROM `toyshop`.`products_description` WHERE products_id=29306
'Limit 0, 1000
'
'-- Date: 2013-10-22 07:01
'*/
'

'===============================  EXTRA FIELDS  =================================================================

'ΑΠΟΘΗΚΕΥΟΝΤΑΙ ΣΤΟ PRODUCTS_DESCRIPTION ΩΣ ΕΞΗΣ
' extra_value_id2,extra_value_id3)
' extra_value_id1=χαρακτηρας/θεμα  extra_value_id2=ηλικια  extra_value_id3=φυλο  13=ΑΓΟΡΙ&ΚΟΡΙΤΣΙ

'INSERT INTO `products_description` (products_id,language_id,products_name,products_description,products_url,
' products_viewed,          extra_value_id1,extra_value_id2,extra_value_id3) extra_value_id1=χαρακτηρας/θεμα  extra_value_id2=ηλικια  extra_value_id3=φυλο  13=ΑΓΟΡΙ&ΚΟΡΙΤΣΙ
' VALUES (29306,4,'CITY LINE ? OVAL ORANGE #90698','<p></p>','',   22,0,0,13);

'----------  ΟΡΙΣΜΟΣ   extra_field_labels  -------------
' extra_field_labels
'1', '4', 'ΧΑΡΑΚΤΗΡΑΣ-ΘΕΜΑ', '1'
'1', '1', 'Character/Theme', '1'
'2', '4', 'ΗΛΙΚΙΑ', '1'    ' αυτο δεν χρησιμοποιείται πιά (δουλευει το products_attributes
'2', '1', 'Age', '1'
'3', '4', 'ΦΥΛΟ', '1'
'3', '1', 'Gender', '1'

'---------------------------  ΤΙΜΕΣ extra_field_values  -----------------------------------------------------
'Π.X. ΕΔΩ ΕΙΝΑΙ ΟΙ ΗΛΙΚΙΕΣ   epf_id=2 and languages_id=4
'-- Query: SELECT * FROM `toyshop`.`extra_field_values` where epf_id=2 and languages_id=4
'INSERT INTO `extra_field_values` (value_id,epf_id,languages_id,parent_id,sort_order,epf_value)
'                          VALUES (5,2,4,0,1,'0 - 12 μηνων');
'INSERT INTO `extra_field_values` (value_id,epf_id,languages_id,parent_id,sort_order,epf_value)
' VALUES (57,2,4,0,4,'4 - 8 ετων');

''Π.X. ΕΔΩ ΕΙΝΑΙ ΤΟ ΦΥΛΟ    epf_id=3 and languages_id=4
'-- Query: SELECT * FROM `toyshop`.`extra_field_values` where epf_id=3 and languages_id=4
'*/
'INSERT INTO `table_name` (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (7,3,4,0,1,'Αγόρι');
'INSERT INTO `table_name` (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (8,3,4,0,2,'Κορίτσι');
'INSERT INTO `table_name` (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (13,3,4,0,3,'Αγόρι & Κορίτσι');

''Π.X. ΕΔΩ ΕΙΝΑΙ ΤΑ ΘΕΜΑΤΑ/ΧΑΡΑΚΤΗΡΕΣ   epf_id=1 and languages_id=4
'-- Query: SELECT * FROM `toyshop`.`extra_field_values` where epf_id=1 and languages_id=4
'value_id , epf_id, languages_id, parent_id, sort_order, epf_value
'14,1,4,0,0,Barbie
'1,1,4,0,0,"Disney Princess"
'15,1,4,0,0,"Dora ? e?e?e???t??a"
'19,1,4,0,0,Spiderman
'20,1,4,0,0,Gormiti
'22,1,4,0,0,"Ben 10"
'23,1,4,0,0,"Hot Wheels"
'25,1,4,0,0,"Polly Pocket"
'26,1,4,0,0,Cars
'27,1,4,0,0,Bratz

'===============================  EXTRA FIELDS ΤΕΛΟΣ =================================================================

Private Sub WEBPRICE_Change()
    WEBPRICE.Text = Replace$(WEBPRICE.Text, ",", ".")
End Sub

Private Sub WEIGHT_Change()
    WEIGHT.Text = Replace$(WEIGHT.Text, ",", ".")
End Sub
                
''**************************************
'' Name: Loose formatting when pasting into an RTF box
'' Description:Gets rid of any text formatting if the text was copied from an RTF box. (EX. if you copied blue text from wordpad, and then paste it into your RTFbox, it would remain blue) This code is useful if your making any sort of file-editors where you dont want formatting to take effect when pasting, like an HTML or SCRIPT editor that uses a RichTextBox so it has no filesize limit.
'' By: (Tim Miron) yar-interactive software
''
'' Side Effects:For some reason, the 'And' doesn't work in this If statement, so I had to use a double-if, weird, And usually works fine...
''
''This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=14721&lngWId=1'for details.'**************************************
'
''Usually, you dont use 2 Ifs in a row, but in this case
''for some reason the 'And' thing didn't work
'Dim CBC As String 'ClipBoard Contents
'Dim a
'If Shift = 2 Then
' If KeyCode = 86 Then 'I know your thinkin "why didn't this guy use the 'And' statement in his if?"
' CBC = Clipboard.GetText 'its because it dont
''function properly if its done that way, i dunnow
''why?
' Clipboard.Clear
' Clipboard.SetText CBC
' 'When you convert it to a string, it looses its
''formatting attributes, then
' 'you re-apply it to the clipboard :)
' End If
'End If
'
'
Private Function greek(ByVal C As String) As String

    Dim cc As String

    C = C + " "
    ClipboardLetTextW C
    cc = Clipboard.GetText
    greek = Replace(cc, "¶", "Ά")
End Function

Private Sub xaraktiras_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Private Sub ZIP_CODE_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub

Sub XaraktiresTree()
'============================================================
         
On Error GoTo 0
         
'         sql = "select  value_id AS ID,parent_id AS PARENT,epf_value AS NAME from extra_field_values where epf_id=1 and languages_id=4 order by  epf_value"
        '  (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (7,3,4,0,1,'?????');
Dim R As New ADODB.Recordset
        
        

144     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'XARAKTIRESANT'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

146     If R(0) = 0 Then
148         Gdb.Execute "CREATE TABLE XARAKTIRESANT " & "([ID] [int] NULL,[PARENTNAME] CHAR(50) NULL, [PARENT] [INT] NULL,[NAME] CHAR(50)  NULL ) "
        End If

150     R.Close

152     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'XARAKTIRESTEL'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

154     If R(0) = 0 Then
            'Gdb.Execute "CREATE TABLE XARAKTIRESTEL ( [ID] [int] NOT NULL,[PARENT] [INT] NOT NULL,[NAME] [varCHAR](50) NULL,[AA] [INT] NOT NULL )ON PRIMARY  "
156         Gdb.Execute "CREATE TABLE XARAKTIRESTEL ( [ID] [int]  NULL,[PARENTNAME] CHAR(50) NULL,[PARENT] [INT]  NULL,[NAME] CHAR(50) NULL,[AA] [INT] NULL )  "
        End If

157     R.Close

        ''ΕΛΕΓΧΩ ΑΝ ΕΝΗΜΕΡΩΘΗΚΕ ΣΗΜΕΡΑ
        '158 r.Open "SELECT OBJECT_NAME(OBJECT_ID) AS DatabaseName, last_user_update AS HME  FROM sys.dm_db_index_usage_stats  WHERE database_id = DB_ID( 'MERCURY')  AND OBJECT_ID=OBJECT_ID('XARAKTIRESANT')", Gdb, adOpenDynamic, adLockOptimistic
        ' If Not r.EOF Then
        '159 If Not IsNull(r!hme) Then
        '160    If Format(r!hme, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
        '         r.Close
        '         GoTo PARAKAMCH
        '    End If
        ' End If
        'End If
        'r.Close
        '
        '
        '

        f_err = 0
161     Gdb.Execute "DELETE FROM XARAKTIRESANT"
162     Gdb.Execute "DELETE FROM XARAKTIRESTEL"

        Dim cName As String

164  '  sql = "select  *  from extra_field_values where epf_id=1 and languages_id=4 order by  epf_value"
'value_id,parent_id ,epf_value

     sql = "select  value_id AS ID,parent_id AS PARENT,epf_value AS NAME from extra_field_values where epf_id=1 and languages_id=4 order by  epf_value"





        R.Open sql, gdblag, adOpenDynamic, adLockOptimistic

166     Do While Not R.EOF

            If f_site = 444 Then
                cName = Left(greek(R!Name), 50)
                cName = Replace(cName, "'", "''")
            Else
                cName = Replace(R!Name, "'", "''")
            End If
          
168         Gdb.Execute "INSERT INTO XARAKTIRESANT (ID,PARENT,NAME) VALUES (" + str(R!ID) + "," + str(R!Parent) + ",'" + cName + "')"
170         R.MoveNext

            If f_err = 1 Then Exit Do
        Loop

172     R.Close

174     Gdb.Execute "INSERT INTO XARAKTIRESTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,1 FROM XARAKTIRESANT WHERE PARENT=0"
        
        Gdb.Execute "DELETE FROM XARAKTIRESANT WHERE PARENT=0"

        Gdb.Execute "UPDATE XARAKTIRESANT SET  PARENTNAME=(SELECT TOP 1 ISNULL(SUBSTRING(NAME,1,25),'.' ) FROM XARAKTIRESTEL WHERE ID=XARAKTIRESANT.PARENT)"
        
        Gdb.Execute "UPDATE XARAKTIRESTEL SET PARENTNAME=NAME"


176     Gdb.Execute "INSERT INTO XARAKTIRESTEL(ID,PARENT,NAME,AA,PARENTNAME) SELECT ID,PARENT,'  '+NAME,2,PARENTNAME  FROM XARAKTIRESANT " ' WHERE PARENT IN (SELECT ID FROM XARAKTIRESTEL)"

'        ' GIA NA FYGEI  TO 1&2  EPIPEDO APO TO XARAKTIRESANT
'178     Gdb.Execute "DELETE FROM XARAKTIRESANT WHERE ID IN (SELECT ID FROM XARAKTIRESTEL)"
'
'180     Gdb.Execute "INSERT INTO XARAKTIRESTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,3  FROM XARAKTIRESANT WHERE PARENT IN (SELECT ID FROM XARAKTIRESTEL)"
'
'        ' GIA NA FYGEI  TO 3  EPIPEDO APO TO XARAKTIRESANT
'182     Gdb.Execute "DELETE FROM XARAKTIRESANT WHERE ID IN (SELECT ID FROM XARAKTIRESTEL)"
'
'184     Gdb.Execute "INSERT INTO XARAKTIRESTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,4  FROM XARAKTIRESANT WHERE PARENT IN (SELECT ID FROM XARAKTIRESTEL)"
'
'        ' GIA NA FYGEI  TO 4  EPIPEDO APO TO XARAKTIRESANT
'186     Gdb.Execute "DELETE FROM XARAKTIRESANT WHERE ID IN (SELECT ID FROM XARAKTIRESTEL)"
'
'188     Gdb.Execute "INSERT INTO XARAKTIRESTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,5  FROM XARAKTIRESANT"
'
'PARAKAMCH:
'
'190     r.Open "SELECT * FROM XARAKTIRESTEL ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic
'        ' 192 r.MoveFirst
'
'        'r.Open "SELECT * FROM CATEG1", Gdb
'194     Do While Not r.EOF
'
'196         error = 0
'            'paterades
'198         mess = "id=" + LTrim(str(r!ID)) + "Parent:" + LTrim(str(r!Parent)) + " " + r!Name
'
'200         If r!Parent = 0 Then
'202             T1.Nodes.ADD , , "r" + LTrim(str(r!ID)), r!Name
'            Else
'
'204             T1.Nodes.ADD "r" + LTrim(str(r!Parent)), tvwChild, "r" + LTrim(str(r!ID)), (r!Name)
'            End If
'
'206         If error = 0 Then
'208             n = n + 1
'210             names(n) = "r" + Format(r!ID, "0000") + "        " + r!Name
'                FC_id(n) = r!ID
'                FC_parenT(n) = r!Parent
'                fc_names(n) = r!Name
'            End If
'
'212         r.MoveNext
'        Loop
'
'214     r.Close
'        fc_max = n



End Sub





