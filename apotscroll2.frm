VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form apot2scroll 
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   225
   ClientTop       =   -2385
   ClientWidth     =   15240
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc EGGTIM2 
      Height          =   495
      Left            =   15000
      Top             =   9960
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
      TabIndex        =   49
      Top             =   3720
      Visible         =   0   'False
      Width           =   1812
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"apotscroll2.frx":0000
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
         TabIndex        =   72
         Top             =   0
         Width           =   900
      End
      Begin VB.ListBox List1 
         Height          =   780
         Left            =   2640
         TabIndex        =   71
         Top             =   0
         Width           =   2655
      End
      Begin VB.CommandButton CMDmovenext 
         Caption         =   ">>"
         Height          =   250
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton CMDmoveprevious 
         Caption         =   "<<"
         Height          =   250
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   200
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Αναζήτηση με αρ.τιμολογίου"
         Height          =   180
         Left            =   10440
         TabIndex        =   67
         Top             =   0
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   900
         TabIndex        =   65
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton neo 
         BackColor       =   &H008080FF&
         Caption         =   "Nέο Είδος"
         Height          =   225
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton upload1 
         Caption         =   "Upload"
         Height          =   465
         Left            =   6720
         TabIndex        =   34
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
         TabIndex        =   70
         Top             =   0
         Width           =   1605
      End
      Begin VB.Label Label7 
         Caption         =   "Barcode"
         Height          =   255
         Left            =   0
         TabIndex        =   66
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
         TabIndex        =   52
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
         Width           =   6645
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
      Left            =   14880
      TabIndex        =   1
      Top             =   0
      Width           =   252
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFC0&
      Height          =   13935
      Left            =   0
      ScaleHeight     =   13875
      ScaleWidth      =   14835
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   14895
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   11760
         TabIndex        =   64
         Top             =   9000
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "DOWNLOAD TABLE PRODUCTS"
         Height          =   375
         Left            =   10440
         TabIndex        =   63
         Top             =   9000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox WEBPRICE 
         Height          =   288
         Left            =   8760
         TabIndex        =   62
         Top             =   120
         Width           =   1332
      End
      Begin VB.TextBox YOU_TUBE 
         DataField       =   "YOU_TUBE"
         DataSource      =   "EID"
         Height          =   288
         Left            =   7200
         TabIndex        =   59
         Top             =   5520
         Width           =   1692
      End
      Begin VB.TextBox ZIP_CODE 
         DataField       =   "TEXT_PRODUCTS_ZIPCODE"
         DataSource      =   "EID"
         Height          =   288
         Left            =   2280
         TabIndex        =   55
         Top             =   5160
         Width           =   1572
      End
      Begin VB.TextBox INDV_PRICE 
         DataField       =   "SHIPPINGPRICE"
         DataSource      =   "EID"
         Height          =   288
         Left            =   2280
         TabIndex        =   54
         Top             =   5520
         Width           =   1572
      End
      Begin VB.TextBox ADDIT_PRICE 
         DataField       =   "ADDITIONALPRICE"
         DataSource      =   "EID"
         Height          =   288
         Left            =   2280
         TabIndex        =   53
         Top             =   5880
         Width           =   1572
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   8760
         MultiLine       =   -1  'True
         TabIndex        =   48
         Text            =   "apotscroll2.frx":0086
         Top             =   9000
         Width           =   1572
      End
      Begin VB.ListBox hlikies 
         Height          =   1500
         Left            =   5640
         MultiSelect     =   2  'Extended
         TabIndex        =   46
         Top             =   3840
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "Διαθέσιμο"
         Height          =   255
         Left            =   12120
         TabIndex        =   44
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox kataskeyasths 
         Height          =   300
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   4080
         Width           =   3255
      End
      Begin VB.ComboBox xaraktiras 
         Height          =   300
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   4800
         Width           =   3255
      End
      Begin VB.ComboBox fylo 
         Height          =   300
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   4440
         Width           =   3255
      End
      Begin VB.TextBox WEIGHT 
         DataField       =   "WEIGHT"
         DataSource      =   "EID"
         Height          =   270
         Left            =   10920
         TabIndex        =   36
         Top             =   120
         Width           =   1095
      End
      Begin MSComctlLib.TreeView T1 
         Height          =   5412
         Left            =   9000
         TabIndex        =   35
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   6360
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
         Left            =   10800
         Top             =   8400
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
         Height          =   2772
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   8412
         _ExtentX        =   14843
         _ExtentY        =   4895
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         _ExtentX        =   3625
         _ExtentY        =   1720
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Top             =   5520
         Width           =   1245
      End
      Begin VB.Label lblEachAdditional 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Each Additional Price"
         Height          =   195
         Left            =   360
         TabIndex        =   58
         Top             =   6000
         Width           =   1425
      End
      Begin VB.Label lblIndvShipping 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Indv. Shipping Price:"
         Height          =   195
         Left            =   360
         TabIndex        =   57
         Top             =   5640
         Width           =   1410
      End
      Begin VB.Label lblZIP_CODE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZIP_CODE"
         Height          =   195
         Left            =   600
         TabIndex        =   56
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
         TabIndex        =   51
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label cΗλικίες 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Ηλικίες"
         Height          =   330
         Left            =   5640
         TabIndex        =   45
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Χαρακτήρας/Θέμα"
         Height          =   180
         Left            =   480
         TabIndex        =   43
         Top             =   4920
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αγόρι/Κορίτσι"
         Height          =   180
         Left            =   480
         TabIndex        =   42
         Top             =   4560
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κατασκευαστής"
         Height          =   180
         Left            =   480
         TabIndex        =   41
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
         TabIndex        =   37
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   10305
         Width           =   45
      End
      Begin VB.Label PIC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         DataField       =   "PIC1"
         DataSource      =   "EID"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   28
         Top             =   10305
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Τιμή καταστηματος"
         Height          =   384
         Left            =   5400
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label PIC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         DataField       =   "PIC0"
         DataSource      =   "EID"
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   17
         Top             =   7800
         Width           =   45
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1560
      End
   End
End
Attribute VB_Name = "apot2scroll"
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
Dim F_id_LOCAL As Long
Dim gdblag As New ADODB.Connection
Dim names(1000) As String
Dim f_exit As Boolean
Dim f_pel As String
Dim f_xaraktiras(500)
Dim f_fylo(10)
Dim f_kataskeyasths(500)
Dim f_hlikies(30)
Dim F_FAKEL_KOINOS As String



Dim f_arxikh_eikona(10) As String




'Private AutoClick As Boolean
'Private Declare Function GetActiveWindow Lib "user32" () As Long

Dim f_error As Integer
                
'**************************************
' Name: Fun with MouseWheel
' Description:Just intercepting MouseWheel event with API. Make an empty project (standard exe) and paste code.
' By: vViktor
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=56768&lngWId=1'for details.'**************************************

Private Const PM_REMOVE = &H1
Private Type POINTAPI
 X As Long
 Y As Long
End Type
Private Type msg
 hWnd As Long
 Message As Long
 wParam As Long
 lParam As Long
 time As Long
 pt As POINTAPI
End Type
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private bCancel As Boolean
Private Const WM_MOUSEWHEEL = 522

'αποθηκευει τα στοιχεια των κατηγοριων
Dim fc_max As Long
Dim FC_id(1000) As Long
Dim FC_parenT(1000) As Long
Dim fc_names(1000) As String




Private Sub ProcessMessages()
 Dim Message As msg
 
 Dim S
 S = 200
 Do While Not bCancel
   If f_exit = True Then
       Unload Me
   End If
 
 
WaitMessage 'Wait For message and...
If PeekMessage(Message, Me.hWnd, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then '...when the mousewheel is used...
 If Message.wParam < 0 Then '...scroll up...
 
       

  If VScroll1.Value + S >= VScroll1.max Then
             VScroll1.Value = VScroll1.max
          Else
              VScroll1.Value = VScroll1.Value + S
          End If
 
 
 
 
     ' Me.Top = Me.Top + 240
 Else '... or scroll down
 
          If VScroll1.Value - S <= VScroll1.min Then
             VScroll1.Value = VScroll1.min
          Else
              VScroll1.Value = VScroll1.Value - S
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
   
   
    On Error Resume Next
    ONOLABEL.Caption = EID.Recordset("ono")
    WEBPRICE.Text = ""
    If IsNull(EID.Recordset("WEBPRICE")) Then
        WEBPRICE.Text = EID.Recordset("lti5")
    Else
        WEBPRICE.Text = EID.Recordset("WEBPRICE")
    End If
    
    
    kodlabel.Caption = EID.Recordset("kod")
    Set msEditor21.mDataSource = EID  ' for code you have to use the 'adoNotes' instead of the Adodc1
    msEditor21.MaxLength = 8000          ' Optional maximum length.

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
    If Len(PIC(N).Caption) > 0 Then
       foto2(N).Picture = LoadPicture()
       foto2(N).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + PIC(N).Caption)
       StretchSourcePictureFromPicture apot2scroll.foto2(N).Picture, apot2scroll.foto2(N)
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
Dim r As New ADODB.Recordset
' r.Open "SELECT * FROM PRODUCT_TO_CATEGORIES WHERE KOD='" + KOD + "'", Gdb, adOpenDynamic, adLockOptimistic
'
' Do While Not r.EOF
'
'   r.MoveNext
' Loop
'
' r.Close
'
 r.Open "SELECT * FROM PRODUCT_ATTRIBUTES WHERE KOD='" + KOD + "'", Gdb, adOpenDynamic, adLockOptimistic
 Do While Not r.EOF
    For k = 0 To hlikies.ListCount - 1
      If f_hlikies(k) = r("OPTIONS_VALUES_ID") Then
          hlikies.Selected(k) = True
          Exit For
      End If
    Next

   r.MoveNext
 Loop
 r.Close


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
196    error = 0
       'paterades
        If FC_parenT(N) = 0 Then
202       T1.Nodes.Add , , "r" + LTrim(str(FC_id(N))), fc_names(N)
       Else

204     T1.Nodes.Add "r" + LTrim(str(FC_parenT(N))), tvwChild, "r" + LTrim(str(FC_id(N))), (fc_names(N))
       End If
Next






 r.Open "SELECT * FROM PRODUCT_TO_CATEGORIES WHERE KOD='" + KOD + "'", Gdb, adOpenDynamic, adLockOptimistic
 Do While Not r.EOF
   ' T1.Nodes.Item(20).Expanded = True
   ' T1.Nodes.Item(7).Selected = True
   For k = 1 To UBound(names) - 1
        If Val(mID(names(k), 2, 4)) = r!categories_id Then
           T1.Nodes.Item(k).Selected = True
           T1.Nodes.Item(k).BackColor = vbCyan
           T1.SetFocus
          ' T1.ShowWhatsThis
        
           Exit For
        End If
   Next
   Exit Do  ' GIA MIA KATHGORIA
   r.MoveNext
 Loop
 r.Close

















 
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
If k > 0 Then
  
  Gdb.Execute "delete PRODUCT_TO_CATEGORIES where KOD='" + KOD + "'", k
  Gdb.Execute "DELETE FROM PRODUCT_ATTRIBUTES where KOD='" + KOD + "' ", k
End If


'=======================================================   σιτε =========================================

Dim m_ID As String
Dim r As New ADODB.Recordset


gdblag.Close
gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131 ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
r.Open "select * from products where products_model='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic



Dim site_foto(6) As String

site_foto(0) = CNull(r!products_image)
site_foto(1) = CNull(r!products_subimage1)
site_foto(2) = CNull(r!products_subimage2)
site_foto(3) = CNull(r!products_subimage3)
site_foto(4) = CNull(r!products_subimage4)
site_foto(5) = CNull(r!products_subimage5)
site_foto(6) = CNull(r!products_subimage6)

m_ID = str(r("products_id"))
r.Close

gdblag.Execute "delete from products_shipping      where products_id=" + m_ID + " limit 1", k
gdblag.Execute "delete from products_description   where products_id=" + m_ID + " limit 1", k

gdblag.Execute "delete from products_to_categories where products_id=" + m_ID + " limit 5", k

gdblag.Execute "DELETE FROM products_attributes    where products_id=" + m_ID + "  limit 9", k

gdblag.Execute "DELETE FROM products               where products_id=" + m_ID + "  limit 1", k

   file = 3
   Open "c:\mercvb\del_foto" For Output As #file
   Print #file, "open www.toys-shop.gr"
   Print #file, "toysftp"
   Print #file, "kP8#H7o2fD"
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

Private Sub cmdΕΝΗΜΕΡΩΣΗ_Click()
 Dim a As String
 Dim m_categ As String
 Dim Y As Long
 Dim m_new As Integer
 
 'If neo.Enabled = True Then
    
 If FONO = "" Then Exit Sub
 
 'Else
    
 KOD.Enabled = True
 
 On Error Resume Next
  
If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty
   On Error Resume Next
   Y = T1.SelectedItem.Index
   
   Y = T1.SelectedItem.Index
   m_categ = mID(names(Y), 2, 4)
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
 
 
 
 Dim kor As String, b As String
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
 
 

Dim r As New ADODB.Recordset
r.Open "select * from PRODUCT_TO_CATEGORIES where KOD='" + KOD + "' and CATEGORIES_ID =" + m_categ, Gdb, adOpenDynamic, adLockOptimistic
If r.EOF Then
  Gdb.Execute "INSERT INTO PRODUCT_TO_CATEGORIES (KOD,PRODUCT_ID,CATEGORIES_ID) VALUES ('" + KOD + "'," + str(F_id_LOCAL) + "," + m_categ + ");"
End If
r.Close

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
      Gdb.Execute "INSERT INTO PRODUCT_ATTRIBUTES " _
      & "(KOD,OPTIONS_ID,OPTIONS_VALUES_ID,PRODUCT_ID) " _
      & " VALUES ('" + KOD + "',1," + str(f_hlikies(k)) + "," + str(EID.Recordset("ID")) + ");"
  End If
Next k
    
    
 'Gdb.Execute "INSERT INTO BARCODES (KOD,ERG) VALUES ('" + KOD + "','" + KOD + "')"
 
 
 Dim ANS As Integer
 Dim K0 As Integer, mc As Integer
 
 For k = 0 To 6
   If f_arxikh_eikona(k) = PIC(k).Caption Then    'ok
   Else
     If Len(f_arxikh_eikona(k)) > 1 Then
         ' ελεγχω μηπως αυτην την φωτογραφια την εχω και άλλη φορα
         mc = 0
         For K0 = 0 To 6
             If f_arxikh_eikona(K0) = PIC(k).Caption Then
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

100 F_EIKONA = 11
110     CommonDialog1.FILENAME = "*.jpg"
120     CommonDialog1.ShowOpen
        Dim m As String
        On Error Resume Next
130     m = "a" + kodlabel.Caption + "-" + Format(Index, "0") ' Text1(0).Text
140     If F_EIKONA = 11 Then
150        m = m + ".JPG"
        End If
160     FileCopy CommonDialog1.FILENAME, F_FAKEL_KOINOS + "\images\" + m
    
170      foto2(Index).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + m)
            
180             StretchSourcePictureFromPicture apot2scroll.foto2(Index).Picture, apot2scroll.foto2(Index)
        PIC(Index).Caption = m
    
        '<EhFooter>
        Exit Sub

Command14_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.pelat2.Command14_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot2scroll.Command14_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub Command2_Click()
Dim r As New ADODB.Recordset
Dim rLOC As New ADODB.Recordset
gdblag.Close
gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131 ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
r.Open "select * from products WHERE products_model<>'' order by products_model", gdblag, adOpenDynamic, adLockOptimistic
Dim N As Long
Dim KODE As String

Do While Not r.EOF()
   KODE = Replace(r!PRODUCTS_MODEL, "'", "~")
   'rLOC.Open "select * from EID WHERE KOD='" + KODE + "'", Gdb, adOpenDynamic, adLockOptimistic
'   If rLOC.EOF Then
'       Gdb.Execute "INSERT INTO EID (KOD) VALUES ('" + KODE + "')"
'   End If
'   rLOC.Close
   
        If Not IsNull(r!products_youtube) Then
           Gdb.Execute "UPDATE EID SET YOU_TUBE='" + Replace(r!products_youtube, "'", "''") + "' WHERE KOD='" + KODE + "'"
        End If
'       ' If Not IsNull(r!products_subimage1) Then Gdb.Execute "UPDATE EID SET PIC1='" + r!products_subimage1 + "' WHERE KOD='" + KODE + "'"
'
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
'
'       If Not IsNull(r!products_WEIGHT) Then Gdb.Execute "UPDATE EID SET WEIGHT=" + str(r!products_WEIGHT) + " WHERE KOD='" + KODE + "'"
'       If Not IsNull(r!products_STATUS) Then Gdb.Execute "UPDATE EID SET DIATHESIMO=" + str(r!products_STATUS) + " WHERE KOD='" + KODE + "'"
       
   '    Dim FPA As String
 '      FPA = "2"
  '     If r!products_tax_class_id = 3 Then
   '        FPA = "1"
    '   ElseIf r!products_tax_class_id = 1 Then
     '      FPA = "2"
'       ElseIf r!products_tax_class_id = 2 Then
 '          FPA = "4"
  '     End If
           
    '
       'Gdb.Execute "UPDATE EID SET FPA=" + FPA + " WHERE KOD='" + KODE + "'"
    '   Gdb.Execute "UPDATE EID SET ID_WEB=" + str(r!PRODUCTS_ID) + " WHERE KOD='" + KODE + "'"
       
       'If Not IsNull(nufacturers_id) Then Gdb.Execute "UPDATE EID SET  MANUFACTURERS_ID=" + str(r!manufacturers_id) + " WHERE KOD='" + KODE + "'"
    '   If Not IsNull(r!manufacturers_id) Then Gdb.Execute "UPDATE EID SET  MANUFACTURERS_ID=" + str(r!manufacturers_id) + " WHERE KOD='" + KODE + "'"
       
       r.MoveNext
       N = N + 1
       DoEvents
       Me.Caption = str(N)
 Loop
 r.Close
 
          
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
Dim r As New ADODB.Recordset
Dim rLOC As New ADODB.Recordset

GoTo 6000



'r.Open "select * from products_description order by products_id", gdblag, adOpenDynamic, adLockOptimistic
'Dim N As Long
'Dim KODE As String
'
'Do While Not r.EOF()
'      ' Gdb.Execute "UPDATE EID SET ONO='" + r!products_name + "' WHERE ID_WEB=" + str(r!PRODUCTS_ID)
'      ' Gdb.Execute "UPDATE EID SET MEMO='" + r!products_description + "' WHERE ID_WEB=" + str(r!PRODUCTS_ID)
'       Gdb.Execute "UPDATE EID SET XARAKTIRAS=" + str(r!extra_value_id1) + " WHERE ID_WEB=" + str(r!PRODUCTS_ID)
'       Gdb.Execute "UPDATE EID SET FYLO=" + str(r!extra_value_id3) + " WHERE ID_WEB=" + str(r!PRODUCTS_ID)
'
'
'       r.MoveNext
'       N = N + 1
'       DoEvents
'       Me.Caption = str(N)
'   '    Exit Do
' Loop
' r.Close
'Exit Sub
'
'
'
'gdblag.Close
'gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=toys-shop.gr ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
'r.Open "select * from products_to_categories", gdblag, adOpenDynamic, adLockOptimistic
'
'Dim rl As New ADODB.Recordset
'
'
'Do While Not r.EOF
'   ' Gdb.Execute ""
'   rl.Open "select * from EID WHERE ID_WEB=" + str(r!PRODUCTS_ID), Gdb, adOpenDynamic, adLockOptimistic
'   If Not rl.EOF Then
'      Gdb.Execute "INSERT INTO PRODUCT_TO_CATEGORIES (KOD,CATEGORIES_ID,PRODUCT_ID) VALUES ('" + Left(rl!KOD, 15) + "'," + str(r!categories_id) + ",0 );"
'   End If
'   rl.Close
'       N = N + 1
'       DoEvents
'       Me.Caption = str(N)
'
'   r.MoveNext
'Loop
'r.Close
'

'ATTRIBUTES

6000

gdblag.Close
gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=toys-shop.gr ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"

r.Open "select * from products_attributes ", gdblag, adOpenDynamic, adLockOptimistic
Do While Not r.EOF
   rLOC.Open "select * from EID WHERE ID_WEB=" + str(r!PRODUCTS_ID), Gdb, adOpenDynamic, adLockOptimistic
   If Not rLOC.EOF Then
        Gdb.Execute "INSERT INTO PRODUCT_ATTRIBUTES " _
      & "(KOD,OPTIONS_ID,OPTIONS_VALUES_ID,PRODUCT_ID) " _
      & " VALUES ('" + rLOC!KOD + "',1," + str(r!options_values_id) + ",0);"
   End If
   rLOC.Close
   N = N + 1
   DoEvents
   Me.Caption = str(N)
   r.MoveNext
Loop
r.Close





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
a = InputBox("ΑΡΙΘΜΟΣ ΤΙΜΟΛ.ΑΓΟΡΑΣ")
b = InputBox("ΗΜΕΡ/ΝΙΑ ΤΙΜΟΛ.ΑΓΟΡΑΣ", , Format(Now, "DD/MM/YYYY"))


EGGTIM2.ConnectionString = gConnect
EGGTIM2.RecordSource = "SELECT KODE FROM EGGTIM WHERE ATIM LIKE '%" + a + "%' AND CONVERT(CHAR(10),HME,103)='" + b + "' ORDER BY PROOD_AJ"
EGGTIM2.Refresh
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




Private Sub FONO_KeyUp(KeyCode As Integer, Shift As Integer)
100     If KeyCode = vbKeyReturn Then
110         KeyCode = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100  Me.AutoRedraw = True
 
102 f_pel = "mpoyg" ' "lagak"
 
     'syndesh xori dsn me connection string  ' ip toys 188.40.85.131
    ' gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=db5.papaki.gr;PORT =3306;DATABASE=vasi;UID=lagakis;Password=3898339214231"
On ERRROR GoTo NOCONNECT
104 gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=toys-shop.gr;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"


 On Error GoTo Form_Load_Err


106  EID.ConnectionString = gConnect
    Dim DUM
108      DUM = ADD_FIELD("EID", "MEMOENG", "TEXT")
110      DUM = ADD_FIELD("EID", "PIC0", "VARCHAR(100)")
112      DUM = ADD_FIELD("EID", "PIC1", "VARCHAR(100)")
114      DUM = ADD_FIELD("EID", "PIC2", "VARCHAR(100)")
116      DUM = ADD_FIELD("EID", "PIC3", "VARCHAR(100)")
118      DUM = ADD_FIELD("EID", "PIC4", "VARCHAR(100)")
120      DUM = ADD_FIELD("EID", "PIC5", "VARCHAR(100)")
122      DUM = ADD_FIELD("EID", "PIC6", "VARCHAR(100)")
    
124 DUM = ADD_FIELD("EID", "DIATHESIMO", "INT")

125 DUM = ADD_FIELD("EID", "ID_WEB", "INT")  ' TO ID TOY SITE
126 DUM = ADD_FIELD("EID", "HME_DIATH", "DATETIME")

128 DUM = ADD_FIELD("EID", "MANUFACTURERS_ID", "INT")
130 DUM = ADD_FIELD("EID", "TEXT_PRODUCTS_ZIPCODE", "VARCHAR(30)")

    DUM = ADD_FIELD("EID", "YOU_TUBE", "VARCHAR(100)")

132 DUM = ADD_FIELD("EID", "ADDITIONALPRICE", "DECIMAL(10,2)")
133 DUM = ADD_FIELD("EID", "SHIPPINGPRICE", "DECIMAL(10,2)")


134 DUM = ADD_FIELD("EID", "WEIGHT", "DECIMAL(10,2)")
135 DUM = ADD_FIELD("EID", "WEBPRICE", "DECIMAL(10,2)")

136 DUM = ADD_FIELD("EID", "FYLO", "INT")
138 DUM = ADD_FIELD("EID", "XARAKTIRAS", "INT")
'' 139 DUM = ADD_FIELD("EID", "CATEGORY", "INT")

    'DIATHESIMO   INTEGER
    'HME_DIATH Date
    'KATASKEYASTHS VarChar(50)
    'TEXT_PRODUCTS_ZIPCODE
    'Indv.Shipping Price:
    'Each Additional Price:
    
    
F_FAKEL_KOINOS = FINDPARAMETROI("APOT2", "F_FAKEL_KOINOS", "C:\MERCVB", "KOINOXPHΣTOΣ ΦΑΚΕΛΟΣ ΕΙΚΟΝΩΝ Π.Χ. \\PC\MERCVB")
    
    
    
    
    
    
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
    Dim r As New ADODB.Recordset
    Dim sql As String

    'On Error Resume Next

    'lagakis
    'sql = "SELECT category_child_id AS ID,category_parent_id AS PARENT,category_name AS NAME   FROM jos_vm_category_xref a inner join jos_vm_category b on a.category_child_id=b.category_id  ORDER BY category_parent_id"

    
    
    'mpoygoudis

'============================= ΚΑΤΗΓΟΡΙΕΣ =====================================================
140 sql = "SELECT a.categories_id as ID,a.parent_id as PARENT,b.CATEGORIES_NAME AS NAME from categories a inner join categories_description b on a.categories_id=b.categories_id order by parent_id,a.categories_id"
     
    Dim N As Long
142 N = 0
    Dim error As Integer

'products_id,options_id,options_values_id

'================================ PRODUCT_ATTRIBUTES ================================
 r.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'PRODUCT_ATTRIBUTES'", Gdb, adOpenDynamic, adLockOptimistic
    'On Error Resume Next

 If r(0) = 0 Then
    Gdb.Execute "CREATE TABLE PRODUCT_ATTRIBUTES ( KOD VARCHAR(16) NOT NULL,OPTIONS_VALUES_ID INT,PRODUCT_ID int NOT NULL,OPTIONS_ID  INT  NULL ) "
    End If
 r.Close


'================================ PRODUCT_TO_CATEGORIES ================================
 r.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'PRODUCT_TO_CATEGORIES'", Gdb, adOpenDynamic, adLockOptimistic
    'On Error Resume Next

 If r(0) = 0 Then
    Gdb.Execute "CREATE TABLE PRODUCT_TO_CATEGORIES ( KOD VARCHAR(16) NOT NULL,PRODUCT_ID int NOT NULL,CATEGORIES_ID INT NOT NULL ) "
    End If
 r.Close





144 r.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'CATEGORIESANT'", Gdb, adOpenDynamic, adLockOptimistic
    'On Error Resume Next

146 If r(0) = 0 Then
148    Gdb.Execute "CREATE TABLE CATEGORIESANT " _
       & "([ID] [int] NOT NULL,[PARENT] [INT] NOT NULL,[NAME] [varCHAR](50) NULL,CONSTRAINT [PK_CATEGORIESANT] PRIMARY KEY CLUSTERED " _
      & "([ID] Asc )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY] "
     End If
150 r.Close


152 r.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'CATEGTEL'", Gdb, adOpenDynamic, adLockOptimistic
    'On Error Resume Next

154 If r(0) = 0 Then
       'Gdb.Execute "CREATE TABLE CATEGTEL ( [ID] [int] NOT NULL,[PARENT] [INT] NOT NULL,[NAME] [varCHAR](50) NULL,[AA] [INT] NOT NULL )ON PRIMARY  "
156    Gdb.Execute "CREATE TABLE CATEGTEL ( [ID] [int] NOT NULL,[PARENT] [INT] NOT NULL,[NAME] [varCHAR](50) NULL,[AA] [INT] NOT NULL )  "
    End If



157 r.Close


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
161 Gdb.Execute "DELETE FROM CATEGORIESANT"
162 Gdb.Execute "DELETE FROM CATEGTEL"

164 r.Open sql, gdblag, adOpenDynamic, adLockOptimistic
166 Do While Not r.EOF
168       Gdb.Execute "INSERT INTO CATEGORIESANT (ID,PARENT,NAME) VALUES (" + str(r!ID) + "," + str(r!Parent) + ",'" + r!Name + "')"
170       r.MoveNext
          If f_err = 1 Then Exit Do
    Loop
172 r.Close



174 Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,1 FROM CATEGORIESANT WHERE PARENT=0"
176 Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,2  FROM CATEGORIESANT WHERE PARENT IN (SELECT ID FROM CATEGTEL)"

    ' GIA NA FYGEI  TO 1&2  EPIPEDO APO TO CATEGORIESANT
178 Gdb.Execute "DELETE FROM CATEGORIESANT WHERE ID IN (SELECT ID FROM CATEGTEL)"

180 Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,3  FROM CATEGORIESANT WHERE PARENT IN (SELECT ID FROM CATEGTEL)"

    ' GIA NA FYGEI  TO 3  EPIPEDO APO TO CATEGORIESANT
182 Gdb.Execute "DELETE FROM CATEGORIESANT WHERE ID IN (SELECT ID FROM CATEGTEL)"



184 Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,4  FROM CATEGORIESANT WHERE PARENT IN (SELECT ID FROM CATEGTEL)"

    ' GIA NA FYGEI  TO 4  EPIPEDO APO TO CATEGORIESANT
186 Gdb.Execute "DELETE FROM CATEGORIESANT WHERE ID IN (SELECT ID FROM CATEGTEL)"



188 Gdb.Execute "INSERT INTO CATEGTEL(ID,PARENT,NAME,AA) SELECT ID,PARENT,NAME,5  FROM CATEGORIESANT"



PARAKAMCH:



190 r.Open "SELECT * FROM CATEGTEL ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic
' 192 r.MoveFirst

    'r.Open "SELECT * FROM CATEG1", Gdb
194 Do While Not r.EOF
       
196    error = 0
       'paterades
198    MESS = "id=" + LTrim(str(r!ID)) + "Parent:" + LTrim(str(r!Parent)) + " " + r!Name
200    If r!Parent = 0 Then
202       T1.Nodes.Add , , "r" + LTrim(str(r!ID)), r!Name
       Else
    
204     T1.Nodes.Add "r" + LTrim(str(r!Parent)), tvwChild, "r" + LTrim(str(r!ID)), (r!Name)
       End If
       
206    If error = 0 Then
208         N = N + 1
210         names(N) = "r" + Format(r!ID, "0000") + "        " + r!Name
            FC_id(N) = r!ID
            FC_parenT(N) = r!Parent
            fc_names(N) = r!Name
       End If
212    r.MoveNext
    Loop
     
     
214  r.Close
fc_max = N
'============================= ΚΑΤΗΓΟΡΙΕΣ =====================================================







     
     '=========================================χαρακτηρας ============================================
     
216 sql = "select *from extra_field_values where epf_id=1 and languages_id=4 order by  epf_value"
    '  (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (7,3,4,0,1,'?????');
    
     Dim n7 As Long
218  n7 = 0
220  r.Open sql, gdblag, adOpenDynamic, adLockOptimistic
222  Do While Not r.EOF
224     xaraktiras.AddItem r!epf_value
226     f_xaraktiras(n7) = r!Value_ID
228     r.MoveNext
230     n7 = n7 + 1
           If f_err = 1 Then Exit Do
     Loop
     xaraktiras.AddItem "                "
     f_xaraktiras(n7) = 0
232  r.Close
234  xaraktiras.ListIndex = -1
     
     
     
    '========================================= FYLO ============================================
236 sql = "select *from extra_field_values where epf_id=3 and languages_id=4 order by epf_value"
    '  (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (7,3,4,0,1,'?????');
238  r.Open sql, gdblag, adOpenDynamic, adLockOptimistic
240  n7 = 0
242  Do While Not r.EOF
244     fylo.AddItem r!epf_value
246     f_fylo(n7) = r!Value_ID
248     n7 = n7 + 1
250     r.MoveNext
           If f_err = 1 Then Exit Do
     Loop
252  r.Close
254  fylo.ListIndex = 0
     
    '========================================= KATASKEYASTHS ============================================
256 sql = "SELECT * FROM manufacturers order by manufacturers_name"
     
258   r.Open sql, gdblag, adOpenDynamic, adLockOptimistic
260  n7 = 0
262  Do While Not r.EOF
264     kataskeyasths.AddItem r!manufacturers_name
266     f_kataskeyasths(n7) = r!manufacturers_id
268     n7 = n7 + 1
270     r.MoveNext
        If f_err = 1 Then Exit Do
     Loop
272  r.Close
274  kataskeyasths.ListIndex = 0
     
     
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
     
276   sql = "SELECT * FROM products_options_values where language_id=4  order by products_options_values_id "
278  n7 = 0
280   r.Open sql, gdblag, adOpenDynamic, adLockOptimistic
282  Do While Not r.EOF
284     hlikies.AddItem r!products_options_values_name
286     f_hlikies(n7) = r!products_options_values_id
288      n7 = n7 + 1
 
290     r.MoveNext
           If f_err = 1 Then Exit Do
     Loop
292  r.Close
294  kataskeyasths.ListIndex = 0
     
296     r.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic


         'R.Index = "pinakes"

         'FPA
         'R.Seek "=", 1, 1
298     Do While Not r.EOF
300         If r("typos") = 1 Then
302             Combo2.AddItem str(r("AYJON")) + " -> " + str(r("TIMH"))
             End If
304         r.MoveNext
            If f_err = 1 Then Exit Do
         Loop
         ' mon.metrhshs
306     r.Close
     
     Combo2.ListIndex = 0
     
     
     If f_err = 1 Then
         MsgBox "Δεν φορτώθηκαν σωστά οι πίνακες."
     End If
     
      
     
     ' Me.Print "Please use now mouse wheel to move this form."
308  Me.SHOW
310 ' ProcessMessages

  Text2.SetFocus








        '<EhFooter>
        Exit Sub
        
        
NOCONNECT:
   MsgBox "ΑΔΥΝΑΤΗ Η ΣΥΝΔΕΣΗ ΜΕ ΤΟ SITE. ΔΟΚΙΜΑΣΤΕ ΑΡΓΟΤΕΡΑ"
   Unload Me
   Exit Sub
Form_Load_Err:
'        MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.apot2scroll.Form_Load " & _
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
  Picture2.Move 0, Picture1.height + Picture1.Top, Me.ScaleWidth - 250, Me.ScaleHeight * 3
 'Picture2.Move 0, 2000, Me.ScaleWidth - 250, Me.ScaleHeight * 3
 
 Picture1.width = Picture2.width
 VScroll1.Left = Picture2.width
On Error Resume Next
VScroll1.height = Me.height - 580
 
 VScroll1.max = Abs(Me.ScaleHeight - Picture2.ScaleHeight)
 VScroll1.LargeChange = VScroll1.max / 10
 VScroll1.SmallChange = VScroll1.max / 100
 VScroll1.min = 0 'Picture1.height

Command1.Top = Picture2.height - 700
Command1.Caption = "Picture2.Height-700"

Picture2.Move 0, Picture1.height + Picture1.Top, Me.ScaleWidth - 250, Me.ScaleHeight * 3
End Sub


Private Sub foto2_DblClick(Index As Integer)
       '<EhHeader>
        On Error GoTo FOTO2_DBLClick_Err
        '</EhHeader>
    Dim F_EIKONA As Long

100 F_EIKONA = 11
110     CommonDialog1.FILENAME = "*.jpg"
120     CommonDialog1.ShowOpen
        Dim m As String
        On Error Resume Next
       
        
130     m = Trim(mID(CommonDialog1.FILENAME, InStrRev(CommonDialog1.FILENAME, "\") + 1, 30))
m = "a" + kodlabel.Caption + "-" + m  ' CommonDialog1.FILENAME ' Text1(0).Text
140     If F_EIKONA = 11 Then
150       ' m = m + ".JPG"
        End If
160     FileCopy CommonDialog1.FILENAME, F_FAKEL_KOINOS + "\images\" + m
    
170      foto2(Index).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + m)
            
180     StretchSourcePictureFromPicture apot2scroll.foto2(Index).Picture, apot2scroll.foto2(Index)
        PIC(Index).Caption = m
    
        '<EhFooter>
        Exit Sub

FOTO2_DBLClick_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.pelat2.Command14_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot2scroll.FOTO2_DBLClick " & "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub foto2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim ANS As Integer
   If Button = 2 Then
       ANS = MsgBox("Να καταργηθεί η εικόνα;", vbYesNo)
       If ANS = vbYes Then
           foto2(Index).Picture = LoadPicture()
       End If
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
   If Val(WEBPRICE.Text) = 0 Then
        WEBPRICE.Text = LTI5.Text
   End If

End Sub

Private Sub cmdmovenext_Click()
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
Dim r As New ADODB.Recordset

'"   στα νέα είδη αυτόματα να βάζει default βάρος=1κιλό & status=διαθέσιμο

r.Open "SELECT * FROM EID WHERE KOD='" + KOD.Text + "'", Gdb, adOpenDynamic, adLockOptimistic

If Not r.EOF() Then
   MsgBox "Υπάρχει ο κωδικός με το όνομα " + r!ONO
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
   
   F_id_LOCAL = EID.Recordset("ID")
   
   
 neo.Enabled = False
 KOD.Enabled = False
 ONO.Enabled = False
 upload1.Enabled = False
 
 
End Sub

Private Sub ONO_LostFocus()

If Len(ONO.Text) = 0 Then
   Exit Sub
End If

Dim r As New ADODB.Recordset
  r.Open "select TOP 100 * from EID WHERE ONO LIKE '" + ONO.Text + "%'", Gdb, adOpenDynamic, adLockOptimistic
  List1.Clear
  Do While Not r.EOF
   List1.AddItem Left(r("KOD") + Space(15), 15) + r("ONO")
   r.MoveNext
   ' Exit Do
 Loop
  r.Close
  
End Sub

Private Sub T1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim Y
Dim nam

If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty
Y = T1.SelectedItem.Index
'Label1.Caption = Str(y)
nam = T1.SelectedItem.Text
Me.Caption = nam & " - " & str(Y) + " names=" + names(Y)
  Set T1.DropHighlight = T1.SelectedItem
End If
End Sub

Private Sub Text2_LostFocus()
Dim r As New ADODB.Recordset

If Len(Text2.Text) = 0 Then
   Exit Sub
End If

If Left(Text2.Text, 1) = "/" Then
   Text2.Text = Trim(mID(Text2.Text, 2, 30))
   r.Open "select TOP 100 * from EID WHERE ERG LIKE '" + Text2.Text + "%'", Gdb, adOpenDynamic, adLockOptimistic
   List1.Clear
   Do While Not r.EOF
      List1.AddItem Left(r("KOD") + Space(15), 15) + r("ONO")
      r.MoveNext
     ' Exit Do
   Loop
   r.Close
   Exit Sub
Else
    r.Open "select * from EID WHERE KOD IN ( SELECT TOP 1 KOD FROM BARCODES WHERE ERG='" + Text2.Text + "')", Gdb, adOpenDynamic, adLockOptimistic
End If

 
If Not r.EOF Then
   KOD.Text = r(0)
   KOD.SetFocus
   FONO.SetFocus
   'UPDATE_CONTROLS
Else
   Exit Sub
End If
   
   F_id_LOCAL = r("ID")



End Sub

Private Sub upload1_Click()
If f_pel = "mpoyg" Then
   mpoyg_upload
   
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
   
   
Else
   lagak_upload
End If

End Sub

Sub mpoyg_upload()
'=====================================================================================================================
Dim a As String
Dim r As New ADODB.Recordset
 
Dim m_ID
Dim m_categ

Dim m_new As Integer
m_new = 0

 
 If FONO = "" Then
   MsgBox "Δεν επιλέξατε είδος "
   Exit Sub
 End If
 
Dim k As Integer
Dim sql As String
 m_categ = "942"
 
 
If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty
   Y = T1.SelectedItem.Index
   nam = T1.SelectedItem.Text
   m_categ = mID(names(Y), 2, 4)
   Set T1.DropHighlight = T1.SelectedItem
Else
   MsgBox "Δεν επιλέξατε κατηγορία "
   Exit Sub
End If

 
 
 
 
 EID.Recordset.Move 0  ' gia na ginei to rich1 opos to mseditor21
 
a = RichToHTML(Rich1, 0&, Len(Rich1.Text))  '(.TextRTF)
a = Replace(a, "'", "''")

gdblag.Close
gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131 ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
r.Open "select * from products where products_model='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic


M_DATE = Format(Now, "YYYY-MM-DD")
If r.EOF Then
   m_new = 1
   sql = "INSERT INTO products (products_youtube,products_quantity,products_model,products_image,products_subimage1,"
   sql = sql + " products_subimage2,products_subimage3,products_subimage4,products_subimage5,products_subimage6,products_price,"
   sql = sql + " products_date_added,products_last_modified,products_date_available,products_weight,products_status,products_tax_class_id,"
   sql = sql + " manufacturers_id,products_ordered) VALUES ('',6,'" + KOD + "','', "
   sql = sql + " NULL,NULL,NULL,NULL,NULL,NULL,'','" + M_DATE + "','" + M_DATE + "',NULL,'0.20',1,1,109,0); "

  ' Exit Sub
   gdblag.Execute sql
   r.Close
   r.Open "select * from products where products_model='" + KOD + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
End If

Dim site_foto(6) As String

site_foto(0) = CNull(r!products_image)
site_foto(1) = CNull(r!products_subimage1)
site_foto(2) = CNull(r!products_subimage2)
site_foto(3) = CNull(r!products_subimage3)
site_foto(4) = CNull(r!products_subimage4)
site_foto(5) = CNull(r!products_subimage5)
site_foto(6) = CNull(r!products_subimage6)





'Exit Sub

m_ID = str(r("products_id"))
r.Close

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
 Dim H As Single
  
 Dim kor As String
 Dim manuf As String
 
  kor = str(f_fylo(fylo.ListIndex))
  
  manuf = str(f_kataskeyasths(kataskeyasths.ListIndex))
  
  
  
  
 H = (100 + g_Fpa(Val(Left(Combo2.Text, 2)))) / 100
 
 
 mTimh = Format(Val(WEBPRICE.Text) / H, "####0.0000")



If IsNull(EID.Recordset("ERG")) Then
   M_ERG = " "
Else
   M_ERG = " #" + EID.Recordset("ERG")
End If



'm_new = 1
If m_new = 1 Then

   gdblag.Execute "insert into products_description(language_id,products_id,extra_value_id3,extra_value_id1,products_name,products_description) values (4," + m_ID + "," + kor + "," + b + ",'" + Replace(FONO.Text, "'", "''") + M_ERG + "','" + Replace(a, "'", "''") + "')"
Else

   r.Open "select * from  products_description  where products_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic
   If r.EOF Then
         gdblag.Execute "insert into products_description(language_id,products_id,extra_value_id3,extra_value_id1,products_name,products_description) values (4," + m_ID + "," + kor + "," + b + ",'" + Replace(FONO.Text, "'", "''") + M_ERG + "','" + Replace(a, "'", "''") + "')"
   Else
       gdblag.Execute "update products_description set  extra_value_id3=" + kor + ",extra_value_id1=" + b + " where products_id=" + m_ID
   End If
   r.Close
   
   ' ", products_name='" + FONO.Text + "', products_description='" + a + "'
End If

Dim energo As String
energo = "0"
If Check1.Value = vbChecked Then
   energo = "1"
End If

gdblag.Execute "update products set products_status=" + energo + ",manufacturers_id=" + manuf + ",products_price=" + Replace(mTimh, ",", ".") + ",products_weight=" + str(Val(WEIGHT)) + "  where products_id=" + m_ID
gdblag.Execute "update products set products_youtube='" + YOU_TUBE.Text + "',products_image='" + PIC(0).Caption + "',products_subimage1='" + PIC(1).Caption + "',products_subimage2='" + PIC(2).Caption + "' where products_id=" + m_ID
gdblag.Execute "update products set products_subimage3='" + PIC(3).Caption + "',products_subimage4='" + PIC(4).Caption + "',products_subimage5='" + PIC(5).Caption + "',products_subimage6='" + PIC(6).Caption + "' where products_id=" + m_ID


 'products_ship_price_two     products_ship_price                 products_ship_zip
 'ADDITIONALPRICE              SHIPPINGPRICE                         TEXT_PRODUCTS_ZIPCODE      YOU_TUBE  WEBPRICE
 'ADDIT_PRICE.text              INDV_PRICE.text                      ZIP_CODE.text
r.Open "select * from products_shipping where products_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic
If r.EOF Then
  gdblag.Execute "INSERT INTO products_shipping (products_id,products_ship_methods_id,products_ship_zip,products_ship_price,products_ship_price_two) VALUES (" + m_ID + ",0,'" + ZIP_CODE.Text + "','" + INDV_PRICE.Text + "','" + ADDIT_PRICE.Text + "');"
Else
  gdblag.Execute "update products_shipping set  products_ship_price_two='" + ADDIT_PRICE.Text + "',products_ship_price='" + INDV_PRICE.Text + "',products_ship_zip='" + ZIP_CODE.Text + "' where products_id=" + m_ID
End If




r.Close
r.Open "select * from products_to_categories where products_id=" + m_ID + " and  categories_id=" + m_categ, gdblag, adOpenDynamic, adLockOptimistic
If r.EOF Then
  gdblag.Execute "INSERT INTO products_to_categories (products_id,categories_id) VALUES (" + m_ID + "," + m_categ + ");"
End If
r.Close


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
gdblag.Execute "update products set products_tax_class_id=" + FPA + " where products_id=" + m_ID

'----------------------------------  CATEGORIES --------------------------------------------------------------------------------------------
r.Open "select count(*) from products_to_categories WHERE products_id=" + m_ID + " AND categories_id=" + m_categ + ";", gdblag, adOpenDynamic, adLockOptimistic
If r(0) > 0 Then
   ' ok yparxei den kano tipota
Else
   gdblag.Execute "insert into  products_to_categories ( products_id,categories_id) values (" + m_ID + "," + m_categ + ") ;"
End If
'----------------------------------  CATEGORIES --------------------------------------------------------------------------------------------



gdblag.Execute "DELETE FROM products_attributes where products_id=" + m_ID + "  limit 9", k
Dim sort As Integer
sort = 0
For k = 0 To hlikies.ListCount - 1
  
  If hlikies.Selected(k) Then
      
      gdblag.Execute "INSERT INTO products_attributes " _
      & "(products_id,options_id,options_values_id,options_values_price,price_prefix,products_options_sort_order) " _
      & " VALUES (" + m_ID + ",1," + str(f_hlikies(k)) + ",'0.0000',''," + str(sort) + ");"
      sort = sort + 1
  End If

Next k


   file = 3
   Open "c:\mercvb\send_foto" For Output As #file
   Print #file, "open www.toys-shop.gr"
   Print #file, "toyshopftp"
   Print #file, "T35m80)!@"
   Print #file, "binary"
   Print #file, "cd httpdocs/images"


Dim nc As Integer
nc = 0


For k = 0 To 6
   Me.Caption = "upload foto" + str(k)
   If Len(PIC(k).Caption) > 0 Then
      
      'την εχει ηδη την φωτογραφια  den kanv tipota
      If site_foto(k) = PIC(k).Caption Then
      Else
        If Len(site_foto(k)) > 1 Then
           Print #file, "del " + Replace(site_foto(k), "*", "-")  'βγαζω το αστερακι μην κανει ζημια μεγαλη
        End If
        
        Print #file, "send " + F_FAKEL_KOINOS + "\images\" + PIC(k).Caption
      
      End If
      
      nc = nc + 1
   End If
Next

   Print #file, "quit"
   Close #file

If nc > 0 Then
   Shell ("ftp -s:c:\mercvb\send_foto")
End If

MsgBox "OK"


End Sub 'P@ssw0rd#!


Sub ftp_Eikones(strLocal As String, strRemote As String)

    Dim lTimer    As Long

    'Dim strRemote As String

    'Dim strLocal  As String
   ' strRemote = "/httpdocs/images"  '             Text4.Text & "/" & Text5.Text
   ' strLocal = "c:\mercvb\images"   ' App.Path & "\Updates\" & Text5.Text
    

   Dim m2ftp As New cFTP
    
    If m2ftp.OpenConnection("toys-shop.gr", "toysftp", "kP8#H7o2fD") Then
   
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






Sub lagak_upload()
'=====================================================================================================================
Dim sql As String
'Dim gdblag As New ADODB.Connection
On Error GoTo 0
'(product_id, auto increment
sql = "INSERT INTO jos_vm_product (vendor_id,product_parent_id,product_sku,product_s_desc,product_desc,product_thumb_image,product_full_image,product_publish,product_weight,product_weight_uom,product_length,product_width,product_height,product_lwh_uom,product_url,product_in_stock,product_available_date,product_availability,product_special,product_discount_id,ship_code_id,cdate,mdate,product_name,product_sales,attribute,custom_attribute,product_tax_id,product_unit,product_packaging,child_options,quantity_options,child_option_ids,product_order_levels) "
'                     (vendor_id,product_parent_id,product_sku,   product_s_desc
sql = sql + " VALUES (1,0,'" + kodlabel + "','s_desc',"

'            ,product_desc,product_thumb_image,
sql = sql + "'" + ONOLABEL.Caption + "','resized/Vero_PC_4e94d2b5cd47c_110x110.jpg',"

'                 product_full_image    product_publish, product_weight,product_weight_uom,
sql = sql + "'4d2b5e2a8b.jpg','Y','0.0000','pounds',"
'              product_length,product_width,product_height,product_lwh_uom,product_url,product_in_stock,product_available_date
sql = sql + " '0.0000','0.0000','0.0000','inches','',-2,1318291200,'24h.gif',"
'         ,product_availability,product_special,product_discount_id,ship_code_id,cdate,mdate,product_name,product_sales,attribute,custom_attribute,product_tax_id,product_unit,product_packaging,child_options,quantity_options,child_option_ids,product_order_levels) "
sql = sql + " 'N',0,NULL,1318333812,1318376117,'" + ONOLABEL.Caption + "',3,'','',0,'piece',2,'Y,N,N,N,N,N,20%,10%,','hide,0,0,1','','0,0');"

gdblag.Execute sql

Dim r As New ADODB.Recordset
Dim m_ID As Long
r.Open "select product_id from jos_vm_product where product_sku='" + kodlabel + "' limit 1", gdblag, adOpenDynamic, adLockOptimistic
m_ID = r(0)


' T1.SelectedItem.Index

sql = "INSERT INTO jos_vm_product_category_xref (`category_id`,`product_id`,`product_list`) Values"
sql = sql + "( " + mID(names(T1.SelectedItem.Index), 2, 3) + "," + str(m_ID) + " ,1)"

gdblag.Execute sql



' key :  product_price_id
sql = "INSERT INTO jos_vm_product_price (product_id,product_price,product_currency,product_price_vdate,product_price_edate,cdate,mdate,shopper_group_id,price_quantity_start,price_quantity_end)"
sql = sql + " VALUES (" + str(m_ID) + ",'" + LTI5.Text + "','EUR',0,0,1318380271,1324931768,5,0,0);"

gdblag.Execute sql



sql = "INSERT INTO jos_vm_product_mf_xref (`product_id`,`manufacturer_id`) Values "
sql = sql + "(" + str(m_ID) + ",1);"

gdblag.Execute sql







'UPDATE `vasi`.`jos_vm_product`
'SET
'`product_id` = {product_id: INT},
'`vendor_id` = {vendor_id: INT},
'`product_parent_id` = {product_parent_id: INT},
'`product_sku` = {product_sku: VARCHAR},
'`product_s_desc` = {product_s_desc: VARCHAR},
'`product_desc` = {product_desc: TEXT},
'`product_thumb_image` = {product_thumb_image: VARCHAR},
'`product_full_image` = {product_full_image: VARCHAR},
'`product_publish` = {product_publish: CHAR},
'`product_weight` = {product_weight: DECIMAL},
'`product_weight_uom` = {product_weight_uom: VARCHAR},
'`product_length` = {product_length: DECIMAL},
'`product_width` = {product_width: DECIMAL},
'`product_height` = {product_height: DECIMAL},
'`product_lwh_uom` = {product_lwh_uom: VARCHAR},
'`product_url` = {product_url: VARCHAR},
'`product_in_stock` = {product_in_stock: INT},
'`product_available_date` = {product_available_date: INT},
'`product_availability` = {product_availability: VARCHAR},
'`product_special` = {product_special: CHAR},
'`product_discount_id` = {product_discount_id: INT},
'`ship_code_id` = {ship_code_id: INT},
'`cdate` = {cdate: INT},
'`mdate` = {mdate: INT},
'`product_name` = {product_name: VARCHAR},
'`product_sales` = {product_sales: INT},
'`attribute` = {attribute: TEXT},
'`custom_attribute` = {custom_attribute: TEXT},
'`product_tax_id` = {product_tax_id: INT},
'`product_unit` = {product_unit: VARCHAR},
'`product_packaging` = {product_packaging: INT},
'`child_options` = {child_options: VARCHAR},
'`quantity_options` = {quantity_options: VARCHAR},
'`child_option_ids` = {child_option_ids: VARCHAR},
'`product_order_levels` = {product_order_levels: VARCHAR}
'WHERE <where_condition>;













End Sub

Private Sub VScroll1_Change()
 VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
 Picture2.Top = -VScroll1.Value + Picture1.height + Picture1.Top
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
