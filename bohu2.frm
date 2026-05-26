VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form BOHU2 
   BackColor       =   &H00FF0000&
   Caption         =   "Παραμ.Προγράμματος"
   ClientHeight    =   8604
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8604
   ScaleWidth      =   15720
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Αποθήκευση-Εξοδος"
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5400
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      _ExtentX        =   23940
      _ExtentY        =   14774
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   529
      BackColor       =   16711680
      TabCaption(0)   =   "Γενικές Παράμετροι"
      TabPicture(0)   =   "bohu2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label12(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label10"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "user"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label23"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label24"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Shape1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label25"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label29"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label30"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label31"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text12(21)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "epitr_fpa"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "pict_timol(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "pict_posot(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "pict_ekpt(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "pel_ono"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "pel_die"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "pel_epa"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "enarxh"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Lhxh"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "MAFM"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "MYDATAPASSW"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "mydataIsDemo"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Check1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "myDataFromEID1"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "afm_doy"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "MYDATAUSER"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "mplafon(0)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "telpel"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "PELARIT"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "PELPOL"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "PELTK"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "GGPS"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "Παράμετροι Αποθήκης"
      TabPicture(1)   =   "bohu2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "M_PROMHU"
      Tab(1).Control(1)=   "m__analysh"
      Tab(1).Control(2)=   "mSyskeyasia"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Παράμετροι Τιμολόγησης"
      TabPicture(2)   =   "bohu2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "metiketes"
      Tab(2).Control(1)=   "yparxei_pos"
      Tab(2).Control(2)=   "tel_pos"
      Tab(2).Control(3)=   "eis_fpa"
      Tab(2).Control(4)=   "epitr_seir"
      Tab(2).Control(5)=   "lian_metr"
      Tab(2).Control(6)=   "kostologhsh"
      Tab(2).Control(7)=   "Timago_lt"
      Tab(2).Control(8)=   "prefix"
      Tab(2).Control(9)=   "eis_ekpt"
      Tab(2).Control(10)=   "auto_ayj"
      Tab(2).Control(11)=   "Label22"
      Tab(2).Control(12)=   "Label19"
      Tab(2).Control(13)=   "Label21"
      Tab(2).Control(14)=   "Label20"
      Tab(2).Control(15)=   "Label18"
      Tab(2).Control(16)=   "Label17"
      Tab(2).Control(17)=   "Label13"
      Tab(2).Control(18)=   "Label9"
      Tab(2).Control(19)=   "Label5"
      Tab(2).Control(20)=   "Label6"
      Tab(2).Control(21)=   "Label7"
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "Παράμετροι Εκτυπώσεων"
      TabPicture(3)   =   "bohu2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Check5"
      Tab(3).Control(1)=   "Check3"
      Tab(3).Control(2)=   "Check2"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Παράμετροι Λογιστικής"
      TabPicture(4)   =   "bohu2.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "mgconnect"
      Tab(4).Control(1)=   "mgdir"
      Tab(4).Control(2)=   "DIRLOG"
      Tab(4).Control(3)=   "Label8"
      Tab(4).Control(4)=   "Label16"
      Tab(4).Control(5)=   "Label11"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Παροχος"
      TabPicture(5)   =   "bohu2.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "userpos"
      Tab(5).Control(1)=   "pwpos"
      Tab(5).Control(2)=   "PAROXKEY"
      Tab(5).Control(3)=   "PAROXPASS"
      Tab(5).Control(4)=   "PAROXUSER"
      Tab(5).Control(5)=   "Label33"
      Tab(5).Control(6)=   "Label32"
      Tab(5).Control(7)=   "Label28"
      Tab(5).Control(8)=   "Label27"
      Tab(5).Control(9)=   "Label26"
      Tab(5).ControlCount=   10
      Begin VB.TextBox userpos 
         Height          =   375
         Left            =   -70080
         TabIndex        =   85
         Text            =   " "
         Top             =   3720
         Width           =   3735
      End
      Begin VB.TextBox pwpos 
         Height          =   375
         Left            =   -70080
         TabIndex        =   83
         Text            =   " "
         Top             =   4440
         Width           =   3735
      End
      Begin VB.CommandButton GGPS 
         Caption         =   "ΓΓΠΣ"
         Height          =   495
         Left            =   10320
         TabIndex        =   82
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox PELTK 
         Height          =   360
         Left            =   11160
         TabIndex        =   78
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox PELPOL 
         Height          =   360
         Left            =   8640
         TabIndex        =   77
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox PELARIT 
         Height          =   360
         Left            =   6600
         TabIndex        =   76
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox PAROXKEY 
         BackColor       =   &H0080C0FF&
         Height          =   375
         Left            =   -70080
         TabIndex        =   75
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox PAROXPASS 
         BackColor       =   &H0080C0FF&
         Height          =   375
         Left            =   -70080
         TabIndex        =   73
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox PAROXUSER 
         BackColor       =   &H0080C0FF&
         Height          =   375
         Left            =   -70080
         TabIndex        =   71
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox telpel 
         Height          =   375
         Left            =   2955
         TabIndex        =   69
         Top             =   1680
         Width           =   2925
      End
      Begin VB.TextBox mplafon 
         Height          =   360
         Index           =   0
         Left            =   8640
         MaxLength       =   5
         TabIndex        =   68
         Top             =   3600
         Width           =   1245
      End
      Begin VB.TextBox MYDATAUSER 
         Height          =   360
         Left            =   2940
         TabIndex        =   67
         Top             =   4080
         Width           =   1965
      End
      Begin VB.TextBox afm_doy 
         Height          =   360
         Left            =   8640
         TabIndex        =   65
         Top             =   825
         Width           =   3855
      End
      Begin VB.ComboBox myDataFromEID1 
         Height          =   288
         ItemData        =   "bohu2.frx":00A8
         Left            =   3000
         List            =   "bohu2.frx":00B2
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   5640
         Width           =   5290
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Δοκιμαστικό Περιβαλλον"
         Height          =   250
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   5290
      End
      Begin VB.CheckBox mydataIsDemo 
         Caption         =   "Δοκιμαστικό Περιβαλλον"
         Height          =   250
         Left            =   3000
         TabIndex        =   62
         Top             =   5280
         Width           =   5290
      End
      Begin VB.TextBox MYDATAPASSW 
         Height          =   360
         Left            =   2940
         TabIndex        =   59
         Top             =   4680
         Width           =   5325
      End
      Begin VB.TextBox MAFM 
         Height          =   360
         Left            =   2940
         TabIndex        =   57
         Top             =   3600
         Width           =   1245
      End
      Begin VB.TextBox mgconnect 
         Height          =   360
         Left            =   -72045
         TabIndex        =   56
         Top             =   1680
         Width           =   2445
      End
      Begin VB.ComboBox M_PROMHU 
         Height          =   315
         ItemData        =   "bohu2.frx":00E1
         Left            =   -74640
         List            =   "bohu2.frx":00EB
         TabIndex        =   54
         Text            =   "Combo1"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.ComboBox m__analysh 
         Height          =   315
         ItemData        =   "bohu2.frx":0119
         Left            =   -74640
         List            =   "bohu2.frx":0123
         TabIndex        =   53
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CheckBox mSyskeyasia 
         Alignment       =   1  'Right Justify
         Caption         =   "Συσκευασία"
         Height          =   375
         Left            =   -74760
         TabIndex        =   52
         Top             =   960
         Width           =   3615
      End
      Begin VB.CheckBox metiketes 
         Alignment       =   1  'Right Justify
         Caption         =   "Eτικέττες μετά την καταχώρηση του τιμολ.αγοράς"
         Height          =   375
         Left            =   -70440
         TabIndex        =   51
         Top             =   1800
         Width           =   4575
      End
      Begin VB.ComboBox yparxei_pos 
         Height          =   315
         ItemData        =   "bohu2.frx":017C
         Left            =   -68280
         List            =   "bohu2.frx":0189
         TabIndex        =   50
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox tel_pos 
         Height          =   315
         ItemData        =   "bohu2.frx":01C2
         Left            =   -68280
         List            =   "bohu2.frx":01CC
         TabIndex        =   49
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox eis_fpa 
         DataField       =   "metatopish_W"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -72105
         TabIndex        =   46
         Top             =   1035
         Width           =   1245
      End
      Begin VB.TextBox epitr_seir 
         DataField       =   "TIME_PREHEAT"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -72105
         TabIndex        =   45
         Top             =   645
         Width           =   1245
      End
      Begin VB.TextBox lian_metr 
         Height          =   360
         Left            =   -72105
         TabIndex        =   44
         Top             =   3360
         Width           =   1245
      End
      Begin VB.TextBox kostologhsh 
         DataField       =   "TIME_PREHEAT"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -72105
         TabIndex        =   41
         Top             =   2880
         Width           =   1245
      End
      Begin VB.TextBox Timago_lt 
         DataField       =   "TIME_PREHEAT"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -72105
         TabIndex        =   39
         Top             =   3840
         Width           =   1245
      End
      Begin VB.TextBox prefix 
         Height          =   360
         Left            =   -72105
         TabIndex        =   36
         Top             =   2400
         Width           =   1245
      End
      Begin VB.TextBox mgdir 
         DataField       =   "default_baros_panioy"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -72045
         TabIndex        =   33
         Top             =   1200
         Width           =   2445
      End
      Begin VB.TextBox DIRLOG 
         DataField       =   "AXRHSIMOPOIHTES_ORES"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -72045
         TabIndex        =   32
         Top             =   735
         Width           =   2445
      End
      Begin MSComCtl2.DTPicker Lhxh 
         Height          =   375
         Left            =   2955
         TabIndex        =   31
         Top             =   2670
         Width           =   1695
         _ExtentX        =   2985
         _ExtentY        =   656
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   186908673
         CurrentDate     =   37735
      End
      Begin MSComCtl2.DTPicker enarxh 
         Height          =   375
         Left            =   2955
         TabIndex        =   30
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2985
         _ExtentY        =   656
         _Version        =   393216
         Format          =   186908673
         CurrentDate     =   37735
      End
      Begin VB.TextBox pel_epa 
         DataField       =   "MPANIO"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   8640
         TabIndex        =   13
         Top             =   1695
         Width           =   3885
      End
      Begin VB.TextBox pel_die 
         DataField       =   "default_baros_panioy"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   2955
         TabIndex        =   12
         Top             =   1290
         Width           =   2925
      End
      Begin VB.TextBox pel_ono 
         DataField       =   "AXRHSIMOPOIHTES_ORES"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   2955
         TabIndex        =   11
         Top             =   825
         Width           =   2925
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "Numbers on Labels"
         Height          =   435
         Left            =   -74115
         TabIndex        =   10
         Top             =   2355
         Width           =   2610
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "Αλλάζω fonts"
         DataField       =   "allazo_fonts"
         DataSource      =   "Data1"
         Height          =   405
         Left            =   -74130
         TabIndex        =   9
         Top             =   1860
         Width           =   2610
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Αλλάζω χρώματα"
         DataField       =   "allazo_xromata"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -74145
         TabIndex        =   8
         Top             =   1395
         Width           =   2610
      End
      Begin VB.TextBox eis_ekpt 
         DataField       =   "WEIGHT_BOTTLE_MIN"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -72105
         TabIndex        =   7
         Top             =   1440
         Width           =   1245
      End
      Begin VB.TextBox auto_ayj 
         DataField       =   "WEIGHT_BOTTLE_MAX"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -72105
         TabIndex        =   6
         Top             =   1920
         Width           =   1245
      End
      Begin VB.TextBox pict_ekpt 
         Height          =   360
         Index           =   0
         Left            =   8640
         MaxLength       =   12
         TabIndex        =   5
         Top             =   2295
         Width           =   1245
      End
      Begin VB.TextBox pict_posot 
         Height          =   360
         Index           =   0
         Left            =   8640
         MaxLength       =   12
         TabIndex        =   4
         Top             =   2685
         Width           =   1245
      End
      Begin VB.TextBox pict_timol 
         Height          =   360
         Index           =   0
         Left            =   8640
         MaxLength       =   12
         TabIndex        =   3
         Top             =   3090
         Width           =   1245
      End
      Begin VB.TextBox epitr_fpa 
         Height          =   360
         Left            =   2955
         TabIndex        =   2
         Top             =   3090
         Width           =   1245
      End
      Begin VB.TextBox Text12 
         Height          =   360
         Index           =   21
         Left            =   8655
         MaxLength       =   7
         TabIndex        =   1
         Top             =   4080
         Width           =   1245
      End
      Begin VB.Label Label33 
         Caption         =   "PASSWORD POS"
         Height          =   255
         Left            =   -71760
         TabIndex        =   87
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label32 
         Caption         =   "USERPOS"
         Height          =   255
         Left            =   -71760
         TabIndex        =   86
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KΛΕΙΔΙ ΠΑΡΟΧΟΥ"
         Height          =   195
         Left            =   -71760
         TabIndex        =   84
         Top             =   2760
         Width           =   1560
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ΤΚ"
         Height          =   195
         Left            =   10440
         TabIndex        =   81
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Πόλη"
         Height          =   195
         Left            =   7320
         TabIndex        =   80
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αριθ."
         Height          =   195
         Left            =   6000
         TabIndex        =   79
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER PASSWORD"
         Height          =   195
         Left            =   -71760
         TabIndex        =   74
         Top             =   2040
         Width           =   1560
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME"
         Height          =   195
         Left            =   -71760
         TabIndex        =   72
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Τηλέφωνο Επιχείρησης"
         Height          =   195
         Left            =   720
         TabIndex        =   70
         Top             =   1680
         Width           =   1725
      End
      Begin VB.Shape Shape1 
         Height          =   1455
         Left            =   360
         Top             =   720
         Width           =   12375
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ΑΦΜ -ΔΟΥ (σε φόρμα)"
         Height          =   195
         Left            =   6720
         TabIndex        =   66
         Top             =   810
         Width           =   1665
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER MYDATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   885
         TabIndex        =   61
         Top             =   4200
         Width           =   1365
      End
      Begin VB.Label user 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ΚΩΔ.MYDATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   945
         TabIndex        =   60
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ΑΦΜ ΕΠΙΧΕΙΡΗΣΗΣ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   58
         Top             =   3660
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Driver database"
         Height          =   465
         Left            =   -74415
         TabIndex        =   55
         Top             =   1620
         Width           =   2280
      End
      Begin VB.Label Label22 
         Caption         =   "Οταν ζητείται ποσότητα μεγαλύτερη του υπολοίπου"
         Height          =   495
         Left            =   -70440
         TabIndex        =   48
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label19 
         Caption         =   "Ποσότητα τιμολόγησης "
         Height          =   495
         Left            =   -70440
         TabIndex        =   47
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Διαλέγω μετρητοίς/πίστωση στις λιανικές πωλήσεις =1 "
         Height          =   495
         Left            =   -74640
         TabIndex        =   43
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Κοστολόγηση μετά το πέρασμα τιμολογίου αγοράς=1"
         Height          =   405
         Left            =   -75000
         TabIndex        =   42
         Top             =   2880
         Width           =   2760
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Στα τιμολόγια αγοράς να προτείνεται η λιανική τιμή=1 "
         Height          =   405
         Left            =   -74880
         TabIndex        =   40
         Top             =   3840
         Width           =   2760
      End
      Begin VB.Label Label17 
         Height          =   375
         Left            =   -74880
         TabIndex        =   38
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Από αξία υπολογίζει ποσότητα=1 τιμή=0"
         Height          =   495
         Left            =   -74520
         TabIndex        =   37
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Κατάλογος Λογιστικής"
         Height          =   435
         Left            =   -74730
         TabIndex        =   35
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Κατάλογος εμπορικού"
         Height          =   285
         Left            =   -74760
         TabIndex        =   34
         Top             =   1230
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Επάγγελμα"
         Height          =   165
         Left            =   6840
         TabIndex        =   28
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Επωνυμία Επιχείρησης"
         Height          =   435
         Left            =   270
         TabIndex        =   27
         Top             =   810
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Διεύθυνση Επιχείρησης"
         Height          =   285
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Επιτρεπόμενες Σειρές"
         Height          =   285
         Left            =   -74925
         TabIndex        =   25
         Top             =   705
         Width           =   2760
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Να αλλάζει το ΦΠΑ"
         Height          =   285
         Left            =   -74880
         TabIndex        =   24
         Top             =   1050
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Εναρξη έτους"
         Height          =   285
         Left            =   210
         TabIndex        =   23
         Top             =   2295
         Width           =   2655
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Να αλλάζει η έκπτωση"
         Height          =   285
         Left            =   -74925
         TabIndex        =   22
         Top             =   1500
         Width           =   2760
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nα αυξάνεται η τιμή πώλησης όταν αυξάνει η τιμή αγοράς"
         Height          =   405
         Left            =   -74880
         TabIndex        =   21
         Top             =   1845
         Width           =   2655
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Δεκαδικά έκπτωσης"
         Height          =   375
         Index           =   0
         Left            =   5880
         TabIndex        =   20
         Top             =   2280
         Width           =   2700
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Δεκαδικά Ποσοτήτων"
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   19
         Top             =   2640
         Width           =   2700
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Δεκαδικά στην τιμή μονάδας"
         Height          =   375
         Index           =   2
         Left            =   5880
         TabIndex        =   18
         Top             =   3075
         Width           =   2700
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Πλαφόν σε πελάτες=1 "
         Height          =   375
         Index           =   3
         Left            =   5880
         TabIndex        =   17
         Top             =   3585
         Width           =   2700
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Επιτρεπόμενα ΦΠΑ"
         Height          =   285
         Left            =   255
         TabIndex        =   16
         Top             =   3150
         Width           =   2655
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Λήξη έτους"
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   2730
         Width           =   2655
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "....."
         Height          =   375
         Index           =   5
         Left            =   5940
         TabIndex        =   14
         Top             =   4065
         Width           =   2700
      End
   End
End
Attribute VB_Name = "BOHU2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

'Sub NOTES()
''Γενικές
''Global gpict_timol, gpict_posot, gcolor,, gm_elegx_ypol
''Global gkatalfpa(1 To 5), gRANTEBOY(1 To 20), gxron_enar, gmhn_enar
''Global  gtimer, gupdating, gpel5, gWN, gNW, gWPN, ghmera
''Global ghme_orist, gsmatim
''==========================Global greek
''==========================Global gProplDemo
''==========================Global gnComPort, gPrefix
''Global ghme_tel
''Global gka_ka, gparas_notes
''Global gm__analysh
''Global gXEIRISTHS, gm_ryumiseis, gAmarkar_par
''Global gegg_xpi, gDRX_TO_EURO
'
''--------------------------ηδη μπήκανε ---
''gpel_ono, gpel_epa, gpel_die,
'
''=========================   Πίνακες  ============================
'' gshmer, gmfpa1, gmfpa2, gmfpa3, gmfpa4, gmfpa5 '
''Global gmfpa6, gmfpa7, gmfpa8, gmfpa9
''Global g_Fpa(1 To 10)
'
'
'
'''Τιμολόγησης
''Global gINTZES, gDYO_TIMES, gm_log, gKODIKOS_2, gnotimol, gkeim2_timol, gtimh_synal ':=0
''Global gmPLAFON, gbuff, gA_proel(1 To 100), gektypoths, gseires_kenes
''Global gLIAN_METR, gANAL_FPA, gauto_ayj, gkeim_timol, gmikro_tim, geis_fpa, geis_ekpt, gmodel_ekt, gepitages
''Global gmETIKETES
'
''Global gtel_pos, gtimago_lt, gOO_P, goo, gARX_KOD_P, garx_kod, gdior_perigraf  ':=0
''Global gyparxei_pos, gpict_ekpt, gores, gBLOCK_UPDATE, gposos_paraggel,
''Global gepitr_fpa
''Global gXarEpik, gXaraKyrEpik, gXarKyrio
''Global gs_Ulian, gm_deltia(1 To 5), gjj_problem
''Global ggm_frashF1, gmikseires_kenes, garperith, gar11perith, gdiktyo
''Global gKOD_ONLY, gparas_skop
''========== ηδη
''gEPITR_SEIR
''
''
'''Λογιστικής
''Global
''Global gkatal_ag(39), grights(7, 12), gxeiristes(7)
''Global gkatal_fag(1 To 39)
''Global gkatal_po(1 To 39), gmcode
''Global gkatal_fpo(1 To 39), gupdating_kin, gDIG_3, gDIG_4, gok_gefyr
''============ ηδη καταχωρημένες
'', gdirlog , genas_kodikos_logistikhs
'
'
'
''
''================== Αποθήκης
''Global gs_Ueis, gs_Aeis, gs_Uejag, gs_Aejag, gs_Uaj_ejag, gs_Aaj_ejag, gs_Uaj_eisag, gs_Aaj_eisag, gs_olon, gs_APOU
''Global gs_Uxre, gs_Axre, gs_Upis, gs_Apis, gUEOR_APOU
''Global gkod_apou_enhm(1 To 12)
''Global gpar(1 To 40), gxart11, gxart08, gxart05, gdiaxvr, gm_POLLES_APOU, gselectable(1 To 40), gryumiseis(1 To 40)
''Global gbarcode, gmSYSKEYASIA, gkathg_eid(199)
''Global gm_apografh, gm_PROMHU
'''Aνθiμoυ
''Global gSOMA, gPROSOPO
''
''
''=============Εκτυπώσεις
''Global ginform, gprint_fonts, gar_Print(1 To 5), gyposel(1 To 5), g_selis, geject
''Global garit_seirvn, gseires_(10), gsynselidas, gTPmeFPA, gEJOYSIODOTHSH, gkostologhsh
''Global gTO_BLEPO_PRIN_TYPOUEI
'
'End Sub

Private Sub Command1_Click()

        'Διαφορετικά απο Clipper
        '==========================
        'r("pict") = pict_ekpt.Text ' einai se allh θεση και θέλη ρύθμιση
        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim TEL As Integer
        
        
        
        If Len(PELARIT.Text) = 0 Or Len(PELPOL.Text) = 0 Or Len(PELTK.Text) = 0 Then
        
         MsgBox "ΣΥΜΠΛΗΡΩΣΤΕ ΑΡΙΘΜΟ ΟΔΟΥ (0 ΑΝ ΔΕΝ ΕΧΕΙ), ΠΟΛΗ,ΤΚ"
         Exit Sub
         
        
        End If

        
                     Dim MERCPDF As Integer, nmax As Integer
                        MERCPDF = 0: nmax = 0
                        For Each xxx In Printers

                            If InStr(UCase(xxx.DeviceName), "MERCURYPDF") > 0 Then
                                ' Set printer as system default.
                                  MERCPDF = 1

                                ' Stop looking for a printer.
                                Exit For

                            End If
                            nmax = nmax + 1
                            If nmax > 50 Then
                               Exit For
                            End If
                        Next
        
                       If Len(Trim(PAROXKEY.Text)) > 5 Then
                           If MERCPDF = 0 Then
                              MsgBox "ΕΓΚΑΤΑΣΤΗΣΤΕ ΤΟΝ MERCURYPDF ΕΚΤΥΠΩΤΗ"
                           End If
                       End If
                       
        
        
        
        
        
        
        
        
        
        

        ' Set db = OpenDatabase(gDir, False, False, gConnect)
        Dim R   As New ADODB.Recordset

100     R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic

        'R.Edit
110     R("C3") = DIRLOG.Text
120     R("pelono") = pel_ono.Text
130     R("pelepa") = pel_epa.Text
140     R("peldie") = pel_die.Text
        R("telpel") = telpel.Text


        R("PELARIT") = PELARIT.Text
        R("PELPOL") = PELPOL.Text
        R("PELTK") = PELTK.Text


R("USERPOS") = userpos.Text
 R("PWDPOS") = pwpos.Text
       


        R("SEIRA1") = afm_doy.Text

  R("afmPEL") = MAFM.Text
  
  R!MYDATAUSER = MYDATAUSER.Text
  R!MYDATAPASW = MYDATAPASSW.Text
  
    
   R!mydataIsDemo = IIf(mydataIsDemo.Value = vbChecked, 1, 0)
  R!myDataFromEID1 = Left(myDataFromEID1.Text, 1)

        R("PAROXUSER") = PAROXUSER.Text
        R("PAROXPASS") = PAROXPASS.Text
        R("PAROXKEY") = PAROXKEY.Text
        
        'pict_timol(0).Text = r("pict")
        'pict_posot(0).Text = r("pictposot")

150     R("pict") = pict_timol(0).Text
160     R("pictposot") = pict_posot(0).Text

170     R("lianmetr") = lian_metr.Text
180     R("eisfpa") = eis_fpa.Text
190     R("eisekpt") = eis_ekpt.Text
200     R("timagolt") = Timago_lt.Text

210     R("telpos") = IIf(tel_pos.Text = tel_pos.List(0), 0, 1)

220     R("etiketes") = IIf(metiketes.Value = Checked, 1, 0)

230     R("AUTOAYJ") = auto_ayj.Text
240     R("n2") = IIf(yparxei_pos.Text = yparxei_pos.List(0), 0, IIf(yparxei_pos.Text = yparxei_pos.List(1), 1, 2))

250     R.Update

        '========================================================================
260     R.MoveNext    'go 2

        'R.Edit
270   '  R("hme") = enarxh.Value


        Dim MEMID As Long
        MEMID = R("ID")
        




        On Error GoTo Command1_Click_Err
        
        Gdb.Execute "UPDATE XRHSEIS SET ENARXH='" + Format(enarxh.Value, "MM/DD/YYYY") + "' WHERE ETOS=" + str(Year(gLhjh))

280     R("pict") = pict_ekpt(0).Text
290     R("n3") = IIf(mSyskeyasia = Checked, 1, 0)
300     R("C1") = mplafon(0).Text

310     R.Update
        '================================================
        'go 3
320     R.MoveNext
        'R.Edit
330     R("hme") = Lhxh.Value
340     R("C3") = mgdir.Text
350     R.Update

        '======================================================================
360     R.MoveNext    '  go 4

        'R.Edit

370     R("c4") = epitr_fpa.Text
380     R("C3") = epitr_seir.Text

390     R("c2") = IIf(M_PROMHU.Text = M_PROMHU.List(0), 0, 1)
400     R("bastim_pro") = prefix
410     R("n2") = kostologhsh
420     R("c1") = m__analysh.Text
430     R.Update

        '======================================================================

440     R.Close



 Gdb.Execute "UPDATE MEM SET HME='" + Format(enarxh.Value, "MM/DD/YYYY") + "' WHERE ID=" + str(MEMID)


450     Unload Me

460     If gApoMenu = True Then
470         SendKeys "%"
480         SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU2.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU2.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command1_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU2.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU2.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        Dim DB As Database

        Dim R  As New ADODB.Recordset

100     mForm_Load Me, fh, fw, ft, fl
110     Me.Picture = LoadPicture(gPicture)
        'Set db = OpenDatabase(gDir, False, False, gConnect)
120     R.Open "select * FROM MEM", Gdb, adOpenDynamic, adLockOptimistic

130     Me.KeyPreview = True

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        '========================================================================

        On Error Resume Next

        SSTab1.Tab = 0 'NA JEKINA APO TO TAB1


 
   MYDATAUSER.Text = R!MYDATAUSER
  MYDATAPASSW.Text = R!MYDATAPASW


 
   mydataIsDemo.Value = IIf(R!mydataIsDemo = 1, vbChecked, vbUnchecked)
  myDataFromEID1.Text = IIf(R!myDataFromEID1 = 1, myDataFromEID1.List(0), myDataFromEID1.List(1))


140     pel_ono.Text = R("pelono")
        MAFM.Text = R("afmPEL")
        
150     pel_epa.Text = R("pelepa")
160     pel_die.Text = R("peldie")



userpos.Text = R("USERPOS")
pwpos.Text = R("PWDPOS")
       

        PELARIT.Text = R("PELARIT")
        PELPOL.Text = R("PELPOL")
        PELTK.Text = R("PELTK")


        telpel.Text = R("telpel")
        afm_doy.Text = R("SEIRA1")
        

170     DIRLOG.Text = R("C3")

180     pict_timol(0).Text = R("pict")
190     pict_posot(0).Text = R("pictposot")

200     lian_metr.Text = R("lianmetr")


PAROXUSER.Text = R("PAROXUSER")
PAROXPASS.Text = R("PAROXPASS")
PAROXKEY.Text = R("PAROXKEY")






210     eis_fpa.Text = R("eisfpa")
220     eis_ekpt.Text = R("eisekpt")
230     Timago_lt.Text = R("timagolt")

240     If IsNull(R("telpos")) Then TEL = 0 Else TEL = R("telpos")
250     tel_pos.Text = tel_pos.List(TEL)

260     If IsNull(R("etiketes")) Then TEL = Unchecked Else TEL = Checked
270     metiketes.Value = TEL

280     If IsNull(R("n2")) Then TEL = 0 Else TEL = R("n2")
290     yparxei_pos.Text = yparxei_pos.List(TEL)

300     metiketes.Value = TEL
310     auto_ayj.Text = R("AUTOAYJ")

        '========================================================================
320     R.MoveNext    'go 2
330     pict_ekpt(0).Text = R("pict")
340     enarxh.Value = gEnarjh ' R("hme")

350     If IsNull(R("n3")) Then TEL = Unchecked Else TEL = IIf(R("n3") = 1, Checked, Unchecked)
360     mSyskeyasia.Value = TEL

370     mplafon(0).Text = R("c1")
        '========================================================================
        'go 3
380     R.MoveNext
390     Lhxh.Value = gLhjh ' R("hme")
400     mgdir.Text = R("C3")

        '======================================================================
410     R.MoveNext    '  go 4

420     mgconnect.Text = R("C3")

430     If IsNull(R("c2")) Then TEL = 0 Else TEL = Val(R("c2"))
440     M_PROMHU.Text = M_PROMHU.List(TEL)

450     epitr_fpa.Text = R("c4")
460     epitr_seir.Text = R("C3")

470     prefix = R("bastim_pro")
480     kostologhsh = R("n2")

490     m__analysh.Text = IIf(IsNull(R("c1")), 0, R("C1"))
        '========================================================================

500     R.Close

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU2.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU2.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    ' mForm_Resize Me, 6.2, 9930, 8580, fh, fw, ft, fl
    ' ResizeFormFor Me
    
End Sub

Private Sub GGPS_Click()

On Error Resume Next

   If Len(Trim(MAFM)) = 9 Then
   
   
     ' pelat1.neo_FIND_AFM Trim(Text1(5).Text), 1
      'Exit Sub
   
   
   
       'OK       Form1.parametros.Caption = "http://lagakis.gr/greek-vat/web-vat-details.php?myafm=028783755&afm=" + Trim(Text1(5).Text) ' 028783755"
      Dim u As String, p As String, AFM As String
      
      
     Dim f_EYRESHAFM_USER, f_EYRESHAFM_PWD, f_EYRESHAFM_AFM As String
      
      
    f_EYRESHAFM_USER = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_USER", "LAGAKIS23041960", "ΧΡΗΣΤΗΣ ΕΥΡΕΣΗΣ ΑΦΜ")
    f_EYRESHAFM_PWD = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_PWD", "d389833921", "ΚΩΔΙΚΟΣ ΕΥΡΕΣΗΣ ΑΦΜ")
    f_EYRESHAFM_AFM = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_AFM", "028783755", "ΑΦΜ ΕΠΙΧΕΙΡΗΣΗ")
       
    Dim ANS As Integer
    ANS = MsgBox("Να αντικατασταθούν τα στοιχεία από την Γενική Γραμματεία Πληροφ.Συστημάτων;", vbYesNo)
    If ANS = vbNo Then
       Exit Sub
    End If
      
      
      
    u = Trim(f_EYRESHAFM_USER)   ' = "LAGAKIS23041960"  ' Ww367868u916
    p = Trim(f_EYRESHAFM_PWD)  '  = "d389833921"   '080353
    AFM = Trim(f_EYRESHAFM_AFM)    ' = "028783755"   '051103283
      
    
 
      Dim EPO As String, DIE As String, EPA As String, pol As String, tk As String, doy As String
      Dim c As String
      c = MAFM 'Text1(5).Text
      Dim ENERGOS As Integer
      ENERGOS = 1
      Par7MyData.requestafm2 u, p, c, EPO, DIE, EPA, pol, tk, doy, ENERGOS
      
      
    Dim sos As Integer
    sos = 0
      
     If sos = 1 Then
        MsgBox "μη εγκυρο ΑΦΜ"
     Else
    
       pel_ono = Replace(EPO, """", "")
       pel_die = Replace(DIE, """", "")
       PELPOL = Replace(pol, """", "")
       pel_epa = Replace(EPA, """", "")
       afm_doy = "ΑΦΜ:" + MAFM + "  ΔΟΥ:" + Replace(doy, """", "")
      ' Me.Text1(7) = Replace(Me.Text1(7), "'", "`")
      ' Me.Text1(13) = Replace(tk, """", "")
   
    
     End If
      
      
      
      
      
      
      
      
      
      
      
      
      
    
   Else
       MsgBox "ΣΥΜΠΛΗΡΩΣΤΕ ΤΟ ΑΦΜ "
      ' Text1(5).SetFocus
       
   
   End If


































End Sub
