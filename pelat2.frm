VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form pelat2OLD 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   12750
   ClientLeft      =   525
   ClientTop       =   630
   ClientWidth     =   18675
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12750
   ScaleWidth      =   18675
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      DataField       =   "ADT"
      DataSource      =   "DATA1"
      Height          =   375
      Index           =   16
      Left            =   1200
      TabIndex        =   122
      Top             =   9840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "KINHTO"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   15
      Left            =   1200
      TabIndex        =   105
      Top             =   9360
      Width           =   3285
   End
   Begin VB.TextBox HMESYMF 
      DataField       =   "HM1"
      DataSource      =   "DATA1"
      Height          =   285
      Left            =   9120
      TabIndex        =   104
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox mhxanhmata 
      DataField       =   "NUM2"
      DataSource      =   "DATA1"
      Height          =   285
      Left            =   10440
      TabIndex        =   102
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox SYMFPOSO 
      DataField       =   "NUM1"
      DataSource      =   "DATA1"
      Height          =   285
      Left            =   8280
      TabIndex        =   99
      ToolTipText     =   "Αποστολή σε Σημειώσεις: F7 Συμφωνία  F5 Eτοιμο κείμενο  "
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox PELPROM 
      Height          =   315
      ItemData        =   "pelat2.frx":0000
      Left            =   7440
      List            =   "pelat2.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   98
      Top             =   0
      Width           =   1935
   End
   Begin VB.ComboBox kentroAdyn 
      Height          =   315
      ItemData        =   "pelat2.frx":0028
      Left            =   9600
      List            =   "pelat2.frx":0032
      Style           =   2  'Dropdown List
      TabIndex        =   97
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ADOMERCNEW.msEditor2 msEditor1 
      Height          =   4815
      Left            =   11160
      TabIndex        =   96
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8493
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "Εξοδος"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   4920
      Width           =   2055
   End
   Begin VB.PictureBox foto2 
      Height          =   3255
      Left            =   11160
      ScaleHeight     =   3195
      ScaleWidth      =   4275
      TabIndex        =   94
      Top             =   5400
      Width           =   4335
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Εύρεση αρχείου εικόνας"
      Height          =   300
      Left            =   11160
      TabIndex        =   93
      Top             =   8640
      Width           =   2280
   End
   Begin VB.TextBox Text1 
      DataField       =   "EMAIL"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   14
      Left            =   1200
      TabIndex        =   91
      Top             =   8760
      Width           =   3285
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "PVLHTHS"
      DataSource      =   "DATA1"
      Height          =   360
      Left            =   1200
      TabIndex        =   88
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ΔΙΑΓΡΑΦΗ ΑΡΧΕΙΟΥ"
      Height          =   225
      Left            =   13560
      TabIndex        =   86
      Top             =   9000
      Width           =   2040
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   11640
      TabIndex        =   85
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   11280
      TabIndex        =   83
      Top             =   9240
      Width           =   4320
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Αποστολή Αρχείου"
      Height          =   225
      Left            =   11280
      TabIndex        =   82
      Top             =   9000
      Width           =   2040
   End
   Begin VB.CommandButton Command7 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3960
      TabIndex        =   81
      Top             =   120
      Width           =   330
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16080
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   4560
      TabIndex        =   48
      Top             =   5400
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   4
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "ΠΕΔΙΑ ΗΜΕΡ/ΝΙΑΣ"
      TabPicture(0)   =   "pelat2.frx":0045
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lhm1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lhm2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lhm3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lhm4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lhm5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lhm6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lhm7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lhm8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lhm9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lhm10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lhm11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dtpicker1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dtpicker2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dtpicker3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dtpicker4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dtpicker5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dtpicker6"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "dtpicker7"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "dtpicker8"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "dtpicker9"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "dtpicker10"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dtpicker11"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "ΚΕΙΜΕΝΟΥ"
      TabPicture(1)   =   "pelat2.frx":0061
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "KEIM1"
      Tab(1).Control(1)=   "keim7"
      Tab(1).Control(2)=   "keim6"
      Tab(1).Control(3)=   "keim5"
      Tab(1).Control(4)=   "keim4"
      Tab(1).Control(5)=   "keim3"
      Tab(1).Control(6)=   "keim2"
      Tab(1).Control(7)=   "lab6"
      Tab(1).Control(8)=   "lab5"
      Tab(1).Control(9)=   "lab4"
      Tab(1).Control(10)=   "lab3"
      Tab(1).Control(11)=   "lab2"
      Tab(1).Control(12)=   "lab1"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "ΑΡΙΘΜΗΤΙΚΑ "
      TabPicture(2)   =   "pelat2.frx":007D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lnum4"
      Tab(2).Control(1)=   "Lnum3"
      Tab(2).Control(2)=   "Lnum2"
      Tab(2).Control(3)=   "Lnum1"
      Tab(2).Control(4)=   "Check1"
      Tab(2).Control(5)=   "Text5"
      Tab(2).Control(6)=   "Text6"
      Tab(2).Control(7)=   "Text7"
      Tab(2).Control(8)=   "Text8"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Πολ.Επιλογών"
      TabPicture(3)   =   "pelat2.frx":0099
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tComb1"
      Tab(3).Control(1)=   "tComb2"
      Tab(3).Control(2)=   "tComb3"
      Tab(3).Control(3)=   "tComb4"
      Tab(3).Control(4)=   "tComb5"
      Tab(3).Control(5)=   "Comb1"
      Tab(3).Control(6)=   "Comb2"
      Tab(3).Control(7)=   "Comb3"
      Tab(3).Control(8)=   "Comb4"
      Tab(3).Control(9)=   "Comb5"
      Tab(3).Control(10)=   "lbcb1"
      Tab(3).Control(11)=   "lbcb2"
      Tab(3).Control(12)=   "lbcb3"
      Tab(3).Control(13)=   "lbcb4"
      Tab(3).Control(14)=   "lbcb5"
      Tab(3).ControlCount=   15
      Begin VB.TextBox dtpicker11 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM11"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   3525
         TabIndex        =   132
         Top             =   5040
         Width           =   2000
      End
      Begin VB.TextBox dtpicker10 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM10"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   3525
         TabIndex        =   130
         Top             =   4680
         Width           =   2000
      End
      Begin VB.TextBox dtpicker9 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM9"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   3525
         TabIndex        =   128
         Top             =   4320
         Width           =   2000
      End
      Begin VB.TextBox dtpicker8 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM8"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   3525
         TabIndex        =   126
         Top             =   3960
         Width           =   2000
      End
      Begin VB.TextBox dtpicker7 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM7"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   3525
         TabIndex        =   124
         Top             =   3480
         Width           =   2000
      End
      Begin VB.TextBox tComb1 
         DataField       =   "COMB1"
         DataSource      =   "DATA1"
         Height          =   375
         Left            =   -69360
         TabIndex        =   121
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox tComb2 
         DataField       =   "COMB2"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -69360
         TabIndex        =   120
         Top             =   1800
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox tComb3 
         DataField       =   "COMB3"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -69360
         TabIndex        =   119
         Top             =   2280
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox tComb4 
         DataField       =   "COMB4"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -69360
         TabIndex        =   118
         Top             =   2760
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox tComb5 
         DataField       =   "COMB5"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -69360
         TabIndex        =   117
         Top             =   3240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox Comb1 
         Height          =   315
         Left            =   -72480
         TabIndex        =   111
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox Comb2 
         Height          =   315
         Left            =   -72480
         TabIndex        =   110
         Top             =   1800
         Width           =   3135
      End
      Begin VB.ComboBox Comb3 
         Height          =   315
         Left            =   -72480
         TabIndex        =   109
         Top             =   2280
         Width           =   3135
      End
      Begin VB.ComboBox Comb4 
         Height          =   315
         Left            =   -72480
         TabIndex        =   108
         Top             =   2760
         Width           =   3135
      End
      Begin VB.ComboBox Comb5 
         Height          =   315
         Left            =   -72480
         TabIndex        =   107
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox KEIM1 
         DataField       =   "CH1"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   90
         Top             =   780
         Width           =   3015
      End
      Begin VB.TextBox keim7 
         BackColor       =   &H0080C0FF&
         DataField       =   "CH7"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   89
         Top             =   3300
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H0080C0FF&
         DataField       =   "NUM1"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   -71790
         TabIndex        =   76
         Top             =   1395
         Width           =   2000
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H0080C0FF&
         DataField       =   "NUM2"
         DataSource      =   "DATA1"
         Height          =   315
         Left            =   -71790
         TabIndex        =   75
         Top             =   1830
         Width           =   2000
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H0080C0FF&
         DataField       =   "NUM3"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   -71790
         TabIndex        =   74
         Top             =   2280
         Width           =   2000
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H0080C0FF&
         DataField       =   "NUM4"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71790
         TabIndex        =   73
         Top             =   2760
         Width           =   2000
      End
      Begin VB.TextBox dtpicker6 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM6"
         DataSource      =   "DATA1"
         Height          =   300
         Left            =   3525
         TabIndex        =   72
         Top             =   3075
         Width           =   2000
      End
      Begin VB.TextBox dtpicker5 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM5"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   3525
         TabIndex        =   71
         Top             =   2685
         Width           =   2000
      End
      Begin VB.TextBox dtpicker4 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM4"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   3525
         TabIndex        =   70
         Top             =   2265
         Width           =   2000
      End
      Begin VB.TextBox dtpicker3 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM3"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   3525
         TabIndex        =   69
         Top             =   1785
         Width           =   2000
      End
      Begin VB.TextBox dtpicker2 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM2"
         DataSource      =   "DATA1"
         Height          =   315
         Left            =   3525
         TabIndex        =   68
         Top             =   1335
         Width           =   2000
      End
      Begin VB.TextBox dtpicker1 
         BackColor       =   &H0080C0FF&
         DataField       =   "HM1"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   3525
         TabIndex        =   67
         Top             =   900
         Width           =   2000
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ενεργός"
         DataField       =   "ENERGOS"
         DataSource      =   "DATA1"
         Height          =   255
         Left            =   -72840
         TabIndex        =   66
         Top             =   780
         Width           =   3615
      End
      Begin VB.TextBox keim6 
         BackColor       =   &H0080C0FF&
         DataField       =   "CH6"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   60
         Top             =   2940
         Width           =   3015
      End
      Begin VB.TextBox keim5 
         BackColor       =   &H0080C0FF&
         DataField       =   "CH5"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   59
         Top             =   2460
         Width           =   3015
      End
      Begin VB.TextBox keim4 
         BackColor       =   &H0080C0FF&
         DataField       =   "CH4"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   58
         Top             =   2100
         Width           =   3015
      End
      Begin VB.TextBox keim3 
         BackColor       =   &H0080C0FF&
         DataField       =   "CH3"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   57
         Top             =   1620
         Width           =   3015
      End
      Begin VB.TextBox keim2 
         BackColor       =   &H0080C0FF&
         DataField       =   "CH2"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   56
         Top             =   1140
         Width           =   3015
      End
      Begin VB.Label lhm11 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lhm11"
         Height          =   255
         Left            =   480
         TabIndex        =   133
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label lhm10 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lhm10"
         Height          =   255
         Left            =   480
         TabIndex        =   131
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label lhm9 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lhm9"
         Height          =   255
         Left            =   480
         TabIndex        =   129
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lhm8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lhm8"
         Height          =   255
         Left            =   480
         TabIndex        =   127
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label lhm7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lhm7"
         Height          =   255
         Left            =   480
         TabIndex        =   125
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label lbcb1 
         Caption         =   "Κειμ.Πολ 1"
         Height          =   255
         Left            =   -74760
         TabIndex        =   116
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lbcb2 
         Caption         =   "Κειμ.Πολ 2"
         Height          =   375
         Left            =   -74760
         TabIndex        =   115
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lbcb3 
         Caption         =   "Κειμ.Πολ 3"
         Height          =   255
         Left            =   -74760
         TabIndex        =   114
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbcb4 
         Caption         =   "Κειμ.Πολ 4"
         Height          =   255
         Left            =   -74760
         TabIndex        =   113
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label lbcb5 
         Caption         =   "Κειμ.Πολ 5"
         Height          =   255
         Left            =   -74760
         TabIndex        =   112
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Lnum1 
         Caption         =   "Aριθμ1"
         Height          =   255
         Left            =   -74835
         TabIndex        =   80
         Top             =   1395
         Width           =   2055
      End
      Begin VB.Label Lnum2 
         Height          =   255
         Left            =   -74835
         TabIndex        =   79
         Top             =   1905
         Width           =   2055
      End
      Begin VB.Label Lnum3 
         Height          =   255
         Left            =   -74835
         TabIndex        =   78
         Top             =   2370
         Width           =   2055
      End
      Begin VB.Label Lnum4 
         Height          =   255
         Left            =   -74835
         TabIndex        =   77
         Top             =   2790
         Width           =   2055
      End
      Begin VB.Label lab6 
         Height          =   255
         Left            =   -74640
         TabIndex        =   65
         Top             =   2940
         Width           =   2535
      End
      Begin VB.Label lab5 
         Height          =   255
         Left            =   -74640
         TabIndex        =   64
         Top             =   2460
         Width           =   2535
      End
      Begin VB.Label lab4 
         Height          =   255
         Left            =   -74640
         TabIndex        =   63
         Top             =   2100
         Width           =   2535
      End
      Begin VB.Label lab3 
         Height          =   255
         Left            =   -74640
         TabIndex        =   62
         Top             =   1620
         Width           =   2535
      End
      Begin VB.Label lab2 
         Height          =   255
         Left            =   -74640
         TabIndex        =   61
         Top             =   1140
         Width           =   2535
      End
      Begin VB.Label lab1 
         Caption         =   "lab1 - keim1"
         Height          =   255
         Left            =   -74880
         TabIndex        =   55
         Top             =   3300
         Width           =   2535
      End
      Begin VB.Label lhm6 
         Height          =   255
         Left            =   480
         TabIndex        =   54
         Top             =   3105
         Width           =   2055
      End
      Begin VB.Label lhm5 
         Height          =   255
         Left            =   480
         TabIndex        =   53
         Top             =   2730
         Width           =   2055
      End
      Begin VB.Label lhm4 
         Height          =   255
         Left            =   480
         TabIndex        =   52
         Top             =   2295
         Width           =   2055
      End
      Begin VB.Label lhm3 
         Height          =   255
         Left            =   480
         TabIndex        =   51
         Top             =   1875
         Width           =   2055
      End
      Begin VB.Label lhm2 
         Height          =   255
         Left            =   480
         TabIndex        =   50
         Top             =   1410
         Width           =   2055
      End
      Begin VB.Label lhm1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lhm1"
         Height          =   255
         Left            =   480
         TabIndex        =   49
         Top             =   900
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ετικεττα"
      Height          =   375
      Left            =   3480
      TabIndex        =   47
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Καρτέλλα"
      Height          =   375
      Left            =   4680
      TabIndex        =   46
      Top             =   4920
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo KATHGORIA 
      Bindings        =   "pelat2.frx":00B5
      DataField       =   "KODGAL"
      DataSource      =   "DATA1"
      Height          =   315
      Left            =   1200
      TabIndex        =   45
      Top             =   7800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "PERIGRAFH"
      BoundColumn     =   "AYJ"
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text1 
      DataField       =   "XRVMA"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   3675
      TabIndex        =   3
      Top             =   1560
      Width           =   840
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "ARPARAG"
      DataSource      =   "DATA1"
      Height          =   360
      Left            =   1200
      TabIndex        =   39
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "PLAISIO"
      DataSource      =   "DATA1"
      Height          =   375
      Left            =   1200
      TabIndex        =   38
      Top             =   6840
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "TYPOS"
      DataSource      =   "DATA1"
      Height          =   360
      ItemData        =   "pelat2.frx":00CA
      Left            =   1200
      List            =   "pelat2.frx":00DA
      TabIndex        =   37
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   3480
      TabIndex        =   36
      ToolTipText     =   "Aλλαγή κωδικού"
      Top             =   120
      Width           =   135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "pelat2.frx":0119
      Height          =   3450
      Left            =   4650
      TabIndex        =   35
      Top             =   1425
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6085
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "EPO"
         Caption         =   "ΕΠΩΝΥΜΟ"
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
         DataField       =   "DIE"
         Caption         =   "ΔΙΕΥΘΥΝΣΗ"
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
      BeginProperty Column02 
         DataField       =   "THL"
         Caption         =   "ΤΗΛΕΦΩΝΟ"
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
      BeginProperty Column03 
         DataField       =   "KOD"
         Caption         =   "ΚΩΔΙΚΟΣ"
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   8520
      Top             =   10080
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
   Begin VB.TextBox Text1 
      DataField       =   "ayp"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1032
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   1140
      TabIndex        =   32
      Top             =   5400
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      DataField       =   "typ"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1032
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   1140
      TabIndex        =   31
      Top             =   5880
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      DataField       =   "SHM2"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   1140
      TabIndex        =   10
      Top             =   4920
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      DataField       =   "PLAFON"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   1140
      TabIndex        =   9
      Top             =   4440
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      DataField       =   "PEK"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   1140
      TabIndex        =   8
      Top             =   3960
      Width           =   2200
   End
   Begin VB.CommandButton Command3 
      Height          =   400
      Left            =   5880
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Διαγραφή"
      Height          =   400
      Left            =   7170
      TabIndex        =   27
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "kod"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1140
      TabIndex        =   18
      Top             =   120
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      DataField       =   "EPO"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1140
      TabIndex        =   0
      Top             =   600
      Width           =   3285
   End
   Begin VB.TextBox Text1 
      DataField       =   "DIE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1140
      TabIndex        =   1
      Top             =   1080
      Width           =   3525
   End
   Begin VB.TextBox Text1 
      DataField       =   "POL"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   1140
      TabIndex        =   2
      Top             =   1560
      Width           =   2040
   End
   Begin VB.TextBox Text1 
      DataField       =   "EPA"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   1140
      TabIndex        =   4
      Top             =   2040
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      DataField       =   "AFM"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   1140
      TabIndex        =   5
      Top             =   2520
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      DataField       =   "THL"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   1140
      TabIndex        =   6
      Top             =   3000
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      DataField       =   "DOY"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   1140
      TabIndex        =   7
      Top             =   3480
      Width           =   2200
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   14
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   13
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   12
      Top             =   0
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   9360
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      DataSourceName  =   "MERCCOMPAQ"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "YPOKATHG"
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
      BackStyle       =   0  'Transparent
      Caption         =   "Α.Δ.Τ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   24
      Left            =   240
      TabIndex        =   123
      Top             =   9840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "KINHTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   106
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label symf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ημερ.Συμφωνίας"
      Height          =   195
      Index           =   2
      Left            =   7560
      TabIndex        =   103
      Top             =   960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label symf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Συμφ.μηχαν"
      Height          =   195
      Index           =   0
      Left            =   9240
      TabIndex        =   101
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label symf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Συμφ.Ποσό"
      Height          =   195
      Index           =   1
      Left            =   7200
      TabIndex        =   100
      Top             =   480
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   92
      Top             =   8880
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Πωλητης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   21
      Left            =   0
      TabIndex        =   87
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label ffbuff 
      Height          =   195
      Left            =   12840
      TabIndex        =   84
      Top             =   5160
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κατηγορία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   20
      Left            =   0
      TabIndex        =   44
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "T.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   19
      Left            =   3225
      TabIndex        =   43
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Τρόπος Πληρωμής"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   18
      Left            =   0
      TabIndex        =   42
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ΜΕΤΑΦΟΡΙΚΉ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   17
      Left            =   0
      TabIndex        =   41
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Καθεστώς ΦΠΑ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   16
      Left            =   0
      TabIndex        =   40
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Αρχ.Υπόλοιπο"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   15
      Left            =   120
      TabIndex        =   34
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Τελικό Υπόλοιπο"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   14
      Left            =   120
      TabIndex        =   33
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδικός Λογιστικής"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   120
      TabIndex        =   30
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Πλαφόν"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   29
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Εκπτωση % Tιμοκατάλογος"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Επωνυμία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Διεύθυνση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Πόλη"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Επάγγελμα"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Α.Φ.Μ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Τηλέφωνο"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Δ.Ο.Υ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A.Φ.Μ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   17
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Επωνυμία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   16
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "pelat2OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ELEGXOS AFM FResize

Dim f_KENTRA_ADYNATISMATOS As Integer
Dim F_CAN_SYGXONEYSI As Integer



Dim FappWord               As Word.Application

Dim FWRDdoc                As Word.Document

Dim F_ArPerEtik            As Long

Dim f_TIMOKAT              As Integer

Dim F_LHM1                 As String

Dim F_LHM2                 As String

Dim f_lHM3                 As String

Dim f_lHM4                 As String

Dim f_lHM5                 As String

Dim f_lHM6                 As String


Dim f_lHM7                 As String

Dim f_lHM8                 As String

Dim f_lHM9                 As String

Dim f_lHM10                 As String

Dim f_lHM11                 As String























Dim f_lab1                 As String

Dim f_lab2                 As String

Dim f_lab3                 As String

Dim f_lab4                 As String

Dim f_lab5                 As String

Dim f_lab6                 As String

Dim f_num1                 As String

Dim f_num2                 As String

Dim f_num3                 As String

Dim f_num4                 As String

Dim f_lbcb1                As String, f_comb1 As String

Dim f_lbcb2                As String, f_comb2 As String

Dim f_lbcb3                As String, f_comb3 As String

Dim f_lbcb4                As String, f_comb4 As String

Dim f_lbcb5                As String, f_comb5 As String

Dim F_DIGpel               As Integer

Dim f_EKPTCAPTION          As String 'Εκπτωση % Tιμοκατάλογος

Dim f_New_Record           As Boolean

Dim f_FirstTab             As Integer

Dim PELPROM_STRING         As String

Dim f_files                As String    ' arxeia office

Dim FSELSTART

Dim FSELLENGTH

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Dim El          As New FORMA

Dim m_IsLoading As Boolean

Dim PROTH

Private Sub Command14_Click()

        '<EhHeader>
        On Error GoTo Command14_Click_Err

        '</EhHeader>
        Dim F_EIKONA As Long

100     F_EIKONA = 11
110     CommonDialog1.FILENAME = "*.jpg"
120     CommonDialog1.ShowOpen

        Dim m As String

        On Error Resume Next

130     m = "e" + Text1(0).Text

140     If F_EIKONA = 11 Then
150         m = m + ".JPG"
        End If

160     FileCopy CommonDialog1.FILENAME, "c:\mercvb\images\" + m
    
170     foto2.Picture = LoadPicture("c:\mercvb\images\" + m)
            
180     StretchSourcePictureFromPicture Me.foto2.Picture, Me.foto2
    
        '<EhFooter>
        Exit Sub

Command14_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command14_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command14_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'xrvma tk
'typos
'plaisio
' pek ekptosi timokatalogoy
'shm2 kod.logistikis
'memo
'typos kathestos fpa
'plaisio metaforikh
'arparag tropos pliromis
'KODGAL KATHGORIA PELATH
'
Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
100     If Text1(1).Enabled Then    'ΑΝ ΕΙΣΑΙ ΣΤΗΝ ΔΙΟΡΘΩΣΗ
            'On Error Resume Next
            '     Data1.Recordset.UpdateBatch

        End If

        ' If Data1.Recordset.RecordCount = 1 Then

        On Error Resume Next

110     Data1.Recordset.Move 0

120     Unload Me
        'Me.Hide

130     If gApoMenu = True Then
140         SendKeys "%"
150         SendKeys "~"    'ENTER
160         SendKeys "{DOWN}"
        Else
170         Unload Me

        End If

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command10_Click()

        '<EhHeader>
        On Error GoTo Command10_Click_Err

        '</EhHeader>
        Dim sFileName As String

100     sFileName = File1.Path + "\" + File1.FILENAME

110     Kill sFileName

120     File1.Refresh

        '<EhFooter>
        Exit Sub

Command10_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command10_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command10_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

        Dim ANS  As Integer

        Dim neos As String, RECS As Integer

        Dim buff As String

100     buff = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)

        '  neos = InputBox("Δώσε τον νέο κωδικό ")
        Dim R As New ADODB.Recordset

110     R.Open "select count(*) from EGG WHERE KOD='" + Data1.Recordset("KOD") + "' and EIDOS='" + buff + "'", Gdb, adOpenForwardOnly, adLockReadOnly

120     If R(0) > 0 Then
130         MsgBox "Αδύνατη η διαγραφή γιατί υπάρχουν κινήσεις"
140         R.Close

            Exit Sub

        End If

        On Error Resume Next

150     ANS = MsgBox("Προσοχή θα διαγραφεί ο λογαριασμός " + Data1.Recordset("epo"), vbYesNo)

160     If ANS = vbYes Then
170         Data1.Recordset.delete

180         MHNYMA2.Timer1.Interval = 2000
190         MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " διεγράφη."
200         MHNYMA2.Label1.Caption = "Διαγραφή εγγραφής"
210         MHNYMA2.SHOW 1

            'MsgBox "Διεγράφη"
        End If

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>
        Dim k As Integer

        On Error GoTo ISLOCKED

100     If Command3.Caption = "Διόρθωση" Then
110         Command3.Caption = "Ενημέρωση"

120         For k = 1 To 15
130             Text1(k).Enabled = True
140             Text1(k).ForeColor = vbBlack
150             Text1(k).MaxLength = Data1.Recordset(Text1(k).DataField).DefinedSize
            Next

160         Text1(1).SetFocus
170         msEditor1.Enabled = True
180         Command7.Enabled = True
        Else
190         Command3.Caption = "Διόρθωση"
            '  DATA1.Recordset.UpdateBatch adAffectAll

            tComb1.Text = Comb1.Text
            tComb2.Text = Comb2.Text
            tComb3.Text = Comb3.Text
            tComb4.Text = Comb4.Text
            tComb5.Text = Comb5.Text

200         Data1.Recordset.Move 0

210         MHNYMA2.Timer1.Interval = 2000
220         MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " αποθηκεύτηκε με επιτυχία"
230         MHNYMA2.SHOW 1

            If f_KENTRA_ADYNATISMATOS = 1 Then

                On Error Resume Next

                Dim R5   As New ADODB.Recordset

                Dim ANS5 As Integer

                ' ANOIGO PROSOPO
                If kentroAdyn.ListIndex = 0 Then ' AN SVMA
                    R5.Open "SELECT * FROM PEL WHERE KOD='10" + Format(Val(Text1(0).Text), "0000") + "'", Gdb, adOpenDynamic, adLockOptimistic

                    If R5.EOF Then
                       
                        ANS5 = MsgBox("να ανοιχθεί και στο πρόσωπο;", vbYesNo)

                        If ANS5 = vbYes Then
                            Gdb.Execute "DROP TABLE DOKPELADYN"
                            Gdb.Execute "SELECT * INTO DOKPELADYN from PEL WHERE EIDOS='e' and KOD='" + Text1(0).Text + "'"
                            Gdb.Execute "ALTER TABLE DOKPELADYN DROP COLUMN ID"
                            Gdb.Execute "UPDATE DOKPELADYN SET NUM3=1, KOD='10" + Format(Val(Text1(0).Text), "0000") + "'"
                            Gdb.Execute "INSERT INTO PEL SELECT TOP 1 * FROM DOKPELADYN", k
                        End If
                    End If
                End If

                ' ANOIGO soma
                If kentroAdyn.ListIndex = 1 Then ' AN prosopo
                    R5.Open "SELECT * FROM PEL WHERE KOD='" + Format(Val(Right(Text1(0).Text, 4)), "0000") + "'", Gdb, adOpenDynamic, adLockOptimistic

                    If R5.EOF Then
                       
                        ANS5 = MsgBox("να ανοιχθεί και στο ΣΏΜΑ;", vbYesNo)

                        If ANS5 = vbYes Then

                            Dim N0 As Long

                            ' On Error Resume Next
                            Gdb.Execute "DROP TABLE DOKPELADYN"
                            Gdb.Execute "SELECT * INTO DOKPELADYN FROM PEL WHERE EIDOS='e' and  KOD='" + Text1(0).Text + "'"
                            Gdb.Execute "ALTER TABLE DOKPELADYN DROP COLUMN ID"
                            Gdb.Execute "UPDATE DOKPELADYN SET NUM3=0, KOD='" + Format(Val(Right(Text1(0).Text, 4)), "0000") + "'"
                            Gdb.Execute "INSERT INTO PEL SELECT TOP 1 * FROM DOKPELADYN", N0
                        End If
                    End If
                End If
               
            End If

240         Text2(0).SetFocus
        End If

        ' KEIM1.BackColor = vbGreen
        ' keim2.BackColor = vbGreen
    
250     keim2.Refresh

        Exit Sub

ISLOCKED:
260     MsgBox "Η εγγραφή ενημερώνεται από άλλο χρήστη"

270     Resume Next

        Exit Sub

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>
        Dim neos As String, RECS As Integer

        Dim buff As String

100     buff = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)
        'BUFF = IIf(PELPROM.ListIndex = 0, "e", "r")
110     neos = InputBox("Δώσε τον νέο κωδικό ")

        Dim R As New ADODB.Recordset

120     R.Open "select count(*),EPO from PEL WHERE KOD='" + neos + "' and EIDOS='" + buff + "'  GROUP BY EPO", Gdb, adOpenForwardOnly, adLockReadOnly

        Dim ANS As Integer, OKSYGX As Integer

130     OKSYGX = 0

        If R.EOF Then
            OKSYGX = 0
        Else
        
        
        If F_CAN_SYGXONEYSI = 0 Then
           ANS = MsgBox("ΥΠΑΡΧΕΙ ΗΔΗ O ΚΩΔΙΚΟΣ " + neos + " " + R(1))
           Exit Sub
        End If
        
        
        

140         If R(0) = 1 Then    ' ΒΡΕΘΗΚΕ ΚΑΙ ΑΛΛΟΣ
150             ANS = MsgBox("ΥΠΑΡΧΕΙ ΗΔΗ O " + R(1) + " ΝΑ ΣΥΓΧΩΝΕΥΘΟΥΝ; ", vbYesNo)

160             If ANS = vbYes Then
170                 OKSYGX = 1
                Else
180                 OKSYGX = 0

                    Exit Sub

                End If

            Else
190             OKSYGX = 0
            End If
        End If

200     If OKSYGX = 0 Then
210         Gdb.Execute "UPDATE PEL SET KOD='" + neos + "' WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'", RECS
220         MsgBox "ΑΡΧΕΙΟ ΠΕΛΑΤΩΝ " + Chr(13) + "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")
        Else
230         Gdb.Execute "delete from  PEL  WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'", RECS

240         If RECS > 0 Then
250             MsgBox "ΟΛΕΣ ΟΙ ΚΙΝΗΣΕΙΣ ΕΧΟΥΝ ΜΕΤΑΦΕΡΘΕΙ ΣΤΟΝ ΚΩΔΙΚΟ " + neos
            Else
260             MsgBox "Η ΔΙΑΔΙΚΑΣΙΑ ΔΙΕΚΟΠΗ"

                Exit Sub

            End If
        End If

270     Gdb.Execute "UPDATE EGG SET KOD='" + neos + "' WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'", RECS
280     MsgBox "ΑΡΧΕΙΟ ΚΙΝΗΣΕΩΝ ΑΠΟΘΗΚΗΣ " + Chr(13) + "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")

        'Gdb.Execute "UPDATE TIM SET PELKOD='" + neos + "' WHERE EIDOS='" + BUFF + "' AND PELKOD='" + Text1(0).Text + "'", RECS
290     MsgBox "ΑΡΧΕΙΟ TIMOΛΟΓΙΩΝ " + Chr(13) + "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")

300     Gdb.Execute "UPDATE TIM SET KPE='" + neos + "' WHERE EIDOS='" + buff + "' AND KPE='" + Text1(0).Text + "'", RECS

310     MsgBox "ΑΡΧΕΙΟ TIMOΛΟΓΙΩΝ " + Chr(13) + "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")

320     Gdb.Execute "UPDATE EGGTIM SET PELKOD='" + neos + "' WHERE EIDOS='" + buff + "' AND PELKOD='" + Text1(0).Text + "'", RECS
330     MsgBox "ΑΡΧΕΙΟ ΚΙΝΗΣΕΩΝ ΑΠΟΘΗΚΗΣ " + Chr(13) + "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")

        ' Else
        '   MsgBox "Υπάρχει ο κωδικός " + neos
        'End If
        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>
        Dim fff As New pelat4

        'fff.SHOW

        ' fff.PELPROM.Text = PELPROM.Text
100     fff.ffbuff = IIf(Left(PELPROM.Text, 1) = "1", "e", "r") + Text1(0).Text
        'fff.ffbuff = IIf(Left(PELPROM.Text, 1) = "1", "e", "r") + Text1(0).Text
110     fff.Command7.Enabled = False
120     fff.APOALLO = "1"

        If f_KENTRA_ADYNATISMATOS = 1 Then
            If kentroAdyn.ListIndex >= 0 Then
                fff.Text2(0).Text = Data1.Recordset("KOD")
                fff.Data1.ConnectionString = gConnect
                fff.kentroAdyn.ListIndex = kentroAdyn.ListIndex ' Comb
            End If
   
        End If

130     fff.resizing = 0
140     fff.SHOW

        Exit Sub

150     pelat4.PELPROM.Text = PELPROM.Text
160     pelat4.SHOW

        ' pelat4.TEXT2(0).Text = Text1(0)

170     pelat4.APOALLO = "1"
180     pelat4.Data1.ConnectionString = gConnect

190     pelat4.Text2(0).Text = Text1(0)

        '  pelat4.Data1.Refresh

200     SendKeys "~"    'ENTER
210     SendKeys "~"    'enter  {DOWN}"

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        ' ETIKETTA ME KEIMENO
        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>

        Dim ll, apoSCANNER As Long

        Dim DB

        Dim PSIFIA As Integer

        Dim fores  As Long, k As Integer

        Dim x      As Printer

        Dim barc(200)

        Dim mONOMA, mBarcode, mERG, mLTI5

        Dim R  As New ADODB.Recordset

        Dim RD As DAO.Recordset

        On Error GoTo Command6_Click_Err

100     mONOMA = Trim(Text1(1).Text)    'ONOMA
110     mBarcode = Text1(0).Text    '  ' BARCODE
120     mERG = Text1(6).Text    ' erg
130     mLTI5 = Format(Val(Text1(13).Text), "###0.00")    ' LTI5

        On Error Resume Next

140     For Each x In Printers

150         If InStr(x.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
160             Set Printer = x
                ' Stop looking for a printer  5200016000239

                Exit For

            End If

        Next

170     apoSCANNER = 1
        'If Scanner Then

        ' μπουγουδης
        'Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
        'Set RD = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM= 1234")
        'If RD.RecordCount = 0 Then
        '   MsgBox "δεν υπάρχουν εγγραφές"
        '   Exit Sub
        'End If

        '  RD.MoveFirst
180     k = 0
190     apoSCANNER = k
        ' End If

        Dim Z

        Dim OSSA As Long

200     OSSA = 0

210     If Len(Dir("C:\OSSA.TXT", vbNormal)) > 2 Then
220         OSSA = 1
        End If

230     Z = F_ArPerEtik

240     fores = Val(InputBox("Πόσες εττικέτες; ", "", 1))

        Dim mFORES As Long

250     For ll = 1 To 1

260         mFORES = fores

            'If Scanner Then
            '   R.Close
            '   R.Open "select * FROM EID WHERE KOD='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

            ' If R.EOF Or Len(Trim(barc(ll))) = 0 Then
            '     mFORES = 0
            ' Else
            '  mONOMA = R("ONO") 'ONOMA
            '  mBarcode = R("KOD") '  ' BARCODE
            '  mERG = R("ERG") ' erg
            '  mLTI5 = Format(R("LTI5"), "###0.00") ' LTI5
            ' End If
            'End If

            '  Dim Z 'As Integer
            ' Z = 1    ' InputBox("METATOPISH")
        
270         For k = 1 To mFORES

280             If OSSA = 0 Then ' ΝΟΡΜΑΛ ΕΚΤΥΠΩΣΗ ΣΕ ΜΙΚΡΕΣ ΕΤΙΚΕΤΕΣ
290                 Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992

300                 If Len(mONOMA) > 16 Then
310                     Printer.FontSize = 10
                    Else
320                     Printer.FontSize = 12
                    End If

330                 Printer.CurrentY = 0
340                 Printer.FontBold = True
350                 Printer.Print Tab(4 + Z); Left(mONOMA, 25)    'Text1(1).Text 'ONOMA
                    'Printer.Print Tab(4 + Z); Mid$(mONOMA, 17, 16) 'Text1(1).Text 'ONOMA
                    'Printer.Print Tab(4 + Z); Mid$(mONOMA, 33, 16) 'Text1(1).Text 'ONOMA
360                 Printer.FontSize = 10
370                 Printer.Print Tab(4 + Z); Text1(2).Text;   ' DIE
380                 Printer.Print Tab(4 + Z); Text1(3).Text;   ' POL
390                 Printer.Print Tab(4 + Z); "THΛ."; Text1(6).Text;    ' THL
400                 Printer.EndDoc
                Else
410                 Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992
420                 Printer.FontSize = 8
430                 Printer.CurrentY = 0
440                 Printer.FontBold = True
450                 Printer.Print Tab(4 + Z); "Ossa Roxani & Sia OE"
460                 Printer.Print Tab(4 + Z); "M.Alexandrou 39 "
470                 Printer.Print Tab(4 + Z); "66200 Mikropoli"
480                 Printer.Print Tab(4 + Z); "Greece"
490                 Printer.Print
500                 Printer.Print
510                 Printer.Print
520                 Printer.FontSize = 12
530                 Printer.Print Tab(4 + Z); Left(keim2, 50)    'Text1(1).Text 'ONOMA

                     Dim jk As Integer: jk = 70
                     Dim aseira As String, bseira As String
                     aseira = keim3.Text
                    If Len(Trim(keim3.Text)) > 70 Then
                      jk = InStr(58, keim3.Text, " ")
                      If jk = 0 Then jk = 70 'σημαινει δεν εχει κενα στο τελος
                      aseira = Left(keim3.Text, jk)
                      bseira = mID$(keim3.Text, 71, 70)
                      Printer.Print Tab(4 + Z); aseira;   ' DIE DATA1.Recordset("die")
                      Printer.Print Tab(4 + Z); bseira;
                    Else
                      Printer.Print Tab(4 + Z); aseira;   ' DIE DATA1.Recordset("die")
                    End If
                     
540                 'Printer.Print Tab(4 + Z); Left(keim3.Text, jk);   ' DIE DATA1.Recordset("die")
550                 Printer.Print Tab(4 + Z); keim4.Text;   ' DIE
560                 Printer.Print Tab(4 + Z); keim5.Text;   ' POL
570                 Printer.Print Tab(4 + Z); keim6.Text;   ' POL
580                 Printer.Print Tab(4 + Z); "TEL."; keim7.Text;   ' POL
590           '      Printer.Print Tab(4 + Z); "TEL."; Text1(6).Text;    ' THL
600                 Printer.EndDoc
                End If

610         Next k

620     Next ll

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command7_Click()

        '<EhHeader>
        On Error GoTo Command7_Click_Err

        '</EhHeader>
        Dim c As String

100     c = get_next_pel("e")

110     Gdb.Execute "insert into PEL (EIDOS,KOD) VALUES ('e','" + c + "')"

120     Text2(0).Text = c
130     Text1(0).Text = c
140     Text2(0).SetFocus

150     Data1.RecordSource = "select *from PEL where  KOD='" + c + "'  AND EIDOS='e' ;"
160     Data1.Refresh

170     Command7.Enabled = False
        '       f_New_Record = True
        '
        '
        '       For K = 0 To 11
        '          Text1(K).Enabled = True
        '          Text1(K).ForeColor = vbBlack
        '       Next

180     Command3.Caption = "Διόρθωση"
        'Command3_Click

190     DataGrid1.SetFocus
200     SendKeys "{enter}{enter}"

        '<EhFooter>
        Exit Sub

Command7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command8_Click()

        '<EhHeader>
        On Error GoTo Command8_Click_Err

        '</EhHeader>
        Dim MF As String

        'ΑΝ ΔΕΝ ΔΙΑΛΕΞΑ ΠΕΛΑΤΗ ΝΑ ΦΕΥΓΕΙ
100     If Len(Trim(Text1(0))) = 0 Then

            Exit Sub

        End If

        '\\Server\data

110     CommonDialog1.ShowOpen
        'If PELPROM.ListIndex = 0 Then
120     MF = f_files + "\MERCVB\files\" + mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1) + Text1(0)
        ' Else
        '   mF = f_files + "\MERCVB\files\r" + Text1(0)
        ' End If

        On Error Resume Next

130     If Len(Dir(f_files, vbDirectory)) = 0 Then
140         MkDir f_files + "\MERCVB\files"
        End If

150     If Len(Dir(MF, vbDirectory)) = 0 Then
160         MkDir MF
        End If

        Dim mf2 As Integer, files2 As String

        On Error Resume Next

170     mf2 = InStrRev(CommonDialog1.FILENAME, "\")
180     files2 = mID(CommonDialog1.FILENAME, mf2 + 1, Len(CommonDialog1.FILENAME) - mf2)

190     FileCopy CommonDialog1.FILENAME, MF + "\" + files2
        'MILSEC 1000
        '       File1.Refresh

        REFRESH_FILES

        Exit Sub

Command8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub REFRESH_FILES()

        Dim N  As Integer

        Dim MF As String
        
100     N = PELPROM.ListIndex

110     If N < 0 Then N = 0
120     MF = UCase(f_files + "\MERCVB\files\" + mID(PELPROM_STRING, N + 1, 1) + Text1(0))

        'Else
        '   mF = UCase(f_files + "\MERCVB\files\r" + Text1(0))
        'End If
        On Error Resume Next

130     If Len(Dir(MF, vbDirectory)) > 0 Then
140         File1.Path = MF
        Else
150         File1.Path = "c:\mercvb\files\dum"
        End If

160     File1.Refresh

End Sub

Private Sub Command9_Click()

        '<EhHeader>
        On Error GoTo Command9_Click_Err

        '</EhHeader>

100     msEditor1.SetFocus
110     SendKeys Format(Now, "dd/MM/YYYY")
120     msEditor1.SetFocus

        '<EhFooter>
        Exit Sub

Command9_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command9_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command9_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Command9_Click()
'  Dim mf As String
'
'  'ΑΝ ΔΕΝ ΔΙΑΛΕΞΑ ΠΕΛΑΤΗ ΝΑ ΦΕΥΓΕΙ
'  If Len(Trim(Text1(0))) = 0 Then
'     Exit Sub
'  End If
'
'
'
'
'
'  If Option1(1).Value = True Then
'     mf = UCase(f_files + "\MERCVB\files\e" + Text1(0))
'  Else
'     mf = UCase(f_files + "\MERCVB\files\r" + Text1(0))
'  End If
'
'  If Len(Dir(mf, vbDirectory)) = 0 Then
'      MkDir mf
'  End If
'
'
'
'
'  CommonDialog1.InitDir = mf
'
'  CommonDialog1.ShowOpen
'
'   Dim sFileName As String
'
'
'sFileName = CommonDialog1.FILENAME
'Set FappWord = New Word.Application
'Set FWRDdoc = FappWord.Documents.Open(sFileName)
'FappWord.Visible = True
'
'
'
'End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DataGrid1_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
            'keybd_event VK_TAB, 0, 0, 0 'send a tab
            MDataGrid1_RowColChange
120         Command3.SetFocus

        End If

        '<EhFooter>
        Exit Sub

DataGrid1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.DataGrid1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.DataGrid1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    MDataGrid1_RowColChange
End Sub

Private Sub MDataGrid1_RowColChange()

        '<EhHeader>
        On Error GoTo DataGrid1_RowColChange_Err

        '</EhHeader>

        Dim MF As String

        'If Option1(1).value = True Then
        On Error Resume Next

        Dim N As Integer

100     N = PELPROM.ListIndex

110     If N < 0 Then N = 0
120     MF = UCase(f_files + "\MERCVB\files\" + mID(PELPROM_STRING, N + 1, 1) + Text1(0))

        'Else
        '   mF = UCase(f_files + "\MERCVB\files\r" + Text1(0))
        'End If
        On Error Resume Next

130     If Len(Dir(MF, vbDirectory)) > 0 Then
140         File1.Path = MF
        Else
150         File1.Path = "c:\mercvb\files\dum"
        End If

160     File1.Refresh

        Dim PI

170     PI = "C:\MERCVB\IMAGES\e" + Trim(Text1(0).Text) + "*" + ".JPG"

180     If Len(Dir(PI)) > 0 Then
190         foto2.Picture = LoadPicture("c:\mercvb\images\" + Dir(PI))
            
200         StretchSourcePictureFromPicture Me.foto2.Picture, Me.foto2
            
            'foto2.PaintPicture foto.Picture, 0, 0, foto2.width, foto2.height, 0, 0, foto.width, foto.height
            'foto.PaintPicture foto.Picture, 0, 0, foto.width, foto.height, 0, 0, foto.width, foto.height
        Else
210         foto2.Picture = LoadPicture()
            'Picture2.Picture = LoadPicture()
        End If

        Set msEditor1.mDataSource = Data1  ' for code you have to use the 'adoNotes' instead of the Adodc1
        msEditor1.MaxLength = 8000          ' Optional maximum length.
        '
        msEditor1.mDataField = "MEMO" '  "nNotes"     ' The field of your text
        msEditor1.Enabled = True

        If IsNull(Data1.Recordset("COMB1")) Then
            Comb1.Text = ""
        Else
            Comb1.Text = Data1.Recordset("COMB1")
        End If
            
        If IsNull(Data1.Recordset("COMB2")) Then
            Comb2.Text = ""
        Else
            Comb2.Text = Data1.Recordset("COMB2")
        End If

        If IsNull(Data1.Recordset("COMB3")) Then
            Comb3.Text = ""
        Else
            Comb3.Text = Data1.Recordset("COMB3")
        End If

        If IsNull(Data1.Recordset("COMB4")) Then
            Comb4.Text = ""
        Else
            Comb4.Text = Data1.Recordset("COMB4")
        End If

        If IsNull(Data1.Recordset("COMB5")) Then
            Comb5.Text = ""
        Else
            Comb5.Text = Data1.Recordset("COMB5")
        End If

        '<EhFooter>
        Exit Sub

DataGrid1_RowColChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.DataGrid1_RowColChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.DataGrid1_RowColChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub File1_DblClick()

        '<EhHeader>
        On Error GoTo File1_DblClick_Err

        '</EhHeader>

        Dim sFileName As String

        Dim xl        As New excel.Application

        Dim xlsheet   As excel.Worksheet

        Dim xlsheet3  As excel.Worksheet

        Dim xlwbook   As excel.workbook

        '
        ' sFileName = File1.Path + "\" + File1.FILENAME
100     sFileName = File1.Path + "\" + File1.FILENAME

        'Exit Sub

110     If UCase(Right(sFileName, 3)) = "DOC" Or UCase(Right(sFileName, 4)) = "DOCX" Then
120         Set FappWord = New Word.Application
130         Set FWRDdoc = FappWord.Documents.Open(sFileName)
140         FappWord.Visible = True

            Exit Sub

        End If

150     If UCase(Right(sFileName, 4)) = "XLSX" Or UCase(Right(sFileName, 3)) = "XLS" Or UCase(Right(sFileName, 3)) = "TXT" Then

160         Set xlwbook = xl.Workbooks.Open(sFileName)
170         MILSEC 2000
180         Set xlsheet = xlwbook.Sheets.Item(1)
190         xl.Visible = True
            '  Set workbook = excel.Workbooks.Add
            '
200         xlwbook.Activate

            Exit Sub

        End If

        Shell "EXPLORER.EXE " + sFileName, vbMaximizedFocus

        '<EhFooter>
        Exit Sub

File1_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.File1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.File1_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "PELAT2"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command1_Click

110     If KeyCode = vbKeyF10 Then
120         ut_rantebou.Label1 = Text1(1).Text

130         ut_rantebou.SHOW 1
        End If
        
        If f_KENTRA_ADYNATISMATOS = 1 Then

            If KeyCode = vbKeyF5 Then

                Dim a

                a = Chr(13) + "ΑΡΙΣ.ΓΟΝΑΤΟ=" + Chr(13) + "ΑΡΙΣ.ΜΗΡΟΣ="
                a = a + Chr(13) + "ΠΕΡΙΦΕΡΕΙΑ="
                a = a + Chr(13) + "ΔΕΞ.ΓΟΝΑΤΟ= " + Chr(13) + "ΔΕΞ.ΜΗΡΟΣ="
                a = a + Chr(13) + "ΚΟΙΛΙΑ=" + Chr(13) + "KGR="

                If f_KENTRA_ADYNATISMATOS = 1 Then
                    msEditor1.Text = msEditor1.Text + a
                End If
            End If
        
            If KeyCode = vbKeyF7 Then
                a = "Συμφ.Ποσό " + Format(Val(SYMFPOSO.Text), "####.00") + Chr(13)
                a = a + "Συμφ.Mηχανήματα " + Format(Val(mhxanhmata.Text), "####") + Chr(13)
                a = a + "Ημερ.Συμφ." + Format(HMESYMF, "DD/MM/YYYY") + Chr(13)
                msEditor1.Text = a + msEditor1.Text
            End If
        
            If KeyCode = vbKeyF9 Then
                SendKeys Format(Now, "DD/MM/YY")
            End If
        
        End If
        
        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
100     If gApoMenu Then  ' ΑΝ ΤΟ ΚΑΛΩ ΑΠΟ ΤΗΝ ΚΑΡΤΕΛΛΑ ΝΑ ΜΗΝ ΑΛΛΑΖΕΙΔΙΑΣΤΑΣΕΙΣ
110         mForm_Load Me, fh, fw, ft, fl
        End If
 
        ' El.mForm_Load Me

        F_CAN_SYGXONEYSI = Val(FINDPARAMETROI(1, "PELAT2", "F_CAN_SYGXONEYSI", "0", "ΜΠΟΡΩ ΜΕ ΔΙΟΡΘΩΣΗ ΝΑ ΣΥΓΧΩΝΕΥΩ ΚΩΔΙΚΟΥΣ=1 ΟΧΙ=0"))

        f_KENTRA_ADYNATISMATOS = Val(FINDPARAMETROI(1, "PELAT2", "f_KENTRA_ADYNATISMATOS", "0", "Αν είναι κέντρο αδυνατίσματος=1"))

        If f_KENTRA_ADYNATISMATOS = 1 Then
            kentroAdyn.Visible = True
            kentroAdyn.ListIndex = 0 ' Combo3.List(0)
            SYMFPOSO.Visible = True
            mhxanhmata.Visible = True
            HMESYMF.Visible = True
            symf(0).Visible = True
            symf(1).Visible = True
            symf(2).Visible = True
            Lnum1.Visible = False
            Lnum2.Visible = False
            Lnum3.Visible = False
            lhm1.Visible = False
    
        End If

        '3. In Form Resize, use:
        'if not m_IsLoading then e1.FormResize.Me

120     Me.KeyPreview = True

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        'f_checkPos = Val(FINDPARAMETROI("PAR1", "F_CHECKPOS", "DELETE", "1=ΕΛΕΓΧOΣ 2=ΑΠΑΓΟΡΕΥΣΗ 0=ΤΙΠΟΤΑ"))

130     PELPROM_STRING = "er"
140     PELPROM.ListIndex = 0

        Dim AA As Long
    
150     AA = GetCurrentTime()
    
160     f_TIMOKAT = Val(FINDPARAMETROI(1, "PELAT2", "F_TIMOKAT", "0", "0=εκπτ.πελάτη Τιμοκατ. -1=τιμής 1=εκπτωσης 333=με 3 χονδρικές 9999=με πελατη & είδος"))

        ' f_tabAP = Val(FINDPARAMETROI(1,"PELAT2", "F_TABAP", "3", "Προεπιλεγμένο TAB "))
        ' SSTab1.Tab = IIf(f_tabAP > SSTab1.Tabs - 1 Or f_tabAP < 0, SSTab1.Tabs - 1, f_tabAP)

        Dim DUM

170     DUM = FINDPARAMETROI(1, "PELAT2", "F_KODPROM", "DELETE", "")
180     DUM = FINDPARAMETROI(1, "PELAT2", "F_RAFI", "DELETE", "")
190     DUM = FINDPARAMETROI(1, "PELAT2", "F_OIKO", "DELETE", "")
200     DUM = FINDPARAMETROI(1, "PELAT2", "F_YPOOIK", "DELETE", "")
210     DUM = FINDPARAMETROI(1, "PELAT2", "F_XTI", "DELETE", "")

220     DUM = FINDPARAMETROI(1, "PELAT2", "F_BOHU", "DELETE", "")
230     DUM = FINDPARAMETROI(1, "PELAT2", "F_LTI", "DELETE", "")
240     DUM = FINDPARAMETROI(1, "PELAT2", "F_PROM", "DELETE", "")
250     DUM = FINDPARAMETROI(1, "PELAT2", "F_STOK", "DELETE", "")

260     F_DIGpel = Val(FINDPARAMETROI(1, "PELAT2", "F_DIGPEL", "4", "ΑΡ.ΨΗΦΙΩΝ ΚΩΔΙΚΟΥ ΠΕΛΑΤΩΝ/ΠΡΟΜΗΘΕΥΤΩΝ"))

        '\\Server\data

270     f_FirstTab = FINDPARAMETROI(1, "PELAT2", "FirstTab", "1", "Πρoεπιλεγμένο Τab ")

280     SSTab1.Tab = IIf(f_FirstTab > SSTab1.Tabs - 1 Or f_FirstTab < 0, SSTab1.Tabs - 1, f_FirstTab)

290     f_files = Trim(FINDPARAMETROI(1, "PELAT2", "F_FILES", "C:", "Δίσκος με αρχεία office"))

300     F_LHM1 = FINDPARAMETROI(1, "PELAT2", "F_lhm1", "Ημερ1", "Ετικέτα ημερ.1")
310     F_LHM2 = FINDPARAMETROI(1, "PELAT2", "F_lhm2", "Ημερ2", "Ετικέτα ημερ.2")
320     f_lHM3 = FINDPARAMETROI(1, "PELAT2", "F_lhm3", "Ημερ3", "Ετικέτα ημερ.3")
330     f_lHM4 = FINDPARAMETROI(1, "PELAT2", "F_lhm4", "Ημερ4", "Ετικέτα ημερ.4")
340     f_lHM5 = FINDPARAMETROI(1, "PELAT2", "F_lhm5", "Ημερ5", "Ετικέτα ημερ.5")
350     f_lHM6 = FINDPARAMETROI(1, "PELAT2", "F_lhm6", "Ημερ6", "Ετικέτα ημερ.6")


f_lHM7 = FINDPARAMETROI(1, "PELAT2", "F_lhm7", "Ημερ6", "Ετικέτα ημερ.6")
f_lHM8 = FINDPARAMETROI(1, "PELAT2", "F_lhm8", "Ημερ6", "Ετικέτα ημερ.6")
f_lHM9 = FINDPARAMETROI(1, "PELAT2", "F_lhm9", "Ημερ6", "Ετικέτα ημερ.6")
f_lHM10 = FINDPARAMETROI(1, "PELAT2", "F_lhm10", "Ημερ6", "Ετικέτα ημερ.6")
f_lHM11 = FINDPARAMETROI(1, "PELAT2", "F_lhm11", "Ημερ6", "Ετικέτα ημερ.6")






        Dim f_EKPTCAPTION As String 'Εκπτωση % Tιμοκατάλογος

360     f_EKPTCAPTION = FINDPARAMETROI(1, "PELAT2", "f_EKPTCAPTION", "Εκπτωση %", "Eτικετα εκπτωσης  ")

370     Label1(11).Caption = f_EKPTCAPTION

380     lhm1.Caption = F_LHM1
390     lhm2.Caption = F_LHM2
400     lhm3.Caption = f_lHM3
410     lhm4.Caption = f_lHM4
420     lhm5.Caption = f_lHM5
430     lhm6.Caption = f_lHM6

lhm7.Caption = f_lHM7
lhm8.Caption = f_lHM8
lhm9.Caption = f_lHM9
lhm10.Caption = f_lHM10
lhm11.Caption = f_lHM11






        ' Gdb.Execute "ALTER TABLE PEL ADD ENERGOS B"

440     f_lab1 = FINDPARAMETROI(1, "PELAT2", "F_lab1", "Κειμ1", "Ετικέτα κειμ.1")
450     f_lab2 = FINDPARAMETROI(1, "PELAT2", "F_lab2", "Κειμ2", "Ετικέτα κειμ.2")
460     f_lab3 = FINDPARAMETROI(1, "PELAT2", "F_lab3", "Κειμ3", "Ετικέτα κειμ.3")
470     f_lab4 = FINDPARAMETROI(1, "PELAT2", "F_lab4", "Κειμ4", "Ετικέτα κειμ.4")
480     f_lab5 = FINDPARAMETROI(1, "PELAT2", "F_lab5", "Κειμ5", "Ετικέτα κειμ.5")
490     f_lab6 = FINDPARAMETROI(1, "PELAT2", "F_lab6", "Κειμ6", "Ετικέτα κειμ.6")
500     lab1.Caption = f_lab1
510     lab2.Caption = f_lab2
520     lab3.Caption = f_lab3
530     lab4.Caption = f_lab4
540     lab5.Caption = f_lab5
550     lab6.Caption = f_lab6

560     f_num1 = FINDPARAMETROI(1, "PELAT2", "F_num1", "Αριθ1", "Ετικέτα κειμ.1")
570     f_num2 = FINDPARAMETROI(1, "PELAT2", "F_num2", "Αριθ2", "Ετικέτα κειμ.2")
580     f_num3 = FINDPARAMETROI(1, "PELAT2", "F_num3", "Αριθ3", "Ετικέτα κειμ.3")
590     f_num4 = FINDPARAMETROI(1, "PELAT2", "F_num4", "Αριθ4", "Ετικέτα κειμ.4")
600     Lnum1.Caption = f_num1
610     Lnum2.Caption = f_num2
620     Lnum3.Caption = f_num3
630     Lnum4.Caption = f_num4

        'Dim f_lbcb1 As String, f_comb1 As String
        'Dim f_lbcb2 As String, f_comb2 As String
        'Dim f_lbcb3 As String, f_comb3 As String
        'Dim f_lbcb4 As String, f_comb4 As String
        'Dim f_lbcb5 As String, f_comb5 As String

        f_lbcb1 = FINDPARAMETROI(1, "PELAT2", "f_lbcb1", "Αριθ1", "Πολ.Eτικ 1")
        f_lbcb2 = FINDPARAMETROI(1, "PELAT2", "f_lbcb2", "Αριθ2", "Πολ.Eτικ 2")
        f_lbcb3 = FINDPARAMETROI(1, "PELAT2", "f_lbcb3", "Αριθ3", "Πολ.Eτικ 3")
        f_lbcb4 = FINDPARAMETROI(1, "PELAT2", "f_lbcb4", "Αριθ4", "Πολ.Eτικ 4")
        f_lbcb5 = FINDPARAMETROI(1, "PELAT2", "f_lbcb5", "Αριθ4", "Πολ.Eτικ 5")
        lbcb1.Caption = f_lbcb1
        lbcb2.Caption = f_lbcb2
        lbcb3.Caption = f_lbcb3
        lbcb4.Caption = f_lbcb4
        lbcb5.Caption = f_lbcb5

        f_comb1 = FINDPARAMETROI(1, "PELAT2", "f_comb1", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        f_comb2 = FINDPARAMETROI(1, "PELAT2", "f_comb2", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        f_comb3 = FINDPARAMETROI(1, "PELAT2", "f_comb3", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        f_comb4 = FINDPARAMETROI(1, "PELAT2", "f_comb4", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        f_comb5 = FINDPARAMETROI(1, "PELAT2", "f_comb5", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")

        Dim gg2, k
     
        If Len(Trim(f_comb1)) > 0 Then
            gg2 = Split(f_comb1, ";")

            For k = LBound(gg2) To UBound(gg2)
                Comb1.AddItem gg2(k)
            Next

        End If

        If Len(Trim(f_comb2)) > 0 Then
            gg2 = Split(f_comb2, ";")

            For k = LBound(gg2) To UBound(gg2)
                Comb2.AddItem gg2(k)
            Next

        End If

        If Len(Trim(f_comb3)) > 0 Then
            gg2 = Split(f_comb3, ";")

            For k = LBound(gg2) To UBound(gg2)
                Comb3.AddItem gg2(k)
            Next

        End If

        If Len(Trim(f_comb4)) > 0 Then
            gg2 = Split(f_comb4, ";")

            For k = LBound(gg2) To UBound(gg2)
                Comb4.AddItem gg2(k)
            Next

        End If

        If Len(Trim(f_comb5)) > 0 Then
            gg2 = Split(f_comb5, ";")

            For k = LBound(gg2) To UBound(gg2)
                Comb5.AddItem gg2(k)
            Next

        End If

        Dim R As New ADODB.Recordset

        On Error GoTo UPDATE_STR

640     R.Open "SELECT TOP 1 HM1 FROM PEL", Gdb, adOpenForwardOnly, adLockReadOnly

        On Error Resume Next

650     R.Close

        'SXOLIA CHAR(50)

        'DTPicker1.Value = Null
        'DTPicker2.Value = Null
        'dtpicker3.Value = Null
        'dtpicker4.Value = Null
        'dtpicker5.Value = Null
        'dtpicker6.Value = Null

660     Me.Picture = LoadPicture(gPicture)
670     PROTH = 1
680     Text2(1) = ""    'goo
690     Command3.Caption = "Διόρθωση"   'ΣΔΦΔΣΦ

700     Me.Picture = LoadPicture(gPicture)

        'tropos plhromhs
        'Dim R As New ADODB.Recordset

710     ADODC2.ConnectionString = gConnect

720     ADODC2.RecordSource = "SELECT STR(AYJON,2,0) AS AYJ,PERIGRAFH FROM PINAKES WHERE TYPOS=13 ORDER BY AYJON"
730     ADODC2.Refresh

740     F_ArPerEtik = Val(FINDPARAMETROI(1, "PELAT2", "F_ArPerEtik", "2", "Αριστερό περιθώριο ετικέττας")) 'posa psifia tha exei h kathe seira
    
750     R.Open "SELECT *FROM PINAKES WHERE TYPOS=12 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

760     Do While Not R.EOF

770         If R("typos") = 12 Then
780             Combo2.AddItem str(R("AYJON")) + "." + R("PERIGRAFH")
            End If

790         R.MoveNext
        Loop

        'KATHGORIA
800     R.Close
810     R.Open "SELECT *FROM PINAKES WHERE TYPOS=13 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

820     If R.EOF Then
830         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH) VALUES (13,1,'ΕΜΠΟΡΟΣ')"
        End If

        'Do While Not R.EOF
        '   If R("typos") = 13 Then
        '      KATHGORIA.A Str(R("AYJON")) + "." '+ R(0) '"PERIGRAFH")
        '   End If
        '   R.MoveNext
        'Loop

840     R.Close

        ' R.Close

        'ΦΟΡΤΩΜΑ ΠΩΛΗΤΩΝ
850     R.Open "SELECT *FROM PINAKES WHERE TYPOS=5 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

860     Do While Not R.EOF
870         Combo3.AddItem (str(R("AYJON"))) + "." + R("PERIGRAFH")
880         R.MoveNext
        Loop

890     R.Close

900     If f_TIMOKAT = 1 Then
910         Label1(11).Caption = "Τιμοκατάλογος"
        End If

        Dim mf2

920     If Len(f_files) = 0 Then
        Else
930         mf2 = UCase(f_files + "\MERCVB\files\")
940         Dir1.Path = mf2
950         Dir1.Refresh
        End If

        '    mForm_Load Me, fh, fw, ft, fl

960     Me.Caption = str(GetCurrentTime() - AA)

        Exit Sub

UPDATE_STR:

        On Error Resume Next

970     Gdb.Execute "ALTER TABLE PEL ADD HM1 DATETIME"
980     Gdb.Execute "ALTER TABLE PEL ADD HM2 DATETIME"
990     Gdb.Execute "ALTER TABLE PEL ADD HM3 DATETIME"
1000    Gdb.Execute "ALTER TABLE PEL ADD HM4 DATETIME"
1010    Gdb.Execute "ALTER TABLE PEL ADD HM5 DATETIME"
1020    Gdb.Execute "ALTER TABLE PEL ADD HM6 DATETIME"

1030    Gdb.Execute "ALTER TABLE PEL ADD CH1 CHAR(30)"
1040    Gdb.Execute "ALTER TABLE PEL ADD CH2 CHAR(30)"
1050    Gdb.Execute "ALTER TABLE PEL ADD CH3 CHAR(30)"
1060    Gdb.Execute "ALTER TABLE PEL ADD CH4 CHAR(30)"
1070    Gdb.Execute "ALTER TABLE PEL ADD CH5 CHAR(30)"
1080    Gdb.Execute "ALTER TABLE PEL ADD CH6 CHAR(30)"
1090    Gdb.Execute "ALTER TABLE PEL ADD CH7 CHAR(30)"
    
1100    Gdb.Execute "ALTER TABLE PEL ADD ENERGOS BIT"

1110    Resume Next

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Paint()

        '<EhHeader>
        On Error GoTo Form_Paint_Err

        '</EhHeader>
100     If PROTH = 1 Then
110         PROTH = 2
            'data1.DatabaseName = gDir

120         Data1.RecordSource = "SELECT TOP 10 *FROM PEL WHERE EIDOS='X'"
130         Data1.ConnectionString = gConnect
140         Data1.Refresh
            'Text2(0).SetFocus
150         PELPROM.SetFocus    '  Option1(1).SetFocus

160         If Len(Trim(ffbuff.Caption)) > 0 Then
170             If Left(ffbuff.Caption, 1) = "e" Then
180                 PELPROM.ListIndex = 0
                Else
190                 PELPROM.ListIndex = 1
                End If

200             Text2(0).Text = Trim(mID(ffbuff, 2, 10))
210             Text2_LostFocus 0

            End If

        End If

        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Form_Paint " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Form_Paint " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
    If gApoMenu Then  ' ΑΝ ΤΟ ΚΑΛΩ ΑΠΟ ΤΗΝ ΚΑΡΤΕΛΛΑ ΝΑ ΜΗΝ ΑΛΛΑΖΕΙΔΙΑΣΤΑΣΕΙΣ
        '   mForm_Resize Me,8, 15000, 10000, fh, fw, ft, fl
    End If

    ' El.mForm_Resize Me, 15000, 10000  ', fh, fw, ft, fl
  '  ResizeFormFor Me
    '3. In Form Resize, use:
    ' If Not m_IsLoading Then El.FormResize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'goo = Text2(1).Text
End Sub

Private Sub hmesymf_LostFocus()

    Dim a As String
  
    If IsDate(HMESYMF) Then
   
    Else
        HMESYMF = ""
   
    End If

End Sub

Private Sub kentroAdyn_Click()

    On Error Resume Next

    If f_KENTRA_ADYNATISMATOS = 1 Then
        If Data1.Recordset("NUM3") = 0 Then ' HTAN STO SVMA
            If kentroAdyn.Text = "ΣΩΜΑ" Then  ' PARAMENEI STO SOMA
                'MHN KANEIS TIPOTA
                Text2(0).Text = Data1.Recordset("kod")
            Else  ' TO GYRIZEI STO PROSOPO
                Text2(0).Text = "10" + Data1.Recordset("KOD")
            End If
        End If
   
        If Data1.Recordset("NUM3") = 1 Then ' HTAN STO PROSOPO
            If kentroAdyn.Text <> "ΣΩΜΑ" Then  ' PARAMENEI STO PROSOPO
                'MHN KANEIS TIPOTA
                Text2(0).Text = Data1.Recordset("kod")
            Else  ' TO GYRIZEI STO SOMA
                Text2(0).Text = mID(Data1.Recordset("KOD"), 3, 4)
            End If
        End If

        Text2(1).Text = ""
        Text2(0).SetFocus
    
        Text2(1).SetFocus
   
    End If

End Sub

'Private Sub Label2_Click()
'
'
'End Sub

Private Sub Label1_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo Label1_Click_Err

        '</EhHeader>

100     If Index = 18 Then
110         UPDATE_PINAKES 12, Label1(Index), Combo2
        End If

        '<EhFooter>
        Exit Sub

Label1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Label1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Label1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'       KeyAscii = 0 'suppress the beep
'       Text2(0).SetFocus
'       'keybd_event VK_TAB, 0, 0, 0 'send a tab
'    End If
'End Sub

'Private Sub Picture1_Click()
'    'αλλαγή γραμματοσειράς
'        '<EhHeader>
'        On Error GoTo Picture1_Click_Err
'        '</EhHeader>
'
'100     CommonDialog1.Flags = cdlCFBoth   ' Flags property must be set
'110     CommonDialog1.ShowFont
'120     With Text3
'            '        .SELSTART = FSELSTART
'            '        .SELLENGTH = FSELLENGTH
'
'130         .SelFontName = CommonDialog1.FontName
'140         .SelFontSize = CommonDialog1.FontSize
'150         .SelBold = CommonDialog1.FontBold
'160         .SelItalic = CommonDialog1.FontItalic
'170         .SelStrikeThru = CommonDialog1.FontStrikethru
'180         .SelUnderline = CommonDialog1.FontUnderline
'
'        End With
'
'        '<EhFooter>
'        Exit Sub
'
'Picture1_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'               "in ADOMERCNEW.pelat2.Picture1_Click " & _
'               "at line " & Erl, _
'               vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Picture1_Click " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub

'Private Sub Picture2_Click()
'    'αλλαγή γραμματοσειράς
'        '<EhHeader>
'        On Error GoTo Picture2_Click_Err
'        '</EhHeader>
'100     CommonDialog1.ShowColor
'
'        '   Text3.SELSTART = FSELSTART
'        '        Text3.SELLENGTH = FSELLENGTH
'
'110     Text3.SelColor = CommonDialog1.Color
'
'
'
'        '<EhFooter>
'        Exit Sub
'
'Picture2_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'               "in ADOMERCNEW.pelat2.Picture2_Click " & _
'               "at line " & Erl, _
'               vbExclamation + vbOKOnly, "Application Error"

'        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Picture2_Click " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
Private Sub Text1_GotFocus(Index As Integer)

        '<EhHeader>
        On Error GoTo Text1_GotFocus_Err

        '</EhHeader>

100     Text1(Index).BackColor = vbYellow

        '<EhFooter>
        Exit Sub

Text1_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Text1_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Text1_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

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
         "in ADOMERCNEW.pelat2.Text1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Text1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_LostFocus(Index As Integer)

        '<EhHeader>
        On Error GoTo Text1_LostFocus_Err

        '</EhHeader>
        Dim B, a

100     Text1(Index).BackColor = vbWhite

110     If Index = 5 Then
120         B = Text1(5).Text
130         a = check_afm(B)
        End If

        '<EhFooter>
        Exit Sub

Text1_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Text1_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Text1_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_GotFocus(Index As Integer)

        '<EhHeader>
        On Error GoTo Text2_GotFocus_Err

        '</EhHeader>

100     Text2(Index).BackColor = vbYellow

        '<EhFooter>
        Exit Sub

Text2_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Text2_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Text2_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Text2_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Text2_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Text2_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Text2_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Text2_LostFocus(Index As Integer)

        '<EhHeader>
        On Error GoTo Text2_LostFocus_Err

        '</EhHeader>

        Dim x As String

100     Text2(Index).BackColor = vbWhite

110     If Len(Trim(Text2(Index))) = 0 Then Exit Sub
        'If Option1(1) Then X = "e" Else X = "r"

120     x = "e"

        On Error Resume Next

130     x = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)
140     Data1.ConnectionString = gConnect
     
        'If gUserId = 2 Then  ' TAMIAS KAGIOYDIS
        '
        '   If Index = 1 Then  'eponumo
        '      DATA1.RecordSource = "select EPO,THL,*from PEL where KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%' AND EIDOS='" + x + "' and EPO like '" + (SameLetters(Text2(1).Text)) + "%';"
        '   End If
        '
        '   If Index = 0 Then 'kodikos
        '      DATA1.RecordSource = "select EPO,THL,*from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND  EIDOS='" + x + "' and KOD like '" + Text2(0).Text + "%';"
        '   End If
        '
        '   If Index = 2 Then 'afm
        '       DATA1.RecordSource = "select EPO,THL.* from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND  EIDOS='" + x + "' and AFM like '" + Text2(2).Text + "%';"
        '   End If

        Dim MSYNT

        MSYNT = ""

        If f_KENTRA_ADYNATISMATOS = 1 Then
            If x = "e" Then
                If kentroAdyn.List(0) = kentroAdyn.Text Then
                    MSYNT = "NUM3=0 AND "
                Else
                    MSYNT = "NUM3=1 AND "
                End If
            End If

        End If

150     '   If gUserId = 2 Then

160     '      If Index = 1 Then  'eponumo
170     '         DATA1.RecordSource = "select *from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND EIDOS='" + X + "' and EPO like '" + (SameLetters(Text2(1).Text)) + "%';"
        '    End If
180     '   If Index = 0 Then    'kodikos
190     '      DATA1.RecordSource = "select *from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND EIDOS='" + X + "' and KOD like '" + (SameLetters(Text2(0).Text)) + "%';"
        ' End If
200     '       If Index = 2 Then    'afm
210     '          DATA1.RecordSource = "select *from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND EIDOS='" + X + "' and AFM like '" + Text2(2).Text + "%';"
        '     End If

        'Else
        Dim FF

        'If Option1(1) Then X = "e" Else X = "r"
220     If Index = 1 Then  'eponumo

230         If Asc(Left(Text2(1).Text, 1)) > 127 Then
240             FF = Text2(1).Text
            Else
250             FF = (SameLetters(Text2(1).Text))
            End If

260         Data1.RecordSource = "select *from PEL where " + MSYNT + " EIDOS='" + x + "' and EPO like '" + FF + "%';"
        End If

270     If Index = 0 Then    'kodikos
280         Data1.RecordSource = "select *from PEL where " + MSYNT + " EIDOS='" + x + "' and KOD like '" + (SameLetters(Text2(0).Text)) + "%';"
        End If

290     If Index = 2 Then    'afm
300         Data1.RecordSource = "select *from PEL where " + MSYNT + " EIDOS='" + x + "' and AFM like '" + Text2(2).Text + "%';"
        End If

        ' End If

        'DataGrid1.Enabled = False

310     Data1.Refresh
320     DataGrid1.columns(0).Width = TextWidth("A") * 30
          
        MDataGrid1_RowColChange
        'DataGrid1.Enabled = True

        Dim MF As String

        '  If Option1(1).value = True Then
330     MF = UCase(f_files + "\MERCVB\files\" + mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1) + Text1(0))

        '  Else
        '     mF = UCase(f_files + "\MERCVB\files\r" + Text1(0))
        '  End If
        On Error Resume Next

340     If Text1(0).Text = Data1.Recordset("KOD") Then  'GIA NA MHN KATHISTEREI
350         If Len(Dir(MF, vbDirectory)) > 0 Then
360             File1.Path = MF
            Else

370             If Len(Dir("c:\mercvb\files\dum", vbDirectory)) = 0 Then
380                 If Len(Dir("c:\mercvb\files", vbDirectory)) = 0 Then
390                     MkDir "c:\mercvb\files"
                    End If

400                 MkDir "c:\mercvb\files\dum"
                End If

410             File1.Path = "c:\mercvb\files\dum"
            End If

420         File1.Refresh

        End If

        Dim PI

430     PI = "C:\MERCVB\IMAGES\e" + Trim(Text1(0).Text) + "*" + ".JPG"

440     If Len(Dir(PI)) > 0 Then
450         foto2.Picture = LoadPicture("c:\mercvb\images\" + Dir(PI))
460         StretchSourcePictureFromPicture Me.foto2.Picture, Me.foto2
        Else
470         foto2.Picture = LoadPicture()
        End If

        On Error Resume Next

480     DataGrid1.SetFocus

        '<EhFooter>
        Exit Sub

Text2_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Text2_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Text2_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub TEXT3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'If Button = vbRightButton Then
'
'  If Check2 Then
'       CommonDialog1.ShowColor
'       TEXT3.SelColor = CommonDialog1.Color
'
'
'  Else
'
'    CommonDialog1.Flags = cdlCFBoth   ' Flags property must be set
'    CommonDialog1.ShowFont
'    With TEXT3
'       .SelFontName = CommonDialog1.FontName
'       .SelFontSize = CommonDialog1.FontSize
'       .SelBold = CommonDialog1.FontBold
'       .SelItalic = CommonDialog1.FontItalic
'       .SelStrikeThru = CommonDialog1.FontStrikethru
'       .SelUnderline = CommonDialog1.FontUnderline
'    End With
'  End If
'End If

'End Sub

Function get_next_pel(meidos As String) As String

        '<EhHeader>
        On Error GoTo get_next_pel_Err

        '</EhHeader>
        Dim DB      As Database

        Dim R       As New ADODB.Recordset

        Dim r0      As New ADODB.Recordset

        Dim k       As Long

        Dim FORMAT1 As String

100     FORMAT1 = Left("00000000000", F_DIGpel)

        On Error Resume Next

        'Set db = OpenDatabase(gDir)
        'Set db = OpenDatabase(gDir, False, False)
        'Set db = OpenDatabase(gDir, False, False, gConnect)

110     If True Then    'eyresi kenoy aritmoy
120         R.Open "select KOD FROM PEL WHERE LEN(KOD)=" + str(F_DIGpel) + " AND EIDOS='" + meidos + "' order by KOD", Gdb, adOpenForwardOnly, adLockReadOnly

130         If R.EOF Then
140             get_next_pel = FORMAT1
150             R.Close: Exit Function
            End If

160         Me.MousePointer = vbHourglass

170         R.MoveFirst
180         k = 0

190         Do While Not R.EOF

200             If Val(R(0)) = 0 Then
210                 R.MoveNext
                Else
220                 k = k + 1

230                 If Left(R(0), 4) = Format(k, FORMAT1) Then
240                     R.MoveNext
                    Else

                        On Error GoTo get_next_pel_Err

250                     r0.Open "select COUNT(*) FROM PEL WHERE KOD='" + Format(k, FORMAT1) + "' AND EIDOS='" + meidos + "' ", Gdb, adOpenForwardOnly, adLockReadOnly

260                     If r0(0) > 0 Then    'ΗΔΗ ΥΠΑΡΧΕΙ
270                         r0.Close

                            Exit Do

                        Else
280                         get_next_pel = Format(k, FORMAT1)
290                         R.Close
300                         Me.MousePointer = vbNormal

                            Exit Function

                        End If

                    End If
                End If

            Loop

310         R.Close

            Dim ll, L

            '

320         For L = k + 1 To 10 ^ F_DIGpel
330             R.Open "select COUNT(*) FROM PEL WHERE KOD='" + Format(L, FORMAT1) + "' AND EIDOS='" + meidos + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
340             ll = R(0)
350             R.Close

360             DoEvents
370             Me.Caption = k

380             If ll = 0 Then

                    Exit For

                End If

            Next

390         get_next_pel = Format(L, FORMAT1)
400         Me.MousePointer = vbNormal

            Exit Function

        End If

410     R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic

420     If meidos = "e" Then
430         get_next_pel = R("epel") + 1
        Else
440         get_next_pel = R("rpel") + 1
        End If

        '<EhFooter>
        Exit Function

get_next_pel_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.get_next_pel " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.get_next_pel " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Private Sub Text3_LostFocus()
'        '<EhHeader>
'        On Error GoTo Text3_LostFocus_Err
'        '</EhHeader>
'
'100     Text3.SelStart = FSELSTART
'110     Text3.SelLength = FSELLENGTH
'
'
'
'
'
'
'
'        '<EhFooter>
'        Exit Sub
'
'Text3_LostFocus_Err:
'        'MsgBox Err.Description & vbCrLf & _
'               "in ADOMERCNEW.pelat2.Text3_LostFocus " & _
'               "at line " & Erl, _
'               vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Text3_LostFocus " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'        '<EhHeader>
'        On Error GoTo Text3_MouseUp_Err
'        '</EhHeader>
'100     FSELSTART = Text3.SelStart
'110     FSELLENGTH = Text3.SelLength
'
'
'        '<EhFooter>
'        Exit Sub
'
'Text3_MouseUp_Err:
'        'MsgBox Err.Description & vbCrLf & _
'               "in ADOMERCNEW.pelat2.Text3_MouseUp " & _
'               "at line " & Erl, _
'               vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Text3_MouseUp " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
