VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form par1 
   BackColor       =   &H00FF0000&
   Caption         =   "Νέο Παραστατικό"
   ClientHeight    =   12510
   ClientLeft      =   -5595
   ClientTop       =   540
   ClientWidth     =   19260
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12510
   ScaleWidth      =   19260
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      Caption         =   "Παραλαβη Τιμ/ΔΑ"
      Height          =   360
      Left            =   12000
      TabIndex        =   191
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton sda_epistrofis 
      Caption         =   "ΣΔΑ Επιστροφής"
      Enabled         =   0   'False
      Height          =   360
      Left            =   10440
      TabIndex        =   186
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox fortARit 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   8520
      TabIndex        =   185
      Top             =   5880
      Width           =   270
   End
   Begin VB.TextBox FORTTK 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   8085
      TabIndex        =   181
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox FORTPOL 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   6240
      TabIndex        =   180
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox FORTDIE 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   6285
      TabIndex        =   179
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton tamiaki 
      Caption         =   "Ταμειακη"
      Height          =   375
      Left            =   13560
      TabIndex        =   174
      Top             =   4320
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   492
      Left            =   18480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   480
      Width           =   492
   End
   Begin VB.CheckBox chkAgor 
      Caption         =   "Αγορές"
      Height          =   255
      Left            =   4440
      TabIndex        =   153
      Top             =   7800
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkPOL 
      Caption         =   "Πωλήσεις"
      Height          =   255
      Left            =   2760
      TabIndex        =   152
      Top             =   7800
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kαρτέλα"
      Height          =   210
      Left            =   9120
      TabIndex        =   150
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton CmdRefreshPel 
      Caption         =   "Ανανέωση"
      Height          =   210
      Left            =   7680
      TabIndex        =   149
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3405
      Left            =   10440
      TabIndex        =   139
      Top             =   720
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   6006
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483635
      TabCaption(0)   =   "Προορισμός"
      TabPicture(0)   =   "par1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblΟδός"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblL"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblΠόλη"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label29"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblΑλλοςΣκοπός"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label33"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "STREET"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "STREETNUMBER"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "POSTALCODE"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CITY"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "BRANCH"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "other"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "ΗΜΕΡ/ΩΡΑ ΠΡΔ"
      TabPicture(1)   =   "par1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "HMERPARAD"
      Tab(1).Control(1)=   "ORAPARAD"
      Tab(1).Control(2)=   "LABEL30"
      Tab(1).Control(3)=   "lblΗΜΕΡΩΡΑ"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Mydata"
      TabPicture(2)   =   "par1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label27"
      Tab(2).Control(1)=   "Label28"
      Tab(2).Control(2)=   "LABEL31"
      Tab(2).Control(3)=   "Combo2TyposParakr"
      Tab(2).Control(4)=   "Combo2OnomaParak"
      Tab(2).Control(5)=   "AJIA2PARAKR"
      Tab(2).Control(6)=   "cmdypolo"
      Tab(2).ControlCount=   7
      Begin VB.TextBox other 
         Height          =   375
         Left            =   1080
         TabIndex        =   169
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox BRANCH 
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   167
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox CITY 
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   162
         Top             =   2040
         Width           =   2000
      End
      Begin VB.TextBox POSTALCODE 
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   161
         Top             =   1560
         Width           =   1000
      End
      Begin VB.TextBox STREETNUMBER 
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   160
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox STREET 
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   1080
         MaxLength       =   80
         TabIndex        =   159
         Top             =   600
         Width           =   2700
      End
      Begin VB.CommandButton cmdypolo 
         Caption         =   "υπολογισμός"
         Height          =   240
         Left            =   -74760
         TabIndex        =   147
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox AJIA2PARAKR 
         BackColor       =   &H0080C0FF&
         Height          =   375
         Left            =   -73920
         TabIndex        =   143
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox Combo2OnomaParak 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   1080
         Width           =   2775
      End
      Begin VB.ComboBox Combo2TyposParakr 
         Height          =   315
         ItemData        =   "par1.frx":0054
         Left            =   -74760
         List            =   "par1.frx":0067
         Style           =   2  'Dropdown List
         TabIndex        =   141
         Top             =   600
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker HMERPARAD 
         Height          =   315
         Left            =   -72960
         TabIndex        =   171
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   16711680
         CalendarTrailingForeColor=   16711680
         Format          =   308477953
         CurrentDate     =   38294
      End
      Begin MSComCtl2.DTPicker ORAPARAD 
         Height          =   315
         Left            =   -72960
         TabIndex        =   172
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   16711680
         CalendarTrailingForeColor=   16711680
         Format          =   308477954
         CurrentDate     =   38294
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".."
         Height          =   195
         Left            =   3720
         TabIndex        =   178
         Top             =   3240
         Width           =   90
      End
      Begin VB.Label LABEL31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".."
         Height          =   195
         Left            =   -74400
         TabIndex        =   176
         Top             =   3000
         Width           =   90
      End
      Begin VB.Label LABEL30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".."
         Height          =   195
         Left            =   -74400
         TabIndex        =   175
         Top             =   3120
         Width           =   90
      End
      Begin VB.Label lblΗΜΕΡΩΡΑ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ΗΜΕΡ/ΩΡΑ ΠΑΡΑΔ"
         Height          =   195
         Left            =   -74640
         TabIndex        =   173
         Top             =   840
         Width           =   1410
      End
      Begin VB.Label lblΑλλοςΣκοπός 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αλλος Σκοπός"
         Height          =   195
         Left            =   0
         TabIndex        =   170
         Top             =   2880
         Width           =   1020
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδ.Υποκ"
         Height          =   195
         Left            =   240
         TabIndex        =   168
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label lblΠόλη 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Πόλη"
         Height          =   195
         Left            =   240
         TabIndex        =   166
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label lblTK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TK"
         Height          =   195
         Left            =   240
         TabIndex        =   165
         Top             =   1560
         Width           =   210
      End
      Begin VB.Label lblL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αριθμός"
         Height          =   195
         Left            =   240
         TabIndex        =   164
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblΟδός 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Οδός"
         Height          =   195
         Left            =   240
         TabIndex        =   163
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αξία"
         Height          =   195
         Left            =   -74640
         TabIndex        =   145
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Δεύτερος φόρος/τέλος/Παρακράτηση"
         Height          =   195
         Left            =   -73920
         TabIndex        =   144
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   1
      Left            =   0
      TabIndex        =   114
      Top             =   11640
      Visible         =   0   'False
      Width           =   10320
      Begin VB.TextBox Textm 
         Height          =   285
         Index           =   24
         Left            =   6120
         MultiLine       =   -1  'True
         TabIndex        =   119
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Textm 
         Height          =   285
         Index           =   25
         Left            =   7080
         TabIndex        =   120
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Textm 
         Height          =   285
         Index           =   27
         Left            =   9120
         MultiLine       =   -1  'True
         TabIndex        =   122
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Textm 
         Height          =   285
         Index           =   21
         Left            =   2160
         TabIndex        =   116
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Textm 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   20
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   115
         Top             =   600
         Width           =   9735
      End
      Begin VB.TextBox Textm 
         Height          =   285
         Index           =   22
         Left            =   3480
         TabIndex        =   117
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Textm 
         Height          =   285
         Index           =   23
         Left            =   4680
         TabIndex        =   118
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Textm 
         Height          =   285
         Index           =   26
         Left            =   8040
         TabIndex        =   121
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "AME EGTL"
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
         Left            =   9000
         TabIndex        =   129
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Diesel Θέρμ"
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
         Left            =   7920
         TabIndex        =   128
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label MTF14 
         BackStyle       =   0  'Transparent
         Caption         =   "Αμολ.100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   127
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label MTF13 
         BackStyle       =   0  'Transparent
         Caption         =   "VPAM98"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   126
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label MTF12 
         BackStyle       =   0  'Transparent
         Caption         =   "Αμόλυβδη 95"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   125
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label MTF11 
         BackStyle       =   0  'Transparent
         Caption         =   "V Power"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   124
         Top             =   0
         Width           =   855
      End
      Begin VB.Label MTF21 
         BackStyle       =   0  'Transparent
         Caption         =   "Diesel Κίν"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   6960
         TabIndex        =   123
         Top             =   0
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   1620
      Index           =   0
      Left            =   0
      TabIndex        =   93
      Top             =   10000
      Visible         =   0   'False
      Width           =   10320
      Begin VB.TextBox tc7 
         Height          =   285
         Left            =   120
         TabIndex        =   190
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TC6 
         Height          =   285
         Left            =   6360
         TabIndex        =   189
         Top             =   1200
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox TC5 
         Height          =   285
         Left            =   1200
         TabIndex        =   188
         Top             =   1200
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.TextBox TextEntoleas 
         Height          =   855
         Left            =   9240
         MultiLine       =   -1  'True
         TabIndex        =   136
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TextOnEntolea 
         Height          =   285
         Left            =   8760
         TabIndex        =   135
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Textm 
         Height          =   285
         Index           =   7
         Left            =   3120
         TabIndex        =   113
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Textm 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   112
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Textm 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   111
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Textm 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   0
         Left            =   7080
         MultiLine       =   -1  'True
         TabIndex        =   102
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Textm 
         Height          =   525
         Index           =   1
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Textm 
         Height          =   525
         Index           =   2
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   96
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Textm 
         Height          =   525
         Index           =   4
         Left            =   6240
         TabIndex        =   95
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Textm 
         Height          =   525
         Index           =   3
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   94
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Εντολέας"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8760
         TabIndex        =   137
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Labf 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   3
         Left            =   6840
         TabIndex        =   134
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Labf 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   2
         Left            =   6600
         TabIndex        =   133
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Labf 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   132
         Top             =   0
         Width           =   45
      End
      Begin VB.Label Labf 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "----"
         Height          =   195
         Index           =   0
         Left            =   6360
         TabIndex        =   131
         Top             =   840
         Width           =   180
      End
      Begin VB.Label MTF21 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Περιγραφή"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   103
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label MTF11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Αριθ.Τεμαχίων"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   101
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label MTF12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Είδος Συσκευασίας"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   100
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label MTF13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Είδος Εμπορεύματος"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   99
         Top             =   0
         Width           =   2865
      End
      Begin VB.Label MTF14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Μικ.Βάρος"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   98
         Top             =   0
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   10680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton metasxhmatismos 
      Caption         =   "Μετασχηματισμός"
      Height          =   420
      Left            =   6720
      TabIndex        =   91
      Top             =   7560
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox ar_apod 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   5640
      TabIndex        =   90
      Top             =   7560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox parat 
      Height          =   285
      Left            =   7185
      MaxLength       =   200
      TabIndex        =   87
      Top             =   7215
      Width           =   3135
   End
   Begin VB.CommandButton Command8 
      Caption         =   "+"
      Height          =   192
      Left            =   8880
      TabIndex        =   74
      Top             =   6000
      Width           =   315
   End
   Begin VB.CommandButton Command7 
      Caption         =   "-"
      Height          =   192
      Left            =   8880
      TabIndex        =   73
      Top             =   6240
      Width           =   315
   End
   Begin VB.TextBox fpaEFK 
      Height          =   405
      Left            =   13845
      TabIndex        =   69
      Top             =   8640
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox SynEfk 
      Height          =   405
      Left            =   13845
      TabIndex        =   68
      Top             =   8220
      Visible         =   0   'False
      Width           =   900
   End
   Begin MSAdodcLib.Adodc DATAPARTIDES 
      Height          =   375
      Left            =   14160
      Top             =   9840
      Width           =   1200
      _ExtentX        =   2117
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
      DataSourceName  =   "perp"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   "epsilonsa"
      RecordSource    =   ""
      Caption         =   "PARTIDES"
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
   Begin VB.Data abat2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox KentroAdyn 
      Height          =   315
      ItemData        =   "par1.frx":00C4
      Left            =   10560
      List            =   "par1.frx":00CE
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton MONTE_CRISTO 
      BackColor       =   &H00FFFF80&
      Caption         =   "ΜΕΤΑΣΧΗΜΑΤΙΣΜΟΣ"
      Height          =   375
      Left            =   10440
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   10920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TelFpa 
      Height          =   345
      Left            =   10920
      TabIndex        =   61
      Top             =   11760
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1680
      Top             =   10200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton kerdos 
      Height          =   255
      Left            =   10680
      TabIndex        =   60
      ToolTipText     =   "Κέρδος"
      Top             =   10680
      Visible         =   0   'False
      Width           =   1020
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2535
      Left            =   14400
      TabIndex        =   59
      Top             =   4560
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   12
      Cols            =   3
   End
   Begin MSAdodcLib.Adodc LASTTIMOL 
      Height          =   336
      Left            =   8520
      Top             =   10080
      Visible         =   0   'False
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
      Bindings        =   "par1.frx":00E1
      Height          =   2352
      Left            =   -120
      TabIndex        =   58
      Top             =   8040
      Width           =   10428
      _ExtentX        =   18415
      _ExtentY        =   4154
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
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2752"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=960,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=960,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=960,.italic=0"
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
   Begin VB.CommandButton Command6 
      Caption         =   "Διόρθωση"
      Height          =   210
      Left            =   9120
      TabIndex        =   57
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton COMMAND5 
      Caption         =   "Επανεκτύπωση Τελευταίου Παραστατικού"
      Enabled         =   0   'False
      Height          =   372
      Left            =   10920
      TabIndex        =   52
      Top             =   11040
      Visible         =   0   'False
      Width           =   2868
   End
   Begin VB.CommandButton Command4 
      Height          =   240
      Left            =   13920
      TabIndex        =   51
      Top             =   1080
      Width           =   165
   End
   Begin MSDataGridLib.DataGrid GridEidon 
      Bindings        =   "par1.frx":00F9
      Height          =   2715
      Left            =   15
      TabIndex        =   29
      Top             =   4845
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   4789
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   11064319
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "Courier New"
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TIMText3 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   45
      TabIndex        =   22
      Top             =   4815
      Width           =   30
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   5880
      TabIndex        =   12
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton selectPEL 
      Caption         =   "Προμηθευτής"
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   1230
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   10
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   105
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox POLAGOR 
      Height          =   315
      ItemData        =   "par1.frx":010B
      Left            =   120
      List            =   "par1.frx":0118
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox PARAS 
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
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
      ItemData        =   "par1.frx":013C
      Left            =   1320
      List            =   "par1.frx":0149
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.ComboBox SEIRA 
      BackColor       =   &H000000FF&
      Height          =   315
      Left            =   14040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox APOTH1 
      Height          =   315
      Left            =   9105
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox APOTH2 
      Height          =   315
      Left            =   9105
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   555
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ΖΥΓΟΣ"
      Height          =   255
      Left            =   11400
      TabIndex        =   37
      Top             =   11400
      Width           =   1140
   End
   Begin VB.ComboBox AYTOKINHTO 
      Height          =   315
      Left            =   7335
      TabIndex        =   36
      Top             =   6570
      Width           =   1485
   End
   Begin VB.CheckBox kostologhsh 
      Caption         =   "Κοστολόγηση"
      Height          =   255
      Left            =   8865
      TabIndex        =   35
      Top             =   6960
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Λήψη Τιμ.Excel"
      Height          =   240
      Left            =   8865
      TabIndex        =   34
      ToolTipText     =   "Κωδ,Ονομα,Ποσότητα,Τιμή"
      Top             =   4860
      Width           =   1470
   End
   Begin VB.CheckBox lianikh 
      Caption         =   "Λιανική"
      Height          =   255
      Left            =   8865
      TabIndex        =   33
      Top             =   6720
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox ektyp 
      Caption         =   "Εκτύπωση"
      Height          =   255
      Left            =   8865
      TabIndex        =   32
      Top             =   6480
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   7320
      TabIndex        =   30
      Top             =   6915
      Width           =   1485
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2625
      Top             =   9600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      Destination     =   1
      PrintFileType   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxGridDeltion 
      Bindings        =   "par1.frx":016D
      Height          =   2655
      Left            =   240
      TabIndex        =   28
      Top             =   4920
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   4683
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc DOK 
      Height          =   330
      Left            =   3825
      Top             =   9600
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "DOK"
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
   Begin MSAdodcLib.Adodc EGGTIM 
      Height          =   375
      Left            =   3195
      Top             =   9270
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "EGGTIM"
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
   Begin MSAdodcLib.Adodc PEL 
      Height          =   375
      Left            =   1155
      Top             =   9270
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "PEL"
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
   Begin MSAdodcLib.Adodc EID 
      Height          =   330
      Left            =   105
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "EID"
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
   Begin VB.Data abat 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data temp 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9465
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox MACtEXT 
      Height          =   285
      Left            =   6945
      TabIndex        =   26
      Text            =   "Text5"
      Top             =   9720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Data CMEM 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\mercvb\reports\reports.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7635
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "mem"
      Top             =   9390
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data cEGGTIM 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\mercvb\reports\reports.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8835
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9390
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data CTIM 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\mercvb\reports\reports.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9390
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox ApoParaggelia 
      Caption         =   "Από παραγγελία"
      Height          =   255
      Left            =   6825
      TabIndex        =   24
      Top             =   4905
      Width           =   1995
   End
   Begin VB.ComboBox PROORISMOS 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "par1.frx":017F
      Left            =   12645
      List            =   "par1.frx":0181
      TabIndex        =   19
      Top             =   12120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox FORTOSH 
      Height          =   315
      ItemData        =   "par1.frx":0183
      Left            =   15600
      List            =   "par1.frx":0185
      TabIndex        =   18
      Top             =   11760
      Width           =   2535
   End
   Begin VB.ComboBox SKOPOS 
      Height          =   315
      ItemData        =   "par1.frx":0187
      Left            =   6285
      List            =   "par1.frx":0189
      TabIndex        =   17
      Top             =   5520
      Width           =   2535
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "par1.frx":018B
      Left            =   6285
      List            =   "par1.frx":018D
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton b_akyr 
      Caption         =   "Εξοδος"
      Height          =   420
      Left            =   9180
      TabIndex        =   21
      Top             =   7560
      Width           =   1140
   End
   Begin VB.CommandButton b_Katax 
      BackColor       =   &H0080FF80&
      Caption         =   "Καταχώρηση"
      Height          =   420
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7560
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   4690
      MaxLength       =   6
      TabIndex        =   3
      Top             =   240
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   3492
      Left            =   0
      TabIndex        =   23
      Top             =   1320
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   6165
      _Version        =   393216
      BackColor       =   16744576
      Rows            =   300
      Cols            =   20
      BackColorSel    =   16744576
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   20
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   6915
      TabIndex        =   5
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   308740097
      CurrentDate     =   38294
   End
   Begin MSDataGridLib.DataGrid GridPelaton 
      Bindings        =   "par1.frx":018F
      Height          =   3480
      Left            =   0
      TabIndex        =   27
      Top             =   1320
      Width           =   10296
      _ExtentX        =   18150
      _ExtentY        =   6138
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid GRIDPARTIDES 
      Bindings        =   "par1.frx":01A1
      Height          =   1020
      Left            =   10440
      TabIndex        =   66
      Top             =   3120
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   1799
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   33023
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ΠΑΡTΙΔΕΣ"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "LOTNUMBER"
         Caption         =   "LOT"
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
         DataField       =   "ΛΗΞΗ"
         Caption         =   "HM.ΛΗΞ"
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
         DataField       =   "ΥΠΟΛ"
         Caption         =   "ΥΠΟΛ"
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
         DataField       =   "ΕΠΙΛ"
         Caption         =   "ΕΠΙΛΟΓΗ"
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
      BeginProperty Column04 
         DataField       =   "ID"
         Caption         =   "ID"
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
         MarqueeStyle    =   4
         Size            =   349
         BeginProperty Column00 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Teliko 
      BackColor       =   &H00C0FFC0&
      Height          =   456
      Left            =   13845
      TabIndex        =   72
      Top             =   9075
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Πρόσθετα"
      Height          =   6555
      Left            =   10320
      TabIndex        =   76
      Top             =   4080
      Width           =   3375
      Begin VB.CommandButton cmdΜΑΖ_ΕΚΤΥ 
         Caption         =   "ΜΑΖ_ΕΚΤΥ"
         Height          =   360
         Left            =   2160
         TabIndex        =   158
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdPOS 
         Caption         =   "POS"
         Height          =   360
         Left            =   120
         TabIndex        =   157
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCAI 
         Caption         =   "ΕΞΑΓΩΓΕΣ"
         Height          =   360
         Left            =   2160
         TabIndex        =   156
         Top             =   840
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.ComboBox CURR 
         Height          =   315
         ItemData        =   "par1.frx":01BC
         Left            =   2040
         List            =   "par1.frx":01C6
         TabIndex        =   138
         Text            =   "Combo1"
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton calcula 
         Caption         =   "υπολογισμός"
         Height          =   240
         Left            =   0
         TabIndex        =   130
         Top             =   3960
         Width           =   3015
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "email"
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   110
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox Combo2PARAKRATISI 
         Height          =   315
         ItemData        =   "par1.frx":01D4
         Left            =   0
         List            =   "par1.frx":01E7
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   3240
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.ComboBox ComboPARAKRATISI 
         Height          =   315
         ItemData        =   "par1.frx":0253
         Left            =   0
         List            =   "par1.frx":0255
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   3600
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox tel_pliroteo 
         BackColor       =   &H000080FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   106
         Top             =   4560
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox parakratisiT 
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   104
         Top             =   4200
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.ComboBox apallagesFPA 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   45
         TabIndex        =   88
         Top             =   2685
         Width           =   2835
      End
      Begin VB.CheckBox metasx 
         BackColor       =   &H00FF8080&
         Caption         =   "Μετασχ.Παραστατικ"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1800
         TabIndex        =   85
         ToolTipText     =   "Με διόρθωση μπορώ να μετατρέψω το παραστατικό χωρίς να χαθεί το αρχικό παραστατικό"
         Top             =   600
         Width           =   1908
      End
      Begin VB.TextBox metaf 
         Height          =   285
         Left            =   120
         TabIndex        =   84
         Top             =   5040
         Width           =   825
      End
      Begin VB.TextBox isotimia 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   83
         Text            =   "1"
         Top             =   5040
         Width           =   690
      End
      Begin VB.CheckBox proepiskopisi 
         BackColor       =   &H00FF8080&
         Caption         =   "Προεπισκόπηση"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox ar_paraggt 
         Height          =   285
         Left            =   480
         TabIndex        =   81
         Top             =   1100
         Width           =   735
      End
      Begin VB.CheckBox SYSKEYASIA 
         BackColor       =   &H00FF8080&
         Caption         =   "Συσκευασία"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1521
         TabIndex        =   80
         Top             =   1170
         Value           =   1  'Checked
         Width           =   1452
      End
      Begin VB.CommandButton submitpayment 
         Caption         =   "submitpayment"
         Height          =   360
         Left            =   3000
         TabIndex        =   187
         Top             =   6240
         Width           =   990
      End
      Begin MSForms.CheckBox AnastolhFPA 
         Height          =   300
         Left            =   360
         TabIndex        =   155
         Top             =   5520
         Width           =   2500
         BackColor       =   16711680
         ForeColor       =   65280
         DisplayStyle    =   4
         Size            =   "4410;529"
         Value           =   "0"
         Caption         =   "Αναστολή Πληρωμής ΦΠΑ"
         FontHeight      =   165
         FontCharSet     =   161
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblΠληρωτέο 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Πληρωτέο"
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
         Height          =   195
         Left            =   120
         TabIndex        =   107
         Top             =   4560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblΠαρακράτηση 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Φόροι/Παρακρ"
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
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   4320
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Απαλλαγές ΦΠΑ"
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
         Height          =   247
         Left            =   117
         TabIndex        =   89
         Top             =   2457
         Width           =   2639
      End
      Begin VB.Label ar_paragg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Αριθ.Παρ"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   79
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Ισοτιμία"
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
         Height          =   210
         Left            =   1080
         TabIndex        =   78
         Top             =   4800
         Width           =   1065
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "άλλα έξοδα"
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
         Height          =   360
         Left            =   0
         TabIndex        =   77
         Top             =   4800
         Width           =   1230
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Γεν.Εκπτωση"
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
      Left            =   6030
      TabIndex        =   31
      Top             =   6930
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "DEL ADDRES              /SXOLIA"
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
      Left            =   5640
      TabIndex        =   38
      Top             =   7260
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "cont          Αρ.Αυτοκ."
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
      Height          =   255
      Left            =   6030
      TabIndex        =   25
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label SynoloG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8850
      TabIndex        =   15
      Top             =   6000
      Width           =   1470
   End
   Begin VB.Label Synolo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8850
      TabIndex        =   13
      Top             =   5100
      Width           =   1470
   End
   Begin VB.Label sfpa 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9300
      TabIndex        =   14
      Top             =   5520
      Width           =   1020
   End
   Begin VB.Label lblΟΔΟΣ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Φ.ΟΔΟΣ"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5640
      TabIndex        =   184
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label lblΠΟΛΗ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Φ.ΠΟΛΗ"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5640
      TabIndex        =   183
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblTK1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Φ.TK"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7560
      TabIndex        =   182
      Top             =   6240
      Width           =   405
   End
   Begin VB.Label lblLabel32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label32"
      Height          =   195
      Left            =   16560
      TabIndex        =   177
      Top             =   10920
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   372
      Left            =   17040
      Top             =   1320
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label RAFI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10440
      TabIndex        =   151
      Top             =   3000
      Width           =   105
   End
   Begin VB.Label ARDEH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3000
      TabIndex        =   148
      Top             =   720
      Width           =   45
   End
   Begin VB.Label lbFPAT 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "ΦΠΑ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   8880
      TabIndex        =   146
      Top             =   5520
      Width           =   405
   End
   Begin VB.Label email 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   7440
      TabIndex        =   140
      Top             =   600
      Width           =   45
   End
   Begin VB.Label par11Value 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   192
      Left            =   6360
      TabIndex        =   92
      Top             =   10680
      Width           =   36
   End
   Begin VB.Label PAR_GIA_AKYROSH 
      Height          =   255
      Left            =   10560
      TabIndex        =   86
      Top             =   8880
      Width           =   1935
   End
   Begin VB.Label ORA 
      Height          =   375
      Left            =   14160
      TabIndex        =   75
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LabTeliko 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Πληρωτέο"
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
      Height          =   150
      Left            =   14730
      TabIndex        =   65
      Top             =   9165
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label LabfpaEFK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ΦΠΑ ΕΦΚ"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   14640
      TabIndex        =   71
      Top             =   8685
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LabSynEfk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ΕΦΚ&"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   14835
      TabIndex        =   70
      Top             =   8205
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lab_strogg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   192
      Left            =   10440
      TabIndex        =   67
      Top             =   5570
      Width           =   36
   End
   Begin VB.Label Label22 
      BackColor       =   &H000000FF&
      Caption         =   "Προσοχή οι τιμες είναι με ΦΠΑ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12240
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "DESTIN    proorismos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   11970
      TabIndex        =   56
      Top             =   12135
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "DEL.TERMS    fort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   15480
      TabIndex        =   55
      Top             =   11160
      Width           =   690
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "DEL.TIM      skopos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   5610
      TabIndex        =   54
      Top             =   5520
      Width           =   660
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "PAYM.TERM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   5595
      TabIndex        =   53
      Top             =   5220
      Width           =   765
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   10440
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2955
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "= : ΕΙΔΗ ΠΕΛΑΤΗ ME TEΛ.TIMH  F5:Τελευταία Τιμή     F8:Ποσότητα από Αξία                   F9:Τιμή χωρίς ΦΠΑ F3:SN"
      ForeColor       =   &H8000000E&
      Height          =   585
      Left            =   0
      TabIndex        =   50
      Top             =   7560
      Width           =   5775
   End
   Begin VB.Label Label15 
      Caption         =   ".."
      Height          =   255
      Left            =   5880
      TabIndex        =   49
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Σειρά"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13920
      TabIndex        =   48
      Top             =   360
      Width           =   630
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Αριθμός"
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
      Height          =   255
      Left            =   4740
      TabIndex        =   47
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Παραστατικό"
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
      Left            =   1560
      TabIndex        =   46
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Είδος"
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
      Left            =   120
      TabIndex        =   45
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Αποθήκη"
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
      Height          =   255
      Left            =   9120
      TabIndex        =   44
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Πωλητής"
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
      Height          =   255
      Left            =   8160
      TabIndex        =   43
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   42
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   41
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Επώνυμο"
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
      Left            =   1080
      TabIndex        =   40
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label10 
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
      Left            =   120
      TabIndex        =   39
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   12000
      Left            =   10440
      Top             =   0
      Width           =   4770
   End
End
Attribute VB_Name = "par1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FORMA 41
Option Explicit
Dim F_EKTYPOTHS           As String '= FINDPARAMETROI(1, "PAR1", "F_EKTYPOTHS", "..", "Εκτυπωτής παραστατικών(με .. προεπιλεγμένος)") '
'' σταθερές για τις στήλες της τιμολόγησης
'Const f_k = 1
'
'Const f_o = 2
'
'Const f_p = 3
'
'Const F_T = 4
'
'Const f_e = 5
'
'Dim f_f ' = 11  '6  ''ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
'
'Const f_a = 7
'
'Const F_M = 8
'
'Const F_mefpa = 9
'
'Const F_PROEL = 10
'
'Const F_mikta = 12

Dim f_pelono, f_peldie, f_pelarit, f_pelpol, f_peltk As String

Dim F_VB6RESIZEFORM As Integer '= Val(FINDPARAMETROI(1, "PAR1", "F_VB6RESIZEFORM ", "0", "1=RESIZE H F???? 0=????????? OS ????")) '

Dim F_meIFESTO As Integer
Dim f_isdiakin As Integer  ' ΕΙΝΑΙ ΔΙΑΚΙΝΗΣΗ  0=ΟΧΙ  1-ΔΙΑΚ+ΑΞΙΑ  2=ΜΟΝΟ ΔΙΑΚΙΝΗΣΗ

Dim f_polagor As Integer
Dim f_EkkremhDeltia As Integer
'
'Dim F_EKPT2  '  = 6  '11   'ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
Dim PAROX_URLADD          As String
' σταθερές για τις στήλες της τιμολόγησης

Dim F_PROEPILFPA As Integer
Dim f_AKRIBOS_TON_KODIKO  As Integer
Dim F_GRIDLISTASTIMOL     As String
Dim F_GRID_WIDTH          As Integer '8 MEGALONEI TA GRAMMATA TOY GRID

Dim F_GRID_FONT_SIZE      As Integer '8 MEGALONEI TA GRAMMATA TOY GRID
Dim F_KYBRAN              As Integer
Dim F_2EKTYPOTHS          As String ' ; EKTYPOTIS AKIROTIKON
Dim fMonimo_MIDENIKO_FPA  As Integer
Dim F_PAROCHOS            As Integer '= Val(FINDPARAMETROI(1, "PAR7MYDATA", "f_parochos", "0", "0=MyData 1=???????"))
Dim F_YPOX_AFM            As Integer ' YPOXREOTIKO AFM STO TIMOLOGIO = IIf(nNull(R!GEF_M) = 99, 1, 0)
Dim fS_Parakratisis       As String ' string parakratisis
Dim F_ARPART_IN_EID       As Integer
Dim F_TIMH_XONDR_TIM_AGOR As Integer
Dim f_k                   As Integer '= 1
Dim f_o                   As Integer '= 2
Dim f_p                   As Integer ' = 3
Dim F_T                   As Integer ' = 4
Dim f_e                   As Integer ' = 5
Dim f_f                   As Integer ' = 11  '6  ''ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
Dim f_a                   As Integer ' = 7
Dim F_M                   As Integer '= 8
Dim F_mefpa  ' = 9
Dim F_PROEL      As Integer ' = 10
Dim F_COL_KOLA   As Integer '= 8
Dim F_KOLA_MIKTA As Integer
Dim F_mikta   '= 12  'XREIAZONTAI GIA TO SYGKENTRVTIKO DELTIO APOSTOLHS

Dim F_EKPT2  '  = 6  '11   'ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
Dim F_mikta_KILA   '= 12



Dim F_CUSTOMER_SCREEN As Integer
Dim F_CUSTOMER_SCR_LEFT As Long


Dim f_colEFK ' 13  sthlh eidikoy foroy
Dim F_GRIDCOLOR               As Long ' COLOR OF GRID1 INVOICE ENTRY
'Dim F_EIDOS_PAR As String

' Dim F_TIMH_XONDR_TIM_AGOR As Integer
Dim F_PosostaOTA              As Single

Dim f_METASX_KAI_HMER         As Integer '1=STA SX.ATIM+HMER  0=ATIM

Dim F_EXCEL_IMPORT            As Long

Dim f_MONO_PELATES_KATHGORIAS As Integer '1=MONO KATHGORIAS  0=XVRIS ELEGXO
Dim F_METAFORIKES             As Integer '1=MONO KATHGORIAS  0=XVRIS ELEGXO

Dim F_METAF3SYNAL             As Integer 'ΜΕΤΑΦΟΡΙΚΕΣ ΜΕ 3 ΣΥΝΑΛΛΑΣΟΜΕΝΟΥΣ =1 ΟΧΙ=0
Dim F_ELAIO_DIK   As Single ' % ΔΙΚΑΙΩΜΑ ΕΛΑΙΟΤΡΙΒΕΙΟΥ ΓΙΑ ΚΟΣΤΟΣ ΕΛΑΙΟΘΛΙΨΗΣ
Dim F_PARAS_3SYNAL            As String 'ΠΑΡΑΣΤΑΤΤΙΚΟ ΓΙΑ ΜΕΤΑΦΟΡΕΣ -ΑΠΟΣΤΟΛΕΣ-ΠΑΡΑΛΗΠΤΗ -ΕΝΤΟΛΕΑ
Dim FKODPELSYGK As String 'ΚΩΔΙΚΟΣ ΠΕΛΑΤΗ ΣΥΓΚΕΝΤΡΟΤΙΚΟΥ


'  ΓΙΑΤΙ Η ΙΔΙΑ ΕΠΙΧΕΙΡΗΣΗ ΕΙΝΑΙ ΚΑΙ ΒΥΤΙΟ ΚΑΙ ΜΕΤΑΦΟΡΕΑΣ Π.Χ. ΤΣΙΑΤΣΙΑΡΗΣ ΗΛΙΑΣ
Dim F_PDFNAME                 As String

Dim f_CAPTION_METASX_DA       As String  ' = FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_DA", "Iaoao?ciaoeoiuo oa AA", "AoeeYoa iaoao?ciaoeoiiy oa AA")
Dim f_CAPTION_METASX_TIM      As String  '  = FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_TIM", "Iaoao?ciaoeoiuo oa TIM", "AoeeYoa iaoao?ciaoeoiiy oa TIM")
Dim f_CAPTION_METASX_LIAN     As String   '= FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_LIAN", "Iaoao?ciaoeoiuo oa EIAN", "AoeeYoa iaoao?ciaoeoiiy oa EEAI")
  
Dim f_TIMES_METASX_DA         As String  ' = FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_DA", "Iaoao?ciaoeoiuo oa AA", "AoeeYoa iaoao?ciaoeoiiy oa AA")
Dim f_TIMES_METASX_TIM        As String  '  = FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_TIM", "Iaoao?ciaoeoiuo oa TIM", "AoeeYoa iaoao?ciaoeoiiy oa TIM")
Dim f_TIMES_METASX_LIAN       As String   '= FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_LIAN", "Iaoao?ciaoeoiuo oa EIAN", "AoeeYoa iaoao?ciaoeoiiy oa EEAI")
  
Dim f_same_eline              As Integer '=0 μονο στην επανεκτυπωση να κραταει το ιδιο eline  (tim.eline)
  
Dim f_elga_typoi_krat         As String
  
Dim f_hotel                   As Integer  ' αυτοματο φορο διαμονης
  
Dim f_parking_auto            As Integer
 
Dim f_CAPTION_MTF11           As String
Dim f_CAPTION_MTF12           As String
Dim f_CAPTION_MTF13           As String
Dim f_CAPTION_MTF14           As String
 
Dim f_CAPTION_MTF21           As String

Dim f_CAPTION_label24 As String
Dim f_CAPTION_label25 As String

Dim F_ANAPODO As String

'Dim f_CAPTION_MTF22 As String
'Dim f_CAPTION_MTF23 As String
 
Dim f_CAPTION_labf0           As String
Dim f_CAPTION_labf1           As String
Dim f_CAPTION_labf2           As String
Dim f_CAPTION_labf3           As String
 
Dim F_MONOENERGOS             As Integer
 
Dim f_TimhForoyDiuamonis
 
'Dim f_CAPTION_METASX_TIM As String  '  = FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_TIM", "Iaoao?ciaoeoiuo oa TIM", "AoeeYoa iaoao?ciaoeoiiy oa TIM")
'Dim f_CAPTION_METASX_LIAN As String   '= FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_LIAN", "Iaoao?ciaoeoiuo oa EIAN", "AoeeYoa iaoao?ciaoeoiiy oa EEAI")
  
'  MDIForm1.METASX_DA.Caption = f_CAPTION_METASX_DA
'  MDIForm1.METASX_TIM.Caption = f_CAPTION_METASX_TIM
'  MDIForm1.metasx_lian.Caption = f_CAPTION_METASX_LIAN
'

Dim F_PLATOSGRIDEIDON      As String

Dim f_IS_PYROSB            As Integer '1=pyrosbesthres  0=oxi
Dim f_auto_kostologhsh     As Integer '1=pyrosbesthres  0=oxi
Dim f_strogg               As Single 'ποσό στρογγυλοποίησης ΦΠΑ
Dim f_strogg_FPA           As Single 'ΚΑΤΗΓΟΡΙΑ ΦΠΑ ΠΟΥ ΘΑ ΓΙΝΕΙ Η ΣΤΡΟΓΓΥΛΟΠΟΙΗΣΗ
'Dim f_PEL_AYTOKINHTO As Integer
Dim f_akyr_idia_mera       As Integer

Dim f_akyr_neo             As Integer  ' αν ακυρωνει με νεο ακυρωτικο (αντιθετο του παραστατικου)

Dim F_CHANGE_TELIKO_POSO   As Integer  '= 1 Then  ΚΑΝΟΝΙΖΕΙ ΤΟ ΠΟΣΟ

Dim F_EKPT_PARAGG          As Integer ' Η ΕΚΠΤΩΣΗ ΠΑΡΑΓΓΕΛΙΑΣ ΑΝ ΘΑ ΜΕΤΑΦΕΡΕΤΑΙ ΣΤΟ ΜΕΤΑΣΧ.ΠΑΡΑΣΤΑΤΙΚΟ Π.Χ. ΤΙΜΟΛΟΓΙΟ

Dim f_is_benzinadiko       As Integer
Dim f_AKYR_benzinadiko     As Integer

Dim f_AKYR_benzPAR         As String

Dim f_EID_AKYR_APOD        As String  ' ΠΟΙΟ ΠΑΡΑΣΤΑΤΤΙΚΟ ΑΚΥΡΩΝΕΙ ΤΗΝ ΑΠΟΔΕΙΞΗ ΤΟΥ ΕΙΣΡΟΩΝ
Dim F_PAR_AYTOK            As String  '           'L','T'

Dim F_THERM                As Integer  ' ΕΙΝΑΙ ΦΟΡΤΗΓΟ ΘΕΡΜΑΝΣΗΣ
Dim f_SDA                  As String '  'H;J  H=ΦΟΡΤΩΣΗς    J = ΕΠΙΣΤΡΟΦΗΣ
Dim F_THER_TIM_XORISAPOD   As Integer

Dim F_IS_BYTIO             As Integer  ' ΕΙΝΑΙ ΦΟΡΤΗΓΟ BYTIO METAFORIKHS

Dim f_2TIMOL_1PAGE         As Integer

Dim F_EIDFOROSKAT          As Integer ' ΕΙΔΙΚΟΣ ΦΟΡΟΣ ΚΑΤΑΝΛΩΣΗΣ ΑΝ ΔΟΥΛΕΥΕΙ ή ΌΧΙ

Dim F_GRID1WKOD            As Long  ' = Val(FINDPARAMETROI(7, "PAR1", "F_GRID1WKOD", "1200", "GRID ΕΙΔΩΝ ΦΑΡΔΟΣ ΚΩΔ."))
Dim F_TYPFOROL             As Integer ' TYPOS FOROLOGIKOY 1=ALGO 2=SIGNPRO  21=rbs101(ics like)

Dim F_KOLA_MIKTA_LABELS    As Integer ' = Val(FINDPARAMETROI(1, "PAR1", "F_KOLA_MIKTA_LABELS", "1", "Γ.37G.ΕΤΙΚΕΤΕΣ ΠΡΟΕΛ/ΜΙΚΤ= 1  ΜΗΚ/ΠΛΑΤΟΣ ΓΙΑ PACKING LIST=2")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))

Dim F_GRID1WONO            As Long  '= Val(FINDPARAMETROI(7, "PAR1", "F_GRID1WONO", "3800", "GRID ΕΙΔΩΝ ΦΑΡΔΟΣ ΟΝΟΜΑ"))
Dim F_GRID1WPOSO           As Long '= Val(FINDPARAMETROI(7, "PAR1", "F_GRID1WPOSO", "1000", "GRID ΕΙΔΩΝ ΦΑΡΔΟΣ ΠΟΣΟΤΗΤΑ"))
Dim F_GRID1WTIMH           As Long '  = Val(FINDPARAMETROI(7, "PAR1", "F_GRID1WTIMH", "1000", "GRID ΕΙΔΩΝ ΦΑΡΔΟΣ TIMH"))
Dim F_GRID1WEKPT           As Long '= Val(FINDPARAMETROI(7, "PAR1", "F_GRID1WEKPT", "1000", "GRID ΕΙΔΩΝ ΦΑΡΔΟΣ ΕΚΠΤΩΣΗ"))
Dim F_GRID1WFPA            As Long ' = Val(FINDPARAMETROI(7, "PAR1", "F_GRID1WFPA", "1000", "GRID ΕΙΔΩΝ ΦΑΡΔΟΣ ΦΠΑ"))

Dim F_DEFA_PLHR            As Integer  ' DEFAULT TROPOS PLIROMIS

Dim f_pedioAFM             As String
Dim f_LabAFM               As String

Dim F_TYPOS_PAR            As String

Dim F_2EKPTOSEIS           As Integer

Dim f_KENTRA_ADYNATISMATOS As Integer

Dim F_ID_NUM_DIOR          As Long ' ΤΟ ID_NUM TOY TIM ΠΟΥ ΘΑ ΔΙΟΡΘΩΣΩ
Dim f_IS_PARTIDES          As Integer

'Dim F_REMOTE As Integer
Dim F_LIANIKES_TIMES       As Integer '1=ΔΟΥΛΕΥΩ ΛΙΑΝΙΚΕΣ ΤΙΜΕΣ ΣΤΑ ΠΑΡΑΣ/ΚΑ ΠΟΥ ΕΧΟΥΝ ΤΗΝ ΕΝΔΕΙΞΗ ΛΙΑΝΙΚΗ

Dim fbuff                  As String

Dim XROMATA(100, 25, 25)    'κρατάει σε πίνακα τα χρώματα

Dim f_YPOK As Integer    ' 0=kentriko  1=1o ypok  2= 2o ypok  3= 3o ypok

Dim F_YPOL_GEN  ' NA EMFANIZEI TA EIDH ASXETA AN EXOYN KINHΘΕΙ

Dim f_loadpar1           As Integer    '1 FORTONEI THN FORMA SYNEXEIA

Dim F_TRANS              As Long    ' 1=BEGINTRANS   - ROLLBACK      0=NOTHING

Dim F_EIKONA             As Long '

Dim f_mono_apoParaggelia As Integer    ' mono apo Paraggelia TIMOLOGO

Dim F_POLHTES            As Integer

Dim f_SYMPYKNVSH         As Integer    ' ΣΥΜΠΥΚΝΩΣΗ ΚΩΔΙΚΩΝ

Dim F_POLHTES_PERIOR     As String

Dim F_PLATOS_FORMAS      As Long

Dim F_YCOS_FORMAS        As Long

Dim F_ELGA               As Single

Dim F_DEBUG              As Long  ' ΑΝΤΙΓΡΑΦΩ ΤΑ ΑΡΧΕΙΑ ΕΓΓΤΙΜ ΤΟΥ ΤΙΜ. ΠΟΥ ΕΒΓΑΛΕ ΑΣΥΜΦΩΝΙΑ ΠΑΡΑΣΤΑΤΙΚΟΥ

Dim F_8                  As Integer ' 1=YPOLOGIZEI POSOTHTA APO AXIA (BENZINADIKA)  2=YPOLOGIZEI TIMH APO AXIA '4-ΥΠΟΛΟΓΙΖΕΙ ΤΙΜΗ ΜΕ ΑΦΑΙΡΕΣΗ 1+0,5 Χ  ΜΟΝΑΔΕΣ

Dim F_PIN_PROELEYS       As Integer    '  1=ΠΡΟΕΛΕΥΣΕΙΣ ΑΠΟ ΠΙΝΑΚΑ
'  ΑΝ ΔΙΠΛΑ ΑΠΟ ΤΗΝ ΠΡΟΕΛΕΥΣΗ ΒΑΛΩ ΚΩΔΙΚΟ ΠΑΡΤΙΔΑΣ ΝΑ ΒΡΙΣΚΕΙ ΤΟΝ ΚΩΔΙΚΟ ΑΠΟ ΠΙΝΑΚΑ ΚΑΙ ΝΑ ΒΑΖΕΙ ΑΠΟ 16-35 ΤΗΝ ΠΡΤΙΔΑ
'αν βάλω 11 STO F_PIN_PROELEYS βαζει στην αποθηκη την προελευση δηλαδή APOT=PROELEYSH
' an balo 2=sakis proeleysh = eid.erg

Dim F_FORMA1             As String, F_FORMA2 As String, F_FORMA3 As String

Dim F_LOCKARITMISI       As Integer     '  1=ΚΛΕΙΔΩΝΕΙ ΤΗΝ ΣΥΝΕΧΕΙΑ ΤΗΣ ΑΡΙΘΜΗΣΗΣ

Dim F_PINAKAS_PROEL(200) As String   '  PINAKAS ME ΠΡΟΕΛΕΥΣΕΙΣ

Dim f_LEKTIKO_ARPARAGG   As String    ' ΛΕΚΤΙΚΟ ΠΟΥ ΘΑ ΓΡΑΦΕΙ ΣΤΑ ΣΧΟΛΙΑ   "Aρ.Παραγγ."  Π.Χ. 123

Dim F_SET_EIDON          As Integer ' ΣΕΤ ΕΙΔΟΝ ΠΟΥ ΧΡΗΣΙΜΟΠΟΙΟ

Dim F_PERIOR_POLHTES     As Integer ' ΠΕΡΙΟΡΙΖΕΙ ΤΟΥΣ ΠΩΛΗΤΕΣ ΣΥΜΦΩΝΑ ΜΕ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ

Dim f_AFHNO_IDIA_SEIRA   As Integer 'ΔΕΝ ΑΦΗΝΕΙ ΙΔΙΑ ΣΕΙΡΑ AN (κωδ<>παλ.κωδ ή ποσ=0) =1    (KOD<>PALIO ή (ποσ=0 & τιμη=0) )=2

' If Grid1.Col < f_p And (Grid1.Text <> TIMText3.Text Or ( Val(Grid1.TextMatrix(Grid1.row, f_p))) = 0)    =1

Dim F_POLLES_SEIRES ' αν θα τυπωνει σε πολλες σειρές την περιγραφή

Dim sqlins(300)         As String

Dim F_AJIDOROKARTAS     As Single

'Dim f_Capitals As Integer
Dim F_PDF               As Integer    ' STELNEI SE PDF

' Dim F_PDFfolder2 As Integer    ' prosorina PDF
Dim F_METRHTA           As Integer    ' ZHTA METPHTA

Dim F_PARAG_KART        As Integer    ' ΝΑ ΒΓΑΙΝΕΙ Η ΠΑΡΑΓΓΕΛΙΑ ΣΤΗΝ ΚΑΡΤΕΛΛΑ

Dim F_UPPER             As Integer ' αν θα τα μετατρεπει σε κεφαλαια ελληνικα ή όχι

Dim F_METRHTAposo       As Single    ' ZHTA METPHTA

Dim F_PARAKRATISI       As Single ' AN EXEI PARAKRATISI TO PARASTATIKO P.X. PAROXHS =20%

Dim f_AmesosPelatisLian As Integer

Dim f_AkyrParas         As String

Dim F_AR_parametroi(200, 2)    ' ΦΟΡΤΩΝΩ ΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ ΣΕ ΠΙΝΑΚΑ ΓΙΑ ΝΑ ΜΗΝ ΑΡΓΕΙ ΣΤΟ ΞΕΚΙΝΗΜΑ ΤΗΣ ΦΟΡΜΑΣ

Dim F_SHMADEMENA_DELTIA As String    ' ΔΕΛΤΙΑ ΠΟΥ ΘΑ ΤΙΜΟΛΟΓΗΘΟΥΝ

Dim F_MONOMEYPOLOIPO    As Integer

Dim f_barcode           As Integer    '1=diabazei apeyueias ta barcode

Dim f_LIANIKHCHECK      As Integer    'ΤΣΕΚΑΡΙΣΜΕΝΗ Η ΛΙΑΝΙΚΗ

Dim f_TIMOKAT           As Long    ' τιμοκατάλογος

Dim f_paras             As String    ' ΚΕΙΜΕΝΟ ΤΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ

Dim f_apeyueias         As String

Dim F_add_barc

Dim f_BARCODE_STOTIM   As Integer   ' 1=ΒΑΖΕΙ ΤΟ ERG ΣΤΟ ΤΙΜΟΛΟΓΙΟ 0=DEFAULT ΤΟΝ ΚΩΔΙΚΟ

Dim F_falcon_Polla_tim As Integer    ' pairnei polla timologia apo to falcon

Dim f_PALIAXRONIA

Dim f_myEID       As String ' ΤΥΠΟΣ ΠΑΡΑΣΤΑΤΤΙΚΟΥ MYDATA  Π.X. 1.1 ΤΙΜΟΛΟΓΙΟ
Dim f_ISMYDATA    As Integer ' 0 anenergo 1=energo gia mydata

Dim f_checkMYDATA As Integer ' 1= online apostolh
Dim f_MYDATAcont  As Integer ' 1= online apostolh

Dim f_FORM_EKTY    'parastat!FORM_EKTY  φορμα εκτύπωσης=99  1=crystal -> timol1 timol1an

Dim f_SeiresTimologioy    'POSA EIDH MPORO NA PERASO

Dim f_SynSeiresTimologioy    'γραμμενες+κενες σειρές

Dim f_MaxSeiresTimologioy    'POSA EIDH MPORO NA PERASO Max TIMOLOGIA Agoras

Dim f_MinSeiresTimologioy    'POSA EIDH MPORO NA PERASO Min TIMOLOGIA polisis

Dim f_Timol2_SeiresTimologioy    'POSA EIDH MPORO NA PERASO sthn forma timol2

Dim F_DESMIA_SEIRES As Integer  ' 0 =KANONIKES SEIRES   2=SEIRES FORMA2





Dim F_TYPONO_KENES     As Integer 'ΑΝ ΘΑ ΤΥΠΩΝΩ ΚΕΝΕΣ ΣΕΙΡΕΣ ΣΤΟ ΤΙΜΟΛΟΓΙΟ

Dim f_AutoChangeLine   As Long    'def="1"  "1=μπορω να αλάξω έκπτωση γραμμής 0=δεν μπορώ

Dim f_CanChangeEkptosi As Long    'def="1"  "1=μπορω να αλάξω έκπτωση γραμμής 0=δεν μπορώ

Dim f_CanChangeFpa     As Long    'def="1"  "1=μπορω να αλάξω το φπα γραμμής 0=δεν μπορώ

Dim f_checkYpol        As Long    'def="0"  "1=ΕΛΕΓΧOΣ 2=ΑΠΑΓΟΡΕΥΣΗ 0=ΤΙΠΟΤΑ")) ελεγχος υπολοιπου

Dim f_psifia_atim      As Long    'DEF=5  T12345A   KALA EINAI NA EXO 6 T123456A

Dim f_PARAGGMARKUP    'paraggelia  pelatoy  0=xondriki   1=me mark up

Dim f_psifiaAjias    'posa psifia tha exei h kathe seira

Dim f_psifiaTIMMON    'posa psifia tha exei h kathe seira

Dim f_1choice_polhseis, f_1choice_agores    ' οταν διαλέγει αγορές ή πωλήσεις σε ποιό παραστατικό να κάθεται
Attribute f_1choice_agores.VB_VarUserMemId = 1073938470

Dim f_dior_mAtim, f_dior_HME    ' το τιμολόγιο που θα πάω να διορθώσω (ενδέχεται να το αλλάξω αριθμό ή ημερομηνία και έτσι χάνω το αρχικό που θα διόρθωνα )
Attribute f_dior_mAtim.VB_VarUserMemId = 1073938472
Attribute f_dior_HME.VB_VarUserMemId = 1073938472
Dim f_dior_ARX_TYPOS    As String    ' O TYPOS ΠΑΡΑΣΤΑΤΙΚΟΥ  που θα πάω να διορθώσω (ενδέχεται να αλλάξω ΤΥΠΟ ΟΠΟΤΕ ΠΡΕΠΕΙ ΝΑ ΠΑΡΕΙ ΝΕΑ ΑΡΙΘΜΗΣΗ ΑΠΟ ΤΟΝ ΝΕΟ ΤΥΠΟ
'      ΔΗΛΑΔΗ ΑΝ ΔΙΟΡΘΩΝΩ ΤΟ ΤΙΜ.ΠΩΛΗΣΗΣ ΚΑΙ ΤΟ ΚΑΝΩ ΔΕΛΤΙΟ ΑΠΟΣΤΟΛΗΣ ΝΑ ΠΑΡΕΙ ΤΗΝ ΑΡΙΘΜΗΣΗ ΤΟΥ ΔΑ

Dim F928                As Integer

'Dim F_PDF As String
Dim F_PDFFOLDER         As String

Dim F_STOPDELTIA        As Integer    'ΝΑ ΜΠΛΟΚΑΡΕΙ ΑΝ ΕΚΚΡΕΜΟΥΝ ΔΕΛΤΙΑ

Dim F_2CHFIA_ZYGIZOMENA As String  '2 ΠΡΩΤΑ ΨΗΦΙΑ ΒΟΗΘ.ΚΩΔ.ΤΩΝ ΖΥΓΙΖΟΜΕΝΩ

Dim f_mc                As Integer    ' σειρες grid γεφυρωσης

Dim f_gkod_pel
Attribute f_gkod_pel.VB_VarUserMemId = 1073938478

Dim SumFpa(9)  ' συνολο ΦΠΑ ΚΑΤΑ ΣΥΝΤΕΛΕΣΤΗ
Attribute SumFpa.VB_VarUserMemId = 1073938479

Dim f_Akyr As Boolean

Dim F_DIOR    ' ΑΝ ΕΙΝΑΙ ΔΙΟΡΘΩΣΗ =1    ΑΝ ΟΧΙ ΔΙΟΡΘΩΣΗ=0
Attribute F_DIOR.VB_VarUserMemId = 1073938481

Dim f_error_gefyr As Integer  ' αν είναι 0 δεν έχει λάθος 1=λάθος στον κωδικό πελάτη

' 2=λάθος στον κωδικό του είδους
Dim f_gDek_ajia
Attribute f_gDek_ajia.VB_VarUserMemId = 1073938483

Dim f_sajia(900), f_fpa(900)    ' κρατά στην μνήμη τοσύνολο κάθε σειράς και το ΦΠΑ
Attribute f_sajia.VB_VarUserMemId = 1073938484

Dim fSCR As Object    ' SCRIPT CONTROL

Dim FlagGridEidon
Attribute FlagGridEidon.VB_VarUserMemId = 1073938487

Dim f_problebgef
Attribute f_problebgef.VB_VarUserMemId = 1073938488

Dim f_matim As String

'Dim f_checkPos

Dim f_dat   As Date

Dim F_POLAGOR_LastValue
Attribute F_POLAGOR_LastValue.VB_VarUserMemId = 1073938492

Dim F_akyr_Delete    '  σβηνω το παραστατικό που ακυρώνεται
Attribute F_akyr_Delete.VB_VarUserMemId = 1073938493

Dim f_use_Sysk    As Integer    ' χρησιμοποιω συσκευασια = 1

Dim F_LOGPLHR(20) As String    ' ΛΟΓΑΡΙΑΣΜΟΙ ΤΡΟΠΟΩΝ ΠΛΗΡΩΜΗΣ

Dim F_PDFPARAGG  ' ΑΝ ΘΑ ΑΝΤΙΓΡΑΦΕΙ ΚΑΙ ΣΤΟ ΦΑΚΕΛΟ C:\PDFPARAGG ΤΑ ΠΑΡΑΑΣΤΑΤΙΚΑ (PDF) ΕΧΟΝΤΑΣ ΣΑΝ ΟΝΟΜΑ ΤΑ 6 ΑΡΙΣΤΕΡΑ ΨΗΦΙΑ ΑΠΟ ΤΙΣ ΠΑΡΑΤΗΡΗΣΕΙΣ

Dim f_mfpa(1 To 9) As Single    'οριζονται στο find_parastat

Dim f_kau(1 To 9)  As Single    'καθαρή αξία κατά φπα

Dim f_gfkod(20)    As String, f_gfaj(20) As Single

'METABLHTES EIDOYS PARASTATIKOY APO PARASTAT
Dim F_TITLOS, F_POS_APOU, F_STADIO, F_AJIA_APOU, F_XONDR, f_pel, f_metasx, f_pol, F_rec_p As Integer, F_rec_m As Integer    'παράμετροι του παραστατικού
Attribute F_TITLOS.VB_VarUserMemId = 1073938500
Attribute F_POS_APOU.VB_VarUserMemId = 1073938500
Attribute F_STADIO.VB_VarUserMemId = 1073938500
Attribute F_AJIA_APOU.VB_VarUserMemId = 1073938500
Attribute F_XONDR.VB_VarUserMemId = 1073938500
Attribute f_pel.VB_VarUserMemId = 1073938500
Attribute f_metasx.VB_VarUserMemId = 1073938500
Attribute f_pol.VB_VarUserMemId = 1073938500
Attribute F_rec_p.VB_VarUserMemId = 1073938500
Attribute F_rec_m.VB_VarUserMemId = 1073938500

Dim f_deltia(1 To 100)    'δελτια που σημαδεύω
Attribute f_deltia.VB_VarUserMemId = 1073938510
Dim f_deltiaMarks(1 To 100)    'τα μαρκ των δελτίων που σημαδεύω
Dim f_IDdeltia(1 To 100)    As Long  'ID δελτιων  που σημαδεύω για να μπορώ εύκολα να τα μαρκάρω

Dim F_File_Timologioy       As String    ' το όνομα του αρχείου που γίνεται η εκτύπωση π.χ. c:\print

Dim F_Type_File_Timologioy  As Integer    '1=typvnei to arxeio 0=den to typonei 2=me entolh print 3=c:\typ.bat

Dim F_File2_Timologioy      As String    ' το όνομα του αρχείου που γίνεται η εκτύπωση π.χ. c:\print

Dim F_Type2_File_Timologioy As Integer    '1=typvnei to arxeio 0=den to typonei

Dim f_F12                   As Integer, f_F1 As Integer  ' ΠΡΟΓΡΑΜΜΑΤΙΣΜΟΣ ΠΛΗΚΤΡΩΝ

Dim F_EXO_XROMATA           As Integer  '1=ΧΡΗΣΙΜΟΠΟΙΩ ΤΑ ΤΡΩΜΑΤΑ ΜΕΓΕΘΗ

Dim F_PAR_MONO_ME_YPOLOIPO  As Integer  '

'Dim F_EIDH_APOT As Integer ' να δειχνει μόνο τα είδη που εχουν κινηθεί στο υποκατάστημα

Dim f_Fkey    ' το πληκτρο που κάνει update
Attribute f_Fkey.VB_VarUserMemId = 1073938520

Dim f_Desmia    As Integer    ' χρησιμοποιω τις παραγγελίες για έλεγχο υπολοίπου  'βλεπε mdesmia

Dim F_EISITHRIA As Integer    '1=XRHΣIMOΠOIΩ ΕΙΣΙΤΗΡΙΑ

Dim f_pelMERC   As String    ' ποιος πελατης του mercury χρησιμοποιει το πρόγραμμα

Dim f_suma    ' teliko poso timologioy
Attribute f_suma.VB_VarUserMemId = 1073938523

Dim f_ColUpdate  ' flag αν ενημερώθηκε η κολόνα του dbgrid1
Attribute f_ColUpdate.VB_VarUserMemId = 1073938524

'Dim f_epitr_seir
Dim fDB         As Database  ' ανοιγω γενικά την database και με την open_data ανοίγω τον πίνακα

Dim ProthFora   As Boolean

Dim fSHMADEMENA As String

Dim mEidos_Par(50)    'ειδος παραστατικού . π.χ.  atim=T
Attribute mEidos_Par.VB_VarUserMemId = 1073938528

Dim F_YP2         As Integer  ' F_YP2 = Val(FindParametroi(1,"PAR1", "F_YP2", "1", "Γ. 12. ΑΡΙΘΜΟΣ ΠΡΟΕΠΙΛ.ΑΠΟΘ.ΠΡΟΟΡΙΣΜΟΥ"))

Dim F_EPIBEB_APOT As Integer

Dim F_METASX_PAR  As Integer  ' ΕΙΝΑΙ ΜΕΤΑΣΧΗΜΑΤΙΣΜΟΣ ΠΑΡΑΣΤΑΤΙΚΟΥ = 1

Dim f_site        As Integer

Dim Prospective   As New TrueOleDBGrid80.Style

Dim Distributors  As New TrueOleDBGrid80.Style

Dim Col           As TrueOleDBGrid80.Column
Dim ColS          As TrueOleDBGrid80.columns

Dim fh(0 To 200), fw(0 To 200), ft(0 To 200), fl(0 To 200)
Attribute fh.VB_VarUserMemId = 1073938529

Private Sub apallagesFPA_Change()
100     apallagesFPA.ToolTipText = apallagesFPA
        'Combo1.ToolTipText = Combo1
End Sub

Private Sub apallagesFPA_Click()
100     apallagesFPA.ToolTipText = apallagesFPA
End Sub

Private Sub ar_paraggt_KeyPress(KeyAscii As Integer)

100     If KeyAscii = 13 Then
102         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

End Sub

Private Sub ar_paraggt_LostFocus()
100     LASTTIMOL.ConnectionString = gConnect
    
        Dim querpol As String
102     If chkPOL.Value = vbChecked Then
104         If chkAgor.Value = vbChecked Then
106             querpol = "  B_N1=" + Left(APOTH1.Text, 1)
            Else
108             querpol = "  B_N1=" + Left(APOTH1.Text, 1) + " AND TIM.EIDOS='e'  "
           
            End If
        Else
110         If chkAgor.Value = vbChecked Then
112             querpol = "  B_N1=" + Left(APOTH1.Text, 1) + " AND TIM.EIDOS='r'  "
            Else
114             querpol = "  B_N1=" + Left(APOTH1.Text, 1)
           
            End If
         
        End If
  
116     If Len(ar_paraggt.Text) > 0 Then
118         LASTTIMOL.RecordSource = "SELECT top 5000 HME,ATIM,AJI AS [AΞΙΑ],PEL.EPO AS [ΕΠΩΝΥΜΙΑ],PEL.KOD AS [ΚΩΔΙΚΟΣ],PARAT AS [ΣΧΟΛΙΑ],SXETIKO AS [ΣΧΕΤ.ΠΑΡ],ID_NUM,ENTITYMARK,SXETMARK,MARKPLHR FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + querpol + " and ATIM LIKE '%" + ar_paraggt.Text + "%' ORDER BY ID_NUM DESC"
        Else
120         If Len(querpol) > 2 Then querpol = " where " + Left$(querpol, 85)
122         LASTTIMOL.RecordSource = "SELECT top 5000 HME,ATIM,AJI AS [AΞΙΑ],PEL.EPO AS [ΕΠΩΝΥΜΙΑ],PEL.KOD AS [ΚΩΔΙΚΟΣ],PARAT AS [ΣΧΟΛΙΑ],SXETIKO AS [ΣΧΕΤ.ΠΑΡ],ID_NUM,ENTITYMARK,SXETMARK,MARKPLHR  FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD " + querpol + " ORDER BY ID_NUM DESC"
        End If

124     LASTTIMOL.Refresh

126     If Not LASTTIMOL.Recordset.EOF Then
  
128         If Len(LASTTIMOL.Recordset(5)) > 1 Then
                'TDBGrid1.ROWS(0).BackColor = vbRed
130             ar_paraggt.BackColor = vbRed
            Else
132             ar_paraggt.BackColor = vbWhite
            End If
        End If

134     TDBGrid1.columns(6).Width = 1544.882
136     TDBGrid1.columns(5).Width = 2610.142
138     TDBGrid1.columns(4).Width = 480.189
140     TDBGrid1.columns(3).Width = 3550.142
142     TDBGrid1.columns(2).Width = 929.7639
144     TDBGrid1.columns(1).Width = 900.2835
146     TDBGrid1.columns(0).Width = 1049.953
  
End Sub

'dim constant a   '   select

'

Private Sub b_akyr_Click()

        '<EhHeader>
        On Error GoTo b_akyr_Click_Err

        '</EhHeader>
        Dim ANS As Integer

        Dim k

        Dim L

100     ANS = vbYes

102     If Len(Grid1.TextMatrix(1, 1)) > 0 Then
104         ANS = MsgBox("Να χαθούν οι αλλαγές;", vbYesNo)
        End If

106     If ANS = vbYes Then
             MDIForm1.StatusBar1.Panels(5).Text = ""

            '--------------------  initialization ----------------------------
108         synolo.Caption = ""
110         sfpa.Caption = ""
112         SynoloG.Caption = ""

114         For k = 1 To 9: f_kau(k) = 0: SumFpa(k) = 0: Next    ' συνολο ΦΠΑ ΚΑΤΑ ΣΥΝΤΕΛΕΣΤΗ
116         For k = 1 To 100: f_fpa(k) = 0: f_sajia(k) = 0: Next    ' κρατά στην μνήμη τοσύνολο κάθε σειράς και το ΦΠΑ
118         For k = 1 To 30: f_IDdeltia(k) = 0: f_deltia(k) = "": Next    'δελτια που σημαδεύω

            ' μηδενιζω το grid1
120         For L = 1 To Grid1.ColS - 1

                'Grid1.Col = l
122             For k = 1 To Grid1.rows - 1
                    'Grid1.Row = K
                    'Grid1.Text = ""
124                 Grid1.TextMatrix(k, L) = ""
                Next
            Next

126         If f_IS_PARTIDES = 1 Then
                'ΣΒΗΝΩ ΤΙΣ ΠΑΡΤΙΔΕΣ ΠΟΥ ΔΗΜΙΟΥΡΓΗΣΑ ΚΑΙ ΔΕΝ ΚΑΤΑΧΩΡΗΘΗΚΕ ΤΟ ΤΙΜΟΛΟΓΙΟ
128             Gdb.Execute "DELETE FROM PARTIDES WHERE POSOAG>0 AND ID_NUM=0"

            End If

130         Text2(0).Enabled = True
132         Text2(1).Enabled = True
134         Text2(2).Enabled = True
136         Text2(0).Text = ""    ' κωδ
138         Text2(1).Text = ""    'επωνυμια
140         Text2(2).Text = ""    'αφμ
142         parat.Text = ""    'ΠΑΡΑΤΗΡΗΣΕΙΣ
144         timtext3.Top = 18000
146         F_DIOR = 0    ' ΔΙΟΡΘΩΣΗ
148         metasxhmatismos.Enabled = True
150         f_loadpar1 = Val(FINDPARAMETROI(1, "PAR1", "F_LOADPAR1", "1", "ΦΟΡΤΩΝΩ ΚΑΘΕ ΦΟΡΑ ΤΗΝ ΦΟΡΜΑ ΠΑΡΑΣΤΑΤΙΚΩΝ=1"))

152         If f_loadpar1 = 1 Then
154             Unload Me
            Else
156             Me.Hide
            End If

            On Error Resume Next

158         SendKeys "%"
160         SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}"

            On Error GoTo b_akyr_Click_Err
        
        End If

        '<EhFooter>
        Exit Sub

b_akyr_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.b_akyr_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.b_akyr_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub b_Katax_Click()

        Dim j As Integer
100     If F_IS_BYTIO >= 1 Then
102         For j = 0 To 7
104             Textm(j) = Textm(20 + j)
            Next
        End If
  
106     If F_YPOX_AFM = 1 Then
108         If Len(Trim$(Text2(2).Text)) < 9 Then
110             MsgBox "Ξεχάσατε να βαλετε ΑΦΜ"
                Exit Sub
            End If

        End If
  
112     If Len(Trim(Text2(0).Text)) = 0 Then
114         MsgBox "ΔΕΝ ΔΙΑΛΕΞΑΤΕ ΠΕΛΑΤΗ"
            Exit Sub
        End If
  
        If Val(Left$(SKOPOS, 2)) = 19 Then
  
            If Len(other.Text) < 2 Then
       
                MsgBox "βαλτε τον αλλο σκοπό επάνω δεξιά"
                other.SetFocus
                Exit Sub
       
            End If
  
        End If

If HMERPARAD.Value < DTPicker1.Value Then
   HMERPARAD.Value = DTPicker1.Value
End If


fortosh.Text = FORTDIE.Text + ";" + fortARit.Text + ";" + FORTPOL.Text + ";" + FORTTK.Text




116     b_Katax.Enabled = False
   
118     If (f_myEID = "9.1" Or f_myEID = "5.1") And f_ISMYDATA = 1 Then   ' pistvtiko διαλεγω το σχετιζομενο παραστατικο
            If f_myEID = "9.1" Then
               par1b.PEL.Text = Trim(FKODPELSYGK)
            Else
120            par1b.PEL.Text = Text2(0)
            End If
122         par1b.SHOW 1
   
        End If

124     If f_ISMYDATA = 1 Then    ' pistvtiko διαλεγω το σχετιζομενο παραστατικο

126         If gVal(SynoloG.Caption) = 0 And InStr(" 9.1 9.2 9.3 10.1 10.2 8.2", f_myEID) = 0 Then ' mhdeniko kai den ανηκει στα δελτια
128             MsgBox "μηδενικά παραστατικά δεν επιτρέπονται στα MyData"
                Exit Sub
            End If
        End If
   
        'If f_myEID = "5.1" Then  ' pistvtiko διαλεγω το σχετιζομενο παραστατικο
        'If Left(PARAS.Text, 1) = "P" Then
        '   par1b.PEL.Text = Text2(0)
        '  par1b.SHOW 1
   
        'End If
   
130     SaveParastat False
   
132     b_Katax.Enabled = True
   
End Sub

Private Sub SaveParastat(ByVal isAkyrotiko As Boolean)

        '====================================================
        '<EhHeader>
        On Error GoTo SaveParastat_Err

        '</EhHeader>

        Dim wrkDefault As Workspace

        Dim R As New ADODB.Recordset

        Dim X, DUM

        Dim F_ATIM
100     fS_Parakratisis = ""

        Dim k, L

        Dim EIDOSKPE

        Dim m_tim, S_POSOTHTA As Single

        Dim par As New ADODB.Recordset, m_r

        Dim m_m

        Dim sql As String

        Dim s1 As Single

        Dim pdf2 As String
        Dim EXEI_MIDENIKO_FPA   ': EXEI_MIDENIKO_FPA = 1
        
102     If f_EID_AKYR_APOD = Right(PARAS.Text, 1) And f_hotel = 0 Then
        
104         If Val(ar_apod.Text) = 0 Then
106             MsgBox "Δεν συμπηρώσατε αποδειξη"
                Exit Sub
            End If
            
            ' το εκανα rem γιατι δεν ειναι ακυρωτικο αλλά αποδ.επιστροφης λιανικης
            ' PAR_GIA_AKYROSH.Caption = "173#" + ar_apod.Text
        
        End If
        
108     For k = 1 To Grid1.rows - 1
110         Grid1.TextMatrix(k, F_T) = Replace(Grid1.TextMatrix(k, F_T), ",", ".")
112         Grid1.TextMatrix(k, f_p) = Replace(Grid1.TextMatrix(k, f_p), ",", ".")
114         Grid1.TextMatrix(k, f_a) = Replace(Grid1.TextMatrix(k, f_a), ",", ".")
            '            If Trim(Grid1.TextMatrix(k, f_f)) <> "5" And gVal(Grid1.TextMatrix(k, f_a)) <> 0 Then
            '                 EXEI_MIDENIKO_FPA = 0
            '            End If

116         If Len(Trim(Grid1.TextMatrix(k, f_k))) > 0 Then
118             If GGET_NVALUE("select count(*) from EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'") = 0 Then
120                 MsgBox "δεν υπάρχει στο αρχείο ο κωδικός " + Grid1.TextMatrix(k, f_k) + " στην " + str(k) + " σειρά"
                    Exit Sub
             
                End If
            End If
            
            'αν η σειρά εχει προιοντα και εχει ασχετο φπα να γυρνάει πίσω
122         If Len(Trim(Grid1.TextMatrix(k, f_k))) > 0 And (Val(Grid1.TextMatrix(k, f_f)) > 9 Or Val(Grid1.TextMatrix(k, f_f)) <= 0) Then
124             MsgBox "Λάθος ΦΠΑ στην " + str(k) + " σειρά"
126             Grid1.row = k: Grid1.Col = f_f
128             timtext3.Top = Grid1.Top + Grid1.CellTop
130             timtext3.Left = Grid1.Left + Grid1.CellLeft
132             timtext3.Width = Grid1.CellWidth - 15
134             timtext3.Height = Grid1.CellHeight - 25
136             timtext3.Visible = True
138             timtext3.Text = Grid1.TextMatrix(k, f_f)
140             timtext3.SetFocus

                Exit Sub

            End If
        
        Next
    
142     FindSynolo
144     If gVal(sfpa.Caption) = 0 And gVal(SynoloG.Caption) <> 0 Then
146         EXEI_MIDENIKO_FPA = 1
        End If
         
148     If EXEI_MIDENIKO_FPA = 1 Then
150         If isAkyrotiko Then
           
            Else
           
152             If fMonimo_MIDENIKO_FPA = 0 Then ' ΔΕΝ ΕΧΕΙ ΜΟΝΙΜΗ ΑΠΑΛΛΑΓΗ
154                 If Len(apallagesFPA.Text) = 0 Then
156                     MsgBox "δηλωστε λόγο απαλλαγής ΦΠΑ"
                        Exit Sub
                    End If
                End If
              
            End If
        End If
        
        '  apallagesFPA.Text = ""

158     If gVal(parakratisiT) > 0 Then
160         If gVal(Left(Combo2PARAKRATISI.Text, 1)) > 0 Then
                'ok
            Else
162             MsgBox "δηλωστε ειδος φόρου"
164             Combo2PARAKRATISI.SetFocus
                Exit Sub
            End If
166         fS_Parakratisis = Left(Combo2PARAKRATISI.Text, 1) + ";" + Left(ComboPARAKRATISI.Text, 2)
           
        End If
        ' AJIA2PARAKR
168     If gVal(AJIA2PARAKR.Text) > 0 Then
170         If gVal(Left(Combo2TyposParakr.Text, 1)) > 0 Then
                'ok
            Else
172             MsgBox "δηλωστε ειδος φόρου"
174             Combo2TyposParakr.SetFocus
                Exit Sub
            End If
176         If Len(fS_Parakratisis) > 0 Then
178             fS_Parakratisis = fS_Parakratisis + ";" + Left(Combo2TyposParakr.Text, 1) + ";" + Left(Combo2OnomaParak.Text, 2)
            Else
180             fS_Parakratisis = Left(Combo2TyposParakr.Text, 1) + ";" + Left(Combo2OnomaParak.Text, 2)
            End If
           
        End If

182     If F_PERIOR_POLHTES = 1 Then '
184         If Val(Left$(Combo8.Text, 2)) = 0 Then
186             MsgBox "Δεν συμπηρώσατε πωλητή"

                Exit Sub

            End If
        End If

188     If Not IsNull(PEL.Recordset("PLAFON")) Then
          
190         If PEL.Recordset("TYP") + Round(Val(DOT(SynoloG.Caption)), 2) > PEL.Recordset("PLAFON") Then
192             If Left(Combo4.Text, 2) = "ΠΙ" And PEL.Recordset("PLAFON") > 0 Then
                    Dim ANS6 As Integer
194                 ANS6 = MsgBox("Υπερβαση πλαφόν " + str(PEL.Recordset("TYP") + Round(Val(DOT(SynoloG.Caption)), 2) - PEL.Recordset("PLAFON")) + " Συνεχίζω Ν/Ο", vbYesNo)
196                 If ANS6 = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
     
        End If

198     If f_is_benzinadiko = 1 Then '
200         If Len(Trim(Text2(2))) < 8 Then
202             MsgBox "Δεν συμπηρώσατε ΑΦΜ"
                Exit Sub
         
            End If
         
204         If F_THERM = 1 Then ' ΓΙΑ ΤΗΝ ΑΠΟΔΕΙΞΗ ΤΗΣ ΘΕΡΜΑΝΣΗΣ ΔΕΝ ΘΕΛΕΙ ΑΡ.ΑΠΟΔΕΙΞΗΣ
                ' ΑΛΛΑ ΓΙΑ ΤΟ ΤΙΜΟΛΟΓΙΟ ΘΕΛΕΙ
                Dim MTIM3  As String
206             MTIM3 = Split(F_PAR_AYTOK, ",")(1)  ' 'L','T'
                
208             If InStr(MTIM3, Right(PARAS.Text, 1)) > 0 Then  'AN EINAI TIMOLOGIO
210                 If F_THER_TIM_XORISAPOD = 1 Then
                        ' ΟΚ ΔΕΝ ΘΕΛΕΙ ΑΠΟΔΕΙΞΗ
                    
                    Else
212                     If Val(ar_apod.Text) = 0 Then
214                         MsgBox "Δεν συμπηρώσατε αποδειξη"
                            Exit Sub
                        End If
                    End If
                End If
                
            Else
216             If Val(ar_apod.Text) = 0 Then
218                 MsgBox "Δεν συμπηρώσατε αποδειξη"
                    Exit Sub
                End If
            End If
         
220         parat.Text = Left(ar_apod.Text + Space(10), 10) + parat.Text
         
        End If

        '------------------------sketh  EKTYPVSH -------------------------------------------------------------
222     If F_PDF = 3 Then
            ' ΑΝΤΙΓΡΑΦΕΙ ΤΟ ΑΡΧΕΙΟ ΑΠΟ ΤΟ C:\PDF\TEMP STO C:\PDF
            ' ΜΕΤΟΝΟΜΑΖΟΝΤΑΣ ΤΑ ΜΕ ΤΟ ΟΝΟΜΑ ΠΟΥ ΒΡΙΣΚΕΙ ΣΤΟ C:\PDF\NAME
      
            ' δεν κάνει τίποτα παρά μόνο πετάει το όνομα sto c:\pdf\name που πρεπει να δοθεί
            ' στο αρχείo c:\pdf\pdftemp
            Dim M_PDF2 As String

            Dim M_PDF2PARAGG

            Dim N122 As Integer

224         N122 = 0
       
            On Error Resume Next

226         If F_PDFPARAGG = 1 Then
228             M_PDF2PARAGG = Dir("C:\PDFPARAGG\NAME\*.*") 'ΟΝΟΜΑ ΤΗΣ ΠΑΡΑΓΓΕΛΙΑΣ
            End If

230         M_PDF2 = Dir(F_PDFFOLDER + "\NAME\*.*")
232         pdf2 = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF")

234         Do While Len(pdf2) > 0

                On Error Resume Next

236             N122 = N122 + 1
238             FileCopy F_PDFFOLDER + "\PDFTEMP\" + pdf2, F_PDFFOLDER + "\" + M_PDF2 + Trim(str(N122)) + ".PDF"
240             MILSEC 1000

242             If F_PDFPARAGG = 1 Then
        
244                 If Len(M_PDF2PARAGG) > 0 Then
246                     FileCopy F_PDFFOLDER + "\PDFTEMP\" + pdf2, "C:\PDFPARAGG\" + M_PDF2PARAGG + ".PDF"
248                     MILSEC 1000
                    End If
                End If

250             pdf2 = ""
                'On Error Resume Next
252             pdf2 = Dir '("C:\PDF\PDFTEMP\*.PDF")
            Loop
  
254         Kill F_PDFFOLDER + "\PDFTEMP\*.PDF"
256         Kill F_PDFFOLDER + "\NAME\" + M_PDF2
       
258         If F_PDFPARAGG = 1 Then
260             Kill "C:\PDFPARAGG\NAME\" + M_PDF2PARAGG
            End If
    
        End If

262     If F_PDF = 1 Or F_PDF = 2 Then ' σβηνει τα αρχεια pdf που μείνανε
            'SETTINGS.INI
            '[PDF Printer]
            '  output=C:\PDF\PDFTEMP\P<guid>.pdf
            '  confirmoverwrite = YES
            '  appendifexists = NO
            '  rememberlastfilename = NO
            '  rememberlastfoldername = YES
            '  openfolder = NO
            '  showpdf = NO
            '  showsaveas = never
            '  showsettings = never
            '  suppresserrors = NO
            '  device = pdfwrite
            '  RUNONSUCCESS=C:\PDF\ACROWRAP.EXE /t  "%1"
            '  resx = 150
            '  resy = 150

            'pdf2 = ".."
264         pdf2 = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF")
   
266         Do While Len(pdf2) > 0

                On Error Resume Next
  
268             FileCopy F_PDFFOLDER + "\PDFTEMP\" + pdf2, F_PDFFOLDER + "\" + pdf2
270             MILSEC 1000

272             If F_PDF = 1 Then
274                 DUM = ShellExecute(0&, "Print", F_PDFFOLDER + "\" + pdf2, "", "", vbNormalFocus)
276                 MILSEC 1000
                End If
    
278             Kill F_PDFFOLDER + "\PDFTEMP\" + pdf2

280             pdf2 = Dir '("C:\PDF\PDFTEMP\*.PDF")
            Loop

        End If

        'βρισκει το είδος του παραστατικού
        'γιατί πολλές φορές το πουλάει σαν λιανική ενώ είναι χονδρική πώληση
282     find_eid_parastat

        'X = mEidos_Par(1 + PARAS.ListIndex)
        'Text1.Text = find_parastat(k, X, 0)

284     f_Akyr = False

        '  On Error GoTo MHNYMA
        Dim Mf_matim As String

        Dim ans0 As Integer

        ' bazo +1 gia na me αφηνει την τελευταία μέρα να γράφω
286     If DTPicker1.Value < gEnarjh Or DTPicker1.Value > gLhjh + 1 Then
288         MsgBox "Ημερομηνία εκτός ορίων χρήσης " + Format(gEnarjh, "DD/MM/YYYY") + " - " + Format(gLhjh, "DD/MM/YYYY")

            Exit Sub

        End If

290     If DTPicker1.Value > Now Then
292         ans0 = MsgBox("Προσοχή  ημερομηνία μεγαλυτερη της σημερινής.  Συνεχίζω;", vbYesNo)

294         If ans0 <> vbYes Then

                Exit Sub

            End If
        End If
        
        'ΟΤΑΝ ΚΑΝΩ ΔΙΟΡΘΩΣΗ Η ΗΜΕΡΟΜΗΝΙΑ ΤΟΥ  ΕΠΟΜΕΝΟΥ ΑΡΙΘΜΟΥ ΤΙΜΟΛΟΓΙΟΥ ΝΑ ΜΗΝ ΕΙΝΑΙ ΜΙΚΡΟΤΕΡΗ ΑΠΟ ΤΟ ΤΡΕΧΟΝ ΤΙΜΟΛΟΓΙΟ
        'ΔΗΛΑΔΗ ΔΙΟΡΘΩΝΩ ΤΟ Τ000100 ΚΑΙ ΤΟΥ ΒΑΖΩ ΗΜΕΡΟΜΗΝΙΑ 15/3/26 ΕΝΩ ΤΟ Τ000101 ΕΧΕΙ ΗΜΕΡΟΜΗΝΙΑ 12/3/26
        'ΗΜΕΡ/ΝΙΑ ΕΠΟΜΕΝΟΥ ΑΡΙΘΜΟΥ :   gget_DvalUE("select HME FROM TIM WHERE ATIM='"+LEFT(F_MATIM,1)+FORMAT(VAL( MID(F_MATIM,2,6) )+1,"000000")+"' ")
        If f_FORM_EKTY > 0 Then
           Dim nextDate As Date
           nextDate = GGET_DVALUE("select HME FROM TIM WHERE ATIM='" + Left(f_matim, 1) + Format(Val(mID(f_matim, 2, 6)) + 1, "000000") + "' ")
           If DTPicker1.Value > nextDate And Year(nextDate) > 1900 Then ' an kanei eof tote einai null kai exei year = 1899
              MsgBox " H HMEΡOMHNIA EINAI ΜΕΓΑΛΥΤΕΡΗ ΑΠΟ ΤΗΝ ΗΜΕΡΟΜΗΝΙΑ ΤΟΥ ΕΠΟΜΕΝΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ"
              Exit Sub
           End If
           ' προηγουμενο τιμολογιο να μην ειμαι πισω σε ημερομηνια του προηγουμενου τιμολογιου
             nextDate = GGET_DVALUE("select HME FROM TIM WHERE ATIM='" + Left(f_matim, 1) + Format(Val(mID(f_matim, 2, 6)) - 1, "000000") + "' ")
           If Val(mID(f_matim, 2, 6)) > 1 And DTPicker1.Value < nextDate And Year(nextDate) > 1900 Then ' an kanei eof tote einai null kai exei year = 1899
              MsgBox " H HMEΡOMHNIA EINAI ΜΙΚΡΟΤΕΡΗ ΑΠΟ ΤΗΝ ΗΜΕΡΟΜΗΝΙΑ ΤΟΥ ΠΡΟΗΓΟΥΜΕΝΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ"
              Exit Sub
           End If
           
           
           
           
           
           
        
        End If
        
        

296     If f_FORM_EKTY > 0 And DateDiff("d", Now, DTPicker1.Value) <= -1 Then
298         ans0 = MsgBox("Προσοχή  ημερομηνία μικρότερη της σημερινής.  Συνεχίζω;", vbYesNo)

300         If ans0 <> vbYes Then

                Exit Sub

            End If
        End If

302     f_paras = PARAS.Text

        Dim RR As New ADODB.Recordset, parast As String

        'ελεγχω μηπως υπάρχει μεγαλύτερο νουμερο απο αυτό
        'If f_FORM_EKTY > 0 And F_DIOR = 0 Then ' εκτυπωνεται
        '   PARAST = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
        '   RR.Open "select * FROM TIM WHERE ATIM>'" + PARAST + "' AND LEFT(ATIM,1)='" + Right(PARAS.Text, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
        '   If Not RR.EOF Then
        '      MsgBox "ΥΠΑΡΧΕΙ ΠΑΡΑΣΤΑΤΙΚΟ ME ΜΕΓΑΛΥΤΕΡΟ AΡΙΘΜΟ"
        '      RR.Close
        '      Exit Sub
        '   End If
        '   RR.Close
        '
        '
        'End If

        'ελεγχος για άδειο παραστατικο
        'x = Open_Data("select sum(poso) as ss from timedit")

        Dim DB As Database

        'If gConnect = "Access" Then
        '  Set db = OpenDatabase(gDir, False, False)
        'Else
        '  Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

304     S_POSOTHTA = 0

306     For k = 1 To Grid1.rows - 1
308         S_POSOTHTA = S_POSOTHTA + Abs(Val(Grid1.TextMatrix(k, f_p)))
        Next

310     If S_POSOTHTA = 0 Then  ' Val(Synolo.Caption)   ' 202329
312         MsgBox "Δεν περάστηκαν είδη"
314         POLAGOR.SetFocus

            Exit Sub

        End If

316     m_tim = GetCurrentTime()

318     If Label1(0).Caption = "" Then 'den ξερω γιατι αλλαξε
320         MsgBox "Δεν Διαλέξατε " + IIf(fbuff = "e", "Πελάτη", "Προμηθευτή")

            Exit Sub

        End If

322     If Val(Text1.Text) = 0 Then
324         MsgBox "Δεν δωσατε αριθμό παραστατικού "

            Exit Sub

        End If

326     If f_is_benzinadiko = 1 Then '
     
            'αν δεν ειναι συγκεντρωτικα δελτια τοτε να ελεγχει αν εβαλε αξια
328         If InStr(f_SDA, Right$(PARAS.Text, 1)) = 0 Then
330             If Val(SynoloG.Caption) <= 0 Then
                    ' MsgBox "λαθος αξια απόδειξης"
                    ' Exit Sub
                End If
            End If
         
        End If

        Dim m500 As Integer

        'F_IS_BYTIO
        
332     If F_IS_BYTIO = 2 And Val(SynoloG.Caption) >= 372 And gVal(parakratisiT.Text) = 0 Then
334         m500 = MsgBox("Προσοχή είναι πάνω από 300€. Δεν βάλατε Παρακτράτηση . Συνεχίζω; ", vbYesNo)

336         If m500 = vbNo Then

                Exit Sub

            End If
            
        Else
338         If Val(synolo.Caption) > 500 And Combo4.Text = Combo4.List(0) Then
                                    'm500 =¨
                 
               
                     MsgBox ("Προσοχή είναι πάνω από 500€. Αλλάξτε τροπο πληρωμης") 'Συνεχίζω; ", vbYesNo)

                                 'If m500 = vbNo Then

                    Exit Sub

                
            End If
     
        End If

344     If F_STADIO = 5 And APOTH1.Text = APOTH2.Text Then    '' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
346         MsgBox "Οι Αποθήκες φόρτωσης και αποστολής πρέπει να διαφέρουν "

            Exit Sub

        End If
    
        Dim ANS3 As Integer

348     If F_EPIBEB_APOT = 1 And fbuff = "r" Then
350         ANS3 = MsgBox("Η αποθήκη που θα μπουν τα εμπορεύματα ειναι η " + Chr(13) + APOTH1.Text + Chr(13) + " Είσατε σίγουροι;", vbYesNo)

352         If ANS3 = vbNo Then

                Exit Sub

            End If
        End If

        ' προσοχή τα δελτία και χωρίς αξία πρέπει να περνάνε
        ' If Val(Synolo.Caption) = 0 Then Exit Sub

354     Me.MousePointer = vbHourglass

356     metasxhmatismos.Enabled = True

358     f_dat = DTPicker1.Value

360     f_matim = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text

362     If F_TRANS = 1 Then
364         Gdb.BeginTrans
        End If

366     F_METASX_PAR = 0 '  As Integer  ' ΕΙΝΑΙ ΜΕΤΑΣΧΗΜΑΤΙΣΜΟΣ ΠΑΡΑΣΤΑΤΙΚΟΥ = 1

368     par.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + Left(f_matim, 1) + "';", Gdb, adOpenDynamic, adLockOptimistic

370     If IsNull(par!FORM_ektyp) Then f_FORM_EKTY = 0 Else f_FORM_EKTY = Val(par!FORM_ektyp)

        'καταργει την διορθωση
372     If metasx.Value = vbChecked Then F_DIOR = 0

374     If F_DIOR = 1 Then    ' ΔΙΟΡΘΩΣΗ

            Dim ANS As Integer

376         If isAkyrotiko Then
378             ANS = vbYes
            Else
380             ANS = MsgBox("Να αποθηκευθεί η διόρθωση;", vbYesNo)
            End If
           
382         If ANS = vbNo Then GoTo OLD145    'Exit Sub

384         ANS = vbNo

386         If IsNull(par!FORM_ektyp) Then f_FORM_EKTY = 0 Else f_FORM_EKTY = Val(par!FORM_ektyp)

            ' ΑΝ ΕΙΝΑΙ ΧΕΙΡΟΓΡΑΦΟ ΤΟ ΠΡΟΣ ΑΚΥΡΩΣΗ ΠΑΡΑΣΤΑΤΙΚΟ ΝΑ ΜΠΟΡΕΙ ΝΑ ΔΙΑΓΡΑΦΕΙ

            'If f_FORM_EKTY > 0 Then    '   fBuff = "e" Then
388         If isAkyrotiko Then
390             ANS = vbYes
            Else
392             ANS = vbNo
            End If
                    
            ' ANS = MsgBox("Να ΕΚΔΟΘΕΙ ΑΚΥΡΩΤΙΚΟ;", vbYesNo)
            ' End If

394         If ANS = vbYes Then

396             f_Akyr = True
398             parat.Text = "Σχ.Παρ. " + f_matim
400             f_paras = "ΑΚΥΡΩΤΙΚΟ ΣΤΟΙΧΕΙΟ  κ"

                'αν δουλεύει με terminal services και είναι το υποκατάστημα
                'τότε να κοβει ακυρωτικό από αλλη σειρά
                'If f_YPOK = 1 And Len(Dir("c:\BSEIRA", vbDirectory)) > 2 Then
                '   Mf_matim = "k" + Format(find_parastat(1, "k", 0), String(f_psifia_atim, "0"))
                'Else
402             Mf_matim = f_AkyrParas + Format(find_parastat(1, f_AkyrParas, 0), String(f_psifia_atim, "0"))
                ' End If

                ' ELEGXOS AN YPARXEI HDH TO AKYROTIKO
404             R.Open "select ATIM from TIM where ATIM='" + Mf_matim + "'", Gdb, adOpenDynamic, adLockOptimistic

406             If Not R.EOF Then
408                 MsgBox "Υπάρχει ήδη  TO ΑΚΥΡΩΤΙΚΟ ΜΕ  ίδιο αριθμό. " + Mf_matim
410                 MsgBox "ΔΙΑΓΡΑΨΤΕ ΤΟ ΠΑΛΙΟ ΑΚΥΡΩΤΙΚΟ ή ΑΛΛΑΞΤΕ ΤΗΝ ΑΡΙΘΜΗΣΗ ΤΟΥ ΑΚΥΡΩΤΙΚΟΥ ΑΠΟ ΤΟ ΠΑΡΑΜΕΤΡΟΙ ΠΑΡΑΣΤΑΤΙΚΩΝ "
412                 Me.MousePointer = vbNormal

                    Exit Sub

                End If

            Else
414             f_Akyr = False
            End If

416         If f_Akyr = True And F_akyr_Delete = 0 Then
                'den σβηνω το  παραστατικό που ακυρωνεται
                
418             Gdb.Execute "UPDATE TIM SET AKYROMENO=1  where ID_NUM=" + str(F_ID_NUM_DIOR) '
                '    left(ATIM," + str(Len(f_dior_mAtim)) + ")='" + f_dior_mAtim + "' and  HME >= '" + Format(f_dior_HME, "mm/dd/yyyy") + "' and  HME <= '" + Format(DateAdd("D", 1, f_dior_HME), "mm/dd/yyyy") + "';"

            Else

                'σβηνω το παλιό παραστατικό
420             sql = "delete  "
422             sql = sql + " from EGGTIM where left(ATIM," + str(Len(f_dior_mAtim)) + ")='" + f_dior_mAtim + "' and  HME >= '" + Format(f_dior_HME, "mm/dd/yyyy") + "' and  HME <= '" + Format(DateAdd("D", 1, f_dior_HME), "mm/dd/yyyy") + "';"
424             Gdb.Execute sql, L

426             sql = "delete  "
428             sql = sql + " from EGG where left(ATIM," + str(Len(f_dior_mAtim)) + ")='" + f_dior_mAtim + "' and  HME >= '" + Format(f_dior_HME, "mm/dd/yyyy") + "' and  HME <= '" + Format(DateAdd("D", 1, f_dior_HME), "mm/dd/yyyy") + "';"
430             Gdb.Execute sql, L

432             sql = "delete  "
434             sql = sql + " from TIM where left(ATIM," + str(Len(f_dior_mAtim)) + ")='" + f_dior_mAtim + "' and  HME >= '" + Format(f_dior_HME, "mm/dd/yyyy") + "' and  HME <= '" + Format(DateAdd("D", 1, f_dior_HME), "mm/dd/yyyy") + "';"
436             Gdb.Execute sql, L
              
438             If f_KENTRA_ADYNATISMATOS = 1 Then
440                 If PEL.Recordset("EIDOS") = "e" Then
442                     Gdb.Execute "UPDATE PEL SET NUM4=(CASE WHEN NUM4 IS NULL THEN 0 ELSE NUM4 END) + " - str(Round(Val(DOT(SynoloG.Caption)), 2)) + " WHERE  EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("kod") + "'"
                    End If
                End If

            End If

            ' SAN MATIM EXO TO AKYROTIKO KAI BAZO SIMERINI HMEROMHNIA
444         If f_Akyr = True Then

446             If F_akyr_Delete = 0 Then
448                 F_POS_APOU = GET_CVALUE("select POS_APOU FROM PARASTAT WHERE EIDOS='" + Left(Mf_matim, 1) + "'")
450                 F_AJIA_APOU = GET_CVALUE("select AJIA_APOU FROM PARASTAT WHERE EIDOS='" + Left(Mf_matim, 1) + "'")
452                 f_pel = GET_CVALUE("select PEL FROM PARASTAT WHERE EIDOS='" + Left(Mf_matim, 1) + "'")
                    
                Else
                    ' ΜΗΔΕΝΙΖΩ ΑΥΤΕΣ ΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ ΓΙΑΤΙ
                    'ΘΑ ΠΑΡΕΙ ΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ ΤΟΥ ΑΡΧΙΚΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ ΚΑΙ ΘΑ ΧΑΛΑΣΕΙ ΤΗΝ ΑΠΟΘΗΚΗ
454                 F_POS_APOU = "0"
456                 F_AJIA_APOU = "0"
458                 f_pel = "0"
                End If
                
460             f_matim = Mf_matim
462             Text1.Text = Val(mID$(Mf_matim, 2, f_psifia_atim))
464             f_dat = Now
                
            End If

        Else  'DIOR=0 ΔΕΝ ΕΙΝΑΙ ΔΙΟΡΘΩΣΗ

            'diadikasia egkrishs
            '----------------------------------------------------------------
466         If f_FORM_EKTY > 0 And Val(SynoloG) > 20 And Val(SynoloG) < 10000 And Combo4.Text = Combo4.List(1) And Len(Dir("c:\mercvb\egkr.txt")) > 1 Then    '<> METRHTOIS ή κάρτα

                Dim ccc As String, NNN, DATE1, Topos

468             ccc = Right(Format(GetCurrentTime(), "#######0"), 4)

470             If InputBox("Αριθμός Εγκρισης για " + ccc) = egkrisi(ccc) Then
                    ' ok
                Else

472                 If InputBox("Eπανάληψη:Αριθμός Εγκρισης για " + ccc) = egkrisi(ccc) Then
                        'ok
                    Else
474                     MsgBox "Λάθος αριθμός έγκρισης"

                        Exit Sub

                    End If
                End If

476             DATE1 = InputBox("Ημερομηνία Εξόφλησης " + Format(DateAdd("d", 30, Now), "dd/MM/YYYY"))
478             Topos = InputBox("Τρόπος Εξόφλησης ")

480             If Not IsDate(DATE1) Then
482                 DATE1 = DateAdd("d", 30, Now)
                Else
484                 DATE1 = DateAdd("h", Hour(Now), DATE1)

                End If

486             NNN = add_schedule(DATE1, Text2(1).Text, Topos)

                Dim SQLMEMO

488             If IsNull(PEL.Recordset("MEMO")) Then
490                 SQLMEMO = Format(DATE1, "DD/MM/YYYY") + " " + Topos + "' WHERE EIDOS='e' and KOD='" + PEL.Recordset("KOD")
                Else
492                 SQLMEMO = PEL.Recordset("MEMO") + Chr(13) + Format(DATE1, "DD/MM/YYYY") + " " + Topos + "' WHERE EIDOS='e' and KOD='" + PEL.Recordset("KOD")
                End If

494             Gdb.Execute "UPDATE PEL SET MEMO='" + SQLMEMO + "'"

                '      'ενημερωση με τον τελευταίο αριθμό
                '   ' γιατί μπορεί ενδιάμεσα να κόπηκε και αλλο παραστατικό
                '    If F_DIOR = 0 And f_FORM_EKTY > 0 Then  ' οχι ΔΙΟΡΘΩΣΗ
                '      'ΓΙΑ ΝΑ ΠΙΑΝΕΙ ΤΟΝΣΩΣΤΟ ΑΡΙΘΜΟ ΑΛΛΑ ΔΗΜΙΟΥΡΓΗΣΕ ΠΡΟΒΛΗΜΑ  5-10-2009 ΒΑΖΟΝΤΑΣ ΤΟΝ ΙΔΙΟ ΑΡΙΘΜΟ ΣΕ ΠΟΛΛΑ ΠΑΡΑΣΤΑΤΙΚΑ
                '      '30/7/2010
                '       Text1.Text = find_parastat(1, Right(PARAS.Text, 1), 0)
                '      f_matim = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
                '    End If
                '
                '    If Check_Dipla = 0 Then
                '        If F_TRANS = 1 Then
                '           Gdb.RollbackTrans
                '        End If
                '        Exit Sub
                '    End If

            End If  '---------------------------------------------------------------

            'ενημερωση με τον τελευταίο αριθμό
            ' γιατί μπορεί ενδιάμεσα να κόπηκε και αλλο παραστατικό
496         If F_DIOR = 0 And f_FORM_EKTY > 0 Then  ' οχι ΔΙΟΡΘΩΣΗ
                'ΓΙΑ ΝΑ ΠΙΑΝΕΙ ΤΟΝΣΩΣΤΟ ΑΡΙΘΜΟ ΑΛΛΑ ΔΗΜΙΟΥΡΓΗΣΕ ΠΡΟΒΛΗΜΑ  5-10-2009 ΒΑΖΟΝΤΑΣ ΤΟΝ ΙΔΙΟ ΑΡΙΘΜΟ ΣΕ ΠΟΛΛΑ ΠΑΡΑΣΤΑΤΙΚΑ
                '30/7/2010
498             Text1.Text = find_parastat(1, Right(PARAS.Text, 1), 0)
500             f_matim = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
            End If

502         If Check_Dipla = 0 Then
504             If F_TRANS = 1 Then
                    '  Gdb.RollbackTrans
                End If

506             Me.MousePointer = vbNormal

                Exit Sub

            End If

        End If

        ' Get default Workspace.
        ' Set wrkDefault = DBEngine.Workspaces(0)
        ' Start of outer transaction.

        'MEXRI TIS 21/6/2010 DEN HTAN REM OI PARAKATO 3 SEIRES
        'META TIS 21/6/2010  METAFERUHKAN OI SEIRES STHN ARXH
        'If F_TRANS = 1 Then
        '  Gdb.BeginTrans
        'End If

508     If Len(f_deltia(1)) > 0 Then

            Dim s As Integer

510         k = 1

512         Do While f_IDdeltia(k) > 0
514             Gdb.Execute "update TIM set ART='1',SXETIKO='" + f_matim + " " + Format(f_dat, "dd/mm/yyyy") + "' where ID_NUM=" + str(f_IDdeltia(k))
516             k = k + 1
            Loop

        End If

        Dim id_num As Long

518     id_num = 0
        'δημιουργία eggtim και ΤΙΜ
520     enhm_eggtim id_num, isAkyrotiko

522     If F_TRANS = -1 Then  'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
524         F_TRANS = 1    ' για να μπορει να ξαναπαίξει
526         Me.MousePointer = vbNormal

            Exit Sub

        End If

528     If f_Akyr = False Or (f_Akyr = True And F_akyr_Delete = 0) Then  ' αν δεν ειναι ακυρωτικό Ή ( ειναι ακυρωτικό αλλά δεν διαγραφει το παλιό)

            'enhm_eggtim

            'ενημέρωση αποθήκης

530         enhm_apot

532         If F_TRANS = -1 Then  'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
534             F_TRANS = 1    ' για να μπορει να ξαναπαίξει
536             Me.MousePointer = vbNormal

                Exit Sub

            End If

            '
            '
            ' If f_pelMERC = "SOK" Then
            '     Dim M_AR_KARTAS, RK As New ADODB.Recordset
            '52     M_AR_KARTAS = InputBox("ΔΩΣΕ ΑΡ.ΔΩΡΟΚΑΡΤΑΣ")
            '     If Val(M_AR_KARTAS) > 0 Then
            '       RK.Open "SELECT * FROM EGG WHERE XRE<>0 AND LEFT(AIT,3)='ΔΩΡ' AND KOD='" + PEL.Recordset("KOD") + "' AND ATIM='L" + Format(M_AR_KARTAS, "000000") + "'", Gdb, adOpenDynamic, adLockOptimistic
            '       If RK.EOF Then
            '          MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ"
            '       Else
            '          MsgBox "ΠΟΣΟ ΔΩΡΟΚΑΡΤΑΣ " + Str(RK("XRE"))
            '          Gdb.Execute "UPDATE EGG SET AIT='ΕΞΑΡΓ." + Str(RK("XRE")) + "-" + Format(Now, "DD/MM") + "',PISTOSI=0,XREOSI=0,XRE=0 WHERE KOD='" + RK("KOD") + "' AND ATIM='" + RK("ATIM") + "' AND XRE<>0 AND LEFT(AIT,3)='ΔΩΡ'"
            '          F_AJIDOROKARTAS = 0  ' ΓΙΑ ΝΑ ΜΗΝ ΠΑΡΕΙ ΔΩΡΟΚΑΡΤΑ
            '       End If
            '
            '       RK.Close
            '
            '       GoTo 52
            '     End If
            '   End If

            'ενημέρωση πελάτη
          
538         f_suma = Round(Val(DOT(SynoloG.Caption)), 2)

            Dim m_suma As Single
          
            '2016-02-27 ΕΤΣΙ ΗΤΑΝ  If F_PARAKRATISI = 0 Then m_suma = f_suma Else m_suma = Round(Val(DOT(tel_pliroteo.Text)), 2)
540         m_suma = f_suma   ' ΕΤΣΙ ΕΓΙΝΕ
            '2016-02-27
             
542         enhm_pel m_suma - F_ELGA, id_num

544         If F_TRANS = -1 Then  'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
546             F_TRANS = 1    ' για να μπορει να ξαναπαίξει
548             Me.MousePointer = vbNormal

                Exit Sub

            End If

550         F_METRHTAposo = 0

552         If F_METRHTA = 1 Then
554             F_METRHTAposo = InputBox("ΜΕΤΡΗΤΑ")
                cmdPOS.Caption = "ΡΕΣΤΑ " + Format(F_METRHTAposo - m_suma, "###0.00")

                F_METRHTAposo = m_suma
            Else
556             F_METRHTAposo = m_suma
            End If

            '--------------   ενημέρωση μετρητων  ----------------------
            Dim Flag_Plir As Boolean

558         Flag_Plir = False

            '1680        If Val(Left(Combo4.Text, 2)) > 0 Then
            '1690            If F_LOGPLHR(Val(Left(Combo4.Text, 2)) - 1) <> "0000" Then    ' αν εχει ορισθει λογαριασμοσ πιστωτικων καρτων
            '1700                Flag_Plir = True
            '                 End If
            '             End If
                 
            Dim NGG As Integer

560         NGG = FIND_LISTINDEX(Combo4, Combo4.Text)

            'If Val(Left(Combo4.Text, 2)) > 0 Then
562         If Val(F_LOGPLHR(NGG)) > 0 Then '  <> "0000" Then    ' αν εχει ορισθει λογαριασμοσ πιστωτικων καρτων
564             Flag_Plir = True
            End If

            ' End If

            'για να πιστώσει πρέπει να είναι στην πρώτη εγγραφή ή να εχει pinakes.timh>0

566         If Combo4.Text = Combo4.List(0) Or Flag_Plir Then    ' METRHTOIS ή κάρτα
568             enhm_met F_METRHTAposo, id_num
            End If

570         If F_TRANS = -1 Then  'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
572             F_TRANS = 1    ' για να μπορει να ξαναπαίξει
574             Me.MousePointer = vbNormal

                Exit Sub

            End If

            Dim RTemp As New ADODB.Recordset

            Dim SQLT As String, TYP As Single

            Dim N22 As Long

            'ENHMEROSI PEL ME TYP
576         SQLT = "SELECT KOD,"
578         SQLT = SQLT + "SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END )  -"
580         SQLT = SQLT + " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END )   AS YP"
582         SQLT = SQLT + " FROM EGG  where HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("kod") + "' GROUP BY KOD"
584         RTemp.Open SQLT, Gdb, adOpenDynamic, adLockOptimistic

586         If IsNull(RTemp("YP")) Then
588             TYP = 0
            Else
590             TYP = RTemp("YP")
            End If

592         Gdb.Execute "UPDATE PEL SET TYP=AYP+ " + Replace(Format(TYP, "###########0.00"), ",", ".") + " WHERE  EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("kod") + "'", N22

594         If F_TRANS = -1 Then  'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
596             F_TRANS = 1    ' για να μπορει να ξαναπαίξει
598             Me.MousePointer = vbNormal

                Exit Sub

            End If

        End If  ' If f_AKYR = False Then  ' αν δεν ειναι ακυρωτικό

600     Me.Caption = str(Abs(m_tim - GetCurrentTime()))

        'ΕΝΗΜΕΡΩΣΗ ΑΡΙΘΜΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ

        '   par.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + left(f_matim, 1) + "';", Gdb, adOpenDynamic, adLockOptimistic

        ' If IsNull(par!FORM_EKTYP) Then f_FORM_EKTY = 0 Else f_FORM_EKTY = Val(par!FORM_EKTYP)

        '---------------γεφυρώσεις-------------------------------------
602     If Len(gdirlog) > 2 And f_error_gefyr = 0 Then
604         EIDOSKPE = PEL.Recordset("eidos") + PEL.Recordset("kod")
            'ΒΡΙΣΚΩ ΑΠΟ ΤΙΣ GEFYRES.DBF ΤΗΝ ΚΑΤΑΛΛΗΛΗ ΕΓΓΡΑΦΗ

606         f_mc = 0    ' ΣΕΙΡΕς GRID ΦΕΦΥΡΩΣΗΣ

            'η χρέωση
            'EIDOS     C  1  0N
            'TITLOS    C 30  0N
            'POL       C  1  0N
            'POS_APOU  C  1  0N
            'AJIA_APOU C  1  0N
            'PEL       C  1  0N
            'SYNT_TITL C  2  0N
            'EKTYP     C  1  0N
            'DESMIA    C  1  0N
            'FORM_EKTYPC  2  0N
            'STADIO    C  1  0N
            'ETIK      C  1  0N
            'METASX    C  1  0N
            'XONDR     C  1  0N
            'GEF_P     N  3  0N
            'GEF_M     N  3  0N
            'ARITMISI  N  5  0N
            'SHM1      C 60  0N

608         m_r = par("gef_p")
610         m_m = par("gef_m")

612         DUM = Gefyroma(f_dat, f_matim, EIDOSKPE, f_gfkod, f_gfaj, m_r, "Σ")
            ' τα μετρητά
614         m_r = m_m
616         DUM = Gefyroma(f_dat, f_matim, EIDOSKPE, f_gfkod, f_gfaj, m_r, "Ε")

618         DB.Close

        End If

        'ΜΑΡΚΆΡΩ ΤΑ ΔΕΛΤΙΑ ΠΟΥ ΧΡΗΣΙΜΟΠΟΙΗΣΑ
620     If Len(fSHMADEMENA) > 5 Then
622         Gdb.Execute "UPDATE TIM set ART='1' where ATIM+CONVERT(CHAR(10),hme,103)  in (" + fSHMADEMENA + ");"
        End If

        'If F_TRANS = 1 And (APOTH1.Text <> APOTH1.List(0)) Then

        ' Ask if the user wants to commit to all the changes
        ' made above.
        '  If MsgBox("Nα αποθηκευθεί το παραστατικό", vbYesNo) = vbYes Then

        '       On Error GoTo commitok

        '            Gdb.CommitTrans
        '  Else
        ' Gdb.RollbackTrans
        ' Me.MousePointer = vbNormal
        ' Exit Sub
        '  End If
        '    On Error Resume Next
        '
        'End If

        'If f_parking_auto = 1 Then
        '  Dim PSIF_ARKYKL As String: PSIF_ARKYKL = Text2(1).Text
        '  Dim MARKA As String: MARKA = Text2(3).Text
        '  MARKA = Left(MARKA, IIf(InStr(MARKA, ":") > 2, Len(MARKA) - InStr(MARKA, ":"), Len(MARKA)))
        '  MARKA = Text2(3).Text
        '
        '  Dim PS_L As Integer
        '  PS_L = InStr(MARKA, ":")
        '
        '
        '
        '  MARKA = Left(MARKA, IIf(PS_L > 2, Max(1, Len(MARKA) - 4), Len(MARKA)))
        '
        '
        '  psifiakopelat.NEWPARKING PSIF_ARKYKL, 2, MARKA, "αυτοκινητο", "", Text1.Text, SynoloG.Caption
        'End If

624     If F_TRANS = 1 Then
626         Gdb.CommitTrans
        End If

        '2020    If Left(f_pelMERC, 3) = "SOK" Then
        '
        '             Dim ARXDOR As String
        '
        '             Dim TELDOR As String
        '
        '             Dim POSODOR As Single
        '
        '2030        If UCase(Left(f_matim, 1)) = "G" Then
        '2040            ARXDOR = InputBox("απο ", "Αρχη εκπτωσης (δωροκάρτα)", Format(Now, "DD/MM/YYYY"))
        '2050            TELDOR = InputBox("απο ", "Αρχη εκπτωσης (δωροκάρτα)", Format(Now, "DD/MM/YYYY"))
        '2060            POSODOR = InputBox("%ΕΚΠΤΩΣΗ ΔΩΡΟΚΑΡΤΑΣ  ", "Ποσοστό εκπτωσης (δωροκάρτας)", 0)
        '                 'Gdb.Execute "update EID SET HM1='" + ARXDOR + "' WHERE KOD IN (SELECT ..."
        '             End If
        '
        '             Dim KSOK
        '
        '             'Kill "C:\KENO.TXT"
        '             ' MILSEC 1000
        '
        '             ' Open "c:\KENO.txt" For Append As #112
        '
        '             '    Print #112, "   "
        '             '   Print #112, "   "
        '             '  Print #112, "   "
        '
        '             '     For KSOK = 1 To 10
        '             '       Print #112, Grid1.TextMatrix(KSOK, f_k)  '   GRAFEI TON KODIKO GIA EPISTROFES
        '             '   Next
        '
        '             ' Close #112
        '
        '         End If

        Dim xxx As Printer

        'Dim foma1 As String
        
628     If isAkyrotiko Then  ' ΝΑ ΔΙΑΒΑΖΕΙ ΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ ΤΟΥ ΑΚΥΡΩΤΙΚΟΥ ΚΑΙ ΟΧΙ ΤΟΥ ΑΡΧΙΚΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ

            ' AN TO AKYΡΩΘΕΝ ΔEN EXEI ΦΟΡMA (XEIROGRAFO) TOTE NA PAREI THN FORMA TOY AKYROTIKOY
630         If f_FORM_EKTY = 0 Then

                ' πρεπει να βαλω το αντιστοιχο ακυρωτικό
                ' και οχι το 'κ'
632             F_FORMA1 = GET_CVALUE("select FORMA1 FROM PARASTAT WHERE EIDOS='" + Left(Mf_matim, 1) + "'")
634             F_FORMA2 = GET_CVALUE("select FORMA1 FROM PARASTAT WHERE EIDOS='" + Left(Mf_matim, 1) + "'")
636             F_FORMA3 = GET_CVALUE("select FORMA1 FROM PARASTAT WHERE EIDOS='" + Left(Mf_matim, 1) + "'")
638             f_FORM_EKTY = GET_NVALUE("select FORMA1 FROM PARASTAT WHERE EIDOS='" + Left(Mf_matim, 1) + "'")
           
            End If
        
        End If

        'εκτύπωση
640     If ektyp.Value = 1 Then    ' ---------------------------------------------

642         If f_FORM_EKTY >= 90 And f_FORM_EKTY <= 99 Then     ' φορμα τύπου f90.txt dot matrix '====================================
644             DUM = printCrystal(f_matim, CDate(DTPicker1.Value))    'δημιουργει τα προσωρινά αρχεία

646             If DUM < -900 Then
648                 MsgBox "αδυνατη η εκτύπωση" + Chr(13) ' + "Διορθώστε το παραστατικό για να ξανατυπωθεί"

                    '2130                If F_TRANS = 1 Then
                    '2140                    Gdb.RollbackTrans
                    '                    End If
                    '
                    '                    Exit Sub

                Else

650                 DUM = print_timol(f_matim, CDate(DTPicker1.Value))
                
                End If
            ElseIf f_FORM_EKTY > 180 And f_FORM_EKTY <= 190 Then ' ελαιοτροβειο συνοδευτικο
               
                    DUM = printCrystal(f_matim, CDate(DTPicker1.Value))    'δημιουργει τα προσωρινά αρχεία
            
             ''''''''''''''''''''''''''ElseIf f_FORM_EKTY = 188 Then      '     λιανικη
                            CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol" + Format(f_FORM_EKTY, "000") + ".rpt"
                            Dim tmima As String
                            tmima = Format(f_FORM_EKTY - 180, "0")
                            If Combo4.Text = Combo4.List(0) Then
                               tamiaki2 tmima, SynoloG.Caption, "0", "0"
                            Else
                               tamiaki2 tmima, "0", SynoloG.Caption, "0"
                            End If
            
            
             
                            Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
            
            
            
            
            
            
            
            
            
            
            
652         ElseIf Len(Trim(F_FORMA1)) > 1 Then
654             DUM = printCrystal(f_matim, CDate(f_dat))

656             If DUM < -900 Then
                    '                    MsgBox "αδυνατη η εκτύπωση" + Chr(13) + "Διορθώστε το παραστατικό για να ξανατυπωθεί"
                    '
                    '                    If F_TRANS = 1 Then
                    '                        Gdb.RollbackTrans
                    '                    End If
                    '
                    '                    Exit Sub

                Else
658                 Frame1.Caption = CNull(F_FORMA1)
660                 CrystalReport1.ReportFileName = F_FORMA1
662                 Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                
664                 If F_PAROCHOS = 1 Then
666                     Dim fFILE As String: fFILE = Trim(f_matim) + ".PDF"  '"c:\mercvb\" +
                        Dim Size As Long
668                     Size = Export2ReportToPDF(F_FORMA1, Trim(f_matim), fFILE)
670                     Dim ubl As String: ubl = GGET_CVALUE("SELECT left(ISNULL(UBL,'000'),3) AS UBL2 FROM PARASTAT WHERE EIDOS='" & Left(f_matim, 1) & "'")
                        'If GGET_NVALUE("SELECT ISNULL(B2G,0) AS DIMOSIO FROM PARASTAT WHERE EIDOS='" & Left(f_matim, 1) & "'") = 1 Then
                        
672                     If ubl <> "000" Then
674                         If GGET_NVALUE("select MYDATAISDEMO from MEM") = 1 Then
676                             PAROX_URLADD = "beta-"
                            Else
678                             PAROX_URLADD = ""
                            End If
680                         uploadAzure id_num, fFILE, Size
                        End If
                        
                       
                        
                        
                    End If
                  

682                 If Len(F_FORMA2) > 1 Then
684                     CrystalReport1.ReportFileName = F_FORMA2
686                     Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                    End If

688                 If Len(F_FORMA3) > 1 Then
690                     CrystalReport1.ReportFileName = F_FORMA3
692                     Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                    End If
                   
694                 If F_PDF = 4 Then
                        ' ΑΝΤΙΓΡΑΦΕΙ ΤΟ ΑΡΧΕΙΟ ΑΠΟ ΤΟ C:\PDF\TEMP STO C:\PDF
                        ' ΜΕΤΟΝΟΜΑΖΟΝΤΑΣ ΤΑ ΜΕ ΤΟ ΟΝΟΜΑ ΠΟΥ ΒΡΙΣΚΕΙ ΣΤΟ C:\PRF\NAME
                        Dim N123 As Integer
                        ' δεν κάνει τίποτα παρά μόνο πετάει το όνομα sto c:\pdf\name που πρεπει να δοθεί
                        ' στο αρχείo c:\pdf\pdftemp
696                     MILSEC 2000
                        Dim Pdf22 As String
                    
698                     Pdf22 = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF")
700                     Do While Len(Pdf22) = 0
702                         MILSEC 1000
704                         N123 = N123 + 1
706                         If N123 > 10 Then
                                Exit Do
                            End If
708                         Pdf22 = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF")
                        Loop
                    
710                     N123 = 0
                        On Error Resume Next
                    
712                     Do While Len(Pdf22) > 0

                            On Error Resume Next

714                         N123 = N123 + 1
716                         FileCopy F_PDFFOLDER + "\PDFTEMP\" + Pdf22, F_PDFFOLDER + "\" + f_matim + Trim(str(N123)) + ".PDF"
718                         MILSEC 1000
720                         Kill F_PDFFOLDER + "\PDFTEMP\" + Pdf22
722                         Pdf22 = ""
724                         Pdf22 = Dir '("C:\PDF\PDFTEMP\*.PDF")
                        Loop
  
726                     Kill F_PDFFOLDER + "\PDFTEMP\*.PDF"
    
                    End If
                
                End If
             
728         ElseIf f_FORM_EKTY > 200 And f_FORM_EKTY <= 299 Then  ' φορολογικος τυπου β
             
                '                DUM = printCrystal(f_matim, CDate(f_dat))    'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
                '
                '                If DUM < -900 Then
                '                    MsgBox "αδυνατη η εκτύπωση" + Chr(13) + "Διορθώστε το παραστατικό για να ξανατυπωθεί"
                '
                '                    If F_TRANS = 1 Then
                '                        Gdb.RollbackTrans
                '                    End If
                '
                '                    Exit Sub
                '
                '                End If
                '
                '                ' f_FORM_EKTY = 90
                '                ' DUM = print_timol(f_matim, CDate(DTPicker1.Value))
                '
                '                CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"
                '                Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                '                ' ΔΗΜΙΟΥΡΓΕΙ ΤΟ ΑΡΧΕΙΟ C:\FOROL\TIM.TXT   //GENERIC PRINTER ME ΕΞΟΔΟ C:\FOROL\TIM.TXT
                '
                '                ' ΚΑΡΑΦΥΛΛΗΣ ΤΥΠΟΥ Β
                '                Dim ExResult As String
                '
                '                Dim Result As Integer
                '
                '                Dim UseDos As Boolean
                '
                '                Dim SupError As Boolean
                '
                '                Dim o As New AlgoDll_webUpdate.Box
                '
                '                '-------------------------------------------------------------------------------------------------------------------
                '                'Using Data File From Windows Application
                '                UseDos = False
                '                SupError = False
                '                o.createTaxFiles = True
                '                o.sendTaxFiles = True
                '                Call o.Sign_("com1", "\FOROL\TIM.TXT", Result, ExResult, UseDos, "", SupError)
                '
                '                Dim DB5 As Database
                '
                '                Set DB5 = OpenDatabase("c:\mercvb\reports\reports.mdb")
                '
                '                If Result = 0 Then
                '
                '                    DB5.Execute "UPDATE MEM SET KEIMTIMOL='" + ExResult + "'"
                '
                '                    CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
                '                    Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                '                    MILSEC 15000
                '                    Call o.Sign_("com1", "\FOROL\TIM.TXT", Result, ExResult, UseDos, "", SupError)
                '
                '                    If Result = 0 Then
                '                        DB5.Execute "UPDATE MEM SET KEIMTIMOL='" + ExResult + "'"
                '                        CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
                '                        Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                '                    End If
                '
                '                Else
                '
                '                End If
                '
                '                DB5.Execute "UPDATE MEM SET KEIMTIMOL=''"
                '                Kill "c:\FOROL\TIM.TXT"
                  
730         ElseIf f_FORM_EKTY > 100 And f_FORM_EKTY <= 199 Then
                '                DUM = printCrystal(f_matim, CDate(f_dat))    'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
                '
                '                If DUM < -900 Then
                '                    MsgBox "αδυνατη η εκτύπωση" + Chr(13) + "Διορθώστε το παραστατικό για να ξανατυπωθεί"
                '
                '                    If F_TRANS = 1 Then
                '                        Gdb.RollbackTrans
                '                    End If
                '
                '                    Exit Sub
                '
                '                End If
                '
                '                For Each xxx In Printers
                '
                '                    If InStr(UCase(xxx.DeviceName), "DESKJET") > 0 Then
                '                        ' Set printer as system default.
                '                        Set Printer = xxx
                '
                '                        ' Stop looking for a printer.
                '                        Exit For
                '
                '                    End If
                '
                '                Next
                '
                '                CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol" + Format(f_FORM_EKTY, "000") + ".rpt"
                '                Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                  
            Else    ' crystal reports timol1.rpt(σφραγίδα)   timol1an.rpt (αντίγραφο)

732             If f_FORM_EKTY > 0 Then    '      PEL.Recordset("eidos") = "e" Then  ' EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE

                    'ΑΝ ΧΡΗΣΙΜΟΠΟΙΕΙ ΤΗΝ ΦΟΡΜΑ ΤΙΜΟΛ2 TOTE NA DINEI ALLES SEIRES APO TO F_FORM2_SEIRESTIMOL
734                 If f_FORM_EKTY = 14 Or f_FORM_EKTY = 12 Then
736                     f_SeiresTimologioy = f_Timol2_SeiresTimologioy
738                     f_SynSeiresTimologioy = f_Timol2_SeiresTimologioy
                    End If

                    ' On Error GoTo b_Katax_Click_Err ' na sbhstei 15-6-2006

740                 If (f_FORM_EKTY >= 51 And f_FORM_EKTY <= 59) Or f_FORM_EKTY = 15 Then                '  ME SQL SERVER H EKTYPOSI

742                     DUM = PrintSqlCrystal(f_matim, CDate(f_dat), id_num)  'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
                    Else  ' ME MDB
744                     DUM = printCrystal(f_matim, CDate(f_dat))    'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
                    End If

                    '   On Error Resume Next

746                 If DUM < -900 Then
748                     MsgBox "αδυνατη η εκτύπωση" + Chr(13) '+ "Διορθώστε το παραστατικό για να ξανατυπωθεί"

                        '2240                    If F_TRANS = 1 Then
                        '2250                        Gdb.RollbackTrans
                        '                        End If
                        '
                        '                        Exit Sub

                    Else

                        'Dim xxx As Printer

750                     For Each xxx In Printers

752                         If InStr(UCase(xxx.DeviceName), "DESKJET") > 0 Then
                                ' Set printer as system default.
754                             Set Printer = xxx

                                ' Stop looking for a printer.
                                Exit For

                            End If

                        Next

756                     If f_FORM_EKTY = 12 Then     '     λιανικη
758                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"
 
760                     ElseIf f_FORM_EKTY = 13 Or f_FORM_EKTY = 32 Or f_FORM_EKTY = 33 Then     '     λιανικη
762                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol3.rpt"


'                        ElseIf f_FORM_EKTY = 188 Then      '     λιανικη
'                            CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol88.rpt"
'                            If Combo4.Text = Combo4.List(0) Then
'                               tamiaki2 SynoloG.Caption, "0", "0"
'                            Else
'                               tamiaki2 "0", SynoloG.Caption, "0"
'                            End If
                               
                            



764                     ElseIf f_FORM_EKTY = 14 Then     '    1 ENTYPO
766                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"

768                     ElseIf f_FORM_EKTY = 31 Then     '    1 ENTYPO
770                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol4.rpt"

772                     ElseIf f_FORM_EKTY = 15 Then     '    1 ENTYPO
774                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol15.rpt"
776                         CrystalReport1.Connect = gConnect
                    
778                     ElseIf (f_FORM_EKTY >= 51 And f_FORM_EKTY <= 59) Then     '    1 ENTYPO
                       
                            ' CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol51.rpt"
780                         CrystalReport1.Connect = gConnect
                            ' Ucr9print.printingcr9  F_EKTYPOTHS, "C:\Mercvb\reports\timol51.rpt"
                        
                            ' ElseIf f_FORM_EKTY = 52 Then     '    1 ENTYPO
                            '    CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol52.rpt"
                            '   CrystalReport1.Connect = gConnect

782                     ElseIf f_FORM_EKTY = 25 Then     '     λιανικη EXOT
784                         Ucr9print.printingCR9 F_EKTYPOTHS, "C:\Mercvb\reports\timolsql.rpt"
                        
786                     ElseIf f_FORM_EKTY = 26 Then     '     λιανικη ESOT
788                         Ucr9print.printingCR9 F_EKTYPOTHS, "C:\Mercvb\reports\timolsql1.rpt"
                        
790                     ElseIf f_FORM_EKTY = 28 Then     '     prosfora
792                         Ucr9print.printingCR9 F_EKTYPOTHS, "C:\Mercvb\reports\timolsql3.rpt"
                        
794                     ElseIf f_FORM_EKTY = 27 Then     '     τιμολ εξοτ ελληνικα-φορολογικος + τιμ.εξοτ.αγγλικα
796                         Ucr9print.printingCR9 F_EKTYPOTHS, "C:\Mercvb\reports\timolsql1.rpt"
798                         Ucr9print.printingCR9 F_EKTYPOTHS, "C:\Mercvb\reports\timolsql.rpt"
                        
                            'CrystalReport1.ReportFileName = "C:\Mercvb\reports\timolsql.rpt"
                            'CrystalReport1.Connect = gConnect

                        Else
800                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
                        End If

                        Dim nnf As Integer

802                     If proepiskopisi.Value = vbChecked Then
                            ' nnf = crptToWindow
804                         CrystalReport1.Destination = crptToWindow
                        Else

806                         CrystalReport1.Destination = crptToPrinter
                        End If

808                     If Len(Dir("C:\MERCVB\REMOTEPRINT.TXT")) = 0 Then
                            ' ΤΥΠΩΝΩ ΚΑΝΟΝΙΚΑ

810                         If f_FORM_EKTY = 10 Then
                                ' CrystalReport1.Action = 1
812                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
814                         ElseIf f_FORM_EKTY = 14 Then
                                ' CrystalReport1.Action = 1
816                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

818                         ElseIf f_FORM_EKTY = 31 Then   ' PAKETOLOGIO TIMOL4
820                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
822                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                                ' CrystalReport1.Action = 1
                                ' CrystalReport1.Action = 1

824                         ElseIf f_FORM_EKTY = 15 Or f_FORM_EKTY = 51 Or f_FORM_EKTY = 52 Then
                                ' CrystalReport1.Action = 1
826                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
828                         ElseIf f_FORM_EKTY = 20 Then
830                             CrystalReport1.Destination = crptToWindow
832                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
834                         ElseIf f_FORM_EKTY = 18 Then
836                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
838                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
840                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
842                         ElseIf f_FORM_EKTY = 13 Then
844                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

846                         ElseIf f_FORM_EKTY = 32 Then
848                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
850                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

852                         ElseIf f_FORM_EKTY = 33 Then
854                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
856                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
858                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

860                         ElseIf f_FORM_EKTY = 23 Then
862                             CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
864                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

866                             CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"
868                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

870                         ElseIf f_FORM_EKTY = 24 Then
872                             CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol.rpt"
874                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

876                             CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol24.rpt"
878                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

880                         ElseIf f_FORM_EKTY = 25 Then
882                             CrystalReport1.ReportFileName = "C:\Mercvb\reports\timo3.rpt"
884                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

886                             CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol25.rpt"
888                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

                            Else
890                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
892                             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                            End If

                        Else
                                        
894                         Open "C:\MERCVB\REPORTS\1.TXT" For Output As #1
896                         Print #1, "-------------"
898                         Close #1
                        End If

                        '------------------------sketh  EKTYPVSH -------------------------------------------------------------
900                     If F_PDF = 3 Then
                            ' δεν κάνει τίποτα παρά μόνο πετάει το όνομα sto c:\pdf\name που πρεπει να δοθεί
                            ' στο αρχείo c:\pdf\pdftemp
902                         Open F_PDFFOLDER + "\name\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" For Output As #5
904                         Print " .."
906                         Close #5
 
908                         If F_PDFPARAGG = 1 Then
910                             k = InStr(parat.Text, " ")

912                             If k < 3 Then
914                                 k = 3 ' ΑΝ ΕΙΝΑΙ ΚΕΝΟ ΝΑ ΜΗΝ ΠΕΤΑΕΙ ΛΑΘΟΣ
                                End If

916                             If Val(mID(parat.Text, 2, k - 2)) > 0 Then
918                                 Open "C:\PDFPARAGG\name\" + Format(Val(mID(parat.Text, 2, k - 2)), "000000") For Output As #5
920                                 Print " .."
922                                 Close #5
                                End If
                            End If  'F_PDFPARAGG = 1 Then

                        End If  'F_PDF=3 Then

                        '------------------ΣΥΓΚΕΝΤΡΩΝΕΙ 2 ΣΕΛΙΔΕΣ ΚΑΙ ΜΕΤΑ ΤΥΠΩΝΕΙ -------------------------------------------------------------------
                        Dim x33, X34

                        Dim nn3

                        Dim k35 As Long

                        Dim PDF1(10) As String, K36 As Long
                
924                     If F_PDF = 2 Then
926                         nn3 = GetCurrentTime()
928                         PDF1(1) = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF", vbNormal)
                     
930                         Do While Len(PDF1(1)) = 0
932                             MILSEC 300

934                             If GetCurrentTime() - nn3 > 15000 Then

                                    Exit Do

                                End If

936                             PDF1(1) = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF", vbNormal)
                            Loop
                     
938                         MILSEC 1000
                     
940                         FileCopy F_PDFFOLDER + "\PDFTEMP\" + PDF1(1), F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(1)) + ".pdf"
942                         MILSEC 1000
                     
944                         nn3 = GetCurrentTime()

                            'ΣΙΓΟΥΡΕΥΟΜΑΙ ΟΤΙ ΟΝΤΩΣ ΑΝΤΙΓΡΑΦΗΚΕ
946                         Do While Len(Dir(F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(1)) + ".PDF")) = 0
948                             MILSEC 1000  '1000  18-9
950                             FileCopy F_PDFFOLDER + "\PDFTEMP\" + PDF1(1), F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(1)) + ".pdf"

952                             If GetCurrentTime() - nn3 > 15000 Then

                                    Exit Do

                                End If

                            Loop

954                         MILSEC 1000
956                         Kill F_PDFFOLDER + "\PDFTEMP\" + PDF1(1)
                            '--------------------------------------------------
958                         MILSEC 2000
                     
960                         PDF1(2) = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF", vbNormal)

962                         If PDF1(2) = PDF1(1) Then
964                             Kill F_PDFFOLDER + "\PDFTEMP\" + PDF1(1)
966                             MILSEC 2000
968                             PDF1(2) = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF", vbNormal)
970                             MILSEC 1000
                            End If
                    
972                         nn3 = GetCurrentTime()

974                         Do While Len(PDF1(2)) = 0
976                             MILSEC 300

978                             If GetCurrentTime() - nn3 > 15000 Then

                                    Exit Do

                                End If

980                             PDF1(2) = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF", vbNormal)
                            Loop
                     
982                         FileCopy F_PDFFOLDER + "\PDFTEMP\" + PDF1(2), F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(2)) + ".pdf"
984                         MILSEC 1000
                     
986                         nn3 = GetCurrentTime()

                            'ΣΙΓΟΥΡΕΥΟΜΑΙ ΟΤΙ ΟΝΤΩΣ ΑΝΤΙΓΡΑΦΗΚΕ
988                         Do While Len(Dir(F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(2)) + ".PDF")) = 0
990                             MILSEC 1000  '1000  18-9
992                             FileCopy F_PDFFOLDER + "\PDFTEMP\" + PDF1(2), F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(2)) + ".pdf"

994                             If GetCurrentTime() - nn3 > 15000 Then

                                    Exit Do

                                End If

                            Loop

996                         Kill F_PDFFOLDER + "\PDFTEMP\" + PDF1(2)
                        End If

                        '------------------ΣΥΓΚΕΝΤΡΩΝΕΙ 2 ΣΕΛΙΔΕΣ ΚΑΙ ΜΕΤΑ ΤΥΠΩΝΕΙ -------------------------------------------------------------------

                        '------------------------MEMONOMENH EKTYPVSH -------------------------------------------------------------
998                     If F_PDF = 1 Then
                  
                            ' πρωτο αντιγραφο
1000                        nn3 = GetCurrentTime()
1002                        PDF1(1) = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF", vbNormal)

1004                        Do While Len(PDF1(1)) = 0
1006                            MILSEC 300

1008                            If GetCurrentTime() - nn3 > 15000 Then

                                    Exit Do

                                End If

1010                            PDF1(1) = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF", vbNormal)
                            Loop

                            ' x33 = ShellExecute(0&, "Print", F_PDFFOLDER + "\PDFTEMP\" + PDF1(1), "", "", vbNormalFocus)
1012                        nn3 = GetCurrentTime()
1014                        MILSEC 1000 '1000  18-9

1016                        FileCopy F_PDFFOLDER + "\PDFTEMP\" + PDF1(1), F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(1)) + ".pdf"
1018                        MILSEC 1000 '18-9
                     
                            'ΣΙΓΟΥΡΕΥΟΜΑΙ ΟΤΙ ΟΝΤΩΣ ΑΝΤΙΓΡΑΦΗΚΕ
1020                        Do While Len(Dir(F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(1)) + ".PDF")) = 0
1022                            MILSEC 1000  '1000  18-9
1024                            FileCopy F_PDFFOLDER + "\PDFTEMP\" + PDF1(1), F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(1)) + ".pdf"

1026                            If GetCurrentTime() - nn3 > 15000 Then

                                    Exit Do

                                End If

                            Loop

1028                        MILSEC 1000 '18-9  1000
1030                        x33 = ShellExecute(0&, "Print", F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(1)) + ".pdf", "", "", vbNormalFocus)
                     
1032                        Kill F_PDFFOLDER + "\PDFTEMP\" + PDF1(1)
                     
                            '2ο και τα επόμενα
1034                        MILSEC 2000 '18-9   1000
                            'ΠΕΡΙΜΕΝΩ ΝΑ ΒΓΕΙ ΤΟ 2ο ΑΡΧΕΙΟ  (ΛΟΓΙΣΤΗΡΙΟ)
1036                        nn3 = GetCurrentTime()

1038                        Do While True ' Len(σ) = 0
1040                            pdf2 = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF", vbNormal)

1042                            If Len(pdf2) > 0 Then Exit Do

1044                            DoEvents

1046                            If GetCurrentTime() - nn3 > 50000 Then

                                    Exit Do

                                End If

                            Loop

1048                        nn3 = 0

1050                        k35 = 1

                            ' pdf2 = Dir("C:\PDF\PDFTEMP\*.PDF")
1052                        Do While Len(pdf2) > 0 And k35 < 50 ' K35<50 ΓΙΑ ΝΑ ΜΗΝ  ΦΑΕΙ ΚΟΛΗΜΑ
1054                            k35 = k35 + 1
1056                            PDF1(k35) = pdf2
1058                            MILSEC 1500  ' 1500 ok 'ayg 2012
1060                            pdf2 = Dir '("C:\PDF\PDFTEMP\*.PDF")
1062                            MILSEC 500
                            Loop

1064                        For K36 = 2 To k35
1066                            MILSEC 1000

1068                            If Len(PDF1(K36)) < 1 Then

                                    Exit For

                                End If

1070                            FileCopy F_PDFFOLDER + "\PDFTEMP\" + PDF1(K36), F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(K36)) + ".pdf"
                     
                                'ΣΙΓΟΥΡΕΥΟΜΑΙ ΟΤΙ ΟΝΤΩΣ ΑΝΤΙΓΡΑΦΗΚΕ
1072                            nn3 = GetCurrentTime()

1074                            Do While Len(Dir(F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(K36)) + ".PDF")) = 0
1076                                MILSEC 1000
1078                                FileCopy F_PDFFOLDER + "\PDFTEMP\" + PDF1(K36), F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(K36)) + ".pdf"

1080                                If GetCurrentTime() - nn3 > 70000 Then

                                        Exit Do

                                    End If

                                Loop
                     
1082                            MILSEC 1000
1084                            x33 = ShellExecute(0&, "Print", F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(str(K36)) + ".pdf", "", "", vbNormalFocus)
1086                            Kill F_PDFFOLDER + "\PDFTEMP\" + PDF1(K36)
                            Next

                        End If   ' PDF=1
                    End If
                 
                End If   ' If f_FORM_EKTY > 0 Then '  EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
                  
            End If  ' If ektyp.Value = 1 Then  '========================================

        End If  '-------------------------------------------------------------

        'ENHMERVSH ARIUMOY PARASTATIKOY
1088    If f_Akyr Then    ' ΑΝΗΜΕΡΩΣΗ ΑΚΥΡΩΤΙΚΟΥ
1090        DUM = find_parastat(1, Left(f_matim, 1), 1)
        Else

1092        If F_DIOR = 1 Then    ' ΔΙΟΡΘΩΣΗ   DEN ΕΝΗΜΕΡΩΝΕΙ ΑΡΙΘΜΗΤΗΡΑ
            Else
1094            Upd_Parastat
            End If
        End If

1096    If Left(f_pelMERC, 3) = "SOK" Then
1098        If f_FORM_EKTY = -1 Then    ' φο
1100            DUM = Shell("c:\TYP.BAT ", vbMinimizedNoFocus)

            End If
        End If
         
        If f_parking_auto = 1 Then
            Dim PSIF_ARKYKL As String: PSIF_ARKYKL = Text2(1).Text
            Dim MARKA As String: MARKA = Text2(3).Text
            MARKA = Left(MARKA, IIf(InStr(MARKA, ":") > 2, Len(MARKA) - InStr(MARKA, ":"), Len(MARKA)))
            MARKA = Text2(3).Text
  
            Dim PS_L As Integer
            PS_L = InStr(MARKA, ":")
  
            MARKA = Left(MARKA, IIf(PS_L > 2, Max(1, Len(MARKA) - 4), Len(MARKA)))
  
            psifiakopelat.NEWPARKING PSIF_ARKYKL, 2, MARKA, "αυτοκινητο", "", Text1.Text, SynoloG.Caption
        End If

        'ENHMERVSH ARIUMOY PARASTATIKOY
        'Upd_Parastat
        '    b_Katax.Enabled = False
        
        '2020-11-30  το εκανα rem γιατι είναι custom για καποιον πελάτη
        ' με περιφερόμενο η/υ αν θυμαμαι καλά
        ' φυσικη απογραφη
        '3630    If F_STADIO = 7 Then
        '
        '            On Error GoTo no_connection
        '
        '            Dim mXRE0 As Single
        '
        '            Dim remGDB As New ADODB.Connection
        '
        '            'remGDB.ConnectionString =
        '3640        remGDB.Open "dsn=MERCAPOT;uid=sa;pwd=p@ssw0rd"
        '
        '            Dim rREM As New ADODB.Recordset
        '
        '            Dim k0 As Long, k0n As Long
        '
        '3650        remGDB.BeginTrans
        '
        '            Dim mRem1 As Single, mRem0 As Single
        '
        '3660        For k0 = 1 To Grid1.ROWS - 1
        '
        '3670            If Len(Grid1.TextMatrix(k0, f_k)) > 0 Then ' AN EXEI KVDIKO
        '                    '3304     rREM.Open "select sum(XRE) as SX,sum(PIS) AS SP FROM EGGTIM WHERE KODE='" + Grid1.TextMatrix(k0, f_k) + "' AND APOT=" + Left(APOTH1.Text, 1), remGDB, adOpenDynamic, adLockOptimistic
        '
        '3680                rREM.Open "select sum(XRE) as SX,sum(PIS) AS SP FROM EGGTIM WHERE KODE='" + Grid1.TextMatrix(k0, f_k) + "' AND APOT=" + Left(APOTH1.Text, 1), remGDB, adOpenDynamic, adLockOptimistic
        '
        '3690                If IsNull(rREM(1)) Then mRem1 = 0 Else mRem1 = rREM(1)
        '3700                If IsNull(rREM(0)) Then mRem0 = 0 Else mRem0 = rREM(0)
        '
        '3710                mXRE0 = mRem1 - mRem0 + Grid1.TextMatrix(k0, f_p)
        '
        '3720                rREM.Close
        '
        '3730                sql = ""
        '3740                sql = "INSERT INTO EGGTIM (PROOD_AJ,HME,ATIM,MONA,TIMM,"
        '3750                sql = sql + "POSO,KAU_AJIA,MIK_AJIA,"
        '3760                sql = sql + "FPA,PELKOD,EIDOS,APOT,"
        '3770                sql = sql + "ONOMA,XRE,PIS,KODE"
        '3780                sql = sql + ") values ("
        '3790                sql = sql + str(k0) + ","    ' prood_aj
        '3800                sql = sql + "'" + Format(f_dat, "MM/DD/YYYY") + "',"  ' HME
        '3810                sql = sql + "'" + Left("S" + mID(f_matim, 2, 10), f_psifia_atim + 2) + "'," 'ATIM
        '3820                sql = sql + "'" + Grid1.TextMatrix(k0, F_M) + "',0,"  'mona,τιμμ
        '3830                sql = sql + str(mXRE0) + ",0,0,"    ' POSO,KAU_aj,MIK_ajIA
        '3840                sql = sql + str(Val(Grid1.TextMatrix(k0, f_f))) + ","  'FPA
        '3850                sql = sql + "'" + PEL.Recordset("kod") + "',"    'PELKOD
        '3860                sql = sql + "'" + PEL.Recordset("EIDOS") + "',"    'EIDOS
        '3870                sql = sql + Left(APOTH1.Text, 2) + ","  'APOT
        '3880                sql = sql + "'" + Left(Grid1.TextMatrix(k0, f_o), 25) + "',"   'ONOMA
        '3890                sql = sql + str(mXRE0) + ","  'XRE
        '3900                sql = sql + "0,"  'PIS
        '3910                sql = sql + "'" + Grid1.TextMatrix(k0, f_k) + "')"  'KODE
        '3920                remGDB.Execute sql, k0n
        '
        '                End If
        '
        '            Next
        '
        '3930        remGDB.CommitTrans
        '3940        MsgBox "OK ΕΝΗΜΕΡΩΘΗΚΕ ΤΟ ΚΕΝΤΡΙΚΟ"
        '
        '        End If

OLD145:
      
        '-----------------------------------------------------------------
        '--------------------  initialization ----------------------------
        '-----------------------------------------------------------------
  
        ' Exit Sub

1102    MIDENIZO_GRID
1104    metasx.Value = vbUnchecked
1106

1108    LASTTIMOL.Refresh


        Dim RR2() As String
   
1266    RR2 = Split(F_GRIDLISTASTIMOL, ";")
          
1268    If UBound(RR2) = 8 Then
            Dim LLC As Integer
1270        For LLC = 0 To 8
1272            TDBGrid1.columns(LLC).Width = RR2(LLC)
            Next
        Else
1274        TDBGrid1.columns(8).Width = 2500 'MARK
1276        TDBGrid1.columns(7).Width = 0 'DI_NUM
1278        TDBGrid1.columns(6).Width = 500 ' 500.882 'SXETIKO
1280        TDBGrid1.columns(5).Width = 1610  ' .142 'PARAT
1282        TDBGrid1.columns(4).Width = 480 ' 480.189 'KOD PELATH
1284        TDBGrid1.columns(3).Width = 3550 'EPO
1286        TDBGrid1.columns(2).Width = 1100 'AJI
1288        TDBGrid1.columns(1).Width = 1700 'ATIM
1290        TDBGrid1.columns(0).Width = 1049 'HME
       
        End If














1110    TelFpa.Text = 0
1112    AnastolhFPA.Value = False
        
1114    POLAGOR.SetFocus
        'PARAS.SetFocus

        'kostologhsh  aploma othonis
        'Gdb.Execute "DROP TABLE KOSTOL" + gXEIRISTHS
        'Gdb.Execute "select EID.KOD,EID.ONO,EID.XTI,EID.POS_KERD,EID.LTI,EID.LTI5,EGGTIM.TIMM  INTO KOSTOL" + gXEIRISTHS + " from EGGTIM inner join EID on EGGTIM.KODE=EID.KOD where left(ATIM,7)='" + f_matim + "'"
        'EID.RecordSource = "SELECT * FROM  KOSTOL" + gXEIRISTHS
        'EID.Refresh

1116    Me.MousePointer = vbNormal

1118    If Val(f_apeyueias) > 0 Then
1120        PARAS_LostFocus
1122        Text2(0) = f_apeyueias
1124        SendKeys "{ENTER} {ENTER}"
        End If

1126    If Len(Dir("C:\SITELAG.TXT")) > 0 And UCase(Left(f_matim, 1)) = "G" Then
            '  apot2vmn.EGGTIM2.ConnectionString = gConnect
            '  apot2vmn.EGGTIM2.RecordSource = "SELECT KODE from EGGTIM where left(ATIM,7)='" + Left(f_matim, 7) + "' and  HME >= '" + Format(f_dat, "mm/dd/yyyy") + "' and  HME <= '" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "';"
            '  apot2vmn.EGGTIM2.Refresh
      
            'apot2vmn.Caption = "SELECT KODE from EGGTIM where left(ATIM," + str(Len(f_dior_mAtim)) + ")='" + f_matim + "' and  HME >= '" + Format(f_dat, "mm/dd/yyyy") + "' and  HME <= '" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "';"
            ' apot2vmn.SHOW
        End If
  
        '                             GEN = "5221275%"
        '                             R0.Open "select max(KOD) FROM EID WHERE KOD LIKE '" + GEN + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        '                             XX = Format(1 + Val(Left(R0(0), 12)), "000000000000")
        '                             R0.Close
        '                             apot2scroll.KOD.Text = XX + ChDigEAN13(XX)    'timtext3.Text
        '                             On Error Resume Next
        '                             gApoMenu = False
        '                             apot2scroll.SHOW
        
1128    If f_AKYR_benzinadiko = 1 Then
1130        If InStr(f_AKYR_benzPAR, Left$(f_matim, 1)) > 0 Then
        
              If F_DIOR = 0 Then  ' στην διορθωση τιμολογιου να μην  βγαζει ξανα συνοδευτικο
        
                 If f_hotel = 1 Then
                     forosDiamonis id_num, f_matim
                 Else
1132                 akyr_benzinadiko id_num, f_matim
                 End If
                 
              End If

            End If
        End If
        
        
        F_DIOR = 0    ' ΔΙΟΡΘΩΣΗ
        
        
        
        Exit Sub

MHNYMA:
1134    HandleError "Par1:b_katax_click"

1136    Resume Next

commitok:
        'Gdb.RollbackTrans
1138    MsgBox "προσοχη δεν αποθηκεύθηκε το παραστατικό"
1140    Me.MousePointer = vbNormal

        Exit Sub

no_connection:
        '4160    MsgBox "το κεντρικο δεν ενημερωθηκε" + str(Erl)
        '
        '        On Error Resume Next
        '
        '4170    remGDB.RollbackTrans
1142    GoTo OLD145

        '<EhFooter>
        Exit Sub

SaveParastat_Err:
        ' 'MsgBox Err.Description & vbCrLf & _
           "in ADOMERCNEW.par1.b_Katax_Click " & _
           "at line " & Erl, _
           vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SaveParastat " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub UPDATE_POL_PARTIDES(ByVal NSEIRA, _
                        ByVal KODE As String, _
                        ByVal id_num As Long, _
                        ByVal mXre As Single, _
                        ByVal mPis As Single, _
                        ByVal EIDOSKPE As String, _
                        ByVal cHME As String, _
                        ByVal ATIM As String)
        '------------------------------------------------------------------------------
        Dim SYNOLO_POS As Single
        Dim TAKE       As Single
        Dim rpart      As New ADODB.Recordset, MTAKE As Single
        'Dim EIDOSKPE As String
        'Dim CHME As String: CHME = Format(DTPicker1.Value, "MM/DD/YYYY")
        'Dim Atim As String: Atim = Right(PARAS.Text, 1) + Format(Text1.Text, String$(6, "0"))
        'EIDOSKPE = PEL.Recordset("eidos") + PEL.Recordset("KOD")
 
100     If F_POS_APOU = 1 Or F_POS_APOU = 4 Then
102         If F_DIOR = 1 Then
104             rpart.Open "select *  FROM PARTIDES WHERE ID_NUM=" + str(F_ID_NUM_DIOR) + " AND KOD='" + KODE + "'  ", Gdb, adOpenDynamic, adLockOptimistic
            Else
106             rpart.Open "select *  FROM PARTIDES WHERE KOD='" + KODE + "' AND POSOPOL>0 ", Gdb, adOpenDynamic, adLockOptimistic
            End If
        Else
 
108         If F_DIOR = 1 Then
110             rpart.Open "select *  FROM PARTIDES WHERE ID_NUM=" + str(F_ID_NUM_DIOR) + " AND KOD='" + KODE + "'  ", Gdb, adOpenDynamic, adLockOptimistic
            Else
112             rpart.Open "select *  FROM PARTIDES WHERE KOD='" + KODE + "' AND POSOPOL>0 ", Gdb, adOpenDynamic, adLockOptimistic
            End If
            ' rpart.Open "select *  FROM PARTIDES WHERE KOD='" + KODE + "' AND POSOPOL>0 ", Gdb, adOpenDynamic, adLockOptimistic
        End If
        Dim mm_ID As Long
 
114     Do While Not rpart.EOF
 
116         mm_ID = rpart("id")
118         If F_POS_APOU = 1 Or F_POS_APOU = 4 Then ' ΠΡΟΣΤΙΘΕΝΤΑΙ ΟΙ ΕΙΣΑΓΩΓΕΣ  KAI OI EPISTROFES  'ΑΓΟΡΕΣ  'UCase(Left(Atim, 1)) = "G" Then
120             Gdb.Execute "UPDATE PARTIDES SET EIDOSKPE='" + EIDOSKPE + "',HME='" + cHME + "',ATIM='" + ATIM + "', POSOPOL=0,ID_NUM=" + str(id_num) + " WHERE ID=" + str(rpart("ID"))
            Else
    
122             If rpart!POSOPOL > 0 Then
                    'dhmioyrg;v to paidi
124                 Gdb.Execute "INSERT INTO PARTIDES(SEIRA,EIDOSKPE,KOD,LOTNUMBER,HMEL,YPOLOIPO,POSOPOL,POSOAG,ID_NUM,HME,ATIM) SELECT " + str(NSEIRA) + ", '" + EIDOSKPE + "',KOD,LOTNUMBER,HMEL,0,0," + str(-rpart!POSOPOL) + "," + str(id_num) + ",'" + cHME + "','" + ATIM + "' FROM PARTIDES WHERE ID=" + str(rpart("ID"))
                    ' ενημερωνω την μανα με το νέο υπόλοιπο
126                 Gdb.Execute "UPDATE PARTIDES SET YPOLOIPO=YPOLOIPO-POSOPOL,POSOPOL=0 WHERE ID=" + str(mm_ID) + ""
                End If
    
            End If
  
128         rpart.MoveNext
        Loop
         
130     rpart.Close

        ' αυτο πηγε μεσα στο enhm_apot sto telos
        ''ΠΡΕΠΕΙ ΝΑ ΔΩ ΑΝ ΣΒΗΣΤΗΚΑΝ ΚΑΠΟΙΟΙ ΚΩΔΙΚΟΙ ΠΟΥ ΚΟΥΒΑΛΟΥΣΑΝ ΠΑΡΤΙΔΕΣ
        '' ΚΑΙ ΝΑ ΣΒΗΣΩ ΤΙΣ ΠΑΡΤΙΔΕΣ ΠΟΥ ΚΟΥΒΑΛΟΥΣΑΝ
        'Dim DEL_PART As String, COUNT_DEL As Integer: DEL_PART = "": COUNT_DEL = 0
        ' If F_POS_APOU = 1 Or F_POS_APOU = 4 Then
        '    If F_DIOR = 1 Then
        '       RPART.Open "select *  FROM PARTIDES WHERE ID_NUM=" + str(F_ID_NUM_DIOR), Gdb, adOpenDynamic, adLockOptimistic
        '       Do While Not RPART.EOF
        '          DEL_PART = DEL_PART + CNull(RPART!LOTNUMBER) + Chr(13)
        '          COUNT_DEL = COUNT_DEL + 1
        '          RPART.MoveNext
        '       Loop
        '       If COUNT_DEL > 0 Then
        '          MsgBox "Προσοχη !!! θα διαγραφουν οι παρτιδες :" + Chr(13) + DEL_PART
        '          Gdb.Execute "delete FROM PARTIDES WHERE ID_NUM=" + str(F_ID_NUM_DIOR)
        '       End If
        '       RPART.Close
        '    End If
        ' End If

End Sub

Function egkrisi(ByVal ccc As String) As Long

        '<EhHeader>
        On Error GoTo egkrisi_Err

        '</EhHeader>

100     egkrisi = Val(mID(ccc, 1, 1)) + Val(mID(ccc, 2, 1)) + Val(mID(ccc, 3, 1)) + Val(mID(ccc, 4, 1))
102     egkrisi = Int(egkrisi / 2) + 135

        '<EhFooter>
        Exit Function

egkrisi_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.egkrisi " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.egkrisi " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Sub MIDENIZO_GRID()

        '===================================================================
        '<EhHeader>
        On Error GoTo MIDENIZO_GRID_Err

        '</EhHeader>
        Dim k As Long

        Dim L As Long

100     synolo.Caption = ""
102     sfpa.Caption = ""
104     SynoloG.Caption = ""

106     For k = 1 To 9: f_kau(k) = 0: SumFpa(k) = 0: Next    ' συνολο ΦΠΑ ΚΑΤΑ ΣΥΝΤΕΛΕΣΤΗ

108     For k = 1 To 100: f_fpa(k) = 0: f_sajia(k) = 0: Next    ' κρατά στην μνήμη τοσύνολο κάθε σειράς και το ΦΠΑ

110     For k = 1 To 100: f_IDdeltia(k) = 0: f_deltia(k) = "": Next    'δελτια που σημαδεύω
112     For k = 1 To 100: f_deltiaMarks(k) = 0: f_deltiaMarks(k) = "": Next    'marks-δελτιων που σημαδεύω
        ' μηδενιζω το grid1  f_deltiaMarks
114     For L = 1 To Grid1.ColS - 1

            'Grid1.Col = l
116         For k = 1 To Grid1.rows - 1
                'Grid1.Row = K
                'Grid1.Text = ""
118             Grid1.TextMatrix(k, L) = ""
            Next
        Next

120     Text2(0).Enabled = True
122     Text2(1).Enabled = True
124     Text2(2).Enabled = True

126     Text2(0).Text = ""    ' κωδ
128     If POLAGOR.ListIndex = 1 Then  ' ΣΤΙς ΠΩΛΗΣΕΙΣ ΤΟ ΑΛΛΑΖΩ

130         Text2(1).Text = ""    'επωνυμια

        End If

132     Text2(2).Text = ""    'αφμ
134     Text2(3).Text = ""    'ΔΙΕΥΘΥΝΣΗ

136     For L = 0 To 7
138         Textm(L).Text = ""
        Next
        
140     parat.Text = ""    'ΠΑΡΑΤΗΡΗΣΕΙΣ

        'On Error GoTo 0
142     timtext3.Top = 18000
144     APOTH2.Enabled = True
146     Text5.Text = "" 'ΓΕΝΙΚΗ ΕΚΠΤΩΣΗ
148     tel_pliroteo.Text = ""
150     parakratisiT.Text = ""
152     tel_pliroteo.Visible = False
154     parakratisiT.Visible = False
        
        ' ComboPARAKRATISI.Text = ""
        ' Combo2PARAKRATISI.Text = ""
156     ComboPARAKRATISI.Visible = False
158     Combo2PARAKRATISI.Visible = False

        '===================================================================
        '<EhFooter>
        Exit Sub

MIDENIZO_GRID_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.MIDENIZO_GRID " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.MIDENIZO_GRID " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Sub elegxoi()
''================================================================
''=======   DEMO TA  TRANSACTIONS  ===============================
''================================================================
''
''   Dim strName As String
''   Dim strMessage As String
''   Dim wrkDefault As Workspace
''   Dim dbsNorthwind As Database
''   Dim rstEmployees As Recordset
''
''   ' Get default Workspace.
''   Set wrkDefault = DBEngine.Workspaces(0)
''   Set dbsNorthwind = OpenDatabase("Northwind.mdb")
''   Set rstEmployees = _
'    '      dbsNorthwind.OpenRecordset("Employees")
''
''   ' Start of outer transaction.
''   wrkDefault.BeginTrans
''   ' Start of main transaction.
''   wrkDefault.BeginTrans
''
''   With rstEmployees
''
''      ' Loop through recordset and ask user if she wants to
''      ' change the title for a specified employee.
''      Do Until .EOF
''         If !title = "Sales Representative" Then
''            strName = !LastName & ", " & !FirstName
''            strMessage = "Employee: " & strName & vbCr & _
'             '               "Change title to Account Executive?"
''
''            ' Change the title for the specified employee.
''            If MsgBox(strMessage, vbYesNo) = vbYes Then
''               .Edit
''               !title = "Account Executive"
''               .Update
''            End If
''         End If
''
''         .MoveNext
''      Loop
''
''      ' Ask if the user wants to commit to all the changes
''      ' made above.
''      If MsgBox("Save all changes?", vbYesNo) = vbYes Then
''         wrkDefault.CommitTrans
''      Else
''         wrkDefault.Rollback
''      End If
''
''      ' Print current data in recordset.
''      .MoveFirst
''      Do While Not .EOF
''         Debug.Print !LastName & ", " & !FirstName & _
'          '            " - " & !title

''         .MoveNext
''      Loop
''
''      ' Roll back any changes made by the user since this is
''      ' a demonstration.
''      wrkDefault.Rollback
''      .Close
''   End With
''
''   dbsNorthwind.Close
''
'''End Sub
''
''
'
''1.να μην υπάρχουν ΦΠΑ>9 και <=0
'End Sub
Sub Upd_Parastat()

        '<EhHeader>
        On Error GoTo Upd_Parastat_Err

        '</EhHeader>

        Dim k   As Integer

        Dim DB  As Database

        Dim R   As Recordset, X As String

        Dim DUM As Integer

100     For k = 0 To SEIRA.ListCount

102         If SEIRA.Text = SEIRA.List(k) Then

                Exit For

            End If

        Next

104     If k <= SEIRA.ListCount Then
106         DUM = find_parastat(k, Right(f_paras, 1), 1)
        End If

        '<EhFooter>
        Exit Sub

Upd_Parastat_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Upd_Parastat " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Upd_Parastat " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function Check_Dipla()

        '************************************************************************
        ' αν υπάρχει ήδη ίδιο παραστατικό  με ίδιο αριθμό  γυρνάει πίσω
        ' στους πελάτες αν υπάρχει ό ίδιος αριθμός με παλιά ημερομηνία μας ενημερώνει
        '<EhHeader>
        On Error GoTo Check_Dipla_Err

        '</EhHeader>
        Dim DB  As Database, re As New ADODB.Recordset, rt As New ADODB.Recordset

        Dim ANS As Integer

100     Check_Dipla = 1

        'If gConnect = "Access" Then
        '  Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

102     If fbuff = "e" And f_FORM_EKTY > 0 Then  ' φορμα τύπου f90.txt Then

104         re.Open "select ATIM,HME from TIM where ATIM='" + Trim(f_matim) + "'", Gdb, adOpenDynamic, adLockOptimistic

106         If Not re.EOF Then
108             re.MoveFirst

110             Do While Not re.EOF

112                 If Format(re("hme"), "dd/mm/yyyy") = Format(f_dat, "dd/mm/yyyy") Then
114                     MsgBox "Υπάρχει ήδη  παραστατικό με ίδιο αριθμό και ημερομηνία "    '        Format(re("hme"), "dd/mm/yyyy")
116                     Check_Dipla = 0
                        
                        Exit Function

                    End If

118                 re.MoveNext
                Loop

120             re.MoveFirst
122             ANS = MsgBox("Να συνεχισθεί η καταχώριση;", vbYesNo, "Υπάρχει ήδη  παραστατικό με ίδιο αριθμό στις " + Format(re("hme"), "dd/mm/yyyy"))

124             If ANS = vbYes Then
126                 Check_Dipla = 1
                Else
128                 Check_Dipla = 0
                End If

                Exit Function

            End If

130         re.Close

            'ΕΛΕΓΧΟΣ ΑΝ ΥΠΑΡΧΕΙ ΤΟ ΠΡΟΗΓΟΥΜΕΝΟ ΠΑΡΑΣΤΑΤΙΚΟ

            Dim reco As String

132         reco = "select * from TIM where  Left(ATIM, 1)  IN (SELECT EIDOS FROM PARASTAT WHERE ARITMISI=(SELECT ARITMISI  FROM PARASTAT WHERE EIDOS='" + Left(f_matim, 1) + "' ))  "
134         reco = reco + " AND SUBSTRING(ATIM,2,6) = '" + Left(Format(Val(mID$(f_matim, 2, f_psifia_atim) - 1), String(f_psifia_atim, "0")) + "   ", 6) + "'" ' + mID$(f_matim, 8, 1) + "%'"
            'SELECT * FROM TIM WHERE ATIM LIKE 'T0005%'
136         re.Open reco, Gdb, adOpenForwardOnly, adLockReadOnly

138         If re.EOF Then   ' And InStr("Ae", Left(f_matim, 1)) = 0

140             If F_LOCKARITMISI = 1 And gVal(mID(f_matim, 2, 7)) - 1 > 0 Then  ' εαν ο προηγουμενο αριθμος ειναι 0
142                 MsgBox "Λάθος στην συνέχεια της αρίθμησης. Ελέγξτε την αρίθμηση"
                    ANS = MsgBox("ΘΑ ΔΗΜΙΟΥΡΓΗΘΕΙ ΚΕΝΟ ΣΤΗΝ ΑΡΙΘΜΗΣΗ" + Chr(13) + "ΝΑ ΣΥΝΕΧΙΣΤΕΙ Η (ΛΑΘΟΣ;) ΚΑΤΑΧΩΡΙΣΗ;", vbYesNo)
                    If ANS = vbYes Then
152                     Check_Dipla = 1
                    Else
154                     Check_Dipla = 0
                    End If
144                 'Check_Dipla = 0
                Else
146                 reco = "Δεν υπάρχει ο προηγούμενος αριθμός.ΘΑ ΔΗΜΙΟΥΡΓΗΘΕΙ ΚΕΝΟ ΣΤΗΝ ΑΡΙΘΜΗΣΗ." + Format(Val(mID$(f_matim, 2, f_psifia_atim)) - 1, String(f_psifia_atim, "0")) + " Να συνεχισθεί η (ΛΑΘΟΣ;) καταχώριση;"
'                      Dim LASTOK As String, trex_atim As String
'                      trex_atim = Right(PARAS.Text, 1)
'
'
'                      LASTOK = GGET_CVALUE("select TOP 1 ATIM FROM TIM WHERE LEFT(ATIM,1)='" + trex_atim + "' ORDER BY ATIM DESC")
'
'                       LASTOK = GGET_CVALUE("select TOP 1 ATIM FROM TIM WHERE LEFT(ATIM,1)='" + trex_atim + "' ORDER BY ATIM DESC")
'                      'Α.ΒΡΙΣΚΩ ΤΟ ΤΕΛΕΥΤΑΙΟ ΠΑΡ/ΚΟ ΑΥΤΗΣ ΤΗΣ ΣΕΙΡΑΣ ΚΑΙ 1.UPDATE ARITMISI   2.BAZV THN ΣΩΣΤΗ ΑΡΙΘΜΗΣΗ ΣΤΟ ΤΙΜΟΛΟΓΙΟ
'                     'Β.ΣΕ ΠΕΡΙΠΤΩΣΗ ΠΟΥ ΕΙΝΑΙ ΤΟ ΠΡΩΤΟ ΤΗΣ ΧΡΟΝΙΑΣ ΜΕ ΡΩΤΑ
'                    If LASTOK = "" Then ' B
148                           ANS = MsgBox(reco, vbYesNo)
                
150                          If ANS = vbYes Then
                                      Check_Dipla = 1
                               Else
                                      Check_Dipla = 0
                               End If
'                     Else ' A .PERIPTOSI
'                       Gdb.Execute "UPDATE ARITMISI SET ARITMISI=" + str(Val(mID(LASTOK, 2, 6))) + " WHERE ID=(SELECT ARITMISI FROM PARASTAT WHERE EIDOS='" + Left(f_matim, 1) + "') "
'
'                       'f_matim = Left(f_matim, 1) + Format(Val(mID(LASTOK, 2, 6)) + 1, "000000")
'                       Text1.Text = Format(Val(mID(LASTOK, 2, 6)) + 1, "#####0")
'
'                     End If
                       
                End If

                Exit Function

            Else    ' YPARXEI ALLA DEN PREPEI NA EXO MIKROTERI HMEROMHNIA APO TO PROHGOYMENO
156             If Not re.EOF Then
158                 If DTPicker1.Value < re("HME") And nNull(re("AKYROMENO")) <> 1 Then
160                     MsgBox "η ημερομηνια είναι μικρότερη από την ημερομηνία του  προηγούμενου παραστατικού "
162                     Check_Dipla = 0
                    End If
                End If
              
            End If

        Else    'ΠΡΟΜΗΘΕΥΤΕΣ
164         re.Open "select ATIM from TIM where HME='" + Format(f_dat, "mm/dd/yyyy") + "' and ATIM='" + f_matim + "'", Gdb, adOpenDynamic, adLockOptimistic

166         If re.RecordCount > 0 Then
168             MsgBox "Υπάρχει ήδη  παραστατικό με ίδια ημερομηνία και ίδιο αριθμό"
170             Check_Dipla = 0
            End If
        End If

        '<EhFooter>
        Exit Function

Check_Dipla_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Check_Dipla " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Check_Dipla " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Sub enhm_eggtim(ByRef m_ID_NUM As Long, ByVal isAkyrotiko)

        Dim SQLTIM As String

        Dim MLITRA As Single

        Dim MEFK   As Single

        Dim SQLpyr As String, rPyr As New ADODB.Recordset, kPyr As Integer, aPyr As String, nPyr As Integer, KatPyr As String, NEAKIN As String
       
        '=========================================
        'δημιουργία eggtim και ΤΙΜ
        '=========================================
        '<EhHeader>
        On Error GoTo enhm_eggtim_Err

        '</EhHeader>

        Dim k, L, m, N As Integer

        Dim MHKOS As Integer

        Dim mPoso, mKAU_AJIA, mMIK_AJIA, mProeleysh

        Dim mXre, mPis, mONOMA, mKERDOS

        Dim sql As String

        Dim mLenOnoma

        'Dim Fdb As New ADODB.Connection
        'Dim Re As New ADODB.Recordset
        'Dim Rt As New ADODB.Recordset

        Dim DB   As Database, re As New ADODB.Recordset, rt As New ADODB.Recordset

        Dim mEID As New ADODB.Recordset

        Dim FEGG(10)

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

        '  On Error GoTo MHNYMA

100     For k = 1 To 9
102         f_kau(k) = 0
        Next

        'DB.Open gDir
        '  End If
        
        ' ΠΩς ΒΡΙΣΚΩ ΤΟ ΜΗΚΟΣ ΤΟΥ ΠΕΔΙΟΥ C1
        'SELECT CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME  = 'TIM' AND COLUMN_NAME='C1'
104     re.Open "SELECT TOP 1 CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME  = 'EGGTIM' AND COLUMN_NAME='ONOMA'", Gdb, adOpenDynamic, adLockOptimistic
106     rt.Open "SELECT top 1 *FROM TIM", Gdb, adOpenDynamic, adLockOptimistic

        Dim mLenPROORISMOS, mLenAYTOKINHTO, mLenSKOPOS, mLenFORTOSH, mLenPARAT

108     mLenPROORISMOS = rt("proor").DefinedSize
110     mLenAYTOKINHTO = rt("aytok").DefinedSize
112     mLenSKOPOS = rt("skopos").DefinedSize

114     mLenFORTOSH = rt("fortosh").DefinedSize
116     mLenPARAT = rt("parat").DefinedSize
118     mLenOnoma = re(0) 're("onoma").DefinedSize

        Dim SYN_ME_FPA As Single

        ' re("MIK_AJIA") = 0

        ''------------------------------ ΕΙΣΑΓΩΓΗ EGGTIM --METAKOMISE 5/7/2014 APO EDO PARAKATO-------------------------
        '200  For k = 1 To 10: FEGG(k) = 0: Next
        '     '------------------------------eggtim -----------------------------------------
        '210  For k = 1 To Grid1.ROWS - 1
        '
        '220      If Len(Grid1.TextMatrix(k, f_k)) > 0 Then
        '
        '230          mEID.Open "SELECT * FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic    'adOpenForwardOnly, adLockReadOnly
        '             'ΠΡΟΣΘΕΤΕΙ ΜΙΑ ΕΓΓΡΑΦΗ ΓΙΑ ΚΑΘΕ ΧΡΩΜΑ ή ΜΕΓΕΘΟΣ
        '240          If F_EXO_XROMATA = 0 Then
        '250              GoSub ADD_RECORD
        '             Else
        '                 Dim K2 As Integer, k3 As Integer
        '                 'ΣΤΗΝ ΔΙΟΡΘΩΣΗ ΔΙΟΡΩΝΕΙ RECORD-RECORD
        '260              If F_DIOR = 1 Then
        '270                  GoSub ADD_RECORD
        '                 Else
        '280                  For K2 = 1 To 25: For k3 = 1 To 25
        '290                          If Val(XROMATA(k, K2, k3)) <> 0 And Val(XROMATA(k, K2, k3)) <> -0.01 Then
        '300                              GoSub ADD_RECORD
        '310                              XROMATA(k, K2, k3) = 0
        '                             End If
        '320                      Next: Next
        '                 End If
        '             End If
        '330          mEID.Close
        '         End If
        '     Next
     
        '==================================== TIM ==================================================
   
120     For k = 1 To Grid1.rows - 1

            '*************  5/7/2014  NEXT LINE *******************************
122        ' If F_XONDR = 0 Then ' lianikh
124        '     mMIK_AJIA = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias)
           ' Else
            
126             If Val(Grid1.TextMatrix(k, f_a)) = 0 Then
128                 mMIK_AJIA = 0
                Else
130                 mMIK_AJIA = Round(Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias) * (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100), f_psifiaAjias)
                End If
           
           ' End If

132         SYN_ME_FPA = SYN_ME_FPA + mMIK_AJIA
        Next

        Dim NXMLTEXT As Integer
134     Dim CXMLTEXT As String: CXMLTEXT = ""
        
136     For NXMLTEXT = 1 To 100
138         If Len(f_deltiaMarks(NXMLTEXT)) > 0 Then
140             CXMLTEXT = CXMLTEXT + f_deltiaMarks(NXMLTEXT) + ","
            Else
                Exit For
            End If
        Next

        Dim OTHERMOVEPURPOSETITLE As String
142     OTHERMOVEPURPOSETITLE = "" + other.Text

144     Dim M_APAL As String: M_APAL = Format(Val(Left(apallagesFPA.Text, 2)), "#0")

146     apallagesFPA.Text = ""
        Dim mAnastolhFPA As String
148     mAnastolhFPA = IIf(AnastolhFPA.Value = True, "1", "0")
150     SQLTIM = "Insert INTO TIM (OTHERMOVEPURPOSETITLE,SXETMARKS,ANASTOLHFPA,INCMARK,B_C2,SXETMARK,APALAGIFPA,HME,ATIM,AJ1,AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9,"
152     SQLTIM = SQLTIM + "FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,"
154     SQLTIM = SQLTIM + "KPE,AJI,ELGA,EIDOS,TRP,EIDPAR,B_N1,B_N2,"
156     SQLTIM = SQLTIM + "PARAT,METAF,SKOPOS,PROOR,FORTOSH,AYTOK,KERDOS,KLEIDI,PARAKRATISI,KR1,ORA,EXCHANGERATE,CURRENCY) VALUES ('" + OTHERMOVEPURPOSETITLE + "','" + CXMLTEXT + "'," + mAnastolhFPA + ",'-','" + fS_Parakratisis + "','" + par11Value.Caption + "'," + M_APAL + ","
158     SQLTIM = SQLTIM + "'" + Format(f_dat, "MM/DD/YYYY") + "',"  ' HME
160     SQLTIM = SQLTIM + "'" + Left(f_matim, f_psifia_atim + 2) + "',"    'ATIM
162     f_suma = 0

164     For k = 1 To 9

166         'If F_XONDR = 0 Then
168         '    SQLTIM = SQLTIM + str(Round(SumFpa(k) / ((100 + g_Fpa(k)) / 100), f_psifiaAjias)) + "," 'AJ1,aj2..aj9
170         '    f_suma = f_suma + Round(SumFpa(k) / ((100 + g_Fpa(k)) / 100), f_psifiaAjias)
            'Else
172             SQLTIM = SQLTIM + str(Round(SumFpa(k), f_psifiaAjias)) + ","    'AJ1,aj2..aj9
174             f_suma = f_suma + Round(SumFpa(k), f_psifiaAjias)
            'End If

        Next
    
        Dim flag2 As Integer

176     flag2 = 0
    
178     For k = 1 To 9

180         If k <> 5 Then
182             If SumFpa(k) > 0 And flag2 = 0 Then
184                 'If F_XONDR = 0 Then  'lianiki
186                 '    SQLTIM = SQLTIM + str(Round(SumFpa(k) / ((100 + g_Fpa(k)) / 100) * g_Fpa(k) / 100 + Val(TelFpa), f_psifiaAjias)) + "," 'fpa1,fpa2,fpa3..fpa9
188                 '    f_suma = f_suma + Round(SumFpa(k) / ((100 + g_Fpa(k)) / 100) * g_Fpa(k) / 100 + Val(TelFpa), f_psifiaAjias)
                    'Else
190                     SQLTIM = SQLTIM + str(Round(SumFpa(k) * g_Fpa(k) / 100 + Val(TelFpa), f_psifiaAjias)) + "," 'fpa1,fpa2,fpa3..fpa9
192                     f_suma = f_suma + Round(SumFpa(k) * g_Fpa(k) / 100 + Val(TelFpa), f_psifiaAjias)
                    'End If

194                 flag2 = 1
            
                Else
            
196                 'If F_XONDR = 0 Then  'lianiki
198                 '    SQLTIM = SQLTIM + str(Round((SumFpa(k) / ((100 + g_Fpa(k)) / 100)) * g_Fpa(k) / 100, f_psifiaAjias)) + ","    'fpa1,fpa2,fpa3..fpa9
200                 '    f_suma = f_suma + Round((SumFpa(k) / ((100 + g_Fpa(k)) / 100)) * g_Fpa(k) / 100, f_psifiaAjias)
                    'Else
202                     SQLTIM = SQLTIM + str(Round(SumFpa(k) * g_Fpa(k) / 100, f_psifiaAjias)) + ","    'fpa1,fpa2,fpa3..fpa9
204                     f_suma = f_suma + Round(SumFpa(k) * g_Fpa(k) / 100, f_psifiaAjias)
                    'End If
                End If
            
            End If

        Next

206     SQLTIM = SQLTIM + "'" + PEL.Recordset("kod") + "',"    'KPE

208     F_ELGA = 0

210     If F_XONDR = 3 Then
212         'F_ELGA = Round(((SYN_ME_FPA - gVal(sfpa.Caption)) * 0.25 / 100) * 1.036, f_psifiaAjias)
            F_ELGA = Round(((SYN_ME_FPA - gVal(sfpa.Caption)) * 0.25 / 100), f_psifiaAjias)
        End If
        
214     If F_XONDR = 24 Then
216         F_ELGA = Round((SYN_ME_FPA * 2.4 / 100), f_psifiaAjias)
        End If

218     SQLTIM = SQLTIM + str(Round(SYN_ME_FPA - F_ELGA + Val(TelFpa), f_psifiaAjias)) + "," 'AJI
220     SQLTIM = SQLTIM + str(Round(F_ELGA, f_psifiaAjias)) + ","    'ELGA
222     SQLTIM = SQLTIM + "'" + PEL.Recordset("EIDOS") + "',"    'EIDOS

        Dim r55 As New ADODB.Recordset

224     r55.Open "select * from PINAKES WHERE ID=" + str(Combo4.ItemData(FIND_LISTINDEX(Combo4, Combo4.Text))), Gdb, adOpenDynamic, adLockOptimistic

226     Dim m_pliromi As String: m_pliromi = "1;ΜΕ"
228     If IsNull(r55!C1) Then

            '    If Combo4.Text = Combo4.List(0) Then    ' METRHTOIS
            ' SQLTIM = SQLTIM + "'" + LTrim(str(r55!AYJON)) + ";" + r55!PERIGRAFH + "'," 'ΤRP
            '       Else
            '620         SQLTIM = SQLTIM + "'ΠΙ',"  'ΤRP
            '      End If
            '  m_pliromi = LTrim(str(nNull(r55!AYJON))) + ";" + CNull(r55!PERIGRAFH)
        Else
            ' m_pliromi = Trim(CNull(r55!C1)) ' C1=>  ESODA ΔΕΝ ΞΕΡΩ ΤΙ ΕΞΥΠΗΡΕΤΕΙ ΚΑΙ ΧΑΛΑΕΙ ΤΑ MYDATA
            ' SQLTIM = SQLTIM + "'" + Trim(r55!C1) + "'," 'ΤRP
        End If
        
230     m_pliromi = LTrim(str(nNull(r55!AYJON))) + ";" + CNull(r55!PERIGRAFH)
        
232     If Trim(m_pliromi) = "" Then  'ΕΜΦΑΝΙΣΤΗΚΕ ΣΕ ΚΑΡΑΓΚΙΑΒΟΥΡΗ 7-9-22 ΝΑ ΕΧΕΙ ΚΕΝΟ ΤΡΟΠΟ ΠΛΗΡΩΜΗΣ
234         m_pliromi = LTrim(str(Combo4.ListIndex + 1)) + ";" + CNull(Combo4.Text)
        End If
        
        PROORISMOS.Text = Trim(STREET.Text) + " " + Trim(STREETNUMBER.Text) + "," + Trim(CITY.Text) + "," + Trim(POSTALCODE.Text)
        
236     SQLTIM = SQLTIM + "'" + m_pliromi + "',"

238     SQLTIM = SQLTIM + "'" + Left(Combo8.Text, 3) + "',"    'ΕΙDPAR
240     SQLTIM = SQLTIM + Left(APOTH1.Text, 2) + ","    ' B_N1
242     SQLTIM = SQLTIM + Left(APOTH2.Text, 2) + ","    ' B_N2

        'On Error GoTo enhm_eggtim_Err
244     If Len(parat.Text) > mLenPARAT Then parat.Text = Left(parat.Text, mLenPARAT)
246     If Len(SKOPOS.Text) > mLenSKOPOS Then parat.Text = Left(SKOPOS.Text, mLenSKOPOS)
248     If Len(fortosh.Text) > mLenFORTOSH Then fortosh.Text = Left(fortosh.Text, mLenFORTOSH)
250     If Len(PROORISMOS.Text) > mLenPROORISMOS Then PROORISMOS.Text = Left(PROORISMOS.Text, mLenPROORISMOS)
252     If Len(AYTOKINHTO.Text) > mLenAYTOKINHTO Then AYTOKINHTO.Text = Left(AYTOKINHTO.Text, mLenAYTOKINHTO)

254     SQLTIM = SQLTIM + "'" + Replace(parat.Text, "'", "`") + "',"    ' PARAT
256     SQLTIM = SQLTIM + str(Val(Replace(metaf.Text, ",", "."))) + ","    ' METAF
258     SQLTIM = SQLTIM + "'" + Replace(SKOPOS.Text, "'", "`") + "',"    ' SKOPOS
260     SQLTIM = SQLTIM + "'" + Replace(PROORISMOS.Text, "'", "`") + "',"    ' PROORISMOS
262     SQLTIM = SQLTIM + "'" + Replace(fortosh.Text, "'", "`") + "',"    ' FORTOSH
264     SQLTIM = SQLTIM + "'" + Replace(AYTOKINHTO.Text, "'", "`") + "',"    ' AYTOKINHTO
266     SQLTIM = SQLTIM + str(get_kerdos) + ","  ' KERDOS

268     If f_pol = "1" Then
270         SQLTIM = SQLTIM + "'" + Left(f_matim, f_psifia_atim + 2) + "',"    ' KEY
        Else
272         SQLTIM = SQLTIM + "'" + Left(f_matim, f_psifia_atim + 2) + PEL.Recordset("kod") + "',"    ' KEY
        End If

274     SQLTIM = SQLTIM + str(Val(Replace(parakratisiT.Text, ",", "."))) + "," + str(Val(Replace(AJIA2PARAKR.Text, ",", "."))) + ",'" + Format(Time$, "HH:mm") + "'," + Replace(str(gVal(isotimia.Text)), ",", ".") + ",'" + CURR.Text + "')"

        Dim NTIM As Long

        Dim RRR  As New ADODB.Recordset

276     RRR.Open "select COUNT(*)  from TIM where ATIM='" + f_matim + "' and HME='" + Format(f_dat, "MM/DD/YYYY") + "'", Gdb, adOpenDynamic, adLockOptimistic

278     If F_TRANS = 1 And RRR(0) > 0 Then
280         Gdb.RollbackTrans
282         MsgBox "Υπάρχει ήδη το παραστατικό " + f_matim + Chr(13) + " δεν κατεχωρήθη"
284         F_TRANS = -1    'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI

            Exit Sub

        End If

286     RRR.Close

288     NTIM = GDBEXE("ENHM_EGGTIM", SQLTIM) ' Gdb.Execute SQLTIM, NTIM

        ' ΣΥΛΛΕΓΩ  TO ID_NUM TOY TIM ΠOY ΔΗΜΙΟΥΡΓΗΘΗΚΕ
        'Dim m_ID_NUM As Long ' id tim

290     RRR.Open "SELECT   MAX(ID_NUM) FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
292     m_ID_NUM = RRR(0)
294     RRR.Close
   
        ' ΔΕΝ ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ RECORD ΣTO ΤΙΜ
   
296     If NTIM = 0 Then
298         If F_TRANS = 1 Then
300             Gdb.RollbackTrans
302             MsgBox "Λάθος στην επικεφαλίδα του παραστατικού" + Chr(13) + " δεν κατεχωρήθη" + Chr(13) + SQLTIM
304             MsgBox Err.Description
                
            Else
306             Gdb.Execute "DELETE FROM EGGTIM WHERE HME='" + Format(f_dat, "MM/DD/YYYY") + "' AND ATIM='" + Left(f_matim, f_psifia_atim + 2) + "'", NTIM
308             MsgBox Err.Description + " Λάθος στην επικεφαλίδα του παραστατικού" + Chr(13) + " δεν κατεχωρήθη" + Chr(13) + SQLTIM
            End If
     
            'SAVE_ERROR Err.Description & " in Project1.Form1.cmdCommand2_Click " & " at line " & Erl
            Dim F As Integer

310         F = FreeFile
312         Open "C:\MERCVB\ERR.TXT" For Append As #F
314         Write #F, Format(Now, "DD/MM/YYYY HH:MM") + SQLTIM

316         Close #F
318         F_TRANS = -1    'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI

            Exit Sub

        End If
        
320     If F_XONDR = 3 Then
            ' F_ELGA = Round((SYN_ME_FPA * 0.25 / 100) , f_psifiaAjias)  F_ELGA
             
322         Gdb.Execute "UPDATE TIM SET B_C2='" + f_elga_typoi_krat + "',PARAKRATISI=" + Replace(str(F_ELGA), ",", ".") + "  WHERE ID_NUM=" + str(m_ID_NUM)  'SYSK  'monada
324        ' Gdb.Execute "UPDATE TIM SET KR1=" + Replace(str(F_ELGA - F_ELGA / 1.036), ",", ".") + "  WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK  'monada
               Gdb.Execute "UPDATE TIM SET KR1=0  WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK  'monada
        End If
        
        Gdb.Execute "UPDATE TIM SET STREET='" + Replace(Left(STREET.Text, 80), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK  'monada
        Gdb.Execute "UPDATE TIM SET STREETNUMBER='" + Replace(Left(STREETNUMBER.Text, 4), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK  'monada
        Gdb.Execute "UPDATE TIM SET POSTALCODE='" + Replace(Left(POSTALCODE.Text, 6), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK  'monada
        Gdb.Execute "UPDATE TIM SET CITY='" + Replace(Left(CITY.Text, 30), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK  'monada
        
        Gdb.Execute "UPDATE TIM SET BRANCH='" + Replace(Left(BRANCH.Text, 30), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK  'monada
        
          Gdb.Execute "UPDATE TIM SET HMEPARAD='" + Format(HMERPARAD.Value, "MM/dd/yyyy") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK  'monada
        
         Gdb.Execute "UPDATE TIM SET ORAPARAD='" + Format(ORAPARAD.Value, "hh:mm") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK  'monada
       
        
        
        
        
326     If F_METAFORIKES >= 1 Then
            
328         Gdb.Execute "UPDATE TIM SET C2='" + Replace(Trim(Left(Textm(2), 120)), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK  'monada
330         Gdb.Execute "UPDATE TIM SET C3='" + Replace(Trim(Left(Textm(3), 120)), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'EMPOREYMA
332         Gdb.Execute "UPDATE TIM SET C1='" + Replace(Trim(Textm(0)), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'PERIGRAFH
            
334         Gdb.Execute "UPDATE TIM SET C12='" + Replace(Trim(Left(Textm(6), 120)), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'SYSK 2
336         Gdb.Execute "UPDATE TIM SET C13='" + Replace(Trim(Left(Textm(7), 120)), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'EMPOREYMA 2
            
            Gdb.Execute "UPDATE TIM SET C5='" + Replace(Trim(TC5.Text), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM)  'EMPOREYMA 2
            Gdb.Execute "UPDATE TIM SET C6='" + Replace(Trim(TC6.Text), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM)  'EMPOREYMA 2
            Gdb.Execute "UPDATE TIM SET C7='" + Replace(Trim(tc7.Text), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM)  'EMPOREYMA 2
            
            
338         If Len(Trim(Textm(1).Text)) = 0 Then
340             Textm(1).Text = "0"
            End If
            
342         If Len(Trim(Textm(4).Text)) = 0 Then
344             Textm(4).Text = "0"
            End If
346         Gdb.Execute "UPDATE TIM SET SKOPOS2='" + Replace(Trim(Left(TextEntoleas, 120)), "'", "`") + "' WHERE ID_NUM=" + str(m_ID_NUM) 'EMPOREYMA 2
348         TextEntoleas = ""
350         TextOnEntolea = ""
            
352         Gdb.Execute "UPDATE TIM SET NUM1=" + Textm(1) + " WHERE ID_NUM=" + str(m_ID_NUM) 'AR TEMAX
354         Gdb.Execute "UPDATE TIM SET NUM11=" + Textm(5) + " WHERE ID_NUM=" + str(m_ID_NUM) 'AR TEMAX 2
356         Gdb.Execute "UPDATE TIM SET NUM2=" + Textm(4) + " WHERE ID_NUM=" + str(m_ID_NUM) 'MIKTO BAROS
            Dim kmetaf As Integer
358         For kmetaf = 0 To 4
360             Textm(kmetaf).Text = ""
            Next
            
        End If

362     rt.Close
        're.Close

        '------------------------------ ΕΙΣΑΓΩΓΗ EGGTIM ---------------------------
364     For k = 1 To 10: FEGG(k) = 0: Next

        '------------------------------eggtim -----------------------------------------
366     For k = 1 To Grid1.rows - 1

368         If Len(Grid1.TextMatrix(k, f_k)) > 0 Then

370             mEID.Open "SELECT * FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic    'adOpenForwardOnly, adLockReadOnly

                'ΠΡΟΣΘΕΤΕΙ ΜΙΑ ΕΓΓΡΑΦΗ ΓΙΑ ΚΑΘΕ ΧΡΩΜΑ ή ΜΕΓΕΘΟΣ
372             If F_EXO_XROMATA = 0 Then
374                 GoSub ADD_RECORD
                Else

                    Dim K2 As Integer, k3 As Integer

                    'ΣΤΗΝ ΔΙΟΡΘΩΣΗ ΔΙΟΡΩΝΕΙ RECORD-RECORD
376                 If F_DIOR = 1 Then
378                     GoSub ADD_RECORD
                    Else

380                     For K2 = 1 To 25: For k3 = 1 To 25

382                             If Val(XROMATA(k, K2, k3)) <> 0 And Val(XROMATA(k, K2, k3)) <> -0.01 Then
384                                 GoSub ADD_RECORD
386                                 XROMATA(k, K2, k3) = 0
                                End If

388                             Next: Next

                        End If
                    End If

390                 mEID.Close
                End If

            Next
            
            'βαζω αποθηκη  = με την προέλευση για κλινική
            '
392         If F_PIN_PROELEYS = 11 Then

394             If F_STADIO = 5 Or PEL.Recordset("EIDOS") = "e" Then     ' ενδοδιακινηση ή πωλήσεις δεν εχει νοημα να βαζω αποθηκες σε καθε σειρά
                Else
396                 Gdb.Execute "UPDATE EGGTIM SET APOT=PROELEYSH WHERE ID_NUM=" + str(m_ID_NUM)
                End If
            End If
            
            'εαν δουλεύω ΕΦΚ τότε αποθήκευσε το στο τιμολόγιο
398         If F_EIDFOROSKAT = 1 And F_AJIA_APOU > 0 Then ' F_PIN_PROELEYS = 12 Then

                Dim refk As New ADODB.Recordset

400             refk.Open "SELECT SUM(EFK) AS SEFK,SUM(LITRA) AS SLITRA FROM EGGTIM WHERE  ID_NUM=" + str(m_ID_NUM), Gdb, adOpenDynamic, adLockOptimistic
402             Gdb.Execute "UPDATE TIM SET EFK=" + Replace(str(refk(0)), ",", ".") + ",LITRA=" + Replace(str(refk(1)), ",", ".") + " WHERE ID_NUM=" + str(m_ID_NUM)
404             Gdb.Execute "UPDATE TIM SET AJ2=AJ2+EFK,FPA2=FPA2+EFK*0.23 WHERE ID_NUM=" + str(m_ID_NUM)
            Else
406             Gdb.Execute "UPDATE TIM SET EFK=0,LITRA=0  WHERE ID_NUM=" + str(m_ID_NUM)
            End If
            
            ' τελικες τιμες στο τιμολογιο να μην εχω "ΑΣΥΜΦΩΝΙΑ ΠΑΡΑΣΤΑΤΙΚΟΥ"
            Dim R_N As New ADODB.Recordset

            Dim kkk

408         Dim AJ(9)  As Double, FP(9) As Double: For kkk = 1 To 9: AJ(kkk) = 0: FP(kkk) = 0: Next

            Dim MSUMA3 As Single

            'Dim mStrogg_Fpa As Single: 'mStrogg_Fpa = 0
            '---------------------------------------------------------------------------------
410         R_N.Open "SELECT SUM(round(POSO*TIMM*(100-EKPT2)/100*(100-EKPT)/100+ 0.000001,2)) AS KAU,FPA FROM EGGTIM WHERE ID_NUM=" + str(m_ID_NUM) + " GROUP BY FPA", Gdb, adOpenDynamic, adLockOptimistic

412         Do While Not R_N.EOF

414             If R_N("FPA") > 0 And R_N("FPA") <= 9 Then
416                 kkk = R_N("FPA")
418                 AJ(kkk) = Round(R_N("KAU") + 0.0001, 2)

420                 If AJ(kkk) <> 0 Then
                        'If Val(lab_strogg.Caption) = 0.01 Then
                        '   mStrogg_Fpa = 0.01
                        '  lab_strogg.Caption = ""
                        ' End If
          
                        ' If lab_strogg.Caption = "-0.01" Then
                        '    mStrogg_Fpa = -0.01
                        '    lab_strogg.Caption = ""
                        'End If
422                     FP(kkk) = Round(0.0001 + R_N("KAU") * (g_Fpa(kkk)) / 100, 2)     '  + mStrogg_Fpa
                        'mStrogg_Fpa = 0
                    End If

424                 MSUMA3 = MSUMA3 + AJ(kkk) + FP(kkk)
                End If

426             R_N.MoveNext
            Loop

428         R_N.Close

            'ΒΡΙΣΚΩ ΤΟ ΜΕΓΑΛΥΤΕΡΟ ΣΕ ΑΞΙΑ ΦΠΑ ΓΙΑ ΝΑ ΤΟ (ΜΕΓΑΛΩΣΩ/ΜΙΚΡΥΝΩ) ΓΙΑ ΝΑ ΣΥΜΦΩΝΕΙ ΜΕ ΤΗΝ ΟΘΟΝΗ
            Dim ll         As Integer

430         Dim mMAXfpa    As Single: mMAXfpa = -9999999

432         Dim mMAXSynt   As Integer: mMAXSynt = 1 ' σε ποιο συντελεστή εμφανίζει το μέγιστο ΦΠΑ

434         Dim mSynoloFpa As Single: mSynoloFpa = 0

436         For ll = 1 To 9

438             If FP(ll) > mMAXfpa Then
440                 mMAXfpa = FP(ll): mMAXSynt = ll
                End If

442             mSynoloFpa = mSynoloFpa + FP(ll)
            Next

            'ΕΔΩ ΔΙΝΩ ΤΗΝ ΔΙΑΦΟΡΑ ΣΤΟ ΜΕΓΑΛΥΤΕΡΟ ΦΠΑ ΠΟΥ ΒΡΗΚΑ ΠΑΡΑΠΑΝΩ
444         If Abs(mSynoloFpa - gVal(sfpa.Caption)) > 0.001 And Abs(mSynoloFpa - gVal(sfpa.Caption)) < 0.02 Then
446             If gVal(sfpa.Caption) > mSynoloFpa Then
448                 FP(mMAXSynt) = FP(mMAXSynt) + Abs(mSynoloFpa - gVal(sfpa.Caption))
450                 MSUMA3 = MSUMA3 + Abs(mSynoloFpa - gVal(sfpa.Caption))
                Else
452                 FP(mMAXSynt) = FP(mMAXSynt) - Abs(mSynoloFpa - gVal(sfpa.Caption))
454                 MSUMA3 = MSUMA3 - Abs(mSynoloFpa - gVal(sfpa.Caption))
                End If
            End If

            'SynoloG.Caption = Format(MSUMA3, "#####0.00")

            '---------------------------------------------------------------------------

            Dim SQL3 As String

456         SQL3 = "UPDATE TIM SET AJ1=" + Replace(Format(AJ(1), "#########.00"), ",", ".") + ","
458         SQL3 = SQL3 + " AJ2=EFK+" + Replace(Format(AJ(2), "#########.00"), ",", ".") + ","
460         SQL3 = SQL3 + " AJ3=" + Replace(Format(AJ(3), "#########.00"), ",", ".") + ","
462         SQL3 = SQL3 + " AJ4=" + Replace(Format(AJ(4), "#########.00"), ",", ".") + ","
464         SQL3 = SQL3 + " AJ5=" + Replace(Format(AJ(5), "#########.00"), ",", ".") + ","
466         SQL3 = SQL3 + " AJ6=" + Replace(Format(AJ(6), "#########.00"), ",", ".") + ","
468         SQL3 = SQL3 + " AJ7=" + Replace(Format(AJ(7), "#########.00"), ",", ".") + ","
470         SQL3 = SQL3 + " AJ8=" + Replace(Format(AJ(8), "#########.00"), ",", ".") + ","
472         SQL3 = SQL3 + " AJ9=" + Replace(Format(AJ(9), "#########.00"), ",", ".") + ","

474         SQL3 = SQL3 + " FPA1=" + Replace(Format(FP(1), "#########.00"), ",", ".") + ","
476         ' SQL3 = SQL3 + " FPA2=EFK*0.23+" + Replace(Format(FP(2), "#########.00"), ",", ".") + ","
            SQL3 = SQL3 + " FPA2=" + Replace(Format(FP(2), "#########.00"), ",", ".") + ","

478         SQL3 = SQL3 + " FPA3=" + Replace(Format(FP(3), "#########.00"), ",", ".") + ","
480         SQL3 = SQL3 + " FPA4=" + Replace(Format(FP(4), "#########.00"), ",", ".") + ","

482         SQL3 = SQL3 + " FPA6=" + Replace(Format(FP(6), "#########.00"), ",", ".") + ","
484         SQL3 = SQL3 + " FPA7=" + Replace(Format(FP(7), "#########.00"), ",", ".") + ","
486         SQL3 = SQL3 + " FPA8=" + Replace(Format(FP(8), "#########.00"), ",", ".") + ","
488         SQL3 = SQL3 + " FPA9=" + Replace(Format(FP(9), "#########.00"), ",", ".")

490         SQL3 = SQL3 + " WHERE ID_NUM=" + str(m_ID_NUM)

492         Gdb.Execute SQL3, kkk
494         SQL3 = "UPDATE TIM SET AJI=AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7+AJ8+AJ9+FPA1+FPA2+FPA3+FPA4+FPA6+FPA7+FPA8+FPA9 "
496         SQL3 = SQL3 + " WHERE ID_NUM=" + str(m_ID_NUM)
498         Gdb.Execute SQL3, kkk

500         R_N.Open "SELECT AJI FROM TIM WHERE ID_NUM=" + str(m_ID_NUM), Gdb, adOpenDynamic, adLockOptimistic
502         SynoloG.Caption = Format(R_N(0), "#####0.00")
504         R_N.Close

            '=================================================================================================
506         If f_IS_PARTIDES = 1 Then

                'ΠΡΕΠΕΙ ΝΑ ΔΩ ΑΝ ΣΒΗΣΤΗΚΑΝ ΚΑΠΟΙΟΙ ΚΩΔΙΚΟΙ ΠΟΥ ΚΟΥΒΑΛΟΥΣΑΝ ΠΑΡΤΙΔΕΣ
                ' ΚΑΙ ΝΑ ΣΒΗΣΩ ΤΙΣ ΠΑΡΤΙΔΕΣ ΠΟΥ ΚΟΥΒΑΛΟΥΣΑΝ
508             Dim DEL_PART As String, COUNT_DEL As Integer: DEL_PART = "": COUNT_DEL = 0

                Dim rpart    As New ADODB.Recordset

510             If F_DIOR = 1 Then
512                 If F_POS_APOU = 1 Or F_POS_APOU = 4 Then
514                     rpart.Open "select *  FROM PARTIDES WHERE ID_NUM=" + str(F_ID_NUM_DIOR), Gdb, adOpenDynamic, adLockOptimistic

516                     Do While Not rpart.EOF
518                         DEL_PART = DEL_PART + CNull(rpart!LOTNUMBER) + Chr(13)
520                         COUNT_DEL = COUNT_DEL + 1
522                         rpart.MoveNext
                        Loop

524                     If COUNT_DEL > 0 Then
526                         MsgBox "Προσοχη !!! θα διαγραφουν οι παρτιδες :" + Chr(13) + DEL_PART
528                         Gdb.Execute "delete FROM PARTIDES WHERE ID_NUM=" + str(F_ID_NUM_DIOR)
                        End If

530                     rpart.Close
                    End If
                End If
            End If

            '===================================================================================================

532         For k = 1 To f_SeiresTimologioy
534             gm_str(k) = ""
            Next
            
536         If f_MYDATAcont = 1 And isAkyrotiko = False Then
538             If Len(Trim(f_myEID)) > 0 And f_ISMYDATA = 1 Then '  As String ' ΤΥΠΟΣ ΠΑΡΑΣΤΑΤΤΙΚΟΥ MYDATA  Π.X. 1.1 ΤΙΜΟΛΟΓΙΟ
                    ' If POLAGOR.ListIndex = 1 Then  ' ΣΤΙς ΠΩΛΗΣΕΙΣ ΤΟ ΑΛΛΑΖΩ
                    'Par7MyData.Visible = False
540                 Dim mStelno As Integer: mStelno = 1
542                 If f_checkMYDATA = 1 Then
544                     If MsgBox("Να αποσταλεί Mydata;", vbYesNo) = vbYes Then
546                         mStelno = 1
                        Else
548                         mStelno = 0
                        End If
                    End If
550                 If mStelno = 1 Then
552                     gisOnlineMydata = 1
554                     Par7MyData.where.Caption = " where ID_NUM=" + str(m_ID_NUM) + "  " ' M_idnum
556                     Par7MyData.Visible = False
                     
558                     If F_PAROCHOS = 1 Then
560                         Par7MyData.loadpar7
562                         Dim ubl As String: ubl = GGET_CVALUE("SELECT left(ISNULL(UBL,'000'),3) AS UBL2 FROM PARASTAT WHERE EIDOS='" & Left(f_matim, 1) & "'")

                            'e?? paroxo mono sthn xondriki kai to parastattiko anti na exei 380 exei "mydata"
564                         If ubl = "000" Then 'mydata
566                             Par7MyData.ToXMLsub 0, ""
                            Else
568                             Par7MyData.ToJason 0, 0
                            End If
                     
                            'Par7MyData.ToJason 0
                        Else
                     
570                         Par7MyData.ToXMLsub 0, ""
                        End If
                        
                    End If
                End If
            End If
            
            ' Gdb.CommitTrans

            Exit Sub

            '----------------------------------------------------------------------------
ADD_RECORD:

572         mPoso = 0: mKAU_AJIA = 0: mMIK_AJIA = 0: mProeleysh = ""

574         If F_EXO_XROMATA = 0 Then

576             mPoso = Val(Grid1.TextMatrix(k, f_p))

578            ' If F_XONDR = 0 Then ' lianikh
580            '     mMIK_AJIA = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias)
582            '     mKAU_AJIA = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias) / (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100)
               ' Else
584                 mKAU_AJIA = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias)
586                 mMIK_AJIA = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias) * (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100)
               ' End If
                
588             If Len(Grid1.TextMatrix(k, F_PROEL)) > 0 Then
590                 mProeleysh = Grid1.TextMatrix(k, F_PROEL)
                Else
                    
592                 mProeleysh = ""
                    'End If
                    
                End If

            Else

594             If XROMATA(k, 0, 0) = -0.01 Then    ' ΣΗΜΑΔΙ ΟΤΙ ΧΡΗΣΙΜΟΠΟΙΗΣΕ ΤΟ F6
                    '  mPoso = XROMATA(k, K2, k3)
                    '  mProeleysh = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + Format(K2, "00") + "-" + Format(k3, "00")

596                 If F_DIOR = 1 Then
598                     mPoso = Val(Grid1.TextMatrix(k, f_p))    'XROMATA(k, K2, k3)
600                     mProeleysh = Left(XROMATA(k, 1, 0), 3) + Replace(Left(XROMATA(k, 0, 1), 2), "*", "") + "*" + Format(K2, "00") + "-" + Format(k3, "00")
                    Else
602                     mPoso = XROMATA(k, K2, k3)
604                     mProeleysh = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + Format(K2, "00") + "-" + Format(k3, "00")
                    End If

                Else

                    '            If F_DIOR = 1 Then
                    '               mPoso = Val(Grid1.TextMatrix(k, f_p))
                    '               'ΑΥΤΑ ΠΟΥ ΕΡΧΟΝΤΑΙ ΜΕ BARCODE TRABANE TIS STHLES ΑΠ ΕΥΘΕΙΑΣ ΑΠΟ ΤΟ TABLE BARCODES TO OPOIO TO TSIMPAO OTAN TO BREI KAI TO BAZO STO XROMATA(k, K2, 0)
                    '               mProeleysh = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + mID(XROMATA(k, K2, 0), 4, 4)   '+ Format(K2, "00") + "-" + Format(k3, "00")
                    '
                    '
                    '
                    '            Else
                    '               mPoso = Val(Grid1.TextMatrix(k, f_p))
                    '               'ΑΥΤΑ ΠΟΥ ΕΡΧΟΝΤΑΙ ΜΕ BARCODE TRABANE TIS STHLES ΑΠ ΕΥΘΕΙΑΣ ΑΠΟ ΤΟ TABLE BARCODES TO OPOIO TO TSIMPAO OTAN TO BREI KAI TO BAZO STO XROMATA(k, K2, 0)
                    '               mProeleysh = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + mID(XROMATA(k, K2, 0), 4, 4)   '+ Format(K2, "00") + "-" + Format(k3, "00")
                    '            End If

606                 If F_DIOR = 1 Then
608                     mPoso = gVal(Grid1.TextMatrix(k, f_p))
                        
                        'ΑΥΤΑ ΠΟΥ ΕΡΧΟΝΤΑΙ ΜΕ BARCODE TRABANE TIS STHLES ΑΠ ΕΥΘΕΙΑΣ ΑΠΟ ΤΟ TABLE BARCODES TO OPOIO TO TSIMPAO OTAN TO BREI KAI TO BAZO STO XROMATA(k, K2, 0)
610                     mProeleysh = Left(XROMATA(k, 1, 0), 3) + Replace(Left(XROMATA(k, 0, 1), 2), "*", "") + "*" + mID(XROMATA(k, K2, 0), 4, 4)   '+ Format(K2, "00") + "-" + Format(k3, "00")
                    Else
612                     mPoso = gVal(Grid1.TextMatrix(k, f_p))
                        
                        'ΑΥΤΑ ΠΟΥ ΕΡΧΟΝΤΑΙ ΜΕ BARCODE TRABANE TIS STHLES ΑΠ ΕΥΘΕΙΑΣ ΑΠΟ ΤΟ TABLE BARCODES TO OPOIO TO TSIMPAO OTAN TO BREI KAI TO BAZO STO XROMATA(k, K2, 0)
614                     mProeleysh = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + mID(XROMATA(k, K2, 0), 4, 4)   '+ Format(K2, "00") + "-" + Format(k3, "00")
                    End If
                End If

616             mKAU_AJIA = Round(mPoso * Val(Grid1.TextMatrix(k, F_T)) * (100 - Val(Grid1.TextMatrix(k, f_e))) / 100, f_psifiaAjias)
618             mMIK_AJIA = mKAU_AJIA * (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100)
            End If

            ' ΒΑΖΕΙ ΤΗΝ ΣΥΣΚΕΥΑΣΙΑ ΜΕΣΑ ΣΤΟ ΟΝΟΜΑ ΓΙΑ ΤΙΜΟΛΟΓΗΣΗ
620         If mLenOnoma > 35 Then
622             mONOMA = Trim(Left(Grid1.TextMatrix(k, f_o), mLenOnoma))    ' 24/8/2011  ήταν 45 και το εκανα 65 για ΟΣΣΑ ΡΩΞΑΝΗ
            Else
624             mONOMA = Left(Grid1.TextMatrix(k, f_o), 35)
            End If

            '   mONOMA = Left(Grid1.TextMatrix(k, f_o), 35)
626         If f_use_Sysk = 1 Then

                '        If f_use_Sysk = 1 Then
                '         If meid("SYSKMAX") > 1 Then
                '           MHKOS = Len(Trim(Format(meid("SYSKMAX"), "####") + "X" + Format(re("poso") / meid("SYSKMAX"), "####")))
                '           re("onoma") = Left(Grid1.TextMatrix(k, f_o) + Space(35), 34 - MHKOS) + " " + Format(meid("SYSKMAX"), "####") + "X" + Format(re("poso") / meid("SYSKMAX"), "####")
                '          End If
                '        End If

628             If mEID("SYSKMAX") > 1 Then

630                 MHKOS = Len(Trim(Format(mEID("SYSKMAX"), "####") + "X" + Format(mPoso / mEID("SYSKMAX"), "####")))

632                 mONOMA = Left(Grid1.TextMatrix(k, f_o) + Space(35), 34 - MHKOS) + " " + Format(mEID("SYSKMAX"), "####") + "X" + Format(mPoso / mEID("SYSKMAX"), "####")

634                 If Left(f_pelMERC, 3) = "OLY" Then
636                     mONOMA = Left(Grid1.TextMatrix(k, f_o), 35)
                    End If

638                 If Left(f_pelMERC, 3) = "PER" Then
                        'MHKOS = Len(Trim(meid("CH2")) + " " + Format(mPoso / meid("SYSKMAX"), "####") * Format(meid("SYSKMAX"), "####"))
                        'MHKOS = Len(Trim(meid("CH2")) + " " + Format(mPoso / meid("SYSKMAX"), "####") + Format(meid("SYSKMAX"), "####"))
640                     MHKOS = Len(Trim(mEID("CH2")) + " " + Format(mPoso / mEID("SYSKMAX"), "####") + Format(mEID("SYSKMAX"), "####"))
642                     mONOMA = Left(Grid1.TextMatrix(k, f_o) + Space(45), 40 - MHKOS) + "    " + Trim(mEID("CH2")) + " " + Format(mPoso / mEID("SYSKMAX"), "####") + " X " + Format(mEID("SYSKMAX"), "####")
                    End If

                Else

644                 If mLenOnoma > 35 Then
646                     mONOMA = Trim(Left(Grid1.TextMatrix(k, f_o), mLenOnoma))
                        ' mONOMA = Left(Grid1.TextMatrix(k, f_o), 45)
                    Else
648                     mONOMA = Left(Grid1.TextMatrix(k, f_o), 35)
                    End If
                End If

            End If

650         mONOMA = Replace(mONOMA, "'", """")

652         mXre = 0: mPis = 0

654         If F_EXO_XROMATA = 0 Then

656             Select Case Val(F_POS_APOU)

                    Case 1    ' +EISAGOGES
658                     mXre = Val(Grid1.TextMatrix(k, f_p))  're("POSO")

660                 Case 2    ' -EISAGOGES
662                     mXre = -Val(Grid1.TextMatrix(k, f_p))  ' re("POSO")

664                 Case 3    ' +EXAGOGES
666                     mPis = Val(Grid1.TextMatrix(k, f_p))    ' re("POSO")

                        'αν ειναι sda επιστροφης τότε να αφαιρεί την επιστρεφόμενη ποσότητα για να γινει το υπόλοιπο
                        'επιστρεφόμενη ποσοτητα= Val(Grid1.TextMatrix(k, f_p)) - Val(Grid1.TextMatrix(k, F_mikta))
                        'επιστρεφόμενη ποσοτητα= mpis - Val(Grid1.TextMatrix(k, F_mikta))
668                     If sda_epistrofis.Enabled = True Then  ' Right(PARAS.Text, 1) = Right("  " + f_SDA, 1) Then
                            '  Grid1.TextMatrix(k, F_mikta) = POL_POS
670                         mPis = mPis - Val(Grid1.TextMatrix(k, F_mikta))
                        End If

672                 Case 4    ' +EISAGOGES
674                     mPis = -Val(Grid1.TextMatrix(k, f_p))  ' re("POSO")
                End Select

            Else

676             Select Case Val(F_POS_APOU)

                    Case 1    ' +EISAGOGES
678                     mXre = mPoso

680                 Case 2    ' -EISAGOGES
682                     mXre = -mPoso

684                 Case 3    ' +EXAGOGES
686                     mPis = mPoso

688                 Case 4    ' +EISAGOGES
690                     mPis = -mPoso
                End Select

            End If

            ' ΕΝΗΜΕΡΏΝΕΙ ΤΟ ΚΈΡΔΟς

692         mKERDOS = 0
694         mEID.Close
696         mEID.Open "SELECT * FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic    'adOpenForwardOnly, adLockReadOnly

698         If PEL.Recordset("EIDOS") = "e" And Val(Grid1.TextMatrix(k, F_T)) > 0 Then
700             If IsNull(mEID("XTI")) Then
702                 mKERDOS = Val(Grid1.TextMatrix(k, F_T))
                Else
704                 mKERDOS = Val(Grid1.TextMatrix(k, F_T)) - mEID("XTI")
                End If

            End If
            
            '            If f_IS_PARTIDES = 1 Then
            '               UPDATE_POL_PARTIDES k, Grid1.TextMatrix(k, f_k), m_ID_NUM, mXre, mPis, PEL.Recordset("EIDOS") + PEL.Recordset("kod"), Format(f_dat, "MM/DD/YYYY"), Left(f_matim, f_psifia_atim + 2)
            '            End If
            
            '            ' na bazei to lotnumber sto eggtim
            '            If f_IS_PARTIDES = 1 Then
            '                           Dim mRpart As New ADODB.Recordset
            '                           mRpart.Open "select LOTNUMBER from PARTIDES WHERE SEIRA=" + str(k) + " and ID_NUM=" + str(m_ID_NUM), Gdb, adOpenDynamic, adLockOptimistic
            '                           If Not mRpart.EOF Then
            '                                mProeleysh = mRpart(0)
            '                           End If
            '                           mRpart.Close
            '            End If
            
            Dim OTHERMEASUREMENTUNITQUANTITY As String, OTHERMEASUREMENTUNITTITLE As String
706         OTHERMEASUREMENTUNITQUANTITY = "0": OTHERMEASUREMENTUNITTITLE = ""

708         sql = ""
710         sql = "INSERT INTO EGGTIM ( OTHERMEASUREMENTUNITQUANTITY , OTHERMEASUREMENTUNITTITLE,PROOD_AJ,HME,ATIM,MONA,EKPT,TIMM,"
712         sql = sql + "POSO,KAU_AJIA,MIK_AJIA, PROELEYSH,"

714         sql = sql + "FPA,PELKOD,EIDOS,FCURRENCY,APOT,"
716         sql = sql + "ONOMA,XRE,PIS,KODE,KERDOS,MIKTA,EKPT2,ID_NUM,KOLA,MIKTAKILA,LITRA,EFK "
718         sql = sql + ") values (" + OTHERMEASUREMENTUNITQUANTITY + ",'" + OTHERMEASUREMENTUNITTITLE + "',"
720         sql = sql + str(k) + ","    ' prood_aj
722         sql = sql + "'" + Format(f_dat, "MM/DD/YYYY") + "',"  ' HME
724         sql = sql + "'" + Left(f_matim, f_psifia_atim + 2) + "',"    'ATIM
726         sql = sql + "'" + Grid1.TextMatrix(k, F_M) + "',"  'mona
728         sql = sql + str(gVal(Grid1.TextMatrix(k, f_e))) + ","  'ekpt

730         'If F_XONDR = 0 Then ' lianikh
732          '   sql = sql + str(gVal(Grid1.TextMatrix(k, F_T)) / (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100)) + "," 'timm
            'Else
734             sql = sql + str(gVal(Grid1.TextMatrix(k, F_T))) + ","  'timm
            'End If

736         sql = sql + str(mPoso) + ","    ' POSO
738         sql = sql + str(Round(mKAU_AJIA, 2)) + ","   ' KAU_aj
740         sql = sql + str(Round(mMIK_AJIA, 2)) + ","    ' MIK_ajIA

742         If F_KOLA_MIKTA = 1 Then
744             sql = sql + "'" + Left(mProeleysh, 40) + "',"    ' PROELEYSH
            Else
746             sql = sql + "'" + Left(mProeleysh, 10) + "',"    ' PROELEYSH
            End If

748         sql = sql + str(Val(Grid1.TextMatrix(k, f_f))) + ","  'FPA
750         sql = sql + "'" + PEL.Recordset("kod") + "',"    'PELKOD
752         sql = sql + "'" + PEL.Recordset("EIDOS") + "',"    'EIDOS
754         sql = sql + str(gVal(isotimia.Text) * Val(Grid1.TextMatrix(k, F_T))) + ","    'FCURRENCY
756         sql = sql + Left(APOTH1.Text, 2) + ","  'APOT
758         sql = sql + "'" + mONOMA + "',"    'ONOMA
760         sql = sql + str(mXre) + ","  'XRE
762         sql = sql + str(mPis) + ","  'PIS
764         sql = sql + "'" + Grid1.TextMatrix(k, f_k) + "',"  'KODE
766         sql = sql + str(mKERDOS) + ", "  'kerdos
768         sql = sql + str(gVal(Grid1.TextMatrix(k, F_mikta))) + ", "      'mikta
770         sql = sql + str(gVal(Grid1.TextMatrix(k, F_EKPT2))) + ","
772         sql = sql + str(m_ID_NUM) + ","
774         sql = sql + str(gVal(Grid1.TextMatrix(k, F_COL_KOLA))) + ","
776         sql = sql + str(gVal(Grid1.TextMatrix(k, F_mikta_KILA))) + ","
  
778         MLITRA = nNull(mEID("LITRA")) * mPoso
780         MEFK = nNull(mEID("EFK")) * mPoso
  
782         sql = sql + Replace(str(MLITRA), ",", ".") + ","
784         sql = sql + Replace(str(MEFK), ",", ".")
  
786         sql = sql + ") " 'ID_NUM

788         N = 0
790         N = GDBEXE("ENHM_EGGTIM", sql)  'Gdb.Execute sql, N
              
            ' η καθε παρτιδα παίρνει και την σειρα του ειδους
792         If f_IS_PARTIDES = 1 Then
794             UPDATE_POL_PARTIDES k, Grid1.TextMatrix(k, f_k), m_ID_NUM, mXre, mPis, PEL.Recordset("EIDOS") + PEL.Recordset("kod"), Format(f_dat, "MM/DD/YYYY"), Left(f_matim, f_psifia_atim + 2)
            End If

            '================================================================================
796         If f_IS_PYROSB = 1 And InStr(Grid1.TextMatrix(k, f_k), "-") > 1 Then  'εαν εχι κωδικο cccc-x
                'ΕΠΙΚΕΦΑΛΙΔΑ ΓΙΑ ΚΑΘΕ ΠΕΛΑΤΗ  EIDOS=0
                '0004     HME        0    P6      P12      P25
                '0004    20050228    1           5-1-46     'ΕΤΣΙ ΕΙΝΑΙ ΣΤΟ ΑΡΧΕΙΟ
                '0004    20050524    1   5-1-14  5-1-21
                '0004    20060219    1           3-1-46
                  
798             kPyr = InStr(Grid1.TextMatrix(k, f_k), "-")
800             aPyr = Left(Grid1.TextMatrix(k, f_k), kPyr - 1)  'CCCC
802             KatPyr = mID(Grid1.TextMatrix(k, f_k), kPyr + 1, 1)  'x
                 
                'συνθετω το 5-1-46
804             NEAKIN = KatPyr + "-" + Format(mPoso, "##0") + "-" + Format(Round(gVal(Grid1.TextMatrix(k, F_T)), 0), "##0")
                 
                'If InStr(Grid1.TextMatrix(k, f_k), "-") > 0 Then
806             rPyr.Open "select * from ALLPYR where EIDOS='0' AND  KODPEL='" + PEL.Recordset("kod") + "'", Gdb, adOpenDynamic, adLockOptimistic
                 
808             If rPyr.EOF Then ' ΠΡΩΤΗ ΦΟΡΑ ΑΝΟΙΓΕΙ ΚΑΡΤΕΛΑ STO 0
810                 SQLpyr = "insert into ALLPYR (KODPEL,HME,EIDOS ) VALUES ('" + PEL.Recordset("kod") + "','" + aPyr + "','0')"
812                 Gdb.Execute SQLpyr
                End If

814             rPyr.Close
                 
816             rPyr.Open "select * from ALLPYR where EIDOS='0' AND  KODPEL='" + PEL.Recordset("kod") + "'", Gdb, adOpenDynamic, adLockOptimistic
                'ψαχνω τα πεδια για να βρώ ποιο πεδιο εχει τιμη = με τον κωδικο cccc
                ' ΤΟ ΟΝΟΜΑ ΤΟΥ ΠΕΔΙΟΥ ΘΑ ΕΙΝΑΙ MC + FORMAT(NPYR,"0")
                 
818             For kPyr = 1 To 9

820                 If aPyr = rPyr("mc" + Format(kPyr, "0")) Then
822                     nPyr = kPyr

                        Exit For

                    End If

824                 If IsNull(rPyr("mc" + Format(kPyr, "0"))) Or Len(rPyr("mc" + Format(kPyr, "0"))) < 1 Then
                        'ΓΕΜΙΖΩ ΤΟ ΠΡΩΤΟ ΚΕΝΟ ΠΕΔΙΟ ΜΕ ΤΟ ΟΝΟΜΑ ΤΟΥ ΠΥΡΟΣΒΕΣΤΗΡΑ
826                     Gdb.Execute "UPDATE ALLPYR SET " + "MC" + Format(kPyr, "0") + " ='" + aPyr + "' WHERE EIDOS='0' AND  KODPEL='" + PEL.Recordset("kod") + "'"
828                     nPyr = kPyr

                        Exit For

                    End If

                Next

830             rPyr.Close
                 
                'ΕΝΗΜΕΡΩΝΩ ΤΟ ΠΕΔΙΟ MC? ΜΕ ΤΗΝ ΚΙΝΗΣΗ (neaKin)
832             rPyr.Open "select * from ALLPYR where HME='" + Format(Now, "YYYYmmdd") + "' and EIDOS='1' AND  KODPEL='" + PEL.Recordset("kod") + "' and MC" + Format(nPyr, "0") + " is NULL", Gdb, adOpenDynamic, adLockOptimistic

834             If rPyr.EOF Then
836                 SQLpyr = "insert into ALLPYR (KODPEL,HME,EIDOS," + "MC" + Format(kPyr, "0") + ") VALUES ('" + PEL.Recordset("kod") + "','" + Format(Now, "YYYYmmdd") + "','1','" + NEAKIN + "' )"
838                 Gdb.Execute SQLpyr
                Else
840                 Gdb.Execute "update ALLPYR SET " + "MC" + Format(kPyr, "0") + " = '" + NEAKIN + "' WHERE ID=" + str(rPyr("ID"))
                End If
                 
842             rPyr.Close
  
            End If

            '==========================================================================================

844         If F_TRANS = 1 And N = 0 Then
846             Gdb.RollbackTrans
848             F_TRANS = -1    'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
850             MsgBox "Λάθος στην σειρά " + Format(k, "###") + " στoν κωδικό " + Grid1.TextMatrix(k, f_k) + Chr(13) + " δεν κατεχωρήθη"

                ' Resume Next  'na sbhstei
                Exit Sub

            End If

            ' 5/7/2014 REM  THIS LINE   1840  SYN_ME_FPA = SYN_ME_FPA + mMIK_AJIA
852         FEGG(Val(Grid1.TextMatrix(k, f_f))) = FEGG(Val(Grid1.TextMatrix(k, f_f))) + mMIK_AJIA

            '------------------------------------------------------------------------------------------------
854         If F_STADIO = 5 Then    ' ενδοδιακινηση  Left(f_matim, 1) = "σ" Then

                ' 5/7/2014 EGINE REM ΑΠΟ ΤΟ   ADDNEW  ΕΩΣ TO UPDATE
                '1870      re.AddNew
                '1880      re("hme") = f_dat
                '1890      re("atim") = f_matim
                '1900      If F_EXO_XROMATA = 0 Then
                '1910          re("poso") = Val(Grid1.TextMatrix(k, f_p))
                '          Else
                '1920          re("poso") = XROMATA(k, K2, k3)
                '1930          re("PROELEYSH") = Left(XROMATA(k, K2, 0), 3) + Left(XROMATA(k, 0, k3), 2) + "*" + Format(K2, "00") + "-" + Format(k3, "00")
                '          End If
                '1940      re("timm") = Val(Grid1.TextMatrix(k, F_T))
                '1950      re("kode") = Grid1.TextMatrix(k, f_k)
                '1960      re("onoma") = Left(Grid1.TextMatrix(k, f_o), 30)
                '1970      re("mona") = Right(Grid1.TextMatrix(k, f_o), 3)
                '1980      re("kau_ajia") = Val(Grid1.TextMatrix(k, f_a))
                '1990      re("fpa") = Val(Grid1.TextMatrix(k, f_f))
                '2000      re("ekpt") = Val(Grid1.TextMatrix(k, f_e))
                '2010      re("pelkod") = PEL.Recordset("kod")
                '2020      re("eidos") = PEL.Recordset("eidos")
                '2030      re("APOT") = Val(Left(APOTH2.Text, 2))
                '2040      re("XRE") = 0: re("PIS") = 0    'ΑΠΟΦΥΓΗ NULL
                '2050      re("XRE") = re("POSO")
                '2060      re.Update

                ' to idio me sql (endodiakinhsh)
856             sql = ""
858             sql = "INSERT INTO EGGTIM (PROOD_AJ,HME,ATIM,MONA,TIMM,"
860             sql = sql + "PROELEYSH,EKPT,POSO,KAU_AJIA,MIK_AJIA,"
862             sql = sql + "FPA,PELKOD,EIDOS,APOT,"
864             sql = sql + "ONOMA,XRE,PIS,KODE"
866             sql = sql + ") values ("
868             sql = sql + str(k) + ","    ' prood_aj
870             sql = sql + "'" + Format(f_dat, "MM/DD/YYYY") + "',"  ' HME
872             sql = sql + "'" + Left(f_matim, f_psifia_atim + 2) + "',"    'ATIM
874             sql = sql + "'" + Grid1.TextMatrix(k, F_M) + "',0,"  'mona,τιμμ
  
                'PROELEYSH ,EKPT, POSO
876             If F_EXO_XROMATA = 0 Then
878                 sql = sql + "'',0," + Replace(str(Val(Grid1.TextMatrix(k, f_p))), ",", ".")
                Else
880                 sql = sql + Left(XROMATA(k, K2, 0), 3) + Left(XROMATA(k, 0, k3), 2) + "*" + Format(K2, "00") + "-" + Format(k3, "00") + ",0,"
882                 sql = sql + Replace(str(XROMATA(k, K2, k3)), ",", ".")
                End If
  
884             sql = sql + ",0,0,"     ' POSO,KAU_aj,MIK_ajIA
886             sql = sql + str(Val(Grid1.TextMatrix(k, f_f))) + ","  'FPA
888             sql = sql + "'" + PEL.Recordset("kod") + "',"    'PELKOD
890             sql = sql + "'" + PEL.Recordset("EIDOS") + "',"    'EIDOS
892             sql = sql + Left(APOTH2.Text, 2) + ","  'APOT
894             sql = sql + "'" + mONOMA + "',"    'ONOMA
896             sql = sql + Replace(str(Val(Grid1.TextMatrix(k, f_p))), ",", ".") + ","  'XRE
898             sql = sql + "0,"   'PIS
900             sql = sql + "'" + Grid1.TextMatrix(k, f_k) + "')"  'KODE
902             N = 0
904             N = GDBEXE("ENHM_EGGTIM", sql)  'Gdb.Execute sql, N

906             If F_TRANS = 1 And N = 0 Then
908                 Gdb.RollbackTrans
910                 F_TRANS = -1    'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
912                 MsgBox "Λάθος στην σειρά (ΕΝΔΟΔΙΑΚΙΝΗΣΗ) " + Format(k, "###") + " στoν κωδικό " + Grid1.TextMatrix(k, f_k) + Chr(13) + " δεν κατεχωρήθη"

                    Exit Sub

                End If
            
            End If
      
            '------------------------------------------------------------------------------------------------
914         If F_STADIO = 7 Then    ' φυσικη απογραφή

916             sql = ""
918             sql = "INSERT INTO EGGTIM (PROOD_AJ,HME,ATIM,MONA,TIMM,"
920             sql = sql + "POSO,KAU_AJIA,MIK_AJIA,"
922             sql = sql + "FPA,PELKOD,EIDOS,APOT,"
924             sql = sql + "ONOMA,XRE,PIS,KODE"
926             sql = sql + ") values ("
928             sql = sql + str(k) + ","    ' prood_aj
930             sql = sql + "'" + Format(f_dat, "MM/DD/YYYY") + "',"  ' HME
932             sql = sql + "'" + Left(f_matim, f_psifia_atim + 2) + "',"    'ATIM
934             sql = sql + "'" + Grid1.TextMatrix(k, F_M) + "',0"  'mona,τιμμ
                'If F_XONDR = 0 Then ' lianikh
                '   sql = sql + str(gVal(Grid1.TextMatrix(k, F_T)) / (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100)) + "," 'timm
                'Else
                '   sql = sql + str(gVal(Grid1.TextMatrix(k, F_T))) + ","  'timm
                'End If
936             sql = sql + str(mPoso) + ",0,0,"    ' POSO,KAU_aj,MIK_ajIA
                '2640  sql = sql + str(mKAU_AJIA) + ","    ' KAU_aj
                '2650  sql = sql + str(mMIK_AJIA) + ","    ' MIK_ajIA
                '2660  sql = sql + "'" + Left(mProeleysh, 10) + "',"    ' PROELEYSH

938             sql = sql + str(Val(Grid1.TextMatrix(k, f_f))) + ","  'FPA
940             sql = sql + "'" + PEL.Recordset("kod") + "',"    'PELKOD
942             sql = sql + "'" + PEL.Recordset("EIDOS") + "',"    'EIDOS
                '2700  sql = sql + str(gVal(isotimia.Text) * Val(Grid1.TextMatrix(k, F_T))) + ","    'FCURRENCY
944             sql = sql + Left(APOTH1.Text, 2) + ","  'APOT
946             sql = sql + "'" + mONOMA + "',"    'ONOMA
948             sql = sql + str(mXre) + ","  'XRE
950             sql = sql + str(mPis) + ","  'PIS
952             sql = sql + "'" + Grid1.TextMatrix(k, f_k) + "')"  'KODE
                '2760  sql = sql + str(mKERDOS) + ", "  'kerdos
                '2770  sql = sql + str(Val(Grid1.TextMatrix(k, F_mikta))) + ", "      'mikta
                '2775  sql = sql + str(gVal(Grid1.TextMatrix(k, F_EKPT2))) + ") "      'EKPT2

                '2780  n = 0
                '2790  n = GDBEXE("ENHM_EGGTIM", sql)  'Gdb.Execute sql, N
      
                '      re.AddNew
                '      re("hme") = f_dat
                '      re("atim") = f_matim
                '
                '
                '      If F_EXO_XROMATA = 0 Then
                '          re("poso") = Val(Grid1.TextMatrix(k, f_p))
                '      Else
                '         re("poso") = XROMATA(k, K2, k3)
                '         re("PROELEYSH") = Left(XROMATA(k, K2, 0), 3) + Left(XROMATA(k, 0, k3), 2) + "*" + Format(K2, "00") + "-" + Format(k3, "00")
                '      End If
                '
                '
                '      re("timm") = Val(Grid1.TextMatrix(k, F_T))
                '      re("kode") = Grid1.TextMatrix(k, f_k)
                '      re("onoma") = Left(Grid1.TextMatrix(k, f_o), 30)
                '      re("mona") = Right(Grid1.TextMatrix(k, f_o), 3)
                '      re("kau_ajia") = Val(Grid1.TextMatrix(k, f_a))
                '      re("fpa") = Val(Grid1.TextMatrix(k, f_f))
                '      re("ekpt") = Val(Grid1.TextMatrix(k, f_e))
                '      re("pelkod") = PEL.Recordset("kod")
                '      re("eidos") = PEL.Recordset("eidos")
                '
                '      re("APOT") = Val(Left(APOTH2.Text, 2))
                '      re("XRE") = 0: re("PIS") = 0    'ΑΠΟΦΥΓΗ NULL
                '      re("XRE") = re("POSO")
                '      re.Update
954             sqlins(k) = sql
      
            End If
      
956         Return

MHNYMA:
958         HandleError "Par1:enhm_eggtim"

960         Resume Next

            '<EhFooter>
            Exit Sub

enhm_eggtim_Err:
            '  'MsgBox Err.Description & vbCrLf & _
                "in ADOMERCNEW.par1.enhm_eggtim " & _
                "at line " & Erl, _
                vbExclamation + vbOKOnly, "Application Error"
            SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.enhm_eggtim " & "at line " & Erl

            Resume Next

            '</EhFooter>

    End Sub

'ενημέρωση αποθήκης
Sub enhm_apot()

        '<EhHeader>
        On Error GoTo enhm_apot_Err

        '</EhHeader>

        Dim G, m, GA, MA, p

        'Dim DB As New ADODB.Connection
        'Dim e As New ADODB.Recordset
        'Dim R As New ADODB.Recordset, K
        Dim DB As Database, R As New ADODB.Recordset, k, L As Integer

        Dim r2 As New ADODB.Recordset

        Dim e As New ADODB.Recordset

        Dim NeaTim

        Dim gm_f_gfkod As String

100     Erase f_gfkod
102     Erase f_gfaj

        ' On Error GoTo MHNYMA

        Dim m2, gm_f_gfFPA

        Dim m_m As Boolean

        Dim A_fpa As Single, a As Single    'αξιες γεφυρώσεων

        '
        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

104     If Len(gdirlog) > 2 Then
            'r2.Open "SELECT * FROM GEFYRES", Gdb
106         r2.Open "SELECT * FROM GEFYRES", Gdb, adOpenKeyset, adLockOptimistic
108         r2.MoveFirst
110         r2.Move F_rec_p - 1
        End If

112     G = "G" + Format(Month(f_dat), "00")
114     m = "M" + Format(Month(f_dat), "00")
116     GA = "GA" + Format(Month(f_dat), "00")
118     MA = "MA" + Format(Month(f_dat), "00")

120     Set fSCR = CreateObject("MSScriptControl.ScriptControl")
122     fSCR.language = "vbscript"
124     fSCR.addObject "eid", EID
126     fSCR.addObject "mactext", mactext

        Dim M_DOROKARTA As Single

        Dim M_DOREID As Single

        Dim M_DORPEL As Single

128     M_DOROKARTA = 0

        'R.MoveFirst
        'e.Index = "EIDko"
        '  DB.Open gDir
130     For k = 1 To Grid1.rows - 1    'Do While Not R.EOF

            'e.Seek "=", R("KODE")

            '    e.Close
132         If Len(Grid1.TextMatrix(k, f_k)) > 0 Then
                'Set gEid = db.OpenRecordset("SELECT *FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'") ' , DB, adOpenKeyset, adLockOptimistic
134             EID.RecordSource = "SELECT *FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'"    ' , DB, adOpenKeyset, adLockOptimistic
136             EID.Refresh

138             If EID.Recordset.RecordCount = 1 Then
140                 p = Val(Grid1.TextMatrix(k, f_p))    '("POSO")
142                 NeaTim = Val(Grid1.TextMatrix(k, F_T)) * (100 - Val(Grid1.TextMatrix(k, f_e))) / 100
144                 a = p * NeaTim
                    'EID.Recordset.Edit

                    'ΥΠΟΛΟΓΙΖΩ ΤΟ ΠΟΣΟ ΤΗΣ ΔΩΡΟΚΑΡΤΑΣ   ΣΕΠ 2012
146                 If Left(f_pelMERC, 3) = "SOK" Then
148                     If IsNull(EID.Recordset("NUM1")) Then
150                         M_DOREID = 0
                        Else
152                         M_DOREID = Val(EID.Recordset("NUM1"))
                        End If

154                     If IsNull(PEL.Recordset("PEK")) Then
156                         M_DORPEL = 0
                        Else
158                         M_DORPEL = Val(PEL.Recordset("PEK"))
                        End If

160                     M_DOROKARTA = M_DOROKARTA + Max(M_DORPEL, M_DOREID) * Val(Grid1.TextMatrix(k, f_p)) * Val(Grid1.TextMatrix(k, F_T)) / 100
162                     F_AJIDOROKARTAS = M_DOROKARTA
                        'parat.Text = F_AJIDOROKARTAS
                    End If

164                 If F_POS_APOU = "1" Or F_AJIA_APOU = "1" Then    ' EISAGOGES=1
166                     EID.Recordset("POS") = EID.Recordset("POS") + p
168                     EID.Recordset(G) = EID.Recordset(G) + p
                        ' EID.Recordset() = EID.Recordset(G) + P

170                     If F_AJIA_APOU = "1" Then
                            ' αν προκειται για αγορες και προμηθευτές
172                         If PEL.Recordset("EIDOS") = "r" And POLAGOR.Text = POLAGOR.List(0) Then
174                             EID.Recordset("prom") = Left(PEL.Recordset("EPO"), 30)
                            End If
                        End If

176                     If IsNull(EID.Recordset("POS_KERD")) Then EID.Recordset("POS_KERD") = 0
                        ' 455

                        '------------------------- ΚΟΣΤΟΛΟΓΗΣΗ ------------------------------------
                        '--------------------------------------------------------------------------
                        
178                     If f_auto_kostologhsh = 1 Then
                            'AN AYJHUHKE H TIMH AGORAS NA ALLAZEI AYTOMATA H TIMH POLISIS
180                         If EID.Recordset("XTI") < NeaTim Then
                                'EID.Recordset("LTI5") = NeaTim * (100 + EID.Recordset("POS_KERD")) / 100 ' PAR11.Text1(4)
182                             EID.Recordset("LTI") = NeaTim * (100 + EID.Recordset("POS_KERD2")) / 100
                                
184                             EID.Recordset("LTI5") = NeaTim * (100 + EID.Recordset("POS_KERD")) / 100 * (100 + g_Fpa(EID.Recordset("fpa"))) / 100     ' PAR11.Text1(4)
                                
                            End If
186                         EID.Recordset("XTI") = NeaTim
                           
                            'ΓΙΑ ΝΑ ΜΗΝ ΠΕΡΑΣΕΙ ΞΑΝΑ ΑΠΟ ΚΟΣΤΟΛΟΓΗΣΗ
188                         kostologhsh.Value = vbUnchecked
                        End If
                        
190                     If kostologhsh And F_STADIO = 4 Then  ' ' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
192                         PAR11.Label10 = EID.Recordset("XTI")
194                         PAR11.label11 = NeaTim

196                         PAR11.ltiproekptosis = Grid1.TextMatrix(k, F_T)
198                         PAR11.UES = EID.Recordset("UES")

200                         If Val(synolo.Caption) <> 0 Then
202                             PAR11.Label4 = 100 * Val(metaf) / Val(synolo)
                            End If

204                         PAR11.POSO.Caption = p
206                         PAR11.Text1(2) = Round(EID.Recordset("POS_KERD"), 3)

208                         If IsNull(EID.Recordset("LTI5")) Then
210                             PAR11.Text1(4) = 0
                            Else
212                             PAR11.Text1(4) = EID.Recordset("LTI5")
                            End If

214                         PAR11.Text1(7) = EID.Recordset("LTI")

216                         If IsNull(EID.Recordset("pos_kerd2")) Then
218                             PAR11.Text1(6) = 0
                            Else
220                             PAR11.Text1(6) = EID.Recordset("pos_kerd2")
                            End If

222                         If IsNull(EID.Recordset("lti2")) Then
224                             PAR11.LTI2 = 0
                            Else
226                             PAR11.LTI2 = EID.Recordset("LTI2")
                            End If
                            
228                         If IsNull(EID.Recordset("lti3")) Then
230                             PAR11.LTI3 = 0
                            Else
232                             PAR11.LTI3 = EID.Recordset("LTI3")
                            End If

                            '   PAR11.Text1(7) = EID.Recordset("LTI2")

                            'PAR11.Text1(2).SetFocus

234                         PAR11.Label1(1).Caption = EID.Recordset("kod") + "-" + EID.Recordset("ono")
236                         PAR11.SHOW 1

238                         If kostologhsh Then  ' μην τυχον και ακυρωσε την κοστολογηση εν τω μεταξύ

240                             If Val(synolo.Caption) <> 0 Then
242                                 EID.Recordset("POS_KERD3") = 100 * Val(metaf) / Val(synolo)
                                End If

244                             If IsNull(EID.Recordset("lastupd")) Then
246                                 EID.Recordset("lastupd") = f_dat
                                Else
248                                 If EID.Recordset("lastupd") < f_dat Then
250                                     EID.Recordset("lastupd") = f_dat
                                    End If
                                
                                End If
                            
252                             EID.Recordset("XTI") = NeaTim

                                '         GoTo 455
                                '  EID.Recordset("POS_KERD") = PAR11.Text1(2)

254                             EID.Recordset("LTI5") = Val(Replace(PAR11.Text1(4), ",", "."))  ' PAR11.Text1(4)

256                             EID.Recordset("LTI") = Val(Replace(PAR11.Text1(7), ",", "."))

258                             EID.Recordset("POS_KERD") = (100 * EID.Recordset("LTI5")) / ((100 + g_Fpa(EID.Recordset("FPA"))) / 100 * EID.Recordset("XTI")) - 100

260                             EID.Recordset("POS_KERD2") = (100 * EID.Recordset("LTI")) / EID.Recordset("XTI") - 100   '  EID.Recordset("XTI") - 100
                                ' If F_XONDRIKH = 3 Then
262                             EID.Recordset("LTI2") = Val(Replace(PAR11.LTI2.Text, ",", "."))
264                             EID.Recordset("LTI3") = Val(Replace(PAR11.LTI3.Text, ",", "."))
                                ' End If
                            End If

                        End If

                        '------------------------- ΚΟΣΤΟΛΟΓΗΣΗ ------------------------------------
                        '--------------------------------------------------------------------------
                    End If

266                 Select Case F_POS_APOU

                        Case "2"    ' -EISAGOGES
268                         EID.Recordset("POS") = EID.Recordset("POS") - p
270                         EID.Recordset(G) = EID.Recordset(G) - p

272                     Case "3"    ' EXAGOGES=1
274                         EID.Recordset("POS") = EID.Recordset("POS") - p
276                         EID.Recordset(m) = EID.Recordset(m) + p

278                     Case "4"    ' -EXAGOGES
280                         EID.Recordset("POS") = EID.Recordset("POS") + p
282                         EID.Recordset(m) = EID.Recordset(m) - p
                    End Select

284                 Select Case F_AJIA_APOU

                        Case "1"    ' EISAGOGES=1

                            ' eid.recordset(GA) = eid.recordset(GA) + a
286                     Case "2"    ' -EISAGOGES

                            '         eid.recordset(GA) = eid.recordset(GA) - a
288                     Case "3"    ' EXAGOGES=1

                            '         eid.recordset(MA) = eid.recordset(MA) + a
290                     Case "4"    ' -EXAGOGES
                            '  eid.recordset(MA) = eid.recordset(MA) - a
                    End Select

292                 If F_EISITHRIA = 1 Then
                  
294                     EID.Recordset("num2") = EID.Recordset("num2") + p
                    End If

296                 EID.Recordset.Update
                    ' e.Close

298                 If Len(gdirlog) > 2 Then

                        ' μαζεύω τις σούμες για τις γεφυρώσεις
                        '====================================================
300                     If f_pol = "1" Then    ' οταν πρόκειται για πωλήσεις

302                         If IsNull(EID.Recordset("kodsynod")) Then
304                             MsgBox "Δεν έχω κωδικό λογιστικής για τον κωδικό " + EID.Recordset("kod")
306                             f_error_gefyr = 2
308                             GoTo OLD1220

                            Else
310                             gm_f_gfkod = EID.Recordset("kodsynod")
                            End If

                        Else    ' agores
312                         gm_f_gfkod = EID.Recordset("kodlogag")
                        End If

                        ' left(eid.recordset("kodsynod')+"018"
314                     gm_f_gfkod = IIf(IsNull(r2("kau" + Grid1.TextMatrix(1, f_f))), "", r2("kau" + Grid1.TextMatrix(1, f_f)))

316                     gvar = mac(gm_f_gfkod)
318                     gm_f_gfkod = gvar    ' 7000018

                        ' "5400"+left(eid.recordset("kodynod"),2)+"04"
320                     gm_f_gfFPA = IIf(IsNull(r2("kFPA" + Grid1.TextMatrix(1, f_f))), "", r2("kfpa" + Grid1.TextMatrix(1, f_f)))
322                     gvar = mac(gm_f_gfFPA)

324                     gm_f_gfFPA = gvar    ' 54007004
326                     A_fpa = a * g_Fpa(Val(Grid1.TextMatrix(1, f_f))) / 100    ' 500 * 18 /100

                        ' τακτοποιώ τις αξίες και τα ΦΠΑ στους πίνακες f_gfkod()  kai f_gfaj()
                        ' Π.Χ.  F_GFKOD(1)="7000018"   F_gfaj(1)=100
                        '       F_GFKOD(2)="54007004"   F_gfaj(2)=18
                        ' τακτοποιώ την καθαρή αξία
328                     For L = 1 To 20

330                         If Len(f_gfkod(L)) = 0 Then    'βρήκα άδειο , σημαίνει ότι δεν βρήκα στα γεμάτα κάποιο να ταιριάζει
332                             f_gfkod(L) = gm_f_gfkod    'den ypirxe to anoigei tora (70-00-00-00)
334                             f_gfaj(L) = a

                                Exit For

                            End If

336                         If f_gfkod(L) = gm_f_gfkod Then    'βρήκα όμοιο στα γεμάτα
338                             f_gfaj(L) = f_gfaj(L) + a    ' soymarei axies

                                Exit For

                            End If

                        Next

                        ' τακτοποιώ το ΦΠΑ
340                     For L = 1 To 20

342                         If Len(f_gfkod(L)) = 0 Then
344                             f_gfkod(L) = gm_f_gfFPA    'den ypirxe to anoigei tora (54-00-70-04)
346                             f_gfaj(L) = A_fpa

                                Exit For

                            End If

348                         If f_gfkod(L) = gm_f_gfFPA Then
350                             f_gfaj(L) = f_gfaj(L) + A_fpa    ' soymarei axies

                                Exit For

                            End If

OLD1220:

                        Next

                    End If    'If Len(gdirlog) Then
                End If  ' If eid.recordset.RecordCount = 1 Then

352             EID.Recordset.Close
            End If    'len( Grid1.TextMatrix(K, f_k) )>0

        Next

354     Set fSCR = Nothing

        Exit Sub

MHNYMA:
356     HandleError "Par1:enhm_apot"

358     Resume Next

        '<EhFooter>
        Exit Sub

enhm_apot_Err:
        '   'MsgBox Err.Description & vbCrLf & _
             "in ADOMERCNEW.par1.enhm_apot " & _
             "at line " & Erl, _
             vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.enhm_apot " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'ενημέρωση πελάτη
Sub enhm_pel(ByVal m_suma As Single, ByVal id_num As Long)

        Dim cc4 As String
        cc4 = CNull(PEL.Recordset("EIDOS"))
        If Len(Trim$(cc4)) = 0 Then
            If InStr(UCase(POLAGOR.Text), "Λ") > 0 Then
                cc4 = "e"
            ElseIf InStr(UCase(POLAGOR.Text), "Γ") > 0 Then
                cc4 = "r"
            Else
                cc4 = "e"
            End If
       
            ' Αγορές
            ' Πωλήσεις
            ' Κιν.Αποθήκης
    
        End If

        Dim ccPel As String
        ' ccPel = CNull(PEL.Recordset("KOD"))
        ' If Len(Trim$(ccPel)) = 0 Then
        ccPel = Text2(0).Text
        '   End If

        '----------------------------------------------------------------------
        ' f_pel = "1"  XREOSI
        ' f_pel = "2"  -XREOSI
        ' f_pel = "3"  PISTOSI
        ' f_pel = "4"  -PISTOSI
        ' f_pel = "5"  0 ΔΑ
        '<EhHeader>
        
        'If Val(Left(Combo2PARAKRATISI.Text, 1)) = 1 Then  'PARAKRATHSEIS
        If gVal(parakratisiT.Text) > 0 Then
        
            If F_ELGA > 0 Then
           
            Else
        
                If Val(Left(Combo2PARAKRATISI.Text, 1)) = 1 Then  'PARAKRATHSEIS
                    m_suma = m_suma - gVal(parakratisiT.Text)
                End If
            
                If Val(Left(Combo2PARAKRATISI.Text, 1)) = 3 Then  'EPIPLEON FOROI
                    m_suma = m_suma + gVal(parakratisiT.Text)
                End If
            End If
           
        End If
        
        On Error GoTo enhm_pel_Err

        '</EhHeader>

        Dim RTemp As New ADODB.Recordset

        Dim SQLT As String, TYP As Single

        Dim N As Long

        Dim X As String

        Dim DB As Database, e As Recordset, R As New ADODB.Recordset, k

        '  On Error GoTo MHNYMA

        Dim SQL22 As String

        'ENHMERVSH ME PROHGOYMENO YPOLOIPO
        '----------------------------------
100     SQLT = "SELECT KOD,"
102     SQLT = SQLT + "SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END )  -"
104     SQLT = SQLT + " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END )   AS YP"
106     SQLT = SQLT + " FROM EGG  where HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND EIDOS='" + cc4 + "' AND KOD='" + ccPel + "' GROUP BY KOD"
108     RTemp.Open SQLT, Gdb, adOpenDynamic, adLockOptimistic

110     If RTemp.EOF Then
112         TYP = 0
        Else

114         If IsNull(RTemp("YP")) Then
116             TYP = 0
            Else
118             TYP = RTemp("YP")
            End If
        End If

120     If IsNull(PEL.Recordset("AYP")) Then
122         TYP = TYP
        Else
124         TYP = TYP + PEL.Recordset("AYP")
        End If
    
126     SQL22 = "UPDATE TIM  SET EKPT5 =" + Replace(Format(TYP, "###########0.00"), ",", ".")
128     SQL22 = SQL22 + " WHERE  ID_NUM=" + str(id_num)
        'HME='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND ATIM='" + f_matim + "'"
130     N = GDBEXE("ENHM_PEL", SQL22)  'Gdb.Execute SQL22, N

132     Gdb.Execute "UPDATE PEL SET ENERGOS=1,TYP=" + Replace(str(TYP), ",", ".") + " WHERE  EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("kod") + "'"
    
        ' R.Open "SELECT TOP 1 * FROM EGG", Gdb, adOpenDynamic, adLockOptimistic

        Dim mXreosi, mPistosi

        Dim mXre

        Dim sql

        Dim n1

134     mXreosi = 0: mPistosi = 0    ' ΑΠΟΦΥΓΗ ΝULL

        Dim FF

136     FF = Left(f_paras, 20)

138     If f_KENTRA_ADYNATISMATOS = 1 Then
140         If PEL.Recordset("EIDOS") = "e" Then
142             Gdb.Execute "UPDATE PEL SET NUM4=(CASE WHEN NUM4 IS NULL THEN 0 ELSE NUM4 END) + " + str(m_suma) + " WHERE  EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("kod") + "'"
            End If
        End If

144     If f_pel = "1" Or f_pel = "4" Or f_pel = "5" Then    ' XREVSH
146         mXre = m_suma

148         If f_pel = "1" Then mXreosi = m_suma
150         If f_pel = "4" Then mPistosi = -m_suma

152         If f_pel = "5" Then
154             mPistosi = 0
156             mXreosi = 0
158             mXre = 0
            End If

        Else
160         mXre = -m_suma

162         If f_pel = "2" Then mXreosi = -m_suma
164         If f_pel = "3" Then mPistosi = m_suma
        End If
    
166     If f_pel = "0" Then
168         mXreosi = 0
170         mPistosi = 0
172         mXre = 0
        End If
    
        '    Dim cc4 As String
        '    cc4 = CNull(PEL.Recordset("eidos"))
        '    If Len(Trim$(cc4)) = 0 Then
        '       If InStr(UCase(POLAGOR.Text), "Λ") > 0 Then
        '          cc4 = "e"
        '       ElseIf InStr(UCase(POLAGOR.Text), "Γ") > 0 Then
        '          cc4 = "r"
        '       Else
        '          cc4 = "e"
        '       End If
        '
        '       ' Αγορές
        '       ' Πωλήσεις
        '       ' Κιν.Αποθήκης
        '
        '    End If
    
174     sql = "INSERT INTO EGG (EIDOS,KOD,HME,USERID,APA,ATIM,AIT,XRE,XREOSI,PISTOSI,IDTIM) VALUES ("
176     sql = sql + "'" + cc4 + "',"    'EIDOS
178     sql = sql + "'" + ccPel + "',"    'KOD

180     sql = sql + "'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "',"  'HME
182     sql = sql + str(gUserId) + ","  'USERID
184     sql = sql + str(Val(mID$(f_matim, 2, 6))) + ","    'APA
186     sql = sql + "'" + f_matim + "',"  'ATIM
188     sql = sql + "'" + FF + "',"  'AIT
190     sql = sql + DOT(str(mXre)) + "," 'XRE
192     sql = sql + DOT(str(mXreosi)) + ","  'XREOSI
194     sql = sql + DOT(str(mPistosi)) + "," + str(id_num) + ")" 'PISTOSI
196     n1 = 0
    
198     If f_pel = "0" Then
200         If F_PARAG_KART = 1 And F_STADIO <= 3 Then  ' ΠΑΡΑΓΓΕΛΙΑ
202             n1 = GDBEXE("ENHM_PEL", sql)  'Gdb.Execute sql, n1
            Else
204             n1 = 1
            End If

        Else
206         n1 = GDBEXE("ENHM_PEL", sql)  ' Gdb.Execute sql, n1
        End If
    
208     If F_TRANS = 1 And n1 = 0 Then
210         Gdb.RollbackTrans
212         MsgBox "Λάθος στην ενημέρωση λογαριασμού " + Chr(13) + " δεν κατεχωρήθη" + Chr(13) + Err.Description

214         F_TRANS = -1    'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI

            Exit Sub

        End If
        
216     If F_TRANS = 0 And n1 = 0 Then
218         MsgBox "Προσοχή δεν ενημερωθηκε η καρτελλα του πελάτη " + Chr(13) + " " + Chr(13) + Err.Description
220         MsgBox "ΝΑ ΕΝΗΜΕΡΩΘΕΙ ΧΕΙΡΟΚΙΝΗΤΑ Η καρτελλα του πελάτη "
        End If

        '610     If f_pelMERC = "SOK" And F_AJIDOROKARTAS > 0 Then
        '            'περναει το ποσό της δωροκάρτας σαν χρέωση
        '620         FF = "ΔΩΡΟΚΑΡΤΑ"  '  Format(F_AJIDOROKARTAS, "###0.00")
        '630         sql = "INSERT INTO EGG (EIDOS,KOD,HME,USERID,APA,ATIM,AIT,XRE,XREOSI,PISTOSI) VALUES ("
        '640         sql = sql + "'" + PEL.Recordset("eidos") + "',"    'EIDOS
        '650         sql = sql + "'" + PEL.Recordset("kod") + "',"    'KOD
        '660         sql = sql + "'" + Format(dtpicker1.Value, "MM/DD/YYYY") + "',"  'HME
        '670         sql = sql + str(gUserId) + ","  'USERID
        '680         sql = sql + mID$(f_matim, 2, 6) + ","    'APA
        '690         sql = sql + "'" + f_matim + "',"  'ATIM
        '700         sql = sql + "'" + FF + "',"  'AIT
        '710         sql = sql + DOT(str(F_AJIDOROKARTAS)) + "," 'XRE
        '720         sql = sql + str(0) + ","  'XREOSI
        '730         sql = sql + DOT(str(F_AJIDOROKARTAS)) + ")"  'PISTOSI
        '
        '740         n1 = GDBEXE("ENHM_PEL", sql) ' Gdb.Execute sql, n1
        '        End If
   
        ' On Error Resume Next

        ' R.Close

222     R.Open "SELECT SUM(XREOSI) AS SX,SUM(PISTOSI) AS SP FROM EGG" & " WHERE EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

224     If IsNull(PEL.Recordset!ayp) Then
226         Gdb.Execute "UPDATE PEL SET AYP=0,TYP=" + DOT(str(R("SX"))) + " - " + DOT(str(R("SP"))) + "" & " WHERE EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("KOD") + "'"
        Else
228         Gdb.Execute "UPDATE PEL SET TYP=AYP + " + DOT(str(R("SX"))) + " - " + DOT(str(R("SP"))) + "" & " WHERE EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("KOD") + "'"
        End If
        
230     R.Close

        Exit Sub

MHNYMA:
232     HandleError "Par1:enhm_pel"

234     Resume Next

ROLL:

236     If F_TRANS = 1 Then
238         Gdb.RollbackTrans
240         MsgBox "Λάθος στην ενημέρωση λογαριασμού " + Chr(13) + " δεν κατεχωρήθη"
242         F_TRANS = -1    'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI

            Exit Sub

        Else
244         MsgBox "Προσοχή δεν πέρασε στην καρτέλα του λογαριασμού "
        End If

246     Resume Next

        '<EhFooter>
        Exit Sub

enhm_pel_Err:
        '  'MsgBox Err.Description & vbCrLf & _
            "in ADOMERCNEW.par1.enhm_pel " & _
            "at line " & Erl, _
            vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.enhm_pel " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub enhm_met(ByVal m_suma As Single, ByVal id_num As Long)
        Dim ccPel As String
        ccPel = Text2(0).Text

        If gVal(parakratisiT.Text) > 0 Then
            If Val(Left(Combo2PARAKRATISI.Text, 1)) = 1 Then  'PARAKRATHSEIS
                m_suma = m_suma - gVal(parakratisiT.Text)
            End If
            
            If Val(Left(Combo2PARAKRATISI.Text, 1)) = 3 Then  'EPIPLEON FOROI
                m_suma = m_suma + gVal(parakratisiT.Text)
            End If
        
        End If

        '
        'ενημέρωση μετρητών
        'ΚΑΙ ΤΑΜΕΙΟ ΚΑΤΑΣΤΗΜΑΤΟΣ ------------------------------------------------
        '----------------------------------------------------------------------
        ' Dim DB As New ADODB.Connection
        ' Dim E As New ADODB.Recordset
        ' Dim R As New ADODB.Recordset, K
        '<EhHeader>
        On Error GoTo enhm_met_Err

        '</EhHeader>
        Dim DB As Database, e As New ADODB.Recordset, R As New ADODB.Recordset, k

        Dim r2 As New ADODB.Recordset

        On Error GoTo MHNYMA

        Dim m_tameio

        'DB.Open gDir
100     R.Open "SELECT TOP 1 * FROM EGG", Gdb, adOpenDynamic, adLockOptimistic

        ' M_SUMA = f_suma ' Round(Val(SynoloG.Caption), 2)

        ' PEL.Recordset.Edit
        Dim TEL_YPOL As Single

        '
        On Error Resume Next

        Dim mXreosi, mPistosi

        Dim mXre

        Dim sql

        Dim n1

        Dim mXPI

102     mXreosi = 0: mPistosi = 0    ' ΑΠΟΦΥΓΗ ΝULL

104     If f_pel = "1" Then  '  TIMOLOGIO , LIANIKH
106         mPistosi = Abs(m_suma)
108         mXPI = "Π"
110     ElseIf f_pel = "2" Then  ' PISTOTIKO TIMOLOGIO
112         mPistosi = -Abs(m_suma)
114         mXPI = "Π"
116     ElseIf f_pel = "3" Then  ' tim.agoras
118         mXreosi = Abs(m_suma)
120         mXPI = "Χ"
122     ElseIf f_pel = "3" Then  ' PISTOTIKO TIMOLOGIO apo promitheyti
124         mXreosi = -Abs(m_suma)
126         mXPI = "Χ"
        End If

128     If PEL.Recordset("eidos") = "e" Then
130         If f_pel = "1" Or f_pel = "4" Then    ' XREVSH
132             mXre = Round(m_suma, 2)
            Else  ' 2 , 3
134             mXre = Round(-m_suma, 2)
            End If

        Else    ' ΠΡΟΜΗΘΕΥΤΕΣ ' ΠΡΟΜΗΘΕΥΤΕΣ   F_PEL=3 ΧΡΕΩΣΗ  4= -ΧΡΕΩΣΗ     1=ΠΙΣΤΩΣΗ   2=-ΠΙΣΤΩΣΗ

136         If f_pel = "2" Then
138             mXre = Round(-m_suma, 2)
            Else    ' 2 , 3
140             mXre = Round(m_suma, 2)
            End If
        End If
        
        Dim CAAXREOPIS As String
        
        Dim R5 As New ADODB.Recordset
142     R5.Open "SELECT * FROM PINAKES WHERE TYPOS=12 AND  ID=" + str(Combo4.ItemData(FIND_LISTINDEX(Combo4, Combo4.Text))), Gdb, adOpenDynamic, adLockOptimistic
144     If IsNull(R5!n1) Then
146         CAAXREOPIS = "0"
        Else
148         CAAXREOPIS = str(nNull(R5!n1))
        End If
        
        Dim cc4 As String
150     cc4 = CNull(PEL.Recordset("eidos"))
152     If Len(Trim$(cc4)) = 0 Then
154         If InStr(UCase(POLAGOR.Text), "Λ") > 0 Then
156             cc4 = "e"
158         ElseIf InStr(UCase(POLAGOR.Text), "Γ") > 0 Then
160             cc4 = "r"
            Else
162             cc4 = "e"
            End If
       
            ' Αγορές
            ' Πωλήσεις
            ' Κιν.Αποθήκης
    
        End If

164     sql = "INSERT INTO EGG (EIDOS,KOD,HME,USERID,APA,ATIM,AIT,XRE,XREOSI,PISTOSI,IDTIM,AAXREOPIS) VALUES ("
166     sql = sql + "'" + cc4 + "',"    'EIDOS
168     sql = sql + "'" + ccPel + "',"    'KOD
170     sql = sql + "'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "',"  'HME
172     sql = sql + str(gUserId) + ","  'USERID
174     sql = sql + str(Val(mID$(f_matim, 2, 6))) + ","    'APA
176     sql = sql + "'" + f_matim + "',"  'ATIM
178     sql = sql + "'" + Left(Combo4.Text, 20) + "',"   'AIT
180     sql = sql + str(mXre) + ","  'XRE
182     sql = sql + str(mXreosi) + ","  'XREOSI
184     sql = sql + str(mPistosi) + "," + str(id_num) + "," + CAAXREOPIS + ")" 'PISTOSI
186     n1 = 0

        On Error Resume Next
188     n1 = GDBEXE("ENHM_MET", sql)
        '430     Gdb.Execute sql, n1

190     If F_TRANS = 1 And n1 = 0 Then
            ' Gdb.RollbackTrans
192         MsgBox "Λάθος στην ενημέρωση Καρτέλας λογαριασμού " + Chr(13) + " δεν κατεχωρήθη ΣΤΗΝ ΚΑΡΤΕΛΑ" + Chr(13) + Err.Description
            'F_TRANS = -1    'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
            ' το εκανα ρεμ γιατι δεν εκοβε το τιμολογιο
            Exit Sub

        End If

        '        '------------------ ΤΑΜΕΙΟ ΚΑΤΑΣΤΗΜΑΤΟΣ ------------------------------------------------
        '480     If Val(Left(Combo4.Text, 2)) > 0 Then
        '490         m_tameio = F_LOGPLHR(Val(Left(Combo4.Text, 2)) - 1)
        '        End If
        '
        '        ' = FindParametroi(1,"PELAT43", "F_TAMEIO", "9990", "ΛΟΓΑΡΙΑΣΜΟΣ ΤΑΜΕΙΟ")
        '
        '500     R2.Open "SELECT * FROM PEL WHERE KOD='" + m_tameio + "' AND EIDOS='e'", Gdb, adOpenForwardOnly, adLockReadOnly
        '
        '510     If R2.EOF Then
        '520         Gdb.Execute "insert into PEL (EIDOS,KOD,EPO) VALUES ('e','" + m_tameio + "','" + Combo4.Text + "')"
        '        End If
        '
        '530     R2.Close
        '
        '540     R.AddNew
        '550     R("hme") = DTPicker1.Value
        '560     R("atim") = f_matim
        '570     R("apa") = Val(mID$(f_matim, 2, 6))
        '580     R("ait") = "ΜΕΤΡΗΤΑ"
        '590     R("XREOSI") = 0: R("PISTOSI") = 0    ' ΑΠΟΦΥΓΗ ΝULL
        '
        '600     If PEL.Recordset("eidos") = "e" Then    ' εισπραξη
        '610         If f_pel = "1" Then R("XREOSI") = m_suma    '
        '620         If f_pel = "2" Then R("XREOSI") = -m_suma
        '630         If f_pel = "3" Then R("PISTOSI") = m_suma
        '640         If f_pel = "4" Then R("XREOSI") = -m_suma
        '        Else
        '650         R("PISTOSI") = m_suma  'πληρωμήρολλ
        '        End If
        '
        '660     R("eidos") = "e"    ' PEL.Recordset("eidos")
        '670     R("kod") = m_tameio    ' PEL.Recordset("kod")
        '680     R.Update
        '
        '690     mXreosi = 0: mPistosi = 0    ' ΑΠΟΦΥΓΗ ΝULL
        '
        '700     If PEL.Recordset("eidos") = "e" Then    ' εισπραξη
        '710         If f_pel = "1" Then mXreosi = m_suma: mXre = m_suma '
        '720         If f_pel = "2" Then mXreosi = -m_suma: mXre = -m_suma '
        '730         If f_pel = "3" Then mPistosi = m_suma: mXre = m_suma '
        '740         If f_pel = "4" Then mXreosi = -m_suma: mXre = -m_suma '
        '        Else
        '750         mPistosi = m_suma  'πληρωμή
        '        End If
        '
        '760     sql = "INSERT INTO EGG (EIDOS,KOD,HME,USERID,APA,ATIM,AIT,XRE,XREOSI,PISTOSI,IDTIM) VALUES ("
        '770     sql = sql + "'e',"    'EIDOS
        '780     sql = sql + "'" + m_tameio + "',"    'KOD
        '790     sql = sql + "'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "',"  'HME
        '800     sql = sql + str(gUserId) + ","  'USERID
        '810     sql = sql + mID$(f_matim, 2, 6) + ","    'APA
        '820     sql = sql + "'" + f_matim + "',"  'ATIM
        '830     sql = sql + "'" + Left("Μετρητά", 20) + "',"   'AIT
        '840     sql = sql + str(mXre) + ","  'XRE
        '850     sql = sql + str(mXreosi) + ","  'XREOSI
        '860     sql = sql + str(mPistosi) + "," + str(ID_NUM) + ")" 'PISTOSI
        '870     n1 = 0
        '
        '        On Error Resume Next
        '
        '880     Gdb.Execute sql, n1
        '
        '        '------------------ ΤΑΜΕΙΟ ΚΑΤΑΣΤΗΜΑΤΟΣ ------------------------------------------------
        '
        '890     R.Close

        Exit Sub

MHNYMA:
194     HandleError "Par1:enhm_met"

196     Resume Next

ROLL:

198     If F_TRANS = 1 Then
200         Gdb.RollbackTrans
202         MsgBox "Λάθος στην ενημέρωση λογαριασμού " + Chr(13) + " δεν κατεχωρήθη"
204         F_TRANS = -1    'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI

            Exit Sub

        Else
206         MsgBox "Προσοχή δεν πέρασαν τα μετρητά  στην καρτέλα του λογαριασμού "
        End If

208     Resume Next

        '<EhFooter>
        Exit Sub

enhm_met_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.enhm_met " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.enhm_met " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'εκτύπωση
Function print_timol(ByVal ATIM As String, ByVal hme As Date)

        '<EhHeader>
        On Error GoTo print_timol_Err

        '</EhHeader>
        Dim DUM

100     Me.Caption = "ektyp_forma"

102     If f_FORM_EKTY = 99 Then
104         DUM = ektyp_forma("c:\mercvb\f90.txt", ATIM, hme)
        Else
106         DUM = ektyp_forma("c:\mercvb\f" + Format(f_FORM_EKTY, "00") + ".txt", ATIM, hme)
        End If

        '
        'GoTo 10

        '<EhFooter>
        Exit Function

print_timol_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.print_timol " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.print_timol " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub b_Katax_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo b_Katax_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = 13 Then
102         KeyAscii = 0
        End If

        '<EhFooter>
        Exit Sub

b_Katax_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.b_Katax_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.b_Katax_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub calcula_Click()
        '   Dim d As Single
        '   d = GET_NVALUE("SELECT ISNULL(TIMH,0) FROM PINAKES WHERE TYPOS=45 AND AYJON=" + Replace(Left(ComboPARAKRATISI.Text, 2), ".", ""))
        '
        '   parakratisiT.Text = Replace(str(gVal(SYNOLO.Caption) * d / 100), ",", ".")

        Dim D    As Single
        Dim mtyp As String
100     Select Case Left(Combo2PARAKRATISI.Text, 1)
            Case "1"
102             mtyp = "45"
104         Case "2"
106             mtyp = "54"
108         Case "3"
110             mtyp = "55"
112         Case "4"
114             mtyp = "56"
        End Select
 
116     D = GET_NVALUE("SELECT ISNULL(TIMH,0) FROM PINAKES WHERE TYPOS=" + mtyp + " AND AYJON=" + Replace(Left(ComboPARAKRATISI.Text, 2), ".", ""))

118     parakratisiT.Text = Replace(str(Round(gVal(synolo.Caption) * D / 100, 2)), ",", ".")

End Sub

Private Sub chkAgor_Click()
100     ar_paraggt_LostFocus
End Sub

Private Sub chkPOL_Click()
100     ar_paraggt_LostFocus
End Sub

Private Sub cmdCAI_Click()
    PAR13.SHOW (1)

End Sub

Private Sub submitpayment_Click()
'submit payment
Dim c, D, SX As String
        'd = TDBGrid1.Columns("ID_NUM").value ' ή TDBGrid1.Col = 7: TDBGrid1.Text

       TDBGrid1.Col = 7  'id_num  ακυρωτικου

       D = TDBGrid1.Text


        Dim midnum As String
       midnum = D '(TDBGrid1.CurrentRow.Cells(7).Value)
        
        Dim cmark As String
        cmark = GGET_CVALUE("select ENTITYMARK FROM TIM WHERE ID_NUM=" + midnum)
        Dim posSignature
        posSignature = GGET_CVALUE("select POSPAYSIGN FROM TIM WHERE ID_NUM=" + midnum)
        Dim POSTransactionId As String
        POSTransactionId = GGET_CVALUE("select TRANSACTIONID FROM TIM WHERE ID_NUM=" + midnum)

        Dim POS_AMOUNT As Single: POS_AMOUNT = GGET_NVALUE("select AJI FROM TIM WHERE ID_NUM=" + midnum)
        Dim AMOUNT As String: AMOUNT = Format(POS_AMOUNT, "###0.00")
        AMOUNT = InputBox("πσο", "ποσο", AMOUNT)

        'sqlDt.Rows(klm)("TRP")
        Dim trp As String
        trp = GGET_CVALUE("select TRP FROM TIM WHERE ID_NUM=" + midnum)


        '"54888937"
        Dim TID As String: TID = GGET_CVALUE("select ISNULL(C1,'') AS C1 from PINAKES where TYPOS=12 AND  AYJON=" + mID(trp, 1, 1))
        'Dim bearer As String = PAR7MYDATA.REQPAROCHOS()
        Par7MyData.loadpar7
        TID = Trim(TID)
        Dim DUM
      DUM = Par7MyData.POS_submitpayment(cmark, midnum, posSignature, TID, AMOUNT, POSTransactionId)
       ' PAR7MYDATA.POS_submitpayment(cMARK, midnum, posSignature, TID, AMOUNT, POSTransactionId)


End Sub

Private Sub cmdPOS_Click()

        Exit Sub
        Dim c, D As String
100     c = "<%PS1;3;" + Replace(SynoloG.Caption, ",", ".") + ">"
        'D = "<%SL;144476123;;;;;;;355;A;201;0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00;EUR;1;15.00;;;;>"
102     D = "<%SL;" + Trim(Text2(2).Text) + ";;;;;;;355;;" + LTrim(Trim(Text1.Text)) + ";0.00;0.00;0.00;0.00;0.00;0.00;0.00;0.00;0.00;0.00;EUR;1;" + LTrim(Trim(Replace(SynoloG.Caption, ",", "."))) + ";;;;>"

        Dim cnn           As ADODB.Connection

        Dim RDB           As ADODB.Recordset

        Dim strConnection As String, strProvider As String, strSource As String

104     strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
106     strSource = "Data Source=C:\MERCVB\REPORTS\REPORTS.mdb;" '& App.Path &

108     Set cnn = New ADODB.Connection
110     strConnection = strProvider & strSource & "Persist Security Info=False"
112     cnn.Open strConnection
114     cnn.Execute "update MEM SET C1='" + D + "'"
116     cnn.Execute "update MEM SET PAR19='" + c + "'"
        ' cnn.Execute "update MEM SET MEMO='" + D + "'"
                
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
                   
118     Ucr9print.printingCR9 F_EKTYPOTHS, "c:\MERCVB\reports\PROEISPR.rpt" ' CrystalReport1.ReportFileName

        '<%PS3;3;15.00>
        '
        'Μαζί λοιπόν με το παραπάνω script, το e.line που θα το συνοδεύει με τα δεδομένα προς τον ΦΗΜ για να γίνει η προείσπραξη είναι το παρακάτω:
        '
        '<%SL;144476123;;;;;;;355;A;201;0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00; 0.00;EUR;1;15.00;;;;>

End Sub

Private Sub cmdPreview_Click()
  
        On Error Resume Next

        ' ' ExportReportToPDF CRXReport, "c:\mercvb\timologio.pdf", "ΤΙΜΟΛΟΓΙΟ"
100   '  Main_Form.ListBox1.AddItem "c:\PDF\PDF.pdf"
 Main_Form.ListBox1.AddItem "c:\MERCVB\timologio.pdf"

102     Main_Form.emailParal = PEL.Recordset("EMAIL")  '"lagakis@otenet.gr"  ' mEMAIL
104     Main_Form.SHOW 1

        'gmail.ListBox1.AddItem "c:\mercvb\timologio.pdf"
        'gmail.TextBox3 = "lagakis@otenet.gr"  ' mEMAIL

        'gmail.SHOW 1

        '
        '  'tdbgrid1.SelBookmarks.Item("id_num"   '
        '    TDBGrid1.Col = 7
        '   ' mID_NUM = TDBGrid1.Text
        '
        '    preview.Label1 = TDBGrid1.Text
        '
        '   preview.SHOW 1
   
End Sub

Private Sub FILL_LOIPOI(ComboPARAKRATISI3 As ComboBox, COMBOTYPOY As ComboBox)

        'Dim R As New ADODB.Recordset

100     If Val(Left(COMBOTYPOY.Text, 1)) = 1 Then ' Or Val(Left(Combo2TyposParakr.Text, 1)) = 1 Then  'PARAKRATHSEIS
            'ΦΟΡΤΩΜΑ parakrathshs
            Dim R26 As New ADODB.Recordset
102         R26.Open "SELECT *FROM PINAKES WHERE TYPOS=45 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
104         ComboPARAKRATISI3.Clear

106         Do While Not R26.EOF
108             ComboPARAKRATISI3.AddItem Format(R26("AYJON"), "#0") + " " + R26("PERIGRAFH")
110             R26.MoveNext
            Loop

112         If ComboPARAKRATISI3.ListCount < 17 Then
114             Gdb.Execute "delete FROM PINAKES WHERE TYPOS=45"

116             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,1,'Περιπτ. β`- Τόκοι - 15% ',15)"
118             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,2,'Περιπτ. γ` - Δικαιώματα - 20% ',20)"

120             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,3,'Περιπτ. δ` - Αμοιβές Συμβουλών Διοίκησης - 20% ',20)"

122             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,4,'Περιπτ. δ` - Τεχνικά Έργα - 3% ',3)"

124             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,5,'Υγρά καύσιμα και προϊόντα καπνοβιομηχανίας 1% ',1)"

126             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,6,'Λοιπά Αγαθά 4% ',4)"

128             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,7,'Παροχή Υπηρεσιών 8% ',8)"

130             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,8,'Προκαταβλητέος Φόρος Αρχιτεκτόνων και Μηχανικών επί Συμβατικών Αμοιβών, για Εκπόνηση Μελετών και Σχεδίων 4% ',4)"

132             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,9,'Προκαταβλητέος Φόρος Αρχιτεκτόνων και Μηχανικών επί Συμβατικών Αμοιβών, που αφορούν οποιασδήποτε άλλης φύσης έργα 10% ',10)"

134             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,10,'Προκαταβλητέος Φόρος στις Αμοιβές Δικηγόρων 15%  ',15)"

136             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,11,'Παρακράτηση Φόρου Μισθωτών Υπηρεσιών παρ. 1 αρ. 15 ν. 4172/2013  ',0)"

138             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,12,'Παρακράτηση Φόρου Μισθωτών Υπηρεσιών παρ. 2 αρ. 15 ν. 4172/2013 - Αξιωματικών Εμπορικού Ναυτικού ',15)"
140             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,13,'Παρακράτηση Φόρου Μισθωτών Υπηρεσιών παρ. 2 αρ. 15 ν. 4172/2013 - Κατώτερο Πλήρωμα Εμπορικού Ναυτικού',10)"

142             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,14,'Παρακράτηση Ειδικής Εισφοράς Αλληλεγγύης(ποσό) ',0)"
144             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,15,'Παρακράτηση Φόρου Αποζημίωσης λόγω Διακοπής Σχέσης Εργασίας παρ. 3 αρ. 15 ν. 4172/2013(ποσό)  ',0)"

146             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,16,'Παρακράτηση συναλλαγών αλλοδαπής βάσει συμβάσεων αποφυγής διπλής φορολογίας(Σ.Α.Δ.Φ.)(ποσό) ',0)"
148             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,17,'Λοιπές Παρακρατήσεις Φόρου(ποσό) ',0)"

150             R26.Close

            End If

        End If

152     If Val(Left(COMBOTYPOY.Text, 1)) = 3 Then '  Or Val(Left(Combo2TyposParakr.Text, 1)) = 3 Then

            'ΦΟΡΤΩΜΑ parakrathshs
            Dim R23 As New ADODB.Recordset
154         R23.Open "SELECT *FROM PINAKES WHERE TYPOS=55 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
156         ComboPARAKRATISI3.Clear

158         Do While Not R23.EOF
160             ComboPARAKRATISI3.AddItem Format(R23("AYJON"), "#0") + " " + R23("PERIGRAFH")
162             R23.MoveNext
            Loop

164         If ComboPARAKRATISI3.ListCount < 17 Then
166             Gdb.Execute "delete FROM PINAKES WHERE TYPOS=55"
         
168             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,1,'1.δεν χρησιμοποιειται ',0)"
170             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,2,'2.δεν χρησιμοποιειται ',0)"
         
172             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,3,'3 β) ασφάλιστρα κλάδου ζωής 4% ',4)"
174             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,4,'4.γ) ασφάλιστρα λοιπών κλάδων 15%. ',15)"
176             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,5,'5.δ) απαλλασσόμενα φόρου ασφαλίστρων 0% ',0)"
178             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,6,'6.Ξενοδοχεία 1-2 αστέρων 0,50 € ποσό ',0.50)"
180             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,7,'7.Ξενοδοχεία 3 αστέρων 1,50 € ποσό ',1.50)"
182             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,8,'8.Ξενοδοχεία 4 αστέρων 3,00 € ποσό   ',3)"
184             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,9,'9.Ξενοδοχεία 4 αστέρων 4,00 € ποσό',4)"
186             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,10,'10.Ενοικιαζόμενα - επιπλωμένα δωμάτια - διαμερίσματα 0,50 € ποσό ',0.50)"
188             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,11,'11.Ειδικός Φόρος στις διαφημίσεις που προβάλλονται από την τηλεόραση (ΕΦΤΔ) 5% ',0.05)"
190             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,12,'12. 3.1 Φόρος πολυτελείας 10% επί της φορολογητέας αξίας για τα ενδοκοινοτικώς αποκτούμενα και εισαγόμενα από τρίτες χώρες 10%  ',0.10)"
192             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,13,'13. 3.2 Φόρος πολυτελείας 10% επί της τιμής πώλησης προ Φ.Π.Α. για τα εγχωρίως παραγόμενα είδη 10%   ',0.10)"
194             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,14,'14.Δικαίωμα του Δημοσίου στα εισιτήρια των καζίνο (80% επί του εισιτηρίου) ',0.80)"
196             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,15,'15.ασφάλιστρα κλάδου πυρός 20%',0.20)"
198             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,16,'16.Τελωνειακοί Δασμοί-Φόροι ποσό  ',0)"
200             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,17,'17.Λοιποί Φόροι ποσό  ',0)"
202             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(55,18,'18.Επιβαρύνσεις Λοιπών Φόρων ποσό ',0)"
            End If

204         R23.Close
        End If

        '3    18,00%
        '4    20,00%
        '5    12,00%
        '6    10,00%
        '7    5,00%
        '8
        '9    2,00%
        '10
        '11  Combo2TyposParakr

206     If Val(Left(COMBOTYPOY.Text, 1)) = 2 Then ' Or Val(Left(Combo2TyposParakr.Text, 1)) = 2 Then  'LOIPOI FOROI

            'ΦΟΡΤΩΜΑ parakrathshs
            Dim R22 As New ADODB.Recordset
208         R22.Open "SELECT *FROM PINAKES WHERE TYPOS=54 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
210         ComboPARAKRATISI3.Clear

212         Do While Not R22.EOF
214             ComboPARAKRATISI3.AddItem Format(R22("AYJON"), "#0") + " " + R22("PERIGRAFH")
216             R22.MoveNext
            Loop

218         If ComboPARAKRATISI3.ListCount < 11 Then
220             Gdb.Execute "delete FROM PINAKES WHERE TYPOS=54"
222             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,1,'1.Για μηνιαίο λογαριασμό μέχρι και 50 ευρώ 12%  ',0.12)"
224             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,2,'2.Για μηνιαίο λογαριασμό από 50,01 μέχρι και 100 ευρώ 15% ',0.15)"
         
226             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,3,'3.Για μηνιαίο λογαριασμό από 100,01 μέχρι και 150 ευρώ 18% ',0.18)"
228             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,4,'4.Για μηνιαίο λογαριασμό από 150,01 ευρώ και άνω 20%',0.20)"
230             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,5,'5.Τέλος καρτοκινητής επί της αξίας του χρόνου ομιλίας (12%) ',0.12)"
232             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,6,'6.Τέλος στη συνδρομητική τηλεόραση 10%   ',0.10)"
234             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,7,'7.Τέλος συνδρομητών σταθερής τηλεφωνίας 5%',0.05)"
236             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,8,'8.Περιβαλλοντικό Τέλος & πλαστικής σακούλας ν. 2339/2001 αρ. 6α 0,07 ευρώ ανά τεμάχιο ποσό ',0.07)"
        
238             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,9,'9.Εισφορά δακοκτονίας 2%',0.02)"
         
240             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,10,'10. Λοιπά τέλη ποσό  ',0)"
242             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,11,'11.Τέλη Λοιπών Φόρων ποσό ',0)"

            End If

244         R22.Close
        End If

        'If Val(Left(Combo2PARAKRATISI.Text, 1)) = 2 Or Val(Left(Combo2TyposParakr.Text, 1)) = 2 Then 'TELH
        '
        '
        '
        '    'ΦΟΡΤΩΜΑ parakrathshs
        '    R.Open "SELECT *FROM PINAKES WHERE TYPOS=54 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
        '    ComboPARAKRATISI3.Clear
        '
        '    Do While Not R.EOF
        '        ComboPARAKRATISI3.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH")
        '        R.MoveNext
        '    Loop
        '
        '    If ComboPARAKRATISI3.ListCount < 11 Then
        '       Gdb.Execute "delete FROM PINAKES WHERE TYPOS=54"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,1,'1.Για μηνιαίο λογαριασμό μέχρι και 50 ευρώ 12%  ',0.12)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,2,'2.Για μηνιαίο λογαριασμό από 50,01 μέχρι και 100 ευρώ 15% ',0.15)"
        '
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,3,'3.Για μηνιαίο λογαριασμό από 100,01 μέχρι και 150 ευρώ 18% ',0.18)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,4,'4.Για μηνιαίο λογαριασμό από 150,01 ευρώ και άνω 20%',0.20)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,5,'5.Τέλος καρτοκινητής επί της αξίας του χρόνου ομιλίας (12%) ',0.12)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,6,'6.Τέλος στη συνδρομητική τηλεόραση 10%   ',0.10)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,7,'7.Τέλος συνδρομητών σταθερής τηλεφωνίας 5%',0.05)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,8,'8.Περιβαλλοντικό Τέλος & πλαστικής σακούλας ν. 2339/2001 αρ. 6α 0,07 ευρώ ανά τεμάχιο ποσό ',0.07)"
        '
        '
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,9,'9.Εισφορά δακοκτονίας 2%',0.02)"
        '
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,10,'10. Λοιπά τέλη ποσό  ',0)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,11,'11.Τέλη Λοιπών Φόρων ποσό ',0)"
        '
        '
        '     End If
        '
        ' R.Close
        'End If

        'If Val(Left(Combo2PARAKRATISI.Text, 1)) = 2 Then 'TELH
        '    'ΦΟΡΤΩΜΑ parakrathshs
        '    R.Open "SELECT *FROM PINAKES WHERE TYPOS=54 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
        '    ComboPARAKRATISI3.Clear
        '    Do While Not R.EOF
        '        ComboPARAKRATISI3.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH")
        '        R.MoveNext
        '    Loop
        '    If ComboPARAKRATISI3.ListCount < 11 Then
        '       Gdb.Execute "delete FROM PINAKES WHERE TYPOS=54"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,1,'1.Για μηνιαίο λογαριασμό μέχρι και 50 ευρώ 12%  ',0.12)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,2,'2.Για μηνιαίο λογαριασμό από 50,01 μέχρι και 100 ευρώ 15% ',0.15)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,3,'3.Για μηνιαίο λογαριασμό από 100,01 μέχρι και 150 ευρώ 18% ',0.18)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,4,'4.Για μηνιαίο λογαριασμό από 150,01 ευρώ και άνω 20%',0.20)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,5,'5.Τέλος καρτοκινητής επί της αξίας του χρόνου ομιλίας (12%) ',0.12)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,6,'6.Τέλος στη συνδρομητική τηλεόραση 10%   ',0.10)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,7,'7.Τέλος συνδρομητών σταθερής τηλεφωνίας 5%',0.05)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,8,'8.Περιβαλλοντικό Τέλος & πλαστικής σακούλας ν. 2339/2001 αρ. 6α 0,07 ευρώ ανά τεμάχιο ποσό ',0.07)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,9,'9.Εισφορά δακοκτονίας 2%',0.02)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,10,'10. Λοιπά τέλη ποσό  ',0)"
        '         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(54,11,'11.Τέλη Λοιπών Φόρων ποσό ',0)"
        '     End If
        '
        ' R.Close
        'End If

246     If Val(Left(COMBOTYPOY.Text, 1)) = 4 Then ' Or Val(Left(Combo2TyposParakr.Text, 1)) = 4 Then  'XARTOS
            'ΦΟΡΤΩΜΑ parakrathshs
            Dim R25 As New ADODB.Recordset
248         R25.Open "SELECT *FROM PINAKES WHERE TYPOS=56 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
250         ComboPARAKRATISI3.Clear
252         Do While Not R25.EOF
254             ComboPARAKRATISI3.AddItem Format(R25("AYJON"), "#0") + " " + R25("PERIGRAFH")
256             R25.MoveNext
            Loop
258         If ComboPARAKRATISI3.ListCount < 4 Or InStr(ComboPARAKRATISI3.List(3), "ψηφ") = 0 Then
260             Gdb.Execute "delete FROM PINAKES WHERE TYPOS=56"
262             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(56,1,'1.Συντελεστής 1,20%  ',1.20)"
264             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(56,2,'2.Συντελεστής 2,40% ',2.40)"
266             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(56,3,'3.Συντελεστής 3,60% ',3.60)"
268             Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(56,4,'4.Λοιπές περιπτώσεις ψηφ.τελους ποσό',0)"
            End If

270         R25.Close
        End If

        '1  2  3  4
End Sub

Private Sub CmdRefreshPel_Click()
100     mmGridPelaton_DblClick 1, 0, 0, 0
End Sub

Private Sub cmdypolo_Click()

        Dim D    As Single
        Dim mtyp As String
100     Select Case Left(Combo2TyposParakr.Text, 1)
            Case "1"
102             mtyp = "45"
104         Case "2"
106             mtyp = "54"
108         Case "3"
110             mtyp = "55"
112         Case "4"
114             mtyp = "56"
        End Select
 
116     D = GET_NVALUE("SELECT ISNULL(TIMH,0) FROM PINAKES WHERE TYPOS=" + mtyp + " AND AYJON=" + Replace(Left(Combo2OnomaParak.Text, 2), ".", ""))

118     AJIA2PARAKR.Text = Replace(str(Round(gVal(synolo.Caption) * D / 100, 2)), ",", ".")

End Sub

Private Sub cmdΜΑΖ_ΕΚΤΥ_Click()
        ' == ΠΡΟΣΟΧΗ ===
        ' DIABAZEI ΤΑ ΣΤΟΙΧΕΙΑ ΑΠΟ ΤΗΝ ΟΘΜΗ (ΠΡΟΟΡΙΣΜΟΣ,ΑΥΤΟΚΙΝΗΤΟ,ΦΟΡΤΩΣΗ

        Dim DB As DAO.Database
100     Set DB = OpenDatabase("c:\mercvb\reports\reports.mdb")

        Dim C1, c2 As String
102     C1 = InputBox("ΔΩΣΕ ΑΡΧΙΚΟ ΠΑΡΑΣΤΑΤΙΚΟ Π.Χ. Τ000123")

104     c2 = InputBox("ΔΩΣΕ ΤΕΛΙΚΟ ΠΑΡΑΣΤΑΤΙΚΟ Π.Χ. Τ000125", "", C1)

106     If Left$(C1, 1) = Left$(c2, 1) Then

        Else
108         MsgBox " ΑΠΟ ΤΗΝ ΙΔΙΑ ΣΕΙΡΑ ΠΡΕΠΕΙ ΝΑ ΕΙΝΑΙ"
            Exit Sub
        End If

110     If c2 < C1 Then
112         MsgBox "TO " + c2 + " PREPEI NA EINAI MEGALYTERO TOY " + C1
            Exit Sub
        End If

        Dim CTIT As String

114     CTIT = GGET_CVALUE("select ISNULL(TITLOS,'')  from PARASTAT WHERE EIDOS='" + Left(C1, 1) + "'")
116     DB.Execute "update MEM SET C1='" + CTIT + "'"

        Dim c    As String
        Dim m_ID As Long

        Dim mF   As String
118     mF = GGET_CVALUE("select ISNULL(FORMA1,'C:\MERCVB\REPORTS\TIMOL1.RPT')  from PARASTAT WHERE EIDOS='" + Left(C1, 1) + "'")

120     F_TYPOS_PAR = GGET_CVALUE("select ISNULL(C1,'173')  from PARASTAT WHERE EIDOS='" + Left(C1, 1) + "'")

        Dim R As New ADODB.Recordset
122     R.Open "SELECT * FROM TIM WHERE LEFT(ATIM,1)='" + Left(C1, 1) + "' AND  ATIM>='" + C1 + "' AND ATIM<='" + c2 + "' ORDER BY ATIM", Gdb, adOpenDynamic, adLockOptimistic
124     R.MoveFirst
126     Do While Not R.EOF

            '        Dim m As Integer
            '
            '        Dim c As String
            '
            '        PAR1.TDBGrid1.Col = 2
            '        c = Replace(PAR1.TDBGrid1.Text, ",", ".")
            '
            '
            '           Dim m_ID As Long
            '         PAR1.TDBGrid1.Col = 7
            '       m_ID = Val(PAR1.TDBGrid1.Text)
        
128         c = str(Val(mID(R!ATIM, 2, 6))) ' Replace(str(R!aji), ",", ".")
130         m_ID = R!id_num
132         Text1.Text = c
            Dim DUM As String
        
            '  PAR1.METASX_SE_TIM 99, C, M_ID, False, 0
            ' PAR1.METASX_SE_TIM 99, C, 0, False

            ' 100     PAR1.diortosis
134         DoEvents
136         f_matim = R!ATIM
138         DTPicker1.Value = R!hme

140         f_same_eline = 0
142         DUM = printCrystal(f_matim, CDate(R!hme))
144         f_same_eline = 0
            Dim plir As String
            'plir = GGET_CVALUE("select PERIGRAFH from PINAKES WHERE TYPOS=12 AND AYJON=" + Left(R!trp, 1))
                
146         plir = GGET_CVALUE("select PERIGRAFH from PINAKES WHERE TYPOS=12 AND AYJON=" + Left(R!trp, 1))
148         DB.Execute "update MEM SET banks='" + plir + "'"
                 
150         CrystalReport1.ReportFileName = mF '  "C:\MERCVB\REPORTS\TIMOL1.RPT" ' F_FORMA1
152         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

            ' EPANEKTYPOSI (2) 'PDF1EKT2ASK0)
        
154         Text2(0).Text = ""
156         Text2(1).Text = ""
        
158         R.MoveNext
        Loop

        'PAR1.Command5_Click
160     Text2(0).Text = "" ' για να μην διορθώση κατα λαθος
        '<EhFooter>

        R.Close

End Sub

Private Sub Combo2PARAKRATISI_Click()
        'If Val(Left(Combo2PARAKRATISI.Text, 1)) = 2 Then
100     FILL_LOIPOI ComboPARAKRATISI, Combo2PARAKRATISI
        'End If

End Sub

Private Sub Combo2TyposParakr_Click()
100     FILL_LOIPOI Combo2OnomaParak, Combo2TyposParakr
        'FILL_LOIPOI ComboPARAKRATISI
End Sub

Private Sub Combo8_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Combo8_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = 13 Then
102         Text2(0).Enabled = True: Text2(0).SetFocus
        End If

        '<EhFooter>
        Exit Sub

Combo8_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Combo8_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Combo8_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub SYMBOL_ONLINE()

        '<EhHeader>
        On Error GoTo SYMBOL_import_Err

        '</EhHeader>

        Dim DB As DAO.Database

        Dim R As DAO.Recordset

        Dim k

        Dim a$, B, X

        Dim D

        Dim FPA As Integer

        Dim rSQL As New ADODB.Recordset

        Dim r2 As New ADODB.Recordset

        '100     Open "C:\APOG.TXT" For Input As #1

100     X = 1

        Dim kod, per, MON, POS, TIMH, TIMOL

        Dim F As String, RECS As Integer

102     k = 0

104     r2.Open "SELECT * FROM TIMAGOR ", Gdb, adOpenDynamic, adLockOptimistic

106     Do While Not r2.EOF

            ' On Error Resume Next

            'Line Input #1, a$

108         X = 1
110         kod = r2("KOD")   '  mID(a$, 1, 13)

112         POS = r2("POSO")  'Val(mID$(a$, 15, 5))

114         TIMOL = 1 'Val(mID$(a$, 20, 6))

            ' If TIMOL = Val(Text1.Text) And Val(kod) > 0 Then
116         rSQL.Open "SELECT KOD FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

118         If rSQL.EOF Then
120             MsgBox "Δεν υπάρχει το είδος " + kod
            Else
122             kod = rSQL(0)
124             rSQL.Close

126             rSQL.Open "SELECT ONO,XTI,FPA FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
128             per = rSQL("ONO")
130             TIMH = rSQL("XTI")
132             FPA = rSQL("FPA")
134             TIMH = Round(rSQL("XTI"), 2)    '  / ((100 + g_Fpa(FPA)) / 100)

            End If

136         rSQL.Close

138         k = k + 1

140         If k > f_SeiresTimologioy Then Exit Do
142         Grid1.TextMatrix(k, f_k) = kod
144         Grid1.TextMatrix(k, f_o) = per
146         Grid1.TextMatrix(k, f_p) = POS
148         Grid1.TextMatrix(k, F_T) = TIMH
150         Grid1.TextMatrix(k, f_f) = FPA
152         Grid1.row = k
154         FindSynolo

            'End If    ' TIMOL=TEXT1

156         r2.MoveNext
            ' If EOF(1) Then Exit Do

        Loop

        ' Close #1

        '<EhFooter>
        Exit Sub

SYMBOL_import_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.SYMBOL_import " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SYMBOL_import " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim DB As DAO.Database

        Dim R As DAO.Recordset

        Dim k

        Dim a$, B, X

        Dim D

        Dim rSQL As New ADODB.Recordset

        Dim rSQL2 As New ADODB.Recordset

100     If Len(Dir("C:\SYMBOL", vbNormal)) > 2 Then
            ' GAZOROS
102         SYMBOL_import

            Exit Sub

        End If

        '        SYMBOL_ONLINE

104     If Len(Dir("C:\SYMBOL_ONLINE", vbNormal)) > 2 Then
            ' DOXATO
106         Me.Caption = "SYMBOL_ONLINE"
108         SYMBOL_ONLINE
            Exit Sub

        End If

110     If Len(Dir("C:\MERCVB\eggtim2.txt", vbNormal)) > 2 Then
    
            Dim SS As Integer
112         SS = MsgBox("ΑΠΟ ΝΕΟ ΠΙΣΤΟΛΑΚΙ;", vbYesNo)
114         If SS = vbYes Then
                ' DOXATO
116             Me.Caption = "SCV"
118             csv_import
                Exit Sub
            End If

        End If

120     If Len(Dir("C:\PAYLOS.EXE", vbNormal)) > 2 Then
            ' GAZOROS
122         PAYLOS_import

            Exit Sub

        End If

124     If Len(Dir("c:\TIMOLOGIA", vbDirectory)) > 2 Then
            ' LOISIR HTR20N007-GP-E-56
126         loisir_import

            Exit Sub

        End If

128     If Len(Dir("c:\WESTNET", vbDirectory)) > 2 Then
            ' WESTNET
130         WESTNET_import

            Exit Sub

        End If

132     If Len(Dir("c:\SOKRATIS", vbDirectory)) > 2 Then
            ' SOKRATIS
134         SOKRATIS_import

            Exit Sub

        End If

        If F_EXCEL_IMPORT = 28028 Then ' Len(Dir("C:\MERCVB\TIMOL.XLSX", vbDirectory)) > 2 Then
            ' SOKRATIS
'TODO: Enter task description here
135         EXCEL28028_IMPORT

            Exit Sub

        End If






136     If F_EXCEL_IMPORT > 99 Then ' Len(Dir("C:\MERCVB\TIMOL.XLSX", vbDirectory)) > 2 Then
            ' SOKRATIS
138         EXCEL_IMPORT

            Exit Sub

        End If

        ' LAMBANEI ENA MONO PARASTAR\TIKO
140     If F_falcon_Polla_tim = 2 Then    ' mpoygoydis

            Dim ANS As Integer

142         ANS = MsgBox("θΈΛΩ ΚΑΙ ΠΟΣΟΤΗΤΕΣ;", vbYesNo)

144         If ANS = vbNo Then

                ' μπουγουδης
146             Set DB = OpenDatabase("C:\LAGEURO\MHXANAKI", False, False, "dBase III;")
148             Set R = DB.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM=" + Left(Text1.Text, 5) + ";")

150             If R.RecordCount = 0 Then
152                 MsgBox "δεν υπάρχουν εγγραφές"

                    Exit Sub

                End If

154             R.MoveFirst
156             k = 0

158             Do While Not R.EOF
160                 k = k + 1

162                 If k > f_SeiresTimologioy Then Exit Do
164                 If IsNull(R("BARCODE")) Then
166                     Grid1.TextMatrix(k, f_k) = " "
                    Else
168                     Grid1.TextMatrix(k, f_k) = R("BARCODE")
                    End If

170                 R.MoveNext
                Loop

172             R.Close
174             DB.Close

                Exit Sub

            End If

        End If
        
        '====================================================
        ' LAMBANEI ENA MONO PARASTAR\TIKO
        '380     If F_REMOTE = 1 Then
        '            REMOTE_IMPORT
        '
        '
        '        End If

        ' LAMBANEI ENA MONO PARASTAR\TIKO
176     If F_falcon_Polla_tim = 0 Or Len(Dir("c:\DIONISIS.EXE", vbDirectory)) > 2 Then

            '
            On Error Resume Next

178         Set DB = OpenDatabase("C:\LAGEURO\MHXANAKI", False, False, "dBase III;")
180         Set R = DB.OpenRecordset("SELECT * FROM TIMAGOR where POS<>0")    '  WHERE ATIM=" + left(Text1.Text, 5))

182         If R.RecordCount = 0 Then
184             MsgBox "δεν υπάρχουν εγγραφές"

                Exit Sub

            End If

186         R.MoveFirst
188         k = 0

            Dim M_CODE

190         Do While Not R.EOF
192             k = k + 1

194             If k > f_SeiresTimologioy Then Exit Do
196             If IsNull(R("BARCODE")) Or R("pos") = 0 Then
198                 Grid1.TextMatrix(k, f_k) = " "
                Else

200                 rSQL.Open "SELECT * FROM BARCODES WHERE left(ERG,13)='" + (Left(R("BARCODE"), 13)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

202                 If rSQL.EOF And Left(R("BARCODE"), 1) = "0" Then
204                     rSQL.Close

206                     rSQL.Open "SELECT * FROM BARCODES WHERE ERG LIKE '" + Trim(mID(R("BARCODE"), 2, 12)) + "%'", Gdb, adOpenForwardOnly, adLockReadOnly
                    End If

208                 If rSQL.EOF Then
210                     M_CODE = " "
212                     Grid1.TextMatrix(k, f_k) = M_CODE
                        '       OK = False
                        '       List2.AddItem rDBF("BARCODE") + ": ΔΕΝ ΥΠΑΡΧΕΙ ΣΤΟ ΑΡΧΕΙΟ ΕΙΔΩΝ"
                    Else
214                     M_CODE = rSQL("KOD")
216                     Grid1.TextMatrix(k, f_k) = M_CODE
                        '   Grid1.TextMatrix(K, F_T) = rSQL("LTI")
218                     Grid1.TextMatrix(k, f_p) = R("POS")
220                     rSQL.Close
222                     rSQL.Open "SELECT * FROM EID WHERE KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
224                     Grid1.TextMatrix(k, F_T) = Round(rSQL("LTI"), 2)
226                     Grid1.TextMatrix(k, f_o) = rSQL("ONO")
228                     Grid1.TextMatrix(k, f_f) = rSQL("FPA")

230                     If f_TIMOKAT = 9999 Then 'περπατουλης τιμοκατάλογος ειδος + πελατης

                            Dim timokPEL As String, timokEID As String

232                         timokPEL = IIf(IsNull(PEL.Recordset("CH4")), "", PEL.Recordset("CH4"))
234                         timokEID = IIf(IsNull(rSQL("CH1")), "", rSQL("CH1"))
                            
236                         rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOKPEL='" + timokPEL + " AND TIMOKEID='" + timokEID + "'", Gdb, adOpenForwardOnly, adLockReadOnly

238                         If rSQL2.EOF Then
240                             Grid1.TextMatrix(k, f_e) = 0
                            Else
242                             Grid1.TextMatrix(k, f_e) = rSQL2(0)
                            End If

244                         rSQL2.Close
                            ' timokatalogos εκπτωσης
246                     ElseIf f_TIMOKAT >= 1 And f_TIMOKAT < 300 And Not IsNull(PEL.Recordset("pek")) Then
248                         rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
250                         If rSQL2.EOF Then
252                             Grid1.TextMatrix(k, f_e) = 0
                            Else
254                             Grid1.TextMatrix(k, f_e) = rSQL2(0)
                            End If

256                         rSQL2.Close

                            'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
258                     ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then

260                         rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly

262                         If rSQL2.EOF Then
                                ' Grid1.TextMatrix(k, F_T) = 0
                            Else
264                             Grid1.TextMatrix(k, F_T) = rSQL2(0)
                            End If

266                         rSQL2.Close

268                     ElseIf f_TIMOKAT = 333 And Not IsNull(PEL.Recordset("pek")) Then
270                         If PEL.Recordset("pek") = 2 Then
272                             Grid1.TextMatrix(k, F_T) = EID.Recordset("lti2")
274                         ElseIf PEL.Recordset("pek") = 3 Then
276                             Grid1.TextMatrix(k, F_T) = EID.Recordset("lti3")
                            Else
278                             Grid1.TextMatrix(k, F_T) = EID.Recordset("lti")
                            End If

                        Else         ' με εκπτωση πελατη
280                         Grid1.TextMatrix(k, f_e) = PEL.Recordset("pek")
                            ' Grid1.Text =
                        End If

282                     Grid1.row = k
284                     FindSynolo

                    End If

286                 rSQL.Close

                End If

288             R.Edit
290             R("pos") = 0
292             R.Update

294             R.MoveNext
            Loop

296         R.Close
298         DB.Close

            Exit Sub

        End If

        '
        On Error Resume Next

300     Set DB = OpenDatabase("C:\LAGEURO\MHXANAKI", False, False, "dBase III;")
302     Set R = DB.OpenRecordset("SELECT * FROM TIMAGOR where ATIM=" + Left(Text1.Text, 5))

304     If R.RecordCount = 0 Then
306         MsgBox "δεν υπάρχουν εγγραφές"

            Exit Sub

        End If

308     R.MoveFirst
310     k = 0

        'Dim M_CODE
312     Do While Not R.EOF
314         k = k + 1

316         If k > f_SeiresTimologioy Then Exit Do
318         If IsNull(R("BARCODE")) Or R("pos") = 0 Then
320             Grid1.TextMatrix(k, f_k) = " "
            Else

322             rSQL.Open "SELECT * FROM BARCODES WHERE LEFT(ERG,13)='" + Left(R("BARCODE"), 13) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

324             If rSQL.EOF And Left(R("BARCODE"), 2) = "00" Then
326                 rSQL.Close

328                 rSQL.Open "SELECT * FROM BARCODES WHERE ERG LIKE '" + mID(R("BARCODE"), 2, 12) + "%'", Gdb, adOpenForwardOnly, adLockReadOnly
                End If

330             If rSQL.EOF Then
332                 M_CODE = " "
334                 Grid1.TextMatrix(k, f_k) = M_CODE
                Else
336                 M_CODE = rSQL("KOD")
338                 Grid1.TextMatrix(k, f_k) = Trim(M_CODE)
340                 Grid1.TextMatrix(k, f_p) = R("POS")
342                 rSQL.Close

344                 rSQL.Open "SELECT * FROM EID WHERE KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
346                 Grid1.TextMatrix(k, F_T) = Round(rSQL("LTI"), 2)
348                 Grid1.TextMatrix(k, f_o) = Trim(rSQL("ONO"))
350                 Grid1.TextMatrix(k, f_f) = rSQL("FPA")

352                 If f_TIMOKAT > 0 And Not IsNull(PEL.Recordset("pek")) Then
354                     rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly

356                     If rSQL2.EOF Then
358                         Grid1.TextMatrix(k, f_e) = 0

                            'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
360                     ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then
362                         rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly

364                         If rSQL2.EOF Then
                                ' Grid1.TextMatrix(k, F_T) = 0
                            Else
366                             Grid1.TextMatrix(k, F_T) = rSQL2(0)
                            End If

368                         rSQL2.Close
370                     ElseIf f_TIMOKAT = 333 And Not IsNull(PEL.Recordset("pek")) Then
372                         If PEL.Recordset("pek") = 2 Then
374                             Grid1.TextMatrix(k, F_T) = EID.Recordset("lti2")
376                         ElseIf PEL.Recordset("pek") = 3 Then
378                             Grid1.TextMatrix(k, F_T) = EID.Recordset("lti3")
                            Else
380                             Grid1.TextMatrix(k, F_T) = EID.Recordset("lti")
                            End If

                        Else
382                         Grid1.TextMatrix(k, f_e) = rSQL2(0)
                        End If

384                     rSQL2.Close
                    Else
386                     Grid1.TextMatrix(k, f_e) = PEL.Recordset("pek")
                        ' Grid1.Text = PEL.Recordset("pek")
                    End If

388                 Grid1.row = k
390                 FindSynolo

                End If

392             rSQL.Close

            End If

            ' R.Edit
            '  R("pos") = 0
            ' R.Update

394         R.MoveNext
        Loop

396     R.Close
398     DB.Close

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub csv_import()
        '0;154;5201890000124 ;88,00;1,2;;0642;100454;30/10/2016
        '0;154;5201219046109 ;42,00;0;;0642;29-931;30/10/2016
        '0;156;5201219046109 ;4,00;0;;;29-931;30/10/2016
        Dim k As Long, a As String
        'On Error Resume Next
        'Close #6

        On Error GoTo LATOS

100     Open "c:\mercvb\eggtim2.txt" For Input As #6
        Dim rSQL   As New ADODB.Recordset
        Dim rSQL2  As New ADODB.Recordset
        Dim rSQL3  As New ADODB.Recordset

        '200     If F_falcon_Polla_tim = 2 Then    ' mpoygoydis
        '
        '            Dim ANS As Integer
        '
        '210         ANS = MsgBox("θΈΛΩ ΚΑΙ ΠΟΣΟΤΗΤΕΣ;", vbYesNo)
        '
        '220         If ANS = vbNo Then
        '
        '                ' μπουγουδης
        '230             Set DB = OpenDatabase("C:\LAGEURO\MHXANAKI", False, False, "dBase III;")
        '240             Set R = DB.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM=" + Left(Text1.Text, 5) + ";")
        '
        '250             If R.RecordCount = 0 Then
        '260                 MsgBox "δεν υπάρχουν εγγραφές"
        '
        '                    Exit Sub
        '
        '                End If
        '
        '270             R.MoveFirst
        '280             k = 0
        '
        '290             Do While Not R.EOF
        '300                 k = k + 1
        '
        '310                 If k > f_SeiresTimologioy Then Exit Do
        '320                 If IsNull(R("BARCODE")) Then
        '330                     Grid1.TextMatrix(k, f_k) = " "
        '                    Else
        '340                     Grid1.TextMatrix(k, f_k) = R("BARCODE")
        '                    End If
        '
        '350                 R.MoveNext
        '                Loop
        '
        '360             R.Close
        '370             DB.Close
        '
        '                Exit Sub
        '
        '            End If
        '
        '        End If
        '
        '        '====================================================
        '

        Dim RR()   As String

        Dim M_CODE As String

102     k = 0
             
        ' LAMBANEI ENA MONO PARASTAR\TIKO
        Dim f_ena As Integer

104     If F_falcon_Polla_tim = 2 Then    ' mpoygoydis
106         f_ena = 1
        Else
108         f_ena = 0
        End If
     
        Dim ANS As Integer

110     ANS = MsgBox("θΈΛΩ ΚΑΙ ΠΟΣΟΤΗΤΕΣ;", vbYesNo)

112     Do While Not EOF(6)
                
114         Line Input #6, a
116         RR = Split(a, ";")
118         M_CODE = RR(7)
120         If (Val(Text1.Text) = Val(RR(1)) And f_ena = 1) Or f_ena = 0 Then '==================================================
122             k = k + 1
124             Grid1.TextMatrix(k, f_k) = M_CODE
126             If ANS = vbYes Then
128                 Grid1.TextMatrix(k, f_p) = RR(3)
130                 rSQL.Open "SELECT * FROM EID WHERE KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
132                 If rSQL.EOF Then  '-------------------------------------------------------
                    Else  ' =========
134                     Grid1.TextMatrix(k, F_T) = Round(rSQL("LTI"), 2)
136                     Grid1.TextMatrix(k, f_o) = Trim(rSQL("ONO"))
138                     Grid1.TextMatrix(k, f_f) = rSQL("FPA")
                    End If ' rSQL.EOF Then  '-------------------------------------------------------
140                 rSQL.Close
142                 FindSynolo
                End If
            End If '  Val(Text1.Text) = Val(RR(1)) Then  '==================================================
144         Grid1.row = k
        Loop

146     Close #6

        Exit Sub

LATOS:
        On Error Resume Next

148     Close #6

End Sub

'Sub LOISIR2_import()
'
'  Dim db As DAO.Database
'  Dim R As DAO.Recordset
'  Dim k
'  Dim A$, b, X
'  Dim d
'  Dim rSQL As New ADODB.Recordset
'
'
'
'
' If Len(Dir("c:\TIMOLOGIA", vbDirectory)) > 2 Then
'   ' LOISIR
'
'
'
'Open "c:\timologia\" + Text1.Text + ".asc" For Binary Access Read As #1
'
''Do While Not EOF(1)
'' Line Input #1, A$
'' Exit Do
''Loop
''Close #1
''Exit Sub
'
'X = 1
'
'
'Dim kod, per, MON, pos, TIMH
'Dim f As String, RECS As Integer
'
'k = 0
'
'Do While Not EOF(1) 'True
'  On Error Resume Next
'    Line Input #1, A$
'   If EOF(1) Then Exit Do
'
'   f = Chr$(10) + Chr$(27) + Chr$(15)
'
''   B = InStr(x, A$, f)
''   If B = 0 Then
' '    GoTo 555
''     Exit Do
'  ' End If
'
'
'
'
'
'X = 1
'
'     f = Chr$(27) + Chr$(15) 'Chr$(10) + Chr$(27) + Chr$(15)
'
'   b = InStr(X, A$, f)
'
'   If b = 0 Then
'     GoTo 555
'   End If
'
'
'
'
'
'
'
'
'
'   d = mID(A$, b + 2, 180)
'   kod = Left(d, 12)
'
'
'
'
'
'
'
'   per = mID$(d, 30, 38)
'   pos = mID$(d, 76, 10)
'
'   MON = mID$(d, 90, 10)
'   TIMH = Trim(mID$(d, 104, 10))
'   TIMH = Replace(TIMH, ",", ".")
'
'   'List1.AddItem d
'
'
'
'rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'   If rSQL.EOF Then
' '      On Error Resume Next
'      Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,XTI,MON) VALUES " _
    '      & "('" + Trim(kod) + "','" + Replace(Left(Trim(per), 35), "'", "~") + "',2," + TIMH + ",'" + Trim(MON) + "')", RECS
'      If RECS = 0 Then
'         MsgBox "Δεν υπάρχει το είδος " + kod
'      Else
'         MsgBox "Ανοίχθηκε το είδος " + kod
'      End If
'   End If
'
'rSQL.Close
'
'   X = b + 3
'
'
'  k = k + 1
'  If k > f_SeiresTimologioy Then Exit Do
'     Grid1.TextMatrix(k, f_k) = kod
'     Grid1.TextMatrix(k, f_o) = per
'     Grid1.TextMatrix(k, f_p) = pos
'     Grid1.TextMatrix(k, F_T) = TIMH
'     Grid1.TextMatrix(k, f_f) = 2
'     Grid1.row = k
'     FindSynolo
'
'555
'Loop
'Close #1
'
'
' Else
' ' μπουγουδης
'Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
'Set R = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM=" + Left(Text1.Text, 5))
'If R.RecordCount = 0 Then
'   MsgBox "δεν υπάρχουν εγγραφές"
'   Exit Sub
'End If
'
'
'R.MoveFirst
'k = 0
'Do While Not R.EOF
'  k = k + 1
'  If k > f_SeiresTimologioy Then Exit Do
'  If IsNull(R("BARCODE")) Then
'     Grid1.TextMatrix(k, f_k) = " "
'  Else
'     Grid1.TextMatrix(k, f_k) = R("BARCODE")
'  End If
'  R.MoveNext
'Loop
'R.Close
'db.Close
'
'End If
'
'
'
'End Sub

Sub loisir_import()

        '<EhHeader>
        On Error GoTo loisir_import_Err

        '</EhHeader>

        Dim DB As DAO.Database

        Dim R As DAO.Recordset

        Dim k

        Dim a$, B, X

        Dim D

        Dim rSQL As New ADODB.Recordset

        Dim ANS

100     ANS = MsgBox("ΜΕ EXCEL", vbYesNo)

102     If ANS = vbYes Then
104         LOISIR2_import

            Exit Sub

        End If

106     If Len(Dir("c:\TIMOLOGIA", vbDirectory)) > 2 Then
            ' LOISIR

108         Open "c:\timologia\" + Text1.Text + ".asc" For Binary Access Read As #1

            'Do While Not EOF(1)
            ' Line Input #1, A$
            ' Exit Do
            'Loop
            'Close #1
            'Exit Sub

110         X = 1

            Dim kod, per, MON, POS, TIMH

            Dim F As String, RECS As Integer

112         k = 0

114         Do While Not EOF(1)    'True

                On Error Resume Next

116             Line Input #1, a$

118             If EOF(1) Then Exit Do

120             F = Chr$(10) + Chr$(27) + Chr$(15)

                '   B = InStr(x, A$, f)
                '   If B = 0 Then
                '    GoTo 555
                '     Exit Do
                ' End If

122             X = 1

                ' f = Chr$(27) + Chr$(15) 'Chr$(10) + Chr$(27) + Chr$(15)
124             F = Chr$(15)  'Chr$(10) + Chr$(27) + Chr$(15)

126             B = InStr(X, a$, F)

128             If B = 0 Then
130                 GoTo OLD555
                End If

132             D = mID(a$, B + 2, 180)
134             kod = mID(D, 3, 11)    ' Left(d, 12)

136             per = mID$(D, 30, 38)
138             POS = mID$(D, 74, 10)

140             MON = mID$(D, 85, 10)
142             TIMH = Trim(mID$(D, 92, 10))
144             TIMH = Replace(TIMH, ",", ".")

                'List1.AddItem d

146             rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

148             If rSQL.EOF Then
                    '      On Error Resume Next
150                 Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,XTI,MON) VALUES " & "('" + Trim(kod) + "','" + Replace(Left(Trim(per), 35), "'", "~") + "',2," + TIMH + ",'" + Trim(MON) + "')", RECS

152                 If RECS = 0 Then
154                     MsgBox "Δεν υπάρχει το είδος " + kod
                    Else
156                     MsgBox "Ανοίχθηκε το είδος " + kod
                    End If
                End If

158             rSQL.Close

160             X = B + 3

162             k = k + 1

164             If k > f_SeiresTimologioy Then Exit Do
166             Grid1.TextMatrix(k, f_k) = kod
168             Grid1.TextMatrix(k, f_o) = per
170             Grid1.TextMatrix(k, f_p) = POS
172             Grid1.TextMatrix(k, F_T) = TIMH
174             Grid1.TextMatrix(k, f_f) = 2
176             Grid1.row = k
178             FindSynolo
OLD555:

            Loop

180         Close #1

        Else
            ' μπουγουδης
182         Set DB = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
184         Set R = DB.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM=" + Left(Text1.Text, 5))

186         If R.RecordCount = 0 Then
188             MsgBox "δεν υπάρχουν εγγραφές"

                Exit Sub

            End If

190         R.MoveFirst
192         k = 0

194         Do While Not R.EOF
196             k = k + 1

198             If k > f_SeiresTimologioy Then Exit Do
200             If IsNull(R("BARCODE")) Then
202                 Grid1.TextMatrix(k, f_k) = " "
                Else
204                 Grid1.TextMatrix(k, f_k) = R("BARCODE")
                End If

206             R.MoveNext
            Loop

208         R.Close
210         DB.Close

        End If

        '<EhFooter>
        Exit Sub

loisir_import_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.loisir_import " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.loisir_import " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub WESTNET_import()

        '<EhHeader>
        On Error GoTo WESTNET_import_Err

        '</EhHeader>

        Dim DB As DAO.Database

        Dim R As DAO.Recordset

        Dim r2 As New ADODB.Recordset

        Dim k

        Dim a$, B, X

        Dim D

        Dim rSQL As New ADODB.Recordset

        Dim ANS

        Dim MONO As String

100     If Len(Dir("c:\WESTNET\*.*", vbNormal)) < 3 Then
102         MsgBox "δεν υπάρχει αρχείο για τιμολόγηση"

            Exit Sub
        
        End If

        ' R2.Open "SELECT COUNT(*) FROM EID WHERE KOD='00106983'", Gdb, adOpenDynamic, adLockOptimistic

104     POLAGOR.Text = POLAGOR.List(0)
106     POLAGOR_LostFocus

108     PARAS.Text = PARAS.List(7)

110     Text2(0).Text = "0064"
112     fbuff = "r"
        ' Text2_KeyUp 2, 13, 1
        ' GridPelaton.SetFocus 'διαλεξε πελάτη
        ' SendKeys "{ENTER}{ENTER}"

114     Open "c:\WESTNET\" + Dir("c:\WESTNET\*.*", vbNormal) For Input As #1
116     X = 1

        Dim kod, per, MON, POS, TIMH

        Dim F As String, RECS As Integer

118     k = 0

120     Do While Not EOF(1)    'True

            On Error Resume Next

122         Line Input #1, a$

124         If EOF(1) Then

                Exit Do

            End If

126         Line Input #1, a$
128         Line Input #1, a$
130         Line Input #1, a$
132         Line Input #1, a$
134         Line Input #1, a$
136         Line Input #1, a$
138         Line Input #1, a$
140         Line Input #1, a$
142         Line Input #1, a$
144         Line Input #1, a$  'ΔΙΑΒΑΖΕΙ ΔΙΕΥΘΥΝΣΗ
146         Text1.Text = mID$(a$, 55, 6)

            'dtpicker1.Value = Replace(mID$(a$, 87, 26), ".", "/")

148         Line Input #1, a$
            'parat.Text = mID$(a$, 62, 14)

150         Line Input #1, a$
152         DTPicker1.Value = Replace(mID$(a$, 87, 26), ".", "/")
154         Text1.Text = mID$(a$, 55, 8)
   
156         Line Input #1, a$

158         parat.Text = mID$(a$, 60, 14)

160         Line Input #1, a$
162         Line Input #1, a$
164         Line Input #1, a$
166         Line Input #1, a$

168         Line Input #1, a$
170         Line Input #1, a$
172         Line Input #1, a$
174         Line Input #1, a$
176         Line Input #1, a$

178         Line Input #1, a$
180         Line Input #1, a$
182         Line Input #1, a$

184         Line Input #1, a$
186         Line Input #1, a$

188         Line Input #1, a$
            ' Line Input #1, a$

            Dim N As Integer

190         For k = 1 To 20

192             Line Input #1, a$    '1o EIDOS

                '    If Len(Trim(Left(kod, 5))) = 0 Then
                '      Exit For
                '  End If

194             If Len(Trim(Left(a$, 5))) = 0 Then    'EISFORA ANAKYKLOSHS 'H KENO
196                 If Len(Trim(Left(a$, 18))) = 0 Then    'KENO

                        'Line Input #1, A$ '1o EIDOS
                        Exit For

                    Else
198                     Line Input #1, a$    '1o EIDOS
                    End If
                End If

200             kod = Trim(mID$(a$, 4, 8))

202             If Left(kod, 2) = "00" Then
204                 kod = Trim(mID$(a$, 5, 8))
                End If
            
206             per = mID$(a$, 15, 38)
208             POS = Val(Replace(mID(a$, 67, 5), ",", "."))

210             POS = Val(Replace(mID(a$, 63, 6), ",", "."))
212             TIMH = Val(Replace(mID(a$, 75, 9), ",", "."))

214             rSQL.Close
216             r2.Close

218             rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenDynamic, adLockOptimistic

220             r2.Open "SELECT COUNT(*) FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenDynamic, adLockOptimistic

222             If r2(0) = 0 Then
224                 Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,XTI,MON) VALUES " & "('" + Trim(kod) + "','" + Replace(Left(Trim(per), 35), "'", "~") + "',2," + str(TIMH) + ",'" + Trim(MON) + "')", RECS

226                 If RECS = 0 Then
228                     MsgBox "Δεν υπάρχει το είδος " + kod
                    Else
230                     MsgBox "Ανοίχθηκε το είδος " + kod
                    End If
                End If

232             If k > f_SeiresTimologioy Then

                    Exit For

                End If

234             Grid1.TextMatrix(k, f_k) = kod
236             Grid1.TextMatrix(k, f_o) = per
238             Grid1.TextMatrix(k, f_p) = POS
240             Grid1.TextMatrix(k, f_e) = 0
242             Grid1.TextMatrix(k, F_T) = TIMH
244             Grid1.TextMatrix(k, f_f) = 2
246             Grid1.row = k
248             FindSynolo
            Next

250         rSQL.Close

252         Line Input #1, a$
254         Line Input #1, a$
256         Line Input #1, a$

258         Line Input #1, a$
260         Line Input #1, a$
262         Line Input #1, a$

264         Line Input #1, a$
266         Line Input #1, a$
268         Line Input #1, a$

270         Line Input #1, a$
272         Line Input #1, a$
274         Line Input #1, a$
276         Line Input #1, a$
278         Line Input #1, a$
280         Line Input #1, a$

282         Line Input #1, a$
284         Line Input #1, a$
286         Line Input #1, a$

            'Line Input #1, A$

288         For k = 1 To 9

290             Text5.Text = "*" + LTrim((Replace(mID(a$, 60, 20), ",", ".")))

292             If Val(mID(Text5.Text, 2, 10)) > 0 Then

                    Exit For

                End If

294             Line Input #1, a$
            Next

            Exit Do

        Loop

296     Close #1

298     Gdb.Execute "update TIM SET ART='1',PARAT='" + Text1.Text + "' WHERE ATIM LIKE 'E%" + Right(parat, 5) + "%'", k

300     Kill "c:\WESTNET\" + Dir("c:\WESTNET\*.*", vbNormal)

        '<EhFooter>
        Exit Sub

WESTNET_import_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.WESTNET_import " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.WESTNET_import " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub LOISIR2_import()

        ' ME EXCEL
        '<EhHeader>
        On Error GoTo LOISIR2_import_Err

        '</EhHeader>
        Dim xl As New excel.Application

        Dim xlsheet As excel.Worksheet

        Dim xlsheet3 As excel.Worksheet

        Dim xlwbook As excel.workbook

100     Set xlwbook = xl.Workbooks.Open("c:\TIMOLOGIA\" + Text1.Text + ".XLS")
102     Set xlsheet = xlwbook.Sheets.Item(1)

        Dim k

        Dim a$, B, X

        Dim D

        Dim FPA As Integer

        Dim rSQL As New ADODB.Recordset

        Dim kod, per, MON, POS, TIMH, TIMOL

        Dim F As String, RECS As Integer

104     k = 1

106     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

108         If IsNull(xlsheet.cells(k, 1)) Then

                Exit Do

            End If

110         If IsEmpty(xlsheet.cells(k, 1)) Then

                Exit Do

            End If

            On Error Resume Next

112         kod = xlsheet.cells(k, 1)

114         POS = xlsheet.cells(k, 3)
116         TIMH = xlsheet.cells(k, 4)

118         rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

120         If rSQL.EOF Then
122             MsgBox "Δεν υπάρχει το είδος " + kod
            Else
124             kod = rSQL(0)
                'rSQL.Close

                'rSQL.Open "SELECT ONO,XTI,FPA FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
126             per = rSQL("ONO")
                ' TIMH = rSQL("XTI")
128             FPA = rSQL("FPA")
                'TIMH = Round(rSQL("XTI"), 2) '  / ((100 + g_Fpa(FPA)) / 100)

            End If

130         rSQL.Close

132         If k > f_SeiresTimologioy Then Exit Do
134         Grid1.TextMatrix(k, f_k) = kod
136         Grid1.TextMatrix(k, f_o) = per
138         Grid1.TextMatrix(k, f_p) = POS
140         Grid1.TextMatrix(k, F_T) = TIMH
142         Grid1.TextMatrix(k, f_f) = FPA
144         Grid1.row = k
146         FindSynolo

148         k = k + 1

            'If EOF(1) Then Exit Do

        Loop

        'Close #1

150     Set xlsheet = Nothing
152     Set xlwbook = Nothing

        '<EhFooter>
        Exit Sub

LOISIR2_import_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.LOISIR2_import " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.LOISIR2_import " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub SOKRATIS_import()

        ' ME EXCEL
        '<EhHeader>
        On Error GoTo SOKRATIS_import_Err

        '</EhHeader>
        Dim xl As New excel.Application

        Dim xlsheet As excel.Worksheet

        Dim xlsheet3 As excel.Worksheet

        Dim xlwbook As excel.workbook

100     Set xlwbook = xl.Workbooks.Open("c:\SOKRATIS\SOK.XLS")
102     Set xlsheet = xlwbook.Sheets.Item(1)

        Dim k

        Dim a$, B, X

        Dim D

        Dim FPA As Integer

        Dim rSQL As New ADODB.Recordset

        Dim kod, per, MON, POS, TIMH, TIMOL

        Dim F As String, RECS As Integer

104     k = 1

        Dim CH1, CH2, CH3

        Dim XTI, lti

        Dim s As String

        Dim nn As Long

        Dim LL2 As Long

106     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

108         If IsNull(xlsheet.cells(k, 1)) Then

                Exit Do

            End If

110         If IsEmpty(xlsheet.cells(k, 1)) Then

                Exit Do

            End If

            On Error Resume Next

112         kod = LTrim(str(xlsheet.cells(k + 1, 1)))    '
114         CH1 = xlsheet.cells(k + 1, 2)    ' PROSFORA
116         CH2 = xlsheet.cells(k + 1, 3)    ' XEIMONIATIKO
118         CH3 = xlsheet.cells(k + 1, 6)    'PROM
120         XTI = xlsheet.cells(k + 1, 4)    ' , ",", ".")
122         lti = xlsheet.cells(k + 1, 5)    ', ",", ".")

124         rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

126         If rSQL.EOF Then
128             s = "INSERT INTO EID (KOD,CH1,CH2,CH3,XTI,LTI5,KATHGORIA,FPA) VALUES ("

                On Error GoTo SOKRATIS_import_Err

130             s = s + "'" + kod + "','" + str(CH1) + "','" + str(CH2) + "','" + str(CH3) + "'," + str(XTI) + "," + str(lti) + ",1,2)"

132             Gdb.Execute s, nn

            Else

134             s = "UPDATE  EID  SET XTI=" + str(XTI) + ",LTI5=" + str(lti) + "  WHERE KOD='" + Trim(kod) + "'"

136             Gdb.Execute s, nn

            End If

138         rSQL.Close

140         If k > f_SeiresTimologioy Then Exit Do
142         XROMATA(k, 0, 0) = -0.01     ' ΣΗΜΑΔΙ ΟΤΙ ΧΡΗΣΙΜΟΠΟΙΗΣΕ ΤΟ F6
144         POS = 0
146         gm_str(k) = ""

148         For LL2 = 7 To 19
150             XROMATA(k, LL2 - 6, 1) = xlsheet.cells(k + 1, LL2)    '
152             gm_str(k) = gm_str(k) + str(XROMATA(k, LL2 - 6, 1)) + ";"
                ' pos = pos + xlsheet.cells(k + 1, LL2)
            Next

154         Grid1.TextMatrix(k, f_k) = kod
156         Grid1.TextMatrix(k, f_o) = ""
158         Grid1.TextMatrix(k, f_p) = POS
160         Grid1.TextMatrix(k, F_T) = XTI
162         Grid1.TextMatrix(k, f_f) = FPA

            'For R = 1 To Grid1.ROWS - 1
            '  For C = 1 To Grid1.ColS - 1
            ' Grid1.TextMatrix(R, C) = Replace(Grid1.TextMatrix(R, C), ";", " ")
            '  gm_str(Val(Label1.Caption)) = gm_str(Val(Label1.Caption)) + Grid1.TextMatrix(R, C) + ";"
            'Next
            'Next

164         Grid1.row = k
166         FindSynolo

168         k = k + 1

            'If EOF(1) Then Exit Do

        Loop

        'Close #1

170     Set xlsheet = Nothing
172     Set xlwbook = Nothing
174     xl.Quit

        '<EhFooter>
        Exit Sub

SOKRATIS_import_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.SOKRATIS_import " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SOKRATIS_import " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub EXCEL28028_IMPORT()
Me.Caption = "28028"
        CommonDialog1.ShowOpen
        Dim mF As String
        mF = CommonDialog1.FILENAME
        If Len(mF) < 2 Then
            MsgBox "δεν διαλεξατε αρχείο"
            Exit Sub
        End If

        ' ME EXCEL
        '<EhHeader>
        On Error GoTo SOKRATIS_import_Err

        '</EhHeader>
        Dim xl As New excel.Application

        Dim xlsheet As excel.Worksheet

        Dim xlsheet3 As excel.Worksheet

        Dim xlwbook As excel.workbook

100     Set xlwbook = xl.Workbooks.Open(mF) ' "c:\MERCVB\TIMOL.XLSX")
102     Set xlsheet = xlwbook.Sheets.Item(1)

        Dim k

        Dim a$, B, X

        Dim D

        Dim FPA As Integer

        Dim rSQL As New ADODB.Recordset

        Dim kod, per, MON, POS, TIMH, TIMOL

        Dim F As String, RECS As Integer

104     k = 2

        Dim CH1, CH2, CH3

        Dim XTI, lti

        Dim s As String

        Dim nn As Long

        Dim LL2 As Long

        Dim ono As String

        'Dim FPA As Integer
        Dim TIMM As Single
        Dim EKPT As Single
        Dim lti5 As Single
        
        Dim mHME As Date
        Dim MATIM As Long
        Dim MAPOT As String
        
        mHME = xlsheet.cells(k + 1, 4)
        MATIM = xlsheet.cells(k + 1, 5)
        Text1.Text = MATIM
        DTPicker1.Value = mHME
        MAPOT = xlsheet.cells(k + 1, 15)
        If Val(MAPOT) = 4 Then
           APOTH1.Text = APOTH1.List(1)
        Else
           APOTH1.Text = APOTH1.List(0)
        
        End If
        
106     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

108         If IsNull(xlsheet.cells(k, 1)) Then

                Exit Do

            End If

110         If IsEmpty(xlsheet.cells(k, 1)) Then

                Exit Do

            End If

            On Error Resume Next

112         kod = xlsheet.cells(k + 1, 6) ' 1)  ' LTrim(str(xlsheet.cells(k + 1, 1)))    '

114         ono = xlsheet.cells(k + 1, 7) ' 2)  ' PERIGRAFH
116         POS = xlsheet.cells(k + 1, 9)    ' posothta
118         TIMM = xlsheet.cells(k + 1, 10) ' timh
            lti5 = xlsheet.cells(k + 1, 8) ' LIT5  timh

            Dim ekpt2(4) As Single
            Dim ll As Integer
            For ll = 1 To 4
                If IsNull(xlsheet.cells(k + 1, 10 + ll)) Then
                     ekpt2(ll) = 1
                Else
                     ekpt2(ll) = (100 - xlsheet.cells(k + 1, 10 + ll)) / 100
                End If
                TIMM = TIMM * ekpt2(ll)
                
            Next ll
              

120         If POS <> 0 Then
                'CH2 = xlsheet.cells(k + 1, 3)    ' XEIMONIATIKO
                'CH3 = xlsheet.cells(k + 1, 6)    'PROM
                'XTI = xlsheet.cells(k + 1, 4)    ' , ",", ".")
                'LTI = xlsheet.cells(k + 1, 5)    ', ",", ".")

122             rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

124             If rSQL.EOF Then
126                ' If F_EXCEL_IMPORT = 28028 Then
128                     Gdb.Execute "insert into EID (KOD,ONO,FPA,MON,XTI,LTI5) VALUES ('" + kod + "','" + ono + "',2,'TEM'," + Replace(Format(TIMM, "####.00"), ",", ".") + "," + Replace(Format(lti5, "####.00"), ",", ".") + ")"
                        '  MsgBox "DENYPARXEI O KVDIKOS " + kod
130                     xlsheet.cells(k, 7) = "δημιουργηθηκε ο κωδικός ME KAT ΦΠΑ=1 & ΜΟΝ=ΤΕΜ"
                   ' Else
132                 '    MsgBox "δεν βρέθηκε ο κωδικός " + Trim(kod)
134                  '   xlsheet.cells(k, 7) = "Δεν βρέθηκε ο κωδικός " + kod

                    'End If
                    
                Else
                
                   If gVal(lti5) > 0 Then
                        Gdb.Execute "UPDATE EID SET  LTI5=" + Replace(Format(lti5, "####.00"), ",", ".") + " WHERE KOD='" + kod + "'"
                    
                   
                   End If
                   
                   
                   If gVal(TIMM) > 0 Then
                   
                           Gdb.Execute "UPDATE EID SET XTI=" + Replace(Format(TIMM, "####.00"), ",", ".") + " WHERE KOD='" + kod + "'"
                    
                   End If
                   
                
                
                   ' Gdb.Execute "UPDATE EID SET XTI=" + Replace(Format(TIMM, "####.00"), ",", ".") + ", LTI5=" + Replace(Format(lti5, "####.00"), ",", ".") + " WHERE KOD='" + kod + "'"
                    
                End If
            
136             ono = rSQL("ONO")
138             FPA = 2 'rSQL("FPA")
140             MON = rSQL("MON")
                
                'F_EXCEL_IMPORT = Val(FINDPARAMETROI(3, "PAR1", "F_EXCEL_IMPORT", "110", "Γ.46.Excel Import=100 Excel+τιμές=110 Excel+τιμές+Αυτόματο ανοιγμα κωδικου=111")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
142             If F_EXCEL_IMPORT < 110 Then
144                 TIMM = rSQL("xti")   ' xlsheet.cells(k + 1, 3)
                End If
146             EKPT = 0 '  xlsheet.cells(k + 1, 4)

                ' End If

148             rSQL.Close

150             If k > f_SeiresTimologioy Then Exit Do

152             Grid1.TextMatrix(k, f_k) = kod
154             Grid1.TextMatrix(k, f_o) = ono
156             Grid1.TextMatrix(k, f_p) = POS
158             Grid1.TextMatrix(k, F_T) = TIMM
160             Grid1.TextMatrix(k, f_e) = EKPT
162             Grid1.TextMatrix(k, f_f) = FPA
164             Grid1.TextMatrix(k, F_M) = MON
            
166             Grid1.row = k
168             FindSynolo

            End If

170         k = k + 1

        Loop

172     Set xlsheet = Nothing
174     Set xlwbook = Nothing
176     xl.Quit

        '<EhFooter>
        Exit Sub

SOKRATIS_import_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.SOKRATIS_import " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SOKRATIS_import " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


Sub EXCEL_IMPORT()

        CommonDialog1.ShowOpen
        Dim mF As String
        mF = CommonDialog1.FILENAME
        If Len(mF) < 2 Then
            MsgBox "δεν διαλεξατε αρχείο"
            Exit Sub
        End If

        ' ME EXCEL
        '<EhHeader>
        On Error GoTo SOKRATIS_import_Err

        '</EhHeader>
        Dim xl As New excel.Application

        Dim xlsheet As excel.Worksheet

        Dim xlsheet3 As excel.Worksheet

        Dim xlwbook As excel.workbook

100     Set xlwbook = xl.Workbooks.Open(mF) ' "c:\MERCVB\TIMOL.XLSX")
102     Set xlsheet = xlwbook.Sheets.Item(1)

        Dim k

        Dim a$, B, X

        Dim D

        Dim FPA As Integer

        Dim rSQL As New ADODB.Recordset

        Dim kod, per, MON, POS, TIMH, TIMOL

        Dim F As String, RECS As Integer

104     k = 1

        Dim CH1, CH2, CH3

        Dim XTI, lti

        Dim s As String

        Dim nn As Long

        Dim LL2 As Long

        Dim ono As String

        'Dim FPA As Integer
        Dim TIMM As Single
        Dim EKPT As Single
        
106     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

108         If IsNull(xlsheet.cells(k, 1)) Then

                Exit Do

            End If

110         If IsEmpty(xlsheet.cells(k, 1)) Then

                Exit Do

            End If

            On Error Resume Next

112         kod = xlsheet.cells(k + 1, 1)  ' LTrim(str(xlsheet.cells(k + 1, 1)))    '

114         ono = xlsheet.cells(k + 1, 2)  ' PERIGRAFH
116         POS = xlsheet.cells(k + 1, 3)    ' posothta
118         TIMM = xlsheet.cells(k + 1, 4) ' timh
          
120         If POS <> 0 Then
                'CH2 = xlsheet.cells(k + 1, 3)    ' XEIMONIATIKO
                'CH3 = xlsheet.cells(k + 1, 6)    'PROM
                'XTI = xlsheet.cells(k + 1, 4)    ' , ",", ".")
                'LTI = xlsheet.cells(k + 1, 5)    ', ",", ".")

122             rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

124             If rSQL.EOF Then
126                 If F_EXCEL_IMPORT = 111 Or F_EXCEL_IMPORT = 101 Then
128                     Gdb.Execute "insert into EID (KOD,ONO,FPA,MON) VALUES ('" + kod + "','" + ono + "',1,'TEM')"
                        '  MsgBox "DENYPARXEI O KVDIKOS " + kod
130                     xlsheet.cells(k, 7) = "δημιουργηθηκε ο κωδικός ME KAT ΦΠΑ=1 & ΜΟΝ=ΤΕΜ"
                    Else
132                     MsgBox "δεν βρέθηκε ο κωδικός " + Trim(kod)
134                     xlsheet.cells(k, 7) = "Δεν βρέθηκε ο κωδικός " + kod
                    End If
                End If
            
136             ono = rSQL("ONO")
138             FPA = rSQL("FPA")
140             MON = rSQL("MON")
                
                'F_EXCEL_IMPORT = Val(FINDPARAMETROI(3, "PAR1", "F_EXCEL_IMPORT", "110", "Γ.46.Excel Import=100 Excel+τιμές=110 Excel+τιμές+Αυτόματο ανοιγμα κωδικου=111")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
142             If F_EXCEL_IMPORT < 110 Then
144                 TIMM = rSQL("xti")   ' xlsheet.cells(k + 1, 3)
                End If
146             EKPT = 0 '  xlsheet.cells(k + 1, 4)

                ' End If

148             rSQL.Close

150             If k > f_SeiresTimologioy Then Exit Do

152             Grid1.TextMatrix(k, f_k) = kod
154             Grid1.TextMatrix(k, f_o) = ono
156             Grid1.TextMatrix(k, f_p) = POS
158             Grid1.TextMatrix(k, F_T) = TIMM
160             Grid1.TextMatrix(k, f_e) = EKPT
162             Grid1.TextMatrix(k, f_f) = FPA
164             Grid1.TextMatrix(k, F_M) = MON
            
166             Grid1.row = k
168             FindSynolo

            End If

170         k = k + 1

        Loop

172     Set xlsheet = Nothing
174     Set xlwbook = Nothing
176     xl.Quit

        '<EhFooter>
        Exit Sub

SOKRATIS_import_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.SOKRATIS_import " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SOKRATIS_import " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub




Sub PAYLOS_import()

        '<EhHeader>
        On Error GoTo PAYLOS_import_Err

        '</EhHeader>

        Dim DB As DAO.Database

        Dim R As DAO.Recordset

        Dim k33

        Dim a$, B, X

        Dim D

        Dim FPA As Integer
    
        Dim YPARXEI, m

        Dim rSQL As New ADODB.Recordset

100     Open "C:\EGGTIM.TXT" For Input As #1

102     X = 1

        Dim kod, per, MON, POS, TIMH, TIMOL

        Dim F As String, RECS As Integer

104     k33 = 0
106     f_use_Sysk = 1

108     Do While Not EOF(1)    'True

            On Error Resume Next

110         Line Input #1, a$

112         X = 1
114         kod = Trim(mID(a$, 2, 10))

116         POS = Val(mID$(a$, 15, 7))

            'TIMOL = Val(mID$(a$, 20, 6))
118         If Val(kod) > 0 Then
120             rSQL.Open "SELECT KOD FROM BARCODES WHERE ERG='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

122             If rSQL.EOF Then
124                 MsgBox "Δεν υπάρχει το είδος " + kod
                Else
126                 kod = rSQL(0)
128                 rSQL.Close
130                 rSQL.Open "SELECT ONO,XTI,FPA FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
132                 per = rSQL("ONO")
134                 TIMH = rSQL("XTI")
136                 FPA = rSQL("FPA")
138                 TIMH = Round(rSQL("XTI"), 2)    '  / ((100 + g_Fpa(FPA)) / 100)
                End If

140             rSQL.Close
                'ελεγχος αν υπαρχει να προσθετει ποσοτητες και να μην ανοιγει καινουριο
142             YPARXEI = 0

144             For m = 1 To k33

146                 If Grid1.TextMatrix(m, f_k) = kod Then
148                     YPARXEI = 1

                        Exit For

                    End If

                Next
            
150             If YPARXEI = 1 Then
152                 Grid1.TextMatrix(m, f_k) = kod
154                 Grid1.TextMatrix(m, f_o) = per
156                 Grid1.TextMatrix(m, f_p) = POS + Val(Grid1.TextMatrix(m, f_p))
158                 Grid1.TextMatrix(m, F_T) = TIMH
160                 Grid1.TextMatrix(m, f_f) = FPA
162                 FindSynolo
                Else
164                 k33 = k33 + 1

166                 If k33 > f_SeiresTimologioy Then Exit Do
168                 Grid1.TextMatrix(k33, f_k) = kod
170                 Grid1.TextMatrix(k33, f_o) = per
172                 Grid1.TextMatrix(k33, f_p) = POS
174                 Grid1.TextMatrix(k33, F_T) = TIMH
176                 Grid1.TextMatrix(k33, f_f) = FPA
178                 Grid1.row = k33
                    'FindSynolo
                End If
            
            End If    ' TIMOL=TEXT1

180         If EOF(1) Then Exit Do
        Loop

182     FindSynolo
184     Close #1

        '<EhFooter>
        Exit Sub

PAYLOS_import_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.PAYLOS_import " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.PAYLOS_import " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub SYMBOL_import()

        '<EhHeader>
        On Error GoTo SYMBOL_import_Err

        '</EhHeader>

        Dim DB As DAO.Database

        Dim R As DAO.Recordset

        Dim k

        Dim a$, B, X

        Dim D

        Dim FPA As Integer

        Dim rSQL As New ADODB.Recordset

100     Open "C:\APOG.TXT" For Input As #1

102     X = 1

        Dim kod, per, MON, POS, TIMH, TIMOL

        Dim F As String, RECS As Integer

104     k = 0

106     Do While Not EOF(1)    'True

            On Error Resume Next

108         Line Input #1, a$

110         X = 1
112         kod = mID(a$, 1, 13)

114         POS = Val(mID$(a$, 15, 5))

116         TIMOL = Val(mID$(a$, 20, 6))

118         If TIMOL = Val(Text1.Text) And Val(kod) > 0 Then
120             rSQL.Open "SELECT KOD FROM BARCODES WHERE ERG='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

122             If rSQL.EOF Then
124                 MsgBox "Δεν υπάρχει το είδος " + kod
                Else
126                 kod = rSQL(0)
128                 rSQL.Close

130                 rSQL.Open "SELECT ONO,XTI,FPA FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
132                 per = rSQL("ONO")
134                 TIMH = rSQL("XTI")
136                 FPA = rSQL("FPA")
138                 TIMH = Round(rSQL("XTI"), 2)    '  / ((100 + g_Fpa(FPA)) / 100)

                End If

140             rSQL.Close

142             k = k + 1

144             If k > f_SeiresTimologioy Then Exit Do
146             Grid1.TextMatrix(k, f_k) = kod
148             Grid1.TextMatrix(k, f_o) = per
150             Grid1.TextMatrix(k, f_p) = POS
152             Grid1.TextMatrix(k, F_T) = TIMH
154             Grid1.TextMatrix(k, f_f) = FPA
156             Grid1.row = k
158             FindSynolo

            End If    ' TIMOL=TEXT1

160         If EOF(1) Then Exit Do

        Loop

162     Close #1

        '<EhFooter>
        Exit Sub

SYMBOL_import_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.SYMBOL_import " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SYMBOL_import " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim a

100     a = toascii("C:\MERCVB\F31.TXT", "C:\LIBRA\HOST.TXT")

102     MsgBox "Τέλος Αποστολής"

        On Error Resume Next

104     a = Shell("C:\LIBRA\LIBRA.EXE 1", vbMaximizedFocus)

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        'ελεγχω αν εχει το δικαιωμα να ανοιξει την καρτελλα
        Dim r1 As New ADODB.Recordset, M_ENABLE As Boolean
        r1.Open "SELECT * FROM ACTIVE_MENUS WHERE ID=5 and  USER_ID=" + str(gUserId), Gdb, adOpenDynamic, adLockOptimistic
        If r1("ENABLED") = 1 Then
            M_ENABLE = True
        Else
            M_ENABLE = False
            Exit Sub
        End If

        '</EhHeader>
        Dim fpelat4 As New pelat4

        'If Index = 0 Then
100     fpelat4.SHOW
102     fpelat4.Text2(0).Text = Text2(0).Text

104     If POLAGOR.ListIndex = 0 Then
106         fpelat4.PELPROM.ListIndex = 1
        Else
        
108         If f_KENTRA_ADYNATISMATOS = 1 Then
110             If kentroAdyn.ListIndex = 1 Then
112                 fpelat4.kentroAdyn.ListIndex = 1
                End If
            End If
        
114         fpelat4.PELPROM.ListIndex = 0
            'fpelat4.Option1(1).value = True
        End If

116     fpelat4.data1.ConnectionString = gConnect
118     fpelat4.Refresh
120     gApoMenu = False

        'pelat4.Text2(0).SetFocus

        'fpelat4.SHOW
122     fpelat4.Text2(0).SetFocus
124     SendKeys "+{ENTER}"

        Dim DUM

126     DUM = fpelat4.KARTELLA(gDir)
        'End If
128     Set fpelat4 = Nothing

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

        'ΔΗΜΙΟΥΡΓΕΙ ΑΥΤΟΜΑΤΑ ΠΑΡΑΣΤΑΤΙΚΟ ΤΠΥ,ΑΠΥ,ΔΑ
        'ΠΑΙΡΝΟΝΤΑΣ ΤΑ ΣΤΟΙΧΕΙΑ ΑΠΟ ΤΟ ΕΙΔΟΣ ΠΟΥ ΕΧΟΥΜΕ ΠΕΡΑΣΕΙ ΜΕΣΑ ΑΠΟ ΤΗΝ ΦΟΡΜΑ ΠΑΡΑΛΑΒΗ
        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>

        Dim ko As String, ajia As String, N As Integer, TYPOS As Integer

100     ko = InputBox("δωσε το barcode του είδους ", , "")
102     ajia = InputBox("δωσε αξία ", , "")
104     TYPOS = Val(InputBox("Αποδ.ΠΥ=6 Δ.Α=9 ΤΠΥ=16 ", , ""))

106     If TYPOS = 6 Or TYPOS = 9 Or TYPOS = 16 Then
            ' OK
        Else
108         MsgBox "Λάθος τύπος παραστατικού"

            Exit Sub

        End If

110     POLAGOR.Text = POLAGOR.List(1)

112     POLAGOR_LostFocus

        Dim R As New ADODB.Recordset, sql As String

        'ΒΡΙΣΚΩ ΤΟ ΕΙΔΟΣ
114     R.Open "SELECT * FROM EID WHERE LEFT(KOD,6)='" + Left(ko, 6) + "'", Gdb, adOpenDynamic, adLockOptimistic

116     If R.EOF Then
118         MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ Ο ΚΩΔΙΚΟΣ"

            Exit Sub

        End If

120     SendKeys ""

122     Text2(0).Text = Trim(R("CH6"))    ' ΚΩΔΙΚΟΣ ΠΕΛΑΤΗ

124     If Left(R("CH6"), 4) = "1112" Then

            'ΕΝΗΜΕΡΩΝΩ  ΤΟΝ ΠΕΛΑΤΗ ΛΙΑΝΙΚΗΣ 1112 ΜΕ ΤΑ ΣΤΟΙΧΕΙΑ ΤΟΥ
126         If TYPOS = 16 Then    ' tpy
128             MsgBox "δεν μπορεις να κόψεις ΤΠΥ χωρίς ΑΦΜ"
            Else
                'sql = "UPDATE PEL SET EPO='" + R("CH2") + "',DIE='" + R("CH3") + "',EPA='" + R("CH4") + "',THL='" + R("CH5") + "' WHERE EIDOS='e' AND KOD='" + Trim(R("CH6")) + "'"
                'Gdb.Execute sql, N

                '      sql = "UPDATE PEL SET EPO='" + Trim(R("CH2")) + "',DIE='" + Trim(R("CH3")) + "',EPA='" + Trim(R("CH4")) + "',THL='" + Left(Trim(R("CH5")), 10) + "' WHERE EIDOS='e' AND KOD='" + Trim(R("CH6")) + "'"
                '     Gdb.Execute sql, N

130             sql = "UPDATE PEL SET EPO='" + Trim(R("CH2")) + "',DIE='" + Trim(R("CH3")) + "',EPA='" + Trim(R("CH4")) + "',THL='" + Left(Trim(R("CH5")), 10) + "' WHERE EIDOS='e' AND KOD='" + Trim(R("CH6")) + "'"
132             Gdb.Execute sql, N

            End If

        End If

        'μπαινω απ'ευθειας στο παραστατικό
134     PARAS.Text = PARAS.List(TYPOS - 1)

        Dim ll As Integer

        Dim gg As String

136     gg = IIf(TYPOS = 6, "y", IIf(TYPOS = 9, "A", "Y"))

138     If gg = "y" And Val(mID(ko, 2, 6)) >= 3398 Then
140         gg = "w"
142         MsgBox "υποκαταστηματος"
        End If

144     For ll = 0 To PARAS.ListCount - 1

146         If Right(PARAS.List(ll), 1) = gg Then
148             PARAS.Text = PARAS.List(ll)
            End If

        Next

150     PARAS_LostFocus
152     SendKeys "{ENTER}{ENTER}"
154     SendKeys "{ENTER}{ENTER}"

        Dim k As Integer

        'ΔΗΜΙΟΥΡΓΩ ΤΙΣ ΣΕΙΡΕΣ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ
156     If TYPOS = 9 Then    'δελτιο αποστολής
158         k = 1
160         Grid1.TextMatrix(k, f_k) = ko
162         Grid1.TextMatrix(k, f_p) = 1
164         Grid1.TextMatrix(k, f_o) = R("ONO")
166         Grid1.TextMatrix(k, F_T) = 0
168         Grid1.TextMatrix(k, f_f) = 2
170         Grid1.TextMatrix(k, f_a) = 0
172         Grid1.row = k: Grid1.Col = 0: FindSynolo
174         SKOPOS.Text = "ΑΠΟ ΕΠΙΣΚΕΥΗ"
176         Gdb.Execute "UPDATE EID SET HM3=GETDATE(),NUM1=" + Text1.Text + " WHERE KOD='" + ko + "'"
        Else
178         k = 1
180         Grid1.TextMatrix(k, f_k) = "223344"
182         timtext3 = "223344"
184         Grid1.TextMatrix(k, f_p) = 1
186         Grid1.TextMatrix(k, f_o) = "ΕΠΙΣΚΕΥΗ"
188         Grid1.TextMatrix(k, F_T) = Round(Val(ajia) / (1 + g_Fpa(2) / 100), 2)
190         Grid1.TextMatrix(k, f_f) = 2
192         Grid1.TextMatrix(k, f_a) = Round(Val(ajia) / (1 + g_Fpa(2) / 100), 2)
194         Grid1.row = k: Grid1.Col = 0: FindSynolo
196         k = 2
198         Grid1.TextMatrix(k, f_k) = ko
200         Grid1.TextMatrix(k, f_p) = 0
202         Grid1.TextMatrix(k, f_o) = R("ONO")
204         Grid1.TextMatrix(k, F_T) = 0
206         Grid1.TextMatrix(k, f_f) = 2
208         Grid1.TextMatrix(k, f_a) = 0
210         Grid1.row = k: Grid1.Col = 0: FindSynolo
212         Gdb.Execute "UPDATE EID SET HM3=GETDATE(),NUM2=" + Text1.Text + " WHERE KOD='" + ko + "'"
        End If

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub Command5_Click()

100     EPANEKTYPOSI (0)
End Sub
   
 Sub EPANEKTYPOSI(PDF1EKT2ASK0 As Integer) '3=EMAIL
   
        ' ΕΠΑΝΕΚΤΥΠΩΣΗ ΤΕΛΕΥΤΑΙΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ

        '       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
        '       CrystalReport1.Destination = crptToWindow
        '       Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>

        Dim DUM

        Dim par As New ADODB.Recordset

100     f_dat = DTPicker1.Value

102     f_matim = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text

104     par.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + Left(f_matim, 1) + "';", Gdb, adOpenDynamic, adLockOptimistic

106     If IsNull(par!FORM_ektyp) Then f_FORM_EKTY = 0 Else f_FORM_EKTY = Val(par!FORM_ektyp)

108     f_paras = PARAS.Text

        Dim ANS4 As Integer
   
110     If PDF1EKT2ASK0 = 0 Then
112         ANS4 = MsgBox("Ναι=>ΣΕ PDF" + Chr(13) + "OXI=> ΣΕ ΕΚΤΥΠΩΤΗ", vbYesNo)
114     ElseIf PDF1EKT2ASK0 = 1 Then
116         ANS4 = vbYes
118     ElseIf PDF1EKT2ASK0 = 2 Then
120         ANS4 = vbNo

        ElseIf PDF1EKT2ASK0 = 3 Then
                 Main_Form.ListBox1.AddItem "C:\PDF\PDF.PDF"
                 
                  DUM = printCrystal(f_matim, CDate(f_dat))
                                                  ' CrystalReport1.ReportFileName = "c:\mercvb\reports\timol1pdfpalio.rpt"
                                                
                                                'Me.Caption = CrystalReport1.ReportFileName
                                               ' Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                 
                 
                 
                 
                 
                 Main_Form.SHOW 1

                 Exit Sub
        End If

122     If ANS4 = vbYes Then


             If Len(Dir("c:\mercvb\reports\timol1pdfpalio.rpt", vbNormal)) > 0 Then
                                                   DUM = printCrystal(f_matim, CDate(f_dat))
                                                   CrystalReport1.ReportFileName = "c:\mercvb\reports\timol1pdfpalio.rpt"
                                                
                                                Me.Caption = CrystalReport1.ReportFileName
                                                Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                                                 MILSEC 1000
                                                Exit Sub
             End If
       











            'σβηνω το παλιο μην τυχον το στειλει λαθος
124         If Len(Dir("C:\Mercvb\timoLOGIO.PDF", vbNormal)) > 0 Then
                On Error Resume Next
126             Kill "C:\Mercvb\timologio.PDF"
                On Error GoTo Command5_Click_Err
            End If
   
            Dim XXX2      As Printer
            Dim pdfExists As Integer
128         pdfExists = 0
            Dim printerName As String
130         For Each XXX2 In Printers

132             If InStr(UCase(XXX2.DeviceName), Trim(F_PDFNAME)) > 0 Then
                    ' Set printer as system default.
134                 printerName = XXX2.DeviceName
136                 Set Printer = XXX2
138                 pdfExists = 1
                    ' Stop looking for a printer.
                    Exit For

                End If

            Next

                                If pdfExists = 0 Then
                                   For Each XXX2 In Printers
                                      If InStr(XXX2.DeviceName, "Print To PDF") > 0 Then
'                                           printerName = "Microsoft Print To PDF"
                                           printerName = XXX2.DeviceName
                                           Set Printer = XXX2
                                           pdfExists = 1
                                           
                                                'F_EKTYPOTHS = printerName ' "MERCURYPDF"
                                                If Len(Dir("c:\mercvb\reports\timol1pdf.rpt", vbNormal)) > 0 Then
                                                   CrystalReport1.ReportFileName = "c:\mercvb\reports\timol1pdf.rpt"
                                                Else
                                                   If Len(Trim(F_FORMA1)) > 1 Then
                                                      CrystalReport1.ReportFileName = F_FORMA1
                                                   Else
                                                       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt" ' F_FORMA1
                                                   End If
                                                End If
                                                'Kill "c:\pdf\pdf.pdf"
                                                Me.Caption = CrystalReport1.ReportFileName
                                                Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                                                MILSEC 1000
                                                Dim ANS3 As Integer
                                                
                                           
                                           
                                                Exit Sub
                                      End If
                                           
                                 ' MsgBox "δεν υπάρχει PDF εκτυπωτής"
                                  'Exit Sub
                                  Next
                                End If

            ' Dim VV As CrystalReport
140         f_same_eline = 0
            'If par!C1 = "215" Then f_same_eline = 1
             
142         If Val(nNull(par!C1)) = 215 Then f_same_eline = 1
144         DUM = printCrystal(f_matim, CDate(f_dat))
146         f_same_eline = 0
            '
            '            VV.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
            '            VV.printerName = Printer.DeviceName
            '            VV.PrinterPort = Printer.Port
            '            VV.PrinterDriver = Printer.DriverName
            '            VV.Action = 1
            ' Set Printer = Printers(Combo1.ListIndex)
148         If Len(Trim(CNull(par!FORMA1))) > 2 Then
150             CrystalReport1.ReportFileName = par!FORMA1
            Else
152             CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
            End If
                   
            On Error Resume Next
                   
154         If pdfExists = 0 Then
156             Shell "C:\MERCVB\REP.EXE  " + CrystalReport1.ReportFileName, vbNormalFocus
                ' preview1.Command1.Caption = CrystalReport1.ReportFileName
                ' preview1.SHOW 1
                   
                'MsgBox "δεν υπάρχει PDF εκτυπωτής"
                Exit Sub
            End If
                   
            '     CrystalReport1.printerName = Printer.DeviceName
            '        CrystalReport1.PrinterPort = Printer.Port
            '        CrystalReport1.PrinterDriver = Printer.DriverName
                   
            'On Error GoTo 0
            Dim crxApp As CRAXdDRT.Application
            Dim crxRpt As CRAXdDRT.Report
                   
158         Set crxApp = New CRAXdDRT.Application
160         Set crxRpt = crxApp.OpenReport(CrystalReport1.ReportFileName)
            'Set CRXReport = CRXApplication.OpenReport("c:\MERCVB\reports\timol1.rpt", 1)
               
            ' ExportReportToPDF crxRpt, "c:\mercvb\reports\timol1.pdf", "titlos"
                   
162         crxRpt.SelectPrinter "", printerName, ""
            ' crxRpt. = "Bullzip PDF Printer"  'Printer.DeviceName ' "Bullzip PDF Printer"
164         crxRpt.PrintOut False, 1  ' NoCopies
                   
            '                      CrystalReport1.printerName = Printer.DeviceName
            '                      CrystalReport1.PrinterPort = Printer.Port
            '                      CrystalReport1.PrinterDriver = Printer.DriverName
            '                      Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
            '                      Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName   ' "C:\Mercvb\reports\timol1.rpt"
                      
            
'            If vbYes = MsgBox("Να σταλει με email;", vbYesNo) Then
'
'                 Main_Form.ListBox1.AddItem "C:\PDF\PDF.PDF"
'                 Main_Form.SHOW 1
'
'
'
'            End If
            
            
            
            
            
            Exit Sub

        End If

166     f_same_eline = 0

        'εκτύπωση
168     If ektyp.Value = 1 Then    ' ---------------------------------------------

170         If f_FORM_EKTY >= 90 Then    ' φορμα τύπου f90.txt  '====================================
172             f_same_eline = 0
174             DUM = printCrystal(f_matim, CDate(DTPicker1.Value))    'δημιουργει τα προσωρινά αρχεία
176             f_same_eline = 0
178             DUM = print_timol(f_matim, CDate(DTPicker1.Value))
            
180         ElseIf Len(Trim(F_FORMA1)) > 1 Then
182             f_same_eline = 0
184             DUM = printCrystal(f_matim, CDate(f_dat))
186             f_same_eline = 0

188             If DUM < -900 Then
190                 MsgBox "αδυνατη η εκτύπωση" + Chr(13) ''   + "Διορθώστε το παραστατικό για να ξανατυπωθεί"
                Else
192                 Frame1.Caption = CNull(F_FORMA1)
194                 CrystalReport1.ReportFileName = F_FORMA1
196                 Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName

198                 If Len(F_FORMA2) > 1 Then
200                     CrystalReport1.ReportFileName = F_FORMA2
202                     Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                    End If

204                 If Len(F_FORMA3) > 1 Then
206                     CrystalReport1.ReportFileName = F_FORMA3
208                     Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                    End If
                   
210                 If F_PDF = 4 Then
                        ' ΑΝΤΙΓΡΑΦΕΙ ΤΟ ΑΡΧΕΙΟ ΑΠΟ ΤΟ C:\PDF\TEMP STO C:\PDF
                        ' ΜΕΤΟΝΟΜΑΖΟΝΤΑΣ ΤΑ ΜΕ ΤΟ ΟΝΟΜΑ ΠΟΥ ΒΡΙΣΚΕΙ ΣΤΟ C:\PRF\NAME
                        Dim N123 As Integer
                        ' δεν κάνει τίποτα παρά μόνο πετάει το όνομα sto c:\pdf\name που πρεπει να δοθεί
                        ' στο αρχείo c:\pdf\pdftemp
212                     MILSEC 2000
                        Dim Pdf22 As String
                    
214                     Pdf22 = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF")
216                     Do While Len(Pdf22) = 0
218                         MILSEC 1000
220                         N123 = N123 + 1
222                         If N123 > 10 Then
                                Exit Do
                            End If
224                         Pdf22 = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF")
                        Loop
  
226                     N123 = 0
                        On Error Resume Next
                    
228                     Do While Len(Pdf22) > 0

                            On Error Resume Next

230                         N123 = N123 + 1
232                         FileCopy F_PDFFOLDER + "\PDFTEMP\" + Pdf22, F_PDFFOLDER + "\" + f_matim + Trim(str(N123)) + ".PDF"
234                         MILSEC 1000
236                         Kill F_PDFFOLDER + "\PDFTEMP\" + Pdf22
238                         Pdf22 = ""
240                         Pdf22 = Dir '("C:\PDF\PDFTEMP\*.PDF")
                        Loop
  
242                     Kill F_PDFFOLDER + "\PDFTEMP\*.PDF"
    
                    End If
                 
                End If
            
            Else    ' crystal reports timol1.rpt(σφραγίδα)   timol1an.rpt (αντίγραφο)

244             If f_FORM_EKTY > 0 Then    '      PEL.Recordset("eidos") = "e" Then  ' EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE

                    'ΑΝ ΧΡΗΣΙΜΟΠΟΙΕΙ ΤΗΝ ΦΟΡΜΑ ΤΙΜΟΛ2 TOTE NA DINEI ALLES SEIRES APO TO F_FORM2_SEIRESTIMOL
246                 If f_FORM_EKTY = 14 Or f_FORM_EKTY = 12 Then
248                     f_SeiresTimologioy = f_Timol2_SeiresTimologioy
250                     f_SynSeiresTimologioy = f_Timol2_SeiresTimologioy
                    End If

                    ' On Error GoTo Command5_Click_Err ' na sbhstei 15-6-2006
252                 If f_FORM_EKTY = 15 Or f_FORM_EKTY = 51 Or f_FORM_EKTY = 52 Then     '  ME SQL SERVER H EKTYPOSI
254                     DUM = PrintSqlCrystal(f_matim, CDate(f_dat), 0)  'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
                    Else  ' ME MDB
256                     f_same_eline = 0
258                     DUM = printCrystal(f_matim, CDate(f_dat))    'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
260                     f_same_eline = 0
                        
                    End If

                    '   On Error Resume Next

262                 If DUM < -900 Then
264                     MsgBox "αδυνατη η εκτύπωση" + Chr(13) '+ "Διορθώστε το παραστατικό για να ξανατυπωθεί"

                    Else
                        Dim ANS As Integer, xxx As Printer
266                     ANS = vbNo ' MsgBox("Σε οθόνη το τιμολόγιο;", vbYesNo)
268                     If ANS = vbYes Then
270                         CrystalReport1.Destination = crptToWindow
                        End If
272                     If f_FORM_EKTY = 12 Then     '     λιανικη
274                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"
276                     ElseIf f_FORM_EKTY = 13 Then     '     λιανικη
278                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol3.rpt"
280                     ElseIf f_FORM_EKTY = 15 Then     '     λιανικη
282                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timolsql.rpt"
284                         CrystalReport1.Connect = gConnect
286                     ElseIf f_FORM_EKTY = 25 Then     '     λιανικη
                            ' CrystalReport1.ReportFileName = "C:\Mercvb\reports\timolsql.rpt"
                            ' CrystalReport1.Connect = gConnect
288                         Ucr9print.printingCR9 F_EKTYPOTHS, "C:\Mercvb\reports\timolsql.rpt"
                        Else
290                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
                        End If

292                     If f_FORM_EKTY = 10 Then
294                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
                            ' Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
296                     ElseIf f_FORM_EKTY = 20 Then

298                         CrystalReport1.Destination = crptToWindow
300                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
302                     ElseIf f_FORM_EKTY = 18 Then
304                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName  'Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
306                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName  'Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
308                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName  'Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
                        Else
310                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName  ' Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
312                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName  'Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
                        End If
                    
                    End If

                End If   ' If f_FORM_EKTY > 0 Then '  EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE

            End If  ' If ektyp.Value = 1 Then  '========================================

        End If  '-------------------------------------------------------------

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
  
   
   
   
   
   
   
'Sub EPANEKTYPOSI(PDF1EKT2ASK0 As Integer)
'
'        ' ΕΠΑΝΕΚΤΥΠΩΣΗ ΤΕΛΕΥΤΑΙΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ
'
'        '       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
'        '       CrystalReport1.Destination = crptToWindow
'        '       Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
'        '<EhHeader>
'        On Error GoTo Command5_Click_Err
'
'        '</EhHeader>
'
'        Dim DUM
'
'        Dim par As New ADODB.Recordset
'
'100     f_dat = DTPicker1.Value
'
'102     f_matim = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
'
'104     par.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + Left(f_matim, 1) + "';", Gdb, adOpenDynamic, adLockOptimistic
'
'106     If IsNull(par!FORM_ektyp) Then f_FORM_EKTY = 0 Else f_FORM_EKTY = Val(par!FORM_ektyp)
'
'108     f_paras = PARAS.Text
'
'        Dim ANS4 As Integer
'
'110     If PDF1EKT2ASK0 = 0 Then
'112         ANS4 = MsgBox("ΣΕ PDF TO ΠAΡAΣTATIKO(TIMOL1.RPT);", vbYesNo)
'114     ElseIf PDF1EKT2ASK0 = 1 Then
'116         ANS4 = vbYes
'118     ElseIf PDF1EKT2ASK0 = 2 Then
'120         ANS4 = vbNo
'        End If
'
'122     If ANS4 = vbYes Then
'            'σβηνω το παλιο μην τυχον το στειλει λαθος
'124         If Len(Dir("C:\Mercvb\timoLOGIO.PDF", vbNormal)) > 0 Then
'                On Error Resume Next
'126             Kill "C:\Mercvb\timologio.PDF"
'                On Error GoTo Command5_Click_Err
'            End If
'
'            Dim XXX2      As Printer
'            Dim pdfExists As Integer
'128         pdfExists = 0
'            Dim printerName As String
'130         For Each XXX2 In Printers
'
'132             If InStr(UCase(XXX2.DeviceName), F_PDFNAME) > 0 Then
'                    ' Set printer as system default.
'134                 printerName = XXX2.DeviceName
'136                 Set Printer = XXX2
'138                 pdfExists = 1
'                    ' Stop looking for a printer.
'                    Exit For
'
'                End If
'
'            Next
'
'            '                    If pdfExists = 0 Then
'            '                      MsgBox "δεν υπάρχει PDF εκτυπωτής"
'            '                      Exit Sub
'            '                    End If
'
'            ' Dim VV As CrystalReport
'140         f_same_eline = 0
'            'If par!C1 = "215" Then f_same_eline = 1
'
'142         If Val(nNull(par!C1)) = 215 Then f_same_eline = 1
'144         DUM = printCrystal(f_matim, CDate(f_dat))
'146         f_same_eline = 0
'            '
'            '            VV.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
'            '            VV.printerName = Printer.DeviceName
'            '            VV.PrinterPort = Printer.Port
'            '            VV.PrinterDriver = Printer.DriverName
'            '            VV.Action = 1
'            ' Set Printer = Printers(Combo1.ListIndex)
'148         If Len(Trim(CNull(par!FORMA1))) > 2 Then
'150             CrystalReport1.ReportFileName = par!FORMA1
'            Else
'152             CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
'            End If
'
'            On Error Resume Next
'
'154         If pdfExists = 0 Then
'156             Shell "C:\MERCVB\REP.EXE  " + CrystalReport1.ReportFileName, vbNormalFocus
'                ' preview1.Command1.Caption = CrystalReport1.ReportFileName
'                ' preview1.SHOW 1
'
'                'MsgBox "δεν υπάρχει PDF εκτυπωτής"
'                Exit Sub
'            End If
'
'            '     CrystalReport1.printerName = Printer.DeviceName
'            '        CrystalReport1.PrinterPort = Printer.Port
'            '        CrystalReport1.PrinterDriver = Printer.DriverName
'
'            'On Error GoTo 0
'            Dim crxApp As CRAXdDRT.Application
'            Dim crxRpt As CRAXdDRT.Report
'
'158         Set crxApp = New CRAXdDRT.Application
'160         Set crxRpt = crxApp.OpenReport(CrystalReport1.ReportFileName)
'            'Set CRXReport = CRXApplication.OpenReport("c:\MERCVB\reports\timol1.rpt", 1)
'
'            ' ExportReportToPDF crxRpt, "c:\mercvb\reports\timol1.pdf", "titlos"
'
'162         crxRpt.SelectPrinter "", printerName, ""
'            ' crxRpt. = "Bullzip PDF Printer"  'Printer.DeviceName ' "Bullzip PDF Printer"
'164         crxRpt.PrintOut False, 1  ' NoCopies
'
'            '                      CrystalReport1.printerName = Printer.DeviceName
'            '                      CrystalReport1.PrinterPort = Printer.Port
'            '                      CrystalReport1.PrinterDriver = Printer.DriverName
'            '                      Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
'            '                      Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName   ' "C:\Mercvb\reports\timol1.rpt"
'
'            Exit Sub
'
'        End If
'
'166     f_same_eline = 0
'
'        'εκτύπωση
'168     If ektyp.Value = 1 Then    ' ---------------------------------------------
'
'170         If f_FORM_EKTY >= 90 Then    ' φορμα τύπου f90.txt  '====================================
'172             f_same_eline = 0
'174             DUM = printCrystal(f_matim, CDate(DTPicker1.Value))    'δημιουργει τα προσωρινά αρχεία
'176             f_same_eline = 0
'178             DUM = print_timol(f_matim, CDate(DTPicker1.Value))
'
'180         ElseIf Len(Trim(F_FORMA1)) > 1 Then
'182             f_same_eline = 0
'184             DUM = printCrystal(f_matim, CDate(f_dat))
'186             f_same_eline = 0
'
'188             If DUM < -900 Then
'190                 MsgBox "αδυνατη η εκτύπωση" + Chr(13) ''   + "Διορθώστε το παραστατικό για να ξανατυπωθεί"
'                Else
'192                 Frame1.Caption = CNull(F_FORMA1)
'194                 CrystalReport1.ReportFileName = F_FORMA1
'196                 Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
'
'198                 If Len(F_FORMA2) > 1 Then
'200                     CrystalReport1.ReportFileName = F_FORMA2
'202                     Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
'                    End If
'
'204                 If Len(F_FORMA3) > 1 Then
'206                     CrystalReport1.ReportFileName = F_FORMA3
'208                     Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
'                    End If
'
'210                 If F_PDF = 4 Then
'                        ' ΑΝΤΙΓΡΑΦΕΙ ΤΟ ΑΡΧΕΙΟ ΑΠΟ ΤΟ C:\PDF\TEMP STO C:\PDF
'                        ' ΜΕΤΟΝΟΜΑΖΟΝΤΑΣ ΤΑ ΜΕ ΤΟ ΟΝΟΜΑ ΠΟΥ ΒΡΙΣΚΕΙ ΣΤΟ C:\PRF\NAME
'                        Dim N123 As Integer
'                        ' δεν κάνει τίποτα παρά μόνο πετάει το όνομα sto c:\pdf\name που πρεπει να δοθεί
'                        ' στο αρχείo c:\pdf\pdftemp
'212                     MILSEC 2000
'                        Dim Pdf22 As String
'
'214                     Pdf22 = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF")
'216                     Do While Len(Pdf22) = 0
'218                         MILSEC 1000
'220                         N123 = N123 + 1
'222                         If N123 > 10 Then
'                                Exit Do
'                            End If
'224                         Pdf22 = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF")
'                        Loop
'
'226                     N123 = 0
'                        On Error Resume Next
'
'228                     Do While Len(Pdf22) > 0
'
'                            On Error Resume Next
'
'230                         N123 = N123 + 1
'232                         FileCopy F_PDFFOLDER + "\PDFTEMP\" + Pdf22, F_PDFFOLDER + "\" + f_matim + Trim(str(N123)) + ".PDF"
'234                         MILSEC 1000
'236                         Kill F_PDFFOLDER + "\PDFTEMP\" + Pdf22
'238                         Pdf22 = ""
'240                         Pdf22 = Dir '("C:\PDF\PDFTEMP\*.PDF")
'                        Loop
'
'242                     Kill F_PDFFOLDER + "\PDFTEMP\*.PDF"
'
'                    End If
'
'                End If
'
'            Else    ' crystal reports timol1.rpt(σφραγίδα)   timol1an.rpt (αντίγραφο)
'
'244             If f_FORM_EKTY > 0 Then    '      PEL.Recordset("eidos") = "e" Then  ' EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
'
'                    'ΑΝ ΧΡΗΣΙΜΟΠΟΙΕΙ ΤΗΝ ΦΟΡΜΑ ΤΙΜΟΛ2 TOTE NA DINEI ALLES SEIRES APO TO F_FORM2_SEIRESTIMOL
'246                 If f_FORM_EKTY = 14 Or f_FORM_EKTY = 12 Then
'248                     f_SeiresTimologioy = f_Timol2_SeiresTimologioy
'250                     f_SynSeiresTimologioy = f_Timol2_SeiresTimologioy
'                    End If
'
'                    ' On Error GoTo Command5_Click_Err ' na sbhstei 15-6-2006
'252                 If f_FORM_EKTY = 15 Or f_FORM_EKTY = 51 Or f_FORM_EKTY = 52 Then     '  ME SQL SERVER H EKTYPOSI
'254                     DUM = PrintSqlCrystal(f_matim, CDate(f_dat), 0)  'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
'                    Else  ' ME MDB
'256                     f_same_eline = 0
'258                     DUM = printCrystal(f_matim, CDate(f_dat))    'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
'260                     f_same_eline = 0
'
'                    End If
'
'                    '   On Error Resume Next
'
'262                 If DUM < -900 Then
'264                     MsgBox "αδυνατη η εκτύπωση" + Chr(13) '+ "Διορθώστε το παραστατικό για να ξανατυπωθεί"
'
'                    Else
'                        Dim ANS As Integer, xxx As Printer
'266                     ANS = vbNo ' MsgBox("Σε οθόνη το τιμολόγιο;", vbYesNo)
'268                     If ANS = vbYes Then
'270                         CrystalReport1.Destination = crptToWindow
'                        End If
'272                     If f_FORM_EKTY = 12 Then     '     λιανικη
'274                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"
'276                     ElseIf f_FORM_EKTY = 13 Then     '     λιανικη
'278                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol3.rpt"
'280                     ElseIf f_FORM_EKTY = 15 Then     '     λιανικη
'282                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timolsql.rpt"
'284                         CrystalReport1.Connect = gConnect
'286                     ElseIf f_FORM_EKTY = 25 Then     '     λιανικη
'                            ' CrystalReport1.ReportFileName = "C:\Mercvb\reports\timolsql.rpt"
'                            ' CrystalReport1.Connect = gConnect
'288                         Ucr9print.printingCR9 F_EKTYPOTHS, "C:\Mercvb\reports\timolsql.rpt"
'                        Else
'290                         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
'                        End If
'
'292                     If f_FORM_EKTY = 10 Then
'294                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
'                            ' Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
'296                     ElseIf f_FORM_EKTY = 20 Then
'
'298                         CrystalReport1.Destination = crptToWindow
'300                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
'302                     ElseIf f_FORM_EKTY = 18 Then
'304                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName  'Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
'306                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName  'Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
'308                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName  'Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
'                        Else
'310                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName  ' Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
'312                         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName  'Ucr9print.printingcr9  F_EKTYPOTHS, CrystalReport1.ReportFileName= 1
'                        End If
'
'                    End If
'
'                End If   ' If f_FORM_EKTY > 0 Then '  EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
'
'            End If  ' If ektyp.Value = 1 Then  '========================================
'
'        End If  '-------------------------------------------------------------
'
'        '<EhFooter>
'        Exit Sub
'
'Command5_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.par1.Command5_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Command5_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub

Private Sub Command6_Click()

        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>
    
        Dim fff As New pelat2

        'fff.ffbuff = f_buff + data1.Recordset("kod") ' Text1(0).Text
   
100     gApoMenu = False
        'fff.COMMAND5.Enabled = False
    
102     If POLAGOR.ListIndex = 0 Then
104         fff.ffbuff = "r" + Text2(0).Text 'data1.Recordset("kod")    ' Text1(0).Text

            Exit Sub

        Else

106         If f_KENTRA_ADYNATISMATOS = 1 Then
108             If kentroAdyn.ListIndex = 1 Then
110                 fff.kentroAdyn.ListIndex = 1
                End If
            Else
112             fff.PELPROM.Text = "e"  'PELPROM.Text
114             fff.Text2(0).Text = Text2(0).Text
                ' fff.Text2_LostFocus 0
                'SendKeys "{ENTER}{ENTER} {ENTER} {ENTER}{ENTER}"
116             fff.Text2_LostFocus 0
                      
118             fff.cmdDiortHosi
120             fff.Frame2.Visible = False
122             fff.Frame1.Top = fff.Toolbar1.Height
124             fff.Toolbar1.Buttons.Item("cmdDiortosi").Value = tbrPressed
126             fff.Toolbar1.Buttons.Item("cmdNew").Enabled = False ' vbRed
      
128             fff.Toolbar1.Buttons.Item("cmdKartela").Enabled = False
130             fff.Toolbar1.Buttons.Item("cmdDelete").Enabled = False
132             fff.Toolbar1.Buttons.Item("cmdEtiketa").Enabled = False
                ' cmdEtiketa
134             fff.Toolbar1.Buttons.Item("cmdChange").Enabled = False
136             fff.Toolbar1.Buttons.Item("cmdUpdate").Enabled = True
       
138             CmdRefreshPel.Visible = True
                     
            End If

140         fff.ffbuff = "e" + Text2(0).Text 'data1.Recordset("kod")    ' Text1(0).Text

            Exit Sub

        End If
  
142     fff.SHOW
      
144     gApoMenu = True
  
        '    'left(data1.Recordset!kod, 5) + data1.Recordset!epo
        '    On Error GoTo LATOS
        '    pelat2.PELPROM.Text = PELPROM.Text
        '    pelat2.Text2(0).Text = Left(data1.Recordset!kod, 5)
        '
        '
        '
        '    pelat2.SHOW
        '    Exit Sub
        '
    
        '    Dim fff As New pelat2
        '
        '    ' fff.ffbuff = f_buff + Text2(0).Text ' Text1(0).Text
        '
        '    If POLAGOR.ListIndex = 0 Then
        '        fff.PELPROM.ListIndex = 0
        '        fff.ffbuff = "r" + Text2(0).Text    ' Text1(0).Text
        '
        '    Else
        '        fff.PELPROM.ListIndex = 1
        '        fff.ffbuff = "e" + Text2(0).Text    ' Text1(0).Text
        '    End If
        '    fff.Command5.Enabled = False
        '    gApoMenu = False
        '    pelat2.SHOW

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Sub old_diortosis()
'        '<EhHeader>
'        On Error GoTo diortosis_Err
'        '</EhHeader>
'
'        Dim FF As PAR2, mkod As String
'
'        Dim Atim
'
'        'POLAGOR.SetFocus
'        'POLAGOR_LostFocus
'        'PARAS.SetFocus
'        'PARAS_LostFocus
'
'
'
'
'
'
'
'
'
'
'
'
'100     MIDENIZO_GRID
'
'110     TDBGrid1.Col = 4
'120     mkod = TDBGrid1.Text
'
'
'
'130     TDBGrid1.Col = 1
'140     Atim = TDBGrid1.Text
'
'
'
'        Dim k As Integer
'
'
'        'ΧΡΕΙΑΖΕΤΑΙ 2 ΦΟΡΕΣ ΓΙΑ ΝΑ ΚΑΝΕΙ ΔΙΟΡΘΒ=ΩΣΗ ΜΕ ΤΗΝ ΠΡΩΤΗ
'        '------------------------------------------------------------------------------
'        'For k = 0 To PARAS.ListCount - 1
'        '   PARAS.ListIndex = k
'        '   If Right(PARAS.Text, 1) = Left(atim, 1) Then
'        '      Exit For
'        '   End If
'        'Next
'        ''  find_eid_parastat
'        '  Text1.Text = Val(mID(atim, 2, 6))
'        '  Text1.SetFocus
'        ''MILSEC 3500
'        ''DoEvents
'        ''  Text2(0).Text = mkod
'        ''      PEL.RecordSource = "select  LEFT(left(EPO+'-'+DIE+'-',56)+KOD,60) as xx,DIE,THL,KOD  from PEL where EIDOS='" + X + "' and KOD like '" + SameLetters(Text2(0).Text) + "%' order by KOD;"
'        '''End If
'        ''
'        ''
'        ''
'        'GridPelaton.Columns(0).width = 6000
'        'GridPelaton.Columns(0).DataField = "XX"
'        '
'        'PEL.Refresh
'        'GridPelaton.Columns(0).width = 6000
'        'GridPelaton.Visible = True
'        'GridPelaton.SetFocus
'        'SendKeys "{ENTER}{ENTER} {ENTER} "
'
'
'
'
'        '------------------------------------------------------------------------------
'
'
'
'
'
'        Dim r As New ADODB.Recordset
'
'150     r.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + Left(Atim, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
'160     If r.EOF Then Exit Sub
'170     If r("pol") > 3 Or r("pol") < 1 Then Exit Sub    ' δεν διορθωνω  asxeta
'
'
'180     If r("pol") = 1 Then    'pvlhseis
'190         POLAGOR.ListIndex = 1
'200     ElseIf r("pol") = 2 Then
'210         POLAGOR.ListIndex = 0    ' agores
'220     ElseIf r("pol") = 3 Then
'230         POLAGOR.ListIndex = 2    ' κινησεις αποθηκης
'        End If
'
'
'
'        '  PARAS.Text = R("titlos") + Space(1) + Left(Atim, 1)
'240     POLAGOR_LostFocus
'        'POLAGOR.SetFocus
'        'PARAS.SetFocus
'        'Dim k As Integer
'250     For k = 0 To PARAS.ListCount - 1
'260         PARAS.ListIndex = k
'270         If Right(PARAS.Text, 1) = Left(Atim, 1) Then
'                Exit For
'            End If
'        Next
'
'
'280     TDBGrid1.Col = 0
'290     DTPicker1.Value = CDate(TDBGrid1.Text)
'
'300     If k = PARAS.ListCount Then Exit Sub
'        ' βρισκω τις παραμετρους του παραστατικού
'310     find_eid_parastat
'
'320     Text1.Text = Val(mID(Atim, 2, 6))
'
'330     If Len(Trim(mID(Atim, 8, 1))) > 0 Then
'           '  SEIRA.Text = mID(Atim, 8, 1)
'        Else
'          '  SEIRA.Text = ""
'        End If
'
'340     Text2(0).Text = mkod
'
'350     Text2(0).Enabled = True
'360     Text2(0).SetFocus
'
'
'SendKeys "{ENTER} {ENTER} {ENTER} "
'
'
'        '<EhFooter>
'        Exit Sub
'
'diortosis_Err:
'        'MsgBox Err.Description & vbCrLf & _
'               "in ADOMERCNEW.par1.diortosis " & _
'               "at line " & Erl, _
'               vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.diortosis " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub

Private Sub Command7_Click()
        '  Dim f_strogg As Single 'ποσό στρογγυλοποίησης ΦΠΑ
        '  Dim f_strogg_FPA As Single 'ΚΑΤΗΓΟΡΙΑ ΦΠΑ ΠΟΥ ΘΑ ΓΙΝΕΙ Η ΣΤΡΟΓΓΥΛΟΠΟΙΗΣΗ

        Dim k As Integer

        'For K = 1 To 9
        '  If SumFpa(K) > 0 Then
        '       f_strogg = -0.01
        '     f_strogg_FPA = K
        ' End If
        'Next
If Val(sfpa.Caption) = 0 Then
    'αν ειναι χωρισ φπα δεν κανει στρογγυλοποιηση
Else
100     If f_strogg = 0.01 Then
102         f_strogg = 0
        Else
104         f_strogg = -0.01
        End If

106     FindSynolo
End If



End Sub

Private Sub Command8_Click()

If Val(sfpa.Caption) = 0 Then
    'αν ειναι χωρισ φπα δεν κανει στρογγυλοποιηση
Else
100     If f_strogg = -0.01 Then
102         f_strogg = 0
        Else
104         f_strogg = 0.01
        End If

106     FindSynolo

End If


End Sub

Private Sub Command9_Click()
    
    
    
    
    


Dim qrl As String ': qrl = "https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=vZGdPC4qtv61lfj3tMcBWCVzle6HNmunsN6JjmX4e9gLqhLIIIUNGzIx6YzWdXmeg0fsTOtxriHc%2bCRo3shXRfuFVdEayG8L9Xt9%2fSH8HaE%3d"

qrl = InputBox("σκανάρισε το qrcode", " ")



           Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)

'"https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=NziZ2b5egY41ooVijo%2fjtSMRj7CvuLB%2fxG1YnEyrsSdlsH7ef9ZcgJZCLIFhuHMUa6Cp8pQjpI4ErLYp4fpsmjaTE7asa34Phn2SIB5cM4c%3d"
'https://mydataapidev.aade.gr/RegisterTransfer
qrl = GGET_CVALUE("select QRURL FROM TIM WHERE ID_NUM=" + str(m_ID))
'https://mydatapi.aade.gr/myDATA/ConfirmDeliveryOutcome
Dim TXML As String
TXML = "<?xml version=""1.0"" encoding=""utf-8""?>"   'qrUrl,outcome"
TXML = TXML + "<ConfirmDeliveryOutcomeRequest>"
TXML = TXML + "<qrUrl>" + qrl + "</qrUrl>"
TXML = TXML + "<outcome>FULL</outcome>"
TXML = TXML + "</ConfirmDeliveryOutcomeRequest>"

Dim DUM As String: DUM = Par7MyData.SendConfirm(TXML, " ")
    
    
    
    
    ' keybd_event VK_ENTER, 0, 0, 0    'send a tab
End Sub

'Private Sub Command7_Click()
    ' Dim metasx_lian As Integer
    'metasx_lian = Val(FindParametroi(1,"PAR1", "METASX_LIAN", "2", "ΠΑΡΑΣΤΑΤΙΚO ΔΛ.ΔΑ(ΜΕΤΑΣΧ.ΣΕ ΛΙΑΝΙΚΗ)0-99"))
    'PAR1.POLAGOR.ListIndex = 1  ' POLHSEIS
    'PAR1.PARAS.ListIndex = 2
    'PAR1.PARAS.SetFocus
        
'End Sub

Private Sub DTPicker1_GotFocus()

        '<EhHeader>
        On Error GoTo DTPicker1_GotFocus_Err

        '</EhHeader>

100     DTPicker1.ToolTipText = "."

        '<EhFooter>
        Exit Sub

DTPicker1_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.DTPicker1_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.DTPicker1_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo DTPicker1_KeyUp_Err

        '</EhHeader>
100     If KeyCode = vbKeyReturn Then
102         KeyCode = 0    'suppress the beep

104         If DTPicker1.ToolTipText = "." Then
106             DTPicker1.ToolTipText = ""
            Else
108             keybd_event VK_TAB, 0, 0, 0    'send a tab
            End If
        End If

        '<EhFooter>
        Exit Sub

DTPicker1_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.DTPicker1_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.DTPicker1_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub menuToTransport(parametros As Integer)


        Dim m As Integer

        Dim c As String

        PAR1.TDBGrid1.Col = 4  'PELKOD
        c = Replace(PAR1.TDBGrid1.Text, ",", ".")
        
        
           Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)

Dim AFMEPIX As String
AFMEPIX = GGET_CVALUE("SELECT TOP 1 AFMPEL FROM MEM ORDER BY ID")




Dim qrl As String:

qrl = GGET_CVALUE("select QRURL FROM TIM WHERE ID_NUM=" + str(m_ID))
Dim AYTOK As String
AYTOK = GGET_CVALUE("select AYTOK FROM TIM WHERE ID_NUM=" + str(m_ID))

'qrl = "https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=vZGdPC4qtv61lfj3tMcBWCVzle6HNmunsN6JjmX4e9gLqhLIIIUNGzIx6YzWdXmeg0fsTOtxriHc%2bCRo3shXRfuFVdEayG8L9Xt9%2fSH8HaE%3d"

Dim TXML As String
TXML = "<?xml version=""1.0"" encoding=""utf-8""?>"   'qrUrl,outcome"
TXML = TXML + "<Transport>"
TXML = TXML + "<qrUrl>" + qrl + "</qrUrl>"

TXML = TXML + "<transportDetail>"
'                                                                      Dim ark As String: ark = "PMT4231"

TXML = TXML + "<vehicleNumber>" + AYTOK + "</vehicleNumber>"

Dim TYPOS As String: TYPOS = "7"
TXML = TXML + "<transportType>" + TYPOS + "</transportType>"

TXML = TXML + "<carrierVatNumber>" + AFMEPIX + "</carrierVatNumber>"

Dim LoN, LAT As String
LoN = "24.12468"
LAT = "41.136"

'txml = txml + "<longitude>" + LoN + "</longitude>"
'txml = txml + "<latitude>" + LAT + "</latitude>"

'Κωδικός Περιγραφή 1 Φορτηγό Δημόσιας Χρήσης 2 Φορτηγό Ιδιωτικής Χρήσης
'3 Πλοίο 4 Τρένο 5 Αεροπλάνο 6 Λοιπά Μεταφορικά Μέσα (π.χ Δίκυκλα, ..) 7 Άνευ


TXML = TXML + "</transportDetail>"

TXML = TXML + "</Transport>"

Dim DUM As String: DUM = Par7MyData.ENARXIAPOS(TXML, " ")

End Sub




Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.category.AddItem "1 .ΠΑΡΑΓΓΕΛΙΕΣ"
102     PARAMETROI.category.AddItem "2. ΜΑΡΚ.ΛΙΑΝΙΚΗ"
         
104     PARAMETROI.category.AddItem "3 .ΓΕΝΙΚΕΣ ΠΑΡΑΜΕΤΡΟΙ"
106     PARAMETROI.category.AddItem "4 .ΦΟΡΜΕΣ-PDF"
108     PARAMETROI.category.AddItem "5 .ΕΙΔΙΚΕΣ ΕΦΑΡΜΟΓΕΣ-CUSTOMS"
110     PARAMETROI.category.AddItem "6.ΠΛΑΤΗ ΣΤΗΛΩΝ ΤΙΜΟΛΟΓΗΣΗΣ"
         
112     PARAMETROI.category.AddItem "7 .ΣΥΜΠΕΡΙΦΟΡΑ ΤΙΜΟΛΟΓΗΣΗΣ"
        
114     PARAMETROI.category.AddItem "8.ΑΡ.ΨΗΦΙΩΝ-ΤΙΜΕΣ"
116     PARAMETROI.category.AddItem "9.ΜΕΤΑΣΧΗΜΑΤΙΣΜΟΣ"
        
118     PARAMETROI.category.Visible = True
        
120     PARAMETROI.PARAM.Caption = "PAR1"
122     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then b_akyr_Click

        ' If KeyCode = 38 Then 'PANO BELOS KANEI SHIFT+TAB
        '    SendKeys "+{TAB}"
        ' End If
102     If KeyCode = 113 Then Shell "calc.exe"

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

100     If Shift = 1 Then
            Dim oleNewColor As OLE_COLOR
102         oleNewColor = ShowColorDialog(Me.hwnd, True, vbRed)
104         If oleNewColor <> -1 Then Grid1.BackColor = oleNewColor

106         Gdb.Execute "UPDATE PARAMETROI SET TIMH=" + str(Grid1.BackColor) + " WHERE  FORMA='PAR1' AND VAR='f_GRIDCOLOR' "

        End If

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
 
 
 If F_VB6RESIZEFORM = 0 Then
 
    Exit Sub
 End If
 
 
 
    If F_PLATOS_FORMAS < 0 Then
        ResizeFormFor Me
    Else
        mForm_Resize Me, 8, F_PLATOS_FORMAS, F_YCOS_FORMAS, fh, fw, ft, fl
    End If
    
    'τα ξαναδινω (4 επομενα) γιατι τα εχανε τα captions
    Labf(0).Caption = f_CAPTION_labf0 '"μον"
    Labf(1).Caption = f_CAPTION_labf1 '"ποσοτητα"

    Labf(2).Caption = f_CAPTION_labf2 '"αριθ.παρ"      'textm(6)  anazhthsh entolea

    Labf(3).Caption = f_CAPTION_labf3 '"π
    Label3(0).Left = Text2(3).Left
    Text1.Height = POLAGOR.Height
    selectPEL.Height = POLAGOR.Height
    DTPicker1.Height = POLAGOR.Height
    Combo8.Height = POLAGOR.Height
    APOTH1.Height = POLAGOR.Height
    Dim mtop As Long
    mtop = POLAGOR.Top
    
    Text1.Top = mtop
    selectPEL.Top = mtop
    DTPicker1.Top = mtop
    Combo8.Top = mtop
    APOTH1.Top = mtop
    
    If F_GRID_WIDTH > 10320 Then
        ' Grid1.Width = F_GRID_WIDTH
    Else
        ' Grid1.Width = 10320
    
    End If
    
    If F_GRID_FONT_SIZE > 8 Then
        Grid1.Font.Size = F_GRID_FONT_SIZE
        TDBGrid1.Font.Size = F_GRID_FONT_SIZE
        GridPelaton.Font.Size = F_GRID_FONT_SIZE
    Else
        'Grid1.Font.Size = 12
        ' TDBGrid1.Font.Size = 12
        ' GridPelaton.Font.Size = 12
           
    End If
 
    ' mForm_Resize par1, 8, 13860, 10830, fh, fw, ft, fl

    Dim k
  '  Dim f_tab As Integer: f_tab = 1
'SSTab1.Tab = IIf(f_tab > SSTab1.Tabs - 1 Or f_tab < 0, SSTab1.Tabs - 1, f_tab)

Label1(0).Left = Text2(1).Left
label11(0).Left = Text2(2).Left
 Label1(0).Top = Label3(0).Top
  label11(0).Top = Label3(0).Top
   Label10(0).Top = Label3(0).Top
 
Label30.Left = 0
Label31.Left = 0
Label33.Left = 0
'sda_epistrofis.Top = apallagesFPA.Top - 600

'STREET.Left = 1080
'HMERPARAD.Left = 2040
'Combo2TyposParakr.Left = 240

'SynoloG.Left = Command7.Left
'Synolo.Left = Command7.Left
'sfpa.Left = Command7.Left
    'k = Label16.Left
    'ResizeFormFor Me
    'Label16.Left = k
    
    timtext3.Height = Grid1.RowHeight(1)
    '  timText3.FontSize = Grid1.Font.Size
   
    If timtext3.Height = Grid1.RowHeight(1) Then  'OK ΠΡΟΣΑΡΜΟΣΤΗΚΕ
    Else
        For k = 0 To Grid1.rows - 1
            Grid1.RowHeight(k) = timtext3.Height  ' Grid1.RowHeight(k) * Me.height / 10830
            '  Grid1.Font.Size = timText3.FontSize
        Next
    End If
    
    Label6.Left = Combo8.Left
    Label7.Left = APOTH1.Left
    Label6.Width = Label6.Width + 100
     Label7.Width = Label7.Width + 100
    
    
    
    
    
End Sub

Private Sub Grid1_DblClick()

        '<EhHeader>
        On Error GoTo Grid1_DblClick_Err

        ' Dim F33 As New APOT2
        ' F33.SHOW
        ' Exit Sub

        '</EhHeader>
        Dim F33 As New apot21
100     F33.WindowState = 2

102     timtext3.Visible = True
        ' kartella
104     F33.Text1.Text = Grid1.TextMatrix(Grid1.row, 1)

106     F33.Text2.Text = Grid1.TextMatrix(Grid1.row, 2)
108     F33.Label1.Caption = Left(APOTH1.Text, 1)
110     F33.Label2.Caption = "0"
        F33.filtroPEL.Caption = " PELKOD='" + Text2(0).Text + "' "

        'Me.Hide

112     F33.SHOW 1
114     Me.WindowState = 2

        ' APOT2.Hide

        '<EhFooter>
        Exit Sub

Grid1_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Grid1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Grid1_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
100     If KeyCode = 39 Then
    
        End If
    
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo Grid1_MouseUp_Err

        '</EhHeader>
        
        ' ΔΕΞΙ ΚΛΙΚ ΒΛΕΠΕΙ ΤΗΝ ΔΙΟΡΘΩΣΗ ΠΕΛΑΤΗ
100     If Button = 2 Then
102         Me.WindowState = 2
104         APOT2.kod.Text = MSFlexGrid1.Text   ' "10"
            ' Dim mapot2 As New APOT2
            '  APOT2.MDIChild = False
           
106         APOT2.kod.Text = Grid1.TextMatrix(Grid1.row, 1)
108         APOT2.kod.SetFocus
110         APOT2.ono.SetFocus
112         APOT2.SHOW

114         SendKeys "{ENTER}"
            
        End If

116     If Shift = 1 And Button = 1 Then
            Dim N As Integer
118         N = Grid1.row
            Dim k As Integer
  
            Dim s As Integer
120         For k = f_SeiresTimologioy To N Step -1
122             For s = 1 To Grid1.ColS - 1
124                 Grid1.TextMatrix(k, s) = Grid1.TextMatrix(k - 1, s)
                Next
      
            Next
  
126         For s = 1 To Grid1.ColS - 1
128             Grid1.TextMatrix(N, s) = ""
            Next
130         timtext3.Text = ""

        End If

        '<EhFooter>
        Exit Sub

Grid1_MouseUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Grid1_MouseUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Grid1_MouseUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_Scroll()

        '<EhHeader>
        On Error GoTo Grid1_Scroll_Err

        '</EhHeader>

100     timtext3.Visible = False

        '<EhFooter>
        Exit Sub

Grid1_Scroll_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Grid1_Scroll " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Grid1_Scroll " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GridEidon_DblClick()
100     take_eidos
     
End Sub

Private Sub GridEidon_GotFocus()

        ' ARXIKO Grid1.width = 5775
        '<EhHeader>
        On Error GoTo GridEidon_GotFocus_Err

        '</EhHeader>

100     GridEidon.Width = Combo4.Left - GridEidon.Left  ' 10320 ' Label18.Left - GridEidon.Left ' 10320

        '<EhFooter>
        Exit Sub

GridEidon_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.GridEidon_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.GridEidon_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GridEidon_LostFocus()

        '<EhHeader>
        On Error GoTo GridEidon_LostFocus_Err

        '</EhHeader>

100     GridEidon.Width = Label18.Left - GridEidon.Left '  10320 '  5775

        '<EhFooter>
        Exit Sub

GridEidon_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.GridEidon_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.GridEidon_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GRIDPARTIDES_DblClick()
        '======================================================================
        On Error Resume Next

        Dim EIDOSKPE As String
100     Dim cHME     As String: cHME = Format(DTPicker1.Value, "MM/DD/YYYY")
102     Dim ATIM     As String: ATIM = Right(PARAS.Text, 1) + Format(Text1.Text, String$(6, "0"))
104     EIDOSKPE = PEL.Recordset("eidos") + PEL.Recordset("KOD")
106     If PEL.Recordset("eidos") = "e" Then Exit Sub

        Dim M_PARTIDA As String
        Dim M_LHXH    As String
108     If DATAPARTIDES.Recordset.RecordCount > 0 Then
110         MsgBox "Σε κάθε σειρά δέχεται μόνο μία παρτίδα"
            Exit Sub
        End If

112     M_PARTIDA = InputBox("Παρτίδα ", "Δωσε την Παρτίδα", "..")
114     If M_PARTIDA = ".." Then
            Exit Sub
        End If

116     M_LHXH = InputBox("Λήξη Παρτίδα ", "Δωσε την Ημερ.Λήξης", "30/01/2010")

118     If Left(M_LHXH, 2) = "00" Then
            Exit Sub
        End If

120     If Len(M_LHXH) = 7 Or Len(M_LHXH) = 10 Then
122         If Len(M_LHXH) = 7 Then
124             If InStr(M_LHXH, "/") = 3 Then
126                 M_LHXH = Left(M_LHXH, 2) + IIf(Val(Left(M_LHXH, 2)) = 2, "/28/", "/30/") + Right(M_LHXH, 4)
                End If
            Else
128             M_LHXH = mID(M_LHXH, 4, 2) + "/" + Left(M_LHXH, 2) + "/" + Right(M_LHXH, 4)
            End If
        End If

        'CSBmk <Type the bookmark name here>

        Dim sql  As String
130     Dim POSO As String: POSO = Grid1.TextMatrix(Grid1.row, f_p)
        Dim m_ID As Long
132     If F_DIOR = 1 Then 'αν προκειται για διορθωση κραταει το id_num της διόρθωσης
134         m_ID = F_ID_NUM_DIOR
        Else
136         m_ID = 0
        End If

138     sql = "INSERT INTO PARTIDES(EIDOSKPE,KOD,LOTNUMBER,HMEL,YPOLOIPO,POSOPOL,POSOAG,ID_NUM,HME,ATIM,SEIRA) VALUES "
140     sql = sql + "('" + EIDOSKPE + "','" + Grid1.TextMatrix(Grid1.row, f_k) + "','" + M_PARTIDA + "','" + M_LHXH + "'," + POSO + "," + POSO + "," + POSO + "," + str(m_ID) + ",'" + cHME + "','" + ATIM + "'," + str(Grid1.row) + ")"
   
142     Gdb.Execute sql
144     GRIDPARTIDES.columns(0).Locked = False
        ' GRIDPARTIDES.Columns(1).Locked = False
        'GRIDPARTIDES.Columns(2).Locked = False
146     DATAPARTIDES.ConnectionString = gConnect
148     DATAPARTIDES.RecordSource = "select LOTNUMBER,CONVERT(CHAR(10),HMEL,3) AS [ΛΗΞΗ],YPOLOIPO AS [ΥΠΟΛ],POSOPOL AS [ΕΠΙΛ],ID FROM PARTIDES WHERE ID_NUM=0 AND SEIRA=" + str(Grid1.row) + " AND KOD='" + Grid1.TextMatrix(Grid1.row, f_k) + "'"
150     DATAPARTIDES.Refresh
 
        ' timText3.SetFocus

End Sub

Private Sub GRIDPARTIDES_HeadClick(ByVal ColIndex As Integer)
        'ΔΙΑΓΡΑΦΗ ΠΑΡΤΙΔΑς
   
        Dim a    As Long
        Dim ANS  As Integer
        Dim m_ID As Long
        Dim k    As Long

        ' to id to bazo sto telos
100     GRIDPARTIDES.Col = GRIDPARTIDES.columns.Count - 1
        'GRIDPARTIDES.row = 1
        'GRIDPARTIDES.Col = 3:
        ' Exit Sub
        'If KeyCode = 46 Then  ' αν δεν ειναι delete tote exit sub
102     If PEL.Recordset("eidos") = "r" Then
104         m_ID = Val(GRIDPARTIDES.Text)

106         GRIDPARTIDES.Col = 0
108         ANS = MsgBox("Να σβηστεί η  παρτίδα " + GRIDPARTIDES.Text, vbYesNo)

110         If ANS = vbYes Then
112             Gdb.Execute "delete from PARTIDES WHERE ID=" + str(m_ID), k
                Dim IID_NUM As Long
114             If F_DIOR = 1 Then
116                 IID_NUM = F_ID_NUM_DIOR
                Else
118                 IID_NUM = 0
                End If
   
120             GRIDPARTIDES.Visible = True
122             DATAPARTIDES.ConnectionString = gConnect
                'ΓΙΑ ΔΙΟΡΘΩΣΗ ΑΓΟΡΩΝ
124             DATAPARTIDES.RecordSource = "select LOTNUMBER,CONVERT(CHAR(10),HMEL,3) AS [ΛΗΞΗ],YPOLOIPO AS [ΥΠΟΛ],POSOPOL AS [ΕΠΙΛ],ID FROM PARTIDES WHERE ID_NUM=" + str(IID_NUM) + " AND SEIRA=" + str(Grid1.row) + "  AND KOD='" + Grid1.TextMatrix(Grid1.row, f_k) + "'"
126             DATAPARTIDES.Refresh
   
128             If k > 0 Then
130                 MsgBox "ok.Διεγράφη"
                End If
   
            End If

        End If
        'End If

End Sub

Private Sub GridPelaton_Click()
    ' mmGridPelaton_DblClick 1, 0, 0   ' 'TIMES0XORIS1 =0  DEFAULT      1=ΧΩΡΙΣ ΤΙΜΗ ( ΣΕ ΔΕΛΤΙΟ)
End Sub

Private Sub GridPelaton_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    'mmGridPelaton_DblClick 1, 0, 0   ' 'TIMES0XORIS1 =0  DEFAULT      1=ΧΩΡΙΣ ΤΙΜΗ ( ΣΕ ΔΕΛΤΙΟ)
End Sub

Private Sub GridPelaton_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'mmGridPelaton_DblClick 1, 0, 0   ' 'TIMES0XORIS1 =0  DEFAULT      1=ΧΩΡΙΣ ΤΙΜΗ ( ΣΕ ΔΕΛΤΙΟ)
End Sub

Private Sub isotimia_LostFocus()

        Exit Sub

        '<EhHeader>
        On Error GoTo isotimia_LostFocus_Err

        '</EhHeader>

        Dim k As Integer

        Dim mISOT

100     mISOT = gVal(isotimia.Text)

102     If mISOT <> 1 Then

104         isotimia.Text = mISOT

106         For k = 1 To Grid1.rows - 1

108             If Val(Grid1.TextMatrix(k, f_p)) <> 0 Then
110                 Grid1.TextMatrix(k, F_T) = Val(Grid1.TextMatrix(k, F_T)) / mISOT
112                 Grid1.row = k: FindSynolo
                End If

            Next

        End If

        '<EhFooter>
        Exit Sub

isotimia_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.isotimia_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.isotimia_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub kentroAdyn_LostFocus()
  
100     Text2(1).Enabled = True
102     Text2(1).SetFocus
End Sub

Private Sub kerdos_Click()

        '<EhHeader>
        On Error GoTo kerdos_Click_Err

        '</EhHeader>

100     MsgBox get_kerdos()

        '<EhFooter>
        Exit Sub

        '' TOPQUARK@OTENET.GR

kerdos_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.kerdos_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.kerdos_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label12_Click()
100     bohu3.defPar.Text = Right(PARAS.Text, 1)
102     bohu3.SHOW
        ' bohu3.defPar.Text = Right(PARAS.Text, 1)
End Sub

Private Sub Label18_Click()

        '<EhHeader>
        On Error GoTo Label18_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 12, Label18, Combo4

        '<EhFooter>
        Exit Sub

Label18_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Label18_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Label18_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label19_Click()

        '<EhHeader>
        On Error GoTo Label19_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 7, Label9, SKOPOS

        '<EhFooter>
        Exit Sub

Label19_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Label19_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Label19_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label20_Click()

        '<EhHeader>
        On Error GoTo Label20_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 8, Label20, fortosh

        '<EhFooter>
        Exit Sub

Label20_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Label20_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Label20_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label21_Click()

        '<EhHeader>
        On Error GoTo Label21_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 9, Label21, PROORISMOS

        '<EhFooter>
        Exit Sub

Label21_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Label21_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Label21_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label5_Click()

        '<EhHeader>
        On Error GoTo Label5_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 10, Label5, AYTOKINHTO

        '<EhFooter>
        Exit Sub

Label5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Label5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Label5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label6_Click()

        '<EhHeader>
        On Error GoTo Label6_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 5, Label6, Combo8

        '<EhFooter>
        Exit Sub

Label6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Label6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Label6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label7_Click()

        '<EhHeader>
        On Error GoTo Label7_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 4, Label7, APOTH1

        '<EhFooter>
        Exit Sub

Label7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Label7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Label7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub MONTE_CRISTO_Click()
        'Text2(0).SetFocus
        'SendKeys "[enter]"
        '   SendKeys "{ENTER} {ENTER}{ENTER} {ENTER}"
   
100     POLAGOR.ListIndex = 1 ' ΠΩΛΗΣΕΙΣ
102     PARAS.ListIndex = Val(Left(MONTE_CRISTO.Caption, 2))
104     PARAS.Enabled = True
106     PARAS.SetFocus

108     DoEvents
   
        'ΚΑΝΟΝΙΖΕΙ ΤΟ ΠΟΣΟ
110     Text5.Text = "*" + Right(MONTE_CRISTO.Caption, 6)
112     Text5_LostFocus
   
114     timtext3.SetFocus
End Sub

Private Sub MSFlexGrid1_Click()

        '<EhHeader>
        On Error GoTo MSFlexGrid1_Click_Err

        '</EhHeader>
100     If Left(f_pelMERC, 3) = "SOK" Then
102         MSFlexGrid1.Col = 1
104         timtext3.Text = MSFlexGrid1.Text
106         timtext3.SetFocus

        End If

        '<EhFooter>
        Exit Sub

MSFlexGrid1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.MSFlexGrid1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.MSFlexGrid1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub par2show_Click()
'    PAR2SHOWING
'End Sub

Private Sub POLAGOR_Change()

        '<EhHeader>
        On Error GoTo POLAGOR_Change_Err

        '</EhHeader>

100     PARAS.Enabled = True

        '<EhFooter>
        Exit Sub

POLAGOR_Change_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.POLAGOR_Change " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.POLAGOR_Change " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub POLAGOR_Click()

        '<EhHeader>
        On Error GoTo POLAGOR_Click_Err

        '</EhHeader>

100     PARAS.Enabled = True

        '<EhFooter>
        Exit Sub

POLAGOR_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.POLAGOR_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.POLAGOR_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub POLAGOR_GotFocus()

        '<EhHeader>
        On Error GoTo POLAGOR_GotFocus_Err

        '</EhHeader>

100     POLAGOR.BackColor = vbYellow
102     F_POLAGOR_LastValue = POLAGOR.Text
        'Me.KeyPreview = False

        '<EhFooter>
        Exit Sub

POLAGOR_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.POLAGOR_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.POLAGOR_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub POLAGOR_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo POLAGOR_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = 13 Then
102         PARAS.Enabled = True
104         PARAS.SetFocus
        End If

        '<EhFooter>
        Exit Sub

POLAGOR_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.POLAGOR_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.POLAGOR_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub POLAGOR_LostFocus()

        ' ANALOGA ME TI DIALEJA ΑΝΟΙΓΕΙ ΤΟ PARASTAT KAI GEMIZEI TO PARAS
        '<EhHeader>
        On Error GoTo POLAGOR_LostFocus_Err

        '</EhHeader>

        Dim X As String, k As Integer

        'Dim db As Database
        Dim R As New ADODB.Recordset

        ' db.Open gDir

        'If gConnect = "Access" Then
        'Set db = OpenDatabase(gDir, False, False)
        'Else
        'Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If
100     POLAGOR.BackColor = vbWhite
        Dim M_FBUFF As String
102     M_FBUFF = fbuff
104     If POLAGOR.Text = POLAGOR.List(0) Then
106         X = "2": fbuff = "r"    'προμηθευτές
108         SKOPOS.Text = SKOPOS.List(8)
110         selectPEL.Caption = "Προμηθευτής"
112         f_SeiresTimologioy = f_MaxSeiresTimologioy    'POSA EIDH MPORO NA PERASO Max TIMOLOGIA Agoras
            'f_SeiresTimologioy = 30 OTAN DIALEGO  POLISEIS

114         Combo4.Text = Combo4.List(1)    ' PISTVSH
            
116     ElseIf POLAGOR.Text = POLAGOR.List(1) Then
118         selectPEL.Caption = "Πελάτης"
120         X = "1": fbuff = "e"    'pelates
122         f_SeiresTimologioy = f_MinSeiresTimologioy    'POSA EIDH MPORO NA PERASO Max tim .polisis
124         SKOPOS.Text = SKOPOS.List(0)    ' ΠΩΛΗΣΗ
            'Text2(0).Text = ""
        Else
126         X = "3": fbuff = "r"    'λογαριασμοί
128         f_SeiresTimologioy = f_MaxSeiresTimologioy    'POSA EIDH MPORO NA PERASO Max TIMOLOGIA Agoras
            ' Text2(0).Text = ""
        End If
         
        ' ΑΝ ΑΛΛΑΞΕ ΚΑΤΗΓΟΡΙΑ ΝΑ ΣΒΗΝΕΙ  ΤΟΝ ΚΩΔΙΚΟ ΤΟΥ ΠΕΛΑΤΗ
130     If M_FBUFF <> fbuff Then
132         Text2(0).Text = ""
        End If

134     If Val(X) = 0 Then X = "1"

        'Set R = OpenRecordset("select pol,titlos,eidos from parastat where pol like '" + x + "%';")

        ' θα κανειrefresh τα παραστατικά μόνο άν άλλαξε
        ' ή βρισκόμαστε στην πρώτη φορά

136     If F_POLAGOR_LastValue <> POLAGOR.Text Or Len(PARAS.Text) = 0 Then    ' MONO THN 1H FORA NA ALLAZEI
            ' If Left(APOTH1.Text) = "1" Then
138         R.Open "select POL,TITLOS,EIDOS from PARASTAT where POL='" + X + "'  AND YP" + Left(APOTH1.Text, 1) + "=1 order by TITLOS", Gdb, adOpenDynamic, adLockOptimistic

140         PARAS.Clear

142         If R.RecordCount = 0 Then

                Exit Sub

            End If

144         R.MoveFirst
146         k = 0

148         Do While Not R.EOF

150             If Not IsNull(R("TITLOS")) And R("pol") = X Then
152                 If F_STOPDELTIA = 99 Then
154                     If R("eidos") = "t" Or R("eidos") = "l" Then
156                         PARAS.AddItem R("TITLOS") + Space(1) + R("eidos")
158                         k = k + 1
                        End If

                    Else
160                     PARAS.AddItem R("TITLOS") + Space(1) + R("eidos")
162                     k = k + 1
                    End If

164                 mEidos_Par(k) = R("eidos")
                End If

166             R.MoveNext
            Loop

168         PARAS.Enabled = True

            On Error Resume Next

170         If X = "2" Then    'agores
172             PARAS.Text = PARAS.List(f_1choice_agores)
174         ElseIf X = "1" Then    'polhseis
                ' PARAS.Text = PARAS.List(0)
176             PARAS.Text = PARAS.List(f_1choice_polhseis)

            Else
178             PARAS.Text = PARAS.List(1)
            End If

        End If

180     DTPicker1.Value = Now

        'PARAS.Enabled = True

182     F_POLAGOR_LastValue = POLAGOR.Text

        ' Me.KeyPreview = True

184     X = mEidos_Par(1 + PARAS.ListIndex)

186     If F_DIOR <> 1 Then
188         Text1.Text = find_parastat(1, X, 0)
        End If

        'db.Close

        '<EhFooter>
        Exit Sub

POLAGOR_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.POLAGOR_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.POLAGOR_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub PARAS_GotFocus()

        '<EhHeader>
        On Error GoTo PARAS_GotFocus_Err

        '</EhHeader>

100     PARAS.BackColor = vbYellow

        '<EhFooter>
        Exit Sub

PARAS_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.PARAS_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.PARAS_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub PARAS_KeyPress(KeyAscii As Integer)
        
        On Error GoTo PARAS_KeyPress_Err
        
100     If KeyAscii = 13 Then
102         If Text1.Enabled = False Then
104             selectPEL.SetFocus
            Else
106             Text1.SetFocus   ' SEIRA.SetFocus
            End If
        End If
        
        Exit Sub

PARAS_KeyPress_Err:
        
108     SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.PARAS_KeyPress " & "at line " & Erl

110     Resume Next

        '</EhFooter>

End Sub

Private Sub PARAS_LostFocus()

        ' STELNEI STO FIND_PARASTAT NA BREI TON A/A KAI PAIRNEI TIS PARAMETROYS APO TO PARASTAT
        '<EhHeader>
        On Error GoTo PARAS_LostFocus_Err

        '</EhHeader>

        Dim k As Integer

        'Dim db As Database
        Dim R As New ADODB.Recordset

        Dim X As String

        On Error GoTo MHNYMA
 
        'βαζει αποθηκη την προελευση
100     If F_PIN_PROELEYS = 11 Then
     
102         Grid1.TextMatrix(0, F_PROEL) = "Αποθήκη"
104         If F_STADIO = 5 Then    ' ενδοδιακινηση
106             Label7.Visible = True
108             APOTH1.Visible = True
110             APOTH2.Visible = True
            Else
112             Label7.Visible = False
114             APOTH1.Visible = False
116             APOTH2.Visible = False
            End If
        End If
        
        Dim m_FormEktyp As Integer

118     PARAS.BackColor = vbWhite
        'db.Open gDir
        'R.Open "select *from parastat;", db, adOpenKeyset, adLockReadOnly ', adCmdTableDirect

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If
120     R.Open "SELECT ISNULL(PROEPILFPA,0) AS PROEPILFPA2,ISNULL(DESMIA,'0') AS DESMIA2,*FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

122     R.MoveFirst

124     Do While Not R.EOF

126         If Right(PARAS.Text, 1) = R("EIDOS") Then
128             F_TITLOS = R("TITLOS")
130             F_POS_APOU = R("POS_APOU")
132             F_AJIA_APOU = R("AJIA_APOU")
134             f_pel = R("pel")
136             f_pol = R("pol")    ' 1=POLISEIS 2=AGORES
138             F_STADIO = R("STADIO")    ' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
140             f_AkyrParas = Left(LTrim(R("akyr")), 1)
142             F_TYPOS_PAR = CNull(R("C1")) ' ΤΥΠΟΣ ΠΑΡΑΣΤΑΤΙΚΟΥ ΓΙΑ Ε_ΛΙΝΕ ΦΟΡΟΛΟΓΙΚΟΥ
                F_PROEPILFPA = R("PROEPILFPA2")
                
                
                
                

                  If nNull(R("B2G")) = 1 Then ' is_B2G = GGET_NVALUE("SELECT ISNULL(B2G,0) AS DIMOSIO FROM PARASTAT WHERE EIDOS='" & Left(sqlDt("ATIM"), 1) & "'")
                      MsgBox ("τα τιμολογια ΔΗΜΟΣΙΟΥ ΕΚΔΩΣΤΕ ΤΑ ΑΠΟ ΤΟ ΠΡΟΓΡΑΜΜΑ ΤΟΥ ΠΑΡΟΧΟΥ")
                      
                  End If

144             If IsNull(R("GEF_P")) Then
146                 F_rec_p = 0
                Else
148                 F_rec_p = R("GEF_P")    ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
                End If

150             If IsNull(R("SET_EIDON")) Then
152                 F_SET_EIDON = 0
                Else
154                 F_SET_EIDON = R("SET_EIDON")    ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
                End If

156             If IsNull(R("POLITES")) Then
158                 F_POLHTES_PERIOR = ""
                Else
160                 F_POLHTES_PERIOR = R("POLITES")    ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
                End If
            f_isdiakin = nNull(R("ISDIAKIN"))
            Dim LDIAK As String
            If f_isdiakin = 1 Then
               LDIAK = "ΑΞΙΑ+ΔΙΑΚΙΝΗΣΗ"
            ElseIf f_isdiakin = 2 Then
               LDIAK = "ΔΙΑΚΙΝΗΣΗ MONO"
            Else
               LDIAK = "ΟΧΙ ΔΙΑΚΙΝΗΣΗ"
            End If
            
            MDIForm1.StatusBar1.Panels(5).Text = LDIAK
            
            
162             F_FORMA1 = Trim(CNull(R("FORMA1")))
164             F_FORMA2 = Trim(CNull(R("FORMA2")))
166             F_FORMA3 = Trim(CNull(R("FORMA3")))
                
168             If nNull(R("DEFAULTAPALL")) > 0 Then
                   
170                 apallagesFPA.Text = apallagesFPA.List(-1 + nNull(R("DEFAULTAPALL")))
                End If
            
172             f_myEID = CNull(R!MYEID)  ' 1.1 τιμολογιο MYDATA

                 If f_myEID = "10.1" Then
                 
                    MsgBox ("ΣΤΑ ΣΧΟΛΙΑ ΓΡΑΨΤΕ ΣΤΗΝ ΑΡΧΗ ΤΟ ΜΑΡΚ ΤΟΥ ΔΕΛΤΙΟΥ ΑΠΟΣΤ.ΠΕΛΑΤΗ")
                 
                 End If



            
174             f_ISMYDATA = nNull(R!ISMYDATA)
            
176             F_YPOX_AFM = IIf(nNull(R!GEF_M) = 99, 1, 0)
            
178             If IsNull(R("PARAKRATISI")) Then
180                 F_PARAKRATISI = 0
                    ' parakrat.Visible = False
182                 parakratisiT.Visible = False
184                 tel_pliroteo.Visible = False
186                 lblΠαρακράτηση.Visible = False
188                 lblΠληρωτέο.Visible = False
190                 ComboPARAKRATISI.Visible = False
192                 Combo2PARAKRATISI.Visible = False
                    
                    '  LBLTEL.Visible = False
                    
                Else
                
194                 If R("PARAKRATISI") = 0 Then
196                     F_PARAKRATISI = 0
                        ' parakrat.Visible = False
198                     parakratisiT.Visible = False
200                     tel_pliroteo.Visible = False
202                     lblΠληρωτέο.Visible = False
204                     ComboPARAKRATISI.Visible = False
206                     Combo2PARAKRATISI.Visible = False
                       
                        'Combo2PARAKRATISI.Text = Combo2PARAKRATISI.List(2)
                        'ComboPARAKRATISI.Text = ComboPARAKRATISI.List(5)
                      
208                     lblΠαρακράτηση.Visible = False
                        ' LBLTEL.Visible = False
                    Else
                        ' parakrat.Visible = True
210                     parakratisiT.Visible = True
212                     tel_pliroteo.Visible = True
214                     lblΠληρωτέο.Visible = True
216                     lblΠαρακράτηση.Visible = True
218                     ComboPARAKRATISI.Visible = True
220                     Combo2PARAKRATISI.Visible = True
                        
222                     Combo2PARAKRATISI.Text = Combo2PARAKRATISI.List(2)
224                     ComboPARAKRATISI.Text = ComboPARAKRATISI.List(5)
                        
226                     If IsNull(R("LPARAKRAT")) Then  'ΔΕΝ ΕΧΩ ΠΡΟΕΠΙΛΕΓΜΕΝΟ
                          
                        Else
228                         If InStr(R("LPARAKRAT"), ";") > 0 Then
                          
                                Dim ARR22()   As String
                                Dim ARR23(10) As String
230                             ARR22 = Split(R("LPARAKRAT"), ";")

232                             If UBound(ARR22) > 0 Then

                                    ' For k = 0 To UBound(ARR)
                                    '     ARR23(k) = Combo8.List(Val(ARR(k)))
                                    ' Next
234                                 If UBound(ARR22) >= 1 Then
236                                     Combo2PARAKRATISI.Text = Combo2PARAKRATISI.List(Max(0, ARR22(0) - 1)) ''         'Val(Split(R("LPARAKRAT"), ";")(0)) - 1)
238                                     ComboPARAKRATISI.Text = ComboPARAKRATISI.List(Max(ARR22(1) - 1, 0)) 'List(Val(Split(R("LPARAKRAT"), ";")(1)) - 1)
                                    End If
                                 
240                                 If UBound(ARR22) >= 3 Then
242                                     Combo2TyposParakr.Text = Combo2TyposParakr.List(Max(ARR22(2) - 1, 0)) ''         'Val(Split(R("LPARAKRAT"), ";")(0)) - 1)
244                                     Combo2OnomaParak.Text = Combo2OnomaParak.List(Max(ARR22(3) - 1, 0)) 'List(Val(Split(R("LPARAKRAT"), ";")(1)) - 1)
                                    End If
                                  
                                End If
                              
                            End If
                        
                        End If
                        
                        ' LBLTEL.Visible = True
246                     F_PARAKRATISI = R("PARAKRATISI")    ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
                        
248                     If R("PARAKRATISI") > 1 Then
250                         SSTab1.Visible = True
                        End If
                        
                    End If
                End If

252             If IsNull(R("XONDR")) Then
254                 F_XONDR = 1
                Else
                   If R("XONDR") = 0 Then
                       F_XONDR = 1
                   Else
                      F_XONDR = R("XONDR")
                   End If
                  
256             '    F_XONDR = R("XONDR")    ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
258             '    If F_XONDR = 2 Then F_XONDR = 1

                End If
                
260             If F_LIANIKES_TIMES = 0 Then
                    '   F_XONDR = 1
                End If
                
262             If IsNull(R("FORM_EKTYP")) Then
264                 f_FORM_EKTY = 0
                Else
266                 f_FORM_EKTY = Val(R("FORM_EKTYP"))    ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
                End If
                
268             If F_2EKPTOSEIS = 1 Then
270                 If f_FORM_EKTY > 0 Then  ' Or F_2EKPTOSEIS = 3 Then
272                     f_f = 6       '  6  ''ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
274                     F_EKPT2 = 11    '  11   'ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
                    Else
276                     f_f = 11       '  6  ''ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
278                     F_EKPT2 = 6    '  11   'ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
                    End If
280                 Grid1.row = 0
282                 Grid1.Col = f_f
284                 Grid1.Text = "ΚΑΤ.ΦΠΑ"
286                 Grid1.Col = F_EKPT2
288                 Grid1.Text = "ΕΚΠΤ 2"
                End If
                
                
               ' If IsNull(R!DESMIA) Then
                '    F_DESMIA_SEIRES = 0
              '  Else
                    F_DESMIA_SEIRES = Val(R!DESMIA2)
               ' End If
                
                'ΑΝ ΧΡΗΣΙΜΟΠΟΙΕΙ ΤΗΝ ΦΟΡΜΑ ΤΙΜΟΛ2 TOTE NA DINEI ALLES SEIRES APO TO F_FORM2_SEIRESTIMOL
290             If F_DESMIA_SEIRES = 2 Or f_FORM_EKTY = 14 Or f_FORM_EKTY = 12 Then
292                 f_SeiresTimologioy = f_Timol2_SeiresTimologioy
294                 f_SynSeiresTimologioy = f_Timol2_SeiresTimologioy
                End If

                '       F_rec_m = R("GEF_M") ' ........  GIA METRITA

296             If IsNull(R("GEF_M")) Then
298                 F_rec_m = 0
                Else
300                 F_rec_m = R("GEF_M")    ' ΠΟΙΑ ΚΑΤΗΓΟΡΙΑ ΘΑ ΕΠΙΤΡΕΠΕΤΑΙ ΝΑ ΚΟΒΕΙ 0=ΟΛΟΙ   RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
                End If

302             If IsNull(R("aytok")) Then
304                 AYTOKINHTO.Text = AYTOKINHTO.List(0)
                Else
306                 AYTOKINHTO.Text = AYTOKINHTO.List(R("aytok"))
                End If

308             f_metasx = R("metasx")    ' απο ποιό παραστατικό μετασχηματίζεται
            End If

310         R.MoveNext
        Loop

312     R.Close

314     If F_STADIO = 5 Then  '' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
316         APOTH2.Enabled = True
        Else
318         APOTH2.Enabled = False
        End If

        'PARAS.Clear
320     geidos_paras = PARAS.ListIndex

322     If F_STADIO = 2 Then
324         Label14.Caption = "MarkUp %"
        Else
326         Label14.Caption = "Μεταφορικά και άλλα έξοδα"
        End If

328     If F_PERIOR_POLHTES = 1 Then '

            'ΦΟΡΤΩΜΑ ΠΩΛΗΤΩΝ
  
330         Combo8.Clear
332         R.Open "SELECT *FROM PINAKES WHERE TYPOS=5 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

334         Do While Not R.EOF
336             Combo8.AddItem LTrim(str(R("AYJON"))) + "." + R("PERIGRAFH")
338             R.MoveNext
            Loop

340         R.Close
342         Combo8.Text = Combo8.List(0)
         
            Dim ARR()    As String

            Dim ARR2(10) As String

344         ARR = Split(F_POLHTES_PERIOR, ",")

346         If UBound(ARR) > 0 Then

348             For k = 0 To UBound(ARR)
350                 ARR2(k) = Combo8.List(Val(ARR(k)))
                Next

352             Combo8.Clear

354             For k = 0 To UBound(ARR)
356                 Combo8.AddItem ARR2(k)
                Next

            End If
        End If  'F_PERIOR_POLHTES = 1 Then

358     X = mEidos_Par(1 + PARAS.ListIndex)

360     If (F_DIOR <> 1) Or F_METASX_PAR = 1 Then

362         Text1.Text = find_parastat(k, X, 0)
'              If GGET_NVALUE("select TOP 1 ISNULL(ARITMISI,0) from PARASTAT WHERE EIDOS='" + Right(PARAS.Text, 1) + "'") > 0 Then
'                     Dim LASTOK As String
'                      LASTOK = GGET_CVALUE("select TOP 1 ATIM FROM TIM WHERE LEFT(ATIM,1)='" + Right(PARAS.Text, 1) + "' ORDER BY ATIM DESC")
'                      'Α.ΒΡΙΣΚΩ ΤΟ ΤΕΛΕΥΤΑΙΟ ΠΑΡ/ΚΟ ΑΥΤΗΣ ΤΗΣ ΣΕΙΡΑΣ ΚΑΙ 1.UPDATE ARITMISI   2.BAZV THN ΣΩΣΤΗ ΑΡΙΘΜΗΣΗ ΣΤΟ ΤΙΜΟΛΟΓΙΟ
'                     'Β.ΣΕ ΠΕΡΙΠΤΩΣΗ ΠΟΥ ΕΙΝΑΙ ΤΟ ΠΡΩΤΟ ΤΗΣ ΧΡΟΝΙΑΣ ΜΕ ΡΩΤΑ
'                    If LASTOK = "" Then ' B
'                           Dim ANS As Integer
'                           ANS = MsgBox("ΔΕΝ ΥΠΑΡΧΕΙ Ο ΠΡΟΗΓΟΥΜΕΝΟΣ ΑΡΙΘΜΟΣ.ΝΑ ΣΥΝΕΧΙΣΩ;", vbYesNo)
'
'                               If ANS = vbYes Then
'                                      'Check_Dipla = 1
'                               Else
'                                      'Check_Dipla = 0
'                               End If
'                     Else ' A .PERIPTOSI
'
'                       If Val(Text1.Text) - Val(mID$(LASTOK, 2, 6)) <> 1 Then
'                            Gdb.Execute "UPDATE ARITMISI SET ARITMISI=" + Format(Val(mID(LASTOK, 2, 6)), "#####0") + " WHERE ID=(SELECT ARITMISI FROM PARASTAT WHERE EIDOS='" + Left(f_matim, 1) + "') "
'
'                           ' f_matim = Left(f_matim, 1) + Format(Val(mID(LASTOK, 2, 6)) + 1, "000000")
'                             Text1.Text = Format(Val(mID(LASTOK, 2, 6)) + 1, "#####0")
'
'                       End If
'
'
'                      ' ")
'
'                     End If
'             End If
'



        End If

        'αν κατα την διορθωση αλλαξε και τυπο παραστατικου τοτε να βρει την νεα αριθμηση




364     If F_DIOR = 1 And f_dior_ARX_TYPOS <> PARAS.Text Then
366         Text1.Text = find_parastat(k, X, 0)
        End If

368     Text1.Enabled = True
370     If f_FORM_EKTY > 0 Then
372         Text1.Enabled = False
            
374         If F_DIOR = 0 Then
                ' 'ΨΑΧΝΩ ΝΑ ΔΩ ΑΝ ΥΠΑΡΧΕΙ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ ΗΔΗ ( ΑΠΟ ΛΑΘΟΣ)
                Dim r60 As New ADODB.Recordset
376             r60.Open "select * FROM TIM WHERE ATIM='" + Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text + "'", Gdb, adOpenDynamic, adLockOptimistic
378             If r60.EOF Then
                    ' OK
                Else
                    'ΞΑΝΑΥΠΑΡΧΕΙ ΤΟ ΤΙΜΟΛΟΓΙΟ ΟΠΟΤΕ ΠΑΙΡΝΕΙ ΤΟ ΕΠΟΜΕΝΟ
380                 r60.Close
382                 r60.Open "select MAX(ATIM)  FROM TIM WHERE LEFT(ATIM,1)='" + Right(PARAS.Text, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
384                 Text1.Text = LTrim(str(Val(mID(r60(0), 2, 6))))
                    Dim DUM As Long
386                 DUM = find_parastat(1, X, 1)  'ΑΠΟΘΗΚΕΥΩ ΤΟΝ ΑΡΙΘΜΟ   ΣΤΟ ΑΡΙΤΜΙΣΙ
388                 Text1.Text = str(Val(Text1.Text) + 1)
                End If
                 
390             r60.Close
            End If
        
        End If

        'PACKING LIST
392     If F_KOLA_MIKTA_LABELS = 2 Then   ' marmaradika
               
394         If InStr("μα", X) > 0 Then  ' προκειται για παψκινγ λιστ
396             f_k = 1: f_o = 2: F_PROEL = 3: F_COL_KOLA = 4: F_mikta_KILA = 5: f_p = 6: F_T = 7: f_e = 8: f_f = 9: f_a = 10: F_M = 11: F_mefpa = 12: F_EKPT2 = 13: F_mikta = 14
398             Grid1.TextMatrix(0, F_COL_KOLA) = "ΠΛΑΤΟΣ"
400             Grid1.TextMatrix(0, F_mikta_KILA) = "ΠΑΧΟΣ"
402             Grid1.TextMatrix(0, f_p) = "ΠΟΣΟΤΗΤΑ"
404             Grid1.TextMatrix(0, F_PROEL) = "ΜΗΚΟΣ"
406             F_KOLA_MIKTA = 1
                 
            Else
408             f_k = 1: f_o = 2: f_p = 3: F_T = 4: f_e = 5: F_EKPT2 = 6: f_a = 7: F_M = 8: F_mefpa = 9: F_PROEL = 10: f_f = 11: F_mikta = 12: F_mikta_KILA = 13: F_COL_KOLA = 14
410             f_colEFK = 15
412             F_KOLA_MIKTA = 0
               
            End If

        End If


414     If Right(PARAS.Text, 1) = Right("  " + Trim(f_SDA), 1) Then
416         sda_epistrofis.Enabled = True
         Else

           sda_epistrofis.Enabled = False
           
        End If

418     If Val(Text1.Text) > 0 Then
420         selectPEL.SetFocus
        End If

422     PARAS.Enabled = False



FORTDIE.Text = f_peldie
fortARit.Text = f_pelarit
FORTTK.Text = f_peltk
FORTPOL.Text = f_pelpol
        
        'ΑΝ ΕΙΝΑΙ ΒΥΤΙΟ ΚΑΙ ΕΧΕΙ ΚΑΙ ΜΕΤΑΦΟΡΕΣ ΝΑ ΑΛΛΑΖΕΙ Ο ΤΙΤΛΟΣ
424     If F_IS_BYTIO = 3 Or F_IS_BYTIO = 3 Then
              
426         If Right$(PARAS, 1) = F_PARAS_3SYNAL Then
              
428             Frame2(0).Top = Frame2(1).Top
430             Frame2(0).Height = Abs(Frame2(0).Top - Grid1.Top)
432             Frame2(0).Width = Frame2(1).Width
434             Frame2(0).Left = Frame2(1).Left
               
436             Frame2(1).Visible = False
438             Frame2(0).Visible = True
                ' Frame2(0).Height = Frame2(0).Height + 300
440             Grid1.Top = 2820 ' 2520
442             Grid1.Height = 2000 ' 2295
444             MTF11(1).Caption = ""
446             MTF12(0).Caption = "-ΟΝΟΜΑ ΑΠΟΣΤ"
448             MTF13(0).Caption = "-ΑΠΟΣΤΟΛΕΑΣ" 'As String
450             MTF14(0).Caption = "-ΟΝΟΜΑ ΠΑΡΑΛ"  'As String
452             MTF21(0).Caption = "-ΠΑΡΑΛΗΠΤΗΣ"  'As String
                'MON  =>C2  Textm(2).
                'POSOTHTA =>num1  => Textm(1)
                'ARIU.PARAST =>c12 => Textm(6)
                'PERIGRAFH =>c13 =>   Textm(7)
                'TEXTM(3)  APOSTOLEAS=>C3   TEXTM(5) ONOMA APOSTOLEA
                'TEXTM(0)  PARALHPTHS=>C0   TEXTM(4) ONOMA PARALHPTH
454             Textm(2).Visible = False
456             Textm(7).Visible = False
458             Textm(1).Visible = False
460             Textm(6).Visible = False
                  
462             Labf(0).Visible = False
464             Labf(1).Visible = False
466             Labf(2).Visible = False
468             Labf(3).Visible = False
                  
470             Form_Resize
                  
            Else
472             Frame2(1).Visible = True
474             Frame2(0).Visible = False
476             Frame2(1).Height = Abs(Frame2(1).Top - Grid1.Top)
              
            End If
        
        End If
        
        Exit Sub

MHNYMA:
478     HandleError "Par1:paras_lostfocus"

480     Resume Next

        '<EhFooter>
        Exit Sub

PARAS_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.PARAS_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.PARAS_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub sda_epistrofis_Click()

        Dim k             As Integer
        Dim R5            As New ADODB.Recordset

        Dim atim_sda_fort As String
        Dim id_num_sda    As Long
        Dim M6KOD         As String
        Dim R6            As New ADODB.Recordset
100     id_num_sda = GET_NVALUE("SELECT ISNULL(ID_NUM,0) AS ID_NUM FROM TIM WHERE LEFT(ATIM,1)='" + Left(f_SDA + " ", 1) + "' ORDER BY ID_NUM DESC")
        Dim ARX   As String, TEL As String

        ' ΓΙΑ ΝΑ ΞΕΧΩΡΙΖΕΙ ΣΤΑ ΣΧΟΛΙΑ ΤΑ ΤΙΜΟΛΟΓΙΑ ΑΠΟ ΤΙΣ ΛΙΑΝΙΚΕΣ
        Dim MLIAN As String
102     MLIAN = Split(F_PAR_AYTOK, ",")(0)
        Dim MTIM As String
104     MTIM = Split(F_PAR_AYTOK, ",")(1)

106     ARX = GET_CVALUE("select TOP 1 isnull(ATIM,'       ') AS TT FROM TIM WHERE LEFT(ATIM,1) IN (" + MLIAN + ") AND ID_NUM>" + str(id_num_sda) + " ORDER BY ID_NUM ")
108     TEL = GET_CVALUE("select TOP 1 ISNULL(ATIM,'       ') AS TT FROM TIM WHERE LEFT(ATIM,1) IN (" + MLIAN + ") ORDER BY ID_NUM DESC")
110     TEL = str(Val(mID(TEL, 2, 6)))
112     ARX = str(Val(mID(ARX, 2, 6)))
114     If Val(ARX) > 0 Then
116         parat.Text = "ΑΠΛ. από " + mID(ARX, 2, 6) + " έως " + mID(TEL, 2, 6)
        End If

118     ARX = GET_CVALUE("select TOP 1 isnull(ATIM,'       ') AS TT FROM TIM WHERE LEFT(ATIM,1) IN (" + MTIM + ") AND ID_NUM>" + str(id_num_sda) + " ORDER BY ID_NUM ")
120     TEL = GET_CVALUE("select TOP 1 ISNULL(ATIM,'       ') AS TT FROM TIM WHERE LEFT(ATIM,1) IN (" + MTIM + ") ORDER BY ID_NUM DESC")
122     TEL = str(Val(mID(TEL, 2, 6)))
124     ARX = str(Val(mID(ARX, 2, 6)))

126     If Val(ARX) > 0 Then
128         parat.Text = parat.Text + " TIM. από " + mID(ARX, 2, 6) + " έως " + mID(TEL, 2, 6)
        End If

        Dim POL_POS As Single
                    
130     k = 1
132     R5.Open "SELECT * FROM EGGTIM WHERE ID_NUM=" + str(id_num_sda) + " ORDER BY ID ", Gdb, adOpenDynamic, adLockOptimistic

134     Do While Not R5.EOF
              
136         Grid1.TextMatrix(k, f_k) = R5("kode")
138         Grid1.TextMatrix(k, f_p) = R5("poso")
140         Grid1.TextMatrix(k, f_e) = R5("ekpt")
                        
            '  F_PAR_AYTOK=    'L','T'
142         POL_POS = GET_NVALUE("select sum(POSO) FROM EGGTIM WHERE KODE='" + R5("KODE") + "' AND  LEFT(ATIM,1) IN (" + F_PAR_AYTOK + ") And ID_NUM > " + str(id_num_sda)) ' χρειαζονται για το συγκεντροτικο"
                        
            'πωληθεντα
            '  On Error GoTo 0
144         Grid1.TextMatrix(k, F_mikta) = POL_POS
                        
            'ypoloipa
            'Grid1.TextMatrix(k, F_mikta_KILA) = 0   ' χρειαζονται για το συγκεντροτικο
                        
            ' Grid1.TextMatrix(k, F_T) = R5!TIMM
              
146         If IsNull(R5!KODE) Then
148             M6KOD = ""
            Else
150             M6KOD = R5!KODE
            End If
              
152         R6.Open "SELECT * FROM EID WHERE KOD='" + M6KOD + "'", Gdb, adOpenDynamic, adLockOptimistic

154         If Not R6.EOF Then
                  
                '  If F_EISITHRIA = 0 Then
156             Grid1.TextMatrix(k, f_o) = R6("ono") ' Left(R6("ono") + Space(60), 60)  '+ m_a
                '  Else
                '      Grid1.TextMatrix(k, f_o) = Trim(R6("ono")) + " " + Left(str(1 + nNull(R6("NUM2"))) + Space(8), 8) + "-"
                '  End If

                ' Grid1.TextMatrix(k, f_o) = R6!ONO
                  
158             Grid1.TextMatrix(k, F_M) = R6("MON")
160             Grid1.TextMatrix(k, f_f) = R6("FPA")
            End If

162         R6.Close
                  
            'Grid1.TextMatrix(k, F_mikta) = EID.Recordset("mikta")    ' χρειαζονται για το συγκεντροτικο
            ' Grid1.TextMatrix(k, f_a) = Round(EID.Recordset("POSO") * EID.Recordset("TIMM"), 2)
164         Grid1.row = k: Grid1.Col = 0: FindSynolo
166         k = k + 1
168         R5.MoveNext
        Loop
           
170     Grid1.row = k

172     If k > 1 Then
            '  timtext3.Text = Grid1.TextMatrix(k + 1, f_k)
        End If

174     Grid1.Col = f_k  ' για να μην ζητάει να ανοιξει νέα καρτέλλα
176     Grid1_EnterCell

End Sub

Private Sub SEIRA_GotFocus()

        '<EhHeader>
        On Error GoTo SEIRA_GotFocus_Err

        '</EhHeader>

100     SEIRA.BackColor = vbYellow

        '<EhFooter>
        Exit Sub

SEIRA_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.SEIRA_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SEIRA_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub SEIRA_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo SEIRA_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = 13 Then Text1.SetFocus

        '<EhFooter>
        Exit Sub

SEIRA_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.SEIRA_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SEIRA_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub SEIRA_LostFocus()

        ' STELNEI STO FIND_PARASTAT NA BREI TON A/A KAI PAIRNEI TIS PARAMETROYS APO TO PARASTAT
        '<EhHeader>
        On Error GoTo SEIRA_LostFocus_Err

        '</EhHeader>

        Dim k As Integer

        Dim X As String

        ' Dim db As Database
        Dim R As New ADODB.Recordset

        On Error GoTo MHNYMA

100     SEIRA.BackColor = vbWhite

102     For k = 0 To SEIRA.ListCount

104         If SEIRA.Text = SEIRA.List(k) Then

                Exit For

            End If

        Next

106     If k <= SEIRA.ListCount Then
            ' Text1.Text = find_parastat(k, Right(PARAS.Text, 1), 0)
108         X = mEidos_Par(1 + PARAS.ListIndex)
110         Text1.Text = find_parastat(k, X, 0)

        End If

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

        ' db.Open gDir
        ' R.Open "SELECT *FROM parastat", db, adOpenKeyset, adLockReadOnly ', adCmdTableDirect

112     R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

114     R.MoveFirst

116     Do While Not R.EOF

118         If Right(PARAS.Text, 1) = R("EIDOS") Then
120             F_TITLOS = R("TITLOS")
122             F_POS_APOU = R("POS_APOU")
124             F_AJIA_APOU = R("AJIA_APOU")
            End If

126         R.MoveNext
        Loop

128     R.Close

130     If Val(Text1.Text) > 0 Then
132         selectPEL.SetFocus
        End If

        Exit Sub

MHNYMA:
134     HandleError "Par1:FINDSYNOLO"

136     Resume Next

        'PARAS.Clear
        '<EhFooter>
        Exit Sub

SEIRA_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.SEIRA_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SEIRA_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub DBGRID1_KeyDown(keycode As Integer, Shift As Integer)
'    If keycode = 27 Then
'        b_Katax.SetFocus
'    End If
'
'End Sub

'Private Sub DBGrid1_KeyUp(keycode As Integer, Shift As Integer)
'    If keycode = 13 Then
'        If f_ColUpdate Then
'            f_ColUpdate = False
'            GridEidon.SetFocus
'        End If
'    End If
'End Sub

'Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    If LastCol = 5 Then
'
'        ' DBGrid1.Row = DBGrid1.Row + 1: DBGrid1.Col = 0
'    End If
'End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Combo4_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
102         KeyAscii = 0    'suppress the beep
104         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Combo4_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Combo4_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Combo4_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub sfpa_Click()
        Dim poso_strogg As String
        Dim mpos        As Single
        On Error Resume Next
        'If PEL.Recordset("eidos") = "r" Then
        '  poso_strogg = InputBox("Ποσό που θα προσθέσει στο ΦΠΑ 0.01/-0.01 ", , 0.01)
        '   mpos = gVal(poso_strogg)
100     If mpos >= 0.01 Then
            'f_poso_strogg = 0.01
            '       lab_strogg.Caption = "+0.01"
        Else
            'f_poso_strogg = "-0.01"
            '        lab_strogg.Caption = "-0.01"
        End If
        'End If
End Sub

Private Sub SKOPOS_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo SKOPOS_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
102         KeyAscii = 0    'suppress the beep
104         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

SKOPOS_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.SKOPOS_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SKOPOS_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub FORTOSH_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo FORTOSH_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
102         KeyAscii = 0    'suppress the beep
104         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

FORTOSH_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.FORTOSH_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.FORTOSH_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub PROORISMOS_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo PROORISMOS_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
102         KeyAscii = 0    'suppress the beep
104         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

PROORISMOS_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.PROORISMOS_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.PROORISMOS_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GridEidon_Click()

    '<EhHeader>
    On Error GoTo GridEidon_Click_Err

    '</EhHeader>
    Dim DELT As Recordset

    'Grid1.SetFocus

    '<EhFooter>
    Exit Sub

GridEidon_Click_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.par1.GridEidon_Click " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.GridEidon_Click " & "at line " & Erl

    Resume Next

    '</EhFooter>

End Sub

Private Sub GridEidon_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>

        '</EhHeader>
100     If KeyCode = 13 Then

102         take_eidos
           
        End If

End Sub

Private Sub take_eidos()

        On Error GoTo GridEidon_KeyUp_Err
        Dim m_a       As String

        Dim k         As Integer

        Dim MKOD(30)  As String

        Dim mPoso(30) As Single

        Dim f5        As String

        Dim ELEM(30)  As String

        Dim DUM       As Integer

        Dim sPOSO     As Single

        Dim arSET     As Single

        Dim mRow

        Dim R    As New ADODB.Recordset

        Dim SQL2 As New ADODB.Recordset
            
100     If True Then     ' KeyCode = 13 Then

            On Error GoTo MHNYMA

102         For k = 1 To 30: MKOD(k) = "": mPoso(k) = 0: Next

            'ΕΑΝ Ο ΚΩΔΙΚΟΣ ΕΧΕΙ ΜΑΖΙ ΤΟΥ ΟΜΑΔΑ ΕΙΔΩΝ ΤΟΤΕ
            'ΦΟΡΤΩΝΕ ΣΕ ΠΙΝΑΚΑ ΤΑ ΕΙΔΗ

104         If EID.Recordset.EOF Then Exit Sub

            '------------------------------------------------------------------------------------------
106         'If IsNull(EID.Recordset("MMEMO")) Then f5 = "" Else f5 = EID.Recordset("MMEMO")
            If IsNull(EID.Recordset("CH6")) Then f5 = "" Else f5 = EID.Recordset("CH6")
108         If Left(f5, 1) = "#" Then



               'Dim r As New ADODB.Recordset
110             R.Open "SELECT CH6 FROM EID WHERE KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
112             f5 = R(0)
114             R.Close
116             DUM = FETES_DELIM(f5, ELEM)

                ' B??S?O ?????S ?O?????S MKOD() ??? ?? ??S?????S ??O MPOSO()
118             sPOSO = 0

120             For k = 1 To 30

122                 If Left(ELEM(k), 1) = "#" Then
124                     DUM = InStr(ELEM(k), "  ")

126                     If DUM - 2 > 0 Then    'AN EXEI KOMMA
128                         MKOD(k) = mID$(ELEM(k), 2, DUM - 2)
130                         mPoso(k) = Val(mID$(ELEM(k), DUM + 1, Len(ELEM(k)) - DUM))
132                         sPOSO = sPOSO + mPoso(k)
                        End If
                    End If

                Next

134             arSET = Val(InputBox("δωστε συνολ.τεμαχια" + Chr(13) + "To set αυτό έχει " + Format(sPOSO, "##") + " τεμαχια " + Chr(13) + "αν θελετε 2 set τοτε " + Chr(13) + " πληκτρολογηστε " + Format(2 * sPOSO, "##"), "Τεμάχια ανα set : " + Format(sPOSO, "##")))

136           ' AN SOYMA TA SYSTATIKA BGAZOYN ENA (SYSTATIKA KOKTEIL P.X)
              ' If sPOSO > 0 Then arSET = arSET / sPOSO

                'sxediazo to grid
138             DUM = Grid1.row

140             For k = 1 To 30

142                 If mPoso(k) <> 0 Then
144                     EID.RecordSource = "select KOD,ONO,XTI,LTI,LTI5,FPA,MON,MEMO,NUM2   from EID WHERE KOD ='" + MKOD(k) + "';"
146                     EID.Refresh

148                     If Not EID.Recordset.EOF Then
150                         Grid1.TextMatrix(DUM + k - 1, f_k) = MKOD(k)
152                         Grid1.TextMatrix(DUM + k - 1, f_p) = mPoso(k) * arSET

154                         If IsNull(EID.Recordset("mon")) Then m_a = Space(3) Else m_a = EID.Recordset("mon")
156                         Grid1.TextMatrix(DUM + k - 1, f_o) = Left(EID.Recordset("ono") + Space(60), 60)
158                         Grid1.TextMatrix(DUM + k - 1, F_M) = m_a

160                         If F_XONDR = 3 Then    'elga
162                             Grid1.TextMatrix(DUM + k - 1, f_f) = 5
                            Else
164                             Grid1.TextMatrix(DUM + k - 1, f_f) = EID.Recordset("FPA")
                            End If

166                         Grid1.TextMatrix(DUM + k - 1, f_e) = PEL.Recordset("ekpt")

168                         If fbuff = "r" Then
170                             Grid1.TextMatrix(DUM + k - 1, F_T) = EID.Recordset("xti")
                            Else

172                             If lianikh Then
174                                 Grid1.TextMatrix(DUM + k - 1, F_T) = EID.Recordset("LTI5") / ((100 + g_Fpa(EID.Recordset("FPA"))) / 100)
176                                 Grid1.TextMatrix(DUM + k - 1, F_mefpa) = EID.Recordset("LTI5")
                                Else
178                                 Grid1.TextMatrix(DUM + k - 1, F_T) = Round(EID.Recordset("LTI"), f_psifiaTIMMON)
                                End If
                            End If

180                         Grid1.row = DUM + k - 1
182                         FindSynolo
                        End If
                    End If

                Next




'                'Dim r As New ADODB.Recordset
'110             R.Open "SELECT MEMO FROM EID WHERE KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
'112             f5 = R(0)
'114             R.Close
'116             DUM = FETES_DELIM(f5, ELEM)
'
'                ' BΡΙΣΚΩ ΠΟΙΟΥΣ ΚΩΔΙΚΟΥΣ MKOD() ΚΑΙ ΤΙ ΠΟΣΟΤΗΤΕΣ ΕΧΩ MPOSO()
'118             sPOSO = 0
'
'120             For k = 1 To 30
'
'122                 If Left(ELEM(k), 1) = "#" Then
'124                     DUM = InStr(ELEM(k), "  ")
'
'126                     If DUM - 2 > 0 Then    'AN EXEI KOMMA
'128                         MKOD(k) = mID$(ELEM(k), 2, DUM - 2)
'130                         mPoso(k) = Val(mID$(ELEM(k), DUM + 1, Len(ELEM(k)) - DUM))
'132                         sPOSO = sPOSO + mPoso(k)
'                        End If
'                    End If
'
'                Next
'
'134             arSET = Val(InputBox("Δώστε συνολ.τεμάχια" + Chr(13) + "To σετ αυτό έχει " + Format(sPOSO, "##") + " τεμάχια" + Chr(13) + "αν θέλετε 2 σετ τότε " + Chr(13) + " πληκτρολογήστε " + Format(2 * sPOSO, "##"), "Τεμάχια ανά σετ : " + Format(sPOSO, "##")))
'
'136             If sPOSO > 0 Then arSET = arSET / sPOSO
'
'                'sxediazo to grid
'138             DUM = Grid1.row
'
'140             For k = 1 To 30
'
'142                 If mPoso(k) > 0 Then
'144                     EID.RecordSource = "select KOD,ONO,XTI,LTI,LTI5,FPA,MON,MEMO,NUM2   from EID WHERE KOD ='" + MKOD(k) + "';"
'146                     EID.Refresh
'
'148                     If Not EID.Recordset.EOF Then
'150                         Grid1.TextMatrix(DUM + k - 1, f_k) = MKOD(k)
'152                         Grid1.TextMatrix(DUM + k - 1, f_p) = mPoso(k) * arSET
'
'154                         If IsNull(EID.Recordset("mon")) Then m_a = Space(3) Else m_a = EID.Recordset("mon")
'156                         Grid1.TextMatrix(DUM + k - 1, f_o) = Left(EID.Recordset("ono") + Space(60), 60)
'158                         Grid1.TextMatrix(DUM + k - 1, F_M) = m_a
'
'160                         If F_XONDR = 3 Then    'elga
'162                             Grid1.TextMatrix(DUM + k - 1, f_f) = 5
'                            Else
'164                             Grid1.TextMatrix(DUM + k - 1, f_f) = EID.Recordset("FPA")
'                            End If
'
'166                         Grid1.TextMatrix(DUM + k - 1, f_e) = PEL.Recordset("ekpt")
'
'168                         If fbuff = "r" Then
'170                             Grid1.TextMatrix(DUM + k - 1, F_T) = EID.Recordset("xti")
'                            Else
'
'172                             If lianikh Then
'174                                 Grid1.TextMatrix(DUM + k - 1, F_T) = EID.Recordset("LTI5") / ((100 + g_Fpa(EID.Recordset("FPA"))) / 100)
'176                                 Grid1.TextMatrix(DUM + k - 1, F_mefpa) = EID.Recordset("LTI5")
'                                Else
'178                                 Grid1.TextMatrix(DUM + k - 1, F_T) = Round(EID.Recordset("LTI"), f_psifiaTIMMON)
'                                End If
'                            End If
'
'180                         Grid1.row = DUM + k - 1
'182                         FindSynolo
'                        End If
'                    End If
'
'                Next

                '------------------------------------------------------------------------------------------
            Else

184             Grid1.Col = f_k

                ' ΓΡΑΦΩ ΤΟΝ ΚΩΔΙΚΟ
186             Grid1.Text = Left(EID.Recordset("kod"), 20)
188             Grid1.Col = f_o

190             If IsNull(EID.Recordset("mon")) Then m_a = Space(3) Else m_a = Left(EID.Recordset("mon"), 3)
                'ΓΡΑΦΩ ΤΗΝ ΠΕΡΙΓΡΑΦΗ ΤΟΥ ΕΙΔΟΥΣ
                If F_PAROCHOS = 1 Then
                    Dim mUBL As String
                    mUBL = GGET_CVALUE("select ISNULL(C2,'') FROM PINAKES WHERE TYPOS=2 AND PERIGRAFH='" + m_a + "'")
                    If Trim(mUBL) = "" Then
                       MsgBox "Η ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ ΔΕΝ ΕΧΕΙ ΑΝΤΙΣΤΟΙΧΟ UBL"
                    End If
                End If
192             Grid1.Col = F_M
194             Grid1.Text = m_a
196             Grid1.Col = f_o

198             RAFI.Caption = CNull(EID.Recordset("UES"))

200             If Left(EID.Recordset("kod"), 1) = "*" Then
                Else

202                 If F_EISITHRIA = 0 Then
204                     Grid1.Text = Trim(Left(EID.Recordset("ono") + Space(60), 150))  '+ m_a
                    Else
206                     Grid1.Text = Trim(EID.Recordset("ono")) + " " + Left(str(1 + nNull(EID.Recordset("NUM2"))) + Space(8), 8) + "-"
                    End If
                
                End If
                
                Dim mSym As String
208             If gAsterisk = 1 Then
210                 mSym = "@"
                Else
212                 mSym = "*"
                End If

214             If Left(EID.Recordset("kod"), 1) = mSym Then

216                 Grid1.Text = InputBox("Δώσε περιγραφή", "", Grid1.Text)
                End If

                '   Grid1.Text = InputBox("Δώσε περιγραφή", "", Grid1.Text)
                'End If

218             Grid1.Col = f_p  '11-3-07
220             Grid1.Text = 1

222             If F_KOLA_MIKTA = 1 Then ' ΚΟΛΑ ΜΙΚΤΑ
224                 Grid1.Col = F_PROEL
                End If

                ' Grid1.Text = 1    '11-3-07

226             If F_KOLA_MIKTA = 1 Then
                    If F_PIN_PROELEYS = 2 Then ' 2=SAKIS PROELEYSH=EID.ERG   '   ΚΟΛΑ ΜΙΚΤΑ
228                        Grid1.Text = EID.Recordset("ERG")
                    End If
                    If F_PIN_PROELEYS = 3 Then ' 2=SAKIS PROELEYSH=EID.ERG   '   ΚΟΛΑ ΜΙΚΤΑ
                        Grid1.Text = GGET_CVALUE("select isnull(CH4,'') FROM EID WHERE KOD='" + EID.Recordset("KOD") + "'")
                          Grid1.Col = Grid1.Col + 1
                           Grid1_EnterCell
                          'keybd_event VK_TAB, 0, 0, 0    'send a tab
                    End If

                    '  SendKeys "{enter}{enter}"
                    'keybd_event VK_TAB, 0, 0, 0 'send a tab
     
                End If

230             If Len(Trim(F_2CHFIA_ZYGIZOMENA)) > 1 And Left(F_2CHFIA_ZYGIZOMENA, 2) = Left(timtext3.Text, 2) Then
            
232                 If mID(F_2CHFIA_ZYGIZOMENA, 3, 1) = "b" Then
234                     Grid1.Text = Val(mID(timtext3.Text, 8, 5)) / 1000
                    Else
236                     Grid1.Text = Round(Val(mID(timtext3.Text, 8, 5)) / EID.Recordset("lti5") / 100, 3)
                    End If
               
                End If

238             Grid1.Col = f_e

                Dim timokPEL As String, timokEID As String
                Dim rSQL2    As New ADODB.Recordset
240             If f_TIMOKAT = 9999 Then

242                 timokPEL = Trim(IIf(IsNull(PEL.Recordset("CH4")), "", PEL.Recordset("CH4")))
244                 timokEID = Trim(IIf(IsNull(EID.Recordset("CH1")), "", EID.Recordset("CH1")))

246                 rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOKPEL='" + timokPEL + "' AND TIMOKEID='" + timokEID + "'", Gdb, adOpenForwardOnly, adLockReadOnly

248                 If rSQL2.EOF Then
250                     Grid1.TextMatrix(Grid1.row, f_e) = 0
                    Else
252                     Grid1.TextMatrix(Grid1.row, f_e) = rSQL2(0)
                    End If

254                 rSQL2.Close

256             ElseIf f_TIMOKAT = 9998 Then  ' οτι και το 9999 αλλα με με τιμη και οχι εκπτωση

                    'Dim timokPEL As String, timokEID As String

258                 timokPEL = Trim(IIf(IsNull(PEL.Recordset("CH4")), "", PEL.Recordset("CH4")))
260                 timokEID = Trim(IIf(IsNull(EID.Recordset("CH1")), "", EID.Recordset("CH1")))

                    'Dim rSQL2 As New ADODB.Recordset

262                 rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOKPEL='" + timokPEL + "' AND TIMOKEID='" + timokEID + "'", Gdb, adOpenForwardOnly, adLockReadOnly

264                 If rSQL2.EOF Then
266                     Grid1.TextMatrix(Grid1.row, F_T) = EID.Recordset("lti")
                    Else
268                     Grid1.TextMatrix(Grid1.row, F_T) = rSQL2(0)
                    End If

270                 rSQL2.Close

272             ElseIf f_TIMOKAT >= 1 And f_TIMOKAT < 300 And Not IsNull(PEL.Recordset("pek")) Then
274                 R.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + str(PEL.Recordset("pek")) + " AND KOD='" + Trim(Left(EID.Recordset("kod"), 13)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

276                 If R.EOF Then
278                     Grid1.Text = ""
                    Else
280                     Grid1.Text = R(0)
                    End If

                    'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
282             ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then

284                 R.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + str(PEL.Recordset("pek")) + " AND KOD='" + Trim(Left(EID.Recordset("kod"), 13)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

286                 If R.EOF Then
                        'Grid1.TextMatrix(k, F_T) = 0
288                     Grid1.TextMatrix(Grid1.row, F_T) = EID.Recordset("lti")
                    Else
290                     Grid1.TextMatrix(Grid1.row, F_T) = R(0)
292                     If F_KYBRAN = 1 Then
                         
294                         If nNull(PEL.Recordset("NUM1")) > 0 Then
296                             Grid1.TextMatrix(Grid1.row, f_e) = nNull(PEL.Recordset("NUM1"))
                            End If
                           
                        End If

                    End If

298                 R.Close
                
300             ElseIf f_TIMOKAT = 333 And Not IsNull(PEL.Recordset("pek")) Then
302                 If PEL.Recordset("pek") = 2 Then
304                     Grid1.TextMatrix(Grid1.row, F_T) = EID.Recordset("lti2")
306                 ElseIf PEL.Recordset("pek") = 3 Then
308                     Grid1.TextMatrix(Grid1.row, F_T) = EID.Recordset("lti3")
310                 ElseIf PEL.Recordset("pek") = 5 Then
312                     Grid1.TextMatrix(Grid1.row, F_mefpa) = gVal(EID.Recordset("LTI5"))
314                     Grid1.TextMatrix(Grid1.row, F_T) = Format(EID.Recordset("LTI5") / (1 + g_Fpa(EID.Recordset("FPA")) / 100), "####0." + String(f_psifiaAjias, "0"))
                    Else
316                     Grid1.TextMatrix(Grid1.row, F_T) = EID.Recordset("lti")
                    End If
                
318             ElseIf f_TIMOKAT = 33333 Then ' agoras
320                 Grid1.TextMatrix(Grid1.row, F_T) = EID.Recordset("xti")
                
                Else

                    ' για δωροκαρτες δεν βαζει εκπτωση
322                 If f_pelMERC = "SOK" Then
                    Else

324                     If IsNull(PEL.Recordset("pek")) Then
326                         Grid1.Text = 0
                        Else
328                         Grid1.Text = PEL.Recordset("pek")
                        End If
                       
                    End If
            
                End If

                'ΓΡΑΦΩ ΤΟ ΦΠΑ ΤΟΥ ΕΙΔΟΥΣ

330             Grid1.Col = f_f

332             If Not IsNull(PEL.Recordset("TYPOS")) Then
334                 If Val(PEL.Recordset("TYPOS")) = 1 Then  'ΝΗΣΙΑ
336                     Grid1.Text = GGET_FPA_NHSION(EID.Recordset("FPA")) '
                        
338                 ElseIf Val(PEL.Recordset("TYPOS")) = 2 Then  'ENDOKOINOTIKO
           
340                     Grid1.Text = 5
                        
342                 ElseIf Val(PEL.Recordset("TYPOS")) = 3 Then  'TX
344                     Grid1.Text = 5
                    Else
346                     If F_XONDR = 3 Then    'elga
348                         Grid1.Text = EID.Recordset("FPA")  'Grid1.Text =5
                        Else
350                         Grid1.Text = EID.Recordset("FPA")
                        End If

                        ' Grid1.Text = EID.Recordset("FPA")
                    End If

                Else
352                 Grid1.Text = EID.Recordset("FPA")
                End If


               If F_PROEPILFPA > 0 Then
                  Grid1.Text = F_PROEPILFPA
               End If
               



                'ΓΡΑΦΩ ΤΗΝ ΤΙΜΗ ΤΟΥ ΕΙΔΟΥΣ
354             If fbuff = "r" Then
356                 Grid1.Col = F_T

358                 If F_TIMH_XONDR_TIM_AGOR = 1 Then
                    
360                     Grid1.Text = Format(EID.Recordset("xti"), "####0." + String(f_psifiaAjias, "0"))
                    
                    Else
362                     Grid1.Text = Format(EID.Recordset("LTI"), "####0." + String(f_psifiaAjias, "0"))
                    
                    End If
                    
364                 If F_STADIO = 3 Or F_STADIO = 5 Then    ' deltio apostolis   Ή ΕΝΔΟΔΙΑΚΙΝΗΣΗ
366                     Grid1.Text = 0
                    End If
                
                    '''''''''''''''''''''''''''''''''''''''''''''   πωλησεις ========================
                Else  'If fbuff = "r" Then
368                 Grid1.Col = F_T

370                 If lianikh Then

372                     Grid1.Col = F_mefpa
374                     Grid1.Text = gVal(EID.Recordset("LTI5"))
376                     Grid1.Col = F_T
378                     Grid1.Text = Format(EID.Recordset("LTI5") / (1 + g_Fpa(EID.Recordset("FPA")) / 100), "####0." + String(f_psifiaTIMMON, "0"))

                    Else

380                     If f_TIMOKAT < 0 Or f_TIMOKAT = 9998 Then ' TIMOKATALOGOS ME KARFOTES TIMES
382                     ElseIf f_TIMOKAT = 333 Then ' 3 XONDTRIKES
384                         If PEL.Recordset("pek") = 2 Then
386                             Grid1.TextMatrix(Grid1.row, F_T) = EID.Recordset("lti2")
388                         ElseIf PEL.Recordset("pek") = 3 Then
390                             Grid1.TextMatrix(Grid1.row, F_T) = EID.Recordset("lti3")
392                         ElseIf PEL.Recordset("pek") = 5 Then
394                             Grid1.Col = F_mefpa
396                             Grid1.Text = gVal(EID.Recordset("LTI5"))
398                             Grid1.Col = F_T
400                             Grid1.Text = Format(EID.Recordset("LTI5") / (1 + g_Fpa(EID.Recordset("FPA")) / 100), "####0." + String(f_psifiaAjias, "0"))
                            Else
402                             Grid1.TextMatrix(Grid1.row, F_T) = EID.Recordset("lti")
                            End If
                            
404                     ElseIf f_TIMOKAT = 33333 Then ' agoras
406                         Grid1.TextMatrix(Grid1.row, F_T) = EID.Recordset("xti")
                        Else
408                         Grid1.Text = Format(EID.Recordset("LTI"), "####0." + String(f_psifiaAjias, "0"))
                        End If

                    End If
                    
                    'εαν χρησιμοποιώ ΕΦΚ ΕΙΔΙΚΟ ΦΟΡΟ ΚΑΤΑΝΑΛΩΣΗΣ
410                 If F_EIDFOROSKAT = 1 And F_AJIA_APOU > 0 Then
                           
                        Dim refk As New ADODB.Recordset
412                     refk.Open "select EFK FROM EID WHERE KOD='" + EID.Recordset("kod") + "'", Gdb, adOpenDynamic, adLockOptimistic
414                     Grid1.TextMatrix(Grid1.row, f_colEFK) = nNull(refk(0))
416                     refk.Close
                           
                    End If

418                 If F_STADIO = 2 Then  ' ΠΑΡΑΓΓΕΛΊΑ

420                     If Val(f_PARAGGMARKUP) = 0 Then
                        Else
422                         Grid1.Text = Format((100 + Val(metaf.Text)) / 100 * EID.Recordset("XTI"), "####0.00")
                        End If

                    End If

424                 If F_STADIO = 3 Or F_STADIO = 5 Then    ' deltio apostolis   Ή ΕΝΔΟΔΙΑΚΙΝΗΣΗ
426                     Grid1.Text = 0
                    End If
                End If   ' If fbuff = "r" Then

428             FindSynolo
            End If  '="#"

            ' ΤΟ ΤΕΧΤ3 ΠΑΡΑΛΑΜΒΑΝΕΙ ΤΗΝ ΤΙΜΗ ΑΠΟ ΤΟ GRID
430         Grid1.Col = f_p

432         If F_KOLA_MIKTA = 1 Then ' ΚΟΛΑ ΜΙΚΤΑ
434             Grid1.Col = F_PROEL
     
436             If F_PIN_PROELEYS = 2 Then ' sakis
438                 Grid1.Col = F_COL_KOLA
                End If
     
            End If

440         timtext3.Text = Grid1.Text
442         timtext3.Visible = False

            ' ΤΟ ΤΕΧΤ3 ΠΑΡΑΛΑΜΒΑΝΕΙ ΤIΣ ΔΙΑΣΤΑΣΕΙΣ ΑΠΟ ΤΟ GRID
444         timtext3.Top = Grid1.Top + Grid1.CellTop
446         timtext3.Left = Grid1.Left + Grid1.CellLeft
448         timtext3.Width = Grid1.CellWidth - 15
450         timtext3.Height = Grid1.CellHeight - 25
452         timtext3.Visible = True
454         timtext3.SetFocus
        End If  ' =CHR(13)

        Dim PI, pi2

        ' PI = "C:\MERCVB\IMAGES\" + CNull(EID.Recordset("kod"))

456     If F_EIKONA > 0 Then

            On Error Resume Next
            'On Error GoTo 0

458         If F_EIKONA = 2 Then    'βοηθητικοσ κωδικοσ
460             PI = "C:\MERCVB\IMAGES\" + CNull(EID.Recordset("ERG"))
462         ElseIf F_EIKONA = 1 Then    ' kvdikos
464             PI = "C:\MERCVB\IMAGES\" + CNull(EID.Recordset("kod"))
466         ElseIf F_EIKONA = 12 Then    'βοηθητικοσ κωδικοσ +.jpg
468             PI = "C:\MERCVB\IMAGES\" + CNull(EID.Recordset("ERG")) + ".JPG"
470         ElseIf F_EIKONA = 11 Then    ' kodikos .jpg
472             PI = "C:\MERCVB\IMAGES\" + CNull(EID.Recordset("kod")) + ".JPG"
474         ElseIf F_EIKONA = 111 Then    ' kodikos .jpg
476             PI = "C:\MERCVB\IMAGES\" + Trim(CNull(EID.Recordset("kod"))) + "*" + ".JPG"

            End If

478         If f_site > 0 Then ' = Val(FindParametroi(1,"MDIFORM1", "F_SITE", "0", "ΚΩΔΙΚΟΣ SITE 1=VM 2=TOYS"))

                Dim R_SITE As New ADODB.Recordset

480             R_SITE.Open "SELECT PIC0 FROM EID WHERE KOD='" + EID.Recordset("kod") + "'", Gdb, adOpenDynamic, adLockOptimistic
482             PI = R_SITE("pic0")
                 
484             If f_site = 222 Then
486                 pi2 = "\\Tsaftsufserver\images\" + PI
                Else
488                 pi2 = "c:\mercvb\images\" + PI
                End If

490             If Not IsNull(pi2) Then
492                 If Len(Dir(pi2)) > 0 Then
494                     Image1.Picture = LoadPicture(pi2)
                    Else
496                     Image1.Picture = LoadPicture()
                    End If

                Else
498                 Image1.Picture = LoadPicture()
                End If
                 
500             Else:   ' End If

502             If Len(Dir(PI)) > 0 Then
                    ' Image1.Picture = LoadPicture("C:\MERCVB\IMAGES\" + Dir(PI))
                    'Image1.Picture = LoadPicture()
504                 Image1.Picture = LoadPicture("C:\MERCVB\IMAGES\" + Dir(PI))
                    '.PaintPicture Picture1.Picture, 0, 0, Picture2.width, Picture2.height, 0, 0, Picture1.width, Picture1.height
                Else
506                 Image1.Picture = LoadPicture()
                    ' Picture2.Picture = LoadPicture()
                End If
             
            End If

        End If

        'If Len(Dir(PI)) > 0 Then
        '   On Error Resume Next
        '   Image1.Picture = LoadPicture(PI)
        'Else
        '  On Error Resume Next
        '   PI = "C:\MERCVB\IMAGES\" + EID.Recordset("kod") + ".JPG"
        '   If Len(Dir(PI)) > 0 Then
        '      Image1.Picture = LoadPicture(PI)
        '   Else
        '      Image1.Picture = LoadPicture()
        '   End If
        'End If

        
        
        Exit Sub

MHNYMA:
508     HandleError "Par1:GridEidon_KeyUp"

510     If Err.Description = "Automation error" Then
            'Gdb.Close
            'Gdb.Open gConnect
        End If

512     Resume Next

        '<EhFooter>
        Exit Sub

GridEidon_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.GridEidon_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.GridEidon_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function FETES_DELIM(LINE, ELEM)

        '<EhHeader>
        On Error GoTo FETES_DELIM_Err

        '</EhHeader>
        Dim KL, KE

        On Error GoTo MHNYMA

        '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
100     For KE = 1 To 30
102         ELEM(KE) = ""
        Next

104     KL = 0   ' metraei xaraktires
106     KE = 0   ' metritis toy pinaka  ELEMENT

        Do
108         KE = KE + 1  ' metritis toy pinaka  ELEMENT
110         KL = KL + 1  ' metraei xaraktires

112         Do While mID$(LINE, KL, 1) <> Chr(13)    ' tab
114             ELEM(KE) = ELEM(KE) + mID$(LINE, KL, 1)
116             KL = KL + 1  ' metraei xaraktires

118             If KL > Len(LINE) Then Exit Do
            Loop

120         KL = KL + 1    ' για να περασει το chr(10)

122         If KL > Len(LINE) Then Exit Do

124     Loop Until KL >= Len(LINE)    'OLO TO MHKOS THS GRAMMHS

126     FETES_DELIM = 0

        Exit Function

MHNYMA:
128     HandleError "Par1:Fetesdelim"

130     Resume Next

        '<EhFooter>
        Exit Function

FETES_DELIM_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.FETES_DELIM " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.FETES_DELIM " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub GridPelaton_DblClick()
100     mmGridPelaton_DblClick 1, 0, 0, 0  ' 'TIMES0XORIS1 =0  DEFAULT      1=ΧΩΡΙΣ ΤΙΜΗ ( ΣΕ ΔΕΛΤΙΟ)
End Sub

Private Sub mmGridPelaton_DblClick(ByVal ask As Integer, _
                                   ByVal m_ID As Long, _
                                   TIMES0XORIS1 As Integer, _
                                   mpar As Integer)
        'TIMES0XORIS1 =0  DEFAULT      1=ΧΩΡΙΣ ΤΙΜΗ ( ΣΕ ΔΕΛΤΙΟ)
        ' MPAR As Integer   METASXHMATIZOMENO PARASTATIKO  DEFAULT=0 SE DIORUVSH-AKYROTIKO=99
        '<EhHeader>
        On Error GoTo GridPelaton_DblClick_Err

        '</EhHeader>
        Dim a, e

        Dim k

        ' ΔΕΝ ΔΙΑΛΕΞΕ ΠΕΛΑΤΗ ΚΑΙ ΕΠΙΣΤΡΕΦΕΙ ΠΙΣΩ
100     If PEL.Recordset.EOF Then
102         Text2(1).SelStart = 0
104         Text2(1).SelLength = Len(Text2(1))
106         Text2(0).Text = ""
108         Text2(1).SetFocus

            Exit Sub

        End If

        ' ΔΙΑΛΕΞΕ ΠΕΛΑΤΗ ΑΛΛΑ ΔΕΝ ΤΟΝ ΒΡΙΣΚΕΙ ΣΤΟ ΑΡΧΕΙΟ ΚΑΙ ΕΠΙΣΤΡΕΦΕΙ ΠΙΣΩ
110     f_gkod_pel = fbuff + PEL.Recordset("KOD")
112     GridPelaton.Visible = False
114     PEL.RecordSource = "select *,isnull(STREETNUMBER,'0') AS STREETNUMBER2,ISNULL(BRANCH,'0') AS BRANCH2 from PEL where EIDOS='" + Left(f_gkod_pel, 1) + "' and KOD='" + mID$(f_gkod_pel, 2, Len(f_gkod_pel) - 1) + "'"
116     PEL.Refresh

118     If PEL.Recordset.RecordCount = 0 Then
120         Text2(1).SetFocus

            Exit Sub

        End If

122     If Len(gdirlog) > 0 Then
124         If IsNull(PEL.Recordset("shm2")) Then
126             f_error_gefyr = 1
            Else

128             If Len(PEL.Recordset("shm2")) < 4 Then
130                 f_error_gefyr = 1
                End If
            End If
        End If

        '---------------------------------------------  ΒΑΖΕΙ ΤΑ ΣΤΟΙΧΕΙΑ ΤΟΥ ΣΤΗΝ ΜΑΣΚΑ ΤΟΥ ΤΙΜΟΛΟΓΙΟΥ
132     Text2(0).Text = PEL.Recordset("KOD")
134     Text2(1).Text = PEL.Recordset("EPO")

       
           





        'ΠΕΡΝΑΩ ΠΕΛΑΤΗ ΚΑΙ ΓΕΜΙΖΕΙ ΤΑ ΚΟΥΤΑΚΙΑ ΤΗΣ ΔΙΕΥΘΥΝΣΗΣ ΑΠΟΣΤΟΛΗΣ
        
        
        STREET.Text = PEL.Recordset("DIE")
        STREETNUMBER.Text = CNull(PEL.Recordset("STREETNUMBER2"))
        POSTALCODE.Text = CNull(PEL.Recordset("XRVMA"))
         
         
         If f_isdiakin > 0 Then
             If Len(Trim$(POSTALCODE)) < 4 Then
               'κανω προσπαθεια για να τον διαβασω απο την ΑΑΔΕ
                Dim ΜΜu As String, ΜΜp As String
                ΜΜu = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_USER", "LAGAKIS23041960", "ΧΡΗΣΤΗΣ ΕΥΡΕΣΗΣ ΑΦΜ")
                ΜΜp = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_PWD", "d389833921", "ΚΩΔΙΚΟΣ ΕΥΡΕΣΗΣ ΑΦΜ")
                Dim newKod As String, ΜΜEPO As String, ΜΜDIE As String, ΜΜEPA As String, ΜΜpol As String, ΜΜtk As String, ΜΜdoy As String, ΜΜc As String, ΜΜENERGOS As Integer
                ΜΜc = PEL.Recordset("afm")  ' Text2(2).Text
                Par7MyData.requestafm2 ΜΜu, ΜΜp, ΜΜc, ΜΜEPO, ΜΜDIE, ΜΜEPA, ΜΜpol, ΜΜtk, ΜΜdoy, ΜΜENERGOS
                  If Len(ΜΜtk) > 4 Then
                      POSTALCODE.Text = ΜΜtk
                      Gdb.Execute "UPDATE PEL SET XRVMA='" + ΜΜtk + "' WHERE EIDOS='" + fbuff + "' AND KOD='" + PEL.Recordset("KOD") + "'"
                      
                  Else
                        MsgBox "δεν εχω ΤΚ διορθωστε τον πελάτη"
                       ' Exit Sub
                      
                  End If
             End If
         End If
         
         
             
             
             
        
        
        CITY.Text = PEL.Recordset("POL")
        BRANCH.Text = PEL.Recordset("BRANCH2")
        If Val(STREETNUMBER.Text) = 0 Then
            STREETNUMBER.Text = "0"
        End If
             
        
        If Val(BRANCH.Text) = 0 Then
            BRANCH.Text = "0"
        End If

136     If IsNull(PEL.Recordset("afm")) Then
138         Text2(2).Text = ""
        Else
140         Text2(2).Text = PEL.Recordset("afm")
        End If

142     Text2(3).Text = Trim(PEL.Recordset("die")) + " TK:" + CNull(PEL.Recordset("XRVMA"))

        '        If f_PEL_AYTOKINHTO = 1 Then
        '          AYTOKINHTO.Text = Trim(PEL.Recordset("ch1"))
        '      End If

        ' ΒΑΖΕΙ ΣΤΟΝ ΤΡΟΠΟ ΠΛΗΡΩΜΗΣ ΟΤΙ ΕΧΟΥΜΕ ΚΑΝΟΝΙΣΕΙ ΓΙΑ ΤΟΝ ΠΕΛΑΤΗ
144     If Not IsNull(PEL.Recordset("ARPARAG")) Then
            Dim nmetritis As Integer
146         nmetritis = Max(0, gVal(PEL.Recordset("ARPARAG")) - 1)
            
148         Combo4.Text = Combo4.List(nmetritis)
            Dim NFF As Integer
150         NFF = Val(Left(PEL.Recordset("ARPARAG"), 2))
            '            If NFF = 0 Then
            '               Combo4.Text = Combo4.List(0)
            '            Else
            '               Combo4.Text = Combo4.List(NFF - 1)
            '            End If

        Else
        
152         Combo4.Text = Combo4.List(F_DEFA_PLHR)  ' PEL.Recordset("ARPARAG")

        End If

154     If F_THERM = 1 Then
       
156         Textm(7).Text = "ΔΕΗ: " + PEL.Recordset("CH1")
158         ARDEH.Caption = "ΔΕΗ: " + PEL.Recordset("CH1")
       
        End If

160     AYTOKINHTO.Text = CNull(PEL.Recordset("CH3"))

        ' ΒΑΖΕΙ ΣΤΟΝ ΠΩΛΗΤΗ ΟΤΙ ΕΧΟΥΜΕ ΚΑΝΟΝΙΣΕΙ ΓΙΑ ΤΟΝ ΠΕΛΑΤΗ
162     If F_POLHTES = 1 Then
164         If Not IsNull(PEL.Recordset("PVLHTHS")) Then

                On Error Resume Next

166             Combo8.ListIndex = Val(PEL.Recordset("PVLHTHS")) - 1
            End If
        End If

        ' If Check1 Then f_metasx = "a"

168     If Len(Trim(f_metasx)) = 0 And ApoParaggelia = Checked Then
170         f_metasx = "a"
        End If

172     FlxGridDeltion.Clear
174     FlxGridDeltion.Visible = False
176     GridEidon.Visible = True

        ' αν ΜΕΤΑΣΧΗΜΑΤΙΖΕΤΑΙ  ΑΠΟ ΚΑΤΙ ΑΛΛΟ ΔΕΙΧΝΕΙ ΤΑ ΣΧΕΤΙΖΟΜΕΝΑ ΠΑΡΑΣΤΑΤΙΚΑ Π.Χ. είναι τιμολόγιο του δείχνει τα δελτία σε εκκρεμότητα
178     If Len(Trim(f_metasx)) > 0 Then
180         EID.RecordSource = "select METAF,ATIM,CONVERT(CHAR(10),HME,103) AS [Ημερομηνία],ID_NUM,isnull(ENTITYMARK,'') AS MARK from TIM where left(ATIM,1) in ('" + Get_meKomma(f_metasx) + "') and KPE='" + PEL.Recordset("kod") + "' and ART IS null AND AKYROMENO=0 ORDER BY HME;"
182         EID.Refresh
184         metasxhmatismos.Visible = True
186         FlagGridEidon = False

188         If EID.Recordset.RecordCount > 0 Then
190             Set FlxGridDeltion.DataSource = EID
192             GridEidon.Visible = False
194             FlxGridDeltion.Visible = True
196             FlxGridDeltion.SetFocus
            End If
        End If

        'Grid1.Col = 0
198     Grid1.row = 1
        'Grid1.SetFocus

200     Text2(0).Enabled = False
202     Text2(1).Enabled = False
204     Text2(2).Enabled = False

        'ΤΟΠΟΘΕΤΕΙ ΣΤΟ ΠΡΩΤΟ ΚΕΛΙ ΤΟ ΚΟΥΤΑΚΙ ΠΟΥ ΧΡΗΣΙΜΟΠΟΙΩ ΓΙΑ ΓΡΑΨΙΜΟ ΣΤΟ ΤΙΜΟΛΟΓΙΟ
206     Grid1.Col = 1
208     timtext3.Top = Grid1.Top + Grid1.CellTop
210     timtext3.Left = Grid1.Left + Grid1.CellLeft
212     timtext3.Width = Grid1.CellWidth - 15
214     timtext3.Height = Grid1.CellHeight - 25

        '- --------------- ΔΙΟΡΘΩΣΗ ΤΙΜΟΛΟΓΙΟΥ ----------------------------------
        Dim re   As New ADODB.Recordset, ANS As Integer

        Dim rEgg As New ADODB.Recordset

216     f_dat = DTPicker1.Value
218     f_matim = Right(PARAS.Text, 1) + Right("0000000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
220     ANS = vbNo
    
222     b_Katax.Enabled = True
224     re.Open "select * from TIM where  HME>='" + Format(f_dat, "mm/dd/yyyy") + "' AND HME<='" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "'  and ATIM='" + Trim(f_matim) + "'", Gdb, adOpenDynamic, adLockOptimistic
           
226     If Not re.EOF Then
228         F_ID_NUM_DIOR = re("ID_NUM")
230         If ask = 1 Then
232             ANS = MsgBox("Υπάρχει ήδη  παραστατικό με ίδια ημερομηνία και ίδιο αριθμό. " + Chr(13) + "Να διορθωθεί;", vbYesNo)
234         ElseIf ask = 2 Then
236             ANS = vbYes
            Else
238             ANS = vbYes
            End If

        Else
        
240         If ask = 2 Then
242             ANS = vbYes
            End If
        
            'ΣΒΗΝΩ ΤΥΧΟΝ ΥΠΟΛΕΙΜΑΤΑ
244         Gdb.Execute "DELETE  from EGGTIM where  HME>='" + Format(f_dat, "mm/dd/yyyy") + "' AND HME<='" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "'  and ATIM='" + f_matim + "'"

246         If PEL.Recordset("EIDOS") = "r" Then
248             re.Close
                '  ΥΠΑΡΧΕΙ ΣΤΟΝ ΙΔΙΟ ΠΡΟΜΗΘΕΥΤΗ ΤΟ ΙΔΙΟ ΠΑΡΑΣΤΑΤΙΚΟ ΣΤΙΣ " + Format(re("HME"), "DD/MM/YYYY")
250             re.Open "select * from TIM where KLEIDI='" + Left(f_matim, f_psifia_atim + 2) + PEL.Recordset("kod") + "'", Gdb, adOpenDynamic, adLockOptimistic

252             If Not re.EOF Then
254                 MsgBox "ΠΡΟΣΟΧΗ ΥΠΑΡΧΕΙ ΣΤΟΝ ΙΔΙΟ ΠΡΟΜΗΘΕΥΤΗ ΤΟ ΙΔΙΟ ΠΑΡΑΣΤΑΤΙΚΟ ΣΤΙΣ " + Format(re("HME"), "DD/MM/YYYY") + Chr(13) + " ΔΕΝ ΘΑ ΕΠΙΤΡΑΠΕΙ Η ΚΑΤΑΧΩΡΗΣΗ"
256                 re.Close
258                 b_Katax.Enabled = False
                End If
            End If
        End If

        ' set ειδων (ΚΤΕΛ) ------------------------------------------------------------------------------------
260     If ANS = vbNo Then
   
262         If IsNull(PEL.Recordset("NUM2")) Or F_SET_EIDON = 0 Then
   
            Else  '========================================================================= ktel ========================================================
     
264             If PEL.Recordset("NUM2") > 0 Then

                    Dim M_SET As Integer

                    Dim R5    As New ADODB.Recordset

                    Dim R6    As New ADODB.Recordset

                    Dim M6KOD As String

266                 M_SET = F_SET_EIDON
268                 If M_SET = 1 Then   ' μονο στα εισιτηρια με προμηθεια
270                     M_SET = PEL.Recordset("NUM2") ' για να παιρνει το σετ απο τον πελατη
                    End If
                    
                    ' STA AMEX DIABAZO TO SET EIDON APO  TON PELATH
272                 If PEL.Recordset("NUM2") > 9 Then
274                     M_SET = PEL.Recordset("NUM2")
                    End If
        
276                 k = 1
278                 R5.Open "SELECT * FROM EIDH_DEFAULT WHERE AA=" + str(M_SET) + " ORDER BY AA_TIMOL ", Gdb, adOpenDynamic, adLockOptimistic

280                 Do While Not R5.EOF
              
282                     Grid1.TextMatrix(k, f_k) = R5("kod")
284                     Grid1.TextMatrix(k, f_p) = 0
286                     Grid1.TextMatrix(k, f_e) = R5("ekpt")
              
288                     Grid1.TextMatrix(k, F_T) = R5!TIMM
              
290                     If IsNull(R5!kod) Then
292                         M6KOD = ""
                        Else
294                         M6KOD = R5!kod
                        End If
              
296                     R6.Open "SELECT * FROM EID WHERE KOD='" + M6KOD + "'", Gdb, adOpenDynamic, adLockOptimistic

298                     If Not R6.EOF Then
                  
300                         If F_EISITHRIA = 0 Then
302                             Grid1.TextMatrix(k, f_o) = R6("ono") ' Left(R6("ono") + Space(60), 60)  '+ m_a
                            Else
304                             Grid1.TextMatrix(k, f_o) = Trim(R6("ono")) + " " + Left(str(1 + nNull(R6("NUM2"))) + Space(8), 8) + "-"
                            End If

                            ' Grid1.TextMatrix(k, f_o) = R6!ONO
                  
306                         Grid1.TextMatrix(k, F_M) = R6("MONA")
308                         Grid1.TextMatrix(k, f_f) = R6("FPA")
                        End If

310                     R6.Close
                  
                        'Grid1.TextMatrix(k, F_mikta) = EID.Recordset("mikta")    ' χρειαζονται για το συγκεντροτικο
                        ' Grid1.TextMatrix(k, f_a) = Round(EID.Recordset("POSO") * EID.Recordset("TIMM"), 2)
312                     Grid1.row = k: Grid1.Col = 0: FindSynolo
314                     k = k + 1
316                     R5.MoveNext
                    Loop
           
318                 Grid1.row = k

320                 If k > 1 Then
322                     timtext3.Text = Grid1.TextMatrix(k + 1, f_k)
                    End If

324                 Grid1.Col = f_k  ' για να μην ζητάει να ανοιξει νέα καρτέλλα
326                 Grid1_EnterCell
                End If 'IsNull(PEL.Recordset("NUM2"))   '========================================================================= ktel ========================================================
            End If  'ANS = vbNo
        End If  ' KTEL

        ' set ειδων (ΚΤΕΛ) ------------------------------------------------------------------------------------

        '--ΑΡΧΗ -------------------------------------- ναι να διορθωθει .//  ΒΛΕΠΩ ΤΑ ΕΙΔΗ ---------------------------------------

328     If ANS = vbYes Then

            'βαζει τον αρχικό πελάτη
            Dim ANS3

330         ANS3 = vbYes
       
            'ΜΕ ΡΩΤΑΕΙ ΑΝ ΘΕΛΩ ΝΑ ΑΛΛΑΞΩ ΠΕΛΑΤΗ ΠΡΟΜΗΘΕΥΤΗ
332         If f_gkod_pel = PEL.Recordset("EIDOS") + re("KPE") Then
            Else
334             ANS3 = MsgBox("ΚΡΑΤΩ ΤΟΝ ΑΡΧΙΚΟ ΠΕΛΑΤΗ/ΠΡΟΜΗΘΕΥΤΗ; ", vbYesNo)
            End If
       
336         If ANS3 = vbYes Then
338             PEL.RecordSource = "select *from PEL where EIDOS='" + Left(f_gkod_pel, 1) + "' and KOD='" + re("kpe") + "'"
340             PEL.Refresh
342             Text2(0).Text = PEL.Recordset("KOD")
344             Text2(1).Text = PEL.Recordset("EPO")

346             If IsNull(PEL.Recordset("afm")) Then
348                 Text2(2).Text = ""
                Else
350                 Text2(2).Text = PEL.Recordset("afm")
                End If

352             Text2(3).Text = PEL.Recordset("die")
            End If

            'Combo8.Text =comb     re("eidpar") ' Combo8.List(0)
  
354         If Val(re("eidpar")) - 1 >= 0 Then
356             Combo8.Text = Combo8.List(Val(re("eidpar")) - 1)
            End If

358         parat.Text = re("parat")

360         If f_is_benzinadiko = 1 Then
362             ar_apod = Left(parat.Text, 10)
364             parat.Text = mID(parat.Text, 11, 30)
            End If


HMERPARAD.Value = re("HMEPARAD")

ORAPARAD.Value = re("ORAPARAD")





366         AYTOKINHTO.Text = re("aytok")
            
368         PROORISMOS.Text = re("PROOR")
370         SKOPOS.Text = re("SKOPOS")
           other.Text = re("OTHERMOVEPURPOSETITLE")
             STREET.Text = re("street")
             STREETNUMBER = re("streetnumber")
             POSTALCODE.Text = re("postalcode")
             CITY.Text = re("city")
             
372         fortosh.Text = re("FORTOSH")
            If InStr(fortosh.Text, ";") > 0 Then
               FORTDIE.Text = Split(fortosh.Text, ";")(0)
               fortARit.Text = Split(fortosh.Text, ";")(1)
               FORTPOL.Text = Split(fortosh.Text, ";")(2)
               FORTTK.Text = Split(fortosh.Text, ";")(3)
            End If
            
            
    
            
            
            

            Dim RR2() As String
           
374         If re("PARAKRATISI") > 0 Then
376             RR2 = Split(re("B_C2"), ";")
378             parakratisiT.Text = re("PARAKRATISI")
380             If Val(RR2(0)) > 0 Then
382                 Combo2PARAKRATISI.Text = Combo2PARAKRATISI.List(Val(RR2(0) - 1))
                Else
384                 Combo2PARAKRATISI.Text = ""
                End If
386             If Val(RR2(1)) > 0 Then
388                 ComboPARAKRATISI.Text = ComboPARAKRATISI.List(Val(RR2(1) - 1))
                Else
390                 Combo2PARAKRATISI.Text = ""
                End If

            End If

            ' For k = 0 To Combo4.ListCount - 1
            '    If InStr(Combo4.List(k), re("TRP")) > 0 Then
            '       Combo4.Text = Combo4.List(k)
            '       Exit For
            '    End If
            ' Next
392         If re("TRP") = "ΜΕ" Then    ' METRHTOIS
394             Combo4.Text = Combo4.List(0)
            Else
396             Combo4.Text = Combo4.List(1)
            End If

398         f_dior_mAtim = f_matim    ' ΤΑ ΚΡΑΤΑΩ ΜΗΝ ΤΥΧΟΝ ΤΑ ΑΛΛΑΞΕΙ ΣΤΗΝ ΦΆΣΗ ΤΗΣ ΔΙΟΡΘΩΣΗΣ
400         f_dior_HME = f_dat    'ΓΙΑ ΝΑ ΜΠΟΡΕΣΩ ΝΑ ΣΒΗΣΩ ΤΟ ΑΡΧΙΚΟ ΤΙΜΟΛΟΓΙΟ
402         f_dior_ARX_TYPOS = PARAS.Text
            
            ''το είχα έτσι αλλά έχανε τα είδη  αν  εβαζα λάθος τις αποθήκες
            ' EID.RecordSource = "select EGGTIM.* from EGGTIM  WHERE HME >= '" + Format(f_dat, "mm/dd/yyyy") + "' AND HME <= '" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "' and ATIM='" + f_matim + "' AND APOT=" + Str(Val(APOTH1)) + " order by PROOD_AJ"
            Dim m_tim As New ADODB.Recordset
            
404         If m_ID > 0 Then ' ΨΑΧΝΩ ΜΕ ID
                Dim RTEST As New ADODB.Recordset
406             RTEST.Open "SELECT COUNT(*) FROM EGGTIM WHERE ID_NUM=" + str(m_ID), Gdb, adOpenDynamic, adLockOptimistic
408             If RTEST(0) > 0 Then
410                 EID.RecordSource = "select E.EKPT2 AS EKPT22, E.*,T.*,E.EKPT AS EKPTOSH from EGGTIM E INNER JOIN TIM T ON E.ID_NUM=T.ID_NUM  WHERE E.ID_NUM=" + str(m_ID) + "  order by PROOD_AJ"
                Else
412                 EID.RecordSource = "select  EGGTIM.EKPT2 AS EKPT22,EGGTIM.*,EGGTIM.EKPT AS EKPTOSH from EGGTIM  WHERE HME >= '" + Format(f_dat, "mm/dd/yyyy") + "' AND HME <= '" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "' and ATIM='" + f_matim + "'  order by PROOD_AJ"
                End If
414             RTEST.Close
416             m_tim.Open "select * from TIM WHERE ID_NUM=" + str(m_ID), Gdb, adOpenDynamic, adLockOptimistic
                 
            Else
418             EID.RecordSource = "select EGGTIM.EKPT2 AS EKPT22,EGGTIM.*,EGGTIM.EKPT AS EKPTOSH from EGGTIM  WHERE HME >= '" + Format(f_dat, "mm/dd/yyyy") + "' AND HME <= '" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "' and ATIM='" + f_matim + "'  order by PROOD_AJ"
420             m_tim.Open "select * from TIM  WHERE HME  >= '" + Format(f_dat, "mm/dd/yyyy") + "' AND HME <= '" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "' and ATIM='" + f_matim + "'  ", Gdb, adOpenDynamic, adLockOptimistic

            End If

422         EID.Refresh
            '   On Error Resume Next
424         EID.Recordset.MoveFirst

            Dim mAji As Double
            Dim MFPA As Double
            
426         If m_tim.EOF Then
428             mAji = 0: MFPA = 0
            Else
430             mAji = m_tim("aji")
432             MFPA = m_tim!FPA1 + m_tim!fpa2 + m_tim!fpa3 + m_tim!FPA4 + m_tim!FPA6 + nNull(m_tim!FPA7) + nNull(m_tim!FPA8) + nNull(m_tim!FPA9)
                
            End If

            ' ΠΑΙΡΝΕΙ ΤΑ ΣΤΟΙΧΕΙΑ ΑΠΟ ΤΟ ΗΔΗ ΥΠΑΡΧΟΝ ΤΙΜΟΛΟΓΙΟ
            STREET.Text = m_tim("STREET")
            STREETNUMBER.Text = CNull(m_tim("STREETNUMBER"))
            POSTALCODE.Text = m_tim("POSTALCODE")
            CITY.Text = m_tim("CITY")
            BRANCH.Text = m_tim("BRANCH")

434         If F_METAFORIKES >= 1 Then
436             Textm(2).Text = m_tim!c2
438             Textm(3).Text = m_tim!c3
440             Textm(0).Text = m_tim!C1
442             Textm(1).Text = str(m_tim!NUM1)
444             Textm(4).Text = str(m_tim!NUM2)
             
446             Textm(5).Text = str(m_tim!NUM11)
448             Textm(6).Text = (m_tim!C12)
450             Textm(7).Text = (m_tim!C13)
                 
                TC5 = m_tim!c5
                TC6 = m_tim!c6
                tc7 = m_tim!c7
             
             
             
             
             
                Dim j As Integer
452             If F_IS_BYTIO >= 1 Then
454                 For j = 0 To 7
456                     Textm(20 + j) = Textm(j)
                    Next
                End If
             
            End If

458         k = 1

            Dim DUMY

            Dim M_APOT

            ' ΟΤΑΝ ΕΧΩ ΕΝΔΟΔΙΑΚΙΝΗΣΗ ΝΑ ΜΗΝ  ΦΕΡΝΕΙ ΔΙΠΛΑ ΤΑ ΕΙΔΗ ΑΛΛΑ ΝΑ ΦΕΡΝΕΙ ΜΟΝΟ ΑΠΟ ΤΗΝ ΜΙΑ ΑΠΟΘΗΚΗ
460         M_APOT = EID.Recordset("APOT")

            Dim r_eidh As New ADODB.Recordset

462         Do While Not EID.Recordset.EOF

464             If M_APOT = EID.Recordset("APOT") Then

466                 XROMATA(k, 1, 0) = Left(EID.Recordset("PROELEYSH"), 3)
468                 XROMATA(k, 0, 1) = mID(EID.Recordset("PROELEYSH"), 4, 2)

470                 Grid1.TextMatrix(k, f_k) = EID.Recordset("kode")
472                 Grid1.TextMatrix(k, f_p) = Replace(str(EID.Recordset("POSO")), ",", ".")

474                 Grid1.TextMatrix(k, F_PROEL) = EID.Recordset("PROELEYSH")
476                 Grid1.TextMatrix(k, F_mikta_KILA) = Replace(str(EID.Recordset("miktakila")), ",", ".")
478                 Grid1.TextMatrix(k, F_COL_KOLA) = Replace(str(EID.Recordset("kola")), ",", ".")
                    
                    'εαν δουλεύω ΕΦΚ τότε αποθήκευσε το στο τιμολόγιο
480                 If F_EIDFOROSKAT = 1 And F_AJIA_APOU > 0 Then  ' F_PIN_PROELEYS = 12 Then
                        'ΕΠΕΙΔΗ ΕΧΩ APOUHKEYMENO TON EFK GIA TO SYNOLO THS POSOTHTAS
482                     If EID.Recordset("POSO") > 0 Then
484                         Grid1.TextMatrix(k, f_colEFK) = Replace(str(EID.Recordset("efk") / EID.Recordset("POSO")), ",", ".")
                        Else
486                         Grid1.TextMatrix(k, f_colEFK) = Replace(str(EID.Recordset("efk")), ",", ".")
                        End If
                    End If

                    ' Grid1.TextMatrix(k, f_e) = Replace(str(nNull(EID.Recordset("EKPT"))), ",", ".")
488                 Grid1.TextMatrix(k, F_EKPT2) = Replace(str(nNull(EID.Recordset("ekpt22"))), ",", ".")
490                 Grid1.TextMatrix(k, f_o) = EID.Recordset("onoma")

492                 Grid1.TextMatrix(k, f_e) = Replace(str(nNull(EID.Recordset("EKPTOSH"))), ",", ".")

                    ' If F_XONDR = 0 Then ' lianikh
                    '    Grid1.TextMatrix(k, F_T) = Replace(str(((100 + g_Fpa(EID.Recordset("FPA"))) / 100) * EID.Recordset("TIMM")), ",", ".")
                    'Else
494                 If TIMES0XORIS1 = 1 Then
496                     Grid1.TextMatrix(k, F_T) = 0
498                 ElseIf TIMES0XORIS1 = 0 Then
500                     Grid1.TextMatrix(k, F_T) = Replace(str(EID.Recordset("TIMM")), ",", ".")
                        
502                     If f_TIMOKAT = 9999 And mpar <> 99 Then
                   
504                         Grid1.TextMatrix(k, F_T) = GGET_NVALUE("SELECT ISNULL(LTI,0) FROM EID WHERE KOD='" + EID.Recordset("kode") + "'")

                            Dim timokPEL As String, timokEID As String

506                         timokPEL = Trim(IIf(IsNull(PEL.Recordset("CH4")), "", PEL.Recordset("CH4")))
508                         timokEID = GGET_CVALUE("SELECT TOP 1 CH1 FROM EID WHERE KOD='" + EID.Recordset("kode") + "'")

                            Dim rSQL2 As New ADODB.Recordset

510                         rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOKPEL='" + timokPEL + "' AND TIMOKEID='" + LTrim(timokEID) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

512                         If rSQL2.EOF Then
514                             Grid1.TextMatrix(k, f_e) = 0
                            Else
516                             Grid1.TextMatrix(k, f_e) = rSQL2(0)
                            End If

518                         rSQL2.Close
                        End If 'f_TIMOKAT = 9999 Then
                        
                    End If 'TIMES0XORIS1 = 1 Then
                    
                    'End If

520                 Grid1.TextMatrix(k, F_mikta) = EID.Recordset("mikta")    ' χρειαζονται για το συγκεντροτικο

                    'ΣΕ ΠΕΡΙΠΤΩΣΗ ΠΟΥ ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΕΙΔΟΣ ΤΟ ΚΑΤΑΛΑΒΑΙΝΕΙ(Π.Χ. ΜΠΟΡΕΙ ΝΑ ΣΒΗΣΤΗΚΕ)
                    ' ΚΑΙ ΕΠΙΣΗΣ ΒΑΖΕΙ ΤΟ ΦΠΑ ΚΑΙ ΤΗΝ ΜΟΝΑΔΑ ΑΠΟ ΤΟ ΕΙΔΟΣ
522                 r_eidh.Open "select FPA,MON from EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic
524                 If r_eidh.EOF Then
526                     MsgBox "Προσοχή δεν υπάρχει στο αρχείο ειδών ο κωδικός " + Grid1.TextMatrix(k, f_k)
528                     MsgBox "***Nα δημιουργηθεί άμεσα ο κωδικός " + Grid1.TextMatrix(k, f_k) + "   λάθος:1100 gridpelaton_dblclick)"
                    Else

                        'Grid1.TextMatrix(K, F_M) = r_eidh("MON")

                        '  If IsNull(r_eidh("FPA")) Then
                        '    Grid1.TextMatrix(K, f_f) = EID.Recordset("FPA")
                        ' Else
                        
                        '    Grid1.TextMatrix(K, f_f) = r_eidh("FPA")
                        ' End If
530                     Grid1.TextMatrix(k, f_f) = EID.Recordset("FPA")
                           
532                     If IsNull(r_eidh("MON")) Then
534                         Grid1.TextMatrix(k, F_M) = EID.Recordset("MON")
                        Else
536                         Grid1.TextMatrix(k, F_M) = r_eidh("MON")
                        End If

                    End If
538                 r_eidh.Close
                    
540                 Grid1.ColAlignment(f_p) = 6 'RIGHT
542                 Grid1.ColAlignment(f_e) = 6 'RIGHT
544                 Grid1.ColAlignment(F_T) = 6 'RIGHT

546                 Grid1.TextMatrix(k, f_a) = Round(EID.Recordset("POSO") * EID.Recordset("TIMM") * (100 - EID.Recordset("ekptOSH")) / 100, 2)

548                 Grid1.row = k: Grid1.Col = 0: FindSynolo

550                 k = k + 1

                End If
                
552             If k = f_MinSeiresTimologioy + 2 Then
554                 MsgBox " ΠΡΟΣΟΧΗ ΠΡΕΠΕΙ ΝΑ ΜΕΓΑΛΩΣΕΤΕ ΤΙΣ ΣΕΙΡΕΣ ΤΟΥ ΤΙΜΟΛΟΓΙΟΥ ΓΙΑΤΙ ΘΑ ΚΟΠΟΥΝ ΤΑ  ΠΑΡΑΠΑΝΙΣΙΑ ΕΙΔΗ " + Chr(13) + " TΩΡΑ EINAI " + str(f_MinSeiresTimologioy)
                
                End If

556             EID.Recordset.MoveNext
            Loop
        
558         FindSynolo
        
560         Grid1.row = k

562         If Round(Abs(gVal(SynoloG.Caption) - mAji), 2) = 0.01 And Round(Abs(gVal(sfpa.Caption) - MFPA), 2) = 0.01 Then
564             SynoloG.Caption = Format(mAji, "#####0.00")
566             sfpa.Caption = Format(MFPA, "#####0.00")
    
568             If mAji > gVal(SynoloG.Caption) Then
570                 f_strogg = -0.01
                Else
572                 f_strogg = 0.01
                End If

            End If

574         If k > 1 Then
576             timtext3.Text = Grid1.TextMatrix(k + 1, f_k)
            End If

578         F_DIOR = 1
580         Grid1.Col = f_k  ' για να μην ζητάει να ανοιξει νέα καρτέλλα
        End If ' ΝΑ ΔΙΟΡΘΩΘΕΙ?

        '----ΤΕΛΟΣ ------------------------------------ ναι να διορθωθει .//  ΒΛΕΠΩ ΤΑ ΕΙΔΗ ---------------------------------------



         ' αν ειναι στα παραστατικα που ειναι αναποδα
         If InStr(F_ANAPODO, Left$(f_matim, 1)) > 0 And F_DIOR = 0 Then
             STREET.Text = FORTDIE.Text
             STREETNUMBER.Text = fortARit.Text
             POSTALCODE.Text = FORTTK.Text
             CITY.Text = FORTPOL.Text
             '----------------------------------
              FORTDIE.Text = PEL.Recordset("DIE")
               fortARit.Text = CNull(PEL.Recordset("STREETNUMBER2"))
               FORTPOL.Text = PEL.Recordset("POL")
               FORTTK.Text = CNull(PEL.Recordset("XRVMA"))
   
        End If











582     If F_METAFORIKES = 1 Then

584         Textm(1).SetFocus

        Else

586         timtext3.Text = Grid1.Text
588         timtext3.SetFocus

        End If

        Exit Sub

MHNYMA:
590     HandleError "Par1:GridPelAton_dblclick"

592     Resume Next

        '<EhFooter>
        Exit Sub

GridPelaton_DblClick_Err:
        ' 'MsgBox Err.Description & vbCrLf & _
           "in ADOMERCNEW.par1.GridPelaton_DblClick " & _
           "at line " & Erl, _
           vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.mmGridPelaton_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GridPelaton_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo GridPelaton_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
102         KeyAscii = 0    'suppress the beep
            ' keybd_event VK_TAB, 0, 0, 0 'send a tab
        End If

        '<EhFooter>
        Exit Sub

GridPelaton_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.GridPelaton_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.GridPelaton_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GridPelaton_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo GridPelaton_KeyUp_Err

        '</EhHeader>
100     If KeyCode = 13 Then
102         KeyCode = 0
104         PAR1.KeyPreview = False

106         mmGridPelaton_DblClick 1, 0, 0, 0

        End If

        '<EhFooter>
        Exit Sub

GridPelaton_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.GridPelaton_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.GridPelaton_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub metasxhmatismos_Click()

        '===================================
        'συγκεντρώνω τα δελτία που σημάδεψε ο χρήστης
        '===================================
        '<EhHeader>
        On Error GoTo metasxhmatismos_Click_Err

        '</EhHeader>
        Dim DELT As Recordset, k As Integer, L As Integer

        Dim e, s As String, D As String, T As Recordset

        Dim X, c, Eidh As Recordset

        Dim R As Recordset

        On Error GoTo MHNYMA

        Dim r2 As New ADODB.Recordset

        Dim R5 As New ADODB.Recordset
    
100     k = 0
        'Set delt = EID.Recordset

102     If EID.Recordset.RecordCount = 0 Then

            Exit Sub

        End If

        ' ΣΥΓΚΕΝΤΡΩΝΩ ΤΑ ΣΗΜΑΔΕΜΕΝΑ ΣΤΟ S

104     s = ""

        Dim mSeira As String

        Dim SX     As String

106     SX = ""
        '"select METAF,ATIM,CONVERT(CHAR(10),HME,103) AS [Ημερομηνία],ID_NUM,isnull(ENTITYMARK,'') AS MARK from TIM where left(ATIM,1) in ('" + Get_meKomma(f_metasx) + "') and KPE='" + PEL.Recordset("kod") + "' and ART IS null ORDER BY HME;"
108     For L = 1 To FlxGridDeltion.rows - 1

110         If Val(FlxGridDeltion.TextMatrix(L, 1)) = 1 Then
112             k = k + 1
114             f_deltia(k) = Left(FlxGridDeltion.TextMatrix(L, 2) + Space(8), 8) + FlxGridDeltion.TextMatrix(L, 3)
116             f_IDdeltia(k) = FlxGridDeltion.TextMatrix(L, 4)

118             f_deltiaMarks(k) = FlxGridDeltion.TextMatrix(L, 5)
120             s = s + "'" + f_deltia(k) + "'" + ","   'ΤΑ ΔΕΛΤΙΑ ΠΟΥ ΕΧΩ ΣΗΜΑΔΕΨΕΙ
            
122             R5.Open "select SYNT_TITL FROM PARASTAT WHERE EIDOS='" + Left(f_deltia(L), 1) + "'", Gdb, adOpenDynamic, adLockOptimistic

124             If IsNull(R5(0)) Then
126                 mSeira = ""
                Else
128                 mSeira = Trim(R5(0)) + "-"
                End If

130             R5.Close
132             If f_METASX_KAI_HMER = 1 Then
134                 SX = SX + mSeira + Format(Val(mID(f_deltia(k), 2, f_psifia_atim)), "######") + " " + FlxGridDeltion.TextMatrix(L, 3) + ","
                Else
136                 SX = SX + mSeira + Format(Val(mID(f_deltia(k), 2, f_psifia_atim)), "######") + ","
                End If
            End If

        Next
138     par11Value.Caption = f_deltiaMarks(k)
140     s = s + "''"

142     If f_metasx = "a" Then
144         parat.Text = Trim(f_LEKTIKO_ARPARAGG) + SX
        Else
146         parat.Text = "Σχετ.Δελτ." + SX
        End If

148     fSHMADEMENA = s

150     If k = 0 Then  ' DEN SHMADECE TIPOTA
152         FlxGridDeltion.Visible = False
154         GridEidon.Visible = True

            Exit Sub

        End If

156     D = EID.RecordSource

        ' σημάδεψε δελτια // ψάχνω να βρώ τα είδη
158     If k = 1 Then
            '"order by PROOD_AJ"
160         EID.RecordSource = "select POSO as SPOSO,KODE,TIMM,ONOMA,FPA,EKPT from EGGTIM where left(ATIM+space(8),8)+CONVERT(CHAR(10),HME,103)  IN (" + s + ") order by PROOD_AJ;"
162         EID.Refresh

        Else    ' polla deltia
164         EID.RecordSource = "select sum(POSO) as SPOSO,KODE,TIMM,ONOMA,FPA,EKPT from EGGTIM where left(ATIM+space(8),8)+CONVERT(CHAR(10),HME,103)  IN (" + s + ") group by KODE,TIMM,ONOMA,FPA,EKPT "
166         EID.Refresh
        End If

        'EID.RecordSource = "select sum(POSO) as SPOSO,KODE,TIMM,ONOMA from EGGTIM where ATIM+CONVERT(CHAR(10),HME,103)  IN (" + S + ") group by KODE,TIMM,ONOMA;"
        ' EID.Refresh

168     F_SHMADEMENA_DELTIA = s

170     If EID.Recordset.RecordCount > 500 Then
172         MsgBox "Πάνω από 500 είδη στο τιμολόγιο. Διαλέξτε λιγότερα δελτία"
174         EID.RecordSource = D
176         EID.Refresh

            Exit Sub

        End If

        'Set fDB = OpenDatabase(gDir, False, False, gConnect)

        Dim EID2 As New ADODB.Recordset

        'EID2.Open "SELECT * FROM EID", Gdb, adOpenDynamic, adLockOptimistic

        'If EID.Recordset.RecordCount = 0 Then
        '   EID.RecordSource = "SELECT *FROM EID"
        '   EID.Refresh
        '   Exit Sub
        'End If

        'ΒΡΙΣΚΕΙ ΤΑ ΕΙΔΗ ΚΑΙ ΒΑΖΕΙ ΠΕΡΙΓΡΑΦΕΣ ΚΑΙ ΤΑ ΑΠΛΩΝΕΙ ΣΤΟ GRID
178     EID.Recordset.MoveFirst
180     k = 1

182     Do While Not EID.Recordset.EOF

184         EID2.Open "SELECT *FROM EID WHERE KOD='" + EID.Recordset("kode") + "'", Gdb, adOpenDynamic, adLockOptimistic
186         Grid1.TextMatrix(k, f_k) = EID.Recordset("kode")
188         Grid1.TextMatrix(k, f_p) = EID.Recordset("SPOSO")

190         If Left(Grid1.TextMatrix(k, f_k), 1) = "*" Then
192             Grid1.TextMatrix(k, f_o) = EID.Recordset("ONOma")
            Else
194             Grid1.TextMatrix(k, f_o) = EID2("ONO")
            End If

196         Grid1.TextMatrix(k, F_M) = EID2("MON")

198         If EID.Recordset("TIMM") = 0 Then
200             Grid1.TextMatrix(k, F_T) = EID2("LTI")
            Else
202             Grid1.TextMatrix(k, F_T) = EID.Recordset("TIMM")    ' ΠΑΡΑΓΓΕΛΙΑ
            End If

            '******************   15-6-2010  ************************************
            'ΕΚΠΤΩΣΕΙΣ
204         Grid1.Col = f_e

206         If f_TIMOKAT > 0 And Not IsNull(PEL.Recordset("pek")) Then
208             r2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + str(PEL.Recordset("pek")) + " AND KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

210             If r2.EOF Then
212                 Grid1.Text = ""
                Else
                    '  Grid1.Text = r2(0)
214                 Grid1.TextMatrix(k, f_e) = r2(0)
                End If

216             r2.Close

                'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
218         ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then
220             r2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + str(PEL.Recordset("pek")) + " AND KOD='" + Trim(Left(EID.Recordset("kod"), 13)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

222             If r2.EOF Then
                    ' Grid1.TextMatrix(k, F_T) = 0
                Else
224                 Grid1.TextMatrix(k, F_T) = r2(0)
                End If

226             r2.Close
228         ElseIf f_TIMOKAT = 9999 Then
230             Dim typos_par As Integer: typos_par = GGET_NVALUE("SELECT ISNULL(STADIO,0) FROM PARASTAT WHERE EIDOS='" + Left$(f_matim, 1) + "'")
232             If typos_par = 3 Then ' D.A.
234                 Grid1.TextMatrix(k, F_T) = 0
                Else
                    Dim timok2PEL As String, timok2EID As String
236                 timok2PEL = Trim(IIf(IsNull(PEL.Recordset("CH4")), "", PEL.Recordset("CH4")))
238                 timok2EID = GGET_CVALUE("SELECT TOP 1 CH1 FROM EID WHERE KOD='" + EID.Recordset("kode") + "'")

                    Dim rSQL2 As New ADODB.Recordset

240                 rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOKPEL='" + timok2PEL + "' AND TIMOKEID='" + LTrim(timok2EID) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

242                 If rSQL2.EOF Then
244                     Grid1.TextMatrix(k, f_e) = 0
                    Else
246                     Grid1.TextMatrix(k, f_e) = rSQL2(0)
                    End If

248                 rSQL2.Close
                End If 'f_
            
            Else
                ' If F_EKPT_PARAGG = 1 Then
                '   Grid1.TextMatrix(K, f_e) = EID.Recordset("EKPT")
              
                ' Else
              
250             Grid1.TextMatrix(k, f_e) = PEL.Recordset("pek")
                ' End If
252             r2.Close

            End If
            
254         If F_EKPT_PARAGG = 1 Then
             
256             Grid1.TextMatrix(k, f_e) = EID.Recordset("EKPT")
258             Grid1.TextMatrix(k, F_T) = EID.Recordset("timm")
            End If

            '******************   15-6-2010  ************* Grid1.TextMatrix(K, f_f) = EID.Recordset("fpa")
            '******************   15-6-2019  **  Grid1.TextMatrix(K, f_f) = EID2("FPA") **********************************
260         Grid1.TextMatrix(k, f_f) = EID2("FPA") ' EID.Recordset("fpa")

262         Grid1.TextMatrix(k, f_a) = Round(EID.Recordset("SPOSO") * Grid1.TextMatrix(k, F_T), 2) * (100 - Val(Grid1.TextMatrix(k, f_e))) / 100

264         Grid1.row = k: Grid1.Col = 0: FindSynolo

266         EID2.Close
268         k = k + 1
270         EID.Recordset.MoveNext
        Loop

272     metasxhmatismos.Enabled = False
        'EID.RecordSource = "SELECT *FROM EID"
        'EID.Refresh
274     FindSynolo
276     Grid1.row = 1
        'Grid1.SetFocus

278     Text2(0).Enabled = False
280     Text2(1).Enabled = False
282     Text2(2).Enabled = False

284     Grid1.Col = 1
286     timtext3.Top = Grid1.Top + Grid1.CellTop
288     timtext3.Left = Grid1.Left + Grid1.CellLeft
290     timtext3.Width = Grid1.CellWidth - 15
292     timtext3.Height = Grid1.CellHeight - 25
294     timtext3.Text = Grid1.Text
296     timtext3.SetFocus
298     GridPelaton.Visible = False
300     FlxGridDeltion.Visible = False
302     GridEidon.Visible = True

        Exit Sub

MHNYMA:
304     HandleError "Par1:metasxhmatismos"

306     Resume Next

        '<EhFooter>
        Exit Sub

metasxhmatismos_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.metasxhmatismos_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.metasxhmatismos_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub selectPEL_Click()

        ' Διαλέγει πελάτη και δείχνει δελτία σε εκκρεμότητα
        '===================================================
        ' Διαλέγει πελάτη και δείχνει δελτία σε εκκρεμότητα
        '===================================================
        '   GridPelaton.Visible = True
        '<EhHeader>
        On Error GoTo selectPEL_Click_Err

        '</EhHeader>

100     Text2(0).Enabled = True
102     Text2(1).Enabled = True
104     Text2(2).Enabled = True
106     Text2(0).SetFocus
        '    SendKeys ""

        ' αν είναι τιμολόγιο του δείχνει τα δελτία σε εκκρεμότητα
108     If Len(f_metasx) > 0 Then
            '
            '  EID.RecordSource = "select metaf,atim,hme from tim where left(atim,1)='" + f_metasx + "' and kpe='" + PEL.Recordset("kod") + "' and art=null;"
            ' EID.Refresh

            ' metasxhmatismos.Visible = True
            ' FlagGridEidon = False
        End If

        '<EhFooter>
        Exit Sub

selectPEL_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.selectPEL_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.selectPEL_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GridEidon_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

        '==================================
        'επιλογή είδους ή δελτίων αποστολής
        '==================================
        '<EhHeader>
        On Error GoTo GridEidon_RowColChange_Err

        '</EhHeader>
        Dim DELT As Recordset

        On Error GoTo MHNYMA

100     If Grid1.Enabled = False Then    ' φάση δελτίων αποστολής

102         If FlagGridEidon Then
                ' FlagGridEidon = False
104             Set DELT = EID.Recordset

106             If IsNull(DELT("metaf")) Or DELT("metaf") = 0 Then
108                 DELT("metaf") = 1: DELT.Update
                Else
110                 DELT("metaf") = 0: DELT.Update
                End If
            End If

        Else

            '   TIMEDIT.Recordset.Edit
            '   TIMEDIT.Recordset("kode") = EID.Recordset("kod")
            '   TIMEDIT.Recordset("onoma") = EID.Recordset("ono")
            '   If fBuff = "r" Then
            '      TIMEDIT.Recordset("timm") = EID.Recordset("xti")
            '   Else
            '      TIMEDIT.Recordset("timm") = EID.Recordset("lti")
            '   End If
            '   TIMEDIT.Recordset("fpa") = EID.Recordset("fpa")
            '   TIMEDIT.Recordset.Update
            '   DBGrid1.Refresh

        End If

        Exit Sub

MHNYMA:
112     HandleError "Par1:GridEidon_RowColChange"

114     Resume Next

        '<EhFooter>
        Exit Sub

GridEidon_RowColChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.GridEidon_RowColChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.GridEidon_RowColChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '==============================
        '==============================
        '==============================
        '==============================
        '<EhHeader>
        On Error GoTo Form_Load_Err
Dim TT As Long: TT = GetCurrentTime()

SSTab1.Tab = 0

Const F_ARR_TYPOS = 0
Const F_ARR_ID = 1
Const F_ARR_AYJON = 2
Const F_ARR_TIMH = 3
Const F_ARR_PERIGRAFH = 4
Const F_ARR_C1 = 5
'
Dim f_Arr_PIN(600, 5)
Dim R_arr As New ADODB.Recordset
 R_arr.Open "SELECT TYPOS,ID,AYJON,TIMH,PERIGRAFH,C1 FROM PINAKES WHERE TYPOS in (4,7,8,9,10,44) ORDER BY TYPOS,AYJON", Gdb, adOpenDynamic, adLockOptimistic
     Dim k_ARR As Long: k_ARR = 0
     Do While Not R_arr.EOF
           k_ARR = k_ARR + 1
           f_Arr_PIN(k_ARR, F_ARR_TYPOS) = R_arr!TYPOS
           f_Arr_PIN(k_ARR, F_ARR_ID) = R_arr!ID

           f_Arr_PIN(k_ARR, F_ARR_AYJON) = R_arr!AYJON
           f_Arr_PIN(k_ARR, F_ARR_TIMH) = R_arr!TIMH
           f_Arr_PIN(k_ARR, F_ARR_PERIGRAFH) = R_arr!PERIGRAFH
           f_Arr_PIN(k_ARR, F_ARR_C1) = R_arr!C1




           R_arr.MoveNext
        Loop
R_arr.Close





        '</EhHeader>

        '    GetLocation Me
        '    CenterForm Me
        '    ResizeForm Me
        
100     CURR.ListIndex = 0 '"" 'CURRENCY.List (0)
        
102     f_same_eline = 0

104     F_METAFORIKES = Val(FINDPARAMETROI(1, "PAR1", "F_METAFORIKES", "0", "Πρόκειται για μεταφορική=1 ΟΧΙ=0")) '
        'Dim F_METAF3SYNAL As Integer 'ΜΕΤΑΦΟΡΙΚΕΣ ΜΕ 3 ΣΥΝΑΛΛΑΣΟΜΕΝΟΥΣ =1 ΟΧΙ=0
106     F_METAF3SYNAL = Val(FINDPARAMETROI(1, "PAR1", "F_METAF3SYNAL", "0", "ΜΕΤΑΦΟΡΙΚΕΣ ΜΕ 3 ΣΥΝΑΛΛΑΣΟΜΕΝΟΥΣ =1 ΟΧΙ=0")) '
        F_ELAIO_DIK = gVal(FINDPARAMETROI(1, "PAR1", "F_ELAIO_DIK", "0", "ΠΟΣΟΤΟ % ΔΙΚΑΙΩΜΑΤΟΣ ΓΙΑ ΕΛΑΙΟΤΡΙΒΕΙΟΠ.Χ. 17.5  ΟΧΙ=0")) '

        F_CUSTOMER_SCREEN = Val(FINDPARAMETROI(1, "PAR1", "F_CUSTOMER_SCREEN", "0", "1= ΟΘΟΝΗ ΠΕΛΑΤΗ ΟΧΙ=0")) '
       F_CUSTOMER_SCR_LEFT = Val(FINDPARAMETROI(1, "PAR1", "F_CUSTOMER_SCR_LEFT", "20000", "LEFT(ΑΡΧΗ) ΟΘΟΝΗΣ ΠΕΛΑΤΗ")) '
        
        If F_CUSTOMER_SCREEN = 1 Then
            PAR1CUSTOMER.Left = F_CUSTOMER_SCR_LEFT
            PAR1CUSTOMER.SHOW
        End If

108     F_MONOENERGOS = Val(FINDPARAMETROI(1, "PAR1", "F_MONOENERGOS", "0", "ΔΕΙΧΝΕΙ ΜΟΝΟ ΤΟΥΣ ΕΝΕΡΓΟΥΣ=1 ΟΛΟΥΣ=0")) '

        f_TimhForoyDiuamonis = Val(FINDPARAMETROI(1, "PAR1", "f_TimhForoyDiuamonis", "0", "ΤΙΜΗ ΦΟΡΟΥ ΔΙΑΜΟΝΗΣ")) '

        F_2EKTYPOTHS = FINDPARAMETROI(1, "PAR1", "F_2EKTYPOTHS", "..", "Εκτυπωτής2 AKYR παραστατικών(με .. προεπιλεγμένος)") '

110     F_EKTYPOTHS = FINDPARAMETROI(1, "PAR1", "F_EKTYPOTHS", "..", "Εκτυπωτής παραστατικών(με .. προεπιλεγμένος)") '

        'F_PARAS_3SYNAL

112     F_THERM = Val(FINDPARAMETROI(1, "PAR1", "F_THERM", "0", "Πρόκειται για ΘΕΡΜΑΝΣΗ=1 ΟΧΙ=0")) '

114     F_THER_TIM_XORISAPOD = Val(FINDPARAMETROI(1, "PAR1", "F_THER_TIM_XORISAPOD", "0", "ΣΤΗΝ ΘΕΡΜΑΝΣΗ ΤΟ ΤΙΜΟΛΟΓΙΟ ΝΑ ΜΗΝ ΖΗΤΑΕΙ ΑΠΟΔΕΙΞΗ=1 ΝΑ ΖΗΤΑ=0")) '

116     F_IS_BYTIO = Val(FINDPARAMETROI(1, "PAR1", "F_IS_BYTIO", "0", "βυτιο=1 δικ.επιμ=2 βυτιο+μεταφ=3 4=ελαιοτριβειο  ΟΧΙ=0")) '
118     F_PARAS_3SYNAL = FINDPARAMETROI(1, "PAR1", "F_PARAS_3SYNAL", ";", "ΧΑΡΑΚΤΗΡ.ΓΡΑΜΜΑ ΜΕΤΑΦ.ΑΠΟΣ-ΠΑΡΑΛ-ΕΝΤΟΛ Π.Χ. Q") '
 

120     f_SDA = FINDPARAMETROI(1, "PAR1", "F_SDA", "Ω;V", "ΧΑΡΑΚΤΗΡ.ΓΡΑΜΜΑ ΣΔΑ ΕΝΑΡΞΗΣ ΚΑΙ ΕΠΙΣΤΡΟΦΗΣ Π.Χ. Φ;Ψ") '
  FKODPELSYGK = FINDPARAMETROI(1, "PAR1", "FKODPELSYGK", "0", "KΩΔΙΚΟΣ ΠΕΛΑΤΗ ΣΥΣΚΕΝΤΡΩΤΙΚΟΥ") '
 
 
 
122     If F_METAFORIKES = 2 Then
 
124         cmdCAI.Visible = True
 
        End If
 
126     If F_METAFORIKES = 1 Then
128         Frame2(0).Height = Frame2(0).Height + 300
130         Grid1.Top = 2820 ' 2520
132         Grid1.Height = 2000 ' 2295
     
134         GridPelaton.Top = 2820 '2520
136         GridPelaton.Height = 2000 '  2295  '2295
138         Frame2(0).Visible = True
140         DoEvents
142         PAR1.Refresh
144         DoEvents
146         PAR1.Refresh
148         Frame2(0).Top = 1320
     
150         If F_IS_BYTIO >= 1 Then
                'Frame2(0).Top = 1320
                'Frame2(0).Height = Frame2(0).Height + 300
152             Frame2(0).Visible = False
       
154             Frame2(1).Top = 1320
156             Frame2(1).Height = Frame2(0).Height '+ 300
158             Frame2(1).Visible = True
160             Textm(20).Height = Frame2(1).Height - Textm(20).Top
       
162             If F_IS_BYTIO = 2 Then  ' dikastikos κλητηρας
                    Dim j As Integer
164                 For j = 1 To 7
166                     Textm(20 + j).Visible = False
                    Next
168                 Label24.Visible = False
170                 Label25.Visible = False
172                 MTF11(1).Visible = False
174                 MTF12(1).Visible = False
176                 MTF13(1).Visible = False
178                 MTF14(1).Visible = False
180                 MTF21(1).Visible = False
             
                End If
                If F_IS_BYTIO = 4 Then ' ελαιτριβειο
                   
                   Textm(23).Visible = False
                   Textm(25).Visible = False
                   'Textm(26).Visible = False
                  ' Textm(27).Visible = False
                   
                   ' Label24.Visible = False
                   ' Label25.Visible = False
                    MTF13(1).Visible = False
                    MTF21(1).Visible = False
                
                
                
                End If
                
       
            End If
     
            ' ΜΕΤΑΦΟΡΙΚΕΣ ΜΕ 3 ΣΥΝΑΛΛΑΣΟΜΕΝΟΥΣ =1 ΟΧΙ=0
 
182         If F_METAF3SYNAL = 1 Then
       
                'MON  =>C2
                'POSOTHTA =>num1
                'ARIU.PARAST =>c12
                'PERIGRAFH =>c13
                'TEXTM(3)  APOSTOLEAS=>C3   TEXTM(5) ONOMA APOSTOLEA
                'TEXTM(0)  PARALHPTHS=>C0   TEXTM(4) ONOMA PARALHPTH
       
                'ΑΝΤΙΜΕΤΑΘΕΤΩ  ΤΟ ΤΕΧΤΜ(2) ΜΕ ΤΟ ΤΕΧΤΜ(5)  ΓΙΑ ΝΑ ΕΧΩ ΤΕΧΤ ΣΤΗΝ ΜΟΝΑΔΑ to textm(2) kai to textm(5) sthn anazhthsh
184             Textm(5).Left = Textm(2).Left
186             Textm(5).Top = Textm(2).Top
188             Textm(5).Width = Textm(2).Width
190             Textm(5).Height = Textm(2).Height
       
192             Textm(0).Height = Frame2(1).Height - Textm(0).Top

194             Textm(3).Height = Frame2(1).Height / 2 ' Frame2(1).Height - Textm(3).Top
196             Textm(0).Height = Frame2(1).Height / 2 ' Textm(3).Height

                '--------------LEFTS -----------------
198             Textm(2).Left = Textm(4).Left
200             Textm(1).Left = Textm(0).Left

                '--------------  TOPS ---------------------
202             Textm(6).Top = Textm(6).Top + 300
204             Textm(1).Top = Textm(6).Top
206             Textm(2).Top = Textm(6).Top
208             Textm(7).Top = Textm(6).Top
                '----------- widths -------------
210             Textm(2).Width = Textm(4).Width

212             Textm(1).Height = Textm(5).Height

                'labels
                Dim pmet As Long
214             pmet = 160

216             Labf(0).Visible = True
218             Labf(1).Visible = True
         
220             Labf(1).Top = Textm(1).Top - pmet
                ' Labf(2).Top = Textm(6).Top - pmet
222             Labf(3).Top = Textm(7).Top - pmet

                'Labf(0).Left = Textm(2).Left 'MON  =>C2
                'Labf(1).Left = Textm(1).Left  'POSOTHTA =>num1
                'Labf(2).Left = Textm(6).Left  'ARIU.PARAST =>c12
                'Labf(3).Left = Textm(7).Left  'PERIGRAFH =>c13

                '  Labf(0).Width = Textm(2).Width
                '   Labf(1).Width = Textm(1).Width
                '   Labf(2).Width = Textm(6).Width
                '   Labf(3).Width = Textm(7).Width

224             Labf(0).Caption = f_CAPTION_labf0 '"μον"
226             Labf(1).Caption = f_CAPTION_labf1 '"ποσοτητα"

228             Labf(2).Caption = f_CAPTION_labf2 '"αριθ.παρ"      'textm(6)  anazhthsh entolea

230             Labf(3).Caption = f_CAPTION_labf3 '"περιγραφη"     'textm(7)    perigrafh entolea
         
                'Labf(0).Visible = True
                'Labf(1).Visible = True
         
232             Textm(3).BackColor = Textm(0).BackColor
         
                '2->3  4->0
                Dim met As Long ' metatopish pro aristera olon ton pedion
234             met = 1200
                '  Textm(2).Left = Textm(2).Left - met
236             Textm(3).Left = Textm(3).Left - met
238             Textm(4).Left = Textm(4).Left - met
240             Textm(5).Left = Textm(5).Left - met
242             Textm(0).Left = Textm(0).Left - met
244             MTF11(0).Left = MTF11(0).Left - met
246             MTF12(0).Left = MTF12(0).Left - met
248             MTF13(0).Left = MTF13(0).Left - met
250             MTF14(0).Left = MTF14(0).Left - met
252             MTF21(0).Left = MTF21(0).Left - met
254             TextOnEntolea.Left = TextOnEntolea.Left - met
256             TextEntoleas.Left = TextEntoleas.Left - met
258             TextEntoleas.Width = Textm(0).Width + 300
260             TextEntoleas.BackColor = Textm(0).BackColor
262             TextEntoleas.Height = Textm(0).Height
         
                ' TextEntoleas.MultiLine = True
         
264             Labf(0).Top = Textm(2).Top: Labf(0).Left = Textm(2).Left - 300: Textm(2).Width = 600 '    - mon.met
266             Labf(2).Top = Textm(6).Top: Labf(2).Left = Textm(6).Left - 600: Textm(6).Width = 1000:  'arium
          
268             Labf(1).Top = Textm(1).Top: Labf(1).Left = Textm(1).Left + 200: Textm(1).Width = 800: Textm(1).Left = Labf(1).Left + 990  'pos
270             Labf(3).Top = Textm(7).Top: Labf(3).Left = Textm(7).Left - 600: Textm(7).Width = 1900 'perigrafh
       
            End If
     
272         If F_THERM = 1 Then
                ' GRACE ΣΤΗΝ ΠΕΡΙΓΡΑΦΗ (ΜΕΤΑΦΟΡΙΚΩΝ) ΤΟΝ ΤΕΛΕΥΤΑΙΟ ΑΡΙΘΜΟ ΣΤΟΥ ΣΔΑ ΕΝΑΡΞΗΣ
274             Textm(0) = str(GET_NVALUE("SELECT ISNULL(A.ARITMISI,0) AS TEL_SDA FROM PARASTAT P INNER JOIN ARITMISI A ON P.ARITMISI=A.ID WHERE P.EIDOS='" + Left$(f_SDA + "  ", 1) + "'"))
       
            End If
     
        End If

276     mForm_Load PAR1, fh, fw, ft, fl
F_VB6RESIZEFORM = Val(FINDPARAMETROI(1, "PAR1", "F_VB6RESIZEFORM ", "1", "0=RESIZE H FORMA 0=OXI ΩΣ ΕΧΕΙ ")) '
278     f_strogg = 0 ' ΣΤΡΟΓΓΥΛΕΜΑ ΦΠΑ/ΤΙΜΟΛΟΓΙΟΥ

280     f_akyr_neo = Val(FINDPARAMETROI(1, "PAR1", "f_akyr_neo", "0", "Γ.49.ΝΕΟ ΑΚΥΡΩΤΙΚΟ ΑΝΤΙΘΕΤΟ ΧΩΡΙΣ ΔΙΑΓΡΑΦΗ=1 ΟΧΙ=0"))
282     f_akyr_idia_mera = Val(FINDPARAMETROI(1, "PAR1", "f_akyr_idia_mera", "1", "Γ.40.ΑΚΥΡΩΤΙΚΟ ΜΟΝΟ ΤΗΝ ΙΔΙΑ ΜΕΡΑ=1 ΟΧΙ=0"))

        ' F_EIDFOROSKAT  F_ARPART_IN_EID
284     F_ARPART_IN_EID = Val(FINDPARAMETROI(1, "PAR1", "F_ARPART_IN_EID", "0", "Γ.47.Η ΠΑΡΤΙΔΑ ΛΑΧΑΝΑΓ. ΜΕΣΑ ΣΤΟ ΕΙΔΟΣ=1 ΟΧΙ=0"))  '
        
286     F_TIMH_XONDR_TIM_AGOR = Val(FINDPARAMETROI(1, "PAR1", "F_TIMH_XONDR_TIM_AGOR", "1", "Γ.37.ΣΤΟ ΤΙΜ.ΑΓΟΡΑΣ ΤΙΜΗ ΧΟΝΔΡΙΚΗΣ=0 ΤΙΜΗ ΑΓΟΡΑΣ=1"))  'αγορ=1 0=ΤΙΜΗ ΧΟΝΔΡΙΚΗΣ
        
288     f_pedioAFM = FINDPARAMETROI(1, "PELAT2", "f_pedioAFM", "CH1", "Ποιό πεδίο θα αναζητάται μαζί με το ΑΦΜ")
        'Dim F_TYPFOROL  As Integer ' TYPOS FOROLOGIKOY 1=ALGO 2=SIGNPRO
290     F_TYPFOROL = Val(FINDPARAMETROI(1, "PAR1", "F_TYPFOROL", "1", "Γ.33.TYΠOΣ ΦOΡOΛOΓIKOY 1=ALGOBOX 2=SIGNPRO 21=rbs")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
        ' f_is_benzinadiko = 1
292     f_is_benzinadiko = Val(FINDPARAMETROI(1, "PAR1", "f_is_benzinadiko", "0", "Γ.34.EINAI BENZINAΔIKO=1 ΟΧΙ=0"))  'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
 
294     f_AKYR_benzinadiko = Val(FINDPARAMETROI(1, "PAR1", "f_AKYR_benzinadiko", "0", "Γ.35.AYTOMATO AKYΡΩΤΙΚΟ ΑΠΟΔ.(BENZINAΔIKO) =1 ΟΧΙ=0"))  'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
296     f_AKYR_benzPAR = FINDPARAMETROI(1, "PAR1", "f_AKYR_benzPAR", "T", "Γ.36.ΠΑΡΑΣΤΑΤΙΚΟ ΠΟΥ ΒΓΑΖΕΙ AYTOMATO AKYΡΩΤΙΚΟ ΑΠΟΔ.(BENZINAΔIKO)")
298     f_EID_AKYR_APOD = FINDPARAMETROI(1, "PAR1", "f_EID_AKYR_APOD", "x", "Γ.35Α.AKYΡΩΤΙΚΟ ΠΑΡΑΣΤΑΤΙΚΟ ΓΙΑ ΑΠΟΔ.ΕΙΣΡΟΩΝ") 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0")
 
        ' F_PAR_AYTOK = "  'L','T'"
 
300     F_PAR_AYTOK = FINDPARAMETROI(1, "PAR1", "F_PAR_AYTOK", "`L`,`T`", "Παραστατικά που κόβονται επί αυτοκινήτου") 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0")
302     F_PAR_AYTOK = Replace(F_PAR_AYTOK, "`", "'")
 
        'f_MYDATAcont
304     f_MYDATAcont = Val(FINDPARAMETROI(1, "PAR1", "f_MYDATAcont", "0", "Γ.38.Συνεχομενη (online) αποστολή Mydata")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
        'f_checkMYDATA
306     f_checkMYDATA = Val(FINDPARAMETROI(1, "PAR1", "f_checkMYDATA", "0", "Γ.39.Επιβεβαίωση πριν αποστολη(Σε online αποστολή Mydata=1"))
 
        'Dim F_EKPT_PARAGG As Integer ' Η ΕΚΠΤΩΣΗ ΠΑΡΑΓΓΕΛΙΑΣ ΑΝ ΘΑ ΜΕΤΑΦΕΡΕΤΑΙ ΣΤΟ ΜΕΤΑΣΧ.ΠΑΡΑΣΤΑΤΙΚΟ Π.Χ. ΤΙΜΟΛΟΓΙΟ
308     F_EKPT_PARAGG = Val(FINDPARAMETROI(1, "PAR1", "F_EKPT_PARAGG", "1", "Γ.36.H ΕΚΠΤΩΣΗ TOY METAΣXΗMATIZOMENOY ΠAΡAΣT.ΜΕΤΑΦΕΡΕΤΑΙ ΣΤΟ NEO=1 ΟΧΙ=0")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
 
        '&H00FF8080&

310     F_GRIDCOLOR = Val(FINDPARAMETROI(1, "PAR1", "f_GRIDCOLOR", "&H00FF8080&", "ΧΡΩΜΑ GRID ΕΙΔΩΝ ΑΡΧ=&H00FF8080& ΑΛΛΑ &H0080C0FF&,&H0080FF80&,&H00C0FFC0&")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
312     Grid1.BackColor = F_GRIDCOLOR
 
314     F_CHANGE_TELIKO_POSO = Val(FINDPARAMETROI(1, "PAR1", "F_CHANGE_TELIKO_POSO", "0", "Προσαρμόζεται το ποσό του μετασχ/νου παρασ/κου με την παραγγελία=1 ΟΧΙ=0")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
 
316     f_2TIMOL_1PAGE = Val(FINDPARAMETROI(1, "PAR1", "f_2TIMOL_1PAGE", "0", "Γ.35.ΠΡΩΤΟΤΥΠΟ+ΑΝΤΙΓΡΑΦΟ ΣΤΗΝ ΙΔΙΑ ΣΕΛΙΔΑ=1 ΟΧΙ=0")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
 
318     f_IS_PYROSB = Val(FINDPARAMETROI(1, "PAR1", "f_IS_PYROSB", "0", "Γ.43.ΠΑΡΑΚΟΛΟΥΘΩ ΠΥΡΟΣΒΕΣΤΗΡΕΣ=1 ΟΧΙ=0")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
 
        F_PLATOSGRIDEIDON = FINDPARAMETROI(3, "PAR1", "F_PLATOSGRIDEIDON", "1000;5000;1000;1000;1000;1000", "ΠΛΑΤΗ ΣΤΗΛΩΝ ΕΠΙΛΟΓΗΣ ΕΙΔΩΝ ΚΟΔ;ΟΝΟΜΑ;ΥΠΟΛ;.. ")
 
320     F_EXCEL_IMPORT = Val(FINDPARAMETROI(3, "PAR1", "F_EXCEL_IMPORT", "110", "Γ.46.Excel Import=100 Excel+τιμές=110 Excel+τιμές+Αυτόματο ανοιγμα κωδικου=111")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
 
        'f_METASX_KAI_HMER
 
 
 
 F_meIFESTO = Val(FINDPARAMETROI(3, "PAR1", "F_meIFESTO", "1", "1=ME ΗΦΑΙΣΤΟ 2=ΦΟΡΟΛ.ΧΩΡΙΣ ΗΦΑΙΣΤΟ(BAZEI DEH+AFM)"))
 
322     f_METASX_KAI_HMER = Val(FINDPARAMETROI(3, "PAR1", "f_METASX_KAI_HMER", "1", "Γ.48.Στα σχετ.δελτ.να εχει και αριθμο+ημερομ=1 μονο αριθμό=0"))  'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
 
324     F_PosostaOTA = Val(FINDPARAMETROI(3, "PAR1", "F_PosostaOTA", "0", "Γ.49.ΠΟΣΟΣΤΌ ΟΤΑ (ΛΑΤΟΜΕΙΑ 4%) ")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
 
326     f_MONO_PELATES_KATHGORIAS = Val(FINDPARAMETROI(3, "PAR1", "f_MONO_PELATES_KATHGORIAS", "0", "Γ.45.Tο παρ/τικό πώλησης μόνο από κατηγ.πελατών=1 ΟΧΙ=0"))  'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
        '  f_PEL_AYTOKINHTO = Val(FINDPARAMETROI(1, "PAR1", "f_PEL_AYTOKINHTO", "0", "Γ.39.ΣΤΟΝ ΑΡ.ΑΥΤΟΚΙΝΗΤΟΥ ΒΑΖΕΙ ΤΟΝ ΑΡ.ΑΥΤΟΚ. ΠΕΛΑΤΗ (CH1)")) 'αγορ=1 0=ΤΙΜΗ ΧΟΝΔΡΙΚΗΣ
 
328     f_auto_kostologhsh = Val(FINDPARAMETROI(1, "PAR1", "f_auto_kostologhsh", "0", "Γ.44.ΑΥΤΟΜΑΤΗ ΚΟΣΤΟΛΟΓΗΣΗ=1 ΟΧΙ=0")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
 
330     F_PDFNAME = FINDPARAMETROI(3, "PAR1", "F_PDFNAME", "PDF", "Ο PDF PRINTER NA ΠΕΡΙΕΧΕΙ ΤΟ ΚΕΙΜΕΝΟ Π.Χ. PDF ")
 
332     f_CAPTION_METASX_DA = FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_DA", "Μετασχημ.σε ΔΑ", "Τίτλος (Μετασχημ.σε ΔΑ)")
334     f_CAPTION_METASX_TIM = FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_TIM", "Μετασχημ. σε TIM", "Τίτλος (Μετασχημ.σε ΤΙΜ)")
336     f_CAPTION_METASX_LIAN = FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_LIAN", "Μετασχημ. σε Αποδ.Λιαν", "Τίτλος (Μετασχημ.σε ΛΙΑΝ)")
  
338     f_CAPTION_MTF11 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_MTF11", "Αριθ.Τεμαχίων", "Τίτλος (Αριθ.Τεμαχίων)")
340     f_CAPTION_MTF12 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_MTF12", "Είδος Συσκευασίας", "Τίτλος (Είδος Συσκευασίας)")
342     f_CAPTION_MTF13 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_MTF13", "Είδος Εμπορεύματος", "Τίτλος (Είδος Εμπορεύματος)")
344     f_CAPTION_MTF14 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_MTF14", "Μικτό Βάρος", "Τίτλος (Μικτό Βάρος)")
346     f_CAPTION_MTF21 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_MTF21", "Περιγραφή", "Τίτλος (Περιγραφή)")
  
  
  f_CAPTION_label24 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_label24", "Diesel Θερμ", "ΒΥΤΙΟ-ΕΛΑΙΟΤΡ 6ο ΤΕΧΤΒΟΧ")
  
  f_CAPTION_label25 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_label25", "AME EGTL", "ΒΥΤΙΟ-ΕΛΑΙΟΤΡ 7ο ΤΕΧΤΒΟΧ")
  
  
348     f_CAPTION_labf0 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_labf0", "μον", "Τίτλος (μον) μεταφορικές")
350     f_CAPTION_labf1 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_labf1", "ποσότητα", "Τίτλος (ποσότητα) μεταφορικές")
352     f_CAPTION_labf2 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_labf2", "αριθ.παρ", "Τίτλος (αριθ.παρ) μεταφορικές")
354     f_CAPTION_labf3 = FINDPARAMETROI(3, "PAR1", "f_CAPTION_labf3", "Περιγραφή", "Τίτλος (Περιγραφή) μεταφορικές")
  
356     MTF11(0).Caption = f_CAPTION_MTF11
358     MTF12(0).Caption = f_CAPTION_MTF12 ' As String
360     MTF13(0).Caption = f_CAPTION_MTF13 'As String
362     MTF14(0).Caption = f_CAPTION_MTF14  'As String
364     MTF21(0).Caption = f_CAPTION_MTF21  'As String
        Label24.Caption = f_CAPTION_label24
        Label25.Caption = f_CAPTION_label25

366     MTF11(1).Caption = f_CAPTION_MTF11
368     MTF12(1).Caption = f_CAPTION_MTF12 ' As String
370     MTF13(1).Caption = f_CAPTION_MTF13 'As String
372     MTF14(1).Caption = f_CAPTION_MTF14  'As String
374     MTF21(1).Caption = f_CAPTION_MTF21  'As String

        '
        ' Dim f_CAPTION_MTF21 As String
        '
  
376     MDIForm1.METASX_DA.Caption = f_CAPTION_METASX_DA
378     MDIForm1.METASX_TIM.Caption = f_CAPTION_METASX_TIM
380     MDIForm1.metasx_lian.Caption = f_CAPTION_METASX_LIAN

        f_parking_auto = Val(FINDPARAMETROI(3, "PAR1", "f_parking_auto", "0", "ειναι parking με αυτόματο ψηφιακο πελατολόγιο=1 /οχι=0"))
        
        F_ANAPODO = FINDPARAMETROI(3, "PAR1", "F_ANAPODO", "ΩΨ", "Κωδικοί παρ/κων με ανάποδο ΦΟΡΤΩΣΗ/ΠΡΟΟΡΙΣΜΟ ;")
        
  
382     'f_elga_typoi_krat = FINDPARAMETROI(3, "PAR1", "f_elga_typoi_krat", "1;2;4;3;", "Τύποι κρατήσεων ΕΛΓΑ για MyData π.χ.  1;2;4;3;")
         f_elga_typoi_krat = FINDPARAMETROI(3, "PAR1", "f_elga_typoi_krat", "1;2;", "Τύποι κρατήσεων ΕΛΓΑ για MyData π.χ.  1;2;4;3;")
384     f_TIMES_METASX_DA = Val(FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_DA", "0", "TO METAΣΧ.(DA) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))
386     f_TIMES_METASX_TIM = Val(FINDPARAMETROI(3, "PAR1", "f_TIMES_METASX_TIM", "0", "TO METAΣΧ.(TIM) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))
388     f_TIMES_METASX_LIAN = Val(FINDPARAMETROI(3, "PAR1", "f_TIMES_METASX_LIAN", "0", "TO METAΣΧ.(LIAN) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))
        f_hotel = Val(FINDPARAMETROI(3, "PAR1", "F_HOTEL", "0", "1=HOTEL ME AYTOMATO FORO 0=OXI"))
 
        'Dim f_TIMES_METASX_TIM As String  '  = FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_TIM", "Iaoao?ciaoeoiuo oa TIM", "AoeeYoa iaoao?ciaoeoiiy oa TIM")
        'Dim f_TIMES_METASX_LIAN As String
 
390     f_LabAFM = FINDPARAMETROI(1, "PELAT2", "f_LabAFM", "ΑΦΜ /κειμ1", "Tίτλος ετικέτας ΑΦΜ")
392     label11(0).Caption = f_LabAFM
 
394     Label18.Caption = "Πληρ"
396     Label9.Caption = "Σχόλια"
398     Label5.Caption = "Αρ.Αυτοκ"
400     Label21.Caption = "Προορισμός"
402     Label20.Caption = "Τόπ.Φόρτωσης"
404     Label19.Caption = "Σκοπός"
 
406     F_KOLA_MIKTA_LABELS = Val(FINDPARAMETROI(5, "PAR1", "F_KOLA_MIKTA_LABELS", "1", "Γ.37G.Ετικέτες Προελ/Μικτ= 1  Μηκ/Πλατ για PACKING LIST=2")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
  
408     If F_KOLA_MIKTA_LABELS = 0 Then
410         Label18.Caption = "Πληρ"
412         Label9.Caption = "Σχόλια"
414         Label5.Caption = "Αρ.Αυτοκ"
416         Label21.Caption = "Προορισμός"
418         Label20.Caption = "Τόπ.Φόρτωσης"
420         Label19.Caption = "Σκοπός"
  
        End If
  
422     If f_is_benzinadiko = 1 Then
424         Label9.Caption = "Αρ.Απόδ-Σχόλια=>"
426         ar_apod.Visible = True
            'parat.Left = 9000 'parat.Left + ar_apod.Width + 1500
428         ar_apod.Top = ar_apod.Top + 200
            ' Me.Refresh
     
        End If
        
        '        F_EIDFOROSKAT = FINDPARAMETROI(6, "PAR1", "F_EIDFOROSKAT", "0", "Δουλεύω ειδικό φόρο κατανάλωσης=1 Οχι=0")
        
430     F_EIDFOROSKAT = FINDPARAMETROI(6, "PAR1", "F_EIDFOROSKAT", "0", "Δουλεύω ειδικό φόρο κατανάλωσης=1 Οχι=0")
        
        'Dim f_psifiaTIMMON    'posa psifia tha exei h kathe seira
432     f_psifiaTIMMON = FINDPARAMETROI(8, "PAR1", "f_psifiaTIMMON", "3", "ΑΡΙΘΜΟΣ ΨΗΦΙΩΝ ΤΙΜ.ΜΟΝ")
        
434     F_PLATOS_FORMAS = FINDPARAMETROI(1, "PAR1", "F_PLATOS_FORMAS", "13860", "Πλάτος φόρμας(13860) όσο μεγαλωνει τοσο στενευει η φόρμα")
        
        ' F_REMOTE = FINDPARAMETROI(1, "PAR1", "F_REMOTE", "0", "K.1.ΦΟΡΗΤΟ ΤΕΡΜΑΤΙΚΟ REMOTE=1 / 0")
        
436     F_YCOS_FORMAS = FINDPARAMETROI(1, "PAR1", "F_YCOS_FORMAS", "10830", "Υψος φόρμας")

438     F_SET_EIDON = 0
440     F_METASX_PAR = 0  ' As Integer  ' ΕΙΝΑΙ ΜΕΤΑΣΧΗΜΑΤΙΣΜΟΣ ΠΑΡΑΣΤΑΤΙΚΟΥ = 1
        Dim k, L As Integer, X As String, DB

        Dim R As New ADODB.Recordset

442     Set Prospective = TDBGrid1.Styles.Add("Prospective")
444     Prospective.Font.Italic = True
446     Prospective.Font.Bold = True
448     Prospective.ForeColor = vbBlue
        
450     Set Distributors = TDBGrid1.Styles.Add("Distributors")
452     Distributors.BackColor = vbRed
454     Distributors.ForeColor = vbWhite


Me.Caption = Me.Caption + "--t1--" + str(GetCurrentTime() - TT)

        'Const f_k = 1
        'Const f_o = 2
        'Const f_p = 3
        'Const F_T = 4
        'Const f_e = 5
        'Dim f_f ' = 11  '6  ''ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
        'Const f_a = 7
        'Const F_M = 8
        'Const F_mefpa = 9
        'Const F_PROEL = 10
        'Const F_mikta = 12
        'Dim F_EKPT2  '  = 6  '11   'ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
        
456     f_k = 1: f_o = 2: f_p = 3: F_T = 4: f_e = 5: F_EKPT2 = 6: f_a = 7: F_M = 8: F_mefpa = 9: F_PROEL = 10: f_f = 11: F_mikta = 12: F_mikta_KILA = 13: F_COL_KOLA = 14
458     f_colEFK = 15
        
460     f_f = 6 ''ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
462     F_EKPT2 = 11  'ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI  069   458
464     F_2EKPTOSEIS = Val(FINDPARAMETROI(7, "PAR1", "F_2EKPTOSEIS", "0", "Γ.32.Εχω 2η έκπτωση ανά είδος =1 οχι=0")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
466     If F_2EKPTOSEIS = 1 Then  ' And f_FORM_EKTY > 0 Then ' Or F_2EKPTOSEIS = 3 Then
            ' f_f = 11       '  6  ''ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
            ' F_EKPT2 = 6    '  11   'ALLAZEI STO FORM_LOAD AN EXO 2PLH EKPTOSI
            
            'ian 2025
468         f_a = 11
470         f_f = 7
472         F_EKPT2 = 6
        End If

474     F_KOLA_MIKTA = Val(FINDPARAMETROI(5, "PAR1", "F_KOLA_MIKTA", "0", "Γ.37.Εχω ΚΟΛΑ,ΜΙΚΤΑ,ΠΡΟΕΛΕΥΣΗ =1 όχι=0")) 'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))
        
        '        If F_KOLA_MIKTA = 1 Then ' ΚΟΛΑ ΜΙΚΤΑ
        '            f_k = 1: f_o = 2: F_PROEL = 3: F_COL_KOLA = 4: F_mikta_KILA = 5: f_p = 6: F_T = 7: f_e = 8: f_f = 9: f_a = 10: F_M = 11: F_mefpa = 12: F_EKPT2 = 13: F_mikta = 14
        '            Grid1.TextMatrix(0, F_COL_KOLA) = "ΚΟΛΑ"
        '            Grid1.TextMatrix(0, F_mikta_KILA) = "ΜΙΚΤΑ ΚG"
        '            Grid1.TextMatrix(0, f_p) = "ΚΑΘΑΡΑ KG"
        '
        '        End If

        '     On Error GoTo MHNYMA

        ' R.Open "SELECT top 1 *FROM TIM", Gdb, adOpenDynamic, adLockOptimistic

        ' R.Close

        ' par1.SHOW

        'R.Open "select * FROM PARAMETROI", Gdb, adOpenDynamic, adLockBatchOptimistic
        '
        'For K = 1 To 200
        '   F_AR_parametroi(K, 0) = R("FORMA")
        '   F_AR_parametroi(K, 1) = R("VAR")
        '   F_AR_parametroi(K, 2) = R("TIMH")
        '   R.MoveNext
        '   If R.EOF() Then
        '     Exit For
        '   End If
        'Next
        '
        'R.Close

476     If gXEIRISTHS < 2 Then
478         TDBGrid1.Visible = False
        End If

480     F_UPPER = Val(FINDPARAMETROI(1, "APOT2", "F_UPPER", "1", "ΜΕΤΑΤΡΟΠΗ ΠΕΡΙΓΡΑΦΩΝ ΣΕ ΚΕΦΑΛΑΙΑ ΕΛΛΗΝΙΚΑ=1/0"))
482     F_EISITHRIA = Val(FINDPARAMETROI(5, "PAR1", "F_EISITHRIA", "0", "Χρησιμοποιώ εισιτήρια=1 Οχι=0"))
     
        'Dim F_LIANIKES_TIMES As Integer '1=ΔΟΥΛΕΥΩ ΛΙΑΝΙΚΕΣ ΤΙΜΕΣ ΣΤΑ ΠΑΡΑΣ/ΚΑ ΠΟΥ ΕΧΟΥΝ ΤΗΝ ΕΝΔΕΙΞΗ ΛΙΑΝΙΚΗ
484     F_LIANIKES_TIMES = Val(FINDPARAMETROI(2, "PAR1", "F_LIANIKES_TIMES", "0", "1=Δουλεύω Λιανικές τιμές στα παραστ/κά ΠΟΥ ΕΧΟΥΝ ΤΗΝ ΕΝΔΕΙΞΗ ΛΙΑΝΙΚΗ"))
     
        'fMonimo_MIDENIKO_FPA
486     fMonimo_MIDENIKO_FPA = Val(FINDPARAMETROI(2, "PAR1", "fMonimo_MIDENIKO_FPA", "0", "1=Δουλεύω μονιμα ΧΩΡΙΣ ΦΠΑ Π.Χ. ΙΑΤΡΟΙ "))
     
488     F_PAROCHOS = Val(FINDPARAMETROI(1, "PAR7MYDATA", "f_parochos", "0", "0=MyData 1=Παροχος"))
         
         MDIForm1.StatusBar1.Panels(6).Text = IIf(F_PAROCHOS = 1, "ΠΑΡΟΧΟΣ", "myData")
490     f_pelMERC = FINDPARAMETROI(5, "PAR1", "F_PELMERC", "..", "Χρήστης MERCURY για custom")
 
492     F_PDFFOLDER = Trim(FINDPARAMETROI(4, "PAR1", "F_PDFFOLDER", "\\user-5rwk428ge6\C\PDF", "Φάκελος με τα PDF"))
        '\\user-5rwk428ge6\C\PDF\

494     f_site = Val(FINDPARAMETROI(1, "MDIFORM1", "F_SITE", "0", "ΚΩΔΙΚΟΣ SITE 1=VM 2=TOYS"))

496     If Left(f_pelMERC, 3) = "SOK" Then
            '   Text14.Visible = True
498         MSFlexGrid1.Visible = True
500         MSFlexGrid1.row = 0

502         MSFlexGrid1.Col = 0
504         MSFlexGrid1.Text = "ΧΡΩΜΑ"

506         MSFlexGrid1.Col = 1
508         MSFlexGrid1.Text = "BARCODE"

510         MSFlexGrid1.Col = 2
512         MSFlexGrid1.Text = "ΥΠΟΛΟΙΠΟ"

        End If

514     F_akyr_Delete = Val(FINDPARAMETROI(3, "PAR1", "F_akyr_Delete", "1", "Γ.31 Σβήνω το παραστατικό που ακυρώνεται=1"))   ' δεν ακυρωνω το παραστατικό που ακυρώνεται

516     F_add_barc = Val(FINDPARAMETROI(2, "PAR1", "add_barc", "1", "Γ.16 Barcode έιναι 1.Κωδ 2.Βοηθ.Κλειδί"))

518     PAR1.KeyPreview = True

520     f_PALIAXRONIA = Trim(FINDPARAMETROI(3, "PAR1", "F_PALIAXRONIA", "2005", "DSN παλιάς χρονιάς"))
522     F_EXO_XROMATA = Val(FINDPARAMETROI(5, "PAR1", "F_EXO_XROMATA", "0", "Χρησιμοποιώ Μεγέθη-Χρώματα=1"))   'posa psifia tha exei h kathe seira

524     F_POLHTES = Val(FINDPARAMETROI(7, "PAR1", "F_POLHTES", "0", " 1=Πωλητής πελάτη->Τιμολόγιο  0=Χωρίς Πωλητές"))

526     f_Fkey = Trim(FINDPARAMETROI(8, "PAR1", "F_FKEY", "112", "Το πλήκτρο που χρησιμοποιώ για καταχώρηση Παρ/κού"))

528     F_EIKONA = Val(FINDPARAMETROI(1, "APOT2", "F_EIKONA", "1", "1=EIKONA=>ΚΩΔΙΚΟΣ 2=>ΜΕ ΒΟΗΘ.ΚΛΕΙΔΙ 11=ΚVDIKOS.JPG 12=BOH.JPG 0=ΟΧΙ ΕΙΚΟΝΕΣ"))

530     F_PAR_MONO_ME_YPOLOIPO = Val(FINDPARAMETROI(1, "PAR1", "F_PAR_MONO_ME_YPOLOIPO", "0", "Παραγγελία πελάτη μόνο αν υπάρχει υπόλοιπο"))

532     F_STOPDELTIA = Val(FINDPARAMETROI(3, "PAR1", "F_STOPDELTIA", "0", "Αποτρέπει την τιμολόγηση αν εκκρεμούν δελτία=1"))   'posa psifia tha exei h kathe seira

534     F_2CHFIA_ZYGIZOMENA = FINDPARAMETROI(5, "PAR1", "F_2CHFIA_ZYGIZOMENA", "  ", "Πρώτα ψηφία Βοηθ.Κωδ.των ζυγιζόμενων")   '2 ΠΡΩΤΑ ΨΗΦΙΑ ΒΟΗΘ.ΚΩΔ.ΤΩΝ ΖΥΓΙΖΟΜΕΝΩ

536     f_LEKTIKO_ARPARAGG = FINDPARAMETROI(1, "PAR1", "f_LEKTIKO_ARPARAGG", "Αρ.Παραγγ", "Λεκτικό στα σχόλια για αρ.Παραγγελίας")  '2 ΠΡΩΤΑ ΨΗΦΙΑ ΒΟΗΘ.ΚΩΔ.ΤΩΝ ΖΥΓΙΖΟΜΕΝΩ
        '

        'ΔΕΝ ΑΦΗΝΕΙ ΙΔΙΑ ΣΕΙΡΑ AN (κωδ<>παλ.κωδ ή ποσ=0) =1    (KOD<>PALIO ή (ποσ=0 & τιμη=0) )=2
538     f_AFHNO_IDIA_SEIRA = Val(FINDPARAMETROI(5, "PAR1", "F_AFHNO_IDIA_SEIRA", "1", "Δεν αφήνει ιδια σειρά αν (κωδ<>παλ.κωδ ή ποσ=0) =1    (KOD<>PALIO ή (ποσ=0 & τιμη=0) )=2"))

540     F_PERIOR_POLHTES = Val(FINDPARAMETROI(7, "PAR1", "F_PERIOR_POLHTES", "0", "Χρησιμοποιώ περιορισμό των πωλητών ανά Παραστατικό=1 Οχι=0"))

542     F_DEBUG = Val(FINDPARAMETROI(5, "PAR1", "F_DEBUG", "0", "Εγκατάσταση με debugging=1 Οχι=0"))

544     F_PDFPARAGG = Val(FINDPARAMETROI(4, "PAR1", "F_PDFPARAGG", "0", "Aντιγράφει το PDF και στο C:\PDFPARAGG =1 Oχι=0"))
    
        'F_EPIBEB_APOT

546     F_EPIBEB_APOT = Val(FINDPARAMETROI(5, "PAR1", "F_EPIBEB_APOT", "0", "Επιβεβαίωση Αποθήκης στα τιμ.Αγοράς=1 Οχι=0"))

548     If Len(Trim(F_2CHFIA_ZYGIZOMENA)) > 1 Then
550         Grid1.ColWidth(7) = 10
552         Grid1.ColWidth(8) = 10

        End If

554     f_Desmia = Val(FINDPARAMETROI(1, "PAR1", "F_DESMIA", "0", "Οι παραγγελίες υπολογίζονται στο Υπόλοιπο=1 "))   ' χρησιμοποιω τις παραγγελίες για έλεγχο υπολοίπου  'βλεπε mdesmia

556     f_mono_apoParaggelia = Val(FINDPARAMETROI(1, "PAR1", "f_mono_apoParaggelia", "0", "Τιμολογώ μόνο αν υπάρχει παραγγελία=1 ")) '
558     f_SYMPYKNVSH = Val(FINDPARAMETROI(7, "PAR1", "f_SYMPYKNVSH", "0", "Συμπύκνωση στον ίδιο κωδικό "))  '
    
560     F_PDF = Val(FINDPARAMETROI(4, "PAR1", "f_PDF", "0", "Στέλνει σε PDF=1 Oχι=0 "))  '
    
562     F_METRHTA = Val(FINDPARAMETROI(7, "PAR1", "f_METRHTA", "0", "Zητά μετρητά μετά την καταχώρηση=1 Oχι=0 "))  '
    
564     F_PARAG_KART = Val(FINDPARAMETROI(1, "PAR1", "F_PARAG_KART", "0", "Φαίνεται η Παραγγελία στην καρτέλα=1 Oχι=0 ")) '

566     F_POLLES_SEIRES = Val(FINDPARAMETROI(4, "PAR1", "F_POLLES_SEIRES", "0", "Η περιγραφή στην εκτύπωση σε πολλές σειρές=1 ")) '

568     F_TRANS = Val(FINDPARAMETROI(3, "PAR1", "F_TRANS", "1", "Χρησιμοποιεί Transactions=1  "))  '
        ' F_TRANS = 1  ' PANTA NA ΔΟΥΛΕΥΕΙ ΤRΑΝSΑCTIONS
570     F_MONOMEYPOLOIPO = Val(FINDPARAMETROI(2, "PAR1", "F_MONOMEYPOLOIPO", "0", "Εμφάνιση μόνο των ειδών με υπόλοιπο=1"))   '
        '  F_EIDH_APOT = Val(FindParametroi(1,"PAR1", "F_EIDH_APOT", "0", "Γ. 9B. ΕΜΦΑΝΙΣΗ ΜΟΝΟ ΤΩΝ ΕΙΔΩΝ ΜΕ ΚΙΝΗΣΗ ΣΤΟ ΥΠΟΚΑΤΑΣΤΗΜΑ=1")) '
572     F_YPOL_GEN = Val(FINDPARAMETROI(7, "PAR1", "F_YPOL_GEN", "1", "Εμφάνιση όλων των ειδών=1 Μόνο με κίνηση=0"))   ' NA EMFANIZEI TA EIDH ASXETA AN EXOYN KINHΘΕΙ

574     f_KENTRA_ADYNATISMATOS = Val(FINDPARAMETROI(1, "PELAT2", "f_KENTRA_ADYNATISMATOS", "0", "Αν είναι κέντρο αδυνατίσματος=1"))

576     f_IS_PARTIDES = Val(FINDPARAMETROI(1, "APOT2", "f_IS_PARTIDES", "0", "Αν Χρησιμοποιώ παρτίδες=1"))

578     If f_KENTRA_ADYNATISMATOS = 1 Then
580         kentroAdyn.Visible = True
582         kentroAdyn.ListIndex = 0 ' Combo3.List(0)
        End If

        '  'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""

        '    'SkinFramework1.ApplyWindow Me.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        Dim DELT  As String

        Dim nDelt As Integer

        Dim sql   As String

        '    sql = "SELECT convert(CHAR(10),HME,3) AS [ΗΜΕΡ/ΝΙΑ],ATIM,SKOPOS,EPO" & " FROM TIM INNER JOIN PEL ON PEL.EIDOS=TIM.EIDOS AND PEL.KOD=TIM.KPE" & " WHERE ART IS NULL AND LEFT(ATIM,1) in ( select EIDOS from PARASTAT WHERE STADIO=3)  and SKOPOS LIKE '%ΩΛ%' AND" & " HME<'" + Format(DateAdd("D", -25, Now), "MM/DD/YYYY") + "' and HME>'" + Format(DateAdd("D", -31, Now), "MM/DD/YYYY") + "' ORDER BY HME"
584     sql = "SELECT  convert(CHAR(10),HME,3) AS [ΗΜΕΡ/ΝΙΑ],ATIM,SKOPOS,EPO" & " FROM TIM INNER JOIN PEL ON PEL.EIDOS=TIM.EIDOS AND PEL.KOD=TIM.KPE" & " WHERE ART IS NULL AND LEFT(ATIM,1) in ( select EIDOS from PARASTAT WHERE STADIO=3)   AND  month(getdate())-month(HME)<2 AND YEAR(HME)=YEAR(GETDATE() )  ORDER BY HME"  'TO BAFEIRESA=> 'and SKOPOS LIKE '%ΩΛ%'

586     R.Open sql, Gdb, adOpenStatic, adLockReadOnly

588     If R.EOF Then
            'OK
        Else
590         nDelt = 0
592         DELT = "Τα παρακάτω δελτία είναι ατιμολόγητα " + Chr(13)

594         Do While Not R.EOF

                ' If InStr(R("skopos"), "ΩΛ") > 0 Then
596             DELT = DELT + R("atim") + " " + R(0) + " " + R("skopos") + " " + R("epo") + Chr(13)
598             nDelt = nDelt + 1
                ' End If

600             R.MoveNext
            Loop

602         If nDelt > 0 And f_EkkremhDeltia = 1 Then
604             MsgBox DELT
            End If

        End If

606     R.Close
 Me.Caption = Me.Caption + "--t2b--" + str(GetCurrentTime() - TT)
608     If F_STOPDELTIA = 1 Then

610         sql = "SELECT convert(CHAR(10),HME,3) AS [ΗΜΕΡ/ΝΙΑ],ATIM,SKOPOS,EPO" & " FROM TIM INNER JOIN PEL ON PEL.EIDOS=TIM.EIDOS AND PEL.KOD=TIM.KPE" & " WHERE ART IS NULL AND LEFT(ATIM,1)='A' and SKOPOS LIKE '%ΩΛ%' AND" & " HME<'" + Format(DateAdd("D", -27, Now), "MM/DD/YYYY") + "' and HME>'" + Format(DateAdd("D", -31, Now), "MM/DD/YYYY") + "' ORDER BY HME"
612         R.Open sql, Gdb, adOpenStatic, adLockReadOnly

614         If R.EOF Then
                'ok
            Else

616             nDelt = 0
618             DELT = "Θα σταματήσει η τιμολόγηση επειδή " + Chr(13) + "Τα παρακάτω δελτία είναι ατιμολόγητα " + Chr(13)

620             Do While Not R.EOF

622                 If InStr(R("skopos"), "ΩΛ") > 0 Then
624                     DELT = DELT + R("atim") + " " + R(0) + " " + R("skopos") + " " + R("epo") + Chr(13)
626                     nDelt = nDelt + 1
                    End If

628                 R.MoveNext
                Loop

630             If nDelt > 0 Then
632                 MsgBox DELT, vbCritical
                End If

            End If

634         R.Close
636         F_STOPDELTIA = 99
        End If

        ' On Error GoTo MHNYMA

638     geidos_paras = 1
        ' On Error GoTo Form_Load_Err
        'tropos plhromhs
        'R.Close
640     R.Open "SELECT *FROM PINAKES WHERE TYPOS=12 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
642     k = 0

644     Do While Not R.EOF

646         If R("typos") = 12 Then
                ' Combo4.AddItem str(R("AYJON")) + "." + R("PERIGRAFH")
648             Combo4.AddItem R("PERIGRAFH")
650             F_LOGPLHR(k) = Format(R("TIMH"), "0000")
652             Combo4.ItemData(k) = R!ID
654             k = k + 1
            End If

656         R.MoveNext
        Loop
658     If k < 7 Then
660         If GGET_NVALUE("select count(*) from PINAKES WHERE TYPOS=12 AND N1=6") = 0 Then
662             Dim NNW As Integer: NNW = GGET_NVALUE("select MAX(AYJON) from PINAKES WHERE TYPOS=12 ")
664             Gdb.Execute "INSERT INTO PINAKES (TYPOS,PERIGRAFH,TIMH,N1,AYJON) VALUES(12,'WEB BANKING',1,6," + str(NNW + 1) + ")"
            End If
666         If GGET_NVALUE("select count(*) from PINAKES WHERE TYPOS=12 AND N1=7") = 0 Then
668             Dim NNW2 As Integer: NNW2 = GGET_NVALUE("select MAX(AYJON) from PINAKES WHERE TYPOS=12 ")
670             Gdb.Execute "INSERT INTO PINAKES (TYPOS,PERIGRAFH,TIMH,N1,AYJON) VALUES(12,'POS/e-POS',1,7," + str(NNW2 + 1) + ")"
            End If
        End If

672     R.Close
674     k = 0

676     F_DIOR = 0

        'ΦΟΡΤΩΜΑ ΑΠΟΘΗΚΩΝ
'678     R.Open "SELECT *FROM PINAKES WHERE TYPOS=4 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
'
'680     Do While Not R.EOF
'682         APOTH1.AddItem LTrim(str(R("AYJON"))) + "." + R("PERIGRAFH")
'684         APOTH2.AddItem LTrim(str(R("AYJON"))) + "." + R("PERIGRAFH")
'686         R.MoveNext
'        Loop
'
'688     R.Close


For k = 0 To UBound(f_Arr_PIN) '(k_ARR, F_ARR_TYPOS) = R_arr!TYPOS
  If f_Arr_PIN(k, F_ARR_TYPOS) = 4 Then
        APOTH1.AddItem LTrim(str(f_Arr_PIN(k, F_ARR_AYJON))) + "." + f_Arr_PIN(k, F_ARR_PERIGRAFH)
        APOTH2.AddItem LTrim(str(f_Arr_PIN(k, F_ARR_AYJON))) + "." + f_Arr_PIN(k, F_ARR_PERIGRAFH)
  End If
Next





690     APOTH1.ListIndex = 0

692     If Len(Dir("C:\LAGEURO\NOPONTOI.TXT", vbNormal)) > 0 Then
694         Command2.Visible = True
        Else
696         Command2.Visible = False
        End If

698     f_use_Sysk = Val(FINDPARAMETROI(7, "PAR1", "F_USE_SYSK", "0", "Γ. 4. Χρησιμοποιώ συσκευασία=1"))   ' πολαπλασιάζει την ποσότητα επί την συσκευασία

700     F_falcon_Polla_tim = Val(FINDPARAMETROI(5, "PAR1", "F_FALCON_POLLA_TIM", "0", "Γ. 5. Πολλαπλά τιμολόγια από PDA=1"))   'posa psifia tha exei h kathe seira

702     f_apeyueias = Trim(FINDPARAMETROI(7, "PAR1", "f_APEYUEIAS", "0", "Α. 3. Κανονικα=0/Απ ευθείας στο παρ/κό με κωδ.πελάτη.."))   '

704     f_1choice_agores = Val(FINDPARAMETROI(3, "PAR1", "f_1choice_agores", "2", "Α. 4. Πρώτη επιλογή αγορών"))   'posa psifia tha exei h kathe seira
706     f_1choice_polhseis = Val(FINDPARAMETROI(3, "PAR1", "f_1choice_polhseis", "2", "Α. 5. Πρώτη επιλογή πωλήσεων"))   'posa psifia tha exei h kathe seira
708     f_AutoChangeLine = Val(FINDPARAMETROI(7, "PAR1", "f_AutoChangeLine", "0", "Α. 6 Αλλάζει αυτόματα σειρά=1")) 'def="1"  "1=μπορω να αλλάξω έκπτωση γραμμής 0=δεν μπορώ

710     f_CanChangeEkptosi = Val(FINDPARAMETROI(7, "PAR1", "f_CanChangeEkptosi", "1", "Α. 7. Αλλάζω έκπτωση γραμμής=1 "))   'def="1"  "1=μπορω να αλλάξω έκπτωση γραμμής 0=δεν μπορώ
712     f_CanChangeFpa = Val(FINDPARAMETROI(7, "PAR1", "f_CanChangeFpa", "1", "Α. 8 Αλλάξω ΦΠΑ γραμμής=1"))   'def="1"  "1=μπορω να αλλάξω  γραμμής 0=δεν μπορώ

714     f_psifiaAjias = Val(FINDPARAMETROI(8, "PAR1", "f_psifiaAjias", "2", "Γ. 6. Δεκαδικά Ψηφία Αξίας Σειρών τιμολογιου"))   'posa psifia tha exei h kathe seira

716     f_PARAGGMARKUP = Val(FINDPARAMETROI(1, "PAR1", "f_PARAGGMARKUP", "0", "Γ. 7. Στην Παραγγελία Πελ.0=Χονδ 1=Markup"))

718     f_MinSeiresTimologioy = Val(FINDPARAMETROI(4, "PAR1", "F_MINSEIRESTIMOLOGIOY", "25", "Β. 1 Σειρές τιμ.Πώλησης"))
  
720     F_TYPONO_KENES = Val(FINDPARAMETROI(4, "PAR1", "F_TYPONO_KENES", "1", "Β.3.1 Θα τυπώνω κενές σειρές στο Τιμολόγιο=1 Οχι=0"))

        'DIM F_TYPONO_KENES AS Integer 'ΑΝ ΘΑ ΤΥΠΩΝΩ ΚΕΝΕΣ ΣΕΙΡΕΣ ΣΤΟ ΤΙΜΟΛΟΓΙΟ

722     f_Timol2_SeiresTimologioy = Val(FINDPARAMETROI(4, "PAR1", "f_TIMOL2_SEIRESTIMOLOGIOY", str(f_MinSeiresTimologioy), "Β. 4 Σειρές τιμ.Πώλησης φορμας TIMOL2"))

724     f_MaxSeiresTimologioy = Val(FINDPARAMETROI(4, "PAR1", "F_MAXSEIRESTIMOLOGIOY", "80", "Β. 2 Σειρές Τιμ.Αγοράς"))

726     If f_MaxSeiresTimologioy > 900 Then
728         f_MaxSeiresTimologioy = 900
        End If
    
730     f_SynSeiresTimologioy = Val(FINDPARAMETROI(4, "PAR1", "f_SynSeiresTimologioy", "25", "Β. 3 Σειρές Τιμ.Πωλ Κενές+Γραμμ"))
    
732     f_AmesosPelatisLian = Val(FINDPARAMETROI(3, "PAR1", "f_AmesosPelatisLian", "0", "Β. 5.Το όνομα πελατη αποθηκεύεται στην λιανική"))

734     If f_AmesosPelatisLian = 1 Then
            ' MDIForm1.Timer1.Enabled = False
736         Text2(3).Enabled = True
        End If
    
738     gCapitals = Val(FINDPARAMETROI(7, "PAR1", "gCapitals", "1", "Β. 6.Το κειμενο αναζήτησης να γίνεται σε κεφαλαία"))
    
740     Grid1.rows = f_MaxSeiresTimologioy
  
        '1180    f_checkPos = Val(FINDPARAMETROI(1, "PAR1", "F_CHECKPOS", "DELETE", " 1=ΕΛΕΓΧOΣ 2=ΑΠΑΓΟΡΕΥΣΗ 0=ΤΙΠΟΤΑ"))

742     f_checkYpol = Val(FINDPARAMETROI(7, "PAR1", "F_checkYpol", "0", "Α. 8.Υπόλοιπο είδους Ελεγχος=1/Απαγόρευση=2/Ελεύθερο=0"))

744     f_TIMOKAT = Val(FINDPARAMETROI(1, "PELAT2", "F_TIMOKAT", "0", "Γ. 8. 0=ΓΕΝ.ΕΚΠΤΩΣΗ ΠΕΛΑΤΗ  1=ΤΙΜΟΚΑΤΑΛΟΓΟΣ ΕΚΠΤ -1=ΤΙΜΟΚ.ΜΕ ΤΙΜΕΣ"))

746     f_psifia_atim = Val(FINDPARAMETROI(8, "PAR1", "F_PSIFIA_ATIM", "5", "Γ. 9. Ψηφία Αριθμού Παραστατικού(5)"))

748     If f_psifia_atim < 5 Then f_psifia_atim = 5

        'Dim f_barcode As Integer '1=diabazei apeyueias ta barcode
750     f_barcode = Val(FINDPARAMETROI(7, "PAR1", "F_BARCODE", "0", "Α. 9 Eίδη: 0=Απ ευθείας ο κωδικός=0 Το Barcode=1"))

752     f_LIANIKHCHECK = Val(FINDPARAMETROI(2, "PAR1", "f_LIANIKHCHECK", "1", "Α.10  1=Λιανική μαρκαρισμένη 0=Οχι"))

754     If f_LIANIKHCHECK = 0 Then lianikh.Value = vbUnchecked Else lianikh.Value = vbChecked

756     F928 = Val(FINDPARAMETROI(4, "PAR1", "F928", "0", "Α.11 Εκτύπωση φόρμας 1=928 0=437"))

758     F_File_Timologioy = FINDPARAMETROI(4, "PAR1", "F_File_Timologioy", "c:\print", "Α.12 Αρχείο εκτύπωσης παραστατικού")
760     F_File2_Timologioy = FINDPARAMETROI(4, "PAR1", "F_File2_Timologioy", "c:\print", "Α.12b Αρχείο εκτύπωσης παραστ.Υποκ")

762     F_Type_File_Timologioy = Val(FINDPARAMETROI(4, "PAR1", "F_Type_File_Timologioy", "1", "Α.13 1=εκτύπωση αρχείου παρ/κού 0=Οχι 2=Print 3=c:\typ.bat"))
764     F_Type2_File_Timologioy = Val(FINDPARAMETROI(4, "PAR1", "F_Type2_File_Timologioy", "1", "Α.13b 1=εκτύπ.αρχείου παρ/κού υποκ. 0=Οχι 2=Print 3=c:\typ.bat"))

766     f_F12 = Val(FINDPARAMETROI(8, "PAR1", "F_F12", "-1", "Α.14 F12:Τύπος Παραστατικού που ενημερώνεται"))
768     f_F1 = Val(FINDPARAMETROI(8, "PAR1", "F_F1", "-1", "Α.15 F1:Τύπος Παραστατικού που ενημερώνεται"))

770     f_BARCODE_STOTIM = Val(FINDPARAMETROI(4, "PAR1", "F_BARCODE_STOTIM", "0", "Α.16  1=Βάζει το ERG στο Τιμολόγιο 0=τον κωδικό(def)"))

772     F_PIN_PROELEYS = Val(FINDPARAMETROI(5, "PAR1", "F_PIN_PROELEYS", "0", "Α.17  1=Προελεύσεις από πίνακα(22)  2=ειδη erg 3=eid.ch4"))
774     F_8 = Val(FINDPARAMETROI(7, "PAR1", "F_8", "1", "Α.18  1=Υπολογίζει ποσοτητα από αξία 2=τιμη απο αξία 3=εκπτ.απο αξια 4=ΑΙΡΒΝΒ"))
    
776     F_LOCKARITMISI = Val(FINDPARAMETROI(3, "PAR1", "F_LOCKARITMISI", "0", "K.3  Κλειδώνει την αρίθμηση=1 όχι=0"))
    
778     If F_PIN_PROELEYS = 1 Then
780         R.Open "SELECT *FROM PINAKES WHERE TYPOS=22 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
            Dim nProel
782         nProel = 1
784         Do While Not R.EOF
786             F_PINAKAS_PROEL(nProel) = R("PERIGRAFH")
788             R.MoveNext
790             nProel = nProel + 1
            Loop
792         R.Close
        End If
        
        Dim PROOR    As String

        Dim APOSTOLH As String

        'PROOR = Val(FindParametroi(1,"PAR1", "PROOR", "1", "ΕΔΡΑ ΠΕΛΑΤΗ,Γ.ΖΕΡΒΟΥ 35"))
        'APOSTOLH = Val(FindParametroi(1,"PAR1", "APOSTOLH", "1", "ΕΔΡΑ ΜΑΣ,Κ.ΠΑΛΑΙΟΛΟΓΟΥ 21"))

'        'ΦΟΡΤΩΜΑ ΣΚΟΠΩΝ
'794     R.Open "SELECT *FROM PINAKES WHERE TYPOS=7 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
'
'796     Do While Not R.EOF
'798         SKOPOS.AddItem Right(" " + Format(R!AYJON, "##"), 2) + "." + R("PERIGRAFH")
'800         R.MoveNext
'        Loop
'        R.Close
        
For k = 0 To UBound(f_Arr_PIN) '(k_ARR, F_ARR_TYPOS) = R_arr!TYPOS
  If f_Arr_PIN(k, F_ARR_TYPOS) = 7 Then
        SKOPOS.AddItem Right(" " + Format(f_Arr_PIN(k, F_ARR_AYJON), "##"), 2) + "." + f_Arr_PIN(k, F_ARR_PERIGRAFH)
  End If
Next
        
        
        
        
        
        
        
        
802     If SKOPOS.ListCount < 20 Or InStr(SKOPOS.List(5), "---") = 0 Then

804         Gdb.Execute "DELETE FROM PINAKES WHERE TYPOS=7"

806         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'Πώληση',7)"
808         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (2,'Πώληση για Λογαριασμό Τρίτων',7)"
810         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (3,'Δειγματισμός',7)"
812         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (4,'Έκθεση',7)"
814         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (5,'Επιστροφή',7)"
816         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (6,'ΚΑΤΑΡΓΗΘΗΚΕ---Φύλαξη',7)"
818         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (7,'Επεξεργασία Συναρμολόγηση',7)"
820         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (8,'Ενδοδιακίνηση.Μεταξύ Εγκαταστάσεων',7)"
822         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (9,'Αγορά',7)"
824         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (10,'Εφοδιασμός πλοίων',7)"
826         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (11,'Δωρεάν διάθεση',7)"
828         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (12,'Εγγύηση',7)"
830         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (13,'Χρησιδανεισμός',7)"
832         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (14,'Αποθήκευση σε Τρίτους',7)"
834         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (15,'ΚΑΤΑΡΓΗΘΗΚΕ---Επιστροφή από Φύλαξη',7)"

836         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (16,'ΚΑΤΑΡΓΗΘΗΚΕ---Ανακύκλωση',7)"
838         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (17,'ΚΑΤΑΡΓΗΘΗΚΕ---Καταστροφή άχρηστου υλικού',7)"
840         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (18,'ΚΑΤΑΡΓΗΘΗΚΕ---Διακίνηση Παγίων ',7)"
842         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (19,'Λοιπές Διακινήσεις',7)"
             Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (20,'Μεταφορές - Ταχυμεταφορές',7)"

            '1Πώληση 2 Πώληση για Λογαριασμό Τρίτων 3 Δειγματισμός 4 Έκθεση 5 Επιστροφή 6 Φύλαξη
            '7 Επεξεργασία Συναρμολόγηση 8 Μεταξύ Εγκαταστάσεων Οντότητας 9 Αγορά 10 Εφοδιασμός πλοίων και αεροσκαφών
            '11 Δωρεάν διάθεση 12 Εγγύηση 13 Χρησιδανεισμός 14 Αποθήκευση σε Τρίτους
            '15 Επιστροφή από Φύλαξη 16 Ανακύκλωση 17 Καταστροφή άχρηστου υλικού   20 Μεταφορές - Ταχυμεταφορές
            '18 Διακίνηση Παγίων (Ενδοδιακίνηση) 19 Λοιπές Διακινήσεις

        End If

844
 Me.Caption = Me.Caption + "--t2bB--" + str(GetCurrentTime() - TT)
        'ΦΟΡΤΩΜΑ FORTOSHS
'846     R.Open "SELECT *FROM PINAKES WHERE TYPOS=8 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
'848     Do While Not R.EOF
'850         FORTOSH.AddItem R("PERIGRAFH")
'852         R.MoveNext
'        Loop
'854     If FORTOSH.ListCount = 0 Then
'856         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'ΕΔΡΑ ΜΑΣ',8)"
'        End If
'
'858     R.Close

Dim mok As Integer
mok = 0
For k = 0 To UBound(f_Arr_PIN) '(k_ARR, F_ARR_TYPOS) = R_arr!TYPOS
  If f_Arr_PIN(k, F_ARR_TYPOS) = 8 Then
        fortosh.AddItem f_Arr_PIN(k, F_ARR_PERIGRAFH)
        mok = 1
  End If
Next
If mok = 0 Then
   Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'ΕΔΡΑ ΜΑΣ',8)"
End If



        'ΦΟΡΤΩΜΑ ΠΡΟΟΡΙΣΜΟΥ
'860     R.Open "SELECT *FROM PINAKES WHERE TYPOS=9 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
'
'        ' If Not R.EOF Then R.MoveFirst
'
'862     Do While Not R.EOF
'864         PROORISMOS.AddItem R("PERIGRAFH")
'866         R.MoveNext
'        Loop
'
'868     If PROORISMOS.ListCount = 0 Then
'870         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'ΕΔΡΑ ΠΕΛΑΤΗ',9)"
'        End If
'
'872     R.Close
mok = 0
For k = 0 To UBound(f_Arr_PIN) '(k_ARR, F_ARR_TYPOS) = R_arr!TYPOS
  If f_Arr_PIN(k, F_ARR_TYPOS) = 9 Then
        PROORISMOS.AddItem f_Arr_PIN(k, F_ARR_PERIGRAFH)
        mok = 1
  End If
Next
If mok = 0 Then
   Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'ΕΔΡΑ ΠΕΛΑΤΗ',9)"
End If



















        'ΦΟΡΤΩΜΑ AYTOKINHTVN
'874     R.Open "SELECT *FROM PINAKES WHERE TYPOS=10 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
'
'876     Do While Not R.EOF
'878         AYTOKINHTO.AddItem R("PERIGRAFH")
'880         R.MoveNext
'        Loop
'882     If AYTOKINHTO.ListCount = 0 Then
'884         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'   ',10)"
'        End If
'886     R.Close

mok = 0
For k = 0 To UBound(f_Arr_PIN) '(k_ARR, F_ARR_TYPOS) = R_arr!TYPOS
  If f_Arr_PIN(k, F_ARR_TYPOS) = 10 Then
        AYTOKINHTO.AddItem f_Arr_PIN(k, F_ARR_PERIGRAFH)
        mok = 1
  End If
Next
If mok = 0 Then
   Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'   ',10)"
End If











        'ΦΟΡΤΩΜΑ ΑΠΑΛΛΑΓΩΝ ΦΠΑ
'888     R.Open "SELECT *FROM PINAKES WHERE TYPOS=44 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
    apallagesFPA.Clear
'
'892     Do While Not R.EOF
'894         apallagesFPA.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH")
'896         R.MoveNext
'        Loop
'        R.Close
        
mok = 0
For k = 0 To UBound(f_Arr_PIN) '(k_ARR, F_ARR_TYPOS) = R_arr!TYPOS
  If f_Arr_PIN(k, F_ARR_TYPOS) = 44 Then
        apallagesFPA.AddItem Format(f_Arr_PIN(k, F_ARR_AYJON), "#0") + "." + f_Arr_PIN(k, F_ARR_PERIGRAFH)
        mok = 1
  End If
Next
        
        
        
        
        
        
        
        
        
        
'Me.Caption = Me.Caption + "--t2bc--" + str(GetCurrentTime() - TT)

898     If apallagesFPA.ListCount < 32 Or InStr(apallagesFPA.List(13), ".28)") = 0 Then 'αν εχει αλλαξει το αρθρο 28  14.Χωρίς ΦΠΑ - άρθρο 33 του Κώδικα ΦΠΑ(π.28)
900         Gdb.Execute "DELETE FROM PINAKES  WHERE TYPOS=44"

902         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,1,'Χωρίς ΦΠΑ - άρθρο 3 του Κώδικα ΦΠΑ','VATEX-EU-132-1F')"
904         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,2,'Χωρίς ΦΠΑ - άρθρο 5 του Κώδικα ΦΠΑ','VATEX-EU-79-C')"

906         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,3,'Χωρίς ΦΠΑ - άρθρο 17 του Κώδικα ΦΠΑ','VATEX-EU-143')"
908         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,4,'Χωρίς ΦΠΑ - άρθρο 18 του Κώδικα ΦΠΑ','VATEX-EU-143')"
910         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,5,'Χωρίς ΦΠΑ - άρθρο 21 του Κώδικα ΦΠΑ','VATEX-EU-143')"

912         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,6,'Χωρίς ΦΠΑ - άρθρο 24 του Κώδικα ΦΠΑ','VATEX-EU-143')"
914         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,7,'Χωρίς ΦΠΑ - άρθρο 27 του Κώδικα ΦΠΑ(πρ.22)','VATEX-EU-132')"
916         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,8,'Χωρίς ΦΠΑ - άρθρο 29 του Κώδικα ΦΠΑ(πρ.24)','VATEX-EU-143-1B')"


918         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,9,'Χωρίς ΦΠΑ - άρθρο 30 του Κώδικα ΦΠΑ(πρ.25)','VATEX-EU-143-1C')"
920         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,10,'10 Χωρίς ΦΠΑ - άρθρο 31 του Κώδικα ΦΠΑ','VATEX-EU-148')"




922         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,11,'Χωρίς ΦΠΑ - άρθρο 32 του Κώδικα ΦΠΑ','VATEX-EU-148')"
924         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,12,'Χωρίς ΦΠΑ - άρθρο 32 - Πλοία Ανοικτής Θαλάσσης του Κώδικα ΦΠΑ','VATEX-EU-148-A')"
926         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,13,'Χωρίς ΦΠΑ - άρθρο 32.1.γ - Πλοία Ανοικτής Θαλάσσης του Κώδικα ΦΠΑ','VATEX-EU-148-A')"
928         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,14,'Χωρίς ΦΠΑ - άρθρο 33 του Κώδικα ΦΠΑ(πρ.28)','VATEX-EU-132')"
930         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,15,'Χωρίς ΦΠΑ - άρθρο 44 του Κώδικα ΦΠΑ','VATEX-EU-143')"

932         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,16,'Χωρίς ΦΠΑ(lapt)-άρθρο 45 του Κώδικα ΦΠΑ(πρ.39α)','VATEX-EU-143')"
934         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,17,'Χωρίς ΦΠΑ - άρθρο 47 του Κώδικα ΦΠΑ','VATEX-EU-148')"
936         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,18,'Χωρίς ΦΠΑ - άρθρο 48 του Κώδικα','VATEX-EU-143')"
938         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,19,'Χωρίς ΦΠΑ - άρθρο 54 του  Κώδικα ΦΠΑ','VATEX-EU-143-1K')"
940         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,20,'ΦΠΑ εμπεριεχόμενος - άρθρο 50 του Κώδικα ΦΠΑ','VATEX-EU-D')"

942         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,21,'ΦΠΑ εμπεριεχόμενος - άρθρο 51 του Κώδικα ΦΠΑ(πρ.44)','VATEX-EU-143')"
944         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,22,'ΦΠΑ εμπεριεχόμενος - άρθρο 52 του Κώδικα ΦΠΑ(πρ.45)','VATEX-EU-F')"
946         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,23,'ΦΠΑ εμπεριεχόμενος - άρθρο 53 του Κώδικα ΦΠΑ(πρ.46)','VATEX-EU-J')"
948         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,24,'Χωρίς ΦΠΑ-Αρθρο 8 του Κώδικα ΦΠΑ(πρ.6)','VATEX-EU-143')"
950         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,25,'Χωρις ΦΠΑ -ΠΟΛ.1029/1995','VATEX-EU-143-1B')"

952         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,26,'Χωρίς ΦΠΑ-ΠΟΛ.1167/2015','VATEX-EU-143')"
954         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,27,'Λοιπές Εξαιρέσεις ΦΠΑ','VATEX-EU-143')"

956         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,28,'άρθρο 29 περ. β` παρ.1 του Κώδικα ΦΠΑ, (Tax Free)','VATEX-EU-143-1B')"
958         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,29,'άρθρο 56 του Κώδικα ΦΠΑ (OSS μη ενωσιακό καθεστώς)','VATEX-EU-143')"
960         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,30,'άρθρο 57 του Κώδικα ΦΠΑ (OSS ενωσιακό καθεστώς)','VATEX-EU-132')"

962         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,31,'άρθρο 58 του Κώδικα ΦΠΑ (IOSS)','VATEX-EU-143-1C')"

964         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,C2) VALUES(44,99,'Δεν χρειάζεται Εξαίρεση ΦΠΑ ΦΠΑ=8','')"

        End If

966

        'ΦΟΡΤΩΜΑ parakrathshs
968     R.Open "SELECT *FROM PINAKES WHERE TYPOS=45 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
970     ComboPARAKRATISI.Clear

972     Do While Not R.EOF
974         ComboPARAKRATISI.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH")
976         R.MoveNext
        Loop

978     If ComboPARAKRATISI.ListCount < 17 Then
980         Gdb.Execute "delete FROM PINAKES WHERE TYPOS=45"

982         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,1,'Περιπτ. β`- Τόκοι - 15% ',15)"
984         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,2,'Περιπτ. γ` - Δικαιώματα - 20% ',20)"

986         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,3,'Περιπτ. δ` - Αμοιβές Συμβουλών Διοίκησης - 20% ',20)"

988         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,4,'Περιπτ. δ` - Τεχνικά Έργα - 3% ',3)"

990         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,5,'Υγρά καύσιμα και προϊόντα καπνοβιομηχανίας 1% ',1)"

992         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,6,'Λοιπά Αγαθά 4% ',4)"

994         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,7,'Παροχή Υπηρεσιών 8% ',8)"

996         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,8,'Προκαταβλητέος Φόρος Αρχιτεκτόνων και Μηχανικών επί Συμβατικών Αμοιβών, για Εκπόνηση Μελετών και Σχεδίων 4% ',4)"

998         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,9,'Προκαταβλητέος Φόρος Αρχιτεκτόνων και Μηχανικών επί Συμβατικών Αμοιβών, που αφορούν οποιασδήποτε άλλης φύσης έργα 10% ',10)"

1000        Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,10,'Προκαταβλητέος Φόρος στις Αμοιβές Δικηγόρων 15%  ',15)"

1002        Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,11,'Παρακράτηση Φόρου Μισθωτών Υπηρεσιών παρ. 1 αρ. 15 ν. 4172/2013  ',0)"

1004        Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,12,'Παρακράτηση Φόρου Μισθωτών Υπηρεσιών παρ. 2 αρ. 15 ν. 4172/2013 - Αξιωματικών Εμπορικού Ναυτικού ',15)"
1006        Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,13,'Παρακράτηση Φόρου Μισθωτών Υπηρεσιών παρ. 2 αρ. 15 ν. 4172/2013 - Κατώτερο Πλήρωμα Εμπορικού Ναυτικού',10)"

1008        Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,14,'Παρακράτηση Ειδικής Εισφοράς Αλληλεγγύης(ποσό) ',0)"
1010        Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,15,'Παρακράτηση Φόρου Αποζημίωσης λόγω Διακοπής Σχέσης Εργασίας παρ. 3 αρ. 15 ν. 4172/2013(ποσό)  ',0)"

1012        Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,16,'Παρακράτηση συναλλαγών αλλοδαπής βάσει συμβάσεων αποφυγής διπλής φορολογίας(Σ.Α.Δ.Φ.)(ποσό) ',0)"
1014        Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES(45,17,'Λοιπές Παρακρατήσεις Φόρου(ποσό) ',0)"

        End If

1016    R.Close
Me.Caption = Me.Caption + "--t2c--" + str(GetCurrentTime() - TT)
        'On Error Resume Next

1018    SKOPOS.Text = SKOPOS.List(0)

1020    fortosh.Text = fortosh.List(0)


        Dim R44 As New ADODB.Recordset

     R44.Open "select * from MEM", Gdb, adOpenDynamic, adLockOptimistic

     f_pelono = R44("pelono")
     f_peldie = R44("peldie")
    f_pelarit = R44("PELARIT")
     f_pelpol = R44("PELPOL")
     f_peltk = R44("PELTK")
     R44.Close



FORTDIE.Text = f_peldie
fortARit.Text = f_pelarit
FORTTK.Text = f_peltk
FORTPOL.Text = f_pelpol








1022    PROORISMOS.Text = PROORISMOS.List(0)

        ' AN EINAI 1o YPOKATASTHMA
1024    If Len(Dir("C:\KENTRIKO.TXT")) = 0 Or Len(Dir("C:\LAGEURO\MPOYGET.EXE", vbNormal)) = 0 Then
1026        APOTH1.Text = APOTH1.List(1)
1028        APOTH2.Text = APOTH2.List(0)
1030        f_YPOK = 1
            '   FORTOSH.Text = FORTOSH.List(1)
            '    Else
            '        APOTH1.Text = APOTH1.List(0)
            '        APOTH2.Text = APOTH2.List(1)
            '        f_YPOK = 0
        End If
    
        'Dim F_YP2 As Integer
 
1032    F_YP2 = Val(FINDPARAMETROI(1, "PAR1", "F_YP2", "1", "Γ. 12. ΑΡΙΘΜΟΣ ΠΡΟΕΠΙΛ.ΑΠΟΘ.ΠΡΟΟΡΙΣΜΟΥ"))

1034    If Len(Dir("C:\KENTRIKO.TXT")) > 0 Or gUserId = 1 Then
1036        APOTH1.Text = APOTH1.List(0)

1038        If F_YP2 > APOTH2.ListCount - 1 Then
1040            F_YP2 = 1
            End If

1042        APOTH2.Text = APOTH2.List(F_YP2)
1044        f_YPOK = 0
        End If

1046    If Len(Dir("C:\YPOK2.TXT")) > 0 Then
1048        APOTH1.Text = APOTH1.List(2)
1050        APOTH2.Text = APOTH2.List(0)
1052        f_YPOK = 2
        End If

1054    If Len(Dir("C:\YPOK3.TXT")) > 0 Then
1056        APOTH1.Text = APOTH1.List(3)
1058        APOTH2.Text = APOTH2.List(0)
1060        f_YPOK = 3
        End If

        'ΦΟΡΤΩΜΑ ΠΩΛΗΤΩΝ
1062    R.Open "SELECT *FROM PINAKES WHERE TYPOS=5 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

1064    Do While Not R.EOF
1066        Combo8.AddItem LTrim(str(R("AYJON"))) + "." + R("PERIGRAFH")
1068        R.MoveNext
        Loop

1070    R.Close

1072    Combo8.Text = Combo8.List(0)

1074    f_SeiresTimologioy = 80    'f_SeiresTimologioy = 30 OTAN DIALEGO  POLISEIS

1076    f_gDek_ajia = "###,###.00"
1078    f_ColUpdate = False
1080    f_error_gefyr = 0

1082    ProthFora = False

        ' πινακας γεφυρώσεων
1084    FlxGridDeltion.Clear
1086    FlxGridDeltion.Visible = False

        Dim fDB As Database

        'Dim R As Recordset

1088    Grid1.ColS = 16

        '2030    Grid1.ColWidth(0) = 300
        '2040    Grid1.ColWidth(f_k) = 1200
        '2050    Grid1.ColWidth(f_o) = 3800
        '2060    Grid1.ColWidth(f_p) = 1000
        '2070    Grid1.ColWidth(F_T) = 1000
        '2080    Grid1.ColWidth(f_e) = 800
        '2090    Grid1.ColWidth(f_f) = 800

1090    F_GRID1WKOD = Val(FINDPARAMETROI(6, "PAR1", "F_GRID1WKOD", "1200", "Πλέγμα Τιμολογησης Πλάτος ΚΩΔ."))
1092    F_GRID1WONO = Val(FINDPARAMETROI(6, "PAR1", "F_GRID1WONO", "3800", "Πλέγμα Τιμολογησης Πλάτος ΟΝΟΜΑ"))
1094    F_GRID1WPOSO = Val(FINDPARAMETROI(6, "PAR1", "F_GRID1WPOSO", "1000", "Πλέγμα Τιμολογησης Πλάτος  ΠΟΣΟΤΗΤΑ"))
1096    F_GRID1WTIMH = Val(FINDPARAMETROI(6, "PAR1", "F_GRID1WTIMH", "1000", "Πλέγμα Τιμολογησης Πλάτος  TIMH"))
1098    F_GRID1WEKPT = Val(FINDPARAMETROI(6, "PAR1", "F_GRID1WEKPT", "1000", "Πλέγμα Τιμολογησης Πλάτος ΕΚΠΤΩΣΗ"))
1100    F_GRID1WFPA = Val(FINDPARAMETROI(6, "PAR1", "F_GRID1WFPA", "1000", "Πλέγμα Τιμολογησης Πλάτος  ΦΠΑ"))

1102    F_GRID_WIDTH = Val(FINDPARAMETROI(6, "PAR1", "F_GRID_WIDTH", "10320", "Πλέγμα Τιμολογησης Πλάτος 10320 KAN"))
1104    F_KYBRAN = Val(FINDPARAMETROI(6, "PAR1", "F_KYBRAN", "0", "0=MONO TIMOKAT.KAΡΦ.TIMES(ΕΚΠ -1 ΓΕΝ) 1=ΚΑΙ ΕΚΠΤΩΣΗ ΣΤΟ NUM1"))

f_polagor = Val(FINDPARAMETROI(6, "PAR1", "f_polagor", "1", "Προεπιλογή 1=Πωλήσεις 0=Αγορές 2=Αποθηκη"))
f_EkkremhDeltia = Val(FINDPARAMETROI(6, "PAR1", "f_EkkremhDeltia", "1", "Προεπιλογή 1=Δειχνει εκκρεμη δελτία 0=δεν δείχνει"))
1106    F_GRID_FONT_SIZE = Val(FINDPARAMETROI(6, "PAR1", "F_GRID_FONT_SIZE", "8", "Πλέγμα Τιμολογησης FONTSIZE 8 KAN"))

1108    f_AKRIBOS_TON_KODIKO = Val(FINDPARAMETROI(6, "PAR1", "F_AKRIBOS_TON_KODIKO", "0", "1=ΑΚΡΙΒΩΣ ΤΟΝ ΚΩΔΙΚΟ 0=ΟΤΙ ΑΡΧΙΖΕΙ ΑΠΟ.."))

1110    F_GRIDLISTASTIMOL = (FINDPARAMETROI(6, "PAR1", "F_GRIDLISTASTIMOL", "1049;1700;1100;3550;480;1610;500;0;2500", "ΠΛΑΤΟΣ ΗΜΕ;ATIM;ΑΞΙΑ;ΣΥΝΑΛΛΑΣ;ΚΩΔ;ΠΑΡΑΤ;ΣΧΕΤ;ID;MARK"))

1112    F_DEFA_PLHR = Val(FINDPARAMETROI(7, "PAR1", "F_DEFA_PLHR", "0", "Προεπιλεγμένος τρόπος πληρωμής"))

1114    Grid1.ColWidth(0) = 400
1116    Grid1.ColWidth(f_k) = F_GRID1WKOD
1118    Grid1.ColWidth(f_o) = F_GRID1WONO
1120    Grid1.ColWidth(f_p) = F_GRID1WPOSO
1122    Grid1.ColWidth(F_T) = F_GRID1WTIMH
1124    Grid1.ColWidth(f_e) = F_GRID1WEKPT
1126    Grid1.ColWidth(f_f) = F_GRID1WFPA

1128    Grid1.Col = 0
1130    Grid1.row = 1
1132    Grid1.Text = " "
1134    Grid1.Col = 0

1136    For k = 1 To Grid1.rows - 1

1138        Grid1.row = k
1140        Grid1.Text = Format(k, "##")
1142        Grid1.RowHeight(k) = 250
        Next

1144    Grid1.row = 0
1146    Grid1.Col = f_k
1148    Grid1.Text = "Κωδικός"
1150    Grid1.Col = f_o
1152    Grid1.Text = "Περιγραφή"
1154    Grid1.Col = f_p
1156    Grid1.Text = "Ποσότητα"
1158    Grid1.Col = F_T
1160    Grid1.Text = "Τιμή Μονάδος"
1162    Grid1.Col = f_e
1164    Grid1.Text = "Εκπτωση"
1166    Grid1.Col = f_a
1168    Grid1.Text = "Αξία"

1170    Grid1.Col = f_f
1172    Grid1.Text = "Kατ.ΦΠΑ"
1174    Grid1.TextMatrix(0, F_mefpa) = "Αξ.με ΦΠΑ"
1176    Grid1.Col = F_PROEL
1178    Grid1.Text = "ΠΡΟΕΛΕΥΣΗ"

1180    Grid1.Col = F_EKPT2
1182    Grid1.Text = "ΕΚΠΤ 2"

1184    GridEidon.columns(0).Width = 18000

1186    If F_KOLA_MIKTA = 1 Then ' ΚΟΛΑ ΜΙΚΤΑ
 
1188        f_k = 1: f_o = 2: F_PROEL = 3: F_COL_KOLA = 4: F_mikta_KILA = 5: f_p = 6: F_T = 7: f_e = 8: f_f = 9: f_a = 10: F_M = 11: F_mefpa = 12: F_EKPT2 = 13: F_mikta = 14
            
1190        If F_METAFORIKES = 1 Then
1192            Grid1.TextMatrix(0, F_COL_KOLA) = "TEM"
1194            Grid1.TextMatrix(0, F_mikta_KILA) = "--"
1196            Grid1.TextMatrix(0, f_p) = "ΠΟΣΟΤΗΤΑ"
1198            Grid1.TextMatrix(0, F_PROEL) = "---"
            
            Else
1200            Grid1.TextMatrix(0, F_COL_KOLA) = "ΚΟΛΑ"
1202            Grid1.TextMatrix(0, F_mikta_KILA) = "ΜΙΚΤΑ ΚG"
1204            Grid1.TextMatrix(0, f_p) = "ΚΑΘΑΡΑ KG"
1206            Grid1.TextMatrix(0, F_PROEL) = "ΠΡΟΕΛΕΥΣΗ"
            End If
            
1208        Grid1.TextMatrix(0, F_T) = "ΤΙΜΗ"
1210        Grid1.TextMatrix(0, f_a) = "ΑΞΙΑ"
1212        If F_KOLA_MIKTA_LABELS = 2 Then

1214            Grid1.TextMatrix(0, F_COL_KOLA) = "ΠΛΑΤΟΣ"
1216            Grid1.TextMatrix(0, F_mikta_KILA) = "ΠΑΧΟΣ"
1218            Grid1.TextMatrix(0, f_p) = "ΠΟΣΟΤΗΤΑ"
1220            Grid1.TextMatrix(0, F_PROEL) = "ΜΗΚΟΣ"

            End If
            
        End If

        'data1.DATABASENAME = gDir
        'data1.RecordSource = "SELECT *FROM PINAKES"

        'Grid1.Row = 20
        'Fdb.Open gDir
        'R.Open "SELECT *FROM mem", db, adOpenKeyset, adLockReadOnly ', adCmdTableDirect

        'If gConnect = "Access" Then
        ' Set fDB = OpenDatabase(gDir, False, False)
        'Else
        '  Set fDB = OpenDatabase(gDir, False, False, gConnect)
        'End If

1222    synolo = " "    ' ΓΙΑ ΝΑ ΜΗΝ ΒΓΑΖΕΙ NULL

        '
        'If gConnect = "Access" Then
        '  Set DB = OpenDatabase(gDir, False, False)
        'Else
        '  Set DB = OpenDatabase(gDir, False, False, gConnect)
        'End If
        'Set r = DB.OpenRecordset("select *from mem")
        'r.MoveFirst
        'r.Move 3
        '
        ''r.MoveNext
        'gEPITR_SEIR = r("c3")
        'r.Close

        'For k = 1 To Len(gEPITR_SEIR)
        '    POLAGOR.AddItem Mid$(gEPITR_SEIR, k, 1)
        'Next

        'TIMEDIT.DatabaseName = gDir
        ' TIMEDIT.RecordSource = "SELECT *FROM TIMEDIT"
        ' TIMEDIT.ConnectionString = gConnect
Me.Caption = Me.Caption + "--t2d--" + str(GetCurrentTime() - TT)
        'EID.DatabaseName = gDir
1224    EID.RecordSource = "SELECT TOP 1 *FROM EID"
1226    EID.ConnectionString = gConnect

        'PEL.DatabaseName = gDir
1228    PEL.RecordSource = "SELECT top 1 *FROM PEL"
1230    PEL.ConnectionString = gConnect

1232    eggtim.ConnectionString = gConnect
1234    eggtim.RecordSource = "SELECT top 10 * FROM EGGTIM"

1236    DOK.ConnectionString = gConnect
1238    DOK.RecordSource = "SELECT top 1 *FROM DOK"

1240    LASTTIMOL.ConnectionString = gConnect

        Dim querpol As String
1242    If chkPOL.Value = vbChecked Then
1244        If chkAgor.Value = vbChecked Then
1246            querpol = " WHERE B_N1=" + Left(APOTH1.Text, 1) '""
            Else
1248            querpol = " WHERE B_N1=" + Left(APOTH1.Text, 1) + " AND TIM.EIDOS='e' "
           
            End If
        Else
1250        If chkAgor.Value = vbChecked Then
1252            querpol = " WHERE B_N1=" + Left(APOTH1.Text, 1) + " AND  TIM.EIDOS='r' "
            Else
1254            querpol = " WHERE B_N1=" + Left(APOTH1.Text, 1) + " AND TIM.EIDOS='' "
           
            End If
         
        End If

1256    LASTTIMOL.RecordSource = "SELECT top 5000 HME,ATIM+' '+(SELECT top 1 LEFT(TITLOS,15) FROM PARASTAT WHERE ASCII(PARASTAT.EIDOS)=ASCII(LEFT(ATIM,1))) AS [ΠΑΡ/ΚΟ],AJI AS [AΞΙΑ],PEL.EPO AS [ΕΠΩΝΥΜΙΑ],PEL.KOD AS [ΚΩΔΙΚΟΣ],PARAT AS [ΣΧΟΛΙΑ],SXETIKO AS [ΣΧΕΤ.ΠΑΡ],ID_NUM,ENTITYMARK,SXETMARK,INCMARK AS [MARK ΑΚΥΡΩΣΗΣ] FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD " + querpol + " ORDER BY ID_NUM DESC"
1258    LASTTIMOL.Refresh
1260    TDBGrid1.MarqueeStyle = 3  'SEIRA MARKARISMENH
1262    TDBGrid1.Splits(0).AlternatingRowStyle = True    'ENALAX XROMATA SEIRON
1264    TDBGrid1.Splits(0).ExtendRightColumn = True
          
        Dim RR() As String
   
1266    RR = Split(F_GRIDLISTASTIMOL, ";")
          
1268    If UBound(RR) = 8 Then
            Dim LLC As Integer
1270        For LLC = 0 To 8
1272            TDBGrid1.columns(LLC).Width = RR(LLC)
            Next
        Else
1274        TDBGrid1.columns(8).Width = 2500 'MARK
1276        TDBGrid1.columns(7).Width = 0 'DI_NUM
1278        TDBGrid1.columns(6).Width = 500 ' 500.882 'SXETIKO
1280        TDBGrid1.columns(5).Width = 1610  ' .142 'PARAT
1282        TDBGrid1.columns(4).Width = 480 ' 480.189 'KOD PELATH
1284        TDBGrid1.columns(3).Width = 3550 'EPO
1286        TDBGrid1.columns(2).Width = 1100 'AJI
1288        TDBGrid1.columns(1).Width = 1700 'ATIM
1290        TDBGrid1.columns(0).Width = 1049 'HME
       
        End If
       
        'TIMEDIT.Enabled = True
1292    EID.Enabled = True
1294    PEL.Enabled = True

1296    DTPicker1.Value = Now


 HMERPARAD.Value = Now
 ORAPARAD.Value = Now
 


1298    PAR1.Picture = LoadPicture(gPicture)

        'R.Open "SELECT *FROM TIMEDIT", Fdb, adOpenKeyset, adLockOptimistic

        ' If R.RecordCount < 40 Then
        '    For K = 1 To 40
        '      R.AddNew
        '      R.Update
        '    Next
        ' Else
        '    For K = 1 To 40
        'TIMEDIT.Recordset.Edit
        '     For l = 0 To R.Fields.Count - 1
        '            R(l) = Null
        '   Next
        '      TIMEDIT.Recordset.Update
        'TIMEDIT.Recordset(0) = "."

        '      TIMEDIT.Recordset.MoveLast

        '     TIMEDIT.Recordset.Move K - 1

        '      R.Update
        '      R.MoveNext
        '    Next
        'End If
        'R.MoveFirst

1300    SEIRA.Clear
1302    X = find_seires(0)

1304    For k = 1 To Len(X)
1306        SEIRA.AddItem mID$(X, k, 1)
        Next

1308    SEIRA.Text = SEIRA.List(0)

        ' Combo4.Clear
        ' Combo4.AddItem "ΜΕΤΡΗΤΟΙΣ"
        ' Combo4.AddItem "ΠΙΣΤΩΣΗ"
        ' Combo4.AddItem "ΠΙΣΤ.ΚΑΡΤΑ"
1310    Combo4.Text = Combo4.List(0)
        ' POLAGOR.SetFocus
        
1312    POLAGOR.Text = POLAGOR.List(f_polagor)
1314    POLAGOR_LostFocus

1316    PARAS.Enabled = True
        ' PARAS.ListIndex = geidos_paras  ' ""

        'DbGrid1.Columns(2).NumberFormat = gDek_pos 'δεκαδικά ποσοτήτων
        'DbGrid1.Columns(3).NumberFormat = gDek_ajia 'δεκαδικά αξίας
        'DbGrid1.Columns(6).NumberFormat = gDek_ajia 'δεκαδικά αξίας

       ' PARAS.SetFocus
        PARAS_LostFocus
        PARAS.Enabled = True
        PARAS.SetFocus
        
        'On Error GoTo Form_Load_Err
        'PARAS.SetFocus

        'μπαινω απ'ευθειας στο παραστατικό
1318    If Val(f_apeyueias) > 0 Then
1320        PARAS_LostFocus
1322        Text2(0) = f_apeyueias
1324        SendKeys "{ENTER} {ENTER}{ENTER} {ENTER}"
        End If

        
'Me.Caption = GetCurrentTime() - tt
        Me.Caption = Me.Caption + "--t2--" + str(GetCurrentTime() - TT)
        Me.Caption = ""
        Exit Sub

MHNYMA:
1326    HandleError "Par1:form_load"

1328    Resume Next

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        '        'MsgBox Err.Description & vbCrLf & _
                  "in ADOMERCNEW.par1.Form_Load " & _
                  "at line " & Erl, _
                  vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function find_parastat(recno As Integer, par As String, ReadWrite As Integer) As Long

        'ReadWrite=0   Βρίσκει τον αριθμο παραστατικού          -------------------------
        'ReadWrite=1   Aποθηκεύει τον αριθμο παραστατικού + 1   -------------------------
        'Βρίσκει τις κατηγορίες φπα
        '<EhHeader>
        On Error GoTo find_parastat_Err

        '</EhHeader>

        Dim k As Integer

        Dim X, Bohuhtiko As String

        '    On Error GoTo MHNYMA

        'Dim db As Database
        Dim R As New ADODB.Recordset

100     R.Open "SELECT * FROM PARASTAT WHERE ASCII(EIDOS)=" + str(Asc(par)), Gdb, adOpenDynamic, adLockOptimistic

102     If IsNull(R("ARITMISI")) Then
104         X = 0
        Else
106         X = R("ARITMISI")
        End If

108     R.Close

        'MEM..MoveFirst
        'For k = 1 To 9: f_mfpa(k) = R("fpa" + LTrim(Str(k))): Next
        Dim N As Integer

110     If X = 0 Then
112         find_parastat = 0

            Exit Function

        End If

114     If ReadWrite = 0 Then
116         R.Open "SELECT * FROM ARITMISI WHERE ID=" + str(X), Gdb, adOpenDynamic, adLockOptimistic

118         If IsNull(R("ARITMISI")) Then
120             find_parastat = 0
            Else
122             find_parastat = R("ARITMISI") + 1
            End If

124         R.Close

        Else
126         Gdb.Execute "UPDATE ARITMISI SET ARITMISI=" + Text1.Text + "  WHERE  ID=" + str(X), N

128         If N = 0 Then
130             MsgBox "Προσοχή αλλάξτε την αρίθμηση από <<Βοηθητικά>> <<Παράμετροι Παραστατικών>> "
            End If
        End If

        Exit Function

MHNYMA:
132     HandleError "Par1:FIND_Parastat"

134     Resume Next

        '<EhFooter>
        Exit Function

find_parastat_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.find_parastat " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.find_parastat " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function find_seires(DUM As Integer) As String

        '<EhHeader>
        On Error GoTo find_seires_Err

        '</EhHeader>

        'Βρίσκει τις υπάρχουσες σειρές
        Dim DB As Database

        Dim X  As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

        'Dim db As Database
        Dim R  As New ADODB.Recordset

        'If gConnect = "Access" Then
        ' Set db = OpenDatabase(gDir, False, False)
        'Else
        ' Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

        ' Data1.DatabaseName = gDir
        ' Data1.RecordSource = "SELECT *FROM PINAKES"

        'Fdb.Open gDir
        'R.Open "SELECT c3 FROM mem", Fdb, adOpenStatic, adLockReadOnly

        'adLockReadOnly ', adCmdTableDirect

100     R.Open "select C3 from MEM", Gdb, adOpenDynamic, adLockOptimistic
102     R.MoveFirst
104     R.Move 3

106     If IsNull(R("c3")) Then
108         find_seires = " "
        Else
110         find_seires = RTrim(R("c3"))
        End If

112     R.Close

        '<EhFooter>
        Exit Function

find_seires_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.find_seires " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.find_seires " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Function Open_Data(ByVal mtable As String) As Recordset
'
'    Set Open_Data = fDB.OpenRecordset(mtable)
'End Function

Private Sub FlxGridDeltion_Click()

        'If Grid1.Enabled = False Then ' φάση δελτίων αποστολής
        '<EhHeader>
        On Error GoTo FlxGridDeltion_Click_Err

        '</EhHeader>

100     FlagGridEidon = True
        '  If ProthFora = False Then
        '    ProthFora = True
        ' Set delt = EID.Recordset

102     DoEvents

104     If Val(FlxGridDeltion.TextMatrix(FlxGridDeltion.row, 1)) = 0 Then    ' If IsNull(EID.Recordset("metaf")) Or EID.Recordset("metaf") = 0 Then
106         FlxGridDeltion.TextMatrix(FlxGridDeltion.row, 1) = 1
108         FlxGridDeltion.Col = 2
110         FlxGridDeltion.CellBackColor = vbYellow
            'EID.Recordset("metaf") = 1: EID.Recordset.Update
        Else
            'EID.Recordset("metaf") = 0: EID.Recordset.Update
112         FlxGridDeltion.TextMatrix(FlxGridDeltion.row, 1) = 0

114         FlxGridDeltion.Col = 2
116         FlxGridDeltion.CellBackColor = vbWhite

        End If

118     FlxGridDeltion.Refresh

        '   Grid1.Enabled = True
        'End If
        'End If
        'EID.Refresh
        '<EhFooter>
        Exit Sub

FlxGridDeltion_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.FlxGridDeltion_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.FlxGridDeltion_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub PAR2SHOWING()

        '<EhHeader>
        On Error GoTo PAR2SHOWING_Err

        '</EhHeader>

        ' Dim par2 As par2
Dim m_ID As String
TDBGrid1.Col = 8
 TDBGrid1.Col = 9: ' M_ID = Left(TDBGrid1.Text, 8)
m_ID = TDBGrid1.columns("ID_NUM")




100     TDBGrid1.Col = 1
        'Set par2 = New par2

102     PAR2.Text1 = Left(TDBGrid1.Text, 8)

104     TDBGrid1.Col = 0
106     PAR2.DTPicker1 = CDate(TDBGrid1.Text)
108     PAR2.DTPicker2 = CDate(TDBGrid1.Text)

110     PAR2.CallFromOther.Caption = "1"

        'GIA OLA TA PARASTATIKA
        '  par2.Combo2OnomaParak.Text = "*"



'112     PAR2.Command1_Click   ANTI AYTO  TA PARAKATO:


       Dim XX As String
         XX = "select '-'+PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.KOD,PEL.EIDOS,TRP,SKOPOS,FORTOSH,ART,PROOR,PARAT,ART,B_N1,B_N2,SXETIKO,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,ID_NUM,ENTITYMARK,ORA   " & " from TIM LEFT join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where ID_NUM=" + m_ID
'        Else
'         XX = "select '-'+PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.KOD,PEL.EIDOS,TRP,SKOPOS,FORTOSH,ART,PROOR,PARAT,ART,B_N1,B_N2,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,ID_NUM,ENTITYMARK,ORA   " & " from TIM LEFT join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' order by HME,PEL.EPO;"
 '       End If
PAR2.TIM.ConnectionString = gConnect
        PAR2.TIM.RecordSource = XX
        PAR2.TIM.Refresh
        











114     PAR2.SHOW_TIMOL
116     PAR2.SHOW

118     PAR2.bohuitiko.SetFocus

120     DoEvents

122     PAR2.Command1.SetFocus
124     PAR1.LASTTIMOL.Refresh

        'On Error GoTo 0

        '   SendKeys "{enter}"    'ENTER  ~

        '<EhFooter>
        Exit Sub

PAR2SHOWING_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.PAR2SHOWING " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.PAR2SHOWING " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub SynoloG_Click()

        '<EhHeader>
        On Error GoTo SynoloG_Click_Err

        '</EhHeader>
        Dim c As String

100     c = InputBox("τελικό ποσό", "Αλλαγή τελικού ποσού", SynoloG.Caption)
102     Text5.Text = "*" + c
104     Text5_LostFocus
106     b_Katax.SetFocus

        '<EhFooter>
        Exit Sub

SynoloG_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.SynoloG_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.SynoloG_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub tamiaki2(Tmhma As String, TR1METR As String, TR2KARTA As String, TR3ALLO As String)
Dim Lektiko23 As String, FPATMHMA As String

If Tmhma = "5" Then
     Lektiko23 = "ΥΠΗΡΕΣΙΕΣ ΕΛΑΙΟΤΡΙΒΕΙΟΥ": FPATMHMA = "13 5"
ElseIf Tmhma = "9" Then
     Lektiko23 = "ΥΠΗΡ.ΕΛΑΙΟΤΡΙΒΕΙΟΥ(ΔΙΚ)": FPATMHMA = "13 5"
     
     
ElseIf Tmhma = "2" Then
     Lektiko23 = "ΑΔΕΙΑ ΔΟΧΕΙΑ           ": FPATMHMA = "24 2"
              
     
ElseIf Tmhma = "6" Then
     Lektiko23 = "ΠΥΡΗΝΑΣ                ": FPATMHMA = "13 6"
     
     
  ElseIf Tmhma = "1" Then
     Lektiko23 = "ΛΑΔΙ                   ": FPATMHMA = "13 1"
     
     
Else
     Lektiko23 = "ΛΑΔΙ                   ": FPATMHMA = "13 1"
     


     
End If
     

      '"plu" 00-0038      Ξ•Ξ™Ξ”ΞΞ£ 1 (Ξ¦Ξ Ξ‘ 6%)              1,000   1230,10     6,00 1
        '"mod" -0,09        123456789012345678901234
        '"plu" 00-0039      Ξ•Ξ™Ξ”ΞΞ£ 2 (Ξ¦Ξ Ξ‘ 13%)             1,000      0,11    13,00 2
        '"mod" -0,10
        '"plu" 00-0040      Ξ•Ξ™Ξ”ΞΞ£ 3 (Ξ¦Ξ Ξ‘ 24%)             2,000      0,10    24,00 3
        '"mod" -0,19
        '"plu" 00-0041      Ξ•Ξ™Ξ”ΞΞ£ 4 (Ξ¦Ξ Ξ‘ 0%)              3,000      0,10     0,00 5
        '"mod" -0,29

        '"dis" 0,01
        '"to1"  1      0.01
        '"to2"  2      0.01
        '"to3"  3      0.01

        'FileOpen(1, "C:\FISLINK\TESTFL.1", OpenMode.Output, OpenAccess.Write)
        'FilePut(1, "SSSSA")
        'FilePut(1, "SSSSA2")
        'FileClose(1)


TR1METR = Replace(TR1METR, ",", ".")
TR2KARTA = Replace(TR2KARTA, ",", ".")
TR3ALLO = Replace(TR3ALLO, ",", ".")


        On Error Resume Next
        Dim c As String
        Dim N As Single

        N = Val(TR1METR) + Val(TR2KARTA) + Val(TR3ALLO)


        If N < 0.03 Then
            MsgBox ("ΔΕΝ ΒΑΛΑΤΕ ΠΟΣΟ")
            Exit Sub
        End If




       ' Dim objStreamWriter As StreamWriter


        'Pass the file path and the file name to the StreamWriter constructor.
        'objStreamWriter = New StreamWriter("C:\FISLINK\TESTFL.1", False, System.Encoding.Default)
        Open "C:\FISLINK\TESTFL.1" For Output As #1




        Dim C0 As String: C0 = Replace(Format(N - 0, "######0.00"), ".", ",")
        Dim N0 As Integer
        N0 = Len(C0)
        'lekt = "ΥΠΗΡΕΣΙΕΣ ΕΛΑΙΟΤΡΙΒΕΙΟΥ"
         c = """plu""" + " 00-0038      " + Lektiko23 + "       1,00" + Space(11 - N0) + C0 + "       " + FPATMHMA  ' + Format(Tmhma, "0")


        'c = """plu""" + " 00-0038      ΥΠΗΡΕΣΙΕΣ ΕΛΑΙΟΤΡΙΒΕΙΟΥ       1,00" + Space(11 - N0) + C0 + "       13 5"

        Print #1, (c)










'        c = """plu""" + " 00-0038      " + mID("....." + Space(24), 1, 24) + "      1,000   " + "   0,01" + "       13 1"
'
'        Print #1, (c)
'
'
'        c = """plu""" + " 00-0038      " + mID("......" + Space(24), 1, 24) + "      1,000   " + "   0,01" + "       13 1"
'
'        Print #1, (c)

       ' C = """plu""" + " 00-0038      " + Format(d1.Value, "dd/MM/yy") + " Ξ•Ξ©Ξ£ " + Format(d2.Value, "dd/MM/yy") + "         1,000   " + "   0,01" + "       13 1"

      '  Print #1, (C)


        'Write a second line of text.
        Dim TR As String
        Dim POSO As Single
        POSO = 0
        TR = "0"

        If Val(TR1METR) > 0 Then
            TR = "1"
            POSO = Val(TR1METR)
            Print #1, ("""to1""" + "  " + TR + "    " + Replace(Format(POSO, "000.00"), ".", ","))



            If Val(TR2KARTA) > 0 Then
                TR = "2"
                POSO = Val(TR2KARTA)
                Print #1, ("""to2""" + "  " + TR + "    " + Replace(Format(POSO, "000.00"), ".", ","))
            End If

        Else ' tr1=0

            If Val(TR2KARTA) > 0 Then
                TR = "2"
                POSO = Val(TR2KARTA)
                Print #1, ("""to1""" + "  " + TR + "    " + Replace(Format(POSO, "000.00"), ".", ","))
            End If




        End If


        '  PRINT #1,("""to1""" + "  " + tr + "    " + Replace(Format(poso, "000.00"), ".", ","))
        '     Else
        '        PRINT #1,("""to1""" + "  0    " + Replace(Format(Val(TR1METR), "000.00"), ".", ","))
        '   End If

        '      PRINT #1,("""to2""" + "  2    " + Replace(Format(Val(TR2KARTA), "000.00"), ".", ","))
        '     PRINT #1,("""to3""" + "  3    " + Replace(Format(Val(TR3ALLO), "000.00"), ".", ","))



        '"to2"  2      0.01
        '"to3"  3      0.01)

        'Close the file.
        'objStreamWriter.Close()
          Close #1


        TR1METR = ""
        TR2KARTA = ""
        TR3ALLO = ""
    




End Sub

Private Sub tamiaki_Click()
 '  tamiaki2 "12", "", ""
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, _
                                   Bookmark As Variant, _
                                   ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
        ' RSClone.Bookmark = Bookmark
        
        '    If ButtonFlag = 0 And RSClone("CustType").Value = 1 Then
        '     RowStyle = Prospective
        '    End If
        
        '   If ButtonFlag = 1 And RSClone("CustType").Value = 4 Then
100     RowStyle = Distributors
        '  End If
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)

        '================================================================================
        '<EhHeader>
        On Error GoTo TDBGrid1_HeadClick_Err

        '</EhHeader>
        Dim sumes(100) As String

        Dim k          As Long

        On Error Resume Next

        '
        ' 'edo parakato exei latos 15/8/2009
        '
        ' For k = 0 To Data1.Recordset.FIELDS.Count - 1
        '        SUMES(k) = TDBGrid.Splits(0).Columns(k).FooterText '  = Format(SUMES(k), "######0.00")
        '  Next
        '
        '
        '

100     If LASTTIMOL.Recordset.Sort = "[" & TDBGrid1.columns(ColIndex).DataField & "] asc" Then   ' strSort
102         LASTTIMOL.Recordset.Sort = "[" & TDBGrid1.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
104         LASTTIMOL.Recordset.Sort = "[" & TDBGrid1.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

        '<EhFooter>
        Exit Sub

TDBGrid1_HeadClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.TDBGrid1_HeadClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.TDBGrid1_HeadClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

        '<EhHeader>
        On Error GoTo TDBGrid1_MouseUp_Err

        '</EhHeader>
100     If Button = 2 Then

            'MDIForm1.runt.Item(0).Caption = "Παρουσίαση"
            'MDIForm1.runt.Item(1).Caption = "Διόρθωση"
            On Error Resume Next
102         TDBGrid1.SetFocus
104         Call Me.PopupMenu(MDIForm1.par1menu)
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

'106     If Button = 1 And Shift = 1 Then  ' shift+ click
'
'            Dim ANS As Integer
'            Dim c, d, SX As String
'108         ANS = MsgBox("Να γίνει επαναποστολή ακυρωτικού;", vbYesNo)
'
'110         If ANS = vbYes Then
'
'112             PAR1.TDBGrid1.Col = 7  'id_num  ακυρωτικου
'
'114             d = TDBGrid1.Text
'                ' d = InputBox("δωσε τον αριθμό του id", 0)
'
'116             SX = GET_CVALUE("select ISNULL(SXETMARK,'00') AS SXX from TIM WHERE ID_NUM=" + d)
'118             If SX = "00" Then
'120                 MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΜΑΡΚ ΠΟΥ ΘΑ ΑΚΥΡΩΘΕΙ"
'                    Exit Sub
'                End If
'                Dim TIMOL As String
'122             TIMOL = GET_CVALUE("select ATIM from TIM WHERE ID_NUM=" + d)
'                   Par7MyData.Visible = False
'124             c = Par7MyData.CANCEL_INVOICE(SX, TIMOL)
'126             If Len(c) > 5 Then
'128                 Gdb.Execute "update TIM set ENTITYMARK='" + c + "' WHERE ID_NUM=" + d
'130                 MsgBox "οκ ακυρωθηκε και το μαρκ ακύρωσης είναι το " + c
'                End If
'            End If
'        End If

        '<EhFooter>
        Exit Sub

TDBGrid1_MouseUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.TDBGrid1_MouseUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.TDBGrid1_MouseUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
Public Sub epanAkyr()
            Dim ANS As Integer
            Dim c, D, SX As String


112             PAR1.TDBGrid1.Col = 7  'id_num  ακυρωτικου

114             D = TDBGrid1.Text
                ' d = InputBox("δωσε τον αριθμό του id", 0)

116             SX = GET_CVALUE("select ISNULL(SXETMARK,'00') AS SXX from TIM WHERE ID_NUM=" + D)
118             If SX = "00" Or Len(SX) = 0 Then
120                 MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΜΑΡΚ ΠΟΥ ΘΑ ΑΚΥΡΩΘΕΙ(SXETMARK)"
                    Exit Sub
                End If
                
                
                
                If Len(TDBGrid1.columns("ENTITYMARK")) > 10 Then
                   MsgBox ("ΕΙΝΑΙ ΗΔΗ ΜΕ ΜΑΡΚ")
                   'If 1 = GET_NVALUE("select ISMYDATA from PARASTAT WHERE EIDOS='" + Left(TDBGrid1.columns(1), 1) + "'") Then
                   
                   Exit Sub
                    'End If
                End If
                Dim N As Integer
                N = GET_NVALUE("select ISMYDATA from PARASTAT WHERE EIDOS='" + Left(TDBGrid1.columns(1), 1) + "'")
                
                If N = 1 Then
                   MsgBox " ΔΕΝ ΕΙΝΑΙ ΑΚΥΡΩΤΙΚΟ. KANTE ΑΠΟΣΤΟΛΗ MYDATA"
                   Exit Sub
                
                End If
                
                
                
                
121              ANS = MsgBox("Να γίνει επαναποστολή ακυρωτικού;", vbYesNo)

122         If ANS = vbYes Then
                
                Dim TIMOL As String
123             TIMOL = GET_CVALUE("select ATIM from TIM WHERE ID_NUM=" + D)
                   Par7MyData.Visible = False
124             c = Par7MyData.CANCEL_INVOICE(SX, TIMOL)
126             If Len(c) > 5 Then
128                 Gdb.Execute "update TIM set ENTITYMARK='" + c + "' WHERE ID_NUM=" + D
130                 MsgBox "οκ ακυρωθηκε και το μαρκ ακύρωσης είναι το " + c
                End If
            End If
End Sub
Private Sub Text1_GotFocus()

        '<EhHeader>
        On Error GoTo Text1_GotFocus_Err

        '</EhHeader>

100     Text1.BackColor = vbYellow
102     Text1.SelStart = 0
104     Text1.SelLength = Len(Text1.Text)

        '<EhFooter>
        Exit Sub

Text1_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Text1_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Text1_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Text1_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = 13 Then
102         KeyAscii = 0

104         If POLAGOR.Text = POLAGOR.List(0) Then
106             DTPicker1.SetFocus
            Else
108             selectPEL.SetFocus
            End If
        End If

        '<EhFooter>
        Exit Sub

Text1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Text1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Text1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_EnterCell()

        '  If Grid1.MouseRow = 0 Then
        '    timtext3.Visible = False
        '    Exit Sub
        '   End If
        '<EhHeader>
        On Error GoTo Grid1_EnterCell_Err

        '</EhHeader>

        On Error GoTo MHNYMA

100     If Grid1.row > f_SeiresTimologioy Then
102         Grid1.row = f_SeiresTimologioy
        End If

104     timtext3.Text = Grid1.Text
106     timtext3.Visible = False
108     timtext3.Top = Grid1.Top + Grid1.CellTop
110     timtext3.Left = Grid1.Left + Grid1.CellLeft
112     timtext3.Width = Grid1.CellWidth - 15
114     timtext3.Height = Grid1.CellHeight - 25

116     timtext3.SelStart = 0
118     timtext3.SelLength = Len(timtext3.Text)

120     timtext3.Visible = True
122     timtext3.SetFocus

        Exit Sub

MHNYMA:
124     HandleError "Par1:EnterCell"

126     Resume Next

        '<EhFooter>
        Exit Sub

Grid1_EnterCell_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Grid1_EnterCell " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Grid1_EnterCell " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_LeaveCell()
    'Grid1.Text = timtext3.Text
End Sub

Private Sub Text1_LostFocus()

        '<EhHeader>
        On Error GoTo Text1_LostFocus_Err

        '</EhHeader>

100     Text1.BackColor = vbWhite

        '<EhFooter>
        Exit Sub

Text1_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Text1_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Text1_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_GotFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text2_GotFocus_Err

        '</EhHeader>
100     If index = 3 Then
102         Text2(index).BackColor = vbYellow
104         SendKeys ""
            ' MILSEC 1000   '500
106         Text2(3).SetFocus
        End If

108     Text2(index).BackColor = vbYellow

110     Text2(index).SelStart = 0
112     Text2(index).SelLength = Len(Text2(index).Text)

        '<EhFooter>
        Exit Sub

Text2_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Text2_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Text2_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)

    '  On Error GoTo MHNYMA
    '
    ' If KeyCode = 13 Then
    '
    '   KeyCode = 0
    '   If Index = 0 Then
    '
    '     If Len(Trim(Text2(0).Text)) = 0 Then
    '        Text2(1).SetFocus ' δώσε όνομα
    '     Else
    '        GridPelaton.Columns(0).width = 5000
    '        GridPelaton.Visible = True
    '        GridPelaton.ZOrder 0
    '        GridPelaton.SetFocus 'διαλεξε πελάτη
    '     End If 'len
    '
    '   End If 'index=0
    '
    '   If Index = 1 Then
    '
    '     If Len(Trim(Text2(1).Text)) = 0 Then
    '
    '        Text2(2).SetFocus ' δώσε ΑΦΜ
    '
    '     Else
    '        GridPelaton.Visible = True
    '        GridPelaton.ZOrder 0
    '        GridPelaton.SetFocus 'διαλεξε πελάτη
    '
    '     End If 'len
    '
    '   End If 'index=0
    '
    '   If Index = 2 Then
    '        GridPelaton.Visible = True
    '        GridPelaton.ZOrder 0
    '        GridPelaton.SetFocus 'διαλεξε πελάτη
    '
    '   End If 'index=0
    '
    'End If ' keyascci=13
    '
    '
    'Exit Sub
    '
    'MHNYMA:
    'HandleError "Par1:text2_keyUP"
    'Resume Next
    '
    '
    '

End Sub

Private Sub Text2_KeyPress(index As Integer, KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Text2_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = 13 Then
102         KeyAscii = 0
            ' selectPEL.SetFocus
        End If

        '<EhFooter>
        Exit Sub

Text2_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Text2_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Text2_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Text2_KeyUp_Err

        '</EhHeader>

        On Error GoTo MHNYMA

100     If KeyCode = 13 Then

102         KeyCode = 0

104         If index = 0 Then

106             If Len(Trim(Text2(0).Text)) = 0 Then
108                 Text2(1).SetFocus    ' δώσε όνομα
                Else
110                 GridPelaton.columns(0).Width = 5000
112                 GridPelaton.Visible = True
114                 GridPelaton.ZOrder 0
116                 GridPelaton.SetFocus    'διαλεξε πελάτη
                End If    'len

            End If    'index=0

118         If index = 1 Then

120             If Len(Trim(Text2(1).Text)) = 0 Then

122                 Text2(2).SetFocus    ' δώσε ΑΦΜ

                Else
                
124                 If f_AmesosPelatisLian = 1 Then
126                     Text2(3).BackColor = vbYellow
128                     Text2(3).SetFocus
                    Else
130                     GridPelaton.Visible = True
132                     GridPelaton.ZOrder 0
134                     GridPelaton.SetFocus    'διαλεξε πελάτη
                    End If

                End If    'len

            End If    'index=0

136         If index = 2 Then
138             GridPelaton.Visible = True
140             GridPelaton.ZOrder 0
142             GridPelaton.SetFocus    'διαλεξε πελάτη

            End If    'index=0

144         If index = 3 Then 'dieyuynsh

146             GetNextPel
          
148             Text2(1).Text = ""
150             Text2(3).Text = ""
          
152             Text2(0).SetFocus
          
                'GridPelaton.Columns(0).width = 5000
                'GridPelaton.Visible = True
                ' GridPelaton.ZOrder 0
                ' GridPelaton.SetFocus    'διαλεξε πελάτη

            End If    'index=0

        End If    ' keyascci=13

        Exit Sub

MHNYMA:
154     HandleError "Par1:text2_keyUP"

156     Resume Next

        '<EhFooter>
        Exit Sub

Text2_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Text2_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Text2_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub GetNextPel()

        'παιρνει το επόμενο νουμερο και εισαγει νεο πελάτη
        '<EhHeader>
        On Error GoTo GetNextPel_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

        On Error Resume Next

        Dim m    As String

        Dim mEpo As String, MDIE As String

100     mEpo = Replace(Text2(1).Text, "'", """")
102     MDIE = Replace(Text2(3).Text, "'", """")

        Dim k As Long

104     k = 0

106     Do While k < 10000
108         Gdb.Execute "update MEM SET EPEL=EPEL+1 WHERE ID=1"
110         R.Open "SELECT EPEL FROM MEM WHERE ID=1", Gdb, adOpenDynamic, adLockOptimistic
112         m = LTrim(str(R(0)))
114         R.Close
116         Me.Caption = "παρακαλω περιμενετε!! " + m
118         k = k + 1
120         R.Open "SELECT COUNT(*) FROM PEL WHERE KOD='" + m + "'", Gdb, adOpenDynamic, adLockOptimistic

122         If R(0) = 0 Then

                Exit Do

            End If

        Loop

124     If k > 9999 Then
126         MsgBox "δεν βρεθηκε κενος κωδικός"

            Exit Sub

        End If
   
128     Gdb.Execute "INSERT INTO PEL (EIDOS,KOD,EPO,DIE) VALUES ('e','" + m + "','" + mEpo + "','" + MDIE + "')", k
130     Text2(0) = m
        ' Text2(0).SetFocus

        '<EhFooter>
        Exit Sub

GetNextPel_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.GetNextPel " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.GetNextPel " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_LostFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text2_LostFocus_Err

        '</EhHeader>

        Dim X As String

        On Error GoTo MHNYMA

100     Text2(index).BackColor = vbWhite

102     If Len(Trim(Text2(index))) = 0 Then Exit Sub
104     If fbuff = "l" Then fbuff = "r"

106     X = fbuff    'If Option1(1) Then x = "e" Else x = "r"

        Dim MSYNT

108     MSYNT = ""

110     If f_KENTRA_ADYNATISMATOS = 1 Then
112         If X = "e" Then
114             If kentroAdyn.List(0) = kentroAdyn.Text Then
116                 MSYNT = "NUM3=0 AND "
118                 Grid1.BackColor = &HFF8080     ' mple
                    'Grid1.BackColor = &H80FF80 ' prasino
                Else
120                 MSYNT = "NUM3=1 AND "
122                 Grid1.BackColor = &H80FF80 ' prasino
                End If
            End If
        End If

124     If f_MONO_PELATES_KATHGORIAS = 1 Then
126         If X = "e" And F_rec_m > 0 Then  '0=OLES OI KATHGORIES
                
128             MSYNT = MSYNT + " KODGAL=" + str(F_rec_m) + " AND "
                   
            End If
        End If

130     If index = 3 Then  ' διευθυνση
132         If f_AmesosPelatisLian = 1 Then
                'οταν περναει τον πελατη λιανικης να μην δεχεται ΑΦΜ ΓΙΑΤΙ ΚΟΛΛΑΕΙ
134             If Len(Text2(3).Text) > 0 Then
136                 Text2(2).Enabled = False
                      
                    ' αν εφυγα με το ποντικι να κανει το ιδιο σαν να εφυγα με εντερ
                    ' δηλαδη να ανοιγει πελατη
138                 If Len(Text2(0).Text) = 0 Then
140                     GetNextPel
142                     Text2(1).Text = ""
144                     Text2(3).Text = ""
146                     Text2(0).SetFocus
                    End If
                      
                End If
            End If
         
        End If

148     If index = 1 Then  'eponumo
            Dim CW4 As String
150         If F_MONOENERGOS = 1 Then CW4 = " AND ISNULL(ENERGOS,1)=1 " Else CW4 = ""
152         If gCapitals = 1 Then
154             PEL.RecordSource = "select LEFT(left(EPO+'-'+ISNULL(DIE,'')+'-',56)+KOD,60) as xx,ISNULL(EPA,'') AS EPA,ISNULL(THL,'') AS THL,ISNULL(DIE,'') AS DIE,KOD,ISNULL(EMAIL,'') AS EMAIL,XRVMA AS TK  from PEL where " + MSYNT + " EIDOS='" + X + "'" + CW4 + " and EPO like '" + (SameLetters(Text2(1).Text)) + "%' order by EPO;"
            Else
156             PEL.RecordSource = "select LEFT(left(EPO+'-'+ISNULL(DIE,'')+'-',56)+KOD,60) as xx,ISNULL(EPA,'') AS EPA,ISNULL(THL,'') AS THL,ISNULL(DIE,'') AS DIE,KOD,ISNULL(EMAIL,'') AS EMAIL,XRVMA AS TK  from PEL where " + MSYNT + " EIDOS='" + X + "'" + CW4 + " and EPO like '" + SameLetters(Text2(1).Text) + "%' order by EPO;"
            End If
        End If

158     If index = 0 Then    'kodikos
160         PEL.RecordSource = "select  LEFT(left(EPO+'-'+ISNULL(DIE,'')+'-',56)+KOD,60) as xx,ISNULL(DIE,'') AS DIE,ISNULL(THL,'') AS THL,KOD,ISNULL(EMAIL,'') AS EMAIL,XRVMA AS TK  from PEL where " + MSYNT + " EIDOS='" + X + "' and KOD like '" + SameLetters(Text2(0).Text) + "%' order by KOD;"
        End If

162     If index = 2 Then    'afm
164         PEL.RecordSource = "select  LEFT(left(EPO+'-'+ISNULL(DIE,'')+'-',56)+KOD,60) AS XX,ISNULL(DIE,'') AS DIE,ISNULL(THL,'') AS THL,KOD,ISNULL(EMAIL,'') AS EMAIL,XRVMA AS TK  from PEL where " + MSYNT + " EIDOS='" + X + "' and (AFM like '%" + Text2(2).Text + "%' OR " + f_pedioAFM + " like '%" + Text2(2).Text + "%' ) order by AFM;"
        End If

166     GridPelaton.columns(0).Width = 6000
168     GridPelaton.columns(0).DataField = "XX"

170     PEL.Refresh
172     GridPelaton.columns(0).Width = 6000
174     GridPelaton.Visible = True
176     GridPelaton.SetFocus

178     If PEL.Recordset.RecordCount = 0 And index = 2 Then 'me afm alla den brika
180         Dim u As String, p As String, formatpel As String: formatpel = FINDPARAMETROI(1, "PELAT1", "F_FORMAT_PEL_KOD", "0000", "ΨΗΦΙΑ ΚΩΔΙΚΟΥ ΠΕΛΑΤΩΝ")
182         u = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_USER", "LAGAKIS23041960", "ΧΡΗΣΤΗΣ ΕΥΡΕΣΗΣ ΑΦΜ")
184         p = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_PWD", "d389833921", "ΚΩΔΙΚΟΣ ΕΥΡΕΣΗΣ ΑΦΜ")
            Dim newKod As String, EPO As String, DIE As String, EPA As String, pol As String, tk As String, doy As String, c As String, ENERGOS As Integer
186         c = Text2(2).Text: ENERGOS = 1
188         Par7MyData.requestafm2 u, p, c, EPO, DIE, EPA, pol, tk, doy, ENERGOS
             
190         Dim ANS As Integer: ANS = MsgBox("Να ανοιχτεί κωδικός για τον " + EPO, vbYesNo)
192         If ANS = vbYes Then
               
194             newKod = pelat1.get_next_pel(X, 0, formatpel)
196             Dim sql As String: sql = "insert into PEL(XRVMA,ENERGOS,EIDOS,KOD,EPO,DIE,POL,EPA,AFM,DOY) VALUES "
198             sql = sql + "('" + tk + "',1,'" + X + "','" + Format(newKod, formatpel) + "','" + EPO + "','" + DIE + "','" + pol + "','" + EPA + "','" + c + "','" + doy + "')"
200             Gdb.Execute sql
202             If X = "r" Then
204                 Gdb.Execute "update MEM SET RPEL=" + str(newKod)
                Else
206                 Gdb.Execute "update MEM SET EPEL=" + str(newKod)
                End If
208             PEL.RecordSource = "select  LEFT(left(EPO+'-'+ISNULL(DIE,'')+'-',56)+KOD,60) AS XX,ISNULL(DIE,'') AS DIE,ISNULL(THL,'') AS THL,KOD,ISNULL(EMAIL,'') AS EMAIL,XRVMA AS TK  from PEL where " + MSYNT + " EIDOS='" + X + "' and (AFM like '%" + Text2(2).Text + "%' OR " + f_pedioAFM + " like '%" + Text2(2).Text + "%' ) order by AFM;"
210             PEL.Refresh
               
            End If
        
        End If

        Exit Sub

MHNYMA:
212     HandleError "Par1:FINDSYNOLO"

214     Resume Next

        '<EhFooter>
        Exit Sub

Text2_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Text2_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Text2_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)

100     If KeyCode = vbKeyReturn Then
102         KeyCode = 0    'suppress the beep
104         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

End Sub

Private Sub Textm_Change(index As Integer)
        
         



100     If F_METAF3SYNAL = 1 Then
            Dim mEpo, MDIE, MAFM
            '2->3  4->0
102         If index = 5 Or index = 4 Then
                Dim mmID As Long
          
104             mmID = GET_NVALUE("SELECT ID FROM PEL WHERE ( AFM LIKE '%" + Textm(index) + "%' OR  EPO LIKE '%" + Textm(index).Text + "%' ) AND EIDOS='e'")
                ' mEpo = GET_CVALUE("SELECT TOP 1 ISNULL(EPO,'') AS MEPO FROM PEL WHERE ( AFM LIKE '%" + Textm(Index) + "%' OR  EPO LIKE '%" + Textm(Index) + "%' ) AND EIDOS='e'")
                'MDIE = GET_CVALUE("SELECT TOP 1 ISNULL(DIE,'') AS MEPO FROM PEL WHERE EPO LIKE '%" + Textm(Index) + "%' AND EIDOS='e'")
                'MAFM = GET_CVALUE("SELECT TOP 1 ISNULL(AFM,'') AS MEPO FROM PEL WHERE EPO LIKE '%" + Textm(Index) + "%' AND EIDOS='e'")
          
106             If mmID > 0 Then
108                 mEpo = GET_CVALUE("SELECT TOP 1 ISNULL(EPO,'') AS MEPO FROM PEL WHERE ID=" + str(mmID))
110                 MDIE = GET_CVALUE("SELECT TOP 1 ISNULL(DIE,'') AS MEPO FROM PEL WHERE ID=" + str(mmID))
112                 MAFM = GET_CVALUE("SELECT TOP 1 ISNULL(AFM,'') AS MEPO FROM PEL WHERE ID=" + str(mmID))
            
114                 Textm(IIf(index = 5, 3, 0)) = "ΕΠΩΝΥΜΙΑ:" + mEpo + Chr(13) + "ΔΙΕΥΘΥΝΣΗ:" + MDIE + Chr(13) + "ΑΦΜ:" + MAFM
116                 Me.Caption = Textm(IIf(index = 5, 3, 0))
                Else
118                 Textm(IIf(index = 5, 3, 0)) = ""
120                 Me.Caption = ""
                End If
       
            End If
   
        End If

End Sub

Private Sub Textm_KeyPress(index As Integer, KeyAscii As Integer)
    enter_to_tab KeyAscii
End Sub

Private Sub Textm_LostFocus(index As Integer)
           'Dim F_ELAIO_DIK   As Single ' % ΔΙΚΑΙΩΜΑ ΕΛΑΙΟΤΡΙΒΕΙΟΥ ΓΙΑ ΚΟΣΤΟΣ ΕΛΑΙΟΘΛΙΨΗΣ
         If index = 21 And F_ELAIO_DIK > 0 Then
             Grid1.TextMatrix(1, f_p) = F_ELAIO_DIK / 100 * gVal(Textm(21).Text)
           '  Grid1.row = 1: Grid1.Col = f_p
               
           
'            If timtext3.Top = Grid1.CellTop Then
'               If timtext3.Left = Grid1.CellLeft Then
'                  timtext3.Text = Grid1.TextMatrix(1, f_p)
'
'               End If
'
'            End If
            
             
             timtext3.SetFocus
             
         End If
End Sub

Private Sub TextOnEntolea_Change()
        Dim mEpo, MDIE, MAFM
        '2->3  4->0
        'If Index = 5 Or Index = 4 Then
        Dim mmID As Long
          
100     mmID = GET_NVALUE("SELECT ID FROM PEL WHERE ( AFM LIKE '%" + TextOnEntolea + "%' OR  EPO LIKE '%" + TextOnEntolea + "%' ) AND EIDOS='e'")
        ' mEpo = GET_CVALUE("SELECT TOP 1 ISNULL(EPO,'') AS MEPO FROM PEL WHERE ( AFM LIKE '%" + Textm(Index) + "%' OR  EPO LIKE '%" + Textm(Index) + "%' ) AND EIDOS='e'")
        'MDIE = GET_CVALUE("SELECT TOP 1 ISNULL(DIE,'') AS MEPO FROM PEL WHERE EPO LIKE '%" + Textm(Index) + "%' AND EIDOS='e'")
        'MAFM = GET_CVALUE("SELECT TOP 1 ISNULL(AFM,'') AS MEPO FROM PEL WHERE EPO LIKE '%" + Textm(Index) + "%' AND EIDOS='e'")
          
102     If mmID > 0 Then
104         mEpo = GET_CVALUE("SELECT TOP 1 ISNULL(EPO,'') AS MEPO FROM PEL WHERE ID=" + str(mmID))
106         MDIE = GET_CVALUE("SELECT TOP 1 ISNULL(DIE,'') AS MEPO FROM PEL WHERE ID=" + str(mmID))
108         MAFM = GET_CVALUE("SELECT TOP 1 ISNULL(AFM,'') AS MEPO FROM PEL WHERE ID=" + str(mmID))
            
110         TextEntoleas = "ΕΠΩΝΥΜΙΑ:" + mEpo + Chr(13) + "ΔΙΕΥΘΥΝΣΗ:" + MDIE + Chr(13) + "ΑΦΜ:" + MAFM
112         Me.Caption = TextEntoleas
        Else
114         TextEntoleas = ""
116         Me.Caption = ""
        End If
       
End Sub

Private Sub TIMText3_GotFocus()

        '<EhHeader>
        On Error GoTo TIMText3_GotFocus_Err

        '</EhHeader>

100     timtext3.BackColor = vbYellow
102     timtext3.SelStart = 0
104     timtext3.SelLength = Len(timtext3.Text)

        '<EhFooter>
        Exit Sub

TIMText3_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.TIMText3_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.TIMText3_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub timText3_KeyDown(KeyCode As Integer, Shift As Integer)

        'ΕΝΤΡΥ ΤΙΜΟΛΟΓΙΟΥ
        '<EhHeader>
        On Error GoTo timText3_KeyDown_Err

        '</EhHeader>
        Dim DUM

        Dim NNN As Long

100     If KeyCode = 13 Then
102         KeyCode = 0
104         DUM = 0
        End If

106     If KeyCode > 95 Then
108         Me.Caption = KeyCode
110         DUM = 0
        End If

        '<EhFooter>
        Exit Sub

timText3_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.timText3_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.timText3_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub timText3_KeyPress(KeyAscii As Integer)

        'ΕΝΤΡΥ ΤΙΜΟΛΟΓΙΟΥ
        '<EhHeader>
        On Error GoTo timText3_KeyPress_Err

        '</EhHeader>
        Dim DUM

100     If KeyAscii = 13 Then
102         KeyAscii = 0
104         DUM = 0
        End If

106     If KeyAscii = 27 Then
108         KeyAscii = 0
110         DUM = 0
        End If

        '<EhFooter>
        Exit Sub

timText3_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.timText3_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.timText3_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub timText3_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo timText3_KeyUp_Err

        '</EhHeader>

        'ΕΝΤΡΥ ΤΙΜΟΛΟΓΙΟΥ  750   1650  4mhnh
        Dim snc1, snc2, snN1

        Dim Synt5

        Dim RSOK As New ADODB.Recordset

        Dim p, T, e, a, c

        Dim KOMA, MTEXT

        Dim r0 As New ADODB.Recordset

        Dim XX As String

        Dim dd2 As String

        Dim intRecord As Integer

        Dim intField As Integer

        Dim ELEM(200) As String

        Dim DUM

        Dim SQL2 As String

        Dim HDH_FORTOMENA As Single, LK0 As Long, n_HDH_FORTOMENA

100     HDH_FORTOMENA = 0

        Dim NNN

        ' On Error Resume Next

102     If Len(Text2(1).Text) = 0 Then
104         MsgBox "Δεν επιλέξατε Πελάτη / Προμηθευτή"
106         Text2(1).SetFocus

            Exit Sub

        End If

        ' On Error GoTo MHNYMA

108     timtext3.Text = Replace(timtext3.Text, "'", "~")

110     If KeyCode = 46 And Grid1.Col = f_k Then    ' DELETE ΣΒΗΝΩ ΤΗΝ ΣΕΙΡΑ
112         Grid1.TextMatrix(Grid1.row, f_k) = ""
114         Grid1.TextMatrix(Grid1.row, f_o) = ""
116         Grid1.TextMatrix(Grid1.row, f_p) = ""
118         Grid1.TextMatrix(Grid1.row, F_T) = ""
120         Grid1.TextMatrix(Grid1.row, f_e) = ""
122         Grid1.TextMatrix(Grid1.row, f_f) = ""
124         Grid1.TextMatrix(Grid1.row, f_a) = ""
126         Grid1.TextMatrix(Grid1.row, F_mefpa) = ""

128         Grid1.TextMatrix(Grid1.row, F_PROEL) = ""
130         Grid1.TextMatrix(Grid1.row, F_COL_KOLA) = "" 'F_mikta_KILA
132         Grid1.TextMatrix(Grid1.row, F_mikta_KILA) = ""
134         Grid1.TextMatrix(Grid1.row, F_mikta) = ""

136         FindSynolo

            Exit Sub

        End If

138     If KeyCode = 27 Then    'esc τελος εισαγωγής
140         KeyCode = 0
142         Combo4.SetFocus
        End If
 
144     If KeyCode = 116 And Shift = 2 Then  'CONTROL  F5
 
            Dim F34 As New apot21
146         F34.filtroPEL.Caption = "  PELKOD='" + Text2(0).Text + "' AND "
148         timtext3.Visible = True
            ' kartella
150         F34.Text1.Text = Trim(Grid1.TextMatrix(Grid1.row, f_k))

152         F34.Text2.Text = Grid1.TextMatrix(Grid1.row, f_o)
154         F34.Label1.Caption = Left(APOTH1.Text, 1)
156         F34.Label2.Caption = "0"

158         F34.SHOW 1

            '     apot21.Text1.Text = Grid1.TextMatrix(Grid1.row, f_k)
            '    apot21.Text2.Text = Grid1.TextMatrix(Grid1.row, f_o)
            '   apot21.Text3.Text = " Αποθ." + APOTH1.Text
            '  apot21.Label1.Caption = str(Val(Left(Me.Combo9.Text, 2)))
            '  apot21.Label2.Caption = IIf(IsNull(DATA1.Recordset("apos")), "", DATA1.Recordset("apos"))
            'Me.Hide
   
            ' apot21.SHOW 1
    
        End If

160     If KeyCode = 114 And Shift = 0 Then    ' --------------  F3 καταχωρηση SN ----------------------

            'snc1, snc2, snN1
162         If Grid1.TextMatrix(Grid1.row, f_k) Then

                '  On Error Resume Next

164             snc2 = ""

166             Do While True
168                 snc1 = InputBox(snc2 + "δωσε το sn")

170                 If Len(Trim(snc1)) = 0 Then

                        Exit Do

                    Else
172                     a = "insert INTO SN (ARIU,KOD,HME,KODPEL,ATIM) VALUES ("
174                     a = a + "'" + snc1 + "',"
176                     a = a + "'" + Grid1.TextMatrix(Grid1.row, f_k) + "',"
178                     a = a + "'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "',"
180                     a = a + "'" + Text2(0).Text + "',"
182                     a = a + "'" + Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text + "')"
184                     Gdb.Execute a, snN1

186                     If snN1 = 1 Then
188                         snc2 = snc2 + snc1 + Chr(13)
                        Else
190                         MsgBox "δεν κατεχωρήθη"
                        End If
                    End If

                Loop

            End If

            Exit Sub

        End If

192     If KeyCode = 121 And Shift = 0 Then    ' --------------  F10 ----------------------
194         SaveParastat False
196         POLAGOR.SetFocus
198         Text2(1).SetFocus
            Exit Sub

        End If

200     If KeyCode = 118 And Shift = 0 And F_PosostaOTA Then    ' --------------  F7 ------ 4% ota

202         Grid1.TextMatrix(Grid1.row, F_T) = Val(Replace(synolo.Caption, ",", ".")) * F_PosostaOTA / 100
              
        End If

        ' --------------  F6 -BGAZEI GRID ----- xroma-megetos -----------------------------------

204     If KeyCode = 117 And Shift = 0 Then   ' --------------  F6 ------ xroma-megetos
            ' UGridEdit.Label1.Caption = 0
206         UGridEdit.Label1.Caption = Grid1.row

            Dim ColS As Integer, rows As Integer

208         ColS = 0: rows = 0

            'ΦΟΡΤΩΝΩ ΤΙΣ ΕΠΙΚΕΦΑΛΙΔΕΣ ΤΟΥ GRID ΧΡΩΜΑΤΩΝ
            Dim MEGETHOS As String, XROMA As String

210         XROMA = Format(29 + (2 * EID.Recordset("KATHGORIA") - 1), "00")
212         MEGETHOS = Format(29 + (2 * EID.Recordset("KATHGORIA")), "00")

214         If Val(XROMA) = 0 Then Exit Sub
216         If Val(MEGETHOS) = 0 Then Exit Sub

218         r0.Open "SELECT *FROM PINAKES WHERE TYPOS=" + XROMA + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

220         Do While Not r0.EOF
222             UGridEdit.Grid1.TextMatrix(r0("AYJON"), 0) = r0("PERIGRAFH")
                '     Exit Sub
224             rows = rows + 1
226             XROMATA(Grid1.row, rows, 0) = r0("PERIGRAFH")
228             r0.MoveNext
            Loop

230         r0.Close
            'ΦΟΡΤΩΝΩ ΤΙΣ ΕΠΙΚΕΦΑΛΙΔΕΣ ΤΟΥ GRID ΧΡΩΜΑΤΩΝ
232         r0.Open "SELECT *FROM PINAKES WHERE TYPOS=" + MEGETHOS + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

234         Do While Not r0.EOF
236             UGridEdit.Grid1.TextMatrix(0, r0("AYJON")) = r0("PERIGRAFH")
238             ColS = ColS + 1
                ' Exit Sub
240             XROMATA(Grid1.row, 0, ColS) = r0("PERIGRAFH")
242             r0.MoveNext
            Loop

244         r0.Close

246         UGridEdit.Grid1.rows = rows + 1
248         UGridEdit.Grid1.ColS = ColS + 1

250         UGridEdit.lti.Caption = EID.Recordset("LTI")
252         UGridEdit.XTI.Caption = EID.Recordset("XTI")
254         UGridEdit.lti5.Caption = EID.Recordset("LTI5")
256         UGridEdit.kod.Caption = EID.Recordset("KOD")
258         UGridEdit.ono.Caption = EID.Recordset("ONO")
260         UGridEdit.CH1.Caption = EID.Recordset("CH1")
262         UGridEdit.hme.Caption = Format(DTPicker1.Value, "DD/MM/YYYY")

264         UGridEdit.SHOW 1

            '000000027311
            ' 27311

            'ΦΟΡΤΩΝΩ ΤΑ ΧΡΩΜΑΤΑ ΣΕ ΠΙΝΑΚΑ
            Dim k0 As Long, K2 As Long, r2 As Long, c2 As Long

            Dim SUMA_XROMATA As Long

266         SUMA_XROMATA = 0
268         K2 = 1

270         For r2 = 1 To rows
272             For c2 = 1 To ColS
274                 k0 = InStr(K2, gm_str(Grid1.row), ";")

276                 If k0 = 0 Then
278                     XROMATA(Grid1.row, r2, c2) = ""
                    Else
280                     XROMATA(Grid1.row, r2, c2) = mID(gm_str(Grid1.row), K2, k0 - K2)
282                     SUMA_XROMATA = SUMA_XROMATA + Val(XROMATA(Grid1.row, r2, c2))
                    End If

284                 K2 = k0 + 1
                Next
            Next

286         If Grid1.Col = f_p Then    'AN PATOYSE PANV STHN POSOTHTA
288             timtext3.Text = SUMA_XROMATA
            Else
290             Grid1.TextMatrix(Grid1.row, f_p) = SUMA_XROMATA
            End If

292         If SUMA_XROMATA > 0 Then
294             XROMATA(Grid1.row, 0, 0) = -0.01    ' ΣΗΜΑΔΙ ΟΤΙ ΧΡΗΣΙΜΟΠΟΙΗΣΕ ΤΟ F6
            End If

            Exit Sub

        End If

        ' --------------  F6 ------ xroma-megetos -----------------------------------

296     If KeyCode = Val(Right(f_Fkey, 3)) And Shift = 0 Then    ' Val(Left(f_Fkey, 1)) Then ' --------------  F1 ----------------------
298         If f_F1 > -1 Then
300             PARAS.Text = PARAS.List(f_F1)
302             PARAS_LostFocus
304             SaveParastat False

                Exit Sub

            End If
        End If

306     If KeyCode = 123 And Shift = 0 Then  ' --------------  F12 ----------------------
308         If f_F12 > -1 Then
310             PARAS.Text = PARAS.List(f_F12)
312             PARAS_LostFocus
314             SaveParastat False

                Exit Sub

            End If
        End If

316     If KeyCode = 13 Then   'enter ------------------------------------------------
318         timtex3_enter
            '*****************************************

        End If    ' KEYCODE=13 -------------------------------------------------------

320     If KeyCode = 116 And Shift = 0 Then    ' f5 teleytaia timh
322         Me.Caption = 30
324         r0.Open "select TOP 1 TIMM,EKPT  from EGGTIM where PELKOD='" + Text2(0) + "' and KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' ORDER BY HME DESC", Gdb, adOpenForwardOnly, adLockReadOnly

326         If Not r0.EOF Then
                'r0.MoveLast
328             Me.Caption = "ΤΕΛ.ΤΙΜΗ:" + Format(r0(0), "###0.000") + " EΚΤΠΩΣΗ :" + Format(r0(1), "###0.00")
330             Grid1.TextMatrix(Grid1.row, F_T) = Format(r0(0), "###0.000")
332             Grid1.TextMatrix(Grid1.row, f_e) = Format(r0(1), "###0.00")

334             If Grid1.Col = F_T Then
336                 timtext3.Text = Grid1.TextMatrix(Grid1.row, F_T)
                End If

338             If Grid1.Col = f_e Then
340                 timtext3.Text = Grid1.TextMatrix(Grid1.row, f_e)
                End If
                
            Else
                Dim cnpalia As String
342             cnpalia = gConnect + ";DATABASE=" + f_PALIAXRONIA
                Dim cnconn As New ADODB.Connection
344             cnconn.Open cnpalia
                Dim R22 As New ADODB.Recordset
346             R22.Open "select TOP 1 TIMM,EKPT  from EGGTIM where PELKOD='" + Text2(0) + "' and KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' ORDER BY HME DESC", cnconn, adOpenForwardOnly, adLockReadOnly
            
348             If Not R22.EOF Then
                    'r0.MoveLast
350                 Me.Caption = "ΤΕΛ.ΤΙΜΗ:" + Format(R22(0), "###0.00") + " EΚΤΠΩΣΗ :" + Format(R22(1), "###0.00")
352                 Grid1.TextMatrix(Grid1.row, F_T) = Format(R22(0), "###0.00")
354                 Grid1.TextMatrix(Grid1.row, f_e) = Format(R22(1), "###0.00")

356                 If Grid1.Col = F_T Then
358                     timtext3.Text = Grid1.TextMatrix(Grid1.row, F_T)
                    End If

360                 If Grid1.Col = f_e Then
362                     timtext3.Text = Grid1.TextMatrix(Grid1.row, f_e)
                    End If
                End If
364             cnconn.Close
            
            End If
        End If

366     If KeyCode = 38 Then   ' ΠΑΝΩ ΒΕΛΟΣ

368         If Grid1.Text <> timtext3.Text Then  ' ΠΡΟΣΟΧΗ ΑΛΛΑΞΕΣ ΤΟ ΚΕΛΙ ΑΡΑ ΕΝΤΕΡ
               
370             If Grid1.Col = f_k Then  ' AN ΑΛΛΑΞΕ ΤΟΝ ΚΩΔΙΚΟ ΔΕΝ ΤΟΝ ΑΠΟΘΗΚΕΥΕΙ ΣΤΟ GRID
                   
                Else
372                 Grid1.Text = timtext3.Text    ' κρατάει το grid το κείμενο
                End If
               
                'Exit Sub
            Else
374             Grid1.Text = timtext3.Text    ' κρατάει το grid το κείμενο
            End If

            '  Grid1.Text = TIMText3.Text    ' κρατάει το grid το κείμενο

376         If Grid1.Col = F_T Or Grid1.Col = f_p Or Grid1.Col = f_e Then
378             FindSynolo
            End If

            ' πάει μία κολόνα αριστερά
380         If Grid1.Col > 1 Then
382             Grid1.Col = Grid1.Col - 1
            Else

                'π'αει μία σειρά επάνω
384             If Grid1.row > 1 Then
386                 Grid1.row = Grid1.row - 1
388                 Grid1.Col = 1
                End If
            End If

390         Grid1_EnterCell
        
        End If

392     If KeyCode = 120 And Shift = 0 Then    ' f9 ypologizo τιμη χορισ φπα
394         KOMA = 1

            Dim DD As Single, SYNTF As Single
            Dim dd3 As Double
396         dd3 = gVal(Grid1.TextMatrix(Grid1.row, F_T))
398         SYNTF = (100 + g_Fpa(Grid1.TextMatrix(Grid1.row, f_f))) / 100

            Dim CB As Double
400         CB = dd3 / SYNTF 'Round(dd3 / (SYNTF), f_psifiaTIMMON)
402         CB = CB

404         If dd3 > 0 Then
                'timtext3.Text = Format(DD / (100 + g_Fpa(2)) / 100, "###0.000")
406             Grid1.TextMatrix(Grid1.row, F_T) = Replace(Format(CB, "####0.00000"), ",", ".")

            End If

        End If

408     If KeyCode = 119 And Shift = 0 Then    ' f8 ypologizo posotita apo ajia

            Dim m_aj, MR

410         m_aj = InputBox("Δώσε αξία ", str(Grid1.row))
412         m_aj = gVal(m_aj)
414         If m_aj < 0 Then
416             m_aj = 0
            End If
             
            ' 1=YPOLOGIZEI POSOTHTA APO AXIA (BENZINADIKA)  2=YPOLOGIZEI TIMH APO AXIA 3=ekptosi apo axia
418         If F_8 = 1 Then
420             If gVal(Grid1.TextMatrix(Grid1.row, F_T)) > 0 Then
422                 Grid1.TextMatrix(Grid1.row, f_p) = Format(m_aj / gVal(Grid1.TextMatrix(Grid1.row, F_T)) / (1 + g_Fpa(gVal(Grid1.TextMatrix(Grid1.row, f_f))) / 100), "###0.00000")
                End If
                 
                FindSynolo
424         ElseIf F_8 = 2 Then
426             If Val(Grid1.TextMatrix(Grid1.row, f_p)) > 0 Then
428                 Grid1.TextMatrix(Grid1.row, F_T) = Format(Round(gVal(m_aj) / gVal(Grid1.TextMatrix(Grid1.row, f_p)) / (1 - Val(Grid1.TextMatrix(Grid1.row, f_e)) / 100), 3), "###0.00000")
                End If
430         ElseIf F_8 = 4 Then
432             If Val(Grid1.TextMatrix(Grid1.row, f_p)) > 0 Then
                 
                    Dim M_FOROS As Single
434                 If Month(DTPicker1.Value) > 10 Or Month(DTPicker1.Value) < 3 Then
436                     M_FOROS = 0.5 * gVal(Grid1.TextMatrix(Grid1.row, f_p))
                    Else
438                     M_FOROS = 2 * gVal(Grid1.TextMatrix(Grid1.row, f_p))
                    End If
440                 m_aj = (m_aj - M_FOROS) / 1.13
                 
442                 Grid1.TextMatrix(Grid1.row, F_T) = Format(Round(gVal(m_aj) / gVal(Grid1.TextMatrix(Grid1.row, f_p)) / (1 - Val(Grid1.TextMatrix(Grid1.row, f_e)) / 100), 3), "###0.00000")
444                 parakratisiT.Text = M_FOROS
                End If
            End If
        End If

446     If KeyCode = 40 Then   ' KATΩ ΒΕΛΟΣ
448         If Grid1.Text <> timtext3.Text Then  ' ΠΡΟΣΟΧΗ ΑΛΛΑΞΕΣ ΤΟ ΚΕΛΙ ΑΡΑ ΕΝΤΕΡ
               
450             If Grid1.Col = f_k Then  ' AN ΑΛΛΑΞΕ ΤΟΝ ΚΩΔΙΚΟ ΔΕΝ ΤΟΝ ΑΠΟΘΗΚΕΥΕΙ ΣΤΟ GRID
                    ' timtex3_enter
                    'MsgBox "Για αλλαγή κωδικού,πατήστε Enter μετά την αλλαγή"
                Else
452                 Grid1.Text = timtext3.Text    ' κρατάει το grid το κείμενο
                End If
               
                'Exit Sub
            Else
454             Grid1.Text = timtext3.Text    ' κρατάει το grid το κείμενο
            End If
           
456         If Grid1.Col = F_T Or Grid1.Col = f_p Or Grid1.Col = f_e Then
458             FindSynolo
            End If
        
            ' μιά σειρά παρακάτω
460         If Grid1.row < Grid1.rows - 1 Then
462             FindSynolo
464             Grid1.row = Grid1.row + 1
466             Grid1_EnterCell
            End If

        End If

        '  If KeyCode = 27 Then 'esc τελος εισαγωγής
        '       KeyCode = 0
        '       Combo4.SetFocus
        '  End If
        Exit Sub

MHNYMA:
468     HandleError "Par1:timtext3_KeyUp"
        'If Err.Description = "Automation error" Then
        '   Gdb.Close
        '   Gdb.Open gConnect

        'End If

470     Resume Next

RECONNECT:

472     Resume Next

        '<EhFooter>
        Exit Sub

timText3_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.timText3_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.timText3_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'===================================================================================================
'===================================================================================================
Sub timtex3_enter()

        '<EhHeader>
        On Error GoTo timtex3_enter_Err

        '</EhHeader>
        Dim snc1, snc2, snN1

        Dim Synt5

        Dim RSOK As New ADODB.Recordset

        Dim p, T, e, a, c

        Dim KOMA, MTEXT

        Dim r0 As New ADODB.Recordset

        Dim XX As String

        Dim dd2 As String

        Dim intRecord As Integer

        Dim intField As Integer

        Dim ELEM(200) As String

        Dim DUM

        Dim SQL2 As String

        Dim HDH_FORTOMENA As Single, LK0 As Long, n_HDH_FORTOMENA

100     HDH_FORTOMENA = 0

        Dim NNN

        Dim KeyCode As Integer

102     KeyCode = 0

        On Error Resume Next

        'On Error GoTo timtex3_enter_Err

        'CSBmk_CS <If Len(Trim(TIMText3.Text)) = 0>

104     If Len(Trim(timtext3.Text)) = 0 Then   ' enter me keno koytaki  ***************
106         Grid1.Text = timtext3.Text    ' κρατάει το grid το κείμενο

            ' μιά σειρά παρακάτω

108         If Grid1.Col < Grid1.ColS - 4 Then  ' πάει στην επόμενη κολόνα

110             If Grid1.Col + 1 = f_f And f_CanChangeFpa = 0 Then  'FPA
                
                Else
112                 Grid1.Col = Grid1.Col + 1
114                 Grid1_EnterCell
                    Exit Sub
                End If

            End If

        End If    '                                                     ****************

        '   GridEidon.width = 10000
        ' GridEidon.columns(0).Width = 1000
        '  GridEidon.columns(1).Width = 5000
        'GridEidon.Columns(3).width = 1000

        ' GridEidon.Columns(2).width = 1000

        ' kod ή όνομα <2 τότε ψάχνουμε

        '******************************************************************************
         
        ' MEXRI 6-1-13 ΗΤΑΝ ΕΤΣΙ
        '    If Grid1.Col < f_p And (Grid1.Text <> TIMText3.Text Or ( Val(Grid1.TextMatrix(Grid1.row, f_p))) = 0) Then  ' αλλαξε ο κωδικός άρα ψάξε στο αρχείο

        '190      If Grid1.Col < f_p And (Grid1.Text <> timText3.Text Or (gVal(Grid1.TextMatrix(Grid1.row, F_T)) = 0 And gVal(Grid1.TextMatrix(Grid1.row, f_p)) = 0)) Then  ' αλλαξε ο κωδικός άρα ψάξε στο αρχείο

        Dim M_f_AFHNO_IDIA_SEIRA As Boolean

        'CSBmk_CS <If f_AFHNO_IDIA_SEIRA = 1>
116     If f_AFHNO_IDIA_SEIRA = 1 Then
118         M_f_AFHNO_IDIA_SEIRA = Grid1.Col < f_p And (Grid1.Text <> timtext3.Text Or (Val(Grid1.TextMatrix(Grid1.row, f_p))) = 0)
            'CSBmk_CS <Else>
        Else
120         M_f_AFHNO_IDIA_SEIRA = Grid1.Col < f_p And (Grid1.Text <> timtext3.Text Or (gVal(Grid1.TextMatrix(Grid1.row, F_T)) = 0 And gVal(Grid1.TextMatrix(Grid1.row, f_p)) = 0))
            'CSBmk_CS <End If>
        End If

        'CSBmk_CS <If F_KOLA_MIKTA = 1>
122     If F_KOLA_MIKTA = 1 Then ' ΚΟΛΑ ΜΙΚΤΑ
124         M_f_AFHNO_IDIA_SEIRA = Grid1.Col < F_PROEL And (Grid1.Text <> timtext3.Text Or (gVal(Grid1.TextMatrix(Grid1.row, F_T)) = 0 And gVal(Grid1.TextMatrix(Grid1.row, f_p)) = 0))
            'CSBmk_CS <End If>
        End If

        'CSBmk_CS <If M_f_AFHNO_IDIA_SEIRA>
126     If M_f_AFHNO_IDIA_SEIRA Then

            ' ψαχνει κατά κωδικό   %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            'CSBmk_CS <    If Grid1.Col = f_k>
128         If Grid1.Col = f_k Then
130             timtext3.Text = Trim(Left(timtext3.Text, 16))

                'CSBmk_CS <        If Left(TIMText3.Text, 1) = ".">
132             If Left(timtext3.Text, 1) = "." Then    ' $$$$$$$
                    'EID.RecordSource = "select SUM(XRE)-SUM(PIS) AS [ΥΠΟΛ],EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO from EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE ERG  LIKE '" + Trim(Mid(timtext3.Text, 2, 14)) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"

                    ' ψαχνει κατά BOHUHTIKO κωδικό  ERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
134                 EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,CH6 " & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" & " WHERE EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND ERG  LIKE '" + SameLetters(Trim(mID(timtext3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,CH6,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1,NUM2 "

                    'CSBmk_CS <            If F_MONOMEYPOLOIPO = 1 And f_pol = 1>
136                 If F_MONOMEYPOLOIPO = 1 And f_pol = 1 Then
138                     EID.RecordSource = EID.RecordSource + " HAVING (select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END)) >0"
                        'CSBmk_CS <            End If>
                    End If

140                 EID.Refresh

                    'ΑΝ ΔΕΝ ΕΧΕΙ ΚΙΝΗΘΕΙ ΚΑΘΟΛΟΥ ΝΑ ΜΠΟΡΕΙ ΝΑ ΔΕΙΞΕΙ
                    'CSBmk_CS <            If EID.Recordset.EOF>
142                 If EID.Recordset.EOF Then
144                     EID.RecordSource = "SELECT KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,CH6 " & " from EID " & " WHERE  ERG  LIKE '" + SameLetters(Trim(mID(timtext3.Text, 2, 14))) + "%'  "
                        'CSBmk_CS <            End If>
                    End If

                    ' & " WHERE EGGTIM.APOT=" + left(APOTH1.Text, 1) + " AND ERG  LIKE '" + SameLetters(Trim(Mid(TIMText3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"

                    'ειδη που αγοραζει ο πελατης
146             ElseIf Left(timtext3.Text, 1) = "=" Then         ' ' POLAGOR.ListIndex = 0 =>agores  $$$$$$$
                    'EID.RecordSource = "select SUM(XRE)-SUM(PIS) AS [ΥΠΟΛ],EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO from EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE ERG  LIKE '" + Trim(Mid(timtext3.Text, 2, 14)) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"

148                 EID.RecordSource = "select KOD AS [ΚΩΔ], EID.ONO,ROUND(XTI,2) AS XTI ,LTI, " & "(SELECT TOP 1 TIMM*(100-EKPT)/100 FROM EGGTIM WHERE EIDOS='" + fbuff + "' AND PELKOD='" + Text2(0).Text + "' AND KODE=EID.KOD  ORDER BY HME DESC  ) AS LTI5, " & " EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,CH6 " & " from EID WHERE KOD IN (SELECT KODE FROM EGGTIM " & " WHERE EIDOS='" + fbuff + "' AND PELKOD='" + Text2(0).Text + "') ORDER BY ONO "

150                 EID.Refresh

                    ' ψαχνει κατά ΚΩΔΙΚΟΥΣ ΜΟΝΟ ΑΥΤΟΥ ΤΟΥ ΠΡΟΜΗΘΕΥΤΗ %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    'CSBmk_CS <        ElseIf Left(TIMText3.Text, 1) = "$" And >
152             ElseIf Left(timtext3.Text, 1) = "$" And POLAGOR.ListIndex = 0 Then        ' ' POLAGOR.ListIndex = 0 =>agores  $$$$$$$
                    'EID.RecordSource = "select SUM(XRE)-SUM(PIS) AS [ΥΠΟΛ],EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO from EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE ERG  LIKE '" + Trim(Mid(timtext3.Text, 2, 14)) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"

154                 EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,CH6 " & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" & " WHERE EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND LEFT(PROM,6) = '" + Trim(Left(Text2(1).Text, 6)) + "' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,CH6,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1,NUM2 ORDER BY KOD;"

156                 EID.Refresh

                    'ΑΝ ΔΕΝ ΕΧΕΙ ΚΙΝΗΘΕΙ ΚΑΘΟΛΟΥ ΝΑ ΜΠΟΡΕΙ ΝΑ ΔΕΙΞΕΙ
                    'CSBmk_CS <            If EID.Recordset.EOF>
158                 If EID.Recordset.EOF Then
160                     EID.RecordSource = "SELECT KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,CH6 " & " from EID " & " WHERE   LEFT(PROM,6) = '" + Left(Text2(1).Text, 6) + "' ORDER BY KOD"
                        'CSBmk_CS <            End If>
                    End If
                    'timtext3.Text = EID.Recordset("KOD")

                    ' ψαχνει κατά κωδικό  Εργοστασίου  KODERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    'CSBmk_CS <        ElseIf Left(TIMText3.Text, 1) = "+">
162             ElseIf Left(timtext3.Text, 1) = "+" Then    ' $$$$$$$
                    'EID.RecordSource = "select SUM(XRE)-SUM(PIS) AS [ΥΠΟΛ],EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO from EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE ERG  LIKE '" + Trim(Mid(timtext3.Text, 2, 14)) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"

164                 EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,CH6 " & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" & " WHERE EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND KODERG = '" + Trim(mID(timtext3.Text, 2, 14)) + "' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,CH6,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1,NUM2 ;"

166                 EID.Refresh

                    'ΑΝ ΔΕΝ ΕΧΕΙ ΚΙΝΗΘΕΙ ΚΑΘΟΛΟΥ ΝΑ ΜΠΟΡΕΙ ΝΑ ΔΕΙΞΕΙ
                    'CSBmk_CS <            If EID.Recordset.EOF>
168                 If EID.Recordset.EOF Then
170                     EID.RecordSource = "SELECT KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,CH6 " & " from EID " & " WHERE  KODERG = '" + Trim(mID(timtext3.Text, 2, 14)) + "'  ;"
                        'CSBmk_CS <            End If>
                    End If

                    ' & " WHERE EGGTIM.APOT=" + left(APOTH1.Text, 1) + " AND ERG  LIKE '" + SameLetters(Trim(Mid(TIMText3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"

                    '----------------kata kodiko otan einai gyrismeno sta barcode ---------------------------------
                    'CSBmk_CS <        ElseIf f_barcode = 1 And Left(TIMText3.T>
172             ElseIf f_barcode = 1 And Left(timtext3.Text, 1) = "/" Then    ' $$$$$$  KATA BARCODE
174                 EID.RecordSource = "select SUM(CASE WHEN APOT<>" + Left(APOTH1.Text, 1) + " OR  XRE IS NULL THEN 0 ELSE XRE END)" & "-SUM(CASE WHEN  APOT<>" + Left(APOTH1.Text, 1) + " OR  PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,CH6 " & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" & " WHERE EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND   KOD  LIKE '" + SameLetters(Trim(mID(timtext3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,CH6,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1,NUM2;"

176                 EID.Refresh

                    '370                 If Left(f_pelMERC, 3) = "SOK" And (Not EID.Recordset.EOF) Then
                    '                        'Dim RSOK As New ADODB.Recordset
                    '
                    '                        'Dim RSOK As New ADODB.Recordset
                    '                        'Dim RSOK2 As New ADODB.Recordset
                    '
                    '380                     NNN = 0
                    '390                     RSOK.Open "select MEMO from EID WHERE KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
                    '                        ' Text14.Text = RSOK("MEMO")
                    '                        'Text14.Text = RSOK(0)
                    '400                     DD2 = CNull(RSOK(0))
                    '410                     RSOK.Close
                    '
                    '420                     RSOK.Open "select BARCODES.*,(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=30 AND AYJON= CAST(SUBSTRING(BARCODES.PROELEYSH,7,2) AS INT ) ) AS XX from BARCODES  WHERE BARCODES.KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
                    '                        ' NNN = RSOK.RecordCount
                    '
                    '430                     MSFlexGrid1.row = 0
                    '                        'DD2 = Text14.Text
                    '
                    '440                     DUM = FETES_DELIM(DD2, ELEM)
                    '
                    '                        'If NNN > 0 Then
                    '
                    '                        '  DUM = FETES_DELIM(DD2, ELEM)
                    '
                    '450                     Do While Not RSOK.EOF
                    '460                         MSFlexGrid1.row = MSFlexGrid1.row + 1
                    '
                    '470                         MSFlexGrid1.Col = 0
                    '480                         MSFlexGrid1.Text = CNull(RSOK("XX"))
                    '                            ' RSOK.MoveNext
                    '
                    '490                         MSFlexGrid1.Col = 1
                    '500                         MSFlexGrid1.Text = RSOK("ERG")
                    '
                    '510                         MSFlexGrid1.Col = 2
                    '                            'MSFlexGrid1.Text = RSOK("PROELEYSH")
                    '
                    '520                         For NNN = 1 To 30
                    '
                    '530                             If InStr(ELEM(NNN), Left(CNull(RSOK("XX")), 3)) > 0 Then
                    '540                                 MSFlexGrid1.Col = 2
                    '
                    '550                                 If mID(ELEM(NNN), 6, 1) = "/" Then
                    '560                                     MSFlexGrid1.Text = mID(ELEM(NNN), 7, 3)
                    '                                    Else
                    '570                                     MSFlexGrid1.Text = mID(ELEM(NNN), 6, 3)
                    '                                    End If
                    '
                    '                                    Exit For
                    '
                    '                                End If
                    '
                    '                            Next
                    '
                    '580                         RSOK.MoveNext
                    '                        Loop
                    '
                    '590                     RSOK.Close
                    '                        'End If
                    '
                    '                    End If

                    'ΑΝ ΔΕΝ ΕΧΕΙ ΚΙΝΗΘΕΙ ΚΑΘΟΛΟΥ ΝΑ ΜΠΟΡΕΙ ΝΑ ΔΕΙΞΕΙ
                    'CSBmk_CS <            If EID.Recordset.EOF>
178                 If EID.Recordset.EOF Then
180                     EID.RecordSource = "SELECT KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,CH6 " & " from EID " & " WHERE  KOD  LIKE '" + SameLetters(Trim(mID(timtext3.Text, 2, 14))) + "%' ;"
                        'CSBmk_CS <            End If>
                    End If

                    'CSBmk_CS <        ElseIf f_barcode = 1 And Left(TIMText3.T>
182             ElseIf f_barcode = 1 And Left(timtext3.Text, 1) = "%" Then    ' $$$$$$  KATA BARCODE

                    'CSBmk_CS <            If fbuff = "r">
184                 If fbuff = "r" Then    '------------------------

                        Dim GEN As String

                        'CSBmk_CS <                If Len(Dir("C:\DIONISIS.EXE")) > 0>
186                     If Len(Dir("C:\DIONISIS.EXE")) > 0 Then
                            '650                         GEN = "5201001%"
                            '660                         r0.Open "select max(ERG) FROM EID WHERE ERG LIKE '" + GEN + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                            '670                         XX = Format(1 + Val(Left(r0(0), 12)), "000000000000")
                            '680                         r0.Close
                            '690                         APOT1.flag.Caption = "1"
                            '700                         APOT1.Text1(4).Text = XX + ChDigEAN13(XX)    'timtext3.Text
                            '                            gApoMenu = False
                            '                            APOT1.SHOW 1

                            '                         ElseIf Len(Dir("C:\SITELAG.TXT")) > 0 Then
                            '                             GEN = "5221275%"
                            '                             R0.Open "select max(KOD) FROM EID WHERE KOD LIKE '" + GEN + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                            '                             XX = Format(1 + Val(Left(R0(0), 12)), "000000000000")
                            '                             R0.Close
                            '                             apot2scroll.KOD.Text = XX + ChDigEAN13(XX)    'timtext3.Text
                            '                             On Error Resume Next
                            '                             gApoMenu = False
                            '                             apot2scroll.SHOW
                            'CSBmk_CS <                ElseIf f_site = 333>
188                     ElseIf f_site = 333 Then
                            'Dim ANSR As Integer
                            'ANSR = MsgBox("Nα ανοιχθεί νέος κωδικός;", vbYesNo)
                            'If ANSR = vbYes Then
                            ' apot2ROULIS.kod = ""
                            ' apot2ROULIS.SHOW
                            'End If
                                
                            'CSBmk_CS <                Else>
                        Else
190                         GEN = "5221275%"
192                         r0.Open "select max(KOD) FROM EID WHERE KOD LIKE '" + GEN + "'", Gdb, adOpenForwardOnly, adLockReadOnly
194                         XX = Format(1 + Val(Left(r0(0), 12)), "000000000000")
196                         r0.Close
198                         APOT1.flag.Caption = "1"
200                         APOT1.Text1(0).Text = XX + ChDigEAN13(XX)    'timtext3.Text
202                         gApoMenu = False
204                         APOT1.SHOW 1
                            'CSBmk_CS <                End If>
                        End If

                        ' gApoMenu = False
                        ' APOT1.SHOW 1
206                     EID.RecordSource = "select KOD,ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,ERG,CH6 from EID WHERE KOD LIKE '" + XX + "%';"
                        'CSBmk_CS <            End If>
                    End If              '------------------------

                    '----------------kata  barcode ---------------------------------
                    'CSBmk_CS <        ElseIf f_barcode = 1>
208             ElseIf f_barcode = 1 Then    ' $$$$$$  KATA BARCODE

                    'CSBmk_CS <            If F_EXO_XROMATA = 0>
210                 If F_EXO_XROMATA = 0 Then

                        Dim MPOSOTHTA2 As String

212                     MPOSOTHTA2 = "EID.POS"

                        'CSBmk_CS <                If Val(Left(APOTH1.Text, 1)) = 1>
214                     If Val(Left(APOTH1.Text, 1)) = 1 Then MPOSOTHTA2 = "EID.POS01"
                        'CSBmk_CS <                If Val(Left(APOTH1.Text, 1)) = 2>
216                     If Val(Left(APOTH1.Text, 1)) = 2 Then MPOSOTHTA2 = "EID.POS02"
                        'CSBmk_CS <                If Val(Left(APOTH1.Text, 1)) = 3>
218                     If Val(Left(APOTH1.Text, 1)) = 3 Then MPOSOTHTA2 = "EID.POS03"
                        'CSBmk_CS <                If Val(Left(APOTH1.Text, 1)) = 4>
220                     If Val(Left(APOTH1.Text, 1)) = 4 Then MPOSOTHTA2 = "EID.POS04"

                        'CSBmk_CS <                If Len(Trim(F_2CHFIA_ZYGIZOMENA)) > 1 An>
222                     If Len(Trim(F_2CHFIA_ZYGIZOMENA)) > 1 And Left(F_2CHFIA_ZYGIZOMENA, 2) = Left(timtext3.Text, 2) Then
224                         EID.RecordSource = "select " + MPOSOTHTA2 + " AS [ΥΠΟΛ],EID.KOD,EID.ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + SameLetters(Trim(mID(timtext3.Text, 3, 5))) + "%';"
                            'CSBmk_CS <                Else>
                        Else

226                         EID.RecordSource = "select " + MPOSOTHTA2 + " AS [ΥΠΟΛ],EID.KOD,EID.ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + Trim(SameLetters(Trim(mID(timtext3.Text, 1, 15)))) + "%';"
                            'CSBmk_CS <                End If>
                        End If

228                     EID.Refresh

                        'ψαχνω κατα βοηθητικο αν δεν βρεθηκε
                        'CSBmk_CS <                If EID.Recordset.EOF>
230                     If EID.Recordset.EOF Then
                            ' ψαχνει κατά BOHUHTIKO κωδικό  ERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
232                         EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,CH6 " & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" & " WHERE  EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND ERG LIKE '" + SameLetters(Trim(mID(timtext3.Text, 1, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,CH6,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1,NUM2 "
234                         EID.Refresh

                            'CSBmk_CS <                    If F_MONOMEYPOLOIPO = 1 And f_pol = 1>
236                         If F_MONOMEYPOLOIPO = 1 And f_pol = 1 Then
238                             EID.RecordSource = EID.RecordSource + " HAVING [ΥΠΟΛ]>0"
                                'CSBmk_CS <                    End If>
                            End If
                            'CSBmk_CS <                End If>
                        End If

                        'CSBmk_CS <            Else>
                    Else    'XROMATA=1

                        On Error GoTo timtex3_enter_Err

240                     EID.RecordSource = "select top 10 EID.KOD,EID.ONO,LTI5,FPA,MON,XTI,LTI,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX,PROELEYSH,KATHGORIA,CH1,NUM2 from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + Trim(SameLetters(Trim(mID(timtext3.Text, 1, 15)))) + "%';"

                        ' ψαχνει κατά BOHUHTIKO κωδικό  ERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                        '       EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" _
                        '      & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                        '     & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                        '    & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                        '   & " WHERE ERG  LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 1, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 "
                        '  EID.Refresh

242                     EID.Refresh

                        'CSBmk_CS <                If Not EID.Recordset.EOF>
244                     If Not EID.Recordset.EOF Then
246                         XROMATA(Grid1.row, 1, 0) = EID.Recordset("PROELEYSH")
248                         XROMATA(Grid1.row, 0, 1) = Replace(mID(EID.Recordset("PROELEYSH"), 4, 2), "*", "")
                            'CSBmk_CS <                End If>
                        End If

                        'CSBmk_CS <                If Left(f_pelMERC, 3) = "SOK" And (Not E>
250                     If Left(f_pelMERC, 3) = "SOK" And (Not EID.Recordset.EOF) Then
252                         NNN = 0
                            '                       RSOK.Open "select MEMO from EID WHERE KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
254                         dd2 = CNull(EID.Recordset("MEMO"))

                            '                                              RSOK.Close
256                         RSOK.Open "select BARCODES.*,(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=30 AND AYJON= CAST(SUBSTRING(BARCODES.PROELEYSH,7,2) AS INT ) ) AS XX from BARCODES  WHERE BARCODES.KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
258                         MSFlexGrid1.row = 0
260                         DUM = FETES_DELIM(dd2, ELEM)

                            'CSBmk_CS <                    Do While Not RSOK.EOF>
262                         Do While Not RSOK.EOF
264                             MSFlexGrid1.row = MSFlexGrid1.row + 1

266                             MSFlexGrid1.Col = 0
268                             MSFlexGrid1.Text = CNull(RSOK("XX"))
                                ' RSOK.MoveNext

270                             MSFlexGrid1.Col = 1
272                             MSFlexGrid1.Text = RSOK("ERG")

274                             MSFlexGrid1.Col = 2

                                'CSBmk_CS <                        For NNN = 1 To 30>
276                             For NNN = 1 To 30

                                    'CSBmk_CS <                            If InStr(ELEM(NNN), Left(CNull(RSOK("XX">
278                                 If InStr(ELEM(NNN), Left(CNull(RSOK("XX")), 3)) > 0 Then
280                                     MSFlexGrid1.Col = 2

                                        'CSBmk_CS <                                If mID(ELEM(NNN), 6, 1) = "/">
282                                     If mID(ELEM(NNN), 6, 1) = "/" Then
284                                         MSFlexGrid1.Text = mID(ELEM(NNN), 7, 3)
                                            'CSBmk_CS <                                Else>
                                        Else
286                                         MSFlexGrid1.Text = mID(ELEM(NNN), 6, 3)
                                            'CSBmk_CS <                                End If>
                                        End If

                                        Exit For

                                        'CSBmk_CS <                            End If>
                                    End If

                                    'CSBmk_CS <                        Next>
                                Next

288                             RSOK.MoveNext
                                'CSBmk_CS <                    Loop>
                            Loop

290                         RSOK.Close
                            'CSBmk_CS <                End If>
                        End If

                        'CSBmk_CS <            End If>
                    End If

                    '----------------kata  barcode  οταν είναι γυρισμένο κατα κωδικό ---------------------------------
                    'CSBmk_CS <        ElseIf Left(TIMText3.Text, 1) = "/">
292             ElseIf Left(timtext3.Text, 1) = "/" Then    ' $$$$$$  KATA BARCODE

                    Dim MBB As String
                    ' ΝΑ ΑΓΝΟΕΙ ΤΟ 0 ΜΠΡΟΣΤΑ
294                 If Left(timtext3.Text, 1) = "0" Then
296                     MBB = Trim(mID(timtext3.Text, 3, 14))
                    Else
298                     MBB = Trim(mID(timtext3.Text, 2, 14))
                    End If

300                 EID.RecordSource = "select EID.KOD,EID.ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX,EID.KATHGORIA,CH1,NUM2 from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + SameLetters(MBB) + "%';"

                    '-------------------ΓΕΝΝΗΤΡΙΑ ΚΩΔΙΚΩΝ -------------------------------------------------
                    'CSBmk_CS <        ElseIf Left(TIMText3.Text, 1) = "%">
302             ElseIf Left(timtext3.Text, 1) = "%" Then   '$$$$$$$

                    'CSBmk_CS <            If fbuff = "r">
304                 If fbuff = "r" Then    '------------------------

                        Dim GEN2 As String

                        'CSBmk_CS <                If Len(Dir("C:\DIONISIS.EXE")) > 0>
306                     If Len(Dir("C:\DIONISIS.EXE")) > 0 Then
308                         GEN2 = "5201001%"
310                         r0.Open "select max(ERG) FROM EID WHERE ERG LIKE '" + GEN2 + "'", Gdb, adOpenForwardOnly, adLockReadOnly
312                         XX = Format(1 + Val(Left(r0(0), 12)), "000000000000")
314                         r0.Close
316                         APOT1.flag.Caption = "1"
318                         APOT1.Text1(4).Text = XX + ChDigEAN13(XX)    'timtext3.Text
                              
320                         gApoMenu = False
322                         APOT1.SHOW 1
                            '                          ElseIf Len(Dir("C:\SITELAG.TXT")) > 0 Then
                            '                             GEN = "5221275%"
                            '                             R0.Open "select max(KOD) FROM EID WHERE KOD LIKE '" + GEN + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                            '                             XX = Format(1 + Val(Left(R0(0), 12)), "000000000000")
                            '                             R0.Close
                            '                             apot2scroll.KOD.Text = XX + ChDigEAN13(XX)    'timtext3.Text
                            '                             On Error Resume Next
                            '                             gApoMenu = False
                            '                             apot2scroll.SHOW
                            'CSBmk_CS <                Else>
                        Else
324                         GEN2 = "5221275%"
326                         r0.Open "select max(KOD) FROM EID WHERE KOD LIKE '" + GEN2 + "'", Gdb, adOpenForwardOnly, adLockReadOnly
328                         XX = Format(1 + Val(Left(r0(0), 12)), "000000000000")
330                         r0.Close
332                         APOT1.flag.Caption = "1"
334                         APOT1.Text1(0).Text = XX + ChDigEAN13(XX)    'timtext3.Text
336                         gApoMenu = False
338                         APOT1.SHOW 1
                          
                            'CSBmk_CS <                End If>
                        End If

                        '1390                      gApoMenu = False
                        '1400                      APOT1.SHOW 1
340                     timtext3.SetFocus
342                     EID.RecordSource = "select KOD,ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,ERG,CH6 from EID WHERE KOD LIKE '" + XX + "%';"
                        'CSBmk_CS <            End If>
                    End If              '------------------------

                    'CSBmk_CS <        Else>
                Else   ' $$$$$$$

                    '----KATA KΩΔIKO ΓΥΡΙΣΜΕΝΟ ΣΤΟ ΚΑΤΑ ΚΩΔΙΚΟ

                    'εαν ειναι ρυθμισμένο να δείχνει όλα τα είδη ή είναι αγορά (οπότε πρέπει να τα δείχνει όλα)
                    '                 If F_YPOL_GEN = 1 Or PEL.Recordset("EIDOS") = "r" Then
                    '                     If Val(Left(APOTH1.Text, 2)) = 1 Then
                    '                         C = "POS01"
                    '                     ElseIf Val(Left(APOTH1.Text, 2)) = 2 Then
                    '                          C = "POS02"
                    '                     ElseIf Val(Left(APOTH1.Text, 2)) = 3 Then
                    '                          C = "POS03"
                    '                     ElseIf Val(Left(APOTH1.Text, 2)) = 4 Then
                    '                          C = "POS04"
                    '                     Else
                    '                          C = "POS"
                    '                     End If
                    '
                    '                     EID.RecordSource = "select " + C + " AS [ΥΠΟΛ], KOD AS [ΚΩΔ]," _
                    '                     & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                    '                     & " from EID " _
                    '                     & " WHERE  ONO  LIKE '" + SameLetters(TIMText3.Text) + "%' ;"
                    '                     EID.Refresh
                    '                 Else
                    'CSBmk_CS <            If F_YPOL_GEN = 1 Or PEL.Recordset("EIDO>
344                 If F_YPOL_GEN = 1 Or PEL.Recordset("EIDOS") = "r" Then
346                     Synt5 = ""
                        'CSBmk_CS <            Else>
                    Else
348                     Synt5 = " EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND "
                        'CSBmk_CS <            End If>
                    End If

                    'CSBmk_CS <            If Len(Trim(F_2CHFIA_ZYGIZOMENA)) > 1 An>
350                 If Len(Trim(F_2CHFIA_ZYGIZOMENA)) > 1 And Left(F_2CHFIA_ZYGIZOMENA, 2) = Left(timtext3.Text, 2) Then
                        'EID.RecordSource = "select EID.KOD,EID.ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 3, 5))) + "%';"
352                     EID.RecordSource = "select KOD,ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,ERG,SYSKMAX,LTI2,LTI3,CH6 from EID WHERE ENERGO=1 AND LEFT(KOD,7)='" + mID(timtext3.Text, 1, 7) + "';"
                        'CSBmk_CS <            Else>
                    Else
354                     If f_AKRIBOS_TON_KODIKO = 1 Then
356                         EID.RecordSource = "select SUM(CASE WHEN APOT<>" + Left(APOTH1.Text, 1) + " OR  XRE IS NULL THEN 0 ELSE XRE END)" & "-SUM(CASE WHEN  APOT<>" + Left(APOTH1.Text, 1) + " OR  PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],EID.ONO,KOD AS [ΚΩΔ],XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3,UES,CH6 " & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" & " WHERE " + Synt5 + " ENERGO=1 AND  KOD  ='" + SameLetters(timtext3.Text) + "' GROUP BY ONO,KOD,XTI,LTI,LTI5,CH6,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3,UES ORDER BY KOD ;"
                        Else
358                         EID.RecordSource = "select SUM(CASE WHEN APOT<>" + Left(APOTH1.Text, 1) + " OR  XRE IS NULL THEN 0 ELSE XRE END)" & "-SUM(CASE WHEN  APOT<>" + Left(APOTH1.Text, 1) + " OR  PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ], EID.ONO,KOD AS [ΚΩΔ],XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3,UES,CH6 " & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" & " WHERE " + Synt5 + " ENERGO=1 AND  KOD  LIKE '" + SameLetters(timtext3.Text) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,CH6,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3,UES ORDER BY KOD ;"
                        End If
                        'CSBmk_CS <            End If>
                    End If
                    
360                 EID.Refresh
                    ' End If

                    'CSBmk_CS <            If EID.Recordset.EOF>
362                 If EID.Recordset.EOF Then
364                     EID.RecordSource = "select KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3,UES,CH6 " & " from EID " & " WHERE ENERGO=1 AND  KOD  LIKE '" + SameLetters(timtext3.Text) + "%' ;"
366                     If EID.Recordset.EOF Then
                        
368                         EID.RecordSource = "select KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3,UES,CH6 " & " from EID " & " WHERE ENERGO=1 AND ( ERG='" + SameLetters(timtext3.Text) + "' OR KOD IN (SELECT TOP 1 KOD FROM BARCODES WHERE ERG='" + timtext3.Text + "')) ;"

                        End If

                        'CSBmk_CS <            End If>
                    End If

                    'CSBmk_CS <        End If>
                End If   ' $$$$$$$

370             EID.Refresh

                Dim RR2() As String
   
11266           RR2 = Split(F_PLATOSGRIDEIDON, ";")
          
11268           If UBound(RR2) = 5 Then
                    Dim LLC As Integer
11270               For LLC = 0 To 4
11272                   GridEidon.columns(LLC).Width = RR2(LLC)
                    Next
                Else

372                 GridEidon.columns(0).Width = 500
374                 GridEidon.columns(1).Width = 5000
376                 GridEidon.columns(2).Width = 1000

378                 GridEidon.SetFocus

                End If
                'ψάχνει κατά όνομα
                'CSBmk_CS <    ElseIf Grid1.Col = f_o>
380         ElseIf Grid1.Col = f_o Then  '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

382             GridEidon.columns(0).Width = 500
384             GridEidon.columns(1).Width = 5000

                'ME * NA MHN PSAXNEI KATA ONOMA
                'CSBmk_CS <        If Left(Grid1.TextMatrix(Grid1.row, f_k)>
                '                Dim mSym As String
                '                If gAsterisk = 1 Then
                '                    mSym = "@"
                '                Else
                '                    mSym = "*"
                '                End If
                '
386             If Left(Grid1.TextMatrix(Grid1.row, f_k), 1) = "*" Then    ' ----------
388                 Grid1.Text = timtext3.Text

                    ' μιά κολονα παρακάτω
                    'CSBmk_CS <            If Grid1.Col < Grid1.ColS - 3>
390                 If Grid1.Col < Grid1.ColS - 3 Then  ' πάει στην επόμενη κολόνα
392                     Grid1.Col = Grid1.Col + 1
394                     Grid1_EnterCell

                        Exit Sub

                        'CSBmk_CS <            End If>
                    End If

                    'CSBmk_CS <        Else>
                Else  ' -------------

                    ' SELECT SUM(XRE-PIS),KODE,EID.ONO FROM  EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE KODE LIKE '01%'  GROUP BY KODE,EID.ONO
                    ' EID.RecordSource = "select SUM(XRE-PIS) AS [ΥΠΟΛ],EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO from EID LEFT JOIN EGGTIM ON EGGTIM.KODE=EID.KOD WHERE EID.ONO LIKE '" + UCase(timtext3.Text) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"
                    Dim ONOM As String

                    Dim KODI As String

                    Dim DOL As Long

                    Dim mpos

                    'SYNDYASMENH ANAZHTHSH ME $  DHLADH   ΣΚΙΑ$99 ΟΛΕΣ ΤΙΣ "ΣΚΙΕΣ" ΜΕ ΚΩΔΙΚΟ ΠΟΥ ΑΡΧΙΖΕΙ ΑΠΟ "99"
396                 ONOM = timtext3.Text

398                 DOL = InStr(1, ONOM, "$")
400                 KODI = ""

                    'CSBmk_CS <            If DOL > 1>
402                 If DOL > 1 Then
404                     KODI = Trim(mID(ONOM, DOL + 1, 14))
406                     ONOM = Left(ONOM, DOL - 1)

                        '--------------- //////////////////
408                     EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ]," & " EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3,CH6 " & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" & " WHERE  EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND EID.ENERGO=1 AND  EID.ONO  LIKE '" + SameLetters(ONOM) + "%' AND KOD LIKE '" + KODI + "%' and EGGTIM.APOT=" + Left(APOTH1.Text, 1) + "  GROUP BY ONO,KOD,XTI,LTI,CH6,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3 ;"
                        'CSBmk_CS <            Else>
                    Else

                        'εαν ειναι ρυθμισμένο να δείχνει όλα τα είδη ή είναι αγορά (οπότε πρέπει να τα δείχνει όλα)
                        '                 If F_YPOL_GEN = 1 Or PEL.Recordset("EIDOS") = "r" Then
                        '                     If Val(Left(APOTH1.Text, 2)) = 1 Then
                        '                         MpOS = "POS01"
                        '                     ElseIf Val(Left(APOTH1.Text, 2)) = 2 Then
                        '                          MpOS = "POS02"
                        '                     ElseIf Val(Left(APOTH1.Text, 2)) = 3 Then
                        '                          MpOS = "POS03"
                        '                     ElseIf Val(Left(APOTH1.Text, 2)) = 4 Then
                        '                          MpOS = "POS04"
                        '                     Else
                        '                          MpOS = "POS"
                        '                     End If
                        '
                        '                     EID.RecordSource = "select " + MpOS + " AS [ΥΠΟΛ], KOD AS [ΚΩΔ]," _
                        '                     & " EID.ONO,LTI5,EID.FPA,MON,XTI,LTI,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                        '                     & " from EID " _
                        '                     & " WHERE  ONO  LIKE '" + SameLetters(TIMText3.Text) + "%' ;"
                        '                 Else

                        'ΠΡΟΣΟΧΗ ΑΝ ΕΧΩ
                        ' WHERE  EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND  EID.ONO  LIKE '" + SameLetters(ONOM) + "%'"
                        'ΒΓΑΖΕΙ ΜΟΝΟ ΤΑ ΕΙΔΗ ΠΟΥ ΕΧΟΥΝ ΚΙΝΗΘΕΙ ΣΤΗΝ ΑΠΟΘΗΚΗ 1 Π.Χ.
                        'ΕΝΩ ΤΟ
                        ' WHERE EID.ONO  LIKE '" + SameLetters(ONOM) + "%'"
                        ' ΒΓΑΖΕΙ ΟΛΑ ΤΑ ΕΙΔΗ,ΚΑΙ ΑΥΤΑ ΠΟΥ ΔΕΝ ΕΧΟΥΝ ΚΙΝΗΘΕΙ
                        'CSBmk_CS <                If F_YPOL_GEN = 1 Or PEL.Recordset("EIDO>
410                     If F_YPOL_GEN = 1 Or PEL.Recordset("EIDOS") = "r" Then
412                         Synt5 = ""
                            'CSBmk_CS <                Else>
                        Else
414                         Synt5 = " EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND "
                            'CSBmk_CS <                End If>
                        End If
                        '--------------- //////////////////
416                     EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE XRE END)" & "-SUM(CASE WHEN PIS IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE PIS END) AS [ΥΠΟΛ]," & " EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3,CH6 " & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" & " WHERE " + Synt5 + " EID.ENERGO=1 AND EID.ONO  LIKE '" + SameLetters(ONOM) + "%'" + "  GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3,CH6 "
                        ' and EGGTIM.APOT=" + left(APOTH1.Text, 1)

                        'CSBmk_CS <                If F_MONOMEYPOLOIPO = 1 And f_pol = 1>
418                     If F_MONOMEYPOLOIPO = 1 And f_pol = 1 Then
420                         EID.RecordSource = EID.RecordSource + " HAVING SUM(CASE WHEN XRE IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE XRE END)" & "-SUM(CASE WHEN PIS IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE PIS END)>0"
                            'CSBmk_CS <                End If>
                        End If

                        'End If

                        'CSBmk_CS <            End If>
                    End If

422                 EID.Refresh

                    Dim RR() As String

1266                RR = Split(F_PLATOSGRIDEIDON, ";")
          
1268                If UBound(RR) = 5 Then
                        Dim LLC2 As Integer
1270                    For LLC2 = 0 To 4
1272                        GridEidon.columns(LLC2).Width = RR(LLC2)
                        Next
                    Else

424                     GridEidon.columns(0).Width = 1000
426                     GridEidon.columns(1).Width = 5000
428                     GridEidon.columns(3).Width = 1000

430                     GridEidon.columns(2).Width = 1000
                    End If
                    'CSBmk_CS <            If EID.Recordset.EOF>
432                 If EID.Recordset.EOF Then
                        '--------------- //////////////////
434                     EID.RecordSource = "select EID.ONO,KOD,POS AS [ΣΥΝ.ΥΠΟΛ], XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,LTI2,LTI3,UES,CH6 " & " from EID " & " WHERE EID.ENERGO=1 AND   EID.ONO  LIKE '" + SameLetters(timtext3.Text) + "%' ORDER BY ONO ;"
436                     EID.Refresh

                        'CSBmk_CS <            End If>
                    End If

                    'CSBmk_CS <            If EID.Recordset.RecordCount = 0>
438                 If EID.Recordset.RecordCount = 0 Then   ' $$$$$$$$$$$$$$$$$$$$$$$$$$

                        ' μιά κολονα παρακάτω
                        'CSBmk_CS <                If Grid1.Col < Grid1.ColS - 3>
440                     If Grid1.Col < Grid1.ColS - 3 Then  ' πάει στην επόμενη κολόνα
442                         Grid1.Col = Grid1.Col + 1
444                         Grid1_EnterCell

                            Exit Sub

                            'CSBmk_CS <                End If>
                        End If
                        'CSBmk_CS <            End If>
                    End If   ' $$$$$$$$$$$$$$$$$$$$$$$$$$

                    'CSBmk_CS <        End If>
                End If    '----------

                'CSBmk_CS <    End If>
            End If   '        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

            'CSBmk_CS <    If EID.Recordset.EOF>
446         If EID.Recordset.EOF Then   '%%%%%%%%%%%%%

                'CSBmk_CS <        If fbuff = "r">
448             If fbuff = "r" Then  ' $$$$$$$$$$$$$$$
                    ' If f_site = 333 Then
                      
                    'CSBmk_CS <            If f_site = 333>
450                 If f_site = 333 Then

                        Dim ANSR As Integer

452                     ANSR = MsgBox("Nα ανοιχθεί νέος κωδικός;", vbYesNo)

                        'CSBmk_CS <                If ANSR = vbYes>
454                     If ANSR = vbYes Then
456                         Grid1.TextMatrix(Grid1.row, f_k) = timtext3.Text
                            'apot2ROULIS.kod = TIMText3.Text
                            'apot2ROULIS.SHOW
                            'apot2ROULIS.NEO_EIDOS
                            'TIMText3.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, f_k)
                            'MSFlexGrid1.TextMatrix(MSFlexGrid1.row + 1, f_k) = TIMText3.Text
                            ' TIMText3.SetFocus
                            'CSBmk_CS <                End If>
                        End If

                        'CSBmk_CS <            Else>
                    Else
                    
458                     If GGET_NVALUE("SELECT COUNT(*) FROM EID WHERE KOD='" + timtext3.Text + "' AND ISNULL(ENERGO,0)=0 ") > 0 Then
460                         If MsgBox("Υπάρχει το είδος σαν ανενεργό." + Chr(13) + "Να γίνει ενεργό;", vbYesNo) = vbYes Then
462                             Gdb.Execute "update EID set ENERGO=1 WHERE KOD='" + timtext3.Text + "'"
                            Else
464                             timtext3.Text = ""
                            End If
                        
                        Else
                    
466                         APOT1.Text1(0).Text = timtext3.Text
468                         APOT1.flag.Caption = "1"
470                         APOT1.SHOW 1
                        End If
472                     timtext3.SetFocus
                        'CSBmk_CS <            End If>
                    End If
                  
                    'CSBmk_CS <        Else>
                Else   ' $$$$$$$$$$$

                    ' If EID.Recordset.RecordCount = 0 Then
                    ' μιά κολονα παρακάτω
                    'CSBmk_CS <            If Grid1.Col < Grid1.ColS - 3>
474                 If Grid1.Col < Grid1.ColS - 3 Then  ' πάει στην επόμενη κολόνα
476                     Grid1.Col = Grid1.Col + 1
478                     Grid1_EnterCell

                        Exit Sub

                        'CSBmk_CS <            End If>
                    End If
                    'CSBmk_CS <        End If>
                End If   ' $$$$$$$$$$$

                'CSBmk_CS <    Else>
            Else  ' %%%%%%%%%%%%%%%%%%

480             GridEidon.SetFocus

                'CSBmk_CS <        If EID.Recordset.RecordCount = 1>
482             If EID.Recordset.RecordCount = 1 Then
                    'CSBmk_CS <            If f_AutoChangeLine = 1>
484                 If f_AutoChangeLine = 1 Then
486                     GridEidon.SetFocus
488                     GridEidon_KeyUp 13, 0
490                     SendKeys "{enter}{enter}"
                        'CSBmk_CS <            Else>
                    Else
492                     GridEidon.SetFocus
494                     GridEidon_KeyUp 13, 0

                        'CSBmk_CS <            End If>
                    End If
                    'CSBmk_CS <        End If>
                End If

496             Me.Caption = "127"
                'CSBmk_CS <    End If>
            End If  ' %%%%%%%%%%%%%%%%%

498         intRecord = Min(10, EID.Recordset.RecordCount)
500         intField = EID.Recordset.FIELDS.Count

            'call the procedure here...
            'CSBmk_CS <    If intRecord > 1>
502         If intRecord > 1 Then
504             'Call AdjustDataGridColumns(intRecord, intField, True)
506             EID.Recordset.MoveFirst
508             GridEidon.SetFocus

                'CSBmk_CS <    End If>
            End If

            '          GridEidon.SetFocus

            ' kod ή όνομα <2 τότε ψάχνουμε
            '****************************
            'CSBmk_CS <Else>
        Else    ' *******************************************************************************************************************************************************
            '*********************************************************************************************************************************************************
            
            'AJ=T*P*(100-EKPT)/100
            'CSBmk_CS <    If Grid1.Col = f_a>
            
510         If Grid1.Col = f_a Then
                  
                'CSBmk_CS <        If Grid1.TextMatrix(Grid1.row, f_p) > 0>
                ' AN EXV 2 EKPTVSEIS ΧΑΛΑΕΙ ΤΗΝ ΤΙΜΗ
                
                ' EMFANIZEI PROBLHMATA 16-5-22 GIAYTO EKANA REM TO EPOMENO IF -ENDIF   14 SEIRES
512             If Grid1.TextMatrix(Grid1.row, f_p) > 0 Then
514                 Dim M_EKP1 As Double: M_EKP1 = (100 - gVal(Grid1.TextMatrix(Grid1.row, f_e))) / 100
516                 Dim M_EKP2 As Double: M_EKP2 = (100 - gVal(Grid1.TextMatrix(Grid1.row, F_EKPT2))) / 100
518                 Dim m_arxiki_timi As Double: m_arxiki_timi = gVal(Grid1.TextMatrix(Grid1.row, F_T))

                    ' εαν η αρχική (υπολόγιζόμενη αξία) - την αξια που εχω πληκτρολογήσει(εαν τυχόν πληκτρολόγησα) ειναι > 0.005 τοτε ξαναυπολογίζω τιμη
520                 If Abs(Round(gVal(Grid1.TextMatrix(Grid1.row, f_p) * M_EKP1 * M_EKP2) * gVal(Grid1.TextMatrix(Grid1.row, F_T)), 2) - gVal(timtext3.Text)) > 0.009 Then
                        ' '''''If Abs(m_arxiki_timi - (gVal(TIMText3.Text) / (gVal(Grid1.TextMatrix(Grid1.row, f_p)) * M_EKP1 * M_EKP2))) > 0.005 Then
522                     Grid1.TextMatrix(Grid1.row, F_T) = Format(gVal(timtext3.Text) / (gVal(Grid1.TextMatrix(Grid1.row, f_p)) * M_EKP1 * M_EKP2), "###0.00000")
                    End If

524                 FindSynolo
                    'CSBmk_CS <        End If>
                End If
            
                'CSBmk_CS <    End If>
            End If
            
            ' EKTOS KOD ,ONO
            
            'CSBmk_CS <    If Grid1.Col = f_f>
526         If Grid1.Col = f_f Then   'φπα ελεγχος για να μην βαζει ασχετες τιμες
                'CSBmk_CS <        If Val(TIMText3.Text) <= 0 Or Val(TIMTex>
528             If Val(timtext3.Text) <= 0 Or Val(timtext3.Text) > 9 Then
530                 MsgBox "Λαθος ΦΠΑ . Πρέπει να είναι απο 1-9"
532                 timtext3.Text = Grid1.TextMatrix(Grid1.row, f_f)
                    'CSBmk_CS <        End If>
                End If
                'CSBmk_CS <    End If>
            End If

            'ΥΠΟΛΟΓΊΖΕΙ ΑΞΊΑ

            '' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            'CSBmk_CS <    If Grid1.Col = f_p>
534         If Grid1.Col = f_p Then

                Dim R As New ADODB.Recordset, mXre As Single, mPis As Single, mDesmia As Single

                'CSBmk_CS <        If f_use_Sysk = 1>
536             If f_use_Sysk = 1 Then
                    
                    'ΠΑΥΛΟΣ ΝΕΧΤ 2 ΛΙΝΕΣ
538                 EID.RecordSource = "SELECT KOD AS [ΚΩΔ]," & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1,NUM2,LTI2,LTI3,CH6 " & " from EID " & " WHERE  KOD= '" + Grid1.TextMatrix(Grid1.row, f_k) + "'"

540                 EID.Refresh
                
                    'CSBmk_CS <            If EID.Recordset("SYSKMAX") > 0 And SYSK>
542                 If EID.Recordset("SYSKMAX") > 0 And SYSKEYASIA.Value = vbChecked Then
544                     timtext3.Text = timtext3.Text * EID.Recordset("SYSKMAX")
                        'CSBmk_CS <            End If>
                    End If
                    'CSBmk_CS <        End If>
                End If

                'CSBmk_CS <        If (fbuff = "e" And (F_POS_APOU = 2 Or F>
546             If (fbuff = "e" And (F_POS_APOU = 2 Or F_POS_APOU = 3) And f_checkYpol > 0 And F_STADIO <> 2) Or (F_PAR_MONO_ME_YPOLOIPO = 1 And F_STADIO = 2) Then    ' F_STADIO<>2 ΔΕΝ ΕΙΝΑΙ ΠΑΡΑΓΓΕΛΙΑ
                    ' 13-4-2006 R.Open "SELECT SUM(XRE),SUM(PIS) FROM EGGTIM WHERE KODE='" + EID.Recordset("KOD") + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
548                 R.Open "SELECT SUM(XRE),SUM(PIS) FROM EGGTIM WHERE  APOT=" + Left(APOTH1.Text, 1) + " and KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' ", Gdb, adOpenForwardOnly, adLockReadOnly

                    'CSBmk_CS <            If R.EOF>
550                 If R.EOF Then
                        'CSBmk_CS <            Else>
                    Else
552                     mXre = IIf(IsNull(R(0)), 0, R(0)): mPis = IIf(IsNull(R(1)), 0, R(1))
                        'CSBmk_CS <            End If>
                    End If

554                 R.Close

                    'ποσοτητα παραγγελιών για δέσμευση

                    'CSBmk_CS <            If IsEmpty(F_SHMADEMENA_DELTIA) Or Len(T>
556                 If IsEmpty(F_SHMADEMENA_DELTIA) Or Len(Trim(F_SHMADEMENA_DELTIA)) <= 2 Then
558                     F_SHMADEMENA_DELTIA = "''"
                        'ΣΑΝ ΔΕΣΜΕΥΜΕΝΑ ΝΟΟΥΝΤΑΙ ΟΛΑ ΤΩΝ ΠΑΡΑΓΓΕΛΙΩΝ ΠΛΗΝ ΑΥΤΟΥ ΤΟΥ ΠΕΛΑΤΗ
560                     SQL2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" & " (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "'  and PELKOD<>'" + Text2(0).Text + "'"
                        'CSBmk_CS <            Else>
                    Else
                        'ΣΑΝ ΔΕΣΜΕΥΜΕΝΑ ΝΟΟΥΝΤΑΙ ΟΛΑ ΤΩΝ ΠΑΡΑΓΓΕΛΙΩΝ ΠΛΗΝ ΤΩΝ ΠΑΡΑΣΤΑΤΙΚΩΝ ΠΟΥ ΤΣΙΜΠΗΣΑ
562                     SQL2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" & " (   NOT EGGTIM.ATIM+CONVERT(CHAR(10),EGGTIM.HME,103)   IN (" + F_SHMADEMENA_DELTIA + ") ) AND (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' "
                        'CSBmk_CS <            End If>
                    End If

564                 R.Open SQL2, Gdb, adOpenForwardOnly, adLockReadOnly

                    'CSBmk_CS <            If R.EOF Or f_Desmia = 0>
566                 If R.EOF Or f_Desmia = 0 Then    ' ΔΕΝ ΒΡΗΚΕ ή ΔΕΝ ΚΑΝΕΙ ΕΛΕΓΧΟ
568                     mDesmia = 0
                        'CSBmk_CS <            Else>
                    Else
570                     mDesmia = IIf(IsNull(R(0)), 0, R(0))
                        'CSBmk_CS <            End If>
                    End If

                    'ελεγχος αν αυτά που στέλνω είναι μέσα σε αυτά που παρήγγειλε o πελάτης
                    'CSBmk_CS <            If f_mono_apoParaggelia = 1 And F_STADIO>
572                 If f_mono_apoParaggelia = 1 And F_STADIO <> 2 Then    'ΔΕΝ ΕΙΝΑΙ ΠΑΡΑΓΓΕΛΙΑ

                        '  On Error GoTo MHNYMA

574                     HDH_FORTOMENA = 0

                        'CSBmk_CS <                For LK0 = 1 To Grid1.ROWS - 1>
576                     For LK0 = 1 To Grid1.rows - 1

                            'CSBmk_CS <                    If LK0 <> Grid1.row And Grid1.TextMatrix>
578                         If LK0 <> Grid1.row And Grid1.TextMatrix(LK0, 1) = Grid1.TextMatrix(Grid1.row, 1) Then
580                             HDH_FORTOMENA = HDH_FORTOMENA + Grid1.TextMatrix(LK0, 3)
                                'CSBmk_CS <                    End If>
                            End If

                            'CSBmk_CS <                Next>
                        Next

                        '    Sql2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" _
                        '    & " (   NOT EGGTIM.ATIM+CONVERT(CHAR(10),EGGTIM.HME,103)   IN (" + F_SHMADEMENA_DELTIA + ") ) AND (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' and PELKOD='" + Text2(0).Text + "'"

                        'Else
582                     SQL2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" & " (   NOT EGGTIM.ATIM+CONVERT(CHAR(10),EGGTIM.HME,103)   IN (" + F_SHMADEMENA_DELTIA + ") ) AND (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' and PELKOD='" + Text2(0).Text + "'"
                        'End If

584                     R.Close
586                     R.Open SQL2, Gdb, adOpenForwardOnly, adLockReadOnly

                        'CSBmk_CS <                If R(0) = 0 Or IsNull(R(0))>
588                     If R(0) = 0 Or IsNull(R(0)) Then

590                         MHNYMA2.Timer1.Interval = 50
592                         MHNYMA2.SHOW 1

                            'Dim ans As Integer
594                         MsgBox "ΔΕΝ ΕΚΚΡΕΜΕΙ ΤΕΤΟΙΟ ΕΙΔΟΣ ΣΤΟΝ ΠΕΛΑΤΗ"
596                         timtext3.Text = 0
                            'CSBmk_CS <                End If>
                        End If

                        'CSBmk_CS <                If HDH_FORTOMENA + Val(TIMText3.Text) > >
598                     If HDH_FORTOMENA + Val(timtext3.Text) > R(0) Then
600                         MHNYMA2.Timer1.Interval = 50
602                         MHNYMA2.SHOW 1

604                         MsgBox "Η παραγγελθείσα ποσότητα είναι " + str(R(0)) + Chr(13) + "και έχετε ήδη τιμολογήσει " + Format(HDH_FORTOMENA + Val(timtext3.Text), "#####.00")
606                         timtext3.Text = 0
                            'CSBmk_CS <                End If>
                        End If

                        'CSBmk_CS <            End If>
                    End If

                    'CSBmk_CS <            If Val(TIMText3.Text) > (mXre - mPis) - >
608                 If Val(timtext3.Text) > (mXre - mPis) - mDesmia Then
610                     MHNYMA2.Timer1.Interval = 50
612                     MHNYMA2.SHOW 1

614                     MsgBox "υπάρχουσα ποσότητα " + Format(mXre - mPis, "###0") + " ΔΕΣΜ:" + Format(mDesmia, "####")

                        'CSBmk_CS <                If f_checkYpol = 2>
616                     If f_checkYpol = 2 Then
618                         timtext3.Text = 0
                            'CSBmk_CS <                End If>
                        End If
                        'CSBmk_CS <            End If>
                    End If
                    'CSBmk_CS <        End If>
                End If

                'CSBmk_CS <        If F_EXO_XROMATA = 1>
620             If F_EXO_XROMATA = 1 Then    ' ΒΑΖΩ ΤΗΝ ΠΟΣΟΤΗΤΑ ΣΤΟΝ ΠΙΝΑΚΑ ΓΙΑ ΝΑ ΤΟΝ ΤΡΑΒΗΞΕΙ ΣΤΟ ENHM_EGGTIM

                    Dim L9 As Long, M9 As Long

                    Dim SUMA9 As Single

622                 SUMA9 = 0

                    'CSBmk_CS <            For L9 = 1 To 25>
624                 For L9 = 1 To 25
                        'CSBmk_CS <                For M9 = 1 To 25>
626                     For M9 = 1 To 25
628                         SUMA9 = SUMA9 + Val(XROMATA(Grid1.row, L9, M9))
                            'CSBmk_CS <                Next>
                        Next
                        'CSBmk_CS <            Next>
                    Next

                    'CSBmk_CS <            If SUMA9 = 0>
630                 If SUMA9 = 0 Then
632                     XROMATA(Grid1.row, 1, 1) = timtext3.Text
                        'CSBmk_CS <            End If>
                    End If
                    'CSBmk_CS <        End If>
                End If
                  
                'CSBmk_CS <        If F_EISITHRIA = 1>
634             If F_EISITHRIA = 1 Then

                    Dim jj As Integer
                     
636                 jj = InStrRev(Grid1.TextMatrix(Grid1.row, f_o), "-")
638                 Grid1.TextMatrix(Grid1.row, f_o) = Left(Grid1.TextMatrix(Grid1.row, f_o), jj + 1) + str(Val(mID(Grid1.TextMatrix(Grid1.row, f_o), jj - 8, 8)) + gVal(timtext3.Text) - 1)
                    'CSBmk_CS <        End If>
                End If

                'CSBmk_CS <        If f_IS_PARTIDES = 1>
640             If f_IS_PARTIDES = 1 Then

                    Dim SYNOLO_POS As Single

                    Dim TAKE As Single

642                 SYNOLO_POS = gVal(timtext3.Text)   ' Grid1.TextMatrix(Grid1.row, f_p))

                    Dim rpart As New ADODB.Recordset, MTAKE As Single
                   
                    On Error GoTo 0
                   
                    'CSBmk_CS <            If F_DIOR = 1>
644                 If F_DIOR = 1 Then   ' ΑΝ ΠΡΟΚΕΙΤΑΙ ΓΙΑ ΔΙΟΡΘΩΣΗ
                 
                        'CSBmk_CS <                If F_POS_APOU = 2 Or F_POS_APOU = 3>
646                     If F_POS_APOU = 2 Or F_POS_APOU = 3 Then      ' 2=-ΕΙΣ  3=ΕΞΑΓΩΓΗ
648                         GRIDPARTIDES.Visible = True
650                         DATAPARTIDES.ConnectionString = gConnect
652                         DATAPARTIDES.RecordSource = "select LOTNUMBER,CONVERT(CHAR(10),HMEL,3) AS [ΛΗΞΗ],POSOPOL AS [ΕΠΙΛ],ID FROM PARTIDES WHERE ID_NUM=" + str(F_ID_NUM_DIOR) + " AND POSOAG<0  AND KOD='" + Grid1.TextMatrix(Grid1.row, f_k) + "'"
                   
                            'ΓΙΑ ΔΙΟΡΘΩΣΗ ΑΓΟΡΩΝ
                            'DATAPARTIDES.RecordSource = "select LOTNUMBER,CONVERT(CHAR(10),HMEL,3) AS [ΛΗΞΗ],POSOPOL AS [ΕΠΙΛ] FROM PARTIDES WHERE ID_NUM=" + str(F_ID_NUM_DIOR) + " AND SEIRA=" + str(Grid1.row) + "  AND KOD='" + Grid1.TextMatrix(Grid1.row, f_k) + "'"
654                         DATAPARTIDES.Refresh
                            'CSBmk_CS <                Else>
                        Else ' If f_pol = 2 Then     ' 1=POLISEIS 2=AGORES
656                         GRIDPARTIDES.Visible = True
658                         DATAPARTIDES.ConnectionString = gConnect
                            'ΓΙΑ ΔΙΟΡΘΩΣΗ ΑΓΟΡΩΝ
660                         DATAPARTIDES.RecordSource = "select LOTNUMBER,CONVERT(CHAR(10),HMEL,3) AS [ΛΗΞΗ],YPOLOIPO AS [ΥΠΟΛ],POSOPOL AS [ΕΠΙΛ],ID FROM PARTIDES WHERE ID_NUM=" + str(F_ID_NUM_DIOR) + " AND SEIRA=" + str(Grid1.row) + "  AND KOD='" + Grid1.TextMatrix(Grid1.row, f_k) + "'"
662                         DATAPARTIDES.Refresh
                            'CSBmk_CS <                End If>
                        End If
                   
                        'CSBmk_CS <            Else>
                    Else

                        'CSBmk_CS <                If F_POS_APOU = 2 Or F_POS_APOU = 3>
664                     If F_POS_APOU = 2 Or F_POS_APOU = 3 Then      ' 2=-ΕΙΣ  3=ΕΞΑΓΩΓΗ                    rpart.Open "select ID,LOTNUMBER,YPOLOIPO ,POSOPOL  FROM PARTIDES WHERE  YPOLOIPO>0 AND KOD='" + Grid1.TextMatrix(Grid1.row, f_k) + "' order by HMEL", Gdb, adOpenDynamic, adLockOptimistic
666                         rpart.Open "select LOTNUMBER,HMEL,YPOLOIPO ,POSOPOL ,ID FROM PARTIDES WHERE  YPOLOIPO>0 AND KOD='" + Grid1.TextMatrix(Grid1.row, f_k) + "' order by HMEL", Gdb, adOpenDynamic, adLockOptimistic

                            'CSBmk_CS <                    Do While Not rpart.EOF>
668                         Do While Not rpart.EOF

                                'CSBmk_CS <                        If rpart("YPOLOIPO") < SYNOLO_POS>
670                             If rpart("YPOLOIPO") < SYNOLO_POS Then
672                                 TAKE = rpart("YPOLOIPO")
                                    'CSBmk_CS <                        Else>
                                Else
674                                 TAKE = SYNOLO_POS
                                    'CSBmk_CS <                        End If>
                                End If

676                             SYNOLO_POS = SYNOLO_POS - TAKE
678                             Gdb.Execute "UPDATE PARTIDES SET POSOPOL=" + Replace(Format(TAKE, "000000.000"), ",", ".") + " WHERE ID=" + str(rpart("ID"))

                                'CSBmk_CS <                        If SYNOLO_POS <= 0>
680                             If SYNOLO_POS <= 0 Then Exit Do
682                             rpart.MoveNext
                                'CSBmk_CS <                    Loop>
                            Loop

684                         rpart.Close
                   
686                         GRIDPARTIDES.Visible = True
688                         DATAPARTIDES.ConnectionString = gConnect
690                         DATAPARTIDES.RecordSource = "select LOTNUMBER,CONVERT(CHAR(10),HMEL,3) AS [ΛΗΞΗ],YPOLOIPO AS [ΥΠΟΛ],POSOPOL  AS [ΕΠΙΛ],ID FROM PARTIDES WHERE  YPOLOIPO>0 AND KOD='" + Grid1.TextMatrix(Grid1.row, f_k) + "'"
692                         DATAPARTIDES.Refresh
                            'CSBmk_CS <                Else>
                        Else  '  f_pol = 2 Then     ' 1=POLISEIS 2=AGORES
694                         GRIDPARTIDES.Visible = True
696                         DATAPARTIDES.ConnectionString = gConnect
698                         DATAPARTIDES.RecordSource = "select LOTNUMBER,CONVERT(CHAR(10),HMEL,3) AS [ΛΗΞΗ],YPOLOIPO AS [ΥΠΟΛ],POSOPOL AS [ΕΠΙΛ],ID FROM PARTIDES WHERE  ID_NUM=0 AND SEIRA=" + str(Grid1.row) + "  AND KOD='" + Grid1.TextMatrix(Grid1.row, f_k) + "'"
700                         DATAPARTIDES.Refresh
                            'CSBmk_CS <                End If>
                        End If
                        'CSBmk_CS <            End If>
                    End If
                    'CSBmk_CS <        End If>
                End If
                'CSBmk_CS <    End If>
            End If     'If Grid1.Col = f_p Then  ' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

702         MTEXT = timtext3.Text
            'CSBmk_CS <    If KOMA > 0>
704         KOMA = InStr(MTEXT, ","): If KOMA > 0 Then Mid$(MTEXT, KOMA, 1) = "."
            ' μετακίνηση σε grid

706         Grid1.Text = MTEXT  ' κρατάει το grid το κείμενο

            'CSBmk_CS <    If Grid1.Col = f_p Or Grid1.Col = F_T Or>
708         If Grid1.Col = f_p Or Grid1.Col = F_T Or Grid1.Col = f_e Or Grid1.Col = f_f Or Grid1.Col = F_EKPT2 Then
710             FindSynolo
                'CSBmk_CS <    End If>
            End If

            'CSBmk_CS <    If F_PIN_PROELEYS = 1 And Grid1.Col = F_>
712         If F_PIN_PROELEYS = 1 And Grid1.Col = F_PROEL Then
                'CSBmk_CS <        If Val(Left(Grid1.TextMatrix(Grid1.row, >
714             If Val(Left(Grid1.TextMatrix(Grid1.row, F_PROEL), 3)) > 0 Then
                    'CSBmk_CS <            If Val(Left(Grid1.TextMatrix(Grid1.row, >
716                 If Val(Left(Grid1.TextMatrix(Grid1.row, F_PROEL), 3)) <= UBound(F_PINAKAS_PROEL) Then
                        
                        'ΑΝ ΔΙΠΛΑ ΑΠΟ ΤΗΝ ΠΡΟΕΛΕΥΣΗ ΒΑΛΩ ΚΩΔΙΚΟ ΠΑΡΤΙΔΑΣ ΝΑ ΒΡΙΣΚΕΙ ΤΟΝ ΚΩΔΙΚΟ ΑΠΟ ΠΙΝΑΚΑ ΚΑΙ ΝΑ ΒΑΖΕΙ ΑΠΟ 16-35 ΤΗΝ ΠΡΤΙΔΑ
                        'CSBmk_CS <                If Len(Trim(mID(Grid1.TextMatrix(Grid1.r>
                        
                        ' ΑΝ ΔΕΝ ΒΑΛΩ ΑΡΙΘΜΟ ΠΑΡΤΙΔΑΣ ΑΛΛΑ ΜΟΝΟ ΑΡΙΘΜΟ ΠΡΟΕΛΕΥΣΗΣ
718                     If Len(Trim(mID(Grid1.TextMatrix(Grid1.row, F_PROEL), 4, 15))) < 2 Then
                        
720                         If F_ARPART_IN_EID = 1 Then
                                Dim RTemp As New ADODB.Recordset
722                             RTemp.Open "SELECT ERG FROM EID WHERE KOD='" + Grid1.TextMatrix(Grid1.row, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic
724                             Grid1.TextMatrix(Grid1.row, F_PROEL) = F_PINAKAS_PROEL(Val(Grid1.TextMatrix(Grid1.row, F_PROEL))) + Space$(15) + CNull(RTemp(0))
726                             RTemp.Close
                             
                            Else
728                             Grid1.TextMatrix(Grid1.row, F_PROEL) = F_PINAKAS_PROEL(Val(Grid1.TextMatrix(Grid1.row, F_PROEL)))
                            End If
                          
                            'CSBmk_CS <                Else>
                        Else
730                         Grid1.TextMatrix(Grid1.row, F_PROEL) = Left(F_PINAKAS_PROEL(Val(Left(Grid1.TextMatrix(Grid1.row, F_PROEL), 3))) + Space(15), 15) + mID(Grid1.TextMatrix(Grid1.row, F_PROEL), 4, 20)
                            'CSBmk_CS <                End If>
                        End If
                        
                        'CSBmk_CS <            End If>
                    End If
                    'CSBmk_CS <        End If>
                End If
                'CSBmk_CS <    End If>
            End If

            ' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

            ' f_CanChangeEkptosi

            'CSBmk_CS <    If Grid1.Col <= Grid1.ColS - 3>
732         If Grid1.Col <= Grid1.ColS - 3 Then  ' πάει στην επόμενη κολόνα
734             Grid1.Col = Grid1.Col + 1

                'εαν δεν επιτρέπεται τότε ξανά στην επόμενη
                'CSBmk_CS <        If Grid1.Col = f_e And f_CanChangeEkptos>
736             If Grid1.Col = f_e And f_CanChangeEkptosi = 0 Then    'EKPTOSI
738                 Grid1.Col = Grid1.Col + 1
                    'CSBmk_CS <        End If>
                End If

                'CSBmk_CS <        If Grid1.Col = f_f And f_CanChangeFpa = >
740             If Grid1.Col = f_f And f_CanChangeFpa = 0 Then    'FPA

742                 HDH_FORTOMENA = 0

                    '  If f_mono_apoParaggelia = 1 Then

                    'CSBmk_CS <            If f_SYMPYKNVSH = 1>
744                 If f_SYMPYKNVSH = 1 Then  '  And Left(timText3.Text, 2) <> Left(F_2CHFIA_ZYGIZOMENA, 2) Then

                        'CSBmk_CS <                For LK0 = 1 To Grid1.row - 1>
746                     For LK0 = 1 To Grid1.row - 1

                            'brika idion kodiko
                            'CSBmk_CS <                    If Grid1.TextMatrix(LK0, 1) = Grid1.Text>
748                         If Grid1.TextMatrix(LK0, 1) = Grid1.TextMatrix(Grid1.row, 1) Then
750                             Grid1.TextMatrix(LK0, f_p) = Val(Grid1.TextMatrix(LK0, f_p)) + Val(Grid1.TextMatrix(Grid1.row, f_p))
752                             Grid1.TextMatrix(LK0, f_a) = Round(Val(Grid1.TextMatrix(LK0, f_p)) * Val(Grid1.TextMatrix(LK0, F_T)) * (100 - Val(Grid1.TextMatrix(LK0, f_e))) / 100, f_psifiaAjias)
                                ' Round(Val(P) * Val(t) * (100 - Val(e)) / 100 * (100 + g_Fpa(Val(Grid1.TextMatrix(R, f_f)))) / 100, f_psifiaAjias)

                                'CSBmk_CS <                        If LK0 <> Grid1.row>
754                             If LK0 <> Grid1.row Then
756                                 Grid1.TextMatrix(Grid1.row, f_a) = 0
758                                 Grid1.TextMatrix(Grid1.row, f_p) = 0
760                                 Grid1.TextMatrix(Grid1.row, F_T) = 0
                                    'CSBmk_CS <                        End If>
                                End If

762                             HDH_FORTOMENA = 1
                                'CSBmk_CS <                    End If>
                            End If

                            'CSBmk_CS <                Next>
                        Next

                        ' Grid1.TextMatrix(Grid1.row, f_a) = 0
                        ' Grid1.TextMatrix(Grid1.row, f_p) = 0
                        ' Grid1.TextMatrix(Grid1.row, F_T) = 0

764                     Grid1.Refresh

766                     FindSynolo
                        'CSBmk_CS <            End If>
                    End If

                    'CSBmk_CS <            If HDH_FORTOMENA = 0>
768                 If HDH_FORTOMENA = 0 Then
770                     Grid1.row = Grid1.row + 1
                        'CSBmk_CS <            Else>
                    Else
772                     Grid1.TextMatrix(Grid1.row, 3) = ""
774                     Grid1.TextMatrix(Grid1.row, 1) = ""
776                     Grid1.TextMatrix(Grid1.row, 2) = ""
778                     Grid1.TextMatrix(Grid1.row, 4) = ""
780                     Grid1.TextMatrix(Grid1.row, 5) = ""

                        'CSBmk_CS <            End If>
                    End If

782                 Grid1.Col = 1
                    'CSBmk_CS <        End If>
                End If

784             Me.Caption = "125"
                'CSBmk_CS <    End If>
            End If

            'CSBmk_CS <    If Grid1.Col > Grid1.ColS - 3 Or (Grid1.>
786         If Grid1.Col > Grid1.ColS - 3 Or (Grid1.Col > f_f + 1) Then ' πάει στην επόμενη κολόνα
                'πάει στην παρακάτω σειρά

                'CSBmk_CS <        If Grid1.row < Grid1.ROWS - 1 And Grid1.>
788             If Grid1.row < Grid1.rows - 1 And Grid1.row < f_SeiresTimologioy Then
790                 Grid1.row = Grid1.row + 1
792                 Grid1.Col = 1
794                 Me.Caption = "124"
                    'CSBmk_CS <        End If>
                End If

796             Me.Caption = "123"
                'CSBmk_CS <    End If>
            End If

            'CSBmk_CS <    If f_IS_PARTIDES = 1 And Grid1.Col > F_T>
798         If f_IS_PARTIDES = 1 And Grid1.Col > F_T Then
800             GRIDPARTIDES.Visible = False
                'CSBmk_CS <    End If>
            End If

            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

802         Grid1_EnterCell

            'CSBmk_CS <End If>
        End If   ' ************************************************

        Exit Sub

        'MHNYMA:
        '3030  HandleError "Par1:timtext3_KeyUp"
        '

        '<EhFooter>

        '       Exit Sub

timtex3_enter_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.timtex3_enter " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.timtex3_enter " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Static Sub FindSynolo()
        '=================================================================================
        '<EhHeader>
        On Error GoTo FindSynolo_Err

        '</EhHeader>
        Dim c, R, p, T, e, k, s, AJ, F, SUMPOS

        '    On Error GoTo MHNYMA

        ' If F_XONDR = 0 Then
        '    Grid1.BackColor = &H80FF80 ' prasino
        '    Label22.Visible = True ' Προσοχή οι τιμες είναι με ΦΠΑ
        '    Find_Lian_Synolo
        '    Exit Sub
        ' End If
        
100     Label22.Visible = False  'Προσοχή οι τιμες είναι με ΦΠΑ
102     Grid1.BackColor = F_GRIDCOLOR ' &HFF8080     ' mple

        ' TelFpa.Text = 0

104     For k = 1 To 9: SumFpa(k) = 0: Next

106     R = Grid1.row
108     c = Grid1.Col  ' απομνημονεύω την τρέχουσα στήλη kai σειρά

110     Grid1.Col = f_p: p = Replace(Grid1.Text, ",", ".") ' Grid1.Text
112     Grid1.Col = F_T: T = gVal(Replace(Grid1.Text, ",", "."))


        If F_CUSTOMER_SCREEN = 1 Then
            Grid1.Col = f_o: PAR1CUSTOMER.Label1.Caption = Grid1.Text
            Dim ffpa As Integer
            Grid1.Col = f_f: ffpa = Val(Grid1.Text)
            
            
            PAR1CUSTOMER.Label2.Caption = T + T * g_Fpa(ffpa) / 100
        End If



114     Grid1.Col = f_e: e = Grid1.Text 'gVal(Grid1.Text) 'Grid1.Text

116     Grid1.Col = f_f: F = Grid1.Text ' FPA

        ' TA EKANA REM TA EPOMENA 3 GIATI MOY DIELYAN TO SDA EPISTROFHS (STHN 12 KOLONA EBAZA THN PVLHUEISA POSOTHTA
        ' Grid1.Col = 12: Grid1.Text = g_Fpa(F) ' % FPA
        '  Grid1.TextMatrix(0, 12) = "ΦΠΑ%"
118     Grid1.TextMatrix(0, f_f) = "KAT.ΦΠΑ"
        '        'αξια με ΦΠΑ
        '180     Grid1.Col = F_mefpa: AJ = 0
        '
        '182     If Val(Grid1.TextMatrix(R, f_f)) > 0 Then
        '185         AJ = Round(0.00001 + Val(p) * Val(T) * (100 - Val(e)) / 100 * (100 + g_Fpa(Val(Grid1.TextMatrix(R, f_f)))) / 100, f_psifiaAjias)
        '188     End If
        '
        '190     Grid1.Text = Replace(Format(AJ, "####0.000"), ",", ".")

120     Grid1.Col = f_a
122     AJ = Round(0.00001 + gVal(p) * T * (100 - gVal(e)) / 100, 2) ' f_psifiaAjias)
124     Grid1.Text = Replace(Format(AJ, "####0.000"), ",", ".")

        Dim E2 As Double
126     E2 = 0
        '2014-01-12 2Η ΕΚΠΤΩΣΗ
128     If gVal(Grid1.TextMatrix(R, F_EKPT2)) > 0 Then
            
130         E2 = gVal(Replace(Grid1.TextMatrix(Grid1.row, F_EKPT2), ",", "."))
132         AJ = Round(0.00001 + (gVal(p) * T * (100 - gVal(e)) / 100) * (100 - gVal(E2)) / 100, f_psifiaAjias)
134         Grid1.TextMatrix(Grid1.row, f_a) = Replace(Format(AJ, "####0.000"), ",", ".")
        End If
        
        'αξια με ΦΠΑ
136     If gVal(Grid1.TextMatrix(R, f_f)) > 0 Then
138         Grid1.TextMatrix(R, F_mefpa) = Format(Round(0.00001 + gVal(p) * gVal(T) * ((100 - gVal(e)) / 100) * ((100 - gVal(E2)) / 100) * (100 + g_Fpa(gVal(Grid1.TextMatrix(R, f_f)))) / 100, f_psifiaAjias), "########.00")
        End If

140     f_sajia(Grid1.row) = AJ
142     f_fpa(Grid1.row) = F  ' ΚΑΤΗΓΟΡΊΑ ΦΠΑ
144     s = 0
146     SUMPOS = 0

        'Dim F_AJIA_ANAFPA(10) As Double

148     For k = 1 To Grid1.rows - 1

150         If Len(Grid1.TextMatrix(k, f_k)) > 0 Then

                f_sajia(k) = Round(0.00001 + (gVal(Grid1.TextMatrix(k, f_p)) * gVal(Grid1.TextMatrix(k, F_T)) * (100 - gVal(Grid1.TextMatrix(k, f_e))) / 100) * (100 - gVal(Grid1.TextMatrix(k, F_EKPT2))) / 100, f_psifiaAjias)

152             ' f_sajia(k) = Round(0.00001 + Val(Grid1.TextMatrix(k, f_a)), 2) '  f_psifiaAjias) '  Val(Grid1.TextMatrix(K, f_p)) * Val(Grid1.TextMatrix(K, F_T)) '
154             SUMPOS = SUMPOS + gVal(Grid1.TextMatrix(k, f_p))
156             s = s + f_sajia(k)
                ' ΣΟΥΜΆΡΕΙ ΣΕ ΚΆΘΕ ΚΑΤΗΓΟΡΊΑ φπα ΤΗΝ ΑΞΙΑ
158             If f_fpa(k) < 9 Then
160                 SumFpa(f_fpa(k)) = SumFpa(f_fpa(k)) + f_sajia(k)
                End If
                'F_AJIA_ANAFPA(Grid1.TextMatrix(k, f_f)) = F_AJIA_ANAFPA(Grid1.TextMatrix(k, f_f)) + f_sajia(k)
            End If

        Next

        'ΑΥΤΟΣ ΕΙΝΑΙ Ο ΥΠΟΛΟΓΙΣΜΟΣ ΠΟΥ ΕΧΩ ΣΤΗΝ ΤΕΛΙΚΗ ΚΑΤΑΧΩΡΗΣΗ ΣΤΟ ΤΙΜ
        'R_N.Open "SELECT SUM(POSO*TIMM*(100-EKPT)/100) AS KAU,FPA FROM EGGTIM WHERE ID_NUM=" + str(m_ID_NUM) + " GROUP BY FPA", Gdb, adOpenDynamic, adLockOptimistic
        'Do While Not R_N.EOF
        '   If R_N("FPA") > 0 And R_N("FPA") <= 9 Then
        '      kkk = R_N("FPA")
        '      AJ(kkk) = Round(R_N("KAU"), 2)
        '      FP(kkk) = Round(0.001 + R_N("KAU") * (g_Fpa(kkk)) / 100, 2)
        '      MSUMA3 = MSUMA3 + AJ(kkk) + FP(kkk)
        '   End If
        '   R_N.MoveNext
        'Loop
        'R_N.Close

162     Label15.Caption = SUMPOS

164     s = Round(0.00001 + s, f_psifiaAjias)

166     synolo.Caption = Format(s, "####0.00")
168     F = 0
        
170     f_strogg_FPA = 0
172     For k = 1 To 9
            ' στο παρακάτω for-next γίνεται έλεγχος ποιό ΦΠΑ ΕΊΝΑΙ ΓΕΜΑΤΟ ΓΙΑ ΝΑ ΠΑΡΕΙ ΤΗΝ ΣΤΡΟΓΓΥΛΟΠΟΙΗΣΗ
174         If SumFpa(k) > 0 Then
176             If f_strogg <> 0 Then
178                 f_strogg_FPA = k
                End If
            End If
180         F = F + Round(SumFpa(k) * g_Fpa(k) / 100, 2)
        Next
182     F = F + f_strogg

184     F = Round(0.00001 + F, 2)  ' f_psifiaAjias)

186     sfpa.Caption = Format(F, "####0.00")
    
188     SynoloG.Caption = Format(F + s, "####0.00")
190     Grid1.Col = c    ' το επαναφέρω εκεί που το βρήκα
192     Grid1.row = R

194     If F_PARAKRATISI > 0 Then
            ' parakratisiT.Text = Round(s * F_PARAKRATISI / 100, 2)
196         tel_pliroteo = gVal(SynoloG.Caption) - gVal(parakratisiT.Text)
        End If
        
        'ΑΝ ΕΧΩ ΕΦΚ ΤΟΤΕ ΠΟΛΑΠΛΑΣΙΑΖΕΙ ΤΙΣ ΠΟΣΟΤΗΤΕΣ ΕΠΙ ΤΟΝ ΕΙΔΙΚΟ ΦΟΡΟ ΚΑΤΑΝΑΛΩΣΗΣ
198     If F_EIDFOROSKAT = 1 And F_AJIA_APOU > 0 Then
            Dim sumEFK As Single
200         sumEFK = 0
202         For k = 1 To Grid1.rows - 1
204             If Len(Grid1.TextMatrix(k, f_k)) > 0 Then
206                 sumEFK = sumEFK + gVal(Grid1.TextMatrix(k, f_p)) * gVal(Grid1.TextMatrix(k, f_colEFK))
                End If
            Next
208         SynEfk.Visible = True
210         LabSynEfk.Visible = True
212         fpaEFK.Visible = True
214         LabfpaEFK.Visible = True
216         Teliko.Visible = True
            'lblΠληρωτέο.Visible = True
218         LabTeliko.Visible = True
           
220         SynEfk.Text = Format(sumEFK, "####0.00")
222         fpaEFK.Text = Format(sumEFK * g_Fpa(2) / 100, "####0.00")
           
224         Teliko.Text = Format(gVal(SynoloG.Caption) + gVal(SynEfk.Text) + gVal(fpaEFK.Text), "####0.00")
           
        End If


        If F_CUSTOMER_SCREEN = 1 Then
           
            PAR1CUSTOMER.synolo.Caption = SynoloG.Caption
        End If




226     Grid1.Refresh
    
        Exit Sub

MHNYMA:
228     HandleError "Par1:FINDSYNOLO"

230     Resume Next

        '<EhFooter>
        Exit Sub

FindSynolo_Err:
        '  'MsgBox Err.Description & vbCrLf & _
            "in ADOMERCNEW.par1.FindSynolo " & _
            "at line " & Erl, _
            vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.FindSynolo " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Static Sub Find_Lian_Synolo() ' ypologismos lianikhs

        '--------------------------------------------------------------------------------------
        On Error GoTo FindSynolo_Err

        Dim c, R, p, T, e, k, s, AJ, F, SUMPOS

        ' TelFpa.Text = 0

100     For k = 1 To 9: SumFpa(k) = 0: Next

102     R = Grid1.row
104     c = Grid1.Col  ' απομνημονεύω την τρέχουσα στήλη kai σειρά

106     Grid1.Col = f_p: p = Replace(Grid1.Text, ",", ".") ' Grid1.Text
108     Grid1.Col = F_T: T = gVal(Replace(Grid1.Text, ",", "."))

110     Grid1.Col = f_e: e = Grid1.Text 'gVal(Grid1.Text) 'Grid1.Text
112     Grid1.Col = f_f: F = Grid1.Text
        'αξια με ΦΠΑ
114     Grid1.Col = F_mefpa: AJ = 0

116     If Val(Grid1.TextMatrix(R, f_f)) > 0 Then
118         AJ = Round(0.00001 + Val(p) * Val(T) * (100 - Val(e)) / 100 * (100 + g_Fpa(Val(Grid1.TextMatrix(R, f_f)))) / 100, f_psifiaAjias)
        End If

120     Grid1.Text = Replace(Format(AJ, "####0.000"), ",", ".")

122     Grid1.Col = f_a
124     AJ = Round(0.00001 + Val(p) * T * (100 - Val(e)) / 100, f_psifiaAjias)
126     Grid1.Text = Replace(Format(AJ, "####0.000"), ",", ".")

        '2014-01-12 2Η ΕΚΠΤΩΣΗ
128     If Val(Grid1.TextMatrix(R, F_EKPT2)) > 0 Then

            Dim E2 As Double

130         E2 = gVal(Replace(Grid1.TextMatrix(Grid1.row, F_EKPT2), ",", "."))
132         AJ = Round(0.00001 + (Val(p) * T * (100 - Val(e)) / 100) * (100 - Val(E2)) / 100, f_psifiaAjias)
134         Grid1.TextMatrix(Grid1.row, f_a) = Replace(Format(AJ, "####0.000"), ",", ".")
        End If

136     f_sajia(Grid1.row) = AJ
138     f_fpa(Grid1.row) = F  ' ΚΑΤΗΓΟΡΊΑ ΦΠΑ
140     s = 0
142     SUMPOS = 0

144     For k = 1 To Grid1.rows - 1

146         If Len(Grid1.TextMatrix(k, f_k)) > 0 Then

148             f_sajia(k) = Round(0.00001 + Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias) '  Val(Grid1.TextMatrix(K, f_p)) * Val(Grid1.TextMatrix(K, F_T)) '
150             SUMPOS = SUMPOS + Val(Grid1.TextMatrix(k, f_p))
152             s = s + f_sajia(k)

                ' ΣΟΥΜΆΡΕΙ ΣΕ ΚΆΘΕ ΚΑΤΗΓΟΡΊΑ φπα ΤΗΝ ΑΞΙΑ
154             If f_fpa(k) < 9 Then
156                 SumFpa(f_fpa(k)) = SumFpa(f_fpa(k)) + f_sajia(k)
                End If
            End If

        Next

158     Label15.Caption = SUMPOS

160     s = Round(0.00001 + s, f_psifiaAjias)

162     synolo.Caption = Format(s, "####0.00")
164     F = 0: For k = 1 To 9: F = F + SumFpa(k) * g_Fpa(k) / 100: Next

166     F = Round(0.00001 + F, f_psifiaAjias)

168     sfpa.Caption = Format(0, "####0.00")
    
170     SynoloG.Caption = Format(0 + s, "####0.00")
172     Grid1.Col = c    ' το επαναφέρω εκεί που το βρήκα
174     Grid1.row = R

176     Grid1.Refresh
    
        Exit Sub

MHNYMA:
178     HandleError "Par1:FINDSYNOLO"

180     Resume Next

        '<EhFooter>
        Exit Sub

FindSynolo_Err:
        '  'MsgBox Err.Description & vbCrLf & _
            "in ADOMERCNEW.par1.FindSynolo " & _
            "at line " & Erl, _
            vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.FindSynolo " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Public Function ExecuteL(Scode As String, Optional fcheckonly As Boolean) As Boolean
'    Dim a$
'    'ExecuteL = EbExecuteLine(StrPtr(Scode + ":a$=gvar"), 0&, 0&, Abs(fcheckonly)) = 0
'
'    '  παράδειγμα χρησιμοποίησης μακροεντολής
'    '  --------------------------------------
'
'    '   Private Sub Command1_Click()
'    '   Dim res As Boolean, var As Single
'    a$ = "GVAR=1"
'    a$ = Left(a$, 2)
'
'    ' ExecuteL = EbExecuteLine(StrPtr(A$), 0&, 0&, Abs(fcheckonly)) = 0
'
'    '
'    '   res = ExecuteLine("var=2+3*(5+6):a$=var")
'    '
'    '   End Sub
'End Function
Function Gefyroma(g_hme, g_atim, g_eidoskpe, gfkod, gfaj, m_r, mgre)

        'εδώ δημιουργούνται οι εγγραφές για το dirlog\temp.dbf
        'και ενημερώνονται οι ανώτεροι λογαριασμοί στο abat και ανοίγονται αν δεν υπάρχουν
        '*********************** γεφυρωμα ****************************
        '<EhHeader>
        On Error GoTo Gefyroma_Err

        '</EhHeader>
        Dim xre_pis(20), xre_Pis_rec(20), xreosis(20), GetList(20)

        Dim DUM

        Dim jj, JJK, m_ait As String

        Dim kodiklog(9), TYP(9), XP(9), kau(9), a As String, mnq As String, eid_par As String

        ' Dim mgre As String
        Dim k As Integer

        Dim atr_xre, atr_pis, lw, DIAFORETIKA

100     atr_xre = 0: atr_pis = 0

        Dim ok_gefyr

        Dim DB       As Database, dbL As Database

        Dim r2       As Recordset, gefyres As New ADODB.Recordset

        Dim bbs      As String

        Dim anoigma  As Integer

        ' Dim PEL As Recordset
        Dim kod_pel  As String

        Dim mbat_pel As String

        ' Dim abat2 As Recordset
        Dim lo       As Integer

        Dim JJ_ROBLEM

        Dim mmhkos, MHKOS(5) As Integer

        Dim MHNAS

        Dim sql, K2 As Integer, logar(5) As Integer

        On Error GoTo MHNYMA

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        '   Set dbL = OpenDatabase(gdirlog, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        '   Set dbL = OpenDatabase(gdirlog, False, False, gConnect)
        'End If

102     Set DB = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
104     Set dbL = OpenDatabase(gdirlog, False, False, "dBASE III;")

        Dim par As Recordset

        'ΒΡΙΣΚΩ ΑΠΟ ΤΙΣ GEFYRES.DBF ΤΗΝ ΚΑΤΑΛΛΗΛΗ ΕΓΓΡΑΦΗ
        'Set par = db.OpenRecordset("SELECT *FROM PARASTAT WHERE EIDOS='" + left(g_atim, 1) + "';")

106     gefyres.Open "SELECT * FROM GEFYRES", Gdb, adOpenDynamic, adLockBatchOptimistic

        'Set gefyres = db.OpenRecordset("GEFYRES")
        'Gefyres.Move par("GEF_P") - 1  ' RECORD ΓΕΦΥΡΏΣΕΩΝ ΕΠΙ ΠΙΣΤΩΣΕΙ
108     gefyres.Move m_r - 1  ' RECORD ΓΕΦΥΡΏΣΕΩΝ ΕΠΙ ΠΙΣΤΩΣΕΙ

110     ok_gefyr = 1

        'a = ctim(0)
112     If g_hme <= ghme_orist Then

114         MsgBox " ΔΕΝ ΕΠΙΤΡΕΠΟΝΤΑΙ ΚΙΝΗΣΕΙΣ  ΣΕ ΟΡΙΣΤΙΚΟΠΟΙΗΜΕΝΕΣ ΕΓΓΡΑΦΕΣ"

            Exit Function

        End If

        '
        '  GoTo 444
116     If InStr("oOοΟ", gefyres("gefyrono")) Then
118         ok_gefyr = 0

            Exit Function

        End If

        Dim EIDPAR

120     eid_par = "  "

122     If gefyres("mEID_PAR") <> "  " Then
124         eid_par = Left(gefyres("mEID_PAR"), 5) + "-" + g_atim    ' Τ.πωλ-Τ00012Γ
        End If

        '  If Gefyres("mmGRE") <> " " Then
        '     mgre = Gefyres("mmGRE") ' Σ
        '  End If
126     m_ait = gefyres("parathrhs")    'π.χ.  ΠΩΛΗΣΕΙΣ

        '
        '
        '
128     For k = 1 To 9
130         a = LTrim(str(k))
132         XP(k) = gefyres("xpi" + a)
134         kodiklog(k) = gefyres("Vkod" + a)  ' 3000****
136         TYP(k) = gefyres("Vtypos" + a)
        Next

138     gefyres.Close

        '*** ψαχνω για διπλοεγγραφές
140     gvar = "Select *from temp where eidhme='1' and hme=#"
142     gvar = gvar + Format(g_hme, "mm/dd/yyyy") + "# and apa='"
144     gvar = gvar + eid_par + "'"

        'EID.DatabaseName = gdirlog
        'EID.RecordSource = gvar

        'EID.Refresh

        '  Set temp = dbL.OpenRecordset(gvar)

146     temp.DatabaseName = gdirlog
148     temp.Connect = "dBase III;"
150     temp.RecordSource = gvar    '  "select * from temp" ' gvar

152     abat.DatabaseName = gdirlog
154     abat.Connect = "dBase III;"

156     abat2.DatabaseName = gdirlog
158     abat2.Connect = "dBase III;"

160     temp.Refresh

162     If temp.Recordset.RecordCount > 0 Then
164         MsgBox ("ΥΠΑΡΧΕΙ ΙΔΙΟΣ ΑΡΙΘΜΟΣ & ΗΜΕΡ/ΝΙΑ ΣΤΗΝ ΓΕΝ.ΛΟΓΙΣΤΙΚΗ ΜΕ ΧΡΕΩΣΗ " + Format(temp.Recordset("XRE"), "###,###,###,###"))
166         MsgBox ("ΑΔΥΝΑΤΗ Η ΓΕΦΥΡΩΣΗ ΜΕ ΤΗΝ ΓΕΝΙΚΗ ΛΟΓΙΣΤΙΚΗ ")

            'temp.Close
            'abat.Close
            Exit Function

        End If

        ' Set gTim = db.OpenRecordset("select *from tim where atim='" + g_atim + "' and hme>=#" + Format(g_hme, "mm/dd/yyyy") + "#;")

168     Set fSCR = CreateObject("MSScriptControl.ScriptControl")

        'cEGGTIM.RecordSource = "select *from eggtim"
        'cEGGTIM.Refresh

170     fSCR.language = "vbscript"

        'DOK.RecordSource = "select * from TIM where ATIM='" + g_atim + "' and HME>='" + Format(g_hme, "mm/dd/yyyy") + "';"

        'DOK.ConnectionString = gConnect
        'DOK.Refresh

        ' DOK.RecordSource = "select *from tim where atim='" + g_atim + "' and hme>='" + Format(g_hme, "mm/dd/yyyy") + "';"
172     DOK.RecordSource = "select * from TIM where ATIM='" + g_atim + "' and HME>='" + Format(g_hme, "mm/dd/yyyy") + "';"

        'CTIM.DatabaseName = gDir
174     DOK.ConnectionString = gConnect
176     DOK.Refresh

178     fSCR.addObject "cTIM", DOK    ' ADO  dok einai to pragmatiko antikeimeno  ctim to antistoixo sthn VBSCRIPT

        'fSCR.AddObject "DOK", DOK ' ADO

180     fSCR.addObject "temp", temp
182     fSCR.addObject "abat", abat
184     fSCR.addObject "abat2", abat2

        'fSCR.AddObject "cMEM", CMEM
186     fSCR.addObject "mactext", mactext

        Dim debug1

        ' εδώ παίρνει το vkod1,vtypos1 έως το vkod9,typos9
188     For k = 1 To 9

190         If kodiklog(k) <> " " Then
                '            sele TEMP

192             debug1 = DOK.Recordset("aji")
194             a = TYP(k)    '
                'dum = ExecuteLine("gvar=" + a)
196             a = mac(a)

198             If Val(a) <> 0 Then    ' γγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγ

                    '
                    '             ' EAN EIMASTE STHN ANALYSH EIDVN
200                 If mID$(kodiklog(k), 3, 2) = "**" Then    ' κωδικός λογιστικής
202                     DIAFORETIKA = 20  ' για νά κάνει πολλά loops ένα για κάθε είδος ΑΠΟΘΗΚΗΣ
                    Else
204                     DIAFORETIKA = 1   ' περνάει μόνο μία φορά
                    End If

206                 For lw = 1 To DIAFORETIKA   ' MONO STHN PERIPTOSH TON EIDVN=20 ALLIOS =1
OLD1010:

                        '* αν δεν έχει ανάλυση να προχωράει
208                     If lw > DIAFORETIKA Then Exit For
210                     If gfaj(lw) = 0 And DIAFORETIKA > 1 Then
212                         lw = lw + 1
214                         GoTo OLD1010
                        End If

                        '
                        '               sele TEMP
                        '
216                     temp.Recordset.AddNew
                        'dum = ExecuteLine("

                        'dum = ExecuteLine("gvar=" + Trim(kodiklog(K)))

                        'dum = ExecuteLine("gvar=" + gvar)

218                     gvar = Trim(kodiklog(k))
220                     bbs = gvar

222                     If InStr("30-50", Left(bbs, 2)) > 0 And InStr(bbs, "***") > 0 Then    ' κωδικός πελάτου
224                         If InStr("30-50", Left(bbs, 2)) Then    '
                                'sele Abat
226                             anoigma = 0
                                '******** εχει ΚΩΔΙΚΟ ΛΟΓΙΣΤΙΚΗΣ *******************
228                             kod_pel = Trim(PEL.Recordset("shm2"))

                                '**  έχει κωδικό λογιστικής
230                             mbat_pel = "3"

232                             If Len(Trim(PEL.Recordset("shm2"))) = gDIG_4 Then
234                                 mbat_pel = "4"
                                End If

236                             If Len(kod_pel) = 0 Then
238                                 MsgBox " δεν έχω κωδικό λογιστικής " + Format(g_hme, "dd/mm/yyyy") + " " + g_atim + " " + g_eidoskpe
240                                 ok_gefyr = 0
                                    'abat.Close
                                    'temp.Close
242                                 Set fSCR = Nothing

                                    Exit Function

                                End If

                                '                              seek mbat_pel + kod_pel
244                             abat.RecordSource = "select *from abat where bat='" + mbat_pel + "' and kod='" + kod_pel + "'"

246                             abat.Refresh

248                             If abat.Recordset.RecordCount > 0 Then
250                                 anoigma = 1
                                Else
                                    ' αλλά δεν υπάρχει στην λογιστική
                                    '** βρίσκω τον τελευταίο κωδικό
252                                 abat.Recordset.AddNew
254                                 abat.Recordset("bat") = mbat_pel
256                                 abat.Recordset("kod") = kod_pel
258                                 abat.Recordset("ono") = PEL.Recordset("epo")
260                                 abat.Recordset("afm") = PEL.Recordset("afm")
262                                 kod_pel = abat.Recordset("kod")
264                                 abat.Recordset.Update

                                    'PEL.Edit
                                    'PEL("SHM2") = abat.recordset("KOD")
                                    'anoigma = 1
                                End If         '  eof()

                                '
                            End If    '  left(kodiklog(k),2) $ '30 50'

                            '
                            '
                            '                      Temp.Edit
266                         temp.Recordset("KOD") = kod_pel

                            'If DRX_TO_EURO = 1 Then
                            '    temp.recordset("XRE") = eround(a / 340.75, 2)
                            'Else
268                         temp.Recordset("XRE") = a
                            'End If

270                     ElseIf mID$(bbs, 1, 4) = "****" Then    ' κωδικός λογιστικής ΑΠΟΘΗΚΗΣ
272                         temp.Recordset("KOD") = Trim(gfkod(lw))
274                         temp.Recordset("XRE") = gfaj(lw)  ' ΤΟ ΔΙΑΧΩΡΙΖΩ ΠΑΡΑΠΑΝΩ ΔΡΧ("ΕΥΡΩ

                        Else

276                         temp.Recordset("kod") = bbs
278                         temp.Recordset("XRE") = a

                        End If

280                     temp.Recordset("APA") = eid_par
282                     temp.Recordset("EIDHME") = "1"
284                     temp.Recordset("AIT") = m_ait
286                     temp.Recordset("HME") = g_hme
288                     temp.Recordset("GRE") = mgre
290                     temp.Recordset("XPI") = XP(k)
292                     temp.Recordset("PROOD") = 0
294                     temp.Recordset("ARIU") = 0
296                     temp.Recordset("SYGK") = ""
298                     jj = Len(Trim(temp.Recordset("kod")))
300                     temp.Recordset.Update

                    Next     ' LW =1 TO DIAFORETIKA

                End If       '     IF &a > 0 γγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγ
            End If    ' if kodiklog(k) <> ' '

            '
            '
        Next  ' K = 1 To 9

302     If Abs(atr_xre - atr_pis) >= 0.01 Or JJ_ROBLEM = 1 Then
        Else
304         temp.Recordset.MoveFirst

306         FlxGridDeltion.Visible = True

            ' ενημερώνω τα αρχείο λογαριασμών abat
308         Do While Not temp.Recordset.EOF    ' and  eidhme+dtos(hme)+left(gre,1)+apa =  '1' + dtos(g_hme) + mgre  + EID_PAR
310             f_mc = f_mc + 1    '
312             FlxGridDeltion.TextMatrix(f_mc, 0) = temp.Recordset(0)
314             FlxGridDeltion.TextMatrix(f_mc, 1) = temp.Recordset("apa")
316             FlxGridDeltion.TextMatrix(f_mc, 2) = temp.Recordset("ait")

318             If temp.Recordset("xpi") = "Π" Then FlxGridDeltion.TextMatrix(f_mc, 4) = Format(temp.Recordset("xre"), "######0.00")
320             If temp.Recordset("xpi") = "Χ" Then FlxGridDeltion.TextMatrix(f_mc, 3) = Format(temp.Recordset("xre"), "######0.00")

322             mmhkos = Len(Trim(temp.Recordset("kod")))

324             If temp.Recordset("eidHME") = "3" Then    ' απογραφή
326                 MHNAS = "00"
                Else
328                 MHNAS = find_mhnas(temp.Recordset("hme"))
                End If

330             sql = "select *from abat WHERE "

332             If mmhkos = gDIG_4 Then sql = sql + "(bat='4' and left(kod," + str(gDIG_4) + ")='" + temp.Recordset("kod") + "') or"
334             If mmhkos >= gDIG_3 Then sql = sql + "(bat='3' and left(kod," + str(gDIG_3) + ")='" + Left(temp.Recordset("kod"), gDIG_3) + "') or "
336             If mmhkos >= 4 Then sql = sql + "(bat='2' and left(kod,4)='" + Left(temp.Recordset("kod"), 4) + "') or "
338             If mmhkos >= 2 Then sql = sql + "(bat='1' and left(kod,2)='" + Left(temp.Recordset("kod"), 2) + "') "

                'συγκεντρώνω τους λογαριασμους που με ενδιαφέρουν
340             abat2.RecordSource = sql
                ' εαν λειπουν τους ανοιγω
342             MHKOS(1) = 2: MHKOS(2) = 4: MHKOS(3) = gDIG_3: MHKOS(4) = gDIG_4

344             If mmhkos = 2 Then lo = 1
346             If mmhkos = 4 Then lo = 2
348             If mmhkos = gDIG_3 Then lo = 3
350             If mmhkos = gDIG_4 Then lo = 4
352             abat2.Refresh

354             If abat2.Recordset.RecordCount < lo Then

356                 For K2 = 1 To lo: logar(K2) = 0: Next

                    On Error Resume Next

358                 abat2.Recordset.MoveFirst

360                 Do While Not abat2.Recordset.EOF
362                     logar(Val(abat2.Recordset("bat"))) = 1
364                     abat2.Recordset.MoveNext

366                     DoEvents
                    Loop

368                 For K2 = 1 To lo

370                     If logar(K2) = 0 Then
372                         abat2.Recordset.AddNew
374                         abat2.Recordset("bat") = LTrim(str(K2))
376                         abat2.Recordset("kod") = Left(temp.Recordset("kod"), MHKOS(K2))
378                         abat2.Recordset.Update
                        End If

                    Next

                End If

380             abat2.Recordset.MoveFirst

382             Do While Not abat2.Recordset.EOF
384                 abat2.Recordset.Edit

386                 If temp.Recordset("xpi") = "Π" Then
388                     If IsNull(abat2.Recordset("ppros" + MHNAS)) Then abat2.Recordset("ppros" + MHNAS) = temp.Recordset("xre") Else abat2.Recordset("ppros" + MHNAS) = abat2.Recordset("ppros" + MHNAS) + temp.Recordset("xre")
                    Else

390                     If IsNull(abat2.Recordset("ppros" + MHNAS)) Then abat2.Recordset("xpros" + MHNAS) = temp.Recordset("xre") Else abat2.Recordset("xpros" + MHNAS) = abat2.Recordset("xpros" + MHNAS) + temp.Recordset("xre")
                    End If

392                 abat2.Recordset.Update
394                 abat2.Recordset.MoveNext
                Loop

396             abat2.Recordset.Close
398             temp.Recordset.MoveNext
            Loop

        End If

400     Set fSCR = Nothing

        Exit Function

MHNYMA:
402     HandleError "Par1:gefyroma"

404     Resume Next

        '<EhFooter>
        Exit Function

Gefyroma_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Gefyroma " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Gefyroma " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function printCrystal(MATIM, char_date)

        '------------------------------------------------------------------------
        'δημιουργεί τα αρχεία TIM  , EGGTIM , MEM    ΣΤΟ  c:\mercvb\reports\reports.mdb"
        'ΓΙΑ ΤΗΝ ΕΚΤΥΠΩΣΗ ΤΟΥ ΤΙΜΟΛΟΓΙΟΥ
        '<EhHeader>
        On Error GoTo printCrystal_Err

        '</EhHeader>
        Dim cID_NUM     As String  ' id_num apo to TIM

        Dim cENTITYMARK As String

        Dim sql         As String, fDB As Database, D1 As Date, d2 As Date

        Dim synt        As String, DB As Database

        Dim DUM

        Dim REGGTIM       As New ADODB.Recordset

        Dim Rtim          As New ADODB.Recordset

        Dim R             As Recordset

        Dim M_KODKAYSIMOY As String

100     M_KODKAYSIMOY = ""

        Dim Rmem As Recordset

        Dim k    As Integer

        Dim KN   As Integer

        Dim r3   As New ADODB.Recordset

        Dim mTimh, mPoso

        Dim PROHG

        '-----------  next lines are OK --------------------------

        On Error GoTo MHNYMA
        
        Dim mf_paras As String
        mf_paras = GET_CVALUE(" SELECT TITLOS FROM PARASTAT WHERE EIDOS='" + Left(MATIM, 1) + "'") + " "

102     printCrystal = 1    ' αν ειναι dum < -900 τοτε δεν τυπωνει

        'ΑΝ ΜΕΝ ΕΙΝΑΙ ΔΙΟΡΘΩΣΗ ΠΡΕΠΕΙ ΝΑ ΚΡΑΤΑ ΤΗΝ ΠΑΛΙΑ ΗΜΕΡΟΜΗΝΙΑ
        'ΑΝ ΔΕ ΕΙΝΑΙ ΑΚΥΡΩΤΙΚΟ ΠΡΕΠΕΙ ΝΑ ΕΧΕΙ ΤΗΝ ΣΗΜΕΡΙΝΗ ΗΜΕΡΟΜΗΝΙΑ
        'TA EKANA REM STIS 26/5/2009
        'If Len(f_dior_mAtim) > 0 Then
        '   d1 = Now
        '   d2 = DateAdd("d", 1, d1)
        'Else
        '   d1 = DTPicker1.value
        '   d2 = DateAdd("d", 1, d1)
        'End If

        ' ΑΝ ΕΙΝΑΙ ΑΚΥΡΩΤΙΚΟ ΝΑ ΕΧΕΙ ΤΗΝ ΗΜΕΡΟΜΗΝΙΑ ΤΗΣ ΕΙΣΟΔΟΥ GWORKDAY
104     If f_Akyr Then
106         D1 = gWorkDay
        Else
108         D1 = char_date
        End If

110     d2 = DateAdd("d", 1, D1)

112     Set DB = OpenDatabase("c:\mercvb\reports\reports.mdb")

114     ADD_FIELDDAO "TIM", "AJ6", "SINGLE", DB
116     ADD_FIELDDAO "TIM", "FPA6", "SINGLE", DB

118     ADD_FIELDDAO "TIM", "AJ7", "SINGLE", DB
120     ADD_FIELDDAO "TIM", "FPA7", "SINGLE", DB

122     ADD_FIELDDAO "TIM", "METAF", "SINGLE", DB
124     ADD_FIELDDAO "TIM", "DEH", "CHAR(12)", DB
126     ADD_FIELDDAO "TIM", "C1", "CHAR(130)", DB
128     ADD_FIELDDAO "TIM", "C2", "CHAR(130)", DB
130     ADD_FIELDDAO "TIM", "ADT", "CHAR(30)", DB
        
132     ADD_FIELDDAO "TIM", "LITRA", "SINGLE", DB
134     ADD_FIELDDAO "TIM", "EFK", "SINGLE", DB
136     ADD_FIELDDAO "TIM", "ORA", "CHAR(8)", DB
        
138     If F_METAFORIKES >= 1 Then
            '         ADD_FIELDDAO "TIM", "C3", "CHAR(45)", dB
            '
            '         ADD_FIELDDAO "TIM", "C12", "CHAR(45)", dB
            '         ADD_FIELDDAO "TIM", "C13", "CHAR(45)", dB
            '
            '          ADD_FIELDDAO "AYTOK", "AYTOK", "CHAR(45)", dB
            '
            '         ADD_FIELDDAO "TIM", "NUM1", "SINGLE", dB
            '          ADD_FIELDDAO "TIM", "NUM2", "SINGLE", dB
            '           ADD_FIELDDAO "TIM", "NUM3", "SINGLE", dB
            '            ADD_FIELDDAO "TIM", "NUM11", "SINGLE", dB
           
        End If
        
140     ADD_FIELDDAO "EGGTIM", "PROELEYSH", "CHAR(10)", DB
142     ADD_FIELDDAO "EGGTIM", "MIKTA", "SINGLE", DB
144     ADD_FIELDDAO "EGGTIM", "BARCODE", "CHAR(14)", DB
146     ADD_FIELDDAO "EGGTIM", "XVRA", "CHAR(20)", DB
148     ADD_FIELDDAO "EGGTIM", "EKPT2", "SINGLE", DB

        ' KOLA , MIKTAKILA
150     ADD_FIELDDAO "EGGTIM", "KOLA", "SINGLE", DB
152     ADD_FIELDDAO "EGGTIM", "MIKTAKILA", "SINGLE", DB
        
154     ADD_FIELDDAO "EGGTIM", "LITRA", "SINGLE", DB
156     ADD_FIELDDAO "EGGTIM", "EFK", "SINGLE", DB
        
158     ADD_FIELDDAO "MEM", "ELINE", "CHAR(250)", DB
        
        ' AYTA TA PEDIA EINAI GIA THN ΘΕΣΣΑΛΟΝΙΚΗ ΝΑ ΤΥΠΩΝΕΙ ΣΤΗΝ ΙΔΙΑ ΣΕΛΙΔΑ ΠΡΩΤΟΤΥΠΟ+ΑΝΤΙΓΡΑΦΟ
        ' ΑΛΛΑ ΔΕΝ ΜΠΟΡΕΙ ΝΑ ΤΑ ΑΝΟΙΞΕΙ ΑΥΤΟΜΑΤΑ ΠΡΕΠΕΙ ΝΑ ΑΝΟΙΧΤΟΥΝ ΣΤΟ ΜΔΒ ΑΠΕΥΘΕΙΑΣ
        
        '                ADD_FIELDDAO "MEM", "ONO1", "CHAR(50)", DB
        '                ADD_FIELDDAO "MEM", "ONO2", "CHAR(50)", DB
        '                ADD_FIELDDAO "MEM", "ONO3", "CHAR(50)", DB
        '                ADD_FIELDDAO "MEM", "ONO4", "CHAR(50)", DB
        '
        '
        '
        '
        '                ADD_FIELDDAO "MEM", "POS1", "SINGLE", DB
        '                ADD_FIELDDAO "MEM", "POS2", "SINGLE", DB
        '                ADD_FIELDDAO "MEM", "POS3", "SINGLE", DB
        '                ADD_FIELDDAO "MEM", "POS4", "SINGLE", DB
        '
        '
        '                ADD_FIELDDAO "MEM", "TIM1", "SINGLE", DB
        '                ADD_FIELDDAO "MEM", "TIM2", "SINGLE", DB
        '                ADD_FIELDDAO "MEM", "TIM3", "SINGLE", DB
        '                ADD_FIELDDAO "MEM", "TIM4", "SINGLE", DB
        '
        '
        '
        '                ADD_FIELDDAO "MEM", "AJ1", "SINGLE", DB
        '                ADD_FIELDDAO "MEM", "AJ2", "SINGLE", DB
        '                ADD_FIELDDAO "MEM", "AJ3", "SINGLE", DB
        '                ADD_FIELDDAO "MEM", "AJ4", "SINGLE", DB
        
        'If gConnect = "Access" Then
        '  Set fDB = OpenDatabase(gDir, False, False)
        '
        'Else
        '  Set fDB = OpenDatabase(gDir, False, False, gConnect)
        'End If
        '                  arn_enhm_apot

        '=====   EGGTIM  ======

        'On Error Resume Next
        '        On Error GoTo 0
160     DB.Execute "delete *from eggtim"

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
        'REGGTIM.MoveNext  other.text
        'Loop

        Dim SYNT6  As String
        
        Dim SQLTIM As String
        
162     synt = " left(ATIM," + str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "'"

164     If F_METAFORIKES >= 1 Then
166         If F_METAF3SYNAL = 1 Then  ' ΤΟ ΔΙΑΦΟΡΕΤΙΚΟ ΕΙΝΑΙ ΣΚΟΠΟΣ2>ADT GIA NA PIASO TON ENTOLEA
168             sql = "select C5,C6,C7,OTHERMOVEPURPOSETITLE,ISNULL(ENTITYUID,'') AS ENTITYUID,ISNULL(AUTHENTICATIONCODE,'') AS AUTHENTICATIONCODE,   AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,TRP,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL,AJ6,FPA6,AJ7,FPA7,ISNULL(ELGA,0)+ISNULL(KR1,0) AS ELGA,PEL.CH1 AS DEH,SKOPOS2 AS ADT,EKPT5 AS PROHGYPOL,ID_NUM,LITRA,PARAKRATISI AS EFK,ORA,TIM.C1,TIM.C2,TIM.C3,TIM.C12,TIM.C13,TIM.NUM1,TIM.NUM2,TIM.NUM3,TIM.NUM11,ISNULL(ELINE,'') AS ELINE "
            Else
170             If F_METAFORIKES = 2 Then  'CAI\
172                 sql = "select C5,C6,C7,OTHERMOVEPURPOSETITLE,ISNULL(ENTITYUID,'') AS ENTITYUID,ISNULL(AUTHENTICATIONCODE,'') AS AUTHENTICATIONCODE,AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,TRP,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL,AJ6,FPA6,AJ7,FPA7,ISNULL(ELGA,0)+ISNULL(KR1,0) AS ELGA,PEL.CH1 AS DEH,SKOPOS2 AS ADT,EKPT5 AS PROHGYPOL,ID_NUM,LITRA,PARAKRATISI AS EFK,ORA,TIM.C1,TIM.C2,TIM.C3,TIM.C12,TIM.C13,TIM.NUM1,TIM.NUM2,TIM.NUM3,TIM.NUM11,ISNULL(ELINE,'') AS ELINE "
                Else
174                 sql = "select C5,C6,C7,OTHERMOVEPURPOSETITLE,ISNULL(ENTITYUID,'') AS ENTITYUID,ISNULL(AUTHENTICATIONCODE,'') AS AUTHENTICATIONCODE,AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,TRP,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL,AJ6,FPA6,AJ7,FPA7,ISNULL(ELGA,0)+ISNULL(KR1,0) AS ELGA,PEL.CH1 AS DEH,ADT,EKPT5 AS PROHGYPOL,ID_NUM,LITRA,PARAKRATISI AS EFK,ORA,TIM.C1,TIM.C2,TIM.C3,TIM.C12,TIM.C13,TIM.NUM1,TIM.NUM2,TIM.NUM3,TIM.NUM11,ISNULL(ELINE,'') AS ELINE "
                End If
              
            End If

        Else
176         sql = "select OTHERMOVEPURPOSETITLE,ISNULL(ENTITYUID,'') AS ENTITYUID,ISNULL(AUTHENTICATIONCODE,'') AS AUTHENTICATIONCODE,AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,TRP,PEL.EPA,PEL.DIE,PEL.AFM,(SELECT SUM(ISNULL(XREOSI,0) ) - SUM(ISNULL(PISTOSI,0) ) FROM EGG WHERE KOD=PEL.KOD AND EIDOS=PEL.EIDOS AND HME<=TIM.HME)+ISNULL(PEL.AYP,0) AS TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL,AJ6,FPA6,AJ7,FPA7,ISNULL(ELGA,0)+ISNULL(KR1,0) AS ELGA,PEL.CH1 AS DEH,ADT,EKPT5 AS PROHGYPOL,ID_NUM,LITRA,PARAKRATISI AS EFK,ORA,ISNULL(ELINE,'') AS ELINE "

        End If
 
178     sql = sql + " ,CURRENCY,isnull(ENTITYMARK,'') as ENTITYMARK,QRURL from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"
        '700     DB.Execute "delete *from tim"
180     Rtim.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly
        
        On Error Resume Next
        
        '---------- qrcode ----------------------------------
        Dim CURL As String

182     CURL = CNull(Rtim!QRURL)
184     Set Image2.Picture = QRCodegenBarcode(CURL)
186     Picture1.Picture = QRCodegenBarcode(CURL, 0, QRCodegenEcc_LOW, 1, 40, , , 500) ' Edited by Saturnian
188     Clipboard.Clear ' Edited by Saturnian
190     Clipboard.SetData QRCodegenBarcode(CURL, 0, QRCodegenEcc_LOW, 1, 40, , , 500) ' Edited by Saturnian
192     SavePicture Picture1.Picture, "C:\MERCVB\FOTOp.BMP"

        Dim stm             As ADODB.Stream

        Dim mAPOLIAsyndesis As Integer

194     mAPOLIAsyndesis = 0

        Dim strFile As String

196     If Len(CURL) < 2 Then
198         strFile = "c:\mercvb\fotoEMPTY.bmp"

200         If GET_NVALUE("SELECT ISNULL(ISMYDATA,0) from PARASTAT WHERE EIDOS='" + Left(MATIM, 1) + "'") = 1 Then
202             mAPOLIAsyndesis = 1
            End If
            
        Else
204         strFile = "c:\mercvb\fotoP.bmp"
        End If

206     Set stm = New ADODB.Stream
208     stm.Type = adTypeBinary
210     stm.Open
212     stm.LoadFromFile strFile

        '---------- qrcode ----------------------------------
        On Error GoTo MHNYMA
        
214     If F_TYPONO_KENES = 1 Then
216         SYNT6 = ""
        Else
218         SYNT6 = "EGGTIM.POSO<>0 AND "
        End If

220     sql = "select EGGTIM.FPA,ATIM,HME as [Ημερομ],KODE as [Kωδ_Είδ],ONOMA as [Περιγραφή],"

        '230     If False Then ' F_XONDR = 0 Then    'LIANIKH
        '240         sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],MIK_AJIA as [Αξία],"
        '        Else
222     sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],ROUND(POSO*TIMM*(100-EKPT)/100*(100-EKPT2)/100,2) as [Αξία],"
        '       End If

224     sql = sql + "MONA,EKPT,EGGTIM.PROELEYSH,(CASE WHEN MIKTA IS NULL THEN 0 ELSE MIKTA END) AS MIKTA,(select TOP 1 ERG FROM BARCODES WHERE KOD=EGGTIM.KODE) AS [BARCODE],ISNULL(  (SELECT ISNULL(CH4,'') FROM EID WHERE KOD=EGGTIM.KODE)  ,'..') AS [XVRA],EKPT2,KOLA,MIKTAKILA  "  ',LITRA,EFK "
226     sql = sql + " from EGGTIM WHERE "             '           INNER JOIN EID ON EGGTIM.KODE=EID.KOD where "
228     sql = sql + SYNT6 ' + " APOT=" + str(Val(Left(APOTH1.Text, 2)))
230     sql = sql + "  ID_NUM=" + str(Rtim("ID_NUM")) + " ORDER BY PROOD_AJ;"
        
232     If f_IS_PARTIDES = 1 Then

234         sql = "select EGGTIM.FPA,EGGTIM.ATIM,EGGTIM.HME as [Ημερομ],KODE as [Kωδ_Είδ],ONOMA as [Περιγραφή],"

236         If False Then '  F_XONDR = 0 Then    'LIANIKH
238             sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],MIK_AJIA as [Αξία],"
            Else
240             sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],ROUND(POSO*TIMM*(100-EKPT)/100*(100-EKPT2)/100,2) as [Αξία],"
            End If

242         sql = sql + "MONA,EKPT,EGGTIM.PROELEYSH,(CASE WHEN MIKTA IS NULL THEN 0 ELSE MIKTA END) AS MIKTA,(select TOP 1 ERG FROM BARCODES WHERE KOD=EGGTIM.KODE) AS [BARCODE],ISNULL(PARTIDES.LOTNUMBER,' ') AS [XVRA],EKPT2,KOLA,MIKTAKILA  "  ',LITRA,EFK "
244         sql = sql + " from EGGTIM  INNER JOIN EID ON EGGTIM.KODE=EID.KOD "
246         sql = sql + " left join PARTIDES ON EGGTIM.KODE=PARTIDES.KOD AND EGGTIM.ATIM=PARTIDES.ATIM AND PARTIDES.SEIRA=EGGTIM.PROOD_AJ where "
            '    sql = sql + SYNT6 + " APOT=" + str(Val(Left(APOTH1.Text, 2)))
            '    sql = sql + " and EGGTIM.ID_NUM=" + str(Rtim("ID_NUM")) + " ORDER BY PROOD_AJ;"
248         sql = sql + SYNT6 + " EGGTIM.ID_NUM=" + str(Rtim("ID_NUM")) + " ORDER BY PROOD_AJ;"
        
        End If
        
        'left(ATIM," + str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' ORDER BY PROOD_AJ;"

        '2/10/13  sql = sql + " from EGGTIM  where APOT=" + str(Val(Left(APOTH1.Text, 2))) + " and  left(ATIM," + str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' ORDER BY PROOD_AJ;"

        On Error GoTo printCrystal_Err

250     REGGTIM.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

        'On Error GoTo 0 ' NA DIAGRAFEI

        'copy to access table eggtim
252     Set R = DB.OpenRecordset("eggtim")
        'On Error Resume Next

254     Set Rmem = DB.OpenRecordset("mem")

        'REGGTIM.Close
        'REGGTIM.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

        Dim syn_ajia As Single

256     syn_ajia = 0

        Dim SYN_EKPT

258     SYN_EKPT = 0
260     REGGTIM.MoveFirst

        On Error Resume Next
     
        'On Error GoTo 0  ' NA DIAGRAFEI
        Dim mCo As Integer

262     If f_2TIMOL_1PAGE = 1 Then

264         For mCo = 1 To 4
266             Rmem.Edit
268             Rmem("ono" + LTrim(str(mCo))) = ""
270             Rmem("pos" + LTrim(str(mCo))) = 0
272             Rmem("tim" + LTrim(str(mCo))) = 0
274             Rmem("aj" + LTrim(str(mCo))) = 0
276             Rmem.Update
            Next

        End If
     
278     mCo = 0
     
        'On Error GoTo printCrystal_Err
280     Do While Not REGGTIM.EOF

282         If printCrystal < 0 Then Exit Do

284         R.AddNew
         
            On Error GoTo 0 ' printCrystal_Err

286         For k = 0 To REGGTIM.FIELDS.Count - 1

                On Error GoTo LATOS

288             If REGGTIM(k).Type = 202 Or REGGTIM(k).Type = 200 Or REGGTIM(k).Type = 129 Then    ' STRING
                    'R(REGGTIM(k).Name) = Left(REGGTIM(k), R(REGGTIM(k).Name).Size)
290                 R(k) = Left(REGGTIM(k), R(k).Size)
                    'DB.Execute "update EGGTIM SET " + REGGTIM(k).Name + "='" + REGGTIM(k) + "' WHERE ????"
                Else
292                 R(k) = REGGTIM(k)
                    'R(REGGTIM(k).Name) = REGGTIM(k)
                    'DB.Execute "update EGGTIM SET " + REGGTIM(k).Name + "=" + str(REGGTIM(k)) + " WHERE ????"
                End If

            Next
            
294         If f_is_benzinadiko = 1 Then ' 1 ΕΙΔΟΣ ΜΟΝΟ ΣΤΑ ΒΕΝΖΙΝΑΔΙΚΑ
              
296             M_KODKAYSIMOY = REGGTIM("Kωδ_Είδ")
              
298             M_KODKAYSIMOY = Trim(LTrim(M_KODKAYSIMOY))
              
300             If InStr("10  11  12  13  20  21  30 31  32  40  50  ", M_KODKAYSIMOY) > 0 Then

                    'ok
302                 If Len(M_KODKAYSIMOY) > 2 Then
304                     M_KODKAYSIMOY = "****"
                    End If

                Else
306                 M_KODKAYSIMOY = "****"
              
                End If
              
            End If

308         mTimh = REGGTIM("Tιμ_Μον")
310         mPoso = REGGTIM("Ποσότητα")

312         If False Then ' F_XONDR = 0 Then    'LIANIKH
            Else
                'Round(Val(P) * Val(t) * (100 - Val(e)) / 100, f_psifiaAjias)
314             R("Αξία") = Round(0.00001 + R("Ποσότητα") * R("Tιμ_Μον") * (100 - R("EKPT")) / 100 * (100 - R("EKPT2")) / 100, 2) ' f_psifiaAjias)
                'r("Αξία") = r("Ποσότητα") * r("Tιμ_Μον") * (100 - r("EKPT")) / 100
            End If

316         mCo = mCo + 1
          
318         If f_2TIMOL_1PAGE = 1 Then
320             If mCo <= 4 Then
322                 Rmem.Edit
324                 Rmem("ono" + LTrim(str(mCo))) = Left(REGGTIM("Περιγραφή"), 30)
326                 Rmem("pos" + LTrim(str(mCo))) = REGGTIM("Ποσότητα")
328                 Rmem("tim" + LTrim(str(mCo))) = REGGTIM("Tιμ_Μον")
330                 Rmem("aj" + LTrim(str(mCo))) = R("Αξία")
332                 Rmem.Update
                End If
            End If

334         syn_ajia = syn_ajia + R("Αξία")
336         SYN_EKPT = SYN_EKPT + Round(0.00001 + R("Ποσότητα") * R("Tιμ_Μον") * R("EKPT") / 100, 2) 'f_psifiaAjias)
338         R.Update

340         REGGTIM.MoveNext
        Loop

        Dim sum_mikta As Single, sum_posothta As Single

342     sum_mikta = 0
344     sum_posothta = 0

346     REGGTIM.Close

        Dim SAJIA2 As Single   ' Η ΑΞΙΑ ΥΠΟΛΟΓΙΖΟΜΕΝΗ

348     R.MoveFirst

350     Do While Not R.EOF

352         If Not IsNull(R("fpa")) Then
354             R.Edit

356             If f_BARCODE_STOTIM = 1 Then
358                 r3.Open "select ERG FROM EID WHERE KOD='" + R("[Kωδ_Είδ]") + "'", Gdb, adOpenDynamic, adLockPessimistic

360                 If Not r3.EOF Then
362                     R("[Kωδ_Είδ]") = r3(0)
                    End If

364                 r3.Close

                End If

366             If R("fpa") = 0 Then
368                 R("FPA") = 2
370                 MsgBox "Προσοχή βρέθηκε ΦΠΑ=0 και μετετράπη σε κατηγ. 2"
                Else
372                 R("fpa") = g_Fpa(R("fpa"))
                End If

374             sum_mikta = sum_mikta + R("[Ποσότητα]") * R("[Tιμ_Μον]")
376             sum_posothta = sum_posothta + R("[Ποσότητα]")
378             R.Update
            End If

380         R.MoveNext
        Loop

        'db.Execute "update EGGTIM set FPA=" + Str(g_Fpa(6)) + " where FPA=6"

        'db.Execute "update EGGTIM set FPA=" + Str(g_Fpa(1)) + " where FPA=1"
        'db.Execute "update EGGTIM set FPA=" + Str(g_Fpa(2)) + " where FPA=2"
        'db.Execute "update EGGTIM set FPA=" + Str(g_Fpa(3)) + " where FPA=3"
        'db.Execute "update EGGTIM set FPA=" + Str(g_Fpa(4)) + " where FPA=4"

        'db.Execute "update EGGTIM set FPA=0 where FPA=5 "

382     KN = f_SynSeiresTimologioy - R.RecordCount  'f_SeiresTimologioy - R.RecordCount

        'sta packing list na mhn bazei paragemismata
384     If F_KOLA_MIKTA_LABELS = 2 Then

        Else

386         For k = 1 To KN
388             R.AddNew: R.Update
            Next

        End If

390     R.Close

        ' synt = " left(ATIM," + str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "'"

        ' αυτο το κομματι δουλευei με odbc  ================================
        'Dim db9 As Database
        'db.Execute "drop table tim"
        'Set db9 = OpenDatabase("", False, False, gConnect)
        'SQL = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL "
        'SQL = SQL + " into tim in 'c:\mercvb\reports\reports.mdb'  from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>=#" + Format(d1, "mm/dd/yyyy") + "# and HME<#" + Format(d2, "mm/dd/yyyy") + "#;"
        'db9.Execute SQL
        'GoTo 22
        ' αυτο το κομματι δουλευei με odbc   ==================================

        '
        ' EINAI OK ALLA TO EBALA STO ENHM_PEL
        'Dim rTemp As New ADODB.Recordset
        'Dim SQLt As String, TYP As Single
        'SQLt = "SELECT EIDOS,KOD,"
        'SQLt = SQLt + "SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END )  -"
        ' SQLt = SQLt + " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END )   AS YP"
        ' SQLt = SQLt + " FROM EGG  where HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("kod") + "' GROUP BY KOD"
        'rTemp.Open SQLt, Gdb, adOpenDynamic, adLockOptimistic
        'If IsNull(rTemp("YP")) Then
        '   TYP = 0
        'Else
        '   TYP = rTemp("YP")
        'End If
        'Gdb.Execute "UPDATE PEL SET TYP=AYP+ " + Format(TYP, "###########0.00")

        '========  TIM ==========
        '680     sql = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL,AJ6,FPA6,ELGA,PEL.CH1 AS DEH,EKPT5 AS PROHGYPOL "
        '690     sql = sql + "  from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"
392     DB.Execute "delete *from tim"
        '710     Rtim.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

394     PROHG = Rtim("prohgYPOL")

        'copy to access table tim
396     Set R = DB.OpenRecordset("tim")

        On Error GoTo LATOS

        '
        'Dim N As Integer
        'On Error GoTo printCrystal_Err
        'For N = 0 To R.FIELDS.Count - 1
        '   If R.FIELDS(N).Name = "DEH" Then
        '      OK = 1
        '   End If
        'Next
        'db.Execute "ALTER TABLE tim add column DEH CHAR(12) "
        '
        Dim ENTITYUID, authenticationCode
        Dim m_Aji As Single

398     Rtim.MoveFirst

        ' select ATIM,
        ' HME as shme,AJI as ajia,PEL.TYP,

        ENTITYUID = Rtim("ENTITYUID"): authenticationCode = Rtim("AUTHENTICATIONCODE")

        '   On Error GoTo printCrystal_Err
400     R.AddNew

        On Error Resume Next

        'Dim ISPIST As Boolean

402     R("SHME") = Rtim("SHME")
404     R("TYP") = Rtim("TYP")

406     R("EPO") = Left(CNull(Rtim("EPO")), Rtim("epo").ActualSize)
408     R("EPA") = Left(CNull(Rtim("EPA")), Rtim("epa").ActualSize)
410     R("DIE") = Left(CNull(Rtim("DIE")), Rtim("die").ActualSize)
412     R("AFM") = Left(CNull(Rtim("AFM")), Rtim("afm").ActualSize)

414     R("DOY") = Left(CNull(Rtim("DOY")), Rtim("doy").ActualSize)
416     R("KOD") = Left(CNull(Rtim("KOD")), Rtim("kod").ActualSize)
418     R("THL") = Left(CNull(Rtim("THL")), Rtim("thl").ActualSize)
420     R("POL") = Left(CNull(Rtim("POL")), Rtim("pol").ActualSize)

422     R("AJIA") = Rtim("AJIA")
424     R("AJI") = Rtim("AJI")

426     m_Aji = Rtim("AJI")

428     R("AJ1") = Rtim("AJ1")
430     R("AJ2") = Rtim("AJ2")    '
432     R("AJ3") = Rtim("AJ3")
434     R("AJ4") = Rtim("AJ4")
436     R("AJ5") = Rtim("AJ5")
438     R("AJ6") = Rtim("AJ6")
440     R("AJ7") = Rtim("AJ7")
         
442     R("FPA1") = Rtim("FPA1")
444     R("FPA2") = Rtim("FPA2")    '
446     R("FPA3") = Rtim("FPA3")
448     R("FPA4") = Rtim("FPA4")
450     R("FPA6") = Rtim("FPA6")
452     R("FPA7") = Rtim("FPA7")
454     R("ORA") = Rtim("ORA")

456     R("ELGA") = Rtim("ELGA")

458     If F_METAFORIKES >= 1 Then
 
460         R("C3") = Rtim("C3")
462         R("C2") = Rtim("C2")
464         R("C1") = Rtim("C1")
          
466         If Len(R!C1) = 0 Then
468             R!C1 = "  "
            End If

470         R("C13") = Rtim("C13")
472         R("C12") = Rtim("C12")
            
474         R("NUM1") = Rtim("NUM1")
476         R("NUM2") = Rtim("NUM2")
478         R("NUM11") = Rtim("NUM11")
480         R("AYTOK") = Rtim("AYTOK")
           
        End If
 
        '                                                                        seira
        Dim typos_par As String

        Dim e_pros    As Integer
 
482     e_pros = 1
        '40  ΔΕΛΤΙΟ ΑΠΟΣΤΟΛΗΣ ΑΠΟ ΠΕΛΑΤΗ 0   1
        '158 Δελτίο Αποστολής    0   1
        '158 Επιστροφή Δελτίου Αποστολής 0   1
        '158 Δελτίο Αποστολής Λιανικής   0   1
        '161 Tιμολόγιο   1   1
        '162 ΤΙΜΟΛΟΓΙΟ ΠΑΡΟΧ. ΥΠΗΡΕΣΙΩΝ  1   1
        '165 ΤΙΜΟΛΟΓΙΟ   1   1
        '169 Πιστωτικό Τιμολόγιο -1  1
        '169 Πιστωτικό Τιμ.Παροχής Εκπτωσης  -1  1
        '173 Απόδειξη Λιανικής Πώλησης   1   1
        '174 Απόδειξη Παροχής Υπηρεσιών  1   1
        '175 Δελτίο Επιστρ.Λιανικώς Πωληθ.   -1  1
        '215 ΕΙΔΙΚΟ ΑΚΥΡΩΤΙΚΟ ΣΤΟΙΧΕΙΟ   -1  1
        '221 ΔΕΛΤΙΟ ΑΠΟΣΤΟΛ.ΤΙΜΟΛΟΓΙΟ    1   1
        '222 Δ. Αποστολής - Τιμολόγιο    1   1
        '223 ΤΙΜΟΛΟΓΙΟ (Παροχης Υπηρεσιων) - ΔΕΛΤΙΟ ΑΠΟΣΤΟΛΗΣ    1   1
        '226 ΤΙΜΟΛΟΓΙΟ (Παροχή Υπηρεσιων) - ΤΙΜΟΛΟΓΙΟ (Πώληση Αγαθών)    1   1
        '227 ΤΙΜΟΛΟΓΙΟ(Παροχή Υπηρεσιων)-ΤΙΜΟΛΟΓΙΟ(Πώληση Αγαθών)-ΔΕΛ.ΑΠ.    1   1
        '229 ΔΕΛΤΙΟ ΠΟΣΟΤΙΚΗΣ ΠΑΡΑΛΑΒΗΣ - ΠΙΣΤΩΤΙΚΟ ΤΙΜΟΛΟΓΙΟ    -1  1
        '231 Απόδειξη Λιανικής Πώλησης  ΔΑ   1   1
        '232 ΑΠΟΔΕΙΞΗ ΠΑΡΟΧΗΣ ΥΠΗΡΕΣΙΩΝ - ΔΕΛΤΙΟ ΑΠΟΣΤΟΛΗΣ   1   1
        '233 ΑΠΟΔΕΙΞΗ ΠΑΡΟΧΗΣ ΥΠΗΡΕΣΙΩΝ - ΑΠΟΔΕΙΞΗ ΛΙΑΝΙΚΗΣ ΠΩΛΗΣΗΣ  1   1
        '235 ΑΠΟΔΕΙΞΗ ΕΠΙΣΤΡΟΦΗΣ - ΔΕΛΤΙΟ ΑΠΟΣΤΟΛΗΣ  -1  1
        '236 Πιστωτικό Τιμολόγιο(ΔΑ) -1  1
        '502 ΑΝΤΙΤΥΠΟ    0   1

        '997981320;064202089;;161;AA;0010;0.00;0.00;4.67;0.00;0.00;0.00;0.00;1.12;0.00;5.79;0
        'PARADEIGMA EKDOSHS AKYROTIKOY (215)  GIA TO PARAPANIO PARASTATIKO
        '997981320;064202089;;215;AA#161#0010#AA;002;0.00;0.00;-4.67;0.00;0.00;0.00;0.00;-1.12;-0.00;-5.79;0

484     If Len(Trim(F_TYPOS_PAR)) > 0 Then
486         typos_par = Trim(F_TYPOS_PAR)

            'ΑΡΝΗΤΙΚΑ
488         If typos_par = 169 Or typos_par = 175 Or typos_par = 215 Or typos_par = 229 Or typos_par = 235 Or typos_par = 236 Then
490             e_pros = -1
            End If
     
        Else

492         If Left(Rtim("ATIM"), 1) = "L" Then
494             typos_par = "231"
496         ElseIf Left(Rtim("ATIM"), 1) = "T" Then
498             typos_par = "222"
500         ElseIf Left(Rtim("ATIM"), 1) = "l" Then
502             typos_par = "167" ' "173"
    
504         ElseIf Left(Rtim("ATIM"), 1) = "y" Then
506             typos_par = "232"
     
508         ElseIf Left(Rtim("ATIM"), 1) = "P" Then
510             typos_par = "169"
512             e_pros = -1
514         ElseIf Left(Rtim("ATIM"), 1) = "p" Then
516             typos_par = "175"
518             e_pros = -1
            End If
        End If
 
        Dim rmem2 As New ADODB.Recordset, AFMPEL As String
 
520     rmem2.Open "select top 1 AFMPEL,ISNULL(PELONO,'') AS PELONO,ISNULL(PELDIE,'') AS PELDIE,ISNULL(PELEPA,'') AS PELEPA,ISNULL(SEIRA1,'') AS SEIRA1,ISNULL(TELPEL,'') AS TELPEL from MEM", Gdb, adOpenForwardOnly, adLockReadOnly

        Dim m_pelono, M_PELEPA, m_peldie, M_SEIRA1, m_telpel As String
 
522     If IsNull(rmem2(0)) Then
524         AFMPEL = ""
        Else
526         AFMPEL = rmem2(0)
        End If
 
528     m_pelono = rmem2("PELONO")
530     M_PELEPA = rmem2("PELEPA")
532     m_peldie = rmem2("PELDIE")
534     M_SEIRA1 = rmem2("SEIRA1")
536     m_telpel = rmem2("TELPEL")
 
538     rmem2.Close
   
        Dim gia_akyr_tim As String
   
        '[<]029235109;047200105   ;;201607141034;222;Γ;00225;0.00;0.00;100.00;0.00;0.00;0.00;0.00;24.00;0.00;124.00;0[>]"
        Dim e_line       As String, E2 As String

        Dim e_AkyrLine   As String
 
        Dim CCC4         As String
 
        'Dim F_TYPFOROL  As Integer ' TYPOS FOROLOGIKOY 1=ALGO 2=SIGNPRO
        Dim diax         As String

540     diax = ";"

542     If F_TYPFOROL = 11 Then diax = "/"
 
        Dim benz      As String ' κωδικοσ & λιτρα καυσιμου
        Dim benz_atim As String
 
        Dim benz2     As String '
        Dim mparat2   As String
      Dim mTher As String: mTher = ""
        'algo
544     If F_TYPFOROL = 1 Or F_TYPFOROL = 11 Or F_TYPFOROL = 21 Then '1=ALGOBOX  21=RBS  11=ALGOBOX ME "/" ANTI ";"
       
'546         If f_is_benzinadiko = 1 Then
'                If M_KODKAYSIMOY = "30" Or M_KODKAYSIMOY = "31" Then
'                  If F_meIFESTO = 2 Then
'                     mTher = "#" + Trim(CNull(Rtim("AFM"))) + "*" + Trim(CNull(Rtim("DEH")))
'                  Else
'                      mTher = ""
'                  End If
'
'                Else
'                    mTher = ""
'                End If
'                'benz = "R?" + M_KODKAYSIMOY + "!" + Replace(Format(sum_posothta, "##0.000"), ",", ".") ' 429.280" ' R?20!ΛΙΤΡΑ
'                '           If typos_par = 215 Then ' ΑΚΥΡΩΤΙΚΟ
'                '
'                '                benz2 = "#" + PAR_GIA_AKYROSH + "#"   '  PAR1.PAR_GIA_AKYROSH.Caption = Trim(CNull(R!c1)) + "#" + mID(c, 2, 7)  ' 221#151   TYPOS ELINE+ARIUMOS
'                '                benz_atim = mID(Rtim("ATIM"), 2, 6)
'                '           Else
'                '               benz2 = "0#173#100"
'                '               benz_atim = parat.Text   '"672967"  ' αριθμος αποδειξης
'                '           End If
'
'                ' einai akyrotiko alla den xero poio akyronei
'                ' ΑΥΤΟ ΧΡΕΙΑΖΕΤΑΙ ΣΤΗΝ ΕΠΑΝΕΚΤΥΠΩΣΗ ΠΟΥ ΔΕΝ ΞΕΡΕΙ ΠΟΙΟ ΘΑ ΑΚΥΡΩΣΕΙ
'548             If typos_par = 215 And Len(Trim(PAR_GIA_AKYROSH)) < 1 Then
'
'550                 If Len(Trim(PAR_GIA_AKYROSH)) < 1 Then
'552                     gia_akyr_tim = GET_CVALUE("select SXETIKO from TIM where " + synt + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';")
'554                     gia_akyr_tim = GET_CVALUE("select ATIM from TIM WHERE ID_NUM=" + gia_akyr_tim)
'
'                        ' e_pros = -1
'556                     PAR_GIA_AKYROSH.Caption = GET_CVALUE("SELECT TRIM(C1) FROM PARASTAT WHERE EIDOS='" + Left(gia_akyr_tim, 1) + "'") + "#" + Trim(mID(gia_akyr_tim, 2, f_psifia_atim))
'
'                    End If
'
'                End If
'
'
'
'
'
'
'558             If Len(Trim(PAR_GIA_AKYROSH)) > 1 Then      'ΑΚΥΡΩΤΙΚΟ=====================
'560                 benz = ""
'562                 benz2 = "0#" + Trim(PAR_GIA_AKYROSH)
'564                 benz_atim = mID(Rtim("ATIM"), 2, 6)
'566                 typos_par = 215
'568                 e_pros = -1
'                Else
'
'570                 If typos_par = "175" Or typos_par = "169" Or typos_par = "173" Or typos_par = "159" Or typos_par = "317" Then
'572                     benz = "?" + M_KODKAYSIMOY + "!" + Replace(Format(sum_posothta, "##0.000"), ",", ".") + mTher '+ "#" + Trim(CNull(Rtim("AFM"))) + "*" + Trim(CNull(Rtim("DEH"))) ' 429.280" ' R?20!ΛΙΤΡΑ
'574                     benz2 = ""
'                        '998481696;;DLD19004842;?20!6.949;202110050732;0005;0027867;0725;173;0;0114051;0.00;0.00;8.06;0.00;0.00;0.00;0.00;1.94;0.00;10.00;0;9CC6FF41D63BC8393C0819912054F562EB4165F6
'576                 ElseIf F_THERM = 1 And F_THER_TIM_XORISAPOD = 1 Then
'578                     benz = "?" + M_KODKAYSIMOY + "!" + Replace(Format(sum_posothta, "####0.000"), ",", ".") + mTher '+ "#" + Trim(CNull(Rtim("AFM"))) + "*" + Trim(CNull(Rtim("DEH")))  ' 429.280" ' R?20!ΛΙΤΡΑ
'580                     benz2 = ""
'                    Else
'
'                        Dim mParat As String
'
'582                     mParat = Trim(Left$(parat.Text, 10))
'
'584                     If Val(mParat) > 99999 Then
'586                         benz2 = "#173#" + Right(Trim(mParat), 6)
'                        Else
'588                         benz2 = "#173#" + Right(Trim(mParat), 5)
'                        End If
'
'590                     benz = "R?" + M_KODKAYSIMOY + "!" + Replace(Format(sum_posothta, "##0.000"), ",", ".") ' 429.280" ' R?20!ΛΙΤΡΑ
'                    End If
'
'592                 benz_atim = mID(Rtim("ATIM"), 2, 6)   ' parat.Text   '"672
'
'                End If  'Len(Trim(PAR_GIA_AKYROSH)) > 3 Then      'ΑΚΥΡΩΤΙΚΟ=====================
'
'            Else   '               f_is_benzinadiko <> 1 Then
'594             benz2 = " "
'                '           If Len(Trim(PAR_GIA_AKYROSH)) > 3 Then      'ΑΚΥΡΩΤΙΚΟ=====================
'                '                benz = ""
'                '                benz2 = "0#" + Trim(PAR_GIA_AKYROSH)
'                '                benz_atim = mID(Rtim("ATIM"), 2, 6)
'                '                typos_par = 215
'                '                e_pros = -1
'
'596             If Len(Trim(PAR_GIA_AKYROSH)) > 3 Then      'ΑΚΥΡΩΤΙΚΟ=====================
'598                 typos_par = 215
'600                 e_pros = -1
'602                 benz2 = "0#" + Trim(PAR_GIA_AKYROSH)
'                End If
'
'                ' PERIPTVSH DIORUVSHS
'                ' einai akyrotiko alla den xero poio akyronei
'604             If typos_par = 215 Then
'
'                    ' Dim gia_akyr_tim As String
'606                 If Len(Trim(PAR_GIA_AKYROSH)) < 1 Then
'608                     gia_akyr_tim = GET_CVALUE("select SXETIKO from TIM where " + synt + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';")
'610                     gia_akyr_tim = GET_CVALUE("select ATIM from TIM WHERE ID_NUM=" + gia_akyr_tim)
'
'612                     e_pros = -1
'614                     PAR_GIA_AKYROSH.Caption = GET_CVALUE("SELECT TRIM(C1) FROM PARASTAT WHERE EIDOS='" + Left(gia_akyr_tim, 1) + "'") + "#" + Trim(mID(gia_akyr_tim, 2, f_psifia_atim))
'616                     benz2 = "0#" + Trim(PAR_GIA_AKYROSH)
'
'                    End If
'
'                End If
'
'618             benz = ""
'620             benz_atim = mID(Rtim("ATIM"), 2, 6)
'            End If  '  f_is_benzinadiko = 1 Then
'
'            '  benzinadiko [<]031051334//R?20!429.280/201811050643/221/0#173#100/672967/0.00/0.00/494.36/0.00/0.00/0.00/0.00/118.65/0.00/613.01/0[>]
'            Dim eisf As Single
'
'622         If F_TYPFOROL = 21 Then 'RBS
'624             If Len(Trim(benz2)) = 0 Then benz2 = "ΑΝΕΥ"
'626             e_line = "[<]" + AFMPEL + diax + Left(CNull(Rtim("AFM")), 9) + diax + benz
'628             eisf = e_pros * nNull(Rtim("EFK"))
'            Else
'
'630             e_line = "[<]" + AFMPEL + diax + Left(CNull(Rtim("AFM")), 9) + diax + benz + diax + Format(R("Shme"), "YYYYMMDDHHmm")
'632             eisf = 0
'            End If
'
'            '/                          221       /173#100/672967/0.00/0.00/
'
'634         e_line = e_line + diax + typos_par + diax + benz2 + diax + benz_atim + diax
'636         E2 = tdot(e_pros * Rtim!aj4 + e_pros * Rtim!aj3) + diax + tdot(e_pros * Rtim!aj1 + e_pros * Rtim!AJ7) + diax + tdot(e_pros * Rtim!aj2 + e_pros * Rtim!aj6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aj5 + eisf) + diax
'638         E2 = E2 + tdot(e_pros * Rtim!FPA4 + e_pros * Rtim!fpa3) + diax + tdot(e_pros * Rtim!FPA1 + e_pros * Rtim!FPA7) + diax + tdot(e_pros * Rtim!fpa2 + e_pros * Rtim!FPA6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aji + eisf)
'
'640         CCC4 = diax + "0[>]"
'
'642         e_line = e_line + E2 + CCC4
'
'644         If f_AKYR_benzinadiko = 1 Then
'                ' ελινε timologioy
'646             benz = "?" + M_KODKAYSIMOY + "!" + Replace(Format(sum_posothta, "##0.000"), ",", ".") ' 429.280" ' R?20!ΛΙΤΡΑ
'648             benz2 = ""
'
'650             e_line = "[<]" + AFMPEL + diax + Left(CNull(Rtim("AFM")), 9) + diax + benz + diax + Format(R("Shme"), "YYYYMMDDHHmm")
'652             e_line = e_line + diax + typos_par + diax + benz2 + diax + benz_atim + diax
'654             E2 = tdot(e_pros * Rtim!aj4 + e_pros * Rtim!aj3) + diax + tdot(e_pros * Rtim!aj1 + e_pros * Rtim!AJ7) + diax + tdot(e_pros * Rtim!aj2 + e_pros * Rtim!aj6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aj5) + diax
'656             E2 = E2 + tdot(e_pros * Rtim!FPA4 + e_pros * Rtim!fpa3) + diax + tdot(e_pros * Rtim!FPA1 + e_pros * Rtim!FPA7) + diax + tdot(e_pros * Rtim!fpa2 + e_pros * Rtim!FPA6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aji)
'658             CCC4 = diax + "0[>]"
'660             e_line = e_line + E2 + CCC4
'                'A PERASMA  [<]998481696;028783755;?21!1.000;202406300000;222;;006968;0.00;0.00;10.00;0.00;0.00;0.00;0.00;2.40;0.00;12.40;0[>]
'                ' eline ακυρωτικου
'                ' benz = ""
'
'662             mparat2 = Trim(Left$(parat.Text, 10))
'
'664             If Val(mparat2) > 99999 Then
'666                 benz2 = "#173#" + Right(Trim(mparat2), 6)
'                Else
'668                 benz2 = "#173#" + Right(Trim(mparat2), 5)
'                End If
'
'670             benz = "?" + M_KODKAYSIMOY + "!" + Replace(Format(Abs(sum_posothta), "##0.000"), ",", ".") ' 429.280" ' R?20!ΛΙΤΡΑ
'672             benz2 = "0#173#" + Right(Trim(mparat2), 6) '   "ariumow apodeixis"
'674             benz_atim = Format(GET_NVALUE("SELECT ARITMISI WHERE ID=56"), "000000")   ' ARITMOS AKYROTIKOY
'676             typos_par = 175 '  215
'678             e_pros = -1
'
'680             E2 = tdot(e_pros * Rtim!aj4 + e_pros * Rtim!aj3) + diax + tdot(e_pros * Rtim!aj1 + e_pros * Rtim!AJ7) + diax + tdot(e_pros * Rtim!aj2 + e_pros * Rtim!aj6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aj5) + diax
'682             E2 = E2 + tdot(e_pros * Rtim!FPA4 + e_pros * Rtim!fpa3) + diax + tdot(e_pros * Rtim!FPA1 + e_pros * Rtim!FPA7) + diax + tdot(e_pros * Rtim!fpa2 + e_pros * Rtim!FPA6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aji)
'684             CCC4 = diax + "0[>]"
'
'686             e_AkyrLine = "[<]" + AFMPEL + diax + "" + diax + benz + diax + Format(R("Shme"), "YYYYMMDDHHmm")
'688             e_AkyrLine = e_AkyrLine + diax + typos_par + diax + benz2 + diax + benz_atim + diax
'690             e_AkyrLine = e_AkyrLine + E2 + CCC4
'
'692             If MATIM <> f_matim And Left(f_matim, 1) = "T" Then ' dhladh tipono to akirotiko THW APODEIJIS POY SYNODEYEI TO TIMOLOGIO
'694                 e_pros = 1 ' giati oi ajies einai hdh arnhtikes
'696                 E2 = tdot(e_pros * Rtim!aj4 + e_pros * Rtim!aj3) + diax + tdot(e_pros * Rtim!aj1 + e_pros * Rtim!AJ7) + diax + tdot(e_pros * Rtim!aj2 + e_pros * Rtim!aj6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aj5) + diax
'698                 E2 = E2 + tdot(e_pros * Rtim!FPA4 + e_pros * Rtim!fpa3) + diax + tdot(e_pros * Rtim!FPA1 + e_pros * Rtim!FPA7) + diax + tdot(e_pros * Rtim!fpa2 + e_pros * Rtim!FPA6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aji)
'700                 benz2 = "#173#" + ar_apod ' Right(Trim(mparat2), 6)
'
'702                 e_AkyrLine = "[<]" + AFMPEL + diax + "" + diax + benz + diax + Format(R("Shme"), "YYYYMMDDHHmm")
'704                 e_AkyrLine = e_AkyrLine + diax + typos_par + diax + benz2 + diax + Right(MATIM, 6) + diax
'706                 e_AkyrLine = e_AkyrLine + E2 + CCC4
'
'708                 e_line = e_AkyrLine
'710                 parat.Text = ar_apod
'                    'TELIKO PISTOTIKOY [<]998481696;;?21!1.000;202406300000;175;#173#123456;000768;0.00;0.00;10.00;0.00;0.00;0.00;0.00;2.40;0.00;12.40;0[>]
'                End If
'                'TELIKO TIMOLOGIOY [<]998481696;028783755;?21!1.000;202406300000;222;;006968;0.00;0.00;10.00;0.00;0.00;0.00;0.00;2.40;0.00;12.40;0[>]
'
'            End If
'
'712         If f_is_benzinadiko = 1 And M_KODKAYSIMOY = "****" Then
'714             e_line = ""
'716             MsgBox ("Τιμολόγιο Λιπαντικών")
'718             e_AkyrLine = ""
'            End If
    
        Else  ' casio ======================  CASIO  PROLINE SIGNUP

            '  Dim benz2     As String ' sxetiko παραστατικο που ακυρωνεται
            
720         If f_is_benzinadiko = 1 Then




'               If M_KODKAYSIMOY = "30" Or M_KODKAYSIMOY = "31" Then
'                  If F_meIFESTO = 2 Then
'                     mTher = "#" + Trim(CNull(Rtim("AFM"))) + "*" + Trim(CNull(Rtim("DEH")))
'                  Else
'                      mTher = ""
'                  End If
'
'                Else
'                    mTher = ""
'                End If



                If M_KODKAYSIMOY = "30" Or M_KODKAYSIMOY = "31" And (Rtim!aji) <> 0 Then
                     If F_meIFESTO = 2 Then
                        mTher = "#" + Trim(CNull(Rtim("AFM"))) + "*" + Left(Trim(CNull(Rtim("DEH"))), 9)
                     Else
                          mTher = ""
                     End If
                Else
                    mTher = ""
                End If
                
722             If Len(Trim(PAR_GIA_AKYROSH)) > 3 Then      'ΑΚΥΡΩΤΙΚΟ=====================
724                 typos_par = 215
726                 e_pros = 1 ' ΓΙΑΤΙ ΕΧΕΙ  ΗΔΗ ΑΡΝΗΤΙΚΕΣ ΑΞΙΕΣ  ΤΟ ΑΚΥΡΩΤΙΚΟ -1
728                 benz2 = "#" + Trim(PAR_GIA_AKYROSH) + "#"
                End If
               
                ' einai akyrotiko alla den xero poio akyronei
                ' ΑΥΤΟ ΧΡΕΙΑΖΕΤΑΙ ΣΤΗΝ ΕΠΑΝΕΚΤΥΠΩΣΗ ΠΟΥ ΔΕΝ ΞΕΡΕΙ ΠΟΙΟ ΘΑ ΑΚΥΡΩΣΕΙ
730             If typos_par = 215 And Len(Trim(PAR_GIA_AKYROSH)) < 1 Then
              
732                 If Len(Trim(PAR_GIA_AKYROSH)) < 1 Then
734                     gia_akyr_tim = GET_CVALUE("select SXETIKO from TIM where " + synt + " and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';")
736                     gia_akyr_tim = GET_CVALUE("select ATIM from TIM WHERE ID_NUM=" + gia_akyr_tim)
                  
                        ' e_pros = -1
738                     PAR_GIA_AKYROSH.Caption = GET_CVALUE("SELECT TRIM(C1) FROM PARASTAT WHERE EIDOS='" + Left(gia_akyr_tim, 1) + "'") + "#" + Trim(mID(gia_akyr_tim, 2, f_psifia_atim))
                  
                    End If
           
                End If
           
740             If Len(Trim(PAR_GIA_AKYROSH)) > 1 Then      'ΑΚΥΡΩΤΙΚΟ=====================
742                 benz = ""
744                 benz2 = "0#" + Trim(PAR_GIA_AKYROSH) + "#"
746                 benz_atim = mID(Rtim("ATIM"), 2, 6)
748                 typos_par = 215
750                 e_pros = -1
                Else
752                 benz = "?" + M_KODKAYSIMOY + "!" + Replace(Format(sum_posothta, "##0.000"), ",", ".") + mTher ' "#" + Trim(CNull(Rtim("AFM"))) + "*" + Trim(CNull(Rtim("DEH"))) ' 429.280" ' R?20!ΛΙΤΡΑ
754                 benz2 = ""
                        
                    'benz= 21!; 10#
                        
                    '998481696;;DLD19004842;?20!6.949;202110050732;0005;0027867;0725;173;0;0114051;0.00;0.00;8.06;0.00;0.00;0.00;0.00;1.94;0.00;10.00;0;9CC6FF41D63BC8393C0819912054F562EB4165F6
756                 benz_atim = mID(Rtim("ATIM"), 2, 6)   ' parat.Text   '"672
                End If  'Len(Trim(PAR_GIA_AKYROSH)) > 3 Then      'ΑΚΥΡΩΤΙΚΟ=====================
          
            Else   '               f_is_benzinadiko <> 1 Then
758             benz2 = " "
760             If Len(Trim(PAR_GIA_AKYROSH)) > 3 Then      'ΑΚΥΡΩΤΙΚΟ=====================
762                 typos_par = 215
764                 e_pros = 1 ' ΓΙΑΤΙ ΕΧΕΙ  ΗΔΗ ΑΡΝΗΤΙΚΕΣ ΑΞΙΕΣ  ΤΟ ΑΚΥΡΩΤΙΚΟ -1
766                 benz2 = "#" + Trim(PAR_GIA_AKYROSH) + "#"
                End If
           
                ' PERIPTVSH DIORUVSHS
                ' einai akyrotiko alla den xero poio akyronei
768             If typos_par = 215 Then

                    ' Dim gia_akyr_tim As String
770                 If Len(Trim(PAR_GIA_AKYROSH)) < 1 Then
772                     gia_akyr_tim = GET_CVALUE("select SXETIKO from TIM where " + synt + " and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';")
774                     gia_akyr_tim = GET_CVALUE("select ATIM from TIM WHERE ID_NUM=" + gia_akyr_tim)
                  
776                     e_pros = -1
778                     PAR_GIA_AKYROSH.Caption = GET_CVALUE("SELECT TRIM(C1) FROM PARASTAT WHERE EIDOS='" + Left(gia_akyr_tim, 1) + "'") + "#" + Trim(mID(gia_akyr_tim, 2, f_psifia_atim))
780                     benz2 = "#" + Trim(PAR_GIA_AKYROSH) + "#"
                  
                    End If
           
                End If
           
782             benz = ""
784             benz_atim = mID(Rtim("ATIM"), 2, 6)
            End If  '  f_is_benzinadiko = 1 Then
     
786         If f_AKYR_benzinadiko = 1 And MATIM <> f_matim And Left(f_matim, 1) = "T" Then ' dhladh tipono to akirotiko THW APODEIJIS POY SYNODEYEI TO TIMOLOGIO
                ' ελινε timologioy
788             benz = "?" + M_KODKAYSIMOY + "!" + Replace(Format(sum_posothta, "##0.000"), ",", ".") + mTher ' 429.280" ' R?20!ΛΙΤΡΑ
790             benz2 = ""
                 
                '   e_line = "[<]" + AFMPEL + diax + Left(CNull(Rtim("AFM")), 9) + diax + benz + diax + Format(R("Shme"), "YYYYMMDDHHmm")
                '  e_line = e_line + diax + typos_par + diax + benz2 + diax + benz_atim + diax
                ' E2 = tdot(e_pros * Rtim!aj4 + e_pros * Rtim!aj3) + diax + tdot(e_pros * Rtim!aj1 + e_pros * Rtim!AJ7) + diax + tdot(e_pros * Rtim!aj2 + e_pros * Rtim!aj6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aj5) + diax
                '         E2 = E2 + tdot(e_pros * Rtim!FPA4 + e_pros * Rtim!fpa3) + diax + tdot(e_pros * Rtim!FPA1 + e_pros * Rtim!FPA7) + diax + tdot(e_pros * Rtim!fpa2 + e_pros * Rtim!FPA6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aji)
                '        CCC4 = diax + "0[>]"
                '       e_line = e_line + E2 + CCC4
       
                ' eline ακυρωτικου
                ' benz = ""

792             mparat2 = Trim(Left$(parat.Text, 10))

794             If Val(mparat2) > 99999 Then
796                 benz2 = "#173#" + Right(Trim(mparat2), 6)
                Else
798                 benz2 = "#173#" + Right(Trim(mparat2), 5)
                End If

800             benz = "?" + M_KODKAYSIMOY + "!" + Replace(Format(Abs(sum_posothta), "##0.000"), ",", ".") ' 429.280" ' R?20!ΛΙΤΡΑ
802             benz2 = "0#173#" + Right(Trim(mparat2), 6) '   "ariumow apodeixis"
804             benz_atim = Format(GET_NVALUE("SELECT ARITMISI WHERE ID=56"), "000000")   ' ARITMOS AKYROTIKOY
                'typos_par = 175 '  215
                'e_pros = -1
                
                'E2 = tdot(e_pros * Rtim!aj4 + e_pros * Rtim!aj3) + diax + tdot(e_pros * Rtim!aj1 + e_pros * Rtim!AJ7) + diax + tdot(e_pros * Rtim!aj2 + e_pros * Rtim!aj6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aj5) + diax
                'E2 = E2 + tdot(e_pros * Rtim!FPA4 + e_pros * Rtim!fpa3) + diax + tdot(e_pros * Rtim!FPA1 + e_pros * Rtim!FPA7) + diax + tdot(e_pros * Rtim!fpa2 + e_pros * Rtim!FPA6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aji)
                'CCC4 = diax + "0[>]"
           
                '  e_AkyrLine = "[<]" + AFMPEL + diax + "" + diax + benz + diax + Format(R("Shme"), "YYYYMMDDHHmm")
                '  e_AkyrLine = e_AkyrLine + diax + typos_par + diax + benz2 + diax + benz_atim + diax
                '  e_AkyrLine = e_AkyrLine + E2 + CCC4
                
806             If MATIM <> f_matim And Left(f_matim, 1) = "T" Then ' dhladh tipono to akirotiko THW APODEIJIS POY SYNODEYEI TO TIMOLOGIO
808                 typos_par = 175 '  215
810                 e_pros = -1
                    
                    ' e_pros = 1 ' giati oi ajies einai hdh arnhtikes
                    ' E2 = tdot(e_pros * Rtim!aj4 + e_pros * Rtim!aj3) + diax + tdot(e_pros * Rtim!aj1 + e_pros * Rtim!AJ7) + diax + tdot(e_pros * Rtim!aj2 + e_pros * Rtim!aj6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aj5) + diax
                    ' E2 = E2 + tdot(e_pros * Rtim!FPA4 + e_pros * Rtim!fpa3) + diax + tdot(e_pros * Rtim!FPA1 + e_pros * Rtim!FPA7) + diax + tdot(e_pros * Rtim!fpa2 + e_pros * Rtim!FPA6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aji)
812                 benz2 = ""  ' γιατι δεν μπορω να το ακυρωσω "#173#" + ar_apod ' Right(Trim(mparat2), 6)
                           
                    ' e_AkyrLine = "[<]" + AFMPEL + diax + "" + diax + benz + diax + Format(R("Shme"), "YYYYMMDDHHmm")
                    ' e_AkyrLine = e_AkyrLine + diax + typos_par + diax + benz2 + diax + Right(MATIM, 6) + diax
                    ' e_AkyrLine = e_AkyrLine + E2 + CCC4
                    
                    ' e_line = e_AkyrLine
814                 parat.Text = ar_apod
                
                End If
       
            End If
            '====================================================================
           
            Dim ssFpa0 As Single
816         ssFpa0 = Rtim!aj5
818         If g_Fpa(6) = 0 Then
820             ssFpa0 = ssFpa0 + Rtim!aj6
            End If
822         If g_Fpa(7) = 0 Then
824             ssFpa0 = ssFpa0 + Rtim!AJ7
            End If
826         If g_Fpa(8) = 0 Then
828             ssFpa0 = ssFpa0 + Rtim!aj8
            End If
           
            ' benz = ""
            ' benz_atim = mID(Rtim("ATIM"), 2, 6)
     
830         e_line = "<%SL " + AFMPEL + ";" + Left(CNull(Rtim("AFM")), 9) + ";;" + benz + ";;;;;" + typos_par + ";" + benz2 + ";" + mID(Rtim("ATIM"), 2, 6) + ";"
     
            'E2 = tdot(e_pros * Rtim!aj4) + ";" + tdot(e_pros * Rtim!aj1) + ";" + tdot(e_pros * Rtim!aj2 + e_pros * Rtim!aj6) + ";" + tdot(e_pros * Rtim!aj3) + ";" + tdot(e_pros * Rtim!aj5) + ";"
832         E2 = tdot(e_pros * Rtim!aj4 + e_pros * Rtim!aj3) + diax + tdot(e_pros * Rtim!aj1) + diax + tdot(e_pros * Rtim!aj2 + e_pros * Rtim!aj6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * ssFpa0) + diax
     
            'E2 = E2 + tdot(e_pros * Rtim!FPA4) + ";" + tdot(e_pros * Rtim!fpa1) + ";" + tdot(e_pros * Rtim!fpa2 + e_pros * Rtim!FPA6) + ";" + tdot(e_pros * Rtim!fpa3) + ";" + tdot(e_pros * Rtim!aji)
834         E2 = E2 + tdot(e_pros * Rtim!FPA4 + e_pros * Rtim!fpa3) + diax + tdot(e_pros * Rtim!FPA1) + diax + tdot(e_pros * Rtim!fpa2 + e_pros * Rtim!FPA6) + diax + tdot(e_pros * 0) + diax + tdot(e_pros * Rtim!aji)
     
836         CCC4 = ";;>"
            
            Dim MYD_CAT   As String
            Dim bale_koma As Integer
838         If F_TYPFOROL = 22 Then ' SIGNPRO ΝΕΟΥ ΤΥΠΟΥ ΤΗΕΝ

                ';EUR;1;113.08                                  ;;(6.00,10.50,1,0),(13.00,25.50,2,0),(24.00,36.80,3,0),(36.00,12.50,4,0),(0.00,10.50,5,18);;>
                ' CCC4 = ";EUR;1;" + tdot(e_pros * Rtim!aji) + ";;"

840             bale_koma = 0
           
                '   mAPOLIAsyndesis = 1
                ' End If
          
                ' IF METRHTA
842             MYD_CAT = GET_CVALUE("SELECT ISNULL(MYCATEG,'1') from PARASTAT WHERE EIDOS='" + Left(MATIM, 1) + "'")
                '  mycateg = Right(MYD_CAT, 1)
          
844             MYD_CAT = Right(MYD_CAT, 1)
          
846             If InStr(Rtim!trp, "POS") = 0 Or typos_par = 215 Then  'αν ειναι μετρητα/πιστωση ή ειναι ακυρωτικο
848                 CCC4 = ";;0;0;;"  'οκ
                Else ' (ME KARTA)
850                 CCC4 = ";;1;" + tdot(e_pros * Rtim!aji) + ";;"
                End If

852             If Rtim!aj3 <> 0 Then
854                 CCC4 = CCC4 + "(" + Replace(Format(g_Fpa(3), "#0.00"), ",", ".") + "," + tdot(e_pros * Rtim!aj3) + "," + MYD_CAT + ",0)" ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
856                 bale_koma = 1
                Else
858                 CCC4 = CCC4 + ""
                End If
          
860             If Rtim!aj4 <> 0 Then
862                 CCC4 = CCC4 + "(" + Replace(Format(g_Fpa(4), "#0.00"), ",", ".") + "," + tdot(e_pros * Rtim!aj4) + "," + MYD_CAT + ",0)" ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
864                 bale_koma = 1
                Else
866                 CCC4 = CCC4 + ""
                End If
          
868             If Rtim!aj1 <> 0 Then
870                 If bale_koma = 1 Then CCC4 = CCC4 + "," ' αν εχει προηγηθεί το 6% βαλε ","
872                 CCC4 = CCC4 + "(13.00," + tdot(e_pros * Rtim!aj1) + "," + MYD_CAT + ",0)"    ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
874                 bale_koma = 1
                Else
876                 CCC4 = CCC4 + ""
                End If
          
878             If Rtim!aj2 + Rtim!aj6 <> 0 Then
880                 If bale_koma = 1 Then CCC4 = CCC4 + "," ' αν εχει προηγηθεί το 6% βαλε "," ' αν εχει προηγηθεί το 6%/13% βαλε ","
882                 CCC4 = CCC4 + "(24.00," + tdot(e_pros * (Rtim!aj2 + Rtim!aj6)) + "," + MYD_CAT + ",0)"    ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
884                 bale_koma = 1
                Else
886                 CCC4 = CCC4 + ""
                End If
                    
888             If ssFpa0 <> 0 Then
890                 If bale_koma = 1 Then CCC4 = CCC4 + "," ' αν εχει προηγηθεί το 6% βαλε "," ' αν εχει προηγηθεί το 6%/13% βαλε ","
                    
892                 CCC4 = CCC4 + "(0.00," + tdot(e_pros * ssFpa0) + "," + MYD_CAT + ",7)"    ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
894                 bale_koma = 1
                Else
896                 CCC4 = CCC4 + ""
                End If
          
                ' CCC4 = CCC4 + ";;;>"
898             CCC4 = CCC4 + ";;>"

900             If typos_par = 215 Or typos_par = 175 Or typos_par = 169 Then 'πιστωτικα ακυρωτικα αποδ.επιστρ.λιανικη οχι pos
                Else
           
902                 If InStr(Rtim!trp, "POS") = 0 Then
             
                    Else
904                     If InStr(Rtim!trp, "ΕΤΕ") = 0 And InStr(Rtim!trp, "ETE") = 0 Then  'ΔΕΝ ΕΙΝΑΙ ΕΤΕΡΟΧΡΟΝΙΣΜΕΝΗ ΕΤΕ ΕΛΛΗΝΙΚΑ ΚΑΙ ETE ΞΕΝΑ ΔΟΥΛΕΥΟΥΝ ΚΑΙ ΤΑ 2
906                         If InStr(Rtim!trp, "POS2") > 0 Then
908                             CCC4 = CCC4 + "  <%PS2;1;" + tdot(e_pros * Rtim!aji) + ">"
                            Else
910                             CCC4 = CCC4 + "  <%PS1;1;" + tdot(e_pros * Rtim!aji) + ">"
                            End If
                       
                        Else  'ειναι ετεροχρονισμενη
912                         If InStr(Rtim!trp, "POS2") > 0 Then
914                             CCC4 = CCC4 + "  <%PS2;2;" + tdot(e_pros * Rtim!aji) + ">"
                            Else
916                             CCC4 = CCC4 + "  <%PS1;2;" + tdot(e_pros * Rtim!aji) + ">"
                            End If
                        End If
                     
                    End If
                End If
          
            End If
          
918         If F_TYPFOROL = 23 Then ' οποσ το 22 αλλα σε 2 σειρεσ το ελινε SIGNPRO PAR20  ΝΕΟΥ ΤΥΠΟΥ ΤΗΕΝ

                ';EUR;1;113.08                                  ;;(6.00,10.50,1,0),(13.00,25.50,2,0),(24.00,36.80,3,0),(36.00,12.50,4,0),(0.00,10.50,5,18);;>
                ' CCC4 = ";EUR;1;" + tdot(e_pros * Rtim!aji) + ";;"

920             bale_koma = 0
           
                '   mAPOLIAsyndesis = 1
                ' End If
          
                ' IF METRHTA
922             MYD_CAT = GET_CVALUE("SELECT ISNULL(MYCATEG,'1') from PARASTAT WHERE EIDOS='" + Left(MATIM, 1) + "'")
                '  mycateg = Right(MYD_CAT, 1)
          
924             MYD_CAT = Right(MYD_CAT, 1)
          
926             If InStr(Rtim!trp, "POS") = 0 Or typos_par = 215 Then  'αν ειναι μετρητα/πιστωση ή ειναι ακυρωτικο
928                 CCC4 = ";;0;0;;"  'οκ
                Else ' (ME KARTA)
930                 CCC4 = ";;1;" + tdot(e_pros * Rtim!aji) + ";;"
                End If

932             If Rtim!aj3 <> 0 Then
934                 CCC4 = CCC4 + "(6.00," + tdot(e_pros * Rtim!aj3) + "," + MYD_CAT + ",0)" ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
936                 bale_koma = 1
                Else
938                 CCC4 = CCC4 + ""
                End If
          
940             If Rtim!aj4 <> 0 Then
942                 CCC4 = CCC4 + "(24.00," + tdot(e_pros * Rtim!aj4) + "," + MYD_CAT + ",0)" ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
944                 bale_koma = 1
                Else
946                 CCC4 = CCC4 + ""
                End If
          
948             If Rtim!aj1 <> 0 Then
950                 If bale_koma = 1 Then CCC4 = CCC4 + "," ' αν εχει προηγηθεί το 6% βαλε ","
952                 CCC4 = CCC4 + "(13.00," + tdot(e_pros * Rtim!aj1) + "," + MYD_CAT + ",0)"    ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
954                 bale_koma = 1
                Else
956                 CCC4 = CCC4 + ""
                End If
          
958             If Rtim!aj2 + Rtim!aj6 <> 0 Then
960                 If bale_koma = 1 Then CCC4 = CCC4 + "," ' αν εχει προηγηθεί το 6% βαλε "," ' αν εχει προηγηθεί το 6%/13% βαλε ","
962                 CCC4 = CCC4 + "(24.00," + tdot(e_pros * (Rtim!aj2 + Rtim!aj6)) + "," + MYD_CAT + ",0)"    ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
964                 bale_koma = 1
                Else
966                 CCC4 = CCC4 + ""
                End If
                    
968             If ssFpa0 <> 0 Then
970                 If bale_koma = 1 Then CCC4 = CCC4 + "," ' αν εχει προηγηθεί το 6% βαλε "," ' αν εχει προηγηθεί το 6%/13% βαλε ","
                    
972                 CCC4 = CCC4 + "(0.00," + tdot(e_pros * ssFpa0) + "," + MYD_CAT + ",7)"    ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
974                 bale_koma = 1
                Else
976                 CCC4 = CCC4 + ""
                End If
          
                ' CCC4 = CCC4 + ";;;>"
978             CCC4 = CCC4 + ";;>"
                  
                ' το μηδενιζω γιατι αν ειναι μετρητοις θα κραταει το προηγουμενο
980             DB.Execute "update MEM SET par20=''"
982             If typos_par = 215 Or typos_par = 175 Or typos_par = 169 Then 'πιστωτικα ακυρωτικα αποδ.επιστρ.λιανικη οχι pos
                Else
           
984                 If InStr(Rtim!trp, "POS") = 0 Then
             
                    Else
                    
                        Dim PS1 As String
986                     If InStr(Rtim!trp, "ΕΤΕ") = 0 And InStr(Rtim!trp, "ETE") = 0 Then  'ΔΕΝ ΕΙΝΑΙ ΕΤΕΡΟΧΡΟΝΙΣΜΕΝΗ ΕΤΕ ΕΛΛΗΝΙΚΑ ΚΑΙ ETE ΞΕΝΑ ΔΟΥΛΕΥΟΥΝ ΚΑΙ ΤΑ 2
988                         If InStr(Rtim!trp, "POS2") > 0 Then
990                             PS1 = "<%PS2;1;" + tdot(e_pros * Rtim!aji) + ">"
                            Else
992                             PS1 = "<%PS1;1;" + tdot(e_pros * Rtim!aji) + ">"
                            End If
                       
                        Else  'ειναι ετεροχρονισμενη
994                         If InStr(Rtim!trp, "POS2") > 0 Then
996                             PS1 = "<%PS2;2;" + tdot(e_pros * Rtim!aji) + ">"
                            Else
998                             PS1 = "<%PS1;2;" + tdot(e_pros * Rtim!aji) + ">"
                            End If
                        End If
1000                    DB.Execute "update MEM SET par20='" + PS1 + "'"
                     
                    End If
                End If
          
            End If
          
1002        If F_TYPFOROL = 220 Then ' SIGNPRO ΝΕΟΥ ΤΥΠΟΥ ΜΕ ΟΛΑ ΤΑ ΦΠΑ

                ';EUR;1;113.08                                  ;;(6.00,10.50,1,0),(13.00,25.50,2,0),(24.00,36.80,3,0),(36.00,12.50,4,0),(0.00,10.50,5,18);;>
                ' CCC4 = ";EUR;1;" + tdot(e_pros * Rtim!aji) + ";;"

1004            bale_koma = 0
           
                '   mAPOLIAsyndesis = 1
                ' End If
          
                ' IF METRHTA
1006            MYD_CAT = GET_CVALUE("SELECT ISNULL(MYCATEG,'1') from PARASTAT WHERE EIDOS='" + Left(MATIM, 1) + "'")
                '  mycateg = Right(MYD_CAT, 1)
          
1008            MYD_CAT = Right(MYD_CAT, 1)
          
1010            If InStr(Rtim!trp, "POS") = 0 Or typos_par = 215 Then  'αν ειναι μετρητα/πιστωση ή ειναι ακυρωτικο
1012                CCC4 = ";;0;0;;"  'οκ
                Else ' (ME KARTA)
1014                CCC4 = ";;1;" + tdot(e_pros * Rtim!aji) + ";;"
                End If

                'If Rtim!aj3 <> 0 Then
1016            CCC4 = CCC4 + "(6.00," + tdot(e_pros * Rtim!aj3) + "," + MYD_CAT + ",0)" ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
1018            bale_koma = 1
                'Else
                '   CCC4 = CCC4 + ""
                'End If
                
1020            If bale_koma = 1 Then CCC4 = CCC4 + "," ' αν εχει προηγηθεί το 6% βαλε ","
1022            CCC4 = CCC4 + "(13.00," + tdot(e_pros * Rtim!aj1) + "," + MYD_CAT + ",0)"    ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
1024            bale_koma = 1
          
                'If Rtim!aj4 <> 0 Then
1026            CCC4 = CCC4 + "(24.00," + tdot(e_pros * Rtim!aj4) + "," + MYD_CAT + ",0)" ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
1028            bale_koma = 1
                'Else
                '   CCC4 = CCC4 + ""
                'End If
                    
1030            If ssFpa0 <> 0 Then
1032                If bale_koma = 1 Then CCC4 = CCC4 + "," ' αν εχει προηγηθεί το 6% βαλε "," ' αν εχει προηγηθεί το 6%/13% βαλε ","
1034                CCC4 = CCC4 + "(0.00," + tdot(e_pros * ssFpa0) + "," + MYD_CAT + ",7)"    ' (Συντελεστής ΦΠΑ-Καθαρή Αξία-myDATA τύπος εσόδου-myDATA λόγος απαλλαγής ΦΠΑ)
1036                bale_koma = 1
                Else
1038                CCC4 = CCC4 + ""
                End If
          
                ' CCC4 = CCC4 + ";;;>"
1040            CCC4 = CCC4 + ";;>"

1042            If typos_par = 215 Then
                Else
           
1044                If InStr(Rtim!trp, "POS") = 0 Then
             
                    Else
1046                    If InStr(Rtim!trp, "ΕΤΕ") = 0 And InStr(Rtim!trp, "ETE") = 0 Then
1048                        CCC4 = CCC4 + "  <%PS1;1;" + tdot(e_pros * Rtim!aji) + ">"
                        Else  'ειναι ετεροχρονισμενη
1050                        CCC4 = CCC4 + "  <%PS1;2;" + tdot(e_pros * Rtim!aji) + ">"
                        End If
                     
                    End If
                End If
          
            End If
          
            '21  Γενικό Σύνολο Παρ/κού   (18:2)
            '22  Κωδικός νομίσματος(1)
            '23  ΠΛΗΘΟΣ ΠΛΗΡΩΜΩΝ ΜΕ ΚΑΡΤΑ    (Ν:2)
            '24  ΣΥΝΟΛΟ ΠΛΗΡΩΜΩΝ ΜΕ ΚΑΡΤΑ    (N 18:2)
            '25  ΤΑΥΤΟΠΟΙΗΣΗ ΠΛΗΡΩΜΗΣ ΜΕ ΚΑΡΤΑ   (C 60)
            '26  myDATA  (C 400)
            '27  Random  (C 10)
            '29  ΠΑΗΨΣ
            '
            'Προσθήκη νέων πεδίων μετά τον κωδικό νομίσματος.
            'ΠΛΗΘΟΣ ΠΛΗΡΩΜΩΝ ΜΕ ΚΑΡΤΑ:
            'Το πλήθος πληρωμών με κάρτα που αφορούν το συγκεκριμένο παραστατικό.
            'ΣΥΝΟΛΟ ΠΛΗΡΩΜΩΝ ΜΕ ΚΑΡΤΑ:
            'Στο πεδίο "σύνολο πληρωμών με κάρτα" τοποθετείται και διαβιβάζεται το συνολικό ποσό που πληρώθηκε για το συγκεκριμένο παραστατικό (απόδειξη εσόδου, απόδειξη λιανικών συναλλαγών, τιμολόγιο), με χρεωστική ή πιστωτική κάρτα, από POS που είναι συνδεδεμένο στον ΦΗΜ.
            '
     
1052        e_line = e_line + E2 + CCC4
        End If
  
1054    cID_NUM = str(Rtim("ID_NUM"))
1056    cENTITYMARK = Rtim!entitymark
  
        '  If typos_par = 215 Then
        '      If Len(CNull(Rtim!eline)) > 20 Then
        '          If InStr(Rtim!eline, ";215") > 0 And InStr(Rtim!eline, "#") > 0 Then
        '               e_line = Rtim!eline
        '          End If
        '      End If
        '  End If
  
1058    If f_same_eline = 1 Then
1060        e_line = Rtim!eline
 
        Else
     
1062        Gdb.Execute "UPDATE TIM SET ELINE='" + e_line + "' WHERE ID_NUM=" + str(Rtim("ID_NUM"))
        End If
 
1064    PAR_GIA_AKYROSH.Caption = ""
 
1066    R("LITRA") = Rtim("LITRA")
    
        Dim M_EFK As Single
       
1068    R("EFK") = Rtim("EFK")
1070    M_EFK = nNull(Rtim("EFK"))
 
1072    R("ADT") = Rtim("ADT")
1074    R("METAF") = Rtim("METAF")

1076    R("ATIM") = Rtim("ATIM")

1078    R("DEH") = CNull(Rtim("DEH"))

        '         'BLEPEI AN EINAI PISTOTIKO
        '         If Left(r("ATIM"), 1) = "P" Then
        '            ISPIST = True
        '         Else
        '            ISPIST = False
        '         End If

        On Error GoTo LATOS

1080    R.Update

        Dim mNomisma As String

1082    mNomisma = CNull(Rtim("CURRENCY"))

        On Error Resume Next

Dim OTHERMOVEPURPOSETITLE As String
OTHERMOVEPURPOSETITLE = CNull(Rtim!OTHERMOVEPURPOSETITLE)


1084    Rtim.Close

1086    R.Close

1088    Set R = DB.OpenRecordset("tim")

1090    If IsNull(R("shme")) Or IsNull(R("epo")) Then
1092        printCrystal = -1000
        End If

        'rtim.MoveFirst
        'Do While Not rtim.EOF
        '   R.AddNew
        '   For K = 0 To R.FIELDS.Count - 1
        '     If UCase(left(R(K).Name, 3)) = "EPA" Then
        '         R(K) = left(rtim(R(K).Name), 16)
        '     Else
        '         R(K) = rtim(rtim(K).Name)
        '     End If
        '
        '   Next
        '   R.Update
        '   rtim.MoveNext
        'Loop
        'rtim.Close

        '=============  KENES SEIRES NA GEMISEI EGGTIM  =========

        '==========

        'Set meggtim = db.OpenRecordset("eggtim")
        Dim SAJIA As Single

1094    Set R = DB.OpenRecordset("SELECT SUM([Αξία]) FROM EGGTIM")
1096    SAJIA = R(0)
1098    R.Close

        ' Set R = db.OpenRecordset("SELECT SUM([Αξία]) AS SSAJ,COUNT(*) AS CCOUNT FROM EGGTIM where [Αξία]<>0")
        'SAJIA = R(0)
        ' R.Close

1100    Set R = DB.OpenRecordset("SELECT * FROM TIM")

1102    If False Then ' F_XONDR = 0 Then    'LIANIKH
1104        DB.Execute "UPDATE TIM  SET AJI=" + str(SAJIA) + " , AJIA=" + str(SAJIA)

1106        DB.Execute "UPDATE EGGTIM SET [Tιμ_Μον]=[Tιμ_Μον]*(1+fpa/100)"
            '  db.Execute "UPDATE EGGTIM SET [Αξία]=round([Αξία],2)"

1108        DB.Execute "UPDATE TIM Set aj1=0,aj2=0,aj3=0,aj4=0,aj5=0,AJ6=0,AJ7=0,fpa1=0,fpa2=0,fpa3=0,fpa4=0,FPA6=0,FPA7=0"

        Else

1110        If Abs(SAJIA) - Abs(R("AJ1") + R("AJ2") + R("AJ3") + R("AJ4") + R("AJ5") + R("AJ6") + R("AJ7")) >= 0.07 Then        '0.05
1112            MsgBox "Aσυμφωνία παραστατικού. Δεν θα τυπωθεί το παραστατικό."
1114            MsgBox "συνολο σειρών " + Format((SAJIA + M_EFK), "#########.00") + Chr(13) + "συνολο TIM " + Format((R("AJ1") + R("AJ2") + R("AJ3") + R("AJ4") + R("AJ5") + R("AJ6") + R("AJ7")), "#########.00")

                '  On Error GoTo 0 'NA DIAGRAFEI
                Dim NCO
         
1116            If F_DEBUG = 1 Then
1118                Open "C:\MERCVB\DEBUG.TXT" For Append As #1
1120                Write #1, MATIM + " ΔΙΑΦΟΡΑ,SAJIA, AJ1 ,AJ2,AJ3,AJ4,AJ5"
1122                Write #1, Abs(SAJIA - (R("AJ1") + R("AJ2") + R("AJ3") + R("AJ4") + R("AJ5") + R("AJ6")))
1124                Write #1, SAJIA
1126                Write #1, R("AJ1")
1128                Write #1, R("AJ2")
1130                Write #1, R("AJ3")
1132                Write #1, R("AJ4")
1134                Write #1, R("AJ5")
1136                Close #1

                    On Error Resume Next

1138                DB.Execute "select * into AS" + MATIM + " from EGGtim WHERE FPA>0"
1140                Gdb.Execute "SELECT * INTO AS" + MATIM + " FROM EGGTIM  WHERE ATIM='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'", NCO
                   
                End If
         
                ' ΑΝ ΕΙΝΑΙ ΔΙΟΡΘΩΣΗ ΘΑ ΧΑΣΕΙ ΤΟ ΠΑΡ/ΚΟ ΑΝ ΤΟ ΣΒΗΣΩ
1142            If F_DIOR = 0 Then
                    'ΕΧΕΙ ΣΤΕΙΛΕΙ ΜΥΔΑΤΑ ΔΕΝ ΠΡΕΠΕΙ ΝΑ ΔΙΑΓΡΑΦΕΙ
                    '                Gdb.Execute "delete from TIM    WHERE ATIM='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'", NCO
                    '                Gdb.Execute "delete from EGGTIM WHERE ATIM='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'", NCO
                    '                Gdb.Execute "delete from EGG    WHERE ATIM='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'", NCO
                End If
                
                '   On Error GoTo 0
                '   Gdb.Execute "delete from EGGTIM WHERE ATIM='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'", NCO
              
1144            printCrystal = -1000
                ' End
            End If

        End If

        '=======================ΕΝΗΜΕΡΩΝΩ ΤΟ MEM ΜΕ ΤΑ ΣΤΑΘΕΡΑ ΣΤΟΙΧΕΙΑ ========
        Dim fmem As Recordset

1146    Set fmem = DB.OpenRecordset("mem")
        'fmem.Edit
        ' debug only On Error GoTo 0
1148    DB.Execute "update MEM SET C1='" + Left(mf_paras, Len(mf_paras) - 1) + "'" ' PARASTATIKO
        '1210    fmem("c1") = Left(mf_paras, Len(mf_paras) - 1)    'parastatiko


'If F_METAFORIKES <> 2 Then

1150    DB.Execute "update MEM SET C2='" + SKOPOS.Text + "'"
        '  fmem("c2") = SKOPOS.Text    'skopos

1152    DB.Execute "update MEM SET c3='" + fortosh.Text + "'"  'fortosi

1154    DB.Execute "update MEM SET c4='" + PROORISMOS.Text + "'" 'paradosi

        DB.Execute "update MEM SET FPO5='" + TC5.Text + "'" 'paradosi
        DB.Execute "update MEM SET FPO6='" + TC6.Text + "'" 'paradosi
        DB.Execute "update MEM SET FPO7='" + tc7.Text + "'" 'paradosi



'End If



1156    DB.Execute "update MEM SET banks='" + Trim(Combo4.Text) + "'" 'τροποσ πληρωμης
1158    DB.Execute "update MEM SET pict='" + Time$ + "'"
1160    DB.Execute "update MEM SET par1='" + Text1.Text + "'"

        DB.Execute "update MEM SET par9='" + ENTITYUID + "'"
        DB.Execute "update MEM SET par10='" + authenticationCode + "'"

1162    DB.Execute "update MEM SET par2 ='" + SEIRA.Text + "'"
1164    DB.Execute "update MEM SET par3='" + AYTOKINHTO.Text + "'"
1166    DB.Execute "update MEM SET par19='" + PEL.Recordset("PLAISIO") + "'" 'METAFORIKH ETAIREIA PELATH
1168    DB.Execute "update MEM SET par6='" + Format(DTPicker1.Value, "dd/mm/yyyy") + "'"

1170    DB.Execute "update MEM SET PAR7='" + mNomisma + "'"

1172    If mNomisma = "EUR" Then
1174        DB.Execute "update MEM SET PAR8='€'"
1176    ElseIf mNomisma = "USD" Then
1178        DB.Execute "update MEM SET PAR8='$'"
        Else
1180        DB.Execute "update MEM SET PAR8='' "
        End If

1182    DB.Execute "update MEM SET n1=" + Replace(str(sum_mikta), ",", ".")

1184    DB.Execute "update MEM SET ELINE='" + e_line + "'"


         DB.Execute "update MEM SET ONO1='" + Format(HMERPARAD.Value, "dd/MM/yyyy") + "'"
         DB.Execute "update MEM SET ONO2='" + Format(ORAPARAD.Value, "hh:mm") + "'"
'OTHERMOVEPURPOSETITLE
         DB.Execute "update MEM SET ONO3='" + CNull(OTHERMOVEPURPOSETITLE) + "'" 'Format(ORAPARAD.Value, "hh:mm") + "'"

        'ENTITYUID, AUTHENTICATIONCODE
        DB.Execute "update MEM SET par9='" + Replace(CNull(PEL.Recordset("ENTITYUID")), ",", ".") + "'"
        DB.Execute "update MEM SET par10='" + Replace(CNull(PEL.Recordset("AUTHENTICATIONCODE")), ",", ".") + "'"

1186    DB.Execute "update MEM SET par11='" + Replace(CNull(PEL.Recordset("CH1")), ",", ".") + "'"
1188    DB.Execute "update MEM SET par12='" + Replace(CNull(PEL.Recordset("CH2")), ",", ".") + "'"
1190    DB.Execute "update MEM SET par13='" + Replace(CNull(PEL.Recordset("CH3")), ",", ".") + "'"
1192    DB.Execute "update MEM SET par14='" + Replace(CNull(PEL.Recordset("CH4")), ",", ".") + "'"
1194    DB.Execute "update MEM SET par15='" + Replace(CNull(PEL.Recordset("CH5")), ",", ".") + "'"

1196    If F_IS_BYTIO >= 1 Then

            Dim mOlogr As String
           
1198        mOlogr = Olografos(gVal(str(Round(m_Aji, 2)))) ' Olografos(Round(m_Aji, 2))  'gVal(SynoloG.Caption))
1200        DB.Execute "update MEM SET par15='" + Left(mOlogr, 40) + "'"
1202        DB.Execute "update MEM SET par16='" + mID(mOlogr, 41, 40) + "'"
           
            ' 9999                        mOlogr = Olografos(gVal(SynoloG.Caption))
            ' GoTo 9999
        
        End If
        
        
1204    If mAPOLIAsyndesis = 1 Then
         If Len(cENTITYMARK) > 10 Then
             DB.Execute "update MEM SET par17='" + cENTITYMARK + "'"
         Else
             DB.Execute "update MEM SET par17='??O???? ???S????S?S'"
         End If
         
1206        'DB.Execute "update MEM SET par17='??O???? ???S????S?S'"
        Else
1208        DB.Execute "update MEM SET par17='" + cENTITYMARK + "'"
        End If
        
        
        
        
        
        
        
'1204    If mAPOLIAsyndesis = 1 Then
'1206        DB.Execute "update MEM SET par17='ΑΠΩΛΕΙΑ ΔΙΑΣΥΝΔΕΣΗΣ'"
'        Else
'1208        DB.Execute "update MEM SET par17='" + cENTITYMARK + "'"
'        End If

1210    DB.Execute "update MEM SET EMAIL='" + Replace(CNull(PEL.Recordset("email")), ",", ".") + "'"

1212    DB.Execute "update MEM SET TK='" + Replace(CNull(PEL.Recordset("xrvma")), ",", ".") + "'"

1214    If f_AKYR_benzinadiko = 1 And f_is_benzinadiko = 1 Then
1216        DB.Execute "update MEM SET PS= " + str(GET_NVALUE("SELECT ARITMISI WHERE ID=56")) ' ARITMOS AKYROTIKOY
1218        DB.Execute "update MEM SET FAG9='" + e_AkyrLine + "'" ' GET_CVALUE(" (80 CHFIA)
        End If

1220    DB.Execute "update MEM SET n2=" + Replace(str(sum_posothta), ",", ".")
1222    DB.Execute "update MEM SET n3=" + Replace(str(mTimh), ",", ".")  '") = mTimh
1224    DB.Execute "update MEM SET n4=" + Replace(str(mPoso), ",", ".")  '")' mPoso
        ' DB.Execute "update MEM SET n5") = SYN_EKPT  '=" + Replace(str(mTimh), ",", ".")  '")

1226    DB.Execute "update MEM SET BASTIM_GID=" + Replace(str(F_METRHTAposo), ",", ".")  '")' ") = F_METRHTAposo
1228    DB.Execute "update MEM SET eisekpt=" + Replace(str(SYN_EKPT), ",", ".")  '")'    ") = SYN_EKPT

1230    DB.Execute "update MEM SET MEMO='" + parat.Text + "'"
        'If Combo4.Text = Combo4.List(0) Then ' METRITOIS
1232    DB.Execute "update MEM SET par4='" + Format(PROHG, "#####.00") + "'"
        'Else
        '  DB.Execute "update MEM SET par4") = Format(PEL.Recordset("TYP") - Val(SynoloG.Caption), "#####.00")
        'End If

        'If ISPIST = True Then DB.Execute "update MEM SET PAR4") = " "

        'fmem.Update

        'M_PELONO, M_PELEPA, M_PELDIE, M_SEIRA1
        ' fmem.Edit
1234    DB.Execute "update MEM SET pelono='" + m_pelono + "'"
1236    DB.Execute "update MEM SET pelepa='" + M_PELEPA + "'" 'skopos

1238    DB.Execute "update MEM SET peldie='" + m_peldie + "'"
1240    DB.Execute "update MEM SET SEIRA1='" + M_SEIRA1 + "'" 'skopos
    
1242    DB.Execute "update MEM SET TELPEL='" + m_telpel + "'" 'skopos
        ' DB.Execute "update MEM SET FOTO ='" + stm.Read + "'"
    
1244    fmem.Edit
1246    fmem("FOTO").Value = stm.Read
1248    fmem.Update

        'If IsNull(R("shme")) Or IsNull(R("epo")) Then
        '   printCrystal = -1000
        'End If

1250    If printCrystal < -900 Then
1252        If F_DIOR = 0 Then
                'exei parei mydata den prepei na sbhsei
                'Gdb.Execute "delete from TIM    WHERE ATIM='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'", NCO
                'Gdb.Execute "delete from EGGTIM WHERE ATIM='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'", NCO
                'Gdb.Execute "delete from EGG    WHERE ATIM='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'", NCO
            End If
        End If

        Exit Function

MHNYMA:
1254    HandleError "Par1:PrintCrystal"
1256    printCrystal = -1000

1258    Resume Next

LATOS:
1260    printCrystal = -1000

1262    Resume Next

        '<EhFooter>
        Exit Function

printCrystal_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.printCrystal " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.printCrystal " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function
Private Function tdot(ByVal num As Single)

100     tdot = Replace$(Format(num, "#####0.00"), ",", ".")
End Function

Function GET_CVALUE(query As String) As String
        Dim R As New ADODB.Recordset
   
        On Error Resume Next
100     R.Open query, Gdb, adOpenDynamic, adLockOptimistic
   
102     GET_CVALUE = ""
104     If Not R.EOF Then
106         GET_CVALUE = R(0)
        End If
108     R.Close
   
End Function

Function GET_NVALUE(query As String) As Single
        Dim R As New ADODB.Recordset
   
        On Error Resume Next
100     R.Open query, Gdb, adOpenDynamic, adLockOptimistic
   
102     GET_NVALUE = 0
104     If Not R.EOF Then
106         GET_NVALUE = R(0)
        End If
   
108     R.Close
   
End Function

Function PrintSqlCrystal(MATIM, char_date, ByVal id_num As Long)
   
        Dim mfile As String
        'f_FORM_EKTY

        mfile = "C:\MERCVB\REPORTS\TIMOL" + Format(f_FORM_EKTY, "00") + ".RPT"
        CrystalReport1.ReportFileName = mfile
  
        'CrystalReport1.SelectionFormula = "{TIM.ID_NUM}= " + str(ID_NUM)   ' R!kod + "'"  '   + Str(ID_NUM)
        '  CrystalReport1.Action = 1

        Gdb.Execute "UPDATE TIM SET ATIM2=SUBSTRING(ATIM,1,1) WHERE ID_NUM =" + str(id_num)
        CrystalReport1.SelectionFormula = "{TIM.ID_NUM}= " + str(id_num)
        If F_METAFORIKES = 1 Then
            Exit Function
        End If
    
        '===============================================================================

        '
        'Dim Application As New CRAXdDRT.Application
        'Dim REPORT As New CRAXdDRT.REPORT
        ''Dim mfile As String
        ''f_FORM_EKTY
        '
        'mfile = "C:\MERCVB\REPORTS\TIMOL" + Format(f_FORM_EKTY, "00") + ".RPT"
        '
        'Dim crxApp As CRAXdDRT.Application   ' .Application
        '
        '
        '
        '
        ''Dim REPORT As CRAXdDRT.REPORT '     .Report
        'Set crxApp = New CRAXdDRT.Application
        'Set REPORT = crxApp.OpenReport(mfile)
        ''Declare a Connection Info Object
        'Dim ConnectionInfo As CRAXdDRT.ConnectionProperties
        ''Set the Connection Info to Connection Properties of ‘the table object
        ''Set ConnectionInfo = crxRpt.Database.Tables(1).ConnectionProperties
        '
        '
        'Set ConnectionInfo = REPORT.Database.Tables(1).ConnectionProperties
        'ConnectionInfo.Item("Use DSN Default Properties") = True '"MERCHP"
        '
        '
        '
        'ConnectionInfo.Item("DSN") = "MERCSQL"
        ''ConnectionInfo.Item("DATABASE") = "MERCURY" '"62.103.69.140,65527" 'DELLAGAKIS\SQL17"
        ''ConnectionInfo.Item("Initial Catalog") = "EMP"
        'ConnectionInfo.Item("User ID") = "sa"
        'ConnectionInfo.Item("Password") = "12345678"
        '
        'On Error GoTo 0
        '
        'Dim Q As String
        ''Q = "SELECT EGGTIM.POSO, EGGTIM.MONA, EGGTIM.KODE, EGGTIM.ERGO, EGGTIM.EKPT, EGGTIM.KAU_AJIA, EGGTIM.ONOMA, EGGTIM.KOLA, EGGTIM.PROELEYSH, EGGTIM.EKPT2, EGGTIM.XVRA, EGGTIM.LITRA, PEL.EPO, PEL.POL, PEL.EPA, PEL.TYP, TIM.ID_NUM "
        '
        'Q = "SELECT EGGTIM.*,TIM.*,PEL.*,PARASTAT.* "
        'Q = Q + " FROM  EGGTIM INNER JOIN  TIM ON EGGTIM.ID_NUM=TIM.ID_NUM "
        'Q = Q + " INNER JOIN  PEL ON (TIM.EIDOS=PEL.EIDOS) AND (TIM.KPE=PEL.KOD) "
        'Q = Q + " INNER JOIN PARASTAT ON LEFT(TIM.ATIM,1)=PARASTAT.EIDOS"
        'Q = Q + " Where TIM.ID_NUM = " + str(ID_NUM) + " ORDER BY TIM.ID_NUM"
        'REPORT.SQLQueryString = Q
        '
        '
        '
        '
        '
        'REPORT.PrintOut False, 1 'NoCopies
        'MsgBox "OK"

        Exit Function

        ' Gdb.Execute "UPDATE TIM SET ATIM2=SUBSTRING(ATIM,1,1) WHERE ID_NUM =" + str(ID_NUM)
        '  CrystalReport1.SelectionFormula = "{TIM.ID_NUM}= " + str(ID_NUM)
        If F_METAFORIKES = 1 Then
            Exit Function
        End If

        '------------------------------------------------------------------------
        'δημιουργεί τα αρχεία TIM  , EGGTIM , MEM    ΣΤΟ  c:\mercvb\reports\reports.mdb"
        'ΓΙΑ ΤΗΝ ΕΚΤΥΠΩΣΗ ΤΟΥ ΤΙΜΟΛΟΓΙΟΥ
        '<EhHeader>
        On Error GoTo PrintSqlCrystal_Err

        '</EhHeader>

        Dim sql  As String, fDB As Database, D1 As Date, d2 As Date

        Dim synt As String, DB As Database

        Dim DUM

        Dim REGGTIM As New ADODB.Recordset

        Dim Rtim    As New ADODB.Recordset

        Dim R       As New ADODB.Recordset

        Dim k       As Integer

        Dim KN      As Integer

        '-----------  next lines are OK --------------------------

        On Error GoTo MHNYMA

100     D1 = DTPicker1.Value
102     d2 = DateAdd("d", 1, D1)

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

        '120     sql = "INSERT INTO EGGTIMREP select FPA,ATIM,HME as [Ημερομ],KODE as [Kωδ_Είδ],ONOMA as [Περιγραφή],"
        '
        '130     If F_XONDR = 0 Then    'LIANIKH
        '140         sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],MIK_AJIA as [Αξία],"
        '        Else
        '150         sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],KAU_AJIA as [Αξία],"
        '        End If
        '
        '160     sql = sql + "MONA,EKPT,PROOD_AJ "

        Dim FIELDS As String
104     FIELDS = "EIDOS ,ATIM,POSO,MONA,TIMM,KERDOS,KODE,HME,ERGO,FPA,PROOD,PROOD_AJ,EKPT,KAU_AJIA,MIK_AJIA,ONOMA,MIKTA,KOLA,PELKOD,PROELEYSH,XRE,PIS,APOT,ATIM2,FCURRENCY,ID_NUM,EKPT2,LITRA,EFK "

106     sql = sql + "INSERT  INTO EGGTIMREP(" + FIELDS + ") SELECT " + FIELDS + " from EGGTIM where ID_NUM=" + str(id_num)

        '        On Error Resume Next

108     Gdb.Execute "DROP TABLE TIMREP"
110     Gdb.Execute "DELETE FROM EGGTIMREP"
        ' 190  Gdb.Execute "DROP TABLE EGGTIMREP"
        ' Gdb.Execute sql, k
        'GoTo 10
     
        ' Gdb.Execute "ALTER TABLE EGGTIMREP ADD IDEGG INT IDENTITY(1,1)     "
     
        On Error GoTo PrintSqlCrystal_Err

112     Gdb.Execute sql, k
        'Dim N
        'For N = 1 To 25 - k
        '  Gdb.Execute "INSERT INTO EGGTIM (ATIM) VALUES ("")"
        'Next

114     Gdb.Execute "update EGGTIMREP set FPA=" + str(g_Fpa(1)) + " where FPA=1"
116     Gdb.Execute "update EGGTIMREP set FPA=" + str(g_Fpa(2)) + " where FPA=2"
118     Gdb.Execute "update EGGTIMREP set FPA=" + str(g_Fpa(3)) + " where FPA=3"
120     Gdb.Execute "update EGGTIMREP set FPA=" + str(g_Fpa(4)) + " where FPA=4"
122     Gdb.Execute "update EGGTIMREP set FPA=0 where FPA=5 "

124     R.Open "SELECT COUNT(*) FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly

        Dim MSEIRES As Integer
      
126     MSEIRES = R(0)
128     KN = f_SynSeiresTimologioy - R(0)   'f_SeiresTimologioy - R.RecordCount
130     R.Close

        Dim NCO As Integer

132     For k = 1 To KN
134         MSEIRES = MSEIRES + 1
            ' Gdb.Execute "INSERT INTO EGGTIMREP (PROOD_AJ,FPA,ATIM,[Ημερομ],[Kωδ_Είδ],[Περιγραφή],[Ποσότητα],[Tιμ_Μον],[Αξία],MONA,EKPT ) VALUES (" + str(MSEIRES) + ",2,' ',GETDATE(),' ',' ',0,0,0,0,0)", NCO
        Next

136     synt = " left(ATIM," + str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "'"

        ' αυτο το κομματι δουλευei με odbc  ================================
        'Dim db9 As Database
        'db.Execute "drop table tim"
        'Set db9 = OpenDatabase("", False, False, gConnect)
        'SQL = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL "
        'SQL = SQL + " into tim in 'c:\mercvb\reports\reports.mdb'  from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>=#" + Format(d1, "mm/dd/yyyy") + "# and HME<#" + Format(d2, "mm/dd/yyyy") + "#;"
        'db9.Execute SQL
        'GoTo 22
        ' αυτο το κομματι δουλευei με odbc   ==================================

138     Gdb.Execute "DROP TABLE TIMREP2"
        '========  TIM ==========
140     sql = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL,TIM.EKPT5 AS PROHGYPOL,EFK,LITRA,"
        
142     sql = sql + "SPACE(100) AS MEMO, SPACE(40) AS C1,SPACE(40) AS C2,SPACE(40) AS C3,SPACE(40) AS C4, SPACE(40) AS PAR1,SPACE(40) AS PAR2,SPACE(40) AS PAR3,SPACE(40) AS PAR4,SPACE(40) AS PAR5,SPACE(40) AS PAR6,SPACE(40) AS PICT,SPACE(40) AS BANKS "

144     sql = sql + " INTO TIMREP2 from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"

146     Gdb.Execute sql

        '=============  KENES SEIRES NA GEMISEI EGGTIM  ===================

        'Set meggtim = db.OpenRecordset("eggtim")
        Dim SAJIA As Single

148     R.Open "SELECT SUM(KAU_AJIA) FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly

150     SAJIA = R(0)
152     R.Close

        'Set R = db.OpenRecordset("SELECT * FROM TIM")
154     R.Open "SELECT * FROM TIMREP2", Gdb, adOpenForwardOnly, adLockReadOnly

        Dim MATIM2

156     MATIM2 = R("ATIM")

158     If F_XONDR = 0 Then    'LIANIKH
160         DB.Execute "UPDATE TIMREP2  SET AJI=" + str(SAJIA) + " , AJIA=" + str(SAJIA)

162         DB.Execute "UPDATE EGGTIMREP SET TIMM=TIMM*(1+fpa/100)"
            '  db.Execute "UPDATE EGGTIM SET [Αξία]=round([Αξία],2)"

164         DB.Execute "UPDATE TIMREP2 Set aj1=0,aj2=0,aj3=0,aj4=0,aj5=0,fpa1=0,fpa2=0,fpa3=0,fpa4=0"

        Else

            'If Abs(SAJIA - (R("AJ1") + R("AJ2") + R("AJ3") + R("AJ4") + R("AJ5"))) > 0.005 Then
            '   MsgBox "Aσυμφωνία παραστατικού. Διαγράψτε το παραστατικό και ξαναπεράστε το"

            '  End

            ' End If

        End If

        '=======================ΕΝΗΜΕΡΩΝΩ ΤΟ MEM ΜΕ ΤΑ ΣΤΑΘΕΡΑ ΣΤΟΙΧΕΙΑ ========

        On Error GoTo PrintSqlCrystal_Err

        Dim fmem As New ADODB.Recordset

        ' Gdb.Execute "DROP TABLE MEMREP"
        ' Gdb.Execute "SELECT TOP 1  C1,C2,C3,C4,BANKS,PICT,PAR1,PAR2,PAR3,PAR4,PAR5,MEMO  INTO MEMREP FROM MEM"

        ' fmem.Open "SELECT COUNT(*) FROM MEMREP", Gdb, adOpenDynamic, adLockOptimistic

        ' If fmem(0) = 0 Then
        '      fmem.AddNew
        '      fmem.Update
        '  End If

        ' fmem.Close

166     fmem.Open "SELECT * FROM TIMREP2", Gdb, adOpenDynamic, adLockOptimistic
        'fmem.Edit
168     fmem("c1") = Left(f_paras, Len(f_paras) - 1)    'parastatiko
170     fmem("c2") = SKOPOS.Text    'skopos
172     fmem("c3") = fortosh.Text    'fortosi
174     fmem("c4") = PROORISMOS.Text    'paradosi
176     fmem("banks") = mID(Combo4.Text, 3, 30)  'τροποσ πληρωμης
178     fmem("pict") = Left(Time$, 5)
180     fmem("par1") = Text1.Text
182     fmem("par2") = SEIRA.Text
184     fmem("par3") = AYTOKINHTO.Text
186     fmem("par5") = MATIM2
188     fmem("par6") = Format(DTPicker1.Value, "dd/mm/yyyy")

190     fmem("MEMO") = parat.Text
        'If Combo4.Text = Combo4.List(0) Then ' METRITOIS
192     fmem("par4") = Format(R("PROHGYPOL"), "#####.00")
        'Else
        '   fmem("par4") = Format(PEL.Recordset("TYP") - Val(SynoloG.Caption), "#####.00")
        'End If
194     fmem.Update
        ' CrystalReport1.ReplaceSelectionFormula

        '2021-04-09 CrystalReport1.SelectionFormula = "{TIM.ID_NUM}=10262"           '+ Left(MATIM, f_psifia_atim)

        '14/12/2014 Gdb.Execute "SELECT TIMREP2.* , MEMREP.*  INTO TIMREP FROM TIMREP2,MEMREP"

        Exit Function

MHNYMA:
196     HandleError "Par1:PrintCrystal"

198     Resume Next

        '<EhFooter>
        Exit Function

PrintSqlCrystal_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.PrintSqlCrystal " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.PrintSqlCrystal " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function ektyp_forma(ByVal arxeio As String, ByVal ATIM As String, ByVal hme As Date)

        ' arxeio π.χ. f90.txt  η φόρμα της εκτύπωσης
        '<EhHeader>
        On Error GoTo ektyp_forma_Err

        '</EhHeader>

        ' atim αριθμός παραστατικού
        ' hme  ημερομηνία

        Dim SA, SB

        Dim DUM

        Dim eField(1 To 50), ePict(1 To 50), eProw(1 To 50), ePcol(1 To 50), eFonts(1 To 50)

        Dim mField(1 To 50), mPict(1 To 50), mProw(1 To 50), mPcol(1 To 50), mFonts(1 To 50)

        Dim sField(1 To 50), sPict(1 To 50), sProw(1 To 50), sPcol(1 To 50), sFonts(1 To 50)

        Dim m_No_of_seir, i, k, m, c

        Dim m_seir_synol, m_seir_eid, sf

        Dim a(1 To 250), ar_ped(1 To 250)

        Dim xa(1 To 250), m_m, npic

        Dim m1

        ' dim gm_str(1 To 250)
        Dim DB     As Database

        Dim TIM    As Recordset

        Dim PEL    As Recordset

        Dim eggtim As Recordset

        Dim m_entol_ektyp

        Dim xart11, kod_tim

        Dim ejodos

        Dim e, MM, s

        Dim ektypoths

        Dim SELIDA

        Dim m_syn_row, syn_row

        On Error GoTo MHNYMA

        'cEGGTIM.RecordSource = "eggtim"
100     Set fSCR = CreateObject("MSScriptControl.ScriptControl")

102     cEGGTIM.RecordSource = "select *from eggtim"
104     cEGGTIM.Refresh

106     CTIM.Connect = "Access"
108     CTIM.DatabaseName = "C:\MERCVB\REPORTS\REPORTS.MDB"
110     CTIM.RecordSource = "select *from tim"
112     CTIM.Refresh

114     fSCR.language = "vbscript"
116     fSCR.addObject "cTIM", CTIM
118     fSCR.addObject "cEGGTIM", cEGGTIM
120     fSCR.addObject "cMEM", CMEM
122     fSCR.addObject "mactext", mactext

        'Set db = OpenDatabase(gDir, False, False, gConnect)
        'fSCR.ExecuteStatement "CPEL.DATABASENAME=" + Chr$(34) + gDir + Chr$(34)
        'fSCR.ExecuteStatement "CTIM.DATABASENAME=" + Chr$(34) + gDir + Chr$(34)
        'fSCR.ExecuteStatement "CEGGTIM.DATABASENAME=" + Chr$(34) + gDir + Chr$(34)
        'fSCR.ExecuteStatement "CMEM.DATABASENAME=" + Chr$(34) + gDir + Chr$(34)
        'fSCR.ExecuteStatement "eid.DATABASENAME=" + Chr$(34) + gDir + Chr$(34)

        'fSCR.ExecuteStatement "CTIM.RECORDSOURCE=" + Chr$(34) + "select *from tim where atim='" + atim + "' and hme>=#" + Format(hme, "mm/dd/yyyy") + "#;" + Chr$(34)

        'fSCR.ExecuteStatement "CPEL.RECORDSOURCE=" + Chr$(34) + "select *from pel where eidos='" + "e" + "' and kod='" + CTIM("kpe") + "';" + Chr$(34)
        'fSCR.ExecuteStatement "CEGGTIM.RECORDSOURCE=" + Chr$(34) + "select *from eggtim where atim='" + atim + "' and hme>=#" + Format(hme, "mm/dd/yyyy") + "#;" + Chr$(34)
        'fSCR.ExecuteStatement "CMEM.RECORDSOURCE=" + Chr$(34) + "SELECT *FROM MEM" + Chr$(34)

        'fSCR.ExecuteStatement "CMEM.RECORDSOURCE=" + Chr$(34) + "SELECT *FROM eid" + Chr$(34)

        'Set gEggtim = db.OpenRecordset("select *from eggtim where atim='" + atim + "' and hme>=#" + Format(hme, "mm/dd/yyyy") + "#;")
124     cEGGTIM.Recordset.MoveFirst

        'Set gTim = db.OpenRecordset("select *from tim where atim='" + atim + "' and hme>=#" + Format(hme, "mm/dd/yyyy") + "#;")
        'Set gmem = db.OpenRecordset("select *from mem")
126     CMEM.Recordset.MoveFirst
        'CMEM.Recordset.Move 90

128     CTIM.Recordset.MoveFirst

        'Set gPel = db.OpenRecordset("select *from pel where eidos='" + "e" + "' and kod='" + gTim("kpe") + "';")
        'Set gEid = db.OpenRecordset("eid")

        'On Error GoTo create_index
        'gEid.Index = "eidko"

        'gEid.Seek "=", "ωωωω"

        ' arxeio = "c:\mercvb\f90.txt"
        'c=ltrim(str(select(),2))  '&& ΅¨«ι «¤ §¨ ¦®γ ©«γ¤ ¦§¦ε ™¨ ©΅ζ£¦¬¤

        'load_forma (arxeio)

        'Dim DUM

130     If Len(Dir("c:\PETREL.TXT")) > 1 Then
132         DUM = load_forma(arxeio, 0)
        Else
134         DUM = load_forma(arxeio, 1)
        End If

        'Me.Caption = "load_forma=OK"

136     For i = 1 To 250

138         If Left$(gm_str(i), 10) = "**********" Then

                Exit For

            End If

        Next

        '* ΣΥΝΟΛΙΚΕΣ ΣΕΙΡΕΣ
140     m_No_of_seir = Val(gm_str(i + 1))

        '* ΣΕΙΡΑ ΠΟΥ ΑΡΧΙΖΟΥΝ ΤΑ ΕΙΔΗ
142     m_seir_eid = Val(gm_str(i + 2))

        '* ΣΕΙΡΑ ΠΟΥ ΑΡΧΙΖΟΥΝ ΤΑ ΣΥΝΟΛΑ
144     m_seir_synol = Val(gm_str(i + 3))

        '* entolh gia ton εκτυπωτή
146     m_entol_ektyp = gm_str(i + 4)

        '**************
        ' PRINTER.
        Dim cSelida

148     If F928 = 1 Then
150         cSelida = "Σελίδα "
        Else
152         cSelida = to437("Σελίδα ")
        End If

154     SELIDA = 1

        Dim mfil As String

        'για να σωζει σε διαφορετικα σημεια
156     If f_YPOK = 0 Then
158         mfil = F_File_Timologioy
        Else
160         mfil = F_File2_Timologioy
        End If

162     Open mfil For Output As #1

        '---------------- επικεφαλίδες
        ' On Error GoTo errorl
164     e = 0

166     Do While Len(Trim(gm_f(e + 1))) > 0
168         e = e + 1
            'MM = Trim(gm_f(e))
            '
            'm_m = ExecuteLine("gvar=" + MM)
            'eField(e) = gvar
170         ePict(e) = Trim(gpic(e))
172         eProw(e) = gm_r(e)
174         ePcol(e) = gm_c(e)
176         eFonts(e) = ""  '     && trim( Fonts )
            ' Printer.Print gm_f(e) + Format(eProw(e), "0000") + "=" + Format(ePcol(e), "0000") + "="
        Loop

        '**************** ΣΕΙΡΕΣ ΜΕ ΕΙΔΗ ****************************************
178     m = 0

180     Do While Len(Trim(gm_f(e + m + 2))) > 0
182         m = m + 1

            'MM = Trim(gm_f(M + e + 1))

            'm_m = ExecuteLine("gvar=" + MM)

            'mField(M) = gvar
184         mPict(m) = Trim(gpic(m + e + 1))
186         mProw(m) = gm_r(m + e + 1)
188         mPcol(m) = gm_c(m + e + 1)
            'mFonts(M) = "" '     && trim( Fonts )
            'MM = mPict(M + e + 1)
        Loop

        Dim fcheckonly As Boolean

        '****************  ΣΥΝΟΛΑ  ****************************************
190     s = 0

192     Do While Len(Trim(gm_f(e + m + s + 3))) > 0
194         s = s + 1
            '         MM = Trim(gm_f(M + e + s + 2))
            '         MM = "gvar=" + MM
            '         m_m = ExecuteLine("gvar=" + MM)
            '
            '         sField(s) = gvar
196         sPict(s) = Trim(gpic(m + e + s + 2))
198         sProw(s) = gm_r(m + e + s + 2)
200         sPcol(s) = gm_c(m + e + s + 2)
            '         sFonts(s) = ""  '   && trim( Fonts )
            '         MM = sPict(M + e + s + 2)
            '         'Debug.Print Format(sField(s), MM)
        Loop

        'Printer.FontSize = 10
        ' Printer.FontName = "Courier New"
        ' Printer.Font.Charset = 161

202     If Len(Trim(m_entol_ektyp)) > 0 Then
            'Printer.Print  ' &_entol_ektyp
        End If

        '===================================== ΕΠΙΚΕΦΑΛΙΔΕΣ ===================
204     syn_row = 0
        ' ΕΚΤΥΠΩΣΗ ΕΠΙΚΕΦΑΛΙΔΩΝ

        ' On Error GoTo errorl
206     m1 = 0

208     For k = 1 To e

210         MM = ePict(k)

            'For i = 1 To eProw(K): Printer.Print: Next
212         For i = 1 To eProw(k)

214             If k = 1 And m_No_of_seir = 1000 Then
                    ' ΜΗΝ ΒΑΖΕΙΣ ΚΕΝΗ ΣΕΙΡΑ ΣΤΗΝ ΑΡΧΗ
                Else
216                 Print #1, " "
                End If

            Next

218         syn_row = syn_row + eProw(k)    'συγκεντρώνω τις σειρές

220         MM = mac(Trim(Trim(gm_f(k))))    ': m_m = ExecuteLine("gvar=" + MM)
            'Printer.Print Tab(ePcol(K)); Format(gvar, ePict(K))

222         If eProw(k + 1) > eProw(k) Then    'αν πρόκειται να τυπώσει παρακάτω μην τυπώνεις δίπλα μου
                '   Print #1, Tab(ePcol(K)); Format(mM, ePict(K))
224             Print #1, Tab(ePcol(k)); Format(Left(MM, Len(ePict(k))), ePict(k));
            Else
226             Print #1, Tab(ePcol(k)); Format(Left(MM, Len(ePict(k))), ePict(k));
            End If

228         m1 = eProw(k)

        Next

230     If m_seir_eid >= m_No_of_seir Then
232         GoTo OLD400
        End If

234     For k = 1 To m_seir_eid - syn_row
            Print #1, Space(80)    'Printer.Print:
        Next    'βρισκω την σειρά  των ειδών

        Dim MPER  As String

        Dim lPer  As Integer

        Dim lSeir As Integer

        Dim kkk   As Integer

        Dim lmid  As Integer

236     syn_row = m_seir_eid
        '  Printer.Print
        'On Error GoTo ektyp_forma_Err
238     cEGGTIM.Recordset.MoveFirst

        '=========================================== ΕΙΔΗ ====================
240     Do While Not cEGGTIM.Recordset.EOF

242         DUM = "select *from eid where kod='" + cEGGTIM.Recordset("Kωδ_Είδ") + "'"

244         If IsNull(cEGGTIM.Recordset("Kωδ_Είδ")) Then Exit Do

            'Set gEid = db.OpenRecordset(dum)
246         For k = 1 To m

248             MM = mac(Trim(Trim(gm_f(e + k + 1))))                 ': m_m = ExecuteLine("gvar=" + MM)
                '   Printer.Print Format(gvar, gpic(e + K + 1));

250             If F_POLLES_SEIRES = 1 And Len(MM) > Len(gpic(e + k + 1)) Then
                
252                 lPer = Len(gpic(e + k + 1))
254                 lSeir = 1 + Int(Len(MM) / lPer) ' ποσες σειρες χρειάζομαι
256                 MPER = MM
258                 lmid = 1

260                 For kkk = 1 To lSeir - 1
262                     Print #1, Tab(mPcol(k)); GFORMAT(mID(MM, lmid, lPer), gpic(e + k + 1))
264                     lmid = lmid + lPer - 1
266                     syn_row = syn_row + 1
                    Next

268                 Print #1, Tab(mPcol(k)); GFORMAT(mID(MM, lmid, lPer), gpic(e + k + 1));
            
                Else
270                 Print #1, Tab(mPcol(k)); GFORMAT(MM, gpic(e + k + 1));
                End If

            Next

272         cEGGTIM.Recordset.MoveNext
274         syn_row = syn_row + 1

            ' ΑΝ ΔΕΝ ΧΩΡΑΝΕ ΟΙ ΣΕΙΡΕΣ ΣΤΟ ΤΙΜΟΛΟΓΙΟ ( Η ΣΕΙΡΑ ΠΟΥ ΤΥΠΏΝΕΤΑΙ ΞΕΠΕΡΝΑ ΤΗΝ ΣΕΙΡΑ ΤΩΝ ΣΥΝΟΛΩΝ)
276         If syn_row > m_seir_synol Then
278             Print #1, ""
280             Print #1, cSelida + Format(SELIDA, "##")    ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
282             SELIDA = SELIDA + 1
284             Print #1, Chr(12)    ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                'If SELIDA > 1 Then
                '   Print #1, "Σελίδα " + Format(SELIDA, "##"); ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                'End If
286             GoSub TYPOSE_HEADER

            End If

        Loop

288     npic = Space(120)

290     For k = 1 To m_seir_synol - syn_row: Print #1,: Next  'βρισκω την σειρά  των synolon

292     For k = 1 To s
294         MM = Trim(gm_f(m + e + k + 2))
            ' m_m = ExecuteLine("gvar=" + MM)

296         MM = mac(Trim(gm_f(m + e + k + 2)))

298         For i = 1 To sProw(k): Print #1,: Next
            'Printer.Print Format(gvar, gpic(M + e + K + 2));
300         Print #1, Tab(sPcol(k)); GFORMAT(MM, gpic(m + e + k + 2));
        Next

        '       If SELIDA > 1 Then
        '                Print #1, "Σελίδα " + Format(SELIDA, "##") ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
        '       End If

OLD400:

302     Close #1

304     cEGGTIM.Recordset.Close

306     Set fSCR = Nothing

308     If f_YPOK = 0 Then
310         If F_Type_File_Timologioy >= 1 Then
312             If F928 = 1 Then
314                 If F_Type_File_Timologioy = 2 Then
316                     MDIForm1.Caption = "PRINTER.PRINT"
318                     DUM = EKTYP_ME_PRINT(mfil)

320                 ElseIf F_Type_File_Timologioy = 3 Then
322                     MDIForm1.Caption = "TYPE"
324                     DUM = Shell("C:\MERCVB\TYP.BAT ", vbMinimizedFocus)

326                 ElseIf F_Type_File_Timologioy = 4 Then
328                     MDIForm1.Caption = "TYPE-EXE"
330                     DUM = Shell("C:\MERCVB\TYPE.EXE ", vbMinimizedFocus)

332                 ElseIf F_Type_File_Timologioy = 5 Then
334                     MDIForm1.Caption = "TYPE me CMD.EXE"
336                     DUM = Shell("cmd.exe /c c:\typ.bat", vbMinimizedFocus)

338                 ElseIf F_Type_File_Timologioy = 6 Then
340                     MDIForm1.Caption = "TYPE me 5sec delay------------"
342                     MILSEC 15000
344                     DUM = Shell("C:\MERCVB\TYP.BAT ", vbMinimizedFocus)
                       
346                 ElseIf F_Type_File_Timologioy = 5 Then
348                     MDIForm1.Caption = "TYPE me CMD.EXE"
350                     DUM = Shell("cmd.exe /c c:\typ.bat", vbMinimizedFocus)

352                 ElseIf Len(Dir("c:\marietos.txt")) > 1 Then
                        ' kanei "type c:\print"  meso dos
354                     DUM = Shell("c:\mercvb\notedos.exe /p " + mfil, vbMaximizedFocus)
356                 ElseIf Len(Dir("c:\AGIOS.txt")) > 1 Then
                        ' kanei "type c:\print"  meso dos
358                     DUM = Shell("c:\mercvb\ff.exe", vbMaximizedFocus)

360                 ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then
                        ' kanei "type c:\print"  meso dos
362                     DUM = EKTYP_ME_PRINT(mfil)
                        ' DUM = Shell("c:\mercvb\ff.exe", vbMaximizedFocus)
                    Else
364                     DUM = Shell("c:\mercvb\notepad.exe /p " + mfil, vbMaximizedFocus)
                    End If

                Else  ' <>928

366                 If F_Type_File_Timologioy = 2 Then
368                     DUM = EKTYP_ME_PRINT(mfil)
370                 ElseIf F_Type_File_Timologioy = 3 Then
372                     MDIForm1.Caption = "TYPE"
374                     DUM = Shell("C:\MERCVB\TYP.BAT ", vbMinimizedFocus)
                     
376                 ElseIf F_Type_File_Timologioy = 5 Then
378                     MDIForm1.Caption = "TYPE me CMD.EXE"
380                     DUM = Shell("cmd.exe /c c:\typ.bat", vbMinimizedFocus)
                     
382                 ElseIf F_Type_File_Timologioy = 6 Then
384                     MDIForm1.Caption = "TYPE me 5sec delay------------"
386                     MILSEC 4000
388                     DUM = Shell("C:\MERCVB\TYP.BAT ", vbMinimizedFocus)

390                 ElseIf F_Type_File_Timologioy = 4 Then
392                     MDIForm1.Caption = "TYPE-EXE"
394                     DUM = Shell("C:\MERCVB\TYPE.EXE ", vbMinimizedFocus)

396                 ElseIf Len(Dir("c:\FAKKAV.TXT")) > 1 Then
398                     DUM = EKTYP_ME_PRINT(mfil)
400                 ElseIf Len(Dir("C:\OLYMPIAN.TXT")) > 1 Then
402                     DUM = Shell("C:\TYP.BAT ", vbMinimizedFocus)

404                 ElseIf Len(Dir("c:\PETREL.TXT")) > 1 Then
406                     Open "c:\PETREL.TXT" For Input As #1
408                     Input #1, SA
410                     SB = Val(SA)
412                     Close #1
414                     DUM = EKTYP_ME_serial(mfil, SB)

                        ' DUM = Shell("C:\T.BAT ", vbMinimizedFocus)
                    Else
416                     DUM = Shell("C:\MERCVB\PRINT.EXE " + Trim(mfil), vbMinimizedFocus)
                    End If
                End If    '928

            Else  'If F_Type_File_Timologioy >= 1 Then

418             If Len(Dir("c:\KENO1.txt")) > 1 Then
                    'MILSEC 8000
420                 MILSEC 1000

422                 Do While Len(Dir(mfil)) > 1
424                     MILSEC 1000
                    Loop

426                 FileCopy "C:\KENO1.TXT", mfil
428                 FileCopy "C:\KENO1.TXT", "C:\PP2"

                    Exit Function

430             ElseIf Len(Dir("c:\DION2.EXE")) > 1 Then
                    ' DUM = EKTYP_ME_PRINT(mfil)
432                 MILSEC 5000
434                 FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
                    '      ElseIf Len(Dir("C:\SOKRATIS.TXT")) > 1 Then
                    '        ' DUM = EKTYP_ME_PRINT(mfil)
                    '
                    '         'FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
                    '
                    '         Printer.Print "     "
                    '         Printer.Print "     "
                    '         Printer.Print "     "
                    '
                    '         Printer.Print "     "
                    '         Printer.Print "     "
                    '         Printer.Print "     "
                    '
                    '         Printer.Print "     "
                    '         Printer.Print "     "
                    '         Printer.Print "     "
                    '         Printer.EndDoc

436             ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then

                    'MILSEC 8000
438                 Do While Len(Dir("c:\FOROL\FOROL.txt")) > 1
440                     MILSEC 500
                    Loop

442                 FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
                    ' kanei "type c:\print"  meso dos
                    '            DUM = EKTYP_ME_PRINT(mfil)

                End If
            End If

        Else

            '-----------------------------------------------
444         If F_Type2_File_Timologioy >= 1 Then
446             If F_Type2_File_Timologioy = 2 Then
448                 DUM = EKTYP_ME_PRINT(mfil)

450             ElseIf F_Type2_File_Timologioy = 3 Then
452                 MDIForm1.Caption = "TYPE"
454                 DUM = Shell("C:\MERCVB\TYP.BAT ", vbMinimizedFocus)
456             ElseIf F_Type2_File_Timologioy = 4 Then
458                 MDIForm1.Caption = "TYPE"
460                 DUM = Shell("C:\MERCVB\TYP.EXE ", vbMinimizedFocus)

462             ElseIf F928 = 1 Then
                Else  ' <>928
464                 DUM = Shell("C:\MERCVB\PRINT.EXE " + mfil, vbMinimizedFocus)
                End If

            Else

466             If Len(Dir("c:\DION2.EXE")) > 1 Then
                    ' DUM = EKTYP_ME_PRINT(mfil)
468                 MILSEC 5000
470                 FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
472             ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then
                End If
            End If
        End If

        Exit Function

create_index:

        'db.Execute "CREATE INDEX eidko ON eid(kod);"
474     Resume Next

errorl:
476     MsgBox str(Erl) + "MM=" + MM    'RECOVER USING objErrorInfo
        'sfalma = 1
        'MsgBox "Ά¦ª ©«¦ §›ε¦ " + Field + "  £.§›ε¦¬ " + Str(i)
        '    Resume Next
478     Close #1
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
480     syn_row = 0
        ' ΕΚΤΥΠΩΣΗ ΕΠΙΚΕΦΑΛΙΔΩΝ

482     m1 = 0

484     For k = 1 To e
486         MM = ePict(k)

488         For i = 1 To eProw(k)

490             If k = 1 And m_No_of_seir = 1000 Then
                    ' ΜΗΝ ΒΑΖΕΙΣ ΚΕΝΗ ΣΕΙΡΑ ΣΤΗΝ ΑΡΧΗ
                Else
492                 Print #1, " "
                End If

            Next

494         syn_row = syn_row + eProw(k)    'συγκεντρώνω τις σειρές
496         MM = mac(Trim(Trim(gm_f(k))))    ': m_m = ExecuteLine("gvar=" + MM)

498         If eProw(k + 1) > eProw(k) Then    'αν πρόκειται να τυπώσει παρακάτω μην τυπώνεις δίπλα μου
500             Print #1, Tab(ePcol(k)); Format(Left(MM, Len(ePict(k))), ePict(k));
            Else
502             Print #1, Tab(ePcol(k)); Format(Left(MM, Len(ePict(k))), ePict(k));
            End If

504         m1 = eProw(k)
        Next

506     If SELIDA > 1 Then
508         Print #1, cSelida + Format(SELIDA, "##")    ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
        End If

510     For k = 1 To m_seir_eid - syn_row - 1
            Print #1, Space(80)    'Printer.Print:
        Next    'βρισκω την σειρά  των ειδών

512     syn_row = m_seir_eid
514     Return

MHNYMA:
516     HandleError "Par1:ektypforma"

518     Resume Next

        '<EhFooter>
        Exit Function

ektyp_forma_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.ektyp_forma " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.ektyp_forma " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function EKTYP_ME_PRINT(mfil)

        '<EhHeader>
        On Error GoTo EKTYP_ME_PRINT_Err

        '</EhHeader>
        Dim a$

100     Printer.FontName = "Arial"
102     Printer.FontSize = 10
104     Printer.FontBold = True

106     Open "c:\mercvb\printing.txt" For Output As #11
108     Open mfil For Input As #1

110     Do While Not EOF(1)
112         Line Input #1, a$
114         Printer.Print a$
116         Print #11, a$
        Loop

118     Printer.EndDoc
120     Close #1
122     Close #11

        ' Printer.EndDoc

        '<EhFooter>
        Exit Function

EKTYP_ME_PRINT_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.EKTYP_ME_PRINT " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.EKTYP_ME_PRINT " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function EKTYP_ME_serial(mfil, Com)

        '<EhHeader>
        On Error GoTo EKTYP_ME_serial_Err

        '</EhHeader>
        Dim a$

        'Dim com1

100     MSComm1.CommPort = Com

102     MSComm1.PortOpen = True

104     Open mfil For Input As #1

106     Do While Not EOF(1)
108         Input #1, a$

110         If IsNull(a$) Then a$ = ""
112         MSComm1.output = a$ + Chr(10)
114         MILSEC 100
        Loop

116     Close #1
        '  Printer.EndDoc
118     MSComm1.PortOpen = False

        '<EhFooter>
        Exit Function

EKTYP_ME_serial_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.EKTYP_ME_serial " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.EKTYP_ME_serial " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function mac(ByVal s As String) As String

        'δινω ενα στρινγκ και μου επιστρέφει μία τιμή
        '<EhHeader>
        On Error GoTo mac_Err

        '</EhHeader>
        Dim X

        'On Error GoTo mac_Err
100     mactext.Text = "   "

        On Error GoTo err2

102     X = "MACtEXT.Text = " + s
104     fSCR.ExecuteStatement X

106     If F928 = 1 Then
108         mac = mactext.Text
        Else
110         mac = to437(mactext.Text)
        End If

        Exit Function

err2:
112     Me.Caption = X
114     mac = "  "

116     Resume Next

        '<EhFooter>
        Exit Function

mac_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.mac " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.mac " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub AYTOKINHTO_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo AYTOKINHTO_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
102         KeyAscii = 0    'suppress the beep
104         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

AYTOKINHTO_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.AYTOKINHTO_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.AYTOKINHTO_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text5_LostFocus()

        '<EhHeader>
        On Error GoTo Text5_LostFocus_Err

        '</EhHeader>
        Dim k As Integer

        Dim genEkpt

100     genEkpt = gVal(Text5.Text)

        Dim MEKPT, MDOR

        Dim ANS

        Dim telikh_timh As Integer

102     telikh_timh = 0

        On Error Resume Next

104     If Left(f_pelMERC, 3) = "SOK" And genEkpt > 0 Then
106         MEKPT = SynoloG.Caption * genEkpt / 100
108         ANS = MsgBox("ΕΚΠΤΩΣΗ " + Format(MEKPT, "###.00") + Chr(13) + "ΔΩΡΟ " + Format(2 * MEKPT, "###.00") + Chr(13) + " ΝΑ ΠΡΟΧΩΡΓΗΣΩ ΣΤΗΝ ΕΚΠΤΩΣΗ; Ν/Ο", vbYesNo)

110         If ANS = vbNo Then

                Exit Sub

            End If
        End If

112     If Left(LTrim(Text5.Text), 1) <> "*" Then
114         If genEkpt > 0 And genEkpt < 100 Then

                ' Text5.Text = genEkpt
116             For k = 1 To Grid1.rows - 1

118                 If Val(Grid1.TextMatrix(k, f_p)) <> 0 Then
120                     Grid1.TextMatrix(k, f_e) = (genEkpt + Val(Grid1.TextMatrix(k, f_e))) - (genEkpt * Val(Grid1.TextMatrix(k, f_e))) / 100
122                     Grid1.row = k: FindSynolo
                    End If

                Next

            End If
        End If

        Dim POS2

        Dim telPOSO As Single, log As Double

        Dim log2, ekpt2

        Dim Zhtoymeno As Single

        Dim mSynoloG
        
124     If Left(LTrim(Text5.Text), 1) = "*" Then
126         telikh_timh = 1
128         telPOSO = gVal(mID(LTrim(Text5.Text), 2, 10)) 'αιτουμενο ποσο
130         Zhtoymeno = telPOSO
            ' log = telPOSO / Val(Replace(SynoloG, ",", ".")) 'συντελεστης σμίκρυνσης
            ' log = telPOSO / Val(Replace(SynoloG, ",", "."))
            
132         mSynoloG = gVal(SynoloG)  ' αυτο που είχε πριν λίγο σαν τελικο ποσό
134         log = telPOSO / mSynoloG
            
136         If telPOSO < mSynoloG Then
138             log2 = Round(mSynoloG - telPOSO, 2)   'συντελεστης σμίκρυνσης
140             ekpt2 = log2 / (mSynoloG * 100)
            End If

            Dim arxSynoloG As Single

142         arxSynoloG = mSynoloG

144         For k = 1 To Grid1.rows - 1

146             If Val(Grid1.TextMatrix(k, F_T)) <> 0 Then

148                 Grid1.TextMatrix(k, f_e) = Grid1.TextMatrix(k, f_e) + Round(ekpt2, 2)   ' + Val(Grid1.TextMatrix(k, f_e))) - (EKPT2 * Val(Grid1.TextMatrix(k, f_e))) / 100
                    ' Grid1.TextMatrix(k, F_T) = Round(log * Val(Grid1.TextMatrix(k, F_T)), 2)
150                 Grid1.row = k: FindSynolo
                End If

            Next
            
            'νεο συνολο που προέκυψε μετα το πρώτο "κουρεμα"
152         mSynoloG = gVal(SynoloG)  ' αυτο που είχε πριν λίγο σαν τελικο ποσό

            Dim maxT As Single

154         maxT = 0    'βρισκω την μεγαλύτερη τιμή για να προσθέσω +0.01€

            Dim colmaxt As Long

156         colmaxt = 0    ' σε ποια σειρά βρέθηκε

158         If Abs(telPOSO - mSynoloG) >= 0.005 Then
160             log = telPOSO / mSynoloG  'epanypologismos συντελεστη σμίκρυνσης

162             For k = 1 To Grid1.rows - 1

164                 If gVal(Grid1.TextMatrix(k, F_T)) <> 0 Then
166                     Grid1.TextMatrix(k, F_T) = Round(log * gVal(Grid1.TextMatrix(k, F_T)), 3)

168                     If Grid1.TextMatrix(k, F_T) > 0 Then
170                         maxT = gVal(Grid1.TextMatrix(k, F_T))
172                         colmaxt = k
                        End If

174                     Grid1.row = k: FindSynolo
                    End If

                Next

            End If

            'νεο συνολο που προέκυψε μετα το πρώτο "κουρεμα"
176         mSynoloG = gVal(SynoloG)  ' αυτο που είχε πριν λίγο σαν τελικο ποσό
            'diaf = gVal(diaf)

178         Text5.Text = "    "

            Dim diaf

180         diaf = telPOSO - mSynoloG

182         If Abs(diaf) >= 0.01 Then   'βαζω την διαφορά στην τιμή του μεγαλύτερου είδους

184             If gVal(Grid1.TextMatrix(colmaxt, f_p)) > 0 Then
186                 Grid1.TextMatrix(colmaxt, F_T) = Round(gVal(Grid1.TextMatrix(colmaxt, F_T)) + (Round(Round(diaf, 2) / ((100 + g_Fpa(Grid1.TextMatrix(colmaxt, f_f))) / 100), 3)) / gVal(Grid1.TextMatrix(colmaxt, f_p)), 4)
                End If

188             FindSynolo
            End If

            'νεο συνολο που προέκυψε μετα το πρώτο "κουρεμα"
190         mSynoloG = Val(Replace(SynoloG, ",", "."))  ' αυτο που είχε πριν λίγο σαν τελικο ποσό

            '510         diaf = telPOSO - mSynoloG
            '520         If Abs(diaf) >= 0.005 Then
            '530             Grid1.TextMatrix(colmaxt, F_T) = gVal(Grid1.TextMatrix(colmaxt, F_T)) + Round(Round(diaf, 2) / ((100 + g_Fpa(Grid1.TextMatrix(colmaxt, f_f))) / 100), 3)
            '540             FindSynolo
            '            End If
            '
            '550         diaf = telPOSO - mSynoloG
            '560         If Abs(diaf) >= 0.005 Then
            '570             Grid1.TextMatrix(colmaxt, F_T) = Grid1.TextMatrix(colmaxt, F_T) + diaf
            '580             f_psifiaAjias = 3
            '590             FindSynolo
            '600             f_psifiaAjias = 2
            '            End If

192         If Abs(mSynoloG - Zhtoymeno) < 0.03 Then
194             TelFpa.Text = Replace(Format(Zhtoymeno - mSynoloG, "###.0000"), ",", ".")
196             SynoloG = Format(Zhtoymeno, "####0.00")
            
            Else
198             TelFpa = 0
            End If
        
200         FindSynolo
        
        End If   '    If Left(LTrim(Text5.Text), 1) = "*" Then

        '<EhFooter
        
202     Text5.Text = ""
        
        Exit Sub

Text5_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Text5_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
204     SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Text5_LostFocus " & "at line " & Erl

206     Resume Next

        '</EhFooter>

End Sub

'Function gVal(A As String) As Single
'
'        '<EhHeader>
'        On Error GoTo gVal_Err
'
'        '</EhHeader>
'        Dim KOMA
'
'        ' KOMA = InStr(A, ","): If KOMA > 0 Then Mid$(A, KOMA, 1) = "."
'100     A = Replace(A, "*", "")
'110     A = Replace(A, ",", ".")
'120     gVal = Val(A) ' τελικο ποσο που ζητάω
'
'        '<EhFooter>
'        Exit Function
'
'gVal_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.par1.gVal " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.gVal " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Function

Function toascii(ByVal arxeio As String, moutput As String)

        ' arxeio π.χ. f90.txt  η φόρμα της εκτύπωσης    host.txt to arxeio me ejagogi ascii
        '<EhHeader>
        On Error GoTo toascii_Err

        '</EhHeader>

        Dim DUM

        Dim eField(1 To 50), ePict(1 To 50), eProw(1 To 50), ePcol(1 To 50), eFonts(1 To 50)

        Dim mField(1 To 50), mPict(1 To 50), mProw(1 To 50), mPcol(1 To 50), mFonts(1 To 50)

        Dim sField(1 To 50), sPict(1 To 50), sProw(1 To 50), sPcol(1 To 50), sFonts(1 To 50)

        Dim m_No_of_seir, i, k, m, c

        Dim m_seir_synol, m_seir_eid, sf

        Dim a(1 To 250), ar_ped(1 To 250)

        Dim xa(1 To 250), m_m, npic

        Dim m1

        ' dim gm_str(1 To 250)
        Dim DB     As Database

        Dim TIM    As Recordset

        Dim PEL    As Recordset

        Dim eggtim As Recordset

        Dim m_entol_ektyp

        Dim xart11, kod_tim

        Dim ejodos

        Dim e, MM, s

        Dim ektypoths

        Dim m_syn_row, syn_row

        Dim CDOK

100     Set fSCR = CreateObject("MSScriptControl.ScriptControl")

102     DOK.ConnectionString = gConnect
104     DOK.RecordSource = "select *from EID  WHERE MON IN ('ΚΙΛ','KIΛ') OR LEFT(ERG,2)='" + F_2CHFIA_ZYGIZOMENA + " ' ORDER BY ONO "    '    LEFT(ERG,3)='210' "
106     DOK.Refresh

108     CTIM.Connect = "Access"
110     CTIM.DatabaseName = "C:\MERCVB\REPORTS\REPORTS.MDB"
112     CTIM.RecordSource = "select *from tim"
114     CTIM.Refresh

116     fSCR.language = "vbscript"
118     fSCR.addObject "cTIM", CTIM
120     fSCR.addObject "DOK", DOK    'cEGGTIM
122     fSCR.addObject "cMEM", CMEM
124     fSCR.addObject "mactext", mactext

126     DOK.Recordset.MoveFirst

128     CMEM.Recordset.MoveFirst

130     CTIM.Recordset.MoveFirst

        'Dim DUM
132     DUM = load_forma(arxeio, 1)

        '-----------    GLOBAL ΜΕΤΑΒΛΗΤΕΣ  ----------------
        'gm_str() οι σειρές του f99.txt
        'gpic(i) το στρινγκ του πεδίου
        'gm_r(i),gm_c(i) σειρά και στήλη του πεδίου
        'gm_f(i) το όνομα του πεδίου

        'gm_str(6)="     ^XXXXXXXXXXXXXXXXXXXXXXXXXXX               ^xxxxxxxx    @XXXXXXXX  ~ono_par  ~SUBS(tim->atim,2,5)  ~TIM->HME
        '        gpic(2)='XXXXXXXXXXXXXXXXXXXXXXXXXXX
        'gm_r(2)=5  gm_c(2)=6    gm_f(2)=ono_par
        '=================================

134     Open moutput For Output As #1

        Dim S2 As String

136     DOK.Recordset.MoveFirst
138     Do While Not DOK.Recordset.EOF
140         S2 = ""

142         For k = 1 To 4
144             MM = mac(Trim(Trim(gm_f(k))))                   ': m_m = ExecuteLine("gvar=" + MM)
146             S2 = S2 + ALIGNFORMAT(MM, gpic(k))    '
            Next

148         Print #1, S2

150         DOK.Recordset.MoveNext
        Loop

152     Close #1

154     DOK.Recordset.Close

156     Set fSCR = Nothing

        '<EhFooter>
        Exit Function

toascii_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.toascii " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.toascii " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function ALIGNFORMAT(MM, PIC)

        ' RIGHT ALIGNMENT
        '<EhHeader>
        On Error GoTo ALIGNFORMAT_Err

        '</EhHeader>
100     If Left(PIC, 1) = "#" Then
102         ALIGNFORMAT = Right(Space(Len(PIC)) + MM, Len(PIC))
        Else
            ' LEFT ALIGNMENT
104         ALIGNFORMAT = Left(MM + Space(Len(PIC)), Len(PIC))
        End If

        '<EhFooter>
        Exit Function

ALIGNFORMAT_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.ALIGNFORMAT " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.ALIGNFORMAT " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Sub AdjustDataGridColumns(intRecord As Integer, _
                          intField As Integer, _
                          Optional AccForHeaders As Boolean)

        '<EhHeader>
        On Error GoTo AdjustDataGridColumns_Err

        '</EhHeader>

100     GridEidon.SetFocus

        'Exit Sub

        'This procedure will adjust DataGrids column width
        'based on longest field in underlying source

        'GRIDEIDON = DataGrid
        'EID = Adodc control
        'intRecord = Number of record
        'intField = Number of field
        'AccForHeaders = True or False

        Dim row      As Long, Col As Long

        Dim Width    As Single, maxWidth As Single

        Dim saveFont As StdFont, saveScaleMode As Integer

        Dim cellText As String

        'If number of records = 0 then exit from the sub
102     If intRecord = 0 Then Exit Sub
        'Save the form's font for DataGrid's font
        'We need this for form's TextWidth method
104     Set saveFont = GridEidon.Parent.Font
106     Set GridEidon.Parent.Font = GridEidon.Font
        'Adjust ScaleMode to vbTwips for the form (parent).
108     saveScaleMode = GridEidon.Parent.ScaleMode
110     GridEidon.Parent.ScaleMode = vbTwips
        'Always from first record...
112     EID.Recordset.MoveFirst
114     maxWidth = 0

        'We begin from the first column until the last column
        Dim MM

116     If intField - 1 > 3 Then MM = 3 Else MM = intField - 1

        Dim FF

118     If intRecord - 1 > 20 Then FF = 20 Else FF = intRecord - 1

120     For Col = 0 To MM    'intField - 1
122         EID.Recordset.MoveFirst

            'Optional param, if true, set maxWidth to
            'width of GRIDEIDON.Parent
124         If AccForHeaders Then
126             maxWidth = GridEidon.Parent.TextWidth(GridEidon.columns(Col).Text) + 200
            End If

            'Repeat from first record again after we have
            'finished process the last record in
            'former column...
128         EID.Recordset.MoveFirst

130         For row = 0 To FF    ' intRecord - 1

                'Get the text from the DataGrid's cell
132             If intField = 1 Then
                Else  'If number of field more than one
134                 cellText = GridEidon.columns(Col).Text
                End If

                'Fix the border...
                'Not for "multiple-line text"...
136             Width = GridEidon.Parent.TextWidth(cellText) + 200

                'Update the maximum width if we found
                'the wider string...
138             If Width > maxWidth Then
140                 maxWidth = Width
142                 GridEidon.columns(Col).Width = maxWidth
                End If

                'Process next record...
144             EID.Recordset.MoveNext
146         Next row

            'Change the column width...
148         GridEidon.columns(Col).Width = maxWidth    'kolom terakhir!
150     Next Col

        'Change the DataGrid's parent property
152     Set GridEidon.Parent.Font = saveFont
154     GridEidon.Parent.ScaleMode = saveScaleMode
        'If finished, then move pointer to first record again
156     EID.Recordset.MoveFirst

        '<EhFooter>
        Exit Sub

AdjustDataGridColumns_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.AdjustDataGridColumns " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.AdjustDataGridColumns " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub  'End of AdjustDataGridColumns

Sub find_eid_parastat()

        '<EhHeader>
        On Error GoTo find_eid_parastat_Err

        '</EhHeader>

        Dim R As New ADODB.Recordset

100     R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
102     R.MoveFirst

104     Do While Not R.EOF

106         If Right(PARAS.Text, 1) = R("EIDOS") Then
108             F_TITLOS = R("TITLOS")

110             f_AkyrParas = Left(LTrim(R("akyr")), 1)

112             F_POS_APOU = R("POS_APOU")
114             F_AJIA_APOU = R("AJIA_APOU")
116             f_pel = R("pel")
118             f_pol = R("pol")    ' 1=POLISEIS 2=AGORES
120             F_STADIO = R("STADIO")    ' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
122             F_TYPOS_PAR = CNull(R("C1")) ' ΤΥΠΟΣ ΠΑΡΑΣΤΑΤΙΚΟΥ ΓΙΑ Ε_ΛΙΝΕ ΦΟΡΟΛΟΓΙΚΟΥ
124             If IsNull(R("GEF_P")) Then
126                 F_rec_p = 0
                Else
128                 F_rec_p = R("GEF_P")    ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
                End If
            
130             If IsNull(R("FORM_EKTYP")) Then
132                 f_FORM_EKTY = 0
                Else
134                 f_FORM_EKTY = Val(R("FORM_EKTYP"))    ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
                End If
            
136             If IsNull(R("XONDR")) Then
138                 F_XONDR = 1
                Else
                     If R("XONDR") = 0 Then   'καταργω την λιανικη
                         F_XONDR = 1
                    Else
                         F_XONDR = R("XONDR")
                     End If
                
140                 '    F_XONDR = R("XONDR")    ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
142                 '    If F_XONDR = 2 Then F_XONDR = 1
                End If
                
                'αν δεν δουλευω λιανικες τιμες να τα κανει ολα χονδρικης
144             If F_LIANIKES_TIMES = 0 Then
                    ' F_XONDR = 1
                End If
                
                'If F_XONDR = 0 Then
                '    Grid1.BackColor = &H80FF80 ' prasino
                'Else
                '    Grid1.BackColor = &HFF8080     ' mple
                'End If
                
146             F_YPOX_AFM = IIf(nNull(R!GEF_M) = 99, 1, 0)
148             F_FORMA1 = Trim(CNull(R("FORMA1")))
150             F_FORMA2 = Trim(CNull(R("FORMA2")))
152             F_FORMA3 = Trim(CNull(R("FORMA3")))
                                
154             f_myEID = CNull(R!MYEID)  ' 1.1 τιμολογιο MYDATA
156             f_ISMYDATA = nNull(R!ISMYDATA)
                
158             If IsNull(R("PARAKRATISI")) Then
160                 F_PARAKRATISI = 0
                    ' parakrat.Visible = False
162                 parakratisiT.Visible = False
164                 tel_pliroteo.Visible = False
166                 lblΠαρακράτηση.Visible = False
168                 lblΠληρωτέο.Visible = False
                    ' LBLTEL.Visible = False
170                 ComboPARAKRATISI.Visible = False
172                 Combo2PARAKRATISI.Visible = False
                    
                Else
                    ' parakrat.Visible = True
174                 parakratisiT.Visible = True
176                 tel_pliroteo.Visible = True
178                 lblΠαρακράτηση.Visible = True
180                 ComboPARAKRATISI.Visible = True
182                 Combo2PARAKRATISI.Visible = True
                    '  Combo2PARAKRATISI.Text = Combo2PARAKRATISI.List(2)
                    ' ComboPARAKRATISI.Text = ComboPARAKRATISI.List(5)
                   
184                 lblΠληρωτέο.Visible = True
                    'LBLTEL.Visible = True
186                 F_PARAKRATISI = R("PARAKRATISI")    ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
                End If
                
            End If

188         R.MoveNext
        Loop

        '<EhFooter>
        Exit Sub

find_eid_parastat_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.find_eid_parastat " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.find_eid_parastat " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function Get_meKomma(ByVal X As String) As String

        ' δινεις το αβγ  και το κανει      α','β','γ
        '<EhHeader>
        On Error GoTo Get_meKomma_Err

        '</EhHeader>

        Dim pol As String

100     pol = ""

        Dim k As Integer

102     For k = 1 To Len(X)
104         pol = pol + "'" + mID$(X, k, 1) + "',"
        Next

106     pol = Left(pol, Len(pol) - 2)
108     pol = mID$(pol, 2, Len(pol) - 1)

110     Get_meKomma = pol

        '<EhFooter>
        Exit Function

Get_meKomma_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Get_meKomma " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.Get_meKomma " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Private Function FIND_PARAMETROI()
'Dim K
'  For K = 1 To 200
'     If FORMA = F_AR_parametroi(K, 0) And Var = F_AR_parametroi(K, 1) Then
'          FIND_PARAMETROI = F_AR_parametroi(K, 2)
'          Exit For
'     End If
'  Next
'
'
'
'End Function

Function get_kerdos()

        '---------------------------------------------------------------------------------------------------------
        '<EhHeader>
        On Error GoTo get_kerdos_Err

        '</EhHeader>
        Dim kerdos As Single

        Dim R      As New ADODB.Recordset

        Dim k, c, p, T, e, F, TIMH

        On Error Resume Next

100     kerdos = 0

        'ypologizei to kerdos toy timologioy
102     For k = 1 To Grid1.rows - 1

104         If Len(Grid1.TextMatrix(k, f_k)) > 0 Then
106             Grid1.row = k
108             c = Grid1.Col  ' απομνημονεύω την τρέχουσα στήλη kai σειρά
110             Grid1.Col = f_p: p = Grid1.Text
112             Grid1.Col = F_T: T = Grid1.Text
114             Grid1.Col = f_e: e = Grid1.Text
                'timh polisis
116             TIMH = Round(Val(T) * (100 - Val(e)) / 100, f_psifiaAjias)
118             R.Open "select XTI FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic

120             If IsNull(R(0)) Then
122                 kerdos = kerdos + Val(p) * TIMH
                Else
124                 kerdos = kerdos + Val(p) * (TIMH - R(0))
                End If
            End If

        Next

126     get_kerdos = kerdos

        '<EhFooter>
        Exit Function

get_kerdos_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.get_kerdos " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.get_kerdos " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function
'Sub FileNames(ByRef fNAMES() As String)  ' GETFILENAMES
'  Dim Fname As String
'  'Dim fNAMES() As String
'  Dim FType As String
'  Dim I As Integer
'
'  FType = "C:\PDF\PRFTEMP\*.PDF"
'  Fname = Dir(FType)
'  Do Until Fname = ""
'    I = I + 1
'    ReDim Preserve fNAMES(1 To I)
'    fNAMES(I) = Fname
'    Fname = Dir
'  Loop
''  If i <> 0 Then
' '   For i = 1 To UBound(fNAMES)
'  '    MsgBox fNAMES(i)
'   ' Next i
' ' End If
'End Sub

'                If F_PDF = 1 Then
'
'                   Dim x33, X34
'                     Dim NN3
'                     NN3 = GetCurrentTime()
'                     Do While Len(Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF", vbNormal)) = 0
'                        DoEvents
'                        If GetCurrentTime() - NN3 > 50000 Then
'                           Exit Do
'                        End If
'                     Loop
'                     NN3 = GetCurrentTime() - NN3
'
'
'
'                     MILSEC 1000
'
'
'
'                     Dim k35 As Long
'                     Dim PDF1(10) As String, K36 As Long
'
'                   'MILSEC 10000
'
'
''                     X33 = ShellExecute(0&, "Print", "C:\print3.pdf", "", "", vbNormalFocus)
'                      MILSEC 2000
'                     NN3 = 0
'
'1035
'
'                    k35 = 0
'                     pdf2 = Dir("C:\PDF\PDFTEMP\*.PDF")
'                   Do While Len(pdf2) > 0
'                      k35 = k35 + 1
'                      PDF1(k35) = pdf2
'                     pdf2 = Dir '("C:\PDF\PDFTEMP\*.PDF")
'                   Loop
'
'                    NN3 = NN3 + 1
'                    If k35 < 2 And NN3 < 5 Then
'                       MILSEC 5000
'                       GoTo 1035
'                    End If
'                   For K36 = 1 To k35
'
'                    ' PDF1(K36) = Dir(F_PDFFOLDER + "\PDFTEMP\*.PDF", vbNormal)
'                     MILSEC 1000
'                     If Len(PDF1(K36)) < 1 Then
'                         Exit For
'                     End If
'
'                     x33 = ShellExecute(0&, "Print", F_PDFFOLDER + "\PDFTEMP\" + PDF1(K36), "", "", vbNormalFocus)
'                     FileCopy F_PDFFOLDER + "\PDFTEMP\" + PDF1(K36), F_PDFFOLDER + "\" + Format(f_dat, "YYYYMMDD") + f_matim + "-" + LTrim(Str(K36)) + ".pdf"
'                   Next
'                End If

Sub METASX_SE_TIM(ByVal mpar As Integer, _
                  ByVal cTel_poso As String, _
                  ByVal m_ID As Long, _
                  ByVal isAkyrotiko As Boolean, _
                  ByVal NORMALTIMES0XORIS1 As Integer)
        '===προσοχη γιατι εξαρταται απο το TDBGRid1 poio record σημαδεψα=====mkod,atim
'112     MKOD = TDBGrid1.Text
'116     ATIM = TDBGrid1.Text
'152     DTPicker1.Value = CDate(TDBGrid1.Text)

        
        '============================================================================================
        ' MONTE-CRISTO ΜΕΤΑΣΧΗΜΑΤΙΣΜΟΣ ΣΕ ΤΙΜΟΛΟΓΙΟ
        
        '99=ΔΙΟΡΘΩΣΗ
        
100     ar_paragg.Visible = True
102     ar_paraggt.Visible = True
        
        On Error GoTo METASX_Err

        Dim METASX_TIM As Integer

104     METASX_TIM = mpar ' Val(FindParametroi(1,"PAR1", "METASX_TIM", "3", "ΠΑΡΑΣΤΑΤΙΚO ΤΠ-ΔΑ(ΜΕΤΑΣΧ.ΣΕ ΤΙΜΟΛ.)0-99"))

        Dim FF As PAR2, MKOD As String

        Dim ATIM

106     F_METASX_PAR = 1  ' XREIAZETAI GIA NA PAIRNEI ARIUMO PARASTATIKOY ME THN ALLAGH TOY COMBO PARASTATIKON

108     MIDENIZO_GRID

110     TDBGrid1.Col = 4
112     MKOD = TDBGrid1.Text

114     TDBGrid1.Col = 1
116     ATIM = TDBGrid1.Text




        Dim k As Integer

        Dim R As New ADODB.Recordset

118     R.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + Left(ATIM, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic

120     If R.EOF Then
122         MsgBox "ΔΕΝ ΕΙΝΑΙ ΔΗΛΩΜΕΝΟ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ"
            Exit Sub
        End If
124     If R("pol") > 3 Or R("pol") < 1 Then
126         MsgBox "ΔΕΝ ΕΙΝΑΙ ΔΗΛΩΜΕΝΟ ΣΤΗ ΛΙΣΤΑ ΠΑΡΑΣΤΑΤΙΚΩΝ "
            Exit Sub    ' δεν διορθωνω  asxeta
        End If

128     If R("pol") = 1 Then    'pvlhseis
130         POLAGOR.ListIndex = 1
132     ElseIf R("pol") = 2 Then
134         POLAGOR.ListIndex = 0    ' agores
136     ElseIf R("pol") = 3 Then
138         POLAGOR.ListIndex = 2    ' κινησεις αποθηκης
140         fbuff = "r"
        End If

        '  PARAS.Text = R("titlos") + Space(1) + Left(Atim, 1)
142     POLAGOR_LostFocus

        'POLAGOR.SetFocus
        'PARAS.SetFocus
        'Dim k As Integer
144     For k = 0 To PARAS.ListCount - 1
146         PARAS.ListIndex = k

148         If Right(PARAS.Text, 1) = Left(ATIM, 1) Then

                Exit For

            End If

        Next

150     TDBGrid1.Col = 0
152     DTPicker1.Value = CDate(TDBGrid1.Text)

154     If k = PARAS.ListCount Then Exit Sub
        ' βρισκω τις παραμετρους του παραστατικού
156     find_eid_parastat

158     Text1.Text = Val(mID(ATIM, 2, 6))
    
160     If Len(Trim(mID(ATIM, 8, 1))) > 0 Then
            '  SEIRA.Text = mID(Atim, 8, 1)
        Else
            '  SEIRA.Text = ""
        End If
    
162     Text2(0).Text = MKOD

164     Text2(0).Enabled = True
        'Text2(0).SetFocus
     
        Dim RR As New ADODB.Recordset
166     RR.Open "SELECT TOP 1 EIDOS,KPE,TRP,B_N1,B_N2,HME,EIDPAR,PARAT FROM TIM WHERE ID_NUM=" + str(m_ID), Gdb, adOpenDynamic, adLockOptimistic
     
168     PEL.RecordSource = "select  *  from PEL where EIDOS='" + RR!EIDOS + "' and KOD='" + RR!KPE + "'" '  IN ( 'like '" + SameLetters(Text2(0).Text) + "%' order by KOD;"
170     PEL.Refresh
     
172     APOTH1.Text = APOTH1.List(RR!B_N1 - 1)
        '   Combo8.Text = RR!EIDPAR
     
174     If RR!B_N2 > 0 Then
176         APOTH2.Text = APOTH1.List(RR!B_N2 - 1)
        End If
     
178     Combo8.ListIndex = Val(Left(RR!EIDPAR, 2))
     
180     DTPicker1.Value = RR!hme
     
182     F_ID_NUM_DIOR = m_ID
     
        'GridPelaton.Columns(0).Width = 6000
        'GridPelaton.Columns(0).DataField = "XX"                  'PAR1.METASX_SE_TIM m, C, m_ID, False
    
        '  GridPelaton.Columns(0).Width = 6000
        ' GridPelaton.Visible = True
        
184     If NORMALTIMES0XORIS1 = 1 Then
        
186         mmGridPelaton_DblClick 2, m_ID, 1, 0
        Else
        
188         mmGridPelaton_DblClick 2, m_ID, 0, mpar
        End If

190     parat.Text = CNull(RR!parat)

        ' PEL.RecordSource = "select  *  from PEL where EIDOS='e' and KOD='" + Text2(0).Text + "'"
        'PEL.Refresh

192     timtex3_enter

194     If mpar = 99 Then
           
            'Combo4.ListIndex = Val(Split(RR!TRP, ";")(1)) - 1
196         F_METASX_PAR = 0  ' XREIAZETAI GIA NA ΜΗΝ PAIRNEI ARIUMO PARASTATIKOY ME THN ALLAGH TOY COMBO PARASTATIKON
            
            Dim Nmetr As Integer
198         Nmetr = Val(Split(RR!trp, ";")(0)) - 1
200         If Nmetr < 0 Then Nmetr = 0
202         Combo4.Text = Combo4.List(Nmetr)
            ' Combo4.ListIndex = Val(Split(RR!TRP, ";")(1)) - 1
        
        Else

204         POLAGOR.ListIndex = 1 ' ΠΩΛΗΣΕΙΣ
206         PARAS.ListIndex = mpar ' Val(Left(MONTE_CRISTO.Caption, 2))
208         PARAS.Enabled = True
210         PARAS.SetFocus

212         DoEvents
   
            'ΚΑΝΟΝΙΖΕΙ ΤΟ ΠΟΣΟ
214         If F_CHANGE_TELIKO_POSO = 1 Then
216             Text5.Text = "*" + cTel_poso ' Right(MONTE_CRISTO.Caption, 6)
218             Text5_LostFocus
            End If
220         Combo4.ListIndex = 0

        End If

222     b_Katax.SetFocus
        
224     If isAkyrotiko Then
226         SaveParastat True
        End If

        '  SendKeys "{ENTER}{ENTER} {ENTER} "

        '<EhFooter>
        Exit Sub

METASX_Err:
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.METASX_SE_TIM " & "at line " & Erl

        Resume Next

End Sub

'Sub METASX_DIORTOSIS(ByVal mPar As Integer, ByVal cTel_poso As String)
''============================================================================================
'' MONTE-CRISTO ΜΕΤΑΣΧΗΜΑΤΙΣΜΟΣ ΣΕ ΤΙΜΟΛΟΓΙΟ
'
'
'        ar_paragg.Visible = True
'        ar_paraggt.Visible = True
'
'
'        On Error GoTo METASX_Err
'
'
'
'Dim METASX_TIM As Integer
'
'20       METASX_TIM = mPar ' Val(FindParametroi(1,"PAR1", "METASX_TIM", "3", "ΠΑΡΑΣΤΑΤΙΚO ΤΠ-ΔΑ(ΜΕΤΑΣΧ.ΣΕ ΤΙΜΟΛ.)0-99"))
'
'
'
'
'        Dim FF As PAR2, mkod As String
'
'        Dim Atim
'
'100     MIDENIZO_GRID
'
'110     TDBGrid1.Col = 4
'120     mkod = TDBGrid1.Text
'
'
'
'130     TDBGrid1.Col = 1
'140     Atim = TDBGrid1.Text
'
'
'
'        Dim k As Integer
'
'
'
'        Dim r As New ADODB.Recordset
'
'150     r.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + Left(Atim, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
'160     If r.EOF Then Exit Sub
'170     If r("pol") > 3 Or r("pol") < 1 Then Exit Sub    ' δεν διορθωνω  asxeta
'
'
'180     If r("pol") = 1 Then    'pvlhseis
'190         POLAGOR.ListIndex = 1
'200     ElseIf r("pol") = 2 Then
'210         POLAGOR.ListIndex = 0    ' agores
'220     ElseIf r("pol") = 3 Then
'230         POLAGOR.ListIndex = 2    ' κινησεις αποθηκης
'        End If
'
'
'
'        '  PARAS.Text = R("titlos") + Space(1) + Left(Atim, 1)
'240     POLAGOR_LostFocus
'        'POLAGOR.SetFocus
'        'PARAS.SetFocus
'        'Dim k As Integer
'250     For k = 0 To PARAS.ListCount - 1
'260         PARAS.ListIndex = k
'270         If Right(PARAS.Text, 1) = Left(Atim, 1) Then
'                Exit For
'            End If
'        Next
'
'
'280     TDBGrid1.Col = 0
'290     DTPicker1.Value = CDate(TDBGrid1.Text)
'
'300     If k = PARAS.ListCount Then Exit Sub
'        ' βρισκω τις παραμετρους του παραστατικού
'310     find_eid_parastat
'
'320     Text1.Text = Val(mID(Atim, 2, 6))
'
'330     If Len(Trim(mID(Atim, 8, 1))) > 0 Then
'           '  SEIRA.Text = mID(Atim, 8, 1)
'        Else
'          '  SEIRA.Text = ""
'        End If
'
'340     Text2(0).Text = mkod
'
'350     Text2(0).Enabled = True
'360     'Text2(0).SetFocus
'
'
'     PEL.RecordSource = "select  *  from PEL where EIDOS='e' and KOD like '" + SameLetters(Text2(0).Text) + "%' order by KOD;"
'     PEL.Refresh
'
'     'GridPelaton.Columns(0).Width = 6000
'     'GridPelaton.Columns(0).DataField = "XX"
'
'   '  GridPelaton.Columns(0).Width = 6000
'    ' GridPelaton.Visible = True
'
'     mmGridPelaton_DblClick 0
'
'
'   ' PEL.RecordSource = "select  *  from PEL where EIDOS='e' and KOD='" + Text2(0).Text + "'"
'    'PEL.Refresh
'
'
'
'timtex3_enter
'
'
'If mPar = 99 Then
'Else
'   POLAGOR.ListIndex = 1 ' ΠΩΛΗΣΕΙΣ
'   PARAS.ListIndex = mPar ' Val(Left(MONTE_CRISTO.Caption, 2))
'   PARAS.Enabled = True
'   PARAS.SetFocus
'   DoEvents
'   'ΚΑΝΟΝΙΖΕΙ ΤΟ ΠΟΣΟ
'   Text5.Text = "*" + cTel_poso ' Right(MONTE_CRISTO.Caption, 6)
'   Text5_LostFocus
'   Combo4.ListIndex = 0
'End If
'
'
'b_Katax.SetFocus
'
'
'
'
'
'
'370   '  SendKeys "{ENTER}{ENTER} {ENTER} "
'
'
'        '<EhFooter>
'        Exit Sub
'
'METASX_Err:
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.METASX_SE_TIM " & "at line " & Erl
'        Resume Next
'End Sub

'Sub REMOTE_IMPORT()
''====================================================================================================
'Dim k
'
'
'            '
'            On Error Resume Next
'         Dim R As New ADODB.Recordset
'         Dim rSQL As New ADODB.Recordset
'         Dim rSQL2 As New ADODB.Recordset
'
'
'         R.Open "SELECT * FROM TIMAGOR WHERE POS<>0", Gdb, adOpenDynamic, adLockOptimistic
'
'         If R.RecordCount = 0 Then
'             MsgBox "δεν υπάρχουν εγγραφές"
'
'                Exit Sub
'
'            End If
'
'         R.MoveFirst
'         k = 0
'
'            Dim M_CODE
'
'         Do While Not R.EOF
'             k = k + 1
'
'             If k > f_SeiresTimologioy Then Exit Do
'             If IsNull(R("BARCODE")) Or R("pos") = 0 Then
'                 Grid1.TextMatrix(k, f_k) = " "
'                Else
'
'                 rSQL.Open "SELECT * FROM BARCODES WHERE left(ERG,13)='" + (Left(R("BARCODE"), 13)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'
'                 If rSQL.EOF And Left(R("BARCODE"), 1) = "0" Then
'                     rSQL.Close
'
'                     rSQL.Open "SELECT * FROM BARCODES WHERE ERG LIKE '" + Trim(mID(R("BARCODE"), 2, 12)) + "%'", Gdb, adOpenForwardOnly, adLockReadOnly
'                    End If
'
'                 If rSQL.EOF Then
'                     M_CODE = " "
'                     Grid1.TextMatrix(k, f_k) = M_CODE
'                        '       OK = False
'                        '       List2.AddItem rDBF("BARCODE") + ": ΔΕΝ ΥΠΑΡΧΕΙ ΣΤΟ ΑΡΧΕΙΟ ΕΙΔΩΝ"
'                    Else
'                     M_CODE = rSQL("KOD")
'                     Grid1.TextMatrix(k, f_k) = M_CODE
'                        '   Grid1.TextMatrix(K, F_T) = rSQL("LTI")
'                     Grid1.TextMatrix(k, f_p) = R("POS")
'                    rSQL.Close
'                     rSQL.Open "SELECT * FROM EID WHERE KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'                     Grid1.TextMatrix(k, F_T) = Round(rSQL("LTI"), 2)
'                     Grid1.TextMatrix(k, f_o) = rSQL("ONO")
'                     Grid1.TextMatrix(k, f_f) = rSQL("FPA")
'
'                     If f_TIMOKAT = 9999 Then 'περπατουλης τιμοκατάλογος ειδος + πελατης
'
'                            Dim timokPEL As String, timokEID As String
'
'                            timokPEL = IIf(IsNull(PEL.Recordset("CH4")), "", PEL.Recordset("CH4"))
'                            timokEID = IIf(IsNull(rSQL("CH1")), "", rSQL("CH1"))
'
'                         rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOKPEL='" + timokPEL + " AND TIMOKEID='" + timokEID + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'
'                         If rSQL2.EOF Then
'                             Grid1.TextMatrix(k, f_e) = 0
'                            Else
'                             Grid1.TextMatrix(k, f_e) = rSQL2(0)
'                            End If
'
'                         rSQL2.Close
'                            ' timokatalogos εκπτωσης
'                     ElseIf f_TIMOKAT >= 1 And f_TIMOKAT < 300 And Not IsNull(PEL.Recordset("pek")) Then
'                         rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'                         If rSQL2.EOF Then
'                             Grid1.TextMatrix(k, f_e) = 0
'                            Else
'                             Grid1.TextMatrix(k, f_e) = rSQL2(0)
'                            End If
'
'                         rSQL2.Close
'
'                            'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
'                     ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then
'
'                            rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'
'                         If rSQL2.EOF Then
'                                ' Grid1.TextMatrix(k, F_T) = 0
'                            Else
'                             Grid1.TextMatrix(k, F_T) = rSQL2(0)
'                            End If
'
'                         rSQL2.Close
'
'                        ElseIf f_TIMOKAT = 333 And Not IsNull(PEL.Recordset("pek")) Then
'                            If PEL.Recordset("pek") = 2 Then
'                               Grid1.TextMatrix(k, F_T) = EID.Recordset("lti2")
'                            ElseIf PEL.Recordset("pek") = 3 Then
'                               Grid1.TextMatrix(k, F_T) = EID.Recordset("lti3")
'                            Else
'                               Grid1.TextMatrix(k, F_T) = EID.Recordset("lti")
'                            End If
'
'                        Else         ' με εκπτωση πελατη
'                         Grid1.TextMatrix(k, f_e) = PEL.Recordset("pek")
'                            ' Grid1.Text =
'                        End If
'
'                    Grid1.row = k
'                     FindSynolo
'
'                    End If
'
'                 rSQL.Close
'
'                End If
'
'             'R.Edit
'             R("pos") = 0
'             R.Update
'
'             R.MoveNext
'            Loop
'
'         R.Close
'        ' DB.Close
'
'            Exit Sub
'
'End Sub

Function FIND_LISTINDEX(MCOMBO As ComboBox, MTEXT As String) As Integer
        Dim N As Integer
    
100     For N = 0 To MCOMBO.ListCount
102         If MTEXT = MCOMBO.List(N) Then
                Exit For
            End If
        Next
    
104     FIND_LISTINDEX = N

End Function

Sub neo_akyr(ByVal id_num As Long, ATIM As String, ByRef ISOK As Integer)

        Dim atimakyr  As String

        Dim LOGOSAKYR As String
100     LOGOSAKYR = InputBox("Λόγος ακύρωσης", "", "λόγω λάθους ")

102     ISOK = 1

        On Error GoTo ROLL
104     Gdb.BeginTrans

        Dim ARITM As Single, akyr As String
        'ΒΡΙΣΚΩ ΜΕ ΠΟΙΟ ΑΚΥΡΩΤΙΚΟ ΑΚΥΡΩΝΕΤΑΙ Π.Χ. "κ"
106     akyr = GET_CVALUE(" SELECT AKYR FROM PARASTAT WHERE EIDOS='" + Left$(ATIM, 1) + "'")

        'ΒΡΙΣΚΩ ΑΠΟ ΠΟΙΟΝ ΑΡΘΜΗΤΗΡΑ ΤΡΑΒΑΕΙ το ακυρωτικο  π.χ. 2
108     ARITM = GET_NVALUE(" SELECT ARITMISI FROM PARASTAT WHERE EIDOS='" + akyr + "'")
   
110     F_TYPOS_PAR = GGET_CVALUE("select ISNULL(C1,'173')  from PARASTAT WHERE EIDOS='" + Left(akyr, 1) + "'")
   
        'ΑΝΕΒΑΖΩ ΤΟΝ ΑΡΘΜΗΤΗΡΑ του ακυρωτικου
112     Gdb.Execute "UPDATE ARITMISI SET ARITMISI=ARITMISI+1 WHERE ID=" + str(ARITM)

        'ΒΡΙΣΚΩ τον αριθμο που θα παρει  το ακυρωτικο  π.χ.  186
        Dim aritm2 As Single
114     aritm2 = GET_NVALUE(" SELECT ARITMISI FROM ARITMISI WHERE ID=" + str(ARITM))

        ' ariumhsh akyrotikoy
116     atimakyr = akyr + Format(aritm2, Left$("0000000000", f_psifia_atim)) '  "κ900001" ' π.χ.

118     Gdb.Execute "select * into DOKTIMAKYR" + Format(id_num, "00000000") + " FROM TIM WHERE ID_NUM=" + str(id_num)
120     Gdb.Execute "select * into DOKEGGTIMAKYR" + Format(id_num, "00000000") + " FROM EGGTIM WHERE ID_NUM=" + str(id_num)
122     Gdb.Execute "select * into DOKEGGAKYR" + Format(id_num, "00000000") + " FROM EGG WHERE IDTIM=" + str(id_num)

124     Gdb.Execute "ALTER TABLE  DOKTIMAKYR" + Format(id_num, "00000000") + "    DROP COLUMN ID_NUM"
126     Gdb.Execute "ALTER TABLE  DOKEGGTIMAKYR" + Format(id_num, "00000000") + " DROP COLUMN ID"
128     Gdb.Execute "ALTER TABLE  DOKEGGAKYR" + Format(id_num, "00000000") + "    DROP COLUMN ID"
   
        Dim cHME As String
130     cHME = Format(Now, "MM/dd/yyyy")

132     Gdb.Execute "UPDATE  DOKTIMAKYR" + Format(id_num, "00000000") + " SET ENTITYMARK='',SXETMARK=ENTITYMARK,SXETIKO='" + str(id_num) + "', HME='" + cHME + "', PARAT='ΣΧ." + ATIM + "', KLEIDI='" + atimakyr + "',ATIM='" + atimakyr + "',AJ1=-AJ1,AJ2=-AJ2,AJ3=-AJ3,AJ4=-AJ4,AJ5=-AJ5,AJ6=-AJ6,AJ7=-AJ7,AJ8=-AJ8,AJ9=-AJ9,FPA1=-FPA1,FPA2=-FPA2,FPA3=-FPA3,FPA4=-FPA4,FPA6=-FPA6,FPA7=-FPA7,FPA8=-FPA8,FPA9=-FPA9,AJI=-AJI"

        Dim nn As Long

        ' ΚΑΝΩ INSERT TO TIM GIA NA PAREI ID_NUM GIA NA TO BALO STO EGGTIM,EGG

134     Gdb.Execute "INSERT INTO TIM select * FROM DOKTIMAKYR" + Format(id_num, "00000000"), nn

        If nn > 0 Then
            'ok
        Else
            Gdb.RollbackTrans
            ISOK = 0
            MsgBox "δεν εγινε η ακυρωση"
            Exit Sub
        End If
        
            



        Dim RRR As New ADODB.Recordset, C_ID_NUM As String
  
136     RRR.Open "SELECT @@IDENTITY", Gdb, adOpenDynamic, adLockOptimistic   'MAX(ID_NUM) FROM TIM"
138     C_ID_NUM = str(RRR(0))
140     RRR.Close
142     If nn > 0 Then
144         Gdb.Execute "update TIM SET PARAT=PARAT+'" + LOGOSAKYR + "' WHERE ID_NUM=" + C_ID_NUM
        
        End If
        
146     Gdb.Execute "UPDATE  DOKEGGTIMAKYR" + Format(id_num, "00000000") + " SET HME='" + cHME + "',KAU_AJIA=-KAU_AJIA,MIK_AJIA=-MIK_AJIA,KERDOS=-KERDOS,POSO=-POSO,XRE=-XRE,PIS=-PIS,ATIM='" + atimakyr + "', ID_NUM=" + C_ID_NUM
148     Gdb.Execute "UPDATE  DOKEGGAKYR" + Format(id_num, "00000000") + " SET  HME='" + cHME + "', XREOSI=-XREOSI,PISTOSI=-PISTOSI,XRE=-XRE,ATIM='" + atimakyr + "',IDTIM=" + C_ID_NUM
  
150     Gdb.Execute "INSERT INTO EGGTIM select * FROM DOKEGGTIMAKYR" + Format(id_num, "00000000")
152     Gdb.Execute "INSERT INTO EGG    select * FROM DOKEGGAKYR" + Format(id_num, "00000000")
   
154     Gdb.Execute "drop TABLE DOKTIMAKYR" + Format(id_num, "00000000")
156     Gdb.Execute "drop TABLE DOKEGGTIMAKYR" + Format(id_num, "00000000")
158     Gdb.Execute "drop TABLE DOKEGGAKYR" + Format(id_num, "00000000")
                
160     Gdb.Execute "UPDATE TIM SET AKYROMENO=1  where ID_NUM=" + str(id_num) '

162     Gdb.CommitTrans
        '
        Dim DUM As String

        '!!!!!!!!!!!!!!! ola ayta παρακατο ta pairnei apo thn οθονη
        'Η ΕΚΤΥΠΩΣΗ ΟΠΟΤΕ ΤΑ  ΓΕΜΙΖΩ
        On Error Resume Next

        Dim r3 As New ADODB.Recordset
164     r3.Open "SELECT * FROM TIM WHERE ID_NUM=" + C_ID_NUM, Gdb, adOpenDynamic, adLockOptimistic

        'Dim f_paras As String

166     f_paras = GET_CVALUE(" SELECT TITLOS FROM PARASTAT WHERE EIDOS='" + akyr + "'") + " "

168     SKOPOS.Text = CNull(r3!SKOPOS)
170     PROORISMOS.Text = CNull(r3!PROOR)
172     fortosh.Text = CNull(r3!fortosh)
174     AYTOKINHTO.Text = CNull(r3!AYTOK)
176     parat.Text = CNull(r3!parat)

178     Combo4.Text = Combo4.List(0)
        Dim k As Integer
   
180     DTPicker1.Value = Now
182     Text1.Text = aritm2

        ' For k = 0 To Combo4.ListCount - 1
        '        If InStr(Left(Combo4.List(k), 3), Split(r3("TRP"), ";")(0)) > 0 Then
        '           Combo4.Text = Combo4.List(k)
        '           Exit For
        '        End If
        '    Next
        '

184     k = Val(Split(r3("TRP"), ";")(0)) - 1
186     If k >= 0 Then
188         Combo4.Text = Combo4.List(k)
        End If

190     r3.Close

        '--------------------------ΕΚΤΥΠΟΣΗ -----------------------------------------
        'απο εδω πρεπει να τον δωσω τιμη για το F_TYPOS_PAR

192     f_Akyr = True
194     DUM = printCrystal(atimakyr, CDate(Now))    'δημιουργει τα προσωρινά αρχεία
196     f_Akyr = False
198     If DUM < -1 Then
200         MsgBox "ΠΡΟΒΛΗΜΑ ΣΤΗΝ ΕΚΤΥΠΩΣΗ  ΑΚΥΡΩΤΙΚΟΥ"
202         Gdb.Execute "UPDATE TIM SET AKYROMENO=0  where ID_NUM=" + str(id_num) '
204         ISOK = 0
            Exit Sub
        End If

        Dim m_FORM_EKTY As Integer
        Dim m_FORMA1    As String
        Dim m_FORMA2    As String
        Dim rakyr       As New ADODB.Recordset

206     rakyr.Open "SELECT isnull(FORMA1,'') AS FORMA1,isnull(FORMA2,'') AS FORMA2,ISNULL(FORM_EKTYP,0) AS FORM_EKTYP  FROM PARASTAT WHERE EIDOS='" + akyr + "'", Gdb, adOpenDynamic, adLockOptimistic
208     m_FORMA1 = Trim(rakyr!FORMA1)
210     m_FORMA2 = Trim(rakyr!FORMA2)
212     m_FORM_EKTY = rakyr!FORM_ektyp
214     rakyr.Close

216     If m_FORM_EKTY >= 90 And m_FORM_EKTY <= 99 Then     ' φορμα τύπου f90.txt dot matrix '====================================
218         DUM = print_timol(f_matim, CDate(DTPicker1.Value))
220     ElseIf m_FORM_EKTY = 10 Or m_FORM_EKTY = 11 Then
222         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
224         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
            'CrystalReport1.Action = 1
226     ElseIf m_FORM_EKTY = 14 Then
228         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"
            'CrystalReport1.Action = 1
230         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
232     ElseIf m_FORM_EKTY = 13 Then
234         CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol3.rpt"
236         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
            'CrystalReport1.Action = 1
238     ElseIf Len(Trim(m_FORMA1)) > 1 Then
240         CrystalReport1.ReportFileName = m_FORMA1
            'CrystalReport1.Action = 1
242         Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
244         If Len(Trim(m_FORMA2)) > 1 Then
246             CrystalReport1.ReportFileName = m_FORMA2
                'CrystalReport1.Action = 1
248             Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
            End If
        Else
            ' CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
            ' CrystalReport1.Action = 1
 
        End If

250     SKOPOS.Text = SKOPOS.List(0) ' ""   CNull(r3!SKOPOS)
252     PROORISMOS.Text = PROORISMOS.List(0) ' CNull(r3!PROOR)
254     fortosh.Text = fortosh.List(0) ' CNull(r3!FORTOSH)
256     AYTOKINHTO.Text = AYTOKINHTO.List(0) 'CNull(r3!AYTOK)
258     parat.Text = "" ' CNull(r3!parat)

        Exit Sub

ROLL:
        'Resume Next
260     MsgBox Err.Description + Chr(13) + " ΔΕΝ ΔΗΜΙΟΥΡΓΗΘΗΚΕ"
262     Gdb.RollbackTrans
264     ISOK = 0

End Sub

Sub akyr_benzinadiko(ByVal id_num As Long, ATIM As String)
        'forosDiamonis
        Dim atimakyr As String
Dim ubl As String
        On Error GoTo ROLL
      If F_PAROCHOS = 1 Then
      Else
100     Gdb.BeginTrans
      End If
      
      
      
        Dim ARITM As Single, akyr As String
        'ΒΡΙΣΚΩ ΜΕ ΠΟΙΟ ΑΚΥΡΩΤΙΚΟ ΑΚΥΡΩΝΕΤΑΙ Π.Χ. "κ"
        ' akyr = GET_CVALUE(" SELECT AKYR FROM PARASTAT WHERE EIDOS='" + Left$(ATIM, 1) + "'")

        ' στο akyr θα βαλω το "επιστροφη λιανικης   "p"
102     akyr = "p"

        Dim nc As Integer
104     nc = GET_NVALUE("select count(*) from PEL WHERE KOD='1111' AND EIDOS='e'")
106     If nc = 0 Then
108         Gdb.Execute "insert into PEL (KOD,EIDOS,EPO,AFM) VALUES ('1111','e','πελάτης λιανικής','000000000') "
        End If

        'ΒΡΙΣΚΩ ΑΠΟ ΠΟΙΟΝ ΑΡΘΜΗΤΗΡΑ ΤΡΑΒΑΕΙ το ακυρωτικο  π.χ. 2
110     ARITM = GET_NVALUE(" SELECT ARITMISI FROM PARASTAT WHERE EIDOS='" + akyr + "'")
   
        'ΑΝΕΒΑΖΩ ΤΟΝ ΑΡΘΜΗΤΗΡΑ του ακυρωτικου
112     Gdb.Execute "UPDATE ARITMISI SET ARITMISI=ARITMISI+1 WHERE ID=" + str(ARITM)

        'ΒΡΙΣΚΩ τον αριθμο που θα παρει  το ακυρωτικο  π.χ.  186
        Dim aritm2 As Single
114     aritm2 = GET_NVALUE(" SELECT ARITMISI FROM ARITMISI WHERE ID=" + str(ARITM))

        ' ariumhsh akyrotikoy
116     atimakyr = akyr + Format(aritm2, Left$("0000000000", f_psifia_atim)) '  "κ900001" ' π.χ.

118     Gdb.Execute "select * into DOKTIMAKYR" + Format(id_num, "00000000") + " FROM TIM WHERE ID_NUM=" + str(id_num)
120     Gdb.Execute "select * into DOKEGGTIMAKYR" + Format(id_num, "00000000") + " FROM EGGTIM WHERE ID_NUM=" + str(id_num)
122     Gdb.Execute "select * into DOKEGGAKYR" + Format(id_num, "00000000") + " FROM EGG WHERE IDTIM=" + str(id_num)

124     Gdb.Execute "ALTER TABLE  DOKTIMAKYR" + Format(id_num, "00000000") + "    DROP COLUMN ID_NUM"
126     Gdb.Execute "ALTER TABLE  DOKEGGTIMAKYR" + Format(id_num, "00000000") + " DROP COLUMN ID"
128     Gdb.Execute "ALTER TABLE  DOKEGGAKYR" + Format(id_num, "00000000") + "    DROP COLUMN ID"
   
        Dim cHME As String
130     cHME = Format(Now, "MM/dd/yyyy")

132     Gdb.Execute "UPDATE  DOKTIMAKYR" + Format(id_num, "00000000") + " SET ENTITYMARK='',SXETIKO='" + ar_apod.Text + "', HME='" + cHME + "', PARAT='ΣΧ.ΑΠΟΔ." + ar_apod.Text + "', KLEIDI='" + atimakyr + "',ATIM='" + atimakyr + "'"
        ' ΔΕΝ ΤΑ ΚΑΝΩ ΑΡΝΗΤΙΚΑ ΓΙΑ ΝΑ ΠΑΝΕ MYDATA      ,AJ1=-AJ1,AJ2=-AJ2,AJ3=-AJ3,AJ4=-AJ4,AJ5=-AJ5,AJ6=-AJ6,AJ7=-AJ7,AJ8=-AJ8,AJ9=-AJ9,FPA1=-FPA1,FPA2=-FPA2,FPA3=-FPA3,FPA4=-FPA4,FPA6=-FPA6,FPA7=-FPA7,FPA8=-FPA8,FPA9=-FPA9,AJI=-AJI"
134     Gdb.Execute "UPDATE  DOKTIMAKYR" + Format(id_num, "00000000") + " SET KPE='1111',JWT='' "
        Dim L As Long

        ' ΚΑΝΩ INSERT TO TIM GIA NA PAREI ID_NUM GIA NA TO BALO STO EGGTIM,EGG
136     Gdb.Execute "INSERT INTO TIM select * FROM DOKTIMAKYR" + Format(id_num, "00000000"), L
  
        Dim RRR As New ADODB.Recordset, C_ID_NUM As String
  
138     RRR.Open "SELECT   MAX(ID_NUM) FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
140     C_ID_NUM = str(RRR(0))
142     RRR.Close
        
        ' Gdb.Execute "UPDATE  DOKEGGTIMAKYR" + Format(ID_NUM, "00000000") + " SET PELKOD='1111',HME='" + cHME + "',KAU_AJIA=-KAU_AJIA,MIK_AJIA=-MIK_AJIA,KERDOS=-KERDOS,POSO=-POSO,ATIM='" + atimakyr + "', ID_NUM=" + C_ID_NUM
144     Gdb.Execute "UPDATE  DOKEGGTIMAKYR" + Format(id_num, "00000000") + " SET PELKOD='1111',HME='" + cHME + "',ATIM='" + atimakyr + "', ID_NUM=" + C_ID_NUM
 
146     Gdb.Execute "UPDATE  DOKEGGAKYR" + Format(id_num, "00000000") + " SET  KOD='1111',HME='" + cHME + "', XREOSI=-XREOSI,PISTOSI=-PISTOSI,XRE=-XRE,ATIM='" + atimakyr + "',IDTIM=" + C_ID_NUM
  
148     Gdb.Execute "INSERT INTO EGGTIM select * FROM DOKEGGTIMAKYR" + Format(id_num, "00000000"), L
150     Gdb.Execute "INSERT INTO EGG    select * FROM DOKEGGAKYR" + Format(id_num, "00000000"), L
   
152     Gdb.Execute "drop TABLE DOKTIMAKYR" + Format(id_num, "00000000")
154     Gdb.Execute "drop TABLE DOKEGGTIMAKYR" + Format(id_num, "00000000")
156     Gdb.Execute "drop TABLE DOKEGGAKYR" + Format(id_num, "00000000")
  
        'ΣΤΕΛΝΩ ΤΗΝ ΑΠΟΔ.ΕΠΙΣΤΡ.ΛΙΑΝΙΚΗΣ ΣΤΟ μυδατα And isAkyrotiko = False
158     If f_MYDATAcont = 1 Then
            
160         Dim SendAkyrToMyData As Integer: SendAkyrToMyData = 0
                
162         SendAkyrToMyData = GGET_NVALUE("SELECT ISNULL(ISMYDATA,0) FROM PARASTAT WHERE EIDOS='" & akyr & "'")
164         If Len(Trim(f_myEID)) > 0 And f_ISMYDATA = 1 Then '  As String ' ΤΥΠΟΣ ΠΑΡΑΣΤΑΤΤΙΚΟΥ MYDATA  Π.X. 1.1 ΤΙΜΟΛΟΓΙΟ
                ' If POLAGOR.ListIndex = 1 Then  ' ΣΤΙς ΠΩΛΗΣΕΙΣ ΤΟ ΑΛΛΑΖΩ
                'Par7MyData.Visible = False
                   
166             If SendAkyrToMyData = 1 Then
                   
168                 Dim mStelno As Integer: mStelno = 1
170                 If mStelno = 1 Then
172                     gisOnlineMydata = 1
174                     Par7MyData.where.Caption = " where ID_NUM=" + C_ID_NUM + "  " ' M_idnum
176                     Par7MyData.Visible = False
                        If F_PAROCHOS = 1 Then
'                                 Dim fFILE110 As String: fFILE110 = Trim(atimakyr) + ".PDF"  '"c:\mercvb\" +
'                                Dim mForma110 As String
'                                mForma110 = GET_CVALUE(" SELECT ISNULL(FORMA1,'C:\MERCVB\REPORTS\TIMOL1.RPT')  FROM PARASTAT WHERE EIDOS='" + akyr + "'") + " "
'                                Dim Size110  As Long
'                                 Size110 = Export2ReportToPDF(mForma110, Trim(atimakyr), fFILE110)
                            Par7MyData.loadpar7
                             ubl = GGET_CVALUE("SELECT left(ISNULL(UBL,'000'),3) AS UBL2 FROM PARASTAT WHERE EIDOS='" & akyr & "'")

                            'e?? paroxo mono sthn xondriki kai to parastattiko anti na exei 380 exei "mydata"
                            If ubl = "000" Then 'mydata
                                Par7MyData.ToXMLsub 0, ""
                            Else
                                Par7MyData.ToJason 0, 0
                            End If
 
                          
                            'Par7MyData.ToJason 0
                        Else
178                         Par7MyData.ToXMLsub 0, ""
                        End If

                    ' Par7MyData.ToXMLsub 0, ""
                    End If
                End If
                   
            End If
        End If
       If F_PAROCHOS = 1 Then
      Else
     'Gdb.BeginTrans
         Gdb.CommitTrans
      End If
180
        '
        Dim DUM As String

        '!!!!!!!!!!!!!!! ola ayta παρακατο ta pairnei apo thn οθονη
        'Η ΕΚΤΥΠΩΣΗ ΟΠΟΤΕ ΤΑ  ΓΕΜΙΖΩ
        On Error Resume Next

        Dim r3 As New ADODB.Recordset
        'tsimpav to record toy akyrotikoy
182     r3.Open "SELECT * FROM TIM WHERE ID_NUM=" + C_ID_NUM, Gdb, adOpenDynamic, adLockOptimistic
        'o titlos toy akyrvtikoy
184     f_paras = GET_CVALUE(" SELECT TITLOS FROM PARASTAT WHERE EIDOS='" + akyr + "'") + " "

        Dim mForma1 As String
186     mForma1 = GET_CVALUE(" SELECT ISNULL(FORMA1,'C:\MERCVB\REPORTS\TIMOL1.RPT')  FROM PARASTAT WHERE EIDOS='" + akyr + "'") + " "

188     mForma1 = IIf(Len(Trim(mForma1)) = 0, "C:\MERCVB\REPORTS\TIMOL1.RPT", mForma1)

190     SKOPOS.Text = CNull(r3!SKOPOS)
192     PROORISMOS.Text = CNull(r3!PROOR)
194     fortosh.Text = CNull(r3!fortosh)
196     AYTOKINHTO.Text = CNull(r3!AYTOK)
        'parat.Text = CNull(r3!parat)

198     Combo4.Text = Combo4.List(0)
        Dim k As Integer
   
200     DTPicker1.Value = Now
             ' τον αριθμο που θα παρει  το ακυρωτικο  π.χ.  186
202     Text1.Text = aritm2

        ' For k = 0 To Combo4.ListCount - 1
        '        If InStr(Left(Combo4.List(k), 3), Split(r3("TRP"), ";")(0)) > 0 Then
        '           Combo4.Text = Combo4.List(k)
        '           Exit For
        '        End If
        '    Next
        '


'bazo ton tropo pliromis
204     k = Val(Split(r3("TRP"), ";")(0)) - 1
206     If k >= 0 Then
208         Combo4.Text = Combo4.List(k)
        End If

210     r3.Close

        '--------------------------ΕΚΤΥΠΟΣΗ -----------------------------------------

        '  f_Akyr = True
212     f_Akyr = False
214     F_TYPOS_PAR = 175 ' για ελινε
216     PAR_GIA_AKYROSH.Caption = ""
218     DUM = printCrystal(atimakyr, CDate(Now))    'δημιουργει τα προσωρινά αρχεία
220     f_Akyr = False
222     If DUM < -1 Then
224         MsgBox "ΠΡΟΒΛΗΜΑ ΣΤΗΝ ΕΚΤΥΠΩΣΗ ΤΗΣ ΑΚΥΡΩΤΙΚΗΣ ΑΠΟΔΕΙΞΗΣ"
            Exit Sub
        End If
        'anebazo to pdf sto azzure
        If F_PAROCHOS = 1 Then
                              If ubl <> "000" Then
                                        If GGET_NVALUE("select MYDATAISDEMO from MEM") = 1 Then
                                           PAROX_URLADD = "beta-"
                                        Else
                                           PAROX_URLADD = ""
                                        End If
                                        
                                Dim fFILE110 As String: fFILE110 = Trim(atimakyr) + ".PDF"  '"c:\mercvb\" +
                                Dim mForma110 As String
                                mForma110 = GET_CVALUE(" SELECT ISNULL(FORMA1,'C:\MERCVB\REPORTS\TIMOL1.RPT')  FROM PARASTAT WHERE EIDOS='" + akyr + "'") + " "
                                Dim Size110  As Long
                                Size110 = Export2ReportToPDF(mForma110, Trim(atimakyr), fFILE110)
                                uploadAzure Val(C_ID_NUM), fFILE110, Size110
                             
                             
                             End If
        End If
        
        
        
        
        
        
        ' If f_FORM_EKTY >= 90 And f_FORM_EKTY <= 99 Then     ' φορμα τύπου f90.txt dot matrix '====================================
        '          DUM = print_timol(f_matim, CDate(DTPicker1.Value))'
        'ElseIf Len(Trim(F_FORMA1)) > 1 Then
226     CrystalReport1.ReportFileName = mForma1
228     Ucr9print.printingCR9 F_2EKTYPOTHS, CrystalReport1.ReportFileName 'ebgala to f_ektypotis
        'CrystalReport1.Action = 1
            
        '        If Len(Trim(F_FORMA2)) > 1 Then
        '           CrystalReport1.ReportFileName = F_FORMA2
        '          Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
        ' CrystalReport1.Action = 1
        '     End If
        'Else
        '     CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
        '    Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
        'CrystalReport1.Action = 1
 
        'End If

        Exit Sub

ROLL:
        'Resume Next
        If F_PAROCHOS = 1 Then
           Resume Next
      Else
         Gdb.RollbackTrans
      End If
     

End Sub

Sub forosDiamonis(ByVal id_num As Long, ATIM As String)
        'forosDiamonis
        Dim atimakyr As String

        On Error GoTo ROLL

100     Gdb.BeginTrans

        Dim ARITM As Single, akyr As String
        'ΒΡΙΣΚΩ ΜΕ ΠΟΙΟ ΑΚΥΡΩΤΙΚΟ ΑΚΥΡΩΝΕΤΑΙ Π.Χ. "κ"
        'akyr = GET_CVALUE(" SELECT AKYR FROM PARASTAT WHERE EIDOS='" + Left$(ATIM, 1) + "'")

        akyr = LTrim(Trim(f_EID_AKYR_APOD))
        ' στο akyr θα βαλω το "επιστροφη λιανικης   "p"
102     '  akyr = "p"

        Dim nc As Integer
104     nc = GET_NVALUE("select count(*) from PEL WHERE KOD='1111' AND EIDOS='e'")
106     If nc = 0 Then
108         Gdb.Execute "insert into PEL (KOD,EIDOS,EPO,AFM) VALUES ('1111','e','πελάτης λιανικής','000000000') "
        End If

        'ΒΡΙΣΚΩ ΑΠΟ ΠΟΙΟΝ ΑΡΘΜΗΤΗΡΑ ΤΡΑΒΑΕΙ το ακυρωτικο  π.χ. 2
110     ARITM = GET_NVALUE(" SELECT ARITMISI FROM PARASTAT WHERE EIDOS='" + akyr + "'")
   
        'ΑΝΕΒΑΖΩ ΤΟΝ ΑΡΘΜΗΤΗΡΑ του ακυρωτικου
112     Gdb.Execute "UPDATE ARITMISI SET ARITMISI=ARITMISI+1 WHERE ID=" + str(ARITM)

        'ΒΡΙΣΚΩ τον αριθμο που θα παρει  το ακυρωτικο  π.χ.  186
        Dim aritm2 As Single
114     aritm2 = GET_NVALUE(" SELECT ARITMISI FROM ARITMISI WHERE ID=" + str(ARITM))

        ' ariumhsh akyrotikoy
116     atimakyr = akyr + Format(aritm2, Left$("0000000000", f_psifia_atim)) '  "κ900001" ' π.χ.

118     Gdb.Execute "select * into DOKTIMAKYR" + Format(id_num, "00000000") + " FROM TIM WHERE ID_NUM=" + str(id_num)
120     Gdb.Execute "select * into DOKEGGTIMAKYR" + Format(id_num, "00000000") + " FROM EGGTIM WHERE ID_NUM=" + str(id_num)
122     Gdb.Execute "select * into DOKEGGAKYR" + Format(id_num, "00000000") + " FROM EGG WHERE IDTIM=" + str(id_num)

124     Gdb.Execute "ALTER TABLE  DOKTIMAKYR" + Format(id_num, "00000000") + "    DROP COLUMN ID_NUM"
126     Gdb.Execute "ALTER TABLE  DOKEGGTIMAKYR" + Format(id_num, "00000000") + " DROP COLUMN ID"
128     Gdb.Execute "ALTER TABLE  DOKEGGAKYR" + Format(id_num, "00000000") + "    DROP COLUMN ID"
   
        'f_TimhForoyDiuamonis
        Dim POSOTHTA As Single
        POSOTHTA = GGET_NVALUE("SELECT SUM(POSO) FROM DOKEGGTIMAKYR" + Format(id_num, "00000000") + " WHERE KODE='1' ")
        POSOTHTA = POSOTHTA * f_TimhForoyDiuamonis
   
        Dim cHME As String
130     cHME = Format(Now, "MM/dd/yyyy")

132     Gdb.Execute "UPDATE  DOKTIMAKYR" + Format(id_num, "00000000") + " SET ENTITYMARK='',SXETIKO='" + ar_apod.Text + "', HME='" + cHME + "', PARAT='ΣΧ.ΑΠΟΔ." + ar_apod.Text + "', KLEIDI='" + atimakyr + "',ATIM='" + atimakyr + "'"
        ' ΔΕΝ ΤΑ ΚΑΝΩ ΑΡΝΗΤΙΚΑ ΓΙΑ ΝΑ ΠΑΝΕ MYDATA      ,AJ1=-AJ1,AJ2=-AJ2,AJ3=-AJ3,AJ4=-AJ4,AJ5=-AJ5,AJ6=-AJ6,AJ7=-AJ7,AJ8=-AJ8,AJ9=-AJ9,FPA1=-FPA1,FPA2=-FPA2,FPA3=-FPA3,FPA4=-FPA4,FPA6=-FPA6,FPA7=-FPA7,FPA8=-FPA8,FPA9=-FPA9,AJI=-AJI"
134     Gdb.Execute "UPDATE  DOKTIMAKYR" + Format(id_num, "00000000") + " SET APALAGIFPA=99,AJI=0,AJ1=0,AJ2=0,AJ3=0,AJ4=0,FPA1=0,FPA2=0,FPA3=0,FPA4=0,PARAKRATISI=" + Replace(Format(POSOTHTA, "##0.00"), ",", ".") + " ,B_C2='3;6'"
        Dim L As Long

        ' ΚΑΝΩ INSERT TO TIM GIA NA PAREI ID_NUM GIA NA TO BALO STO EGGTIM,EGG
136     Gdb.Execute "INSERT INTO TIM select * FROM DOKTIMAKYR" + Format(id_num, "00000000"), L
  
        Dim RRR As New ADODB.Recordset, C_ID_NUM As String
  
138     RRR.Open "SELECT   MAX(ID_NUM) FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
140     C_ID_NUM = str(RRR(0))
142     RRR.Close
        
        ' Gdb.Execute "UPDATE  DOKEGGTIMAKYR" + Format(ID_NUM, "00000000") + " SET PELKOD='1111',HME='" + cHME + "',KAU_AJIA=-KAU_AJIA,MIK_AJIA=-MIK_AJIA,KERDOS=-KERDOS,POSO=-POSO,ATIM='" + atimakyr + "', ID_NUM=" + C_ID_NUM
144     Gdb.Execute "UPDATE  DOKEGGTIMAKYR" + Format(id_num, "00000000") + " SET KAU_AJIA=0,MIK_AJIA=0,TIMM=0,PELKOD='1111',HME='" + cHME + "',ATIM='" + atimakyr + "', ID_NUM=" + C_ID_NUM
 
146     Gdb.Execute "UPDATE  DOKEGGAKYR" + Format(id_num, "00000000") + " SET  KOD='1111',HME='" + cHME + "', XREOSI=-XREOSI,PISTOSI=-PISTOSI,XRE=-XRE,ATIM='" + atimakyr + "',IDTIM=" + C_ID_NUM
  
148     Gdb.Execute "INSERT INTO EGGTIM select * FROM DOKEGGTIMAKYR" + Format(id_num, "00000000"), L
150

        Gdb.Execute "INSERT INTO EGG    select * FROM DOKEGGAKYR" + Format(id_num, "00000000"), L
   
152     Gdb.Execute "drop TABLE DOKTIMAKYR" + Format(id_num, "00000000")
154     Gdb.Execute "drop TABLE DOKEGGTIMAKYR" + Format(id_num, "00000000")
156     Gdb.Execute "drop TABLE DOKEGGAKYR" + Format(id_num, "00000000")
  
        'ΣΤΕΛΝΩ ΤΗΝ ΑΠΟΔ.ΕΠΙΣΤΡ.ΛΙΑΝΙΚΗΣ ΣΤΟ μυδατα And isAkyrotiko = False
158     If f_MYDATAcont = 1 Then
            
160         Dim SendAkyrToMyData As Integer: SendAkyrToMyData = 0
                
162         SendAkyrToMyData = GGET_NVALUE("SELECT ISNULL(ISMYDATA,0) FROM PARASTAT WHERE EIDOS='" & akyr & "'")
164         If Len(Trim(f_myEID)) > 0 And f_ISMYDATA = 1 Then '  As String ' ΤΥΠΟΣ ΠΑΡΑΣΤΑΤΤΙΚΟΥ MYDATA  Π.X. 1.1 ΤΙΜΟΛΟΓΙΟ
                ' If POLAGOR.ListIndex = 1 Then  ' ΣΤΙς ΠΩΛΗΣΕΙΣ ΤΟ ΑΛΛΑΖΩ
                'Par7MyData.Visible = False
                   
166             If SendAkyrToMyData = 1 Then
                   
168                 Dim mStelno As Integer: mStelno = 1
170                 If mStelno = 1 Then
172                     gisOnlineMydata = 1
174                     Par7MyData.where.Caption = " where ID_NUM=" + C_ID_NUM + "  " ' M_idnum
176                     Par7MyData.Visible = False
                        If F_PAROCHOS = 1 Then
                            Par7MyData.ToJason 0, 0
                        Else
178                         Par7MyData.ToXMLsub 0, ""
                        End If
                    End If
                End If
                   
            End If
        End If

180     Gdb.CommitTrans

664     If F_PAROCHOS = 1 And f_ISMYDATA = 1 Then
666         Dim fFILE As String: fFILE = Trim(atimakyr) + ".PDF"  '"c:\mercvb\" +
            Dim Size  As Long
668         Size = Export2ReportToPDF(F_FORMA1, Trim(atimakyr), fFILE)
670         Dim ubl As String: ubl = GGET_CVALUE("SELECT left(ISNULL(UBL,'000'),3) AS UBL2 FROM PARASTAT WHERE EIDOS='" & Left(atimakyr, 1) & "'")
            'If GGET_NVALUE("SELECT ISNULL(B2G,0) AS DIMOSIO FROM PARASTAT WHERE EIDOS='" & Left(f_matim, 1) & "'") = 1 Then
                        
672         If ubl <> "000" Then
674             If GGET_NVALUE("select MYDATAISDEMO from MEM") = 1 Then
676                 PAROX_URLADD = "beta-"
                Else
678                 PAROX_URLADD = ""
                End If
680             uploadAzure Val(C_ID_NUM), fFILE, Size
            End If
        End If

        '
        Dim DUM As String

        '!!!!!!!!!!!!!!! ola ayta παρακατο ta pairnei apo thn οθονη
        'Η ΕΚΤΥΠΩΣΗ ΟΠΟΤΕ ΤΑ  ΓΕΜΙΖΩ
        On Error Resume Next

        Dim r3 As New ADODB.Recordset
182     r3.Open "SELECT * FROM TIM WHERE ID_NUM=" + C_ID_NUM, Gdb, adOpenDynamic, adLockOptimistic

184     f_paras = GET_CVALUE(" SELECT TITLOS FROM PARASTAT WHERE EIDOS='" + akyr + "'") + " "

        Dim mForma1 As String
186     mForma1 = GET_CVALUE(" SELECT ISNULL(FORMA1,'C:\MERCVB\REPORTS\TIMOL1.RPT')  FROM PARASTAT WHERE EIDOS='" + akyr + "'") + " "

188     mForma1 = IIf(Len(Trim(mForma1)) = 0, "C:\MERCVB\REPORTS\TIMOL1.RPT", mForma1)

190     SKOPOS.Text = CNull(r3!SKOPOS)
192     PROORISMOS.Text = CNull(r3!PROOR)
194     fortosh.Text = CNull(r3!fortosh)
196     AYTOKINHTO.Text = CNull(r3!AYTOK)
        'parat.Text = CNull(r3!parat)

198     Combo4.Text = Combo4.List(0)
        Dim k As Integer
   
200     DTPicker1.Value = Now
202     Text1.Text = aritm2

        ' For k = 0 To Combo4.ListCount - 1
        '        If InStr(Left(Combo4.List(k), 3), Split(r3("TRP"), ";")(0)) > 0 Then
        '           Combo4.Text = Combo4.List(k)
        '           Exit For
        '        End If
        '    Next
        '

204     k = Val(Split(r3("TRP"), ";")(0)) - 1
206     If k >= 0 Then
208         Combo4.Text = Combo4.List(k)
        End If

210     r3.Close

        '--------------------------ΕΚΤΥΠΟΣΗ -----------------------------------------

        '  f_Akyr = True
212     f_Akyr = False
214     F_TYPOS_PAR = 175 ' για ελινε
216     PAR_GIA_AKYROSH.Caption = ""
218     DUM = printCrystal(atimakyr, CDate(Now))    'δημιουργει τα προσωρινά αρχεία
220     f_Akyr = False
222     If DUM < -1 Then
224         MsgBox "ΠΡΟΒΛΗΜΑ ΣΤΗΝ ΕΚΤΥΠΩΣΗ ΤΗΣ ΑΚΥΡΩΤΙΚΗΣ ΑΠΟΔΕΙΞΗΣ"
            Exit Sub
        End If
        ' If f_FORM_EKTY >= 90 And f_FORM_EKTY <= 99 Then     ' φορμα τύπου f90.txt dot matrix '====================================
        '          DUM = print_timol(f_matim, CDate(DTPicker1.Value))'
        'ElseIf Len(Trim(F_FORMA1)) > 1 Then
226     CrystalReport1.ReportFileName = mForma1
228     Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
        'CrystalReport1.Action = 1
            
        '        If Len(Trim(F_FORMA2)) > 1 Then
        '           CrystalReport1.ReportFileName = F_FORMA2
        '          Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
        ' CrystalReport1.Action = 1
        '     End If
        'Else
        '     CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
        '    Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
        'CrystalReport1.Action = 1
 
        'End If

        Exit Sub

ROLL:
        'Resume Next

230     'Resume Next
        'Resume Next
        'Resume Next
       'Resume Next
       Gdb.RollbackTrans

End Sub

Public Function GGET_FPA_NHSION(ByVal CATFPA As Integer) As Integer  ' EIDRecordset("FPA")) '

        'TYPOS   AYJON PERIGRAFH   TIMH    ID  LOG1    N1  C1  N2  C2  N3  C3  N4  C4
        '1        6                  4     2   NULL    6   NULL    NULL    NULL    NULL    NULL    NULL    NULL
        '1        1   NULL          13     4   NULL    2   EEEEEEEE                NULL    NULL    NULL    NULL    NULL    NULL
        '1        2   NULL          24     5   NULL    1   NULL    NULL    NULL    NULL    NULL    NULL    NULL

        '1        3   NULL           6     6   NULL    3   NULL    NULL    NULL    NULL    NULL    NULL    NULL

        '1        4   NULL          17     7   NULL    4   NULL    NULL    NULL    NULL    NULL    NULL    NULL

        '1        5   NULL           0     8   NULL    7   NULL    NULL    NULL    NULL    NULL    NULL    NULL
        '1        7                  9                 5

        '                        '    mydata.AddItem "1 ΦΠΑ συντελεστής 24%"
        '                        '    mydata.AddItem "2 ΦΠΑ συντελεστής 13% "
        '                        '    mydata.AddItem "3 ΦΠΑ συντελεστής 6%"
        '                        '    mydata.AddItem "4 ΦΠΑ συντελεστής 17%"    N1=MYDATA   AYJON=ΚΑΤΗΓ ΦΠΑ
        '                        '    mydata.AddItem "5 ΦΠΑ συντελεστής 9%"
        '                        '    mydata.AddItem "6 ΦΠΑ συντελεστής 4%"
        '                        '    mydata.AddItem "7 Άνευ Φ.Π.Α. 0%"
        '                        '    mydata.AddItem "8 Εγγραφές χωρίς ΦΠΑ (πχ Μισθοδοσία, Αποσβέσεις)"

        Dim N As Integer, L As Integer
100     N = GGET_NVALUE("SELECT TOP 1 TIMH FROM PINAKES WHERE TYPOS=1 AND  AYJON=" + str(CATFPA))  ' => 24

102     If N = 24 Then
104         L = GGET_NVALUE("SELECT TOP 1 AYJON FROM PINAKES WHERE N1=4 AND TYPOS=1")  ' => 17% (MYDATA=4 =>N1=4 TIMH=17% AYJON=4)

106     ElseIf N = 13 Then
108         L = GGET_NVALUE("SELECT TOP 1 AYJON FROM PINAKES WHERE N1=5 AND TYPOS=1")  ' => 9% (MYDATA=5 =>N1=5 TIMH=9% AYJON=7)

110     ElseIf N = 6 Then
112         L = GGET_NVALUE("SELECT TOP 1 AYJON FROM PINAKES WHERE N1=6 AND TYPOS=1")  ' => 4% (MYDATA=6 =>N1=6 TIMH=4% AYJON=4)    =>6
        Else
114         L = CATFPA
        End If

116     If L = 0 Then
118         L = CATFPA
        End If

120     GGET_FPA_NHSION = L

End Function
 
    ' ExportReportToPDF(ByVal FILENAME As String, ByVal ReportTitle As String, pdf_filename As String) As Long
Function Export2ReportToPDF(ByVal FILENAME As String, _
                            ByVal ReportTitle As String, _
                            pdf_filename As String) As Long

        On Error GoTo lab999

        ' Dim XXX2 As Printer
        '   Dim pdfExists As Integer
        '   pdfExists = 0
        '   Dim printerName As String
        '                  For Each XXX2 In Printers
        '
        '                     If InStr(UCase(XXX2.DeviceName), F_PDFNAME) > 0 Then
        '                            ' Set printer as system default.
        '                            printerName = XXX2.DeviceName
        '                            Set Printer = XXX2
        '                             pdfExists = 1
        '                            ' Stop looking for a printer.
        '                            Exit For
        '
        '                        End If
        '
        '                    Next
                    
        ' ΚΡΑΤΑΩ ΤΗΝ ΡΥΘΜΙΣΗ ΓΙΑ ΝΑ ΜΗΝ ΤΗΝ ΧΑΛΑΣΩ
100     Dim M_F_EKTYPOTHS As String: M_F_EKTYPOTHS = F_EKTYPOTHS
102     F_EKTYPOTHS = "MERCURYPDF"
104     CrystalReport1.ReportFileName = F_FORMA1
106     Kill "c:\pdf\pdf.pdf"
108     Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
110     MILSEC 1000
112     FileCopy "c:\pdf\pdf.pdf", "c:\pdf\" + pdf_filename

114     Export2ReportToPDF = FileLen("c:\pdf\" + pdf_filename)

        'ΕΠΑΝΑΦΕΡΩ ΤΗΝ ΡΥΘΜΙΣΗ
116     F_EKTYPOTHS = M_F_EKTYPOTHS

        '  'pdf_filename = pdf_filename
        '        Dim crxApp As CRAXdDRT.Application
        '        Dim ReportObject As CRAXdDRT.Report
        '       Set crxApp = New CRAXdDRT.Application
        '        'ReportObject = crxApp.OpenReport(FILENAME)
        '
        '        Set ReportObject = crxApp.OpenReport(FILENAME)
        '        Dim objExportOptions As CRAXdDRT.ExportOptions
        '        ReportObject.ReportTitle = ReportTitle
        '        With ReportObject
        '            .EnableParameterPrompting = False
        '            .MorePrintEngineErrorMessages = True
        '        End With
        '        ' objExportOptions = ReportObject.ExportOptions
        '        Set objExportOptions = ReportObject.ExportOptions
        '        With objExportOptions
        '            .DestinationType = CRAXdDRT.CRExportDestinationType.crEDTDiskFile '   CRAXDDRT.crEDTDiskFile
        '            .DiskFileName = "c:\pdf\" + pdf_filename ' FILENAME
        '            .FormatType = CRAXdDRT.CRExportFormatType.crEFTPortableDocFormat 'crEFTPortableDocFormat
        '            .PDFExportAllPages = True
        '        End With
        '        ReportObject.Export False ' True
        '        'Dim myFile As New FileInfo(pdf_filename)
        '        ' Dim sizeInBytes As Long =
        '        Export2ReportToPDF = 1 ' myFile.Length
        Exit Function
        ' MsgBox "δεν αποθηκευτηκε"
lab999:
        'ReportObject.Export True ' χειροκινητα
118     Resume Next

End Function

'νεο κομματι
' θα μπει μεσα στο par1

'                Dim FcrystalFilename As String: FcrystalFilename = F_FORMA1
'                FcrystalFilename = "C:\MERCVB\REPORTS\TIMOL1.RPT"
'                f_matim = "00000" 'DEBUG
'                Dim fFILE As String: fFILE = Trim(f_matim) + ".PDF"  '"c:\mercvb\" +
'                Dim Size As Long: Size = ExportReportToPDF(FcrystalFilename, "τιμολογιο", fFILE)
'                       fFILE = "C:\PDF\PDF.PDF"  'debug
'                       uploadAzure str(id_num), fFILE, Size

Public Sub uploadAzure(ByVal id_num As Long, FILENAME As String, Size As Long)

        Dim URL As String

        URL = get_url(id_num, FILENAME, Size)

128     Dim upfilename As String: upfilename = "C:\pdf\" + FILENAME
            
            
           ' URL = get_url(id_num, FILENAME, Size)
            
            
        On Error Resume Next
           
130     UploadAfIle URL, upfilename

132     '  VerifyUpload (id_num)

        Dim k As Integer

        For k = 1 To 4
           Dim MCC As String
            MCC = VerifyUpload(id_num)
            If Len(MCC) = "True" Then
                'continue
                Exit For
            Else

                k = k ' continue

            End If
            
            
           ' MCC = VerifyUpload(id_num)
            

            'url = get_url(id_num, FILENAME, Size)
            
           ' url = get_url(id_num, FILENAME, Size)
            
            UploadAfIle URL, upfilename
         
        Next

134    ' UploadAfIle URL, upfilename


download (id_num)

End Sub

Function get_url(ByVal id_num As Long, FILENAME As String, Size As Long) As String
'======================================================================================
  
        Dim ResultJSON  As String
            
        Dim jsonContent As String
100         jsonContent = ResultJSON ' " {""externa

        Dim REQ As WinHttp.WinHttpRequest
102     Set REQ = New WinHttp.WinHttpRequest
            
104     Dim URI As String: URI = gfURL1 + "api/uploadFileRequest" '?Bearer=" + gf_Bearer /api/
        Dim Q   As String
            
      
106     Dim URL As String: URL = ""
    On Error Resume Next
    
108     ResultJSON = "{  ""externalSystemId"":""" + LTrim(Trim(str(id_num))) + """ , ""filename"":""" + "C:\\PDF\\" + FILENAME + """, ""fileSize"": " + " " + LTrim(str(Size)) + "}   "  'debug
      
110     With REQ
112         .Open "POST", URI, async:=False
114         .setRequestHeader "Content-Type", "application/hal+json"
116         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
118         .setRequestHeader "Authorization", "Bearer " + gf_Bearer
120         .send (ResultJSON)
        End With
     
            Dim p

122         Dim sInputJson As String: ResultJSON = REQ.responseText
           
 
124        ' Set p = JSON.parse(sInputJson)
126        ' url = p.Item("uploadUrl")  ' Split(Split(sInputJson, ":")(1), ",")(0)

   'get_url = url
   
'   Set p = JSON.parse(sInputJson)
'   url = p.Item("uploadUrl")  ' Split(Split(sInputJson, ":")(1), ",")(0)
'   get_url = url
   
   
   Set p = JSON.parse(ResultJSON)
   URL = p.Item("uploadUrl")  ' Split(Split(sInputJson, ":")(1), ",")(0)
  ' ResultJSON = url
   
   
   get_url = URL
   
End Function




Private Sub uploadAzurefile(url2 As String, FILENAME As String)

        On Error GoTo lab777

        'Dim ResultJSON As String = JsonConvert.SerializeObject(SOU).ToString
        'Debug.Print(ResultJSON) '

        '  Exit Sub
        'Dim jsonContent As String = ResultJSON ' " {""externa

        Dim REQ As WinHttp.WinHttpRequest
        'REQ = New WinHttp.WinHttpRequest
100     Dim URI As String: URI = url2 '+ " HTTP/1.1" ' "https://beta-srv.parochos.gr/api/uploadFileRequest"  '?Bearer=" + gf_Bearer /api/
            
102     Set REQ = New WinHttp.WinHttpRequest
        Dim Q          As String

        Dim pdfBytes() As Byte
        'ΝΑ ΒΡΕΘΕΙ ΤΟ ΑΝΤΙΣΤΟΙΧΟ VB6  pdfBytes = file.ReadAllBytes(FILENAME)

        Dim F          As Long
        'Dim B() As Byte
 
104     F = FreeFile
106     Open FILENAME For Binary As F
        'ReDim pdfBytes(LOF(F))
        'Get F, , pdfBytes()
108     pdfBytes = input(LOF(F), #F)
110     Close F

112     With REQ
114         .Open "PUT", URI, async:=False
116         .setRequestHeader "Content-Type", "application/pdf"
118         .setRequestHeader "x-ms-blob-type", "BlockBlob"
120         .send (pdfBytes) '("c:\mercvb\" + filename)

        End With
        
122     Dim sInputJson As String: sInputJson = REQ.responseText
        
124     MsgBox REQ.responseText

        Exit Sub

lab777:

126     MsgBox Err.Description
128     Resume Next
End Sub
       
Public Sub UploadAfIle(surl As String, sFileName As String)
        Dim adoStream As Object
100     Set adoStream = CreateObject("ADODB.Stream")
102     adoStream.Mode = 3          ' read write
104     adoStream.Type = 1          ' adTypeBinary
106     adoStream.Open
108     adoStream.LoadFromFile (sFileName)
110     With CreateObject("Microsoft.XMLHTTP")
112         adoStream.Position = 0
114         .Open "PUT", surl, False
116         .setRequestHeader "Content-Length", "0" 'this is not a must
118         .setRequestHeader "x-ms-blob-type", "BlockBlob"
120         .send adoStream.Read(adoStream.Size)
        End With
122     Set adoStream = Nothing
End Sub
       
Private Function VerifyUpload(mexternalSystemId As Long) As String
100     Dim ResultJSON As String: ResultJSON = "{    ""externalSystemId"": """ + " " + LTrim(str(mexternalSystemId)) + """ }" ' JsonConvert.SerializeObject(SOU).toString

102     Dim jsonContent As String: jsonContent = ResultJSON '


ResultJSON = "{    ""externalSystemId"": """ + LTrim(str(mexternalSystemId)) + """ }"  ' JsonConvert.Serialize
jsonContent = ResultJSON '

        Dim REQ         As WinHttp.WinHttpRequest
104     Set REQ = New WinHttp.WinHttpRequest
106     Dim URI: URI = gfURL1 + "api/Finalize" '?Bearer=" + gf_Bearer /api/

108     Dim URL As String: URL = ""

110     With REQ
112         .Open "POST", URI, async:=False
114         .setRequestHeader "Content-Type", "application/hal+json"
116         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
118         .setRequestHeader "Authorization", "Bearer " + gf_Bearer

120         .send (jsonContent) '"{    ""externalSystemId"": """ + mexternalSystemId + """ }")
        End With
        
122     Dim sInputJson As String: sInputJson = REQ.responseText
         Dim p As Object
         On Error Resume Next
          Set p = JSON.parse(sInputJson)
          Dim ERRORS As String
         ERRORS = p.Item("errorMessage")  ' Split(Split(sInputJson, ":")(1), ",")(0)
         
         VerifyUpload = False
        VerifyUpload = p.Item("signing")("pdfUploaded")
         
       ' VerifyUpload = ERRORS
End Function

Private Sub download(mexternalSystemId As Long)

100     Dim ResultJSON  As String: ResultJSON = "{    ""externalSystemId"": """ + " " + LTrim(str(mexternalSystemId)) + """ }"  ' JsonConvert.SerializeObject(SOU).toString
102     Dim jsonContent As String: jsonContent = ResultJSON '

        Dim REQ         As WinHttp.WinHttpRequest
104     Set REQ = New WinHttp.WinHttpRequest
106     Dim URI: URI = gfURL1 + "api/getFiles" '?Bearer=" + gf_Bearer /api/
108     Dim URL As String: URL = ""
110     With REQ
112         .Open "POST", URI, async:=False
114         .setRequestHeader "Content-Type", "application/hal+json"
116         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
118         .setRequestHeader "Authorization", "Bearer " + gf_Bearer
120         .send (jsonContent) '"{    ""externalSystemId"": """ + mexternalSystemId + """ }")
                
        End With
            
122     Dim sInputJson As String: sInputJson = REQ.responseText
            
End Sub


  Public Sub toMydata()
  Dim m_ID_NUM As String
  
  m_ID_NUM = TDBGrid1.columns("ID_NUM") ' ή TDBGrid1.Col = 7: TDBGrid1.Text
  

536         If f_MYDATAcont = 0 Then
              MsgBox "ενεργοποιηστε την αυτοματη αποστολη"

            Else
                Dim stelnei As Integer
                Dim MATIM As String
                TDBGrid1.Col = 1
                MATIM = TDBGrid1.Text
                Dim MARK As String
                TDBGrid1.Col = 8
                MARK = TDBGrid1.columns("ENTITYMARK")
                
                
                
                Dim markReal As String: markReal = ""
                stelnei = GGET_NVALUE("SELECT ISNULL(ISMYDATA,0) FROM PARASTAT WHERE EIDOS='" & Left(MATIM, 1) & "'")
                
538             If stelnei = 0 Then
                   MsgBox "ΔΕΝ ΕΙΝΑΙ ΡΥΘΜΙΣΜΕΝΟ ΓΙΑ MYDATA ΤΟ ΒΠΑΡΑΣΤΑΤΙΚΟ"
                Else

550                 If Val(Left(MARK, 1)) > 0 Or Left(MARK, 3) = "ΔΟΚ" Then
                        MsgBox "εχει ηδη ανεβει στο Mydata"
                    Else
552                     gisOnlineMydata = 1
554                     Par7MyData.where.Caption = " where ID_NUM=" + str(m_ID_NUM) + "  " ' M_idnum
556                     Par7MyData.Visible = False
                     
558                     If F_PAROCHOS = 1 Then
                            Dim fFILE As String ': fFILE = Trim(MATIM) + ".PDF"  '"c:\mercvb\" +
                            Dim Size As Long
668                         'Size = Export2ReportToPDF(F_FORMA1, Left(MATIM, 7), fFILE)
                            fFILE = Left(MATIM, 7) + ".PDF" '"c:\mercvb\" +
                            
                            Size = Export2ReportToPDF(F_FORMA1, Left(MATIM, 7), fFILE)

560                         Par7MyData.loadpar7
562                         Dim ubl As String: ubl = GGET_CVALUE("SELECT left(ISNULL(UBL,'000'),3) AS UBL2 FROM PARASTAT WHERE EIDOS='" & Left(MATIM, 1) & "'")

                            'e?? paroxo mono sthn xondriki kai to parastattiko anti na exei 380 exei "mydata"
564                         If ubl = "000" Then 'mydata
566                             Par7MyData.ToXMLsub 0, markReal
                            Else
568                             Par7MyData.ToJason 0, 1  ' kane kai to upload
 
'                              If ubl <> "000" Then
'                                If GGET_NVALUE("select MYDATAISDEMO from MEM") = 1 Then
'                                   PAROX_URLADD = "beta-"
'                                Else
'                                   PAROX_URLADD = ""
'                                End If
'                                uploadAzure (m_ID_NUM), fFILE, Size
'                              End If




                            End If
                     
                            'Par7MyData.ToJason 0
                        Else
                     
570                         Par7MyData.ToXMLsub 0, markReal
                        End If
                        LASTTIMOL.Refresh
                        MsgBox markReal
                        
                    End If
                End If
            End If


    End Sub

Public Sub menu_ektyp_par(ByVal Ektyp1Email2 As Integer)
      
        
        '<EhHeader>
        On Error GoTo ektyp_par_Click_Err

        '</EhHeader>

        Dim m As Integer

        Dim c As String

        PAR1.TDBGrid1.Col = 2
        c = Replace(PAR1.TDBGrid1.Text, ",", ".")
        
        
           Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)
        
        
        PAR1.METASX_SE_TIM 99, c, m_ID, False, 0
        ' PAR1.METASX_SE_TIM 99, C, 0, False

        ' 100     PAR1.diortosis
110     DoEvents
If Ektyp1Email2 = 2 Then
     EPANEKTYPOSI (3)
Else
     EPANEKTYPOSI (0)   '  PAR1.Command5_Click
End If
    
    
    
        PAR1.Text2(0).Text = "" ' για να μην διορθώση κατα λαθος
        PAR1.PARAS.Text = ""
        PAR1.Text1.Text = ""
        PAR1.b_Katax.Enabled = False
        F_DIOR = 0
        
        
        '<EhFooter>
        Exit Sub

ektyp_par_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.ektyp_par_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.ektyp_par_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


Public Sub menu_diak_status()
'https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=NziZ2b5egY41ooVijo%2fjtSMRj7CvuLB%2fxG1YnEyrsSdlsH7ef9ZcgJZCLIFhuHMUa6Cp8pQjpI4ErLYp4fpsmjaTE7asa34Phn2SIB5cM4c%3d
''https://mydataapidev.aade.gr/GetDeliveryNoteStatus?mark=%7bmark%7d
'Dim MARK As String


     Dim c As String

        PAR1.TDBGrid1.Col = 8
        c = Replace(PAR1.TDBGrid1.Text, "ΔΟΚ", "")

       ' MARK = c '  InputBox("δωσε τo mark", "μαρκ")

Dim DUM As String
DUM = Par7MyData.GetDeliveryStatus(c, "aaa")
If DUM = "REGISTERED" Then
   DUM = "ΕΚΔΟΣΗ (REGISTERED)"
ElseIf DUM = "IN_TRANSIT" Then
   DUM = "ΣΕ ΜΕΤΑΦΟΡΑ (IN_TRANSIT)"
ElseIf DUM = "DELIVERED_BY_CARRIER" Then
   DUM = "ΠΑΡΕΛΗΦΘΗ (DELIVERED_BY_CARRIER)"
Else

End If



MsgBox DUM


End Sub

Public Sub menuParadosis()



Dim qrl As String: qrl = "https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=vZGdPC4qtv61lfj3tMcBWCVzle6HNmunsN6JjmX4e9gLqhLIIIUNGzIx6YzWdXmeg0fsTOtxriHc%2bCRo3shXRfuFVdEayG8L9Xt9%2fSH8HaE%3d"

           Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)

'"https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=NziZ2b5egY41ooVijo%2fjtSMRj7CvuLB%2fxG1YnEyrsSdlsH7ef9ZcgJZCLIFhuHMUa6Cp8pQjpI4ErLYp4fpsmjaTE7asa34Phn2SIB5cM4c%3d"
'https://mydataapidev.aade.gr/RegisterTransfer
qrl = GGET_CVALUE("select QRURL FROM TIM WHERE ID_NUM=" + str(m_ID))
'https://mydatapi.aade.gr/myDATA/ConfirmDeliveryOutcome
Dim TXML As String
TXML = "<?xml version=""1.0"" encoding=""utf-8""?>"   'qrUrl,outcome"
TXML = TXML + "<ConfirmDeliveryOutcomeRequest>"
TXML = TXML + "<qrUrl>" + qrl + "</qrUrl>"
TXML = TXML + "<outcome>FULL</outcome>"
TXML = TXML + "</ConfirmDeliveryOutcomeRequest>"

Dim DUM As String: DUM = Par7MyData.SendConfirm(TXML, " ")

End Sub


Public Sub menu_enhmPLHR()

Dim c, D, SX As String
        'd = TDBGrid1.Columns("ID_NUM").value ' ή TDBGrid1.Col = 7: TDBGrid1.Text




        Dim midnum As String:
           PAR1.TDBGrid1.Col = 7
       midnum = Val(PAR1.TDBGrid1.Text)
        
       ' midnum = (tdbgrid1.CurrentRow.cells(7).Value)
        'midnum = InputBox("id_num")
        
        
        
        Dim cmark  As String
        cmark = GGET_CVALUE("select ENTITYMARK FROM TIM WHERE ID_NUM=" + midnum)
        
        
        
        Dim posSignature As String: posSignature = GGET_CVALUE("select POSPAYSIGN FROM TIM WHERE ID_NUM=" + midnum)
        
        
        Dim POSTransactionId As String: POSTransactionId = GGET_CVALUE("select TRANSACTIONID FROM TIM WHERE ID_NUM=" + midnum)

        Dim POS_AMOUNT As Single: POS_AMOUNT = GGET_NVALUE("select AJI FROM TIM WHERE ID_NUM=" + midnum)
        
        Dim AMOUNT As String: AMOUNT = Format(POS_AMOUNT, "###0.00")
        
        
        'AMOUNT = InputBox("??S?", "??S?", AMOUNT)

        'sqlDt.Rows(klm)("TRP")
        
        
        
        Dim trp As String: trp = GGET_CVALUE("select TRP FROM TIM WHERE ID_NUM=" + midnum)


     
        Dim TID As String: TID = GGET_CVALUE("select ISNULL(C1,'') AS C1 from PINAKES where TYPOS=12 AND  AYJON=" + mID(trp, 1, 1))
        'Dim bearer As String = PAR7MYDATA.REQPAROCHOS()
        Par7MyData.loadpar7
        TID = Trim(TID)
        Par7MyData.POS_submitpayment cmark, midnum, posSignature, TID, AMOUNT, POSTransactionId




End Sub
