VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form bohu3 
   BackColor       =   &H00FF0000&
   Caption         =   "Παραμ.Παραστατικών"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   6615
      Left            =   0
      TabIndex        =   21
      Top             =   2160
      Width           =   15615
      Begin TabDlg.SSTab SSTab1 
         Height          =   5352
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   13176
         _ExtentX        =   23230
         _ExtentY        =   9446
         _Version        =   393216
         Tabs            =   6
         Tab             =   2
         TabHeight       =   420
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "ΛΟΓΙΣΤΙΚΗ"
         TabPicture(0)   =   "bohu3.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lblLabels(13)"
         Tab(0).Control(1)=   "Label6"
         Tab(0).Control(2)=   "LABEL1(51)"
         Tab(0).Control(3)=   "LABEL1(52)"
         Tab(0).Control(4)=   "LABEL1(53)"
         Tab(0).Control(5)=   "LABEL1(54)"
         Tab(0).Control(6)=   "LABEL1(55)"
         Tab(0).Control(7)=   "Label8"
         Tab(0).Control(8)=   "LABEL1(57)"
         Tab(0).Control(9)=   "LABEL1(56)"
         Tab(0).Control(10)=   "Label7"
         Tab(0).Control(11)=   "Label9"
         Tab(0).Control(12)=   "Label111"
         Tab(0).Control(13)=   "Label13"
         Tab(0).Control(14)=   "txtFields(14)"
         Tab(0).Control(15)=   "Command1"
         Tab(0).Control(16)=   "log1(1)"
         Tab(0).Control(17)=   "log1(2)"
         Tab(0).Control(18)=   "log1(3)"
         Tab(0).Control(19)=   "log1(4)"
         Tab(0).Control(20)=   "log1(5)"
         Tab(0).Control(21)=   "log1(0)"
         Tab(0).Control(22)=   "log1(7)"
         Tab(0).Control(23)=   "log1(6)"
         Tab(0).Control(24)=   "log2(1)"
         Tab(0).Control(25)=   "log2(2)"
         Tab(0).Control(26)=   "log2(3)"
         Tab(0).Control(27)=   "log2(4)"
         Tab(0).Control(28)=   "log2(5)"
         Tab(0).Control(29)=   "log2(0)"
         Tab(0).Control(30)=   "mCopy"
         Tab(0).Control(31)=   "log2(7)"
         Tab(0).Control(32)=   "log2(6)"
         Tab(0).Control(33)=   "log2(16)"
         Tab(0).Control(34)=   "log2(17)"
         Tab(0).Control(35)=   "mPaste"
         Tab(0).ControlCount=   36
         TabCaption(1)   =   "MYDATA-ΜΥΦ-ΦΟΡΟΛ"
         TabPicture(1)   =   "bohu3.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblLabels(25)"
         Tab(1).Control(1)=   "lblLabels(24)"
         Tab(1).Control(2)=   "lblLabels(23)"
         Tab(1).Control(3)=   "lblLabels(22)"
         Tab(1).Control(4)=   "lblLabel11"
         Tab(1).Control(5)=   "lblΑριθμΤαμειακής"
         Tab(1).Control(6)=   "Label12"
         Tab(1).Control(7)=   "lblLabels(10)"
         Tab(1).Control(8)=   "Label15"
         Tab(1).Control(9)=   "lblLabels(26)"
         Tab(1).Control(10)=   "lblLabels(27)"
         Tab(1).Control(11)=   "lblΠροσοχήΤο"
         Tab(1).Control(12)=   "MYISAYTOTIM"
         Tab(1).Control(13)=   "mycateg"
         Tab(1).Control(14)=   "myTypEsod"
         Tab(1).Control(15)=   "MYEID"
         Tab(1).Control(16)=   "myf"
         Tab(1).Control(17)=   "tameiaki"
         Tab(1).Control(18)=   "typos_par"
         Tab(1).Control(19)=   "ELINES"
         Tab(1).Control(20)=   "txtFields(10)"
         Tab(1).Control(21)=   "mydata"
         Tab(1).Control(22)=   "myTypEXod"
         Tab(1).Control(23)=   "myCatExod"
         Tab(1).Control(24)=   "ISMYDATA"
         Tab(1).Control(25)=   "isDiakin"
         Tab(1).Control(26)=   "ComboUBL"
         Tab(1).Control(27)=   "Checkb2g"
         Tab(1).ControlCount=   28
         TabCaption(2)   =   "ΦΟΡΜΕΣ-ΑΡΙΘΜΗΣΗ"
         TabPicture(2)   =   "bohu3.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "ΑΡΙΘΜΗΣΗ"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "lblLabels(16)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "lblLabels(17)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "lblLabels(11)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "lblLabels(12)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "lblΠροεπιλεγμένηΑπαλλαγή"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Label10"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "SEIRES"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "Frame3"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "ARITMISI"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "txtFields(11)"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "txtFields(16)"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "txtFields(17)"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "txtFields(12)"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "List1"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "chKvAT"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).Control(16)=   "DEFAULTAPALL"
         Tab(2).Control(16).Enabled=   0   'False
         Tab(2).Control(17)=   "DESMIA"
         Tab(2).Control(17).Enabled=   0   'False
         Tab(2).Control(18)=   "ApallFpa"
         Tab(2).Control(18).Enabled=   0   'False
         Tab(2).ControlCount=   19
         TabCaption(3)   =   "Κινήσεις-Ενημερώσεις"
         TabPicture(3)   =   "bohu3.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lblLabels(15)"
         Tab(3).Control(1)=   "lblLabels(3)"
         Tab(3).Control(2)=   "lblLabels(4)"
         Tab(3).Control(3)=   "lblLabels(5)"
         Tab(3).Control(4)=   "lblLabels(6)"
         Tab(3).Control(5)=   "Label3"
         Tab(3).Control(6)=   "Label5"
         Tab(3).Control(7)=   "txtFields(15)"
         Tab(3).Control(8)=   "txtFields(3)"
         Tab(3).Control(9)=   "AJIA_APOU"
         Tab(3).Control(10)=   "POS_APOU"
         Tab(3).Control(11)=   "PEL"
         Tab(3).ControlCount=   12
         TabCaption(4)   =   "Μετασχηματισμοί Ακυρώσεις"
         TabPicture(4)   =   "bohu3.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lblLabels(14)"
         Tab(4).Control(1)=   "lblLabels(18)"
         Tab(4).Control(2)=   "lblLabels(8)"
         Tab(4).Control(3)=   "lblLabels(7)"
         Tab(4).Control(4)=   "txtFields(13)"
         Tab(4).Control(5)=   "akyr"
         Tab(4).Control(6)=   "txtFields(8)"
         Tab(4).Control(7)=   "STADIO"
         Tab(4).Control(8)=   "ComboAKYR"
         Tab(4).Control(9)=   "combometasx"
         Tab(4).ControlCount=   10
         TabCaption(5)   =   "Λοιπά"
         TabPicture(5)   =   "bohu3.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "lblLabels(19)"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "lblΠΩΛΗΤΕΣ"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).Control(2)=   "Label2"
         Tab(5).Control(2).Enabled=   0   'False
         Tab(5).Control(3)=   "Label11"
         Tab(5).Control(3).Enabled=   0   'False
         Tab(5).Control(4)=   "Label16"
         Tab(5).Control(4).Enabled=   0   'False
         Tab(5).Control(5)=   "AYTOKINHTO"
         Tab(5).Control(5).Enabled=   0   'False
         Tab(5).Control(6)=   "YP1"
         Tab(5).Control(6).Enabled=   0   'False
         Tab(5).Control(7)=   "yp2"
         Tab(5).Control(7).Enabled=   0   'False
         Tab(5).Control(8)=   "yp3"
         Tab(5).Control(8).Enabled=   0   'False
         Tab(5).Control(9)=   "POLITES"
         Tab(5).Control(9).Enabled=   0   'False
         Tab(5).Control(10)=   "Parakratisi"
         Tab(5).Control(10).Enabled=   0   'False
         Tab(5).Control(11)=   "yp4"
         Tab(5).Control(11).Enabled=   0   'False
         Tab(5).Control(12)=   "log1(8)"
         Tab(5).Control(12).Enabled=   0   'False
         Tab(5).Control(13)=   "foros"
         Tab(5).Control(13).Enabled=   0   'False
         Tab(5).Control(14)=   "katforou"
         Tab(5).Control(14).Enabled=   0   'False
         Tab(5).Control(15)=   "proepilFPA"
         Tab(5).Control(15).Enabled=   0   'False
         Tab(5).ControlCount=   16
         Begin VB.TextBox proepilFPA 
            Height          =   288
            Left            =   -70560
            TabIndex        =   155
            Top             =   4920
            Width           =   1692
         End
         Begin VB.ComboBox katforou 
            Height          =   315
            Left            =   -68040
            Style           =   2  'Dropdown List
            TabIndex        =   154
            Top             =   3360
            Width           =   4692
         End
         Begin VB.ComboBox foros 
            Height          =   315
            ItemData        =   "bohu3.frx":00A8
            Left            =   -68040
            List            =   "bohu3.frx":00B8
            Style           =   2  'Dropdown List
            TabIndex        =   153
            Top             =   2760
            Width           =   4692
         End
         Begin VB.ComboBox ApallFpa 
            Height          =   315
            Left            =   0
            TabIndex        =   152
            Text            =   "Combo1"
            Top             =   3600
            Width           =   4932
         End
         Begin VB.TextBox DESMIA 
            Height          =   288
            Left            =   2400
            TabIndex        =   150
            Top             =   2520
            Width           =   372
         End
         Begin VB.CheckBox Checkb2g 
            Alignment       =   1  'Right Justify
            Caption         =   "Είναι για Δημόσιο"
            Height          =   375
            Left            =   -67026
            TabIndex        =   146
            Top             =   4920
            Width           =   1935
         End
         Begin VB.ComboBox ComboUBL 
            Height          =   288
            ItemData        =   "bohu3.frx":0115
            Left            =   -67146
            List            =   "bohu3.frx":0128
            Style           =   2  'Dropdown List
            TabIndex        =   145
            Top             =   4560
            Width           =   2055
         End
         Begin VB.ComboBox isDiakin 
            Height          =   288
            ItemData        =   "bohu3.frx":0195
            Left            =   -72600
            List            =   "bohu3.frx":01A2
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   2400
            Width           =   4215
         End
         Begin VB.TextBox DEFAULTAPALL 
            Height          =   285
            Left            =   5040
            TabIndex        =   140
            Top             =   3600
            Width           =   300
         End
         Begin VB.CheckBox chKvAT 
            Caption         =   "Υποχρεωτικό ΑΦΜ"
            Height          =   255
            Left            =   2400
            TabIndex        =   139
            Top             =   3120
            Width           =   3375
         End
         Begin VB.ComboBox combometasx 
            Height          =   288
            Left            =   -70080
            Style           =   2  'Dropdown List
            TabIndex        =   136
            Top             =   2760
            Width           =   5175
         End
         Begin VB.ComboBox ComboAKYR 
            Height          =   288
            Left            =   -70080
            Style           =   2  'Dropdown List
            TabIndex        =   135
            Top             =   1680
            Width           =   5175
         End
         Begin VB.TextBox log1 
            BackColor       =   &H0080FFFF&
            DataField       =   "LPARAKRAT"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   8
            Left            =   -70560
            TabIndex        =   133
            Top             =   3960
            Width           =   1692
         End
         Begin VB.CheckBox ISMYDATA 
            Alignment       =   1  'Right Justify
            Caption         =   "Ενημερώνει MyData"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -74724
            TabIndex        =   131
            Top             =   3550
            Width           =   2292
         End
         Begin VB.ComboBox myCatExod 
            Height          =   288
            ItemData        =   "bohu3.frx":0208
            Left            =   -72600
            List            =   "bohu3.frx":0239
            Style           =   2  'Dropdown List
            TabIndex        =   128
            Top             =   3240
            Width           =   7514
         End
         Begin VB.ComboBox myTypEXod 
            Height          =   288
            ItemData        =   "bohu3.frx":0523
            Left            =   -72600
            List            =   "bohu3.frx":0605
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   2760
            Width           =   7514
         End
         Begin VB.CommandButton mPaste 
            BackColor       =   &H0080FF80&
            Caption         =   "Επικόλληση Γεφύρωσης"
            Height          =   255
            Left            =   -67200
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   5280
            Width           =   2535
         End
         Begin VB.TextBox log2 
            DataField       =   "LFPA17"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   17
            Left            =   -68400
            TabIndex        =   85
            Top             =   4440
            Width           =   1455
         End
         Begin VB.TextBox log2 
            DataField       =   "LFPA16"
            DataSource      =   "adodc1"
            Height          =   375
            HideSelection   =   0   'False
            Index           =   16
            Left            =   -68400
            TabIndex        =   84
            Top             =   3960
            Width           =   1455
         End
         Begin VB.TextBox log2 
            DataField       =   "LFPA6"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   6
            Left            =   -70440
            TabIndex        =   83
            Top             =   3960
            Width           =   1815
         End
         Begin VB.TextBox log2 
            DataField       =   "LFPA7"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   7
            Left            =   -70440
            TabIndex        =   82
            Top             =   4440
            Width           =   1815
         End
         Begin VB.CommandButton mCopy 
            BackColor       =   &H0080FF80&
            Caption         =   "Αντιγραφή Γεφύρωσης"
            Height          =   255
            Left            =   -70440
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   5280
            Width           =   2535
         End
         Begin VB.TextBox log2 
            DataField       =   "PISPEL"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   0
            Left            =   -70440
            TabIndex        =   80
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox log2 
            DataField       =   "LFPA5"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   5
            Left            =   -70440
            TabIndex        =   79
            Top             =   3360
            Width           =   3015
         End
         Begin VB.TextBox log2 
            DataField       =   "LFPA4"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   4
            Left            =   -70440
            TabIndex        =   78
            Top             =   2880
            Width           =   3015
         End
         Begin VB.TextBox log2 
            DataField       =   "LFPA3"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   3
            Left            =   -70440
            TabIndex        =   77
            Top             =   2400
            Width           =   3015
         End
         Begin VB.TextBox log2 
            DataField       =   "LFPA2"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   2
            Left            =   -70440
            TabIndex        =   76
            Top             =   1920
            Width           =   3015
         End
         Begin VB.TextBox log2 
            DataField       =   "LFPA1"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   1
            Left            =   -70440
            TabIndex        =   75
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox log1 
            DataField       =   "LAJ6"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   6
            Left            =   -73200
            TabIndex        =   74
            Top             =   3960
            Width           =   2535
         End
         Begin VB.TextBox log1 
            DataField       =   "LAJ7"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   7
            Left            =   -73200
            TabIndex        =   73
            Top             =   4440
            Width           =   2535
         End
         Begin VB.TextBox log1 
            DataField       =   "XREPEL"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   0
            Left            =   -73200
            TabIndex        =   72
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox log1 
            DataField       =   "LAJ5"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   5
            Left            =   -73200
            TabIndex        =   71
            Top             =   3360
            Width           =   2535
         End
         Begin VB.TextBox log1 
            DataField       =   "LAJ4"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   4
            Left            =   -73200
            TabIndex        =   70
            Top             =   2880
            Width           =   2535
         End
         Begin VB.TextBox log1 
            DataField       =   "LAJ3"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   3
            Left            =   -73200
            TabIndex        =   69
            Top             =   2400
            Width           =   2535
         End
         Begin VB.TextBox log1 
            DataField       =   "LAJ2"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   2
            Left            =   -73200
            TabIndex        =   68
            Top             =   1920
            Width           =   2535
         End
         Begin VB.TextBox log1 
            DataField       =   "LAJ1"
            DataSource      =   "adodc1"
            Height          =   375
            Index           =   1
            Left            =   -73200
            TabIndex        =   67
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0080FF80&
            Caption         =   "Γέφυρες"
            Height          =   255
            Left            =   -73200
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   5280
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.ComboBox mydata 
            BackColor       =   &H000080FF&
            Height          =   288
            Left            =   -63360
            TabIndex        =   65
            Text            =   "Combo1"
            Top             =   2160
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H000080FF&
            DataField       =   "ETIK"
            DataSource      =   "adodc1"
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   -67440
            TabIndex        =   64
            Top             =   5640
            Visible         =   0   'False
            Width           =   5292
         End
         Begin VB.ComboBox ELINES 
            Height          =   288
            Left            =   -71760
            TabIndex        =   63
            Text            =   "Combo1"
            Top             =   4200
            Width           =   6669
         End
         Begin VB.TextBox typos_par 
            Height          =   285
            Left            =   -72600
            TabIndex        =   62
            Top             =   4200
            Width           =   855
         End
         Begin VB.TextBox tameiaki 
            DataField       =   "TAMEIAKI"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   -72600
            TabIndex        =   61
            Top             =   3840
            Width           =   1935
         End
         Begin VB.ComboBox myf 
            Height          =   288
            ItemData        =   "bohu3.frx":1616
            Left            =   -72600
            List            =   "bohu3.frx":1632
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "GEF_P"
            DataSource      =   "adodc1"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   14
            Left            =   -65040
            TabIndex        =   59
            Top             =   840
            Width           =   1692
         End
         Begin VB.ComboBox PEL 
            Height          =   288
            ItemData        =   "bohu3.frx":16B8
            Left            =   -72360
            List            =   "bohu3.frx":16CE
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   1680
            Width           =   3375
         End
         Begin VB.ComboBox POS_APOU 
            Height          =   288
            ItemData        =   "bohu3.frx":1738
            Left            =   -72360
            List            =   "bohu3.frx":174B
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   2160
            Width           =   3375
         End
         Begin VB.ComboBox AJIA_APOU 
            Height          =   288
            ItemData        =   "bohu3.frx":17A3
            Left            =   -72360
            List            =   "bohu3.frx":17B6
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   2580
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -72360
            TabIndex        =   55
            Top             =   1200
            Width           =   3375
         End
         Begin VB.ComboBox STADIO 
            Height          =   288
            ItemData        =   "bohu3.frx":180E
            Left            =   -70920
            List            =   "bohu3.frx":1827
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   2160
            Width           =   6015
         End
         Begin VB.ListBox List1 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "bohu3.frx":18A1
            Left            =   3240
            List            =   "bohu3.frx":18E1
            TabIndex        =   53
            Top             =   1932
            Width           =   2535
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00C0FFC0&
            DataField       =   "FORM_EKTYP"
            DataSource      =   "adodc1"
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   2400
            TabIndex        =   52
            Top             =   1932
            Width           =   372
         End
         Begin VB.TextBox txtFields 
            Enabled         =   0   'False
            Height          =   285
            Index           =   17
            Left            =   2400
            TabIndex        =   51
            Top             =   4428
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ARITMISI"
            DataSource      =   "adodc1"
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   2400
            TabIndex        =   50
            Top             =   4080
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -70920
            TabIndex        =   49
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "EKTYP"
            DataSource      =   "adodc1"
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   2400
            TabIndex        =   48
            Top             =   1560
            Width           =   3375
         End
         Begin VB.TextBox akyr 
            Height          =   285
            Left            =   -70920
            TabIndex        =   47
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   -70920
            TabIndex        =   46
            Top             =   3720
            Width           =   6015
         End
         Begin VB.TextBox txtFields 
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   -72360
            TabIndex        =   45
            Top             =   3480
            Width           =   3375
         End
         Begin VB.ListBox ARITMISI 
            Height          =   1035
            Left            =   5880
            TabIndex        =   44
            Top             =   4080
            Width           =   6015
         End
         Begin VB.CheckBox yp4 
            Alignment       =   1  'Right Justify
            Caption         =   "Eνεργό σε Υποκ3"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -72000
            TabIndex        =   43
            Top             =   3144
            Width           =   1695
         End
         Begin VB.TextBox Parakratisi 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   -70560
            TabIndex        =   42
            Top             =   3564
            Width           =   1692
         End
         Begin VB.TextBox POLITES 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   -70560
            TabIndex        =   41
            Top             =   4584
            Width           =   1692
         End
         Begin VB.CheckBox yp3 
            Alignment       =   1  'Right Justify
            Caption         =   "Eνεργό σε Υποκ2"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -72000
            TabIndex        =   40
            Top             =   2784
            Width           =   1695
         End
         Begin VB.CheckBox yp2 
            Alignment       =   1  'Right Justify
            Caption         =   "Eνεργό σε Υποκ1"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -72000
            TabIndex        =   39
            Top             =   2424
            Width           =   1695
         End
         Begin VB.CheckBox YP1 
            Alignment       =   1  'Right Justify
            Caption         =   "Eνεργό σε Κεντρ"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -72000
            TabIndex        =   38
            Top             =   2064
            Width           =   1695
         End
         Begin VB.ComboBox AYTOKINHTO 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   -70560
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1560
            Width           =   1575
         End
         Begin VB.ComboBox MYEID 
            Height          =   288
            ItemData        =   "bohu3.frx":1A9D
            Left            =   -72600
            List            =   "bohu3.frx":1B37
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   720
            Width           =   7514
         End
         Begin VB.ComboBox myTypEsod 
            Height          =   288
            ItemData        =   "bohu3.frx":2454
            Left            =   -72600
            List            =   "bohu3.frx":24BB
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1080
            Width           =   7514
         End
         Begin VB.ComboBox mycateg 
            Height          =   288
            ItemData        =   "bohu3.frx":2CBC
            Left            =   -72600
            List            =   "bohu3.frx":2CE4
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1560
            Width           =   7514
         End
         Begin VB.CheckBox MYISAYTOTIM 
            Height          =   252
            Left            =   -72600
            TabIndex        =   33
            Top             =   1920
            Width           =   252
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Εναλλακτικά"
            Height          =   2292
            Left            =   5880
            TabIndex        =   23
            Top             =   1440
            Width           =   6015
            Begin VB.CommandButton CMDFORMA1 
               Caption         =   "?"
               Height          =   300
               Left            =   4800
               TabIndex        =   29
               Top             =   240
               Width           =   252
            End
            Begin VB.TextBox FORMA1 
               Height          =   285
               Left            =   2520
               TabIndex        =   28
               Top             =   240
               Width           =   2175
            End
            Begin VB.CommandButton CMDFORMA2 
               Caption         =   "?"
               Height          =   300
               Left            =   4800
               TabIndex        =   27
               Top             =   840
               Width           =   252
            End
            Begin VB.TextBox FORMA2 
               Height          =   285
               Left            =   2520
               TabIndex        =   26
               Top             =   840
               Width           =   2175
            End
            Begin VB.CommandButton CMDFORMA3 
               Caption         =   "?"
               Height          =   300
               Left            =   4800
               TabIndex        =   25
               Top             =   1440
               Width           =   252
            End
            Begin VB.TextBox FORMA3 
               Height          =   285
               Left            =   2520
               TabIndex        =   24
               Top             =   1440
               Width           =   2175
            End
            Begin VB.Image Image1 
               Height          =   225
               Index           =   3
               Left            =   5280
               Picture         =   "bohu3.frx":2F4A
               Top             =   1440
               Width           =   240
            End
            Begin VB.Image Image1 
               Height          =   225
               Index           =   2
               Left            =   5280
               Picture         =   "bohu3.frx":347C
               Top             =   840
               Width           =   240
            End
            Begin VB.Image Image1 
               Height          =   225
               Index           =   1
               Left            =   5280
               Picture         =   "bohu3.frx":39AE
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Φόρμα 1ης Εκτύπωσης"
               Height          =   192
               Index           =   0
               Left            =   240
               TabIndex        =   32
               Top             =   360
               Width           =   1632
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Φόρμα 2ης Εκτύπωσης"
               Height          =   192
               Index           =   2
               Left            =   240
               TabIndex        =   31
               Top             =   840
               Width           =   1632
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Φόρμα 3ης Εκτύπωσης"
               Height          =   192
               Index           =   3
               Left            =   240
               TabIndex        =   30
               Top             =   1440
               Width           =   1632
            End
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Προεπιλεγμένο ΦΠΑ για όλα τα είδη"
            ForeColor       =   &H00FF0000&
            Height          =   192
            Left            =   -74040
            TabIndex        =   156
            Top             =   4920
            Width           =   2616
         End
         Begin VB.Label SEIRES 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "......"
            Height          =   192
            Left            =   3720
            TabIndex        =   151
            Top             =   2520
            Width           =   1536
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Σειρές εκτύπωσης 0/2"
            Height          =   192
            Left            =   600
            TabIndex        =   149
            Top             =   2520
            Width           =   1572
         End
         Begin VB.Label lblΠροσοχήΤο 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Προσοχή το 6% να είναι στην κατηγ.3"
            Height          =   195
            Left            =   -65040
            TabIndex        =   147
            Top             =   4200
            Width           =   2850
         End
         Begin VB.Label lblΠροεπιλεγμένηΑπαλλαγή 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Προεπιλεγμένη Απαλλαγή ΦΠΑ"
            Height          =   192
            Left            =   0
            TabIndex        =   141
            Top             =   3360
            Width           =   2316
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Αν θα φαίνονται στην καρτέλλα του πελάτη/προμηθευτή και σε ποιά στήλη(Χρ/Πισ)"
            Height          =   195
            Left            =   -68760
            TabIndex        =   138
            Top             =   1800
            Width           =   6300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Τα ακυρωτικά βαζω 1 για να φαίνονται στην λίστα αλλοιώς δεν τυπώνονται"
            Height          =   195
            Left            =   -68760
            TabIndex        =   137
            Top             =   1200
            Width           =   5745
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Προεπιλ.Κατηγ και Τύπος Φόρου π.χ. 3;6 Για λοιποί φόροι(3) και ξενοδοχ.1-2 αστέρια(6)"
            ForeColor       =   &H00FF0000&
            Height          =   384
            Left            =   -74040
            TabIndex        =   134
            Top             =   3960
            Width           =   3372
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "Δύο φόροι=3 Ενα Φόρος=1  Καθόλου=0"
            ForeColor       =   &H00FF0000&
            Height          =   252
            Left            =   -74040
            TabIndex        =   132
            Top             =   3600
            Width           =   3132
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MyData ΚΑΤΗΓ εξοδου"
            ForeColor       =   &H00FF0000&
            Height          =   192
            Index           =   27
            Left            =   -74640
            TabIndex        =   130
            Top             =   3240
            Width           =   1704
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MyData Τυπος Εξόδου"
            ForeColor       =   &H00FF0000&
            Height          =   192
            Index           =   26
            Left            =   -74640
            TabIndex        =   129
            Top             =   2880
            Width           =   1668
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Προτεινόμενοι συνδυασμοί Είδους-Τύπου-Κατηγορίας"
            Height          =   192
            Left            =   -71760
            TabIndex        =   126
            Top             =   1920
            Width           =   3864
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Πίστωση ΦΠΑ"
            Height          =   252
            Left            =   -68280
            TabIndex        =   125
            Top             =   3720
            Width           =   1452
         End
         Begin VB.Label Label111 
            BackStyle       =   0  'Transparent
            Caption         =   "Xρέωση ΦΠΑ"
            Height          =   252
            Left            =   -69960
            TabIndex        =   124
            Top             =   3720
            Width           =   1332
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "ΛΟΓ.ΑΞΙΑΣ ΦΠΑ"
            Height          =   252
            Left            =   -70320
            TabIndex        =   123
            Top             =   1200
            Width           =   2412
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "ΛΟΓ.ΠΙΣΤΩΣΗΣ"
            Height          =   252
            Left            =   -70200
            TabIndex        =   122
            Top             =   480
            Width           =   2172
         End
         Begin VB.Label LABEL1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   ".."
            ForeColor       =   &H00FF0000&
            Height          =   372
            Index           =   56
            Left            =   -74160
            TabIndex        =   121
            Top             =   4080
            Width           =   852
         End
         Begin VB.Label LABEL1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   ".."
            ForeColor       =   &H00FF0000&
            Height          =   252
            Index           =   57
            Left            =   -74640
            TabIndex        =   120
            Top             =   4440
            Width           =   1332
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "ΛΟΓ.ΚΑΘ.ΠΟΣΟΥ"
            Height          =   252
            Left            =   -73080
            TabIndex        =   119
            Top             =   1200
            Width           =   2412
         End
         Begin VB.Label LABEL1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   ".."
            ForeColor       =   &H00FF0000&
            Height          =   252
            Index           =   55
            Left            =   -74640
            TabIndex        =   118
            Top             =   3360
            Width           =   1332
         End
         Begin VB.Label LABEL1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   ".."
            ForeColor       =   &H00FF0000&
            Height          =   372
            Index           =   54
            Left            =   -74640
            TabIndex        =   117
            Top             =   3000
            Width           =   1332
         End
         Begin VB.Label LABEL1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            ForeColor       =   &H00FF0000&
            Height          =   372
            Index           =   53
            Left            =   -74760
            TabIndex        =   116
            Top             =   2520
            Width           =   1332
         End
         Begin VB.Label LABEL1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            ForeColor       =   &H00FF0000&
            Height          =   372
            Index           =   52
            Left            =   -74760
            TabIndex        =   115
            Top             =   2040
            Width           =   1332
         End
         Begin VB.Label LABEL1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            ForeColor       =   &H00FF0000&
            Height          =   372
            Index           =   51
            Left            =   -74760
            TabIndex        =   114
            Top             =   1560
            Width           =   1332
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "ΛΟΓ.ΧΡΕΩΣΗΣ"
            Height          =   252
            Left            =   -72960
            TabIndex        =   113
            Top             =   480
            Width           =   2172
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "MyData Κωδ"
            ForeColor       =   &H00FF0000&
            Height          =   264
            Index           =   10
            Left            =   -69360
            TabIndex        =   112
            Top             =   5280
            Width           =   1452
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Τύπος Παρ/κού για ELINE"
            ForeColor       =   &H00FF0000&
            Height          =   372
            Left            =   -74760
            TabIndex        =   111
            Top             =   4200
            Width           =   1932
         End
         Begin VB.Label lblΑριθμΤαμειακής 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Αριθμ.Ταμειακής"
            ForeColor       =   &H00FF0000&
            Height          =   192
            Left            =   -74040
            TabIndex        =   110
            Top             =   3840
            Width           =   1248
         End
         Begin VB.Label lblLabel11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ΜΥΦ"
            ForeColor       =   &H00FF0000&
            Height          =   192
            Left            =   -73320
            TabIndex        =   109
            Top             =   4644
            Width           =   396
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Γέφυρα αρ.εγγραφής"
            ForeColor       =   &H00FF0000&
            Height          =   252
            Index           =   13
            Left            =   -66720
            TabIndex        =   108
            Top             =   960
            Width           =   1572
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Αξία Αποθήκης"
            ForeColor       =   &H00FF0000&
            Height          =   252
            Index           =   6
            Left            =   -74280
            TabIndex        =   107
            Top             =   2640
            Width           =   1812
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ποσότητα Αποθήκης"
            ForeColor       =   &H00FF0000&
            Height          =   252
            Index           =   5
            Left            =   -74280
            TabIndex        =   106
            Top             =   2196
            Width           =   1812
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Χρέωση/ Πίστωση"
            ForeColor       =   &H00FF0000&
            Height          =   252
            Index           =   4
            Left            =   -74280
            TabIndex        =   105
            Top             =   1680
            Width           =   1812
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Πωλήσες=1 Αγορές=2"
            ForeColor       =   &H00FF0000&
            Height          =   252
            Index           =   3
            Left            =   -74280
            TabIndex        =   104
            Top             =   1200
            Width           =   1812
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1Προσφ 2Παραγγ 3=Δ.Α 4=τιμ 5=ενδοδ 6=ΠΥ"
            ForeColor       =   &H00FF8080&
            Height          =   360
            Index           =   7
            Left            =   -74880
            TabIndex        =   103
            Top             =   2160
            Width           =   3852
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Φόρμα Εκτύπωσης"
            ForeColor       =   &H00FF8080&
            Height          =   252
            Index           =   12
            Left            =   480
            TabIndex        =   102
            ToolTipText     =   "11 LASER 10 LASER ΜΙΑ ΕΚΤΥΠΩΣΗ 12 LASER ΦΟΡΜΑ ΤΙΜΟL2  99 ΑΚΙΔΩΝ F90.TXT ΦΟΡΜΑ  18 τριπλο 13=TIMOL3"
            Top             =   1920
            Width           =   1812
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Εκτυπώνεται"
            ForeColor       =   &H00FF8080&
            Height          =   252
            Index           =   11
            Left            =   480
            TabIndex        =   101
            Top             =   1560
            Width           =   1812
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Τελευταίος αριθμός :"
            ForeColor       =   &H00FF8080&
            Height          =   360
            Index           =   17
            Left            =   720
            TabIndex        =   100
            Top             =   4488
            Width           =   1812
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Αρίθμηση από μετρητή Ν"
            ForeColor       =   &H00FF8080&
            Height          =   360
            Index           =   16
            Left            =   720
            TabIndex        =   99
            Top             =   4152
            Width           =   1812
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Μετασχηματίζεται από"
            ForeColor       =   &H00FF8080&
            Height          =   252
            Index           =   8
            Left            =   -72840
            TabIndex        =   98
            Top             =   2760
            Width           =   1812
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ακυρωτικό"
            ForeColor       =   &H00FF8080&
            Height          =   252
            Index           =   18
            Left            =   -72360
            TabIndex        =   97
            Top             =   1800
            Width           =   1332
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Τιμολόγηση μόνο από Παλάτες κατηγορίας:"
            ForeColor       =   &H00FF8080&
            Height          =   432
            Index           =   14
            Left            =   -74640
            TabIndex        =   96
            Top             =   3720
            Width           =   3828
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Χονδ=1  ΕΛΓΑ=3"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   15
            Left            =   -74160
            TabIndex        =   95
            Top             =   3540
            Width           =   1215
         End
         Begin VB.Label ΑΡΙΘΜΗΣΗ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ΜΕΤΡΗΤΕΣ ΑΡΙΘΜΗΣΗΣ ΠΑΡΑΣΤΑΤΙΚΩΝ"
            Height          =   252
            Left            =   6240
            TabIndex        =   94
            Top             =   3840
            Width           =   3612
         End
         Begin VB.Label LABEL1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% Παρακράτηση"
            ForeColor       =   &H00FF0000&
            Height          =   216
            Index           =   1
            Left            =   -72240
            TabIndex        =   93
            Top             =   3624
            Width           =   1308
         End
         Begin VB.Label lblΠΩΛΗΤΕΣ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ΠΩΛΗΤΕΣ"
            ForeColor       =   &H00FF0000&
            Height          =   192
            Left            =   -72000
            TabIndex        =   92
            Top             =   4680
            Width           =   732
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Αυτοκίνητο (προεπιλεγμένο)"
            ForeColor       =   &H00FF0000&
            Height          =   192
            Index           =   19
            Left            =   -74040
            TabIndex        =   91
            Top             =   1560
            Width           =   3012
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MyData Ειδος"
            ForeColor       =   &H00FF0000&
            Height          =   264
            Index           =   22
            Left            =   -74640
            TabIndex        =   90
            Top             =   840
            Width           =   1452
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MyData Τυπος Εσόδου"
            ForeColor       =   &H00FF0000&
            Height          =   264
            Index           =   23
            Left            =   -74640
            TabIndex        =   89
            Top             =   1200
            Width           =   1812
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MyData ΚΑΤΗΓ εσόδου"
            ForeColor       =   &H00FF0000&
            Height          =   192
            Index           =   24
            Left            =   -74640
            TabIndex        =   88
            Top             =   1560
            Width           =   1704
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MyData AYTOTIMOΛΟΓΗΣΗ"
            ForeColor       =   &H00FF0000&
            Height          =   264
            Index           =   25
            Left            =   -74640
            TabIndex        =   87
            Top             =   1920
            Width           =   1932
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   15615
      Begin VB.CommandButton cmdAkyro 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ακυρο"
         Height          =   492
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   240
         Width           =   1572
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Εξοδος"
         Height          =   492
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1572
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Ενημέρωση"
         Height          =   492
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1572
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Eπανυπολογισμός Αποθήκης"
         Enabled         =   0   'False
         Height          =   492
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1572
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Αντιγραφή Παραστατικού"
         Enabled         =   0   'False
         Height          =   492
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1572
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   11280
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox defPar 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   9120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Διόρθωση"
      Height          =   300
      Left            =   10680
      TabIndex        =   1
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Ακύρωση"
      Height          =   300
      Left            =   12000
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   336
      Left            =   4200
      Top             =   9120
      Visible         =   0   'False
      Width           =   1812
      _ExtentX        =   3201
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
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   372
      Left            =   5760
      Top             =   9120
      Visible         =   0   'False
      Width           =   4332
      _ExtentX        =   7646
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   372
      Left            =   840
      Top             =   9120
      Visible         =   0   'False
      Width           =   4332
      _ExtentX        =   7646
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   9120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":3EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":4432
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":4974
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":4EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":5748
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":5C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":5DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":62FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":6850
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":6DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":7334
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu3.frx":78C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   15615
      Begin VB.CommandButton cmdΤαξινόμησηΚατά 
         Caption         =   "Ταξινόμηση κατά Κωδικό Παρ/κού"
         Height          =   315
         Left            =   2640
         TabIndex        =   148
         Top             =   120
         Width           =   2895
      End
      Begin VB.ComboBox combofiltro 
         Height          =   315
         ItemData        =   "bohu3.frx":7E18
         Left            =   120
         List            =   "bohu3.frx":7E25
         TabIndex        =   143
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   9360
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   9360
         TabIndex        =   7
         Top             =   240
         Width           =   6015
      End
      Begin VB.TextBox txtFields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   9360
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         Enabled         =   0   'False
         Height          =   765
         Index           =   9
         Left            =   11520
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   3855
      End
      Begin MSComctlLib.ImageCombo Combo2 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "ΠΑΡΑΣΤΑΤΙΚΑ"
         ImageList       =   "ImageList1"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Χαρ.γράμμα"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   14
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Τίτλος Παραστατικού"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Σειρά (myData)"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Index           =   2
         Left            =   7320
         TabIndex        =   12
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Παρατηρήσεις"
         ForeColor       =   &H00FF8080&
         Height          =   270
         Index           =   9
         Left            =   10320
         TabIndex        =   11
         Top             =   600
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   252
         Left            =   3120
         TabIndex        =   10
         Top             =   60
         Width           =   372
      End
      Begin VB.Label mID2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   252
         Left            =   3720
         TabIndex        =   9
         Top             =   60
         Width           =   732
      End
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Φόρμα 2ης Εκτυύπωσης"
      Height          =   192
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   1716
   End
End
Attribute VB_Name = "bohu3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim f_Timol2_SeiresTimologioy As Integer      '= Val(FINDPARAMETROI(4, "PAR1", "f_TIMOL2_SEIRESTIMOLOGIOY", str(f_MinSeiresTimologioy), "Β. 4 Σειρές τιμ.Πώλησης φορμας TIMOL2"))
Dim f_SynSeiresTimologioy  As Integer '= Val(FINDPARAMETROI(4, "PAR1", "f_SynSeiresTimologioy", "25", "Β. 3 Σειρές Τιμ.Πωλ Κενές+Γραμμ"))


Dim gColor1 As Long
 Dim gColor2 As Long
 Dim gColor3 As Long
 Dim gColor4 As Long
 Dim gDirection    As Direction

Dim fmeArxWidt As Long



Dim F_ALLAGH        As String  ' ΤΟ ΕΙΔΟΣ ΠΟΥ ΑΛΛΑΞΕ ΣΥΜΠΕΡΙΦΟΡΑ ΣΤΗΝ ΑΠΟΘΗΚΗ

Dim mLog1(18)       As String 'για αντιγραφη επικόλληση γεφύρωσης

Dim mLog2(18)       As String 'για αντιγραφη επικόλληση γεφύρωσης

'Dim Adodc1 As Recordset  '
Dim mbChangedByCode As Boolean

Dim mvBookMark      As Variant

Dim mbEditFlag      As Boolean

Dim mbAddNewFlag    As Boolean

Dim mbDataChanged   As Boolean

Dim F(20)           As String

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub ApallFpa_Click()
  DEFAULTAPALL.Text = Split(ApallFpa.Text, ";")(0)
  
End Sub

Private Sub cmdAkyro_Click()
    Combo2.Enabled = True
End Sub

Private Sub CMDFORMA1_Click()

   cd1.ShowOpen
   FORMA1.Text = cd1.FILENAME
   




End Sub

Private Sub CMDFORMA2_Click()
 cd1.ShowOpen
   FORMA2.Text = cd1.FILENAME
   
End Sub

Private Sub CMDFORMA3_Click()
   cd1.ShowOpen
   FORMA3.Text = cd1.FILENAME
   
End Sub

Private Sub cmdΤαξινόμησηΚατά_Click()
 LOAD_LISTA_PAR "", "EIDOS"
End Sub

Private Sub Combo2_Click()


        Command3.Enabled = True
        


        Combo2.Enabled = False
        
        '<EhHeader>
        On Error GoTo Combo2_Click_Err

        '</EhHeader>
        Dim oTEXT, k
        ComboAKYR.Clear
        combometasx.Clear
        
        Dim r3 As New ADODB.Recordset
        r3.Open "select * from PARASTAT WHERE POL=1", Gdb, adOpenDynamic, adLockOptimistic
        Do While Not r3.EOF
            ComboAKYR.AddItem r3("EIDOS") + "-" + r3!TITLOS
            combometasx.AddItem r3("EIDOS") + "-" + r3!TITLOS
            r3.MoveNext
        Loop
        r3.Close
        ComboAKYR.Text = ""
        
    
        
        
        
        
        

100     Adodc1.RecordSource = "SELECT *,ISNULL(ISDIAKIN,0) AS ISDIAKIN2 FROM PARASTAT WHERE ASCII(EIDOS)=" + str(Asc(Right(Combo2.Text, 1)))
110     Adodc1.Refresh


        
'       dum2 = ADD_FIELD("PARASTAT", "MYEID", "nvarCHAR(10)")
'dum2 = ADD_FIELD("PARASTAT", "MYTYPESOD", "nvarCHAR(20)")
'dum2 = ADD_FIELD("PARASTAT", "MYCATEG", "nvarCHAR(20)")
'dum2 = ADD_FIELD("PARASTAT", "MYISAYTOTIM", "INT")
'
   ' MYEID.Text = CNull(Adodc1.Recordset("MYEID"))
   ' myTypEsod.Text = CNull(Adodc1.Recordset("MYTYPESOD"))
   ' mycateg.Text = CNull(Adodc1.Recordset("MYCATEG"))
    MYISAYTOTIM.Value = IIf(nNull(Adodc1.Recordset("MYISAYTOTIM")) = 1, vbChecked, vbUnchecked)
    
    
    ISMYDATA.Value = IIf(nNull(Adodc1.Recordset("ISMYDATA")) = 1, vbChecked, vbUnchecked)
    
    
    MYEID.ListIndex = -1
    For k = 0 To MYEID.ListCount - 1
      If Split(MYEID.List(k), " ")(0) = CNull(Adodc1.Recordset("MYEID")) Then
         MYEID.ListIndex = k
      End If
    Next
    
    
    ComboUBL.ListIndex = -1
    For k = 0 To ComboUBL.ListCount - 1
      If Split(ComboUBL.List(k), " ")(0) = CNull(Adodc1.Recordset("UBL")) Then
         ComboUBL.ListIndex = k
      End If
    Next
    
    
    
    
    
    myTypEsod.ListIndex = -1
    For k = 0 To myTypEsod.ListCount - 1
      If Split(myTypEsod.List(k), " ")(0) = CNull(Adodc1.Recordset("myTypEsod")) Then
         myTypEsod.ListIndex = k
      End If
    Next
    
    mycateg.ListIndex = -1
    For k = 0 To mycateg.ListCount - 1
      If Split(mycateg.List(k), " ")(0) = CNull(Adodc1.Recordset("mycateg")) Then
         mycateg.ListIndex = k
      End If
    Next


    myTypEXod.ListIndex = -1
    For k = 0 To myTypEXod.ListCount - 1
      If Split(myTypEXod.List(k), " ")(0) = CNull(Adodc1.Recordset("myTypEXod")) Then
         myTypEXod.ListIndex = k
      End If
    Next
    
    myCatExod.ListIndex = -1
    For k = 0 To myCatExod.ListCount - 1
      If Split(myCatExod.List(k), " ")(0) = CNull(Adodc1.Recordset("mycateXOD")) Then
         myCatExod.ListIndex = k
      End If
    Next



    
    isDiakin.ListIndex = -1
    For k = 0 To isDiakin.ListCount - 1
      If Left(isDiakin.List(k), 1) = CNull(Adodc1.Recordset("isdiakin2")) Then
         isDiakin.ListIndex = k
      End If
    Next



Checkb2g.Value = IIf(nNull(Adodc1.Recordset("B2G")) = 1, vbChecked, vbUnchecked)

         If IsNull(Adodc1.Recordset("EIDOS")) Then
             Exit Sub
         End If
         cmdUpdate.Enabled = True
         



        Dim R As New ADODB.Recordset

        '   For Each k In bohu3.txtFields

120     For Each oTEXT In Me.txtFields

130         oTEXT.Enabled = True
        Next
        txtFields(0).Enabled = False ' ΧΑΡΑΚΤΗΡΙΣΤΙΚΟ ΓΡΑΜΜΑ ΓΙΑ ΝΑ ΜΗΝ ΤΟ ΧΑΛΑΣΕΙ

140     For k = 0 To 3

            On Error Resume Next

150         If IsNull(Adodc1.Recordset(F(k))) Then
160             txtFields(k).Text = " "
            Else
170             txtFields(k).Text = Adodc1.Recordset(F(k))
            End If

        Next

180     For k = 8 To 16

            On Error Resume Next

190         If IsNull(Adodc1.Recordset(F(k))) Then
200             txtFields(k).Text = " "
            Else
210             txtFields(k).Text = Adodc1.Recordset(F(k))
            End If

        Next

220     If IsNull(Adodc1.Recordset("akyr")) Then
230         akyr.Text = " "
        Else
240         akyr.Text = Adodc1.Recordset("akyr")
        End If

        'AYTOKINHTO.Text = adodc1.Recordset("akyr")

250     If IsNull(Adodc1.Recordset("aytok")) Then
260         AYTOKINHTO.Text = AYTOKINHTO.List(0)
        Else
270         AYTOKINHTO.Text = AYTOKINHTO.List(Adodc1.Recordset("aytok"))
        End If

280     STADIO.Text = STADIO.List(Adodc1.Recordset("STADIO") - 1)



    AJIA_APOU.ListIndex = -1
    For k = 0 To AJIA_APOU.ListCount - 1
      If Left(AJIA_APOU.List(k), 1) = Left(CNull(Adodc1.Recordset("AJIA_APOU")), 1) Then
         AJIA_APOU.ListIndex = k
      End If
    Next
290    ' AJIA_APOU.Text = AJIA_APOU.List(Adodc1.Recordset("AJIA_APOU") - 1) 'AJIA_APOU


    POS_APOU.ListIndex = -1
    For k = 0 To POS_APOU.ListCount - 1
      If Left(POS_APOU.List(k), 1) = Left(CNull(Adodc1.Recordset("POS_APOU")), 1) Then
         POS_APOU.ListIndex = k
      End If
    Next


  PEL.ListIndex = -1
    For k = 0 To PEL.ListCount - 1
      If Left(PEL.List(k), 1) = Left(CNull(Adodc1.Recordset("PEL")), 1) Then
         PEL.ListIndex = k
      End If
    Next
DESMIA.Text = ""
  DESMIA = Adodc1.Recordset("DESMIA")

300   '  POS_APOU.Text = POS_APOU.List(Adodc1.Recordset("POS_APOU") - 1) 'POS_APOU
310   '  PEL.Text = PEL.List(Adodc1.Recordset("PEL") - 1) 'PEL
     
320     log1(1).Text = Adodc1.Recordset("LAJ1")
330     log1(2).Text = Adodc1.Recordset("LAJ2")
340     log1(3).Text = Adodc1.Recordset("LAJ3")
350     log1(4).Text = Adodc1.Recordset("LAJ4")
360     log1(5).Text = Adodc1.Recordset("LAJ5")
    
370     log1(6).Text = Adodc1.Recordset("LAJ6")
380     log1(7).Text = Adodc1.Recordset("LAJ7")
        log1(8).Text = Adodc1.Recordset("LPARAKRAT")
        
        Dim ll As Integer
        
        For ll = 51 To 57
          Label1(ll) = "ΦΠΑ " + Format(g_Fpa(ll - 50), "#0.0") + "%"
        Next
        
        
        
        
        proepilFPA.Text = CNull(Adodc1.Recordset("proepilFPA"))
        
        
        DEFAULTAPALL.Text = CNull(Adodc1.Recordset("DEFAULTAPALL"))
        
        
        
        FORMA1.Text = CNull(Adodc1.Recordset("FORMA1"))
        FORMA2.Text = CNull(Adodc1.Recordset("FORMA2"))
        FORMA3.Text = CNull(Adodc1.Recordset("FORMA3"))
        
        If nNull(Adodc1.Recordset("GEF_M")) = 99 Then
           chKvAT.Value = vbChecked
        Else
            chKvAT.Value = vbUnchecked
        End If
        
        
        
        
        
        
        
        
        
        
        
        mID2.Caption = Adodc1.Recordset("id")
    
        If IsNull(Adodc1.Recordset("PARAKRATISI")) Then
            Parakratisi.Text = ""
        Else
382         Parakratisi.Text = Adodc1.Recordset("PARAKRATISI")
        End If
    
390     log2(1).Text = Adodc1.Recordset("LFPA1")
400     log2(2).Text = Adodc1.Recordset("LFPA2")
410     log2(3).Text = Adodc1.Recordset("LFPA3")
420     log2(4).Text = Adodc1.Recordset("LFPA4")
430     log2(5).Text = Adodc1.Recordset("LFPA5")
    
440     log2(6).Text = Adodc1.Recordset("LFPA6")
450     log2(7).Text = Adodc1.Recordset("LFPA7")
       
460     log2(16).Text = Adodc1.Recordset("LFPA16")
470     log2(17).Text = Adodc1.Recordset("LFPA17")
       
480     log1(0).Text = Adodc1.Recordset("XREPEL")
490     log2(0).Text = Adodc1.Recordset("PISPEL")

typos_par.Text = CNull(Adodc1.Recordset("C1"))

        POLITES.Text = Adodc1.Recordset("POLITES")

        If IsNull(Adodc1.Recordset("yp1")) Then
            YP1.Value = vbUnchecked
        Else

            If Adodc1.Recordset("yp1") = 1 Then
                YP1.Value = vbChecked
            Else
                YP1.Value = vbUnchecked
            End If
        End If

        If IsNull(Adodc1.Recordset("yp2")) Then
            yp2.Value = vbUnchecked
        Else

            If Adodc1.Recordset("yp2") = 1 Then
                yp2.Value = vbChecked
            Else
                yp2.Value = vbUnchecked
            End If
        End If

        If IsNull(Adodc1.Recordset("yp3")) Then
            yp3.Value = vbUnchecked
        Else

            If Adodc1.Recordset("yp3") = 1 Then
                yp3.Value = vbChecked
            Else
                yp3.Value = vbUnchecked
            End If
        End If

        If IsNull(Adodc1.Recordset("MYF")) Then
            myf.ListIndex = 0
        Else
            myf.ListIndex = Adodc1.Recordset("MYF")
        End If

        '.Refresh
        
        
        '             'HEADER
'         Gdb.Execute "UPDATE PARASTAT SET TITLOS='" + txtFields(1) + "' where ID='" + B
          txtFields(1) = Adodc1.Recordset("titlos")
'         Gdb.Execute "UPDATE PARASTAT SET SYNT_TITL='" + txtFields(2) + "' where ID='" + B
          txtFields(2) = Adodc1.Recordset("SYNT_TITL")
'         Gdb.Execute "UPDATE PARASTAT SET EIDOS='" + txtFields(0) + "' where ID='" + B
          txtFields(0) = Adodc1.Recordset("eidos")

'         Gdb.Execute "UPDATE PARASTAT SET SHM1='" + txtFields(9) + "' where ID='" + B
          txtFields(9) = Adodc1.Recordset("SHM1")

'
'        'tab κινησεις-ενημερωσεις
'         Gdb.Execute "UPDATE PARASTAT SET POL=" + RTrim(LTrim(txtFields(3))) + " where ID='" + B + "'"
          txtFields(3) = Adodc1.Recordset("POL")
'         Gdb.Execute "UPDATE PARASTAT SET XONDR=" + RTrim(LTrim(txtFields(15))) + " where ID='" + B + "'"
          txtFields(15) = Adodc1.Recordset("XONDR")
'
'         'TAB Μετασχηματισμοί Ακυρώσεις
'         Gdb.Execute "UPDATE PARASTAT SET METASX='" + RTrim(LTrim(txtFields(8))) + " where ID='" + B + "'"
          txtFields(8) = Adodc1.Recordset("METASX")

'         Gdb.Execute "UPDATE PARASTAT SET GEF_M=" + RTrim(LTrim(txtFields(13))) + " where ID='" + B + "'"
          txtFields(13) = Adodc1.Recordset("gef_m")


          txtFields(0) = Adodc1.Recordset("eidos")
          
'         Gdb.Execute "UPDATE PARASTAT SET STADIO=" + Left$(STADIO.Text, 1) + " where ID='" + B + "'"
          STADIO.Text = Adodc1.Recordset("STADIO")
'         Gdb.Execute "UPDATE PARASTAT SET AKYR='" + Left(akyr.Text, 1) + "' where ID='" + B + "'"
          akyr.Text = Adodc1.Recordset("AKYR")
'
          txtFields(11) = Adodc1.Recordset("EKTYP")
          txtFields(12) = Adodc1.Recordset("FORM_EKTYP")
          txtFields(16) = Adodc1.Recordset("ARITMISI")
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        

500     If Val(txtFields(16).Text) > 0 Then
510         Adodc2.ConnectionString = gConnect
520         Adodc2.RecordSource = "SELECT *FROM ARITMISI WHERE ID=" + txtFields(16).Text
530         Adodc2.Refresh
540         txtFields(17) = nNull(Adodc2.Recordset("aritmisi"))
550         txtFields(17).SetFocus
        End If

560     txtFields(1).SetFocus

Dim r2 As New ADODB.Recordset

610     r2.Open "select *from PARASTAT  WHERE ARITMISI>0   order by ARITMISI,TITLOS ", Gdb, adOpenDynamic, adLockOptimistic
620     r2.MoveFirst
630     k = 0

640     Do While Not r2.EOF
650         ARITMISI.AddItem Format(r2("ARITMISI"), "##") + "---" + " " + r2("TITLOS") + " " + r2("eidos")
       
660         r2.MoveNext
        Loop
r2.Close














        '<EhFooter>
        Exit Sub

Combo2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.Combo2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.Combo2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub ComboAKYR_Click()
  akyr.Text = Left(ComboAKYR.Text, 1)
End Sub



Private Sub combofiltro_Click()

If combofiltro = "" Then
    LOAD_LISTA_PAR "", "POL,TITLOS"
    Exit Sub
End If

If Left(combofiltro.Text, 1) = "1" Then
    LOAD_LISTA_PAR "", "POL,TITLOS"
    Exit Sub
End If


If Left(combofiltro.Text, 1) = "2" Then
    LOAD_LISTA_PAR " WHERE ISMYDATA=1 ", "POL,TITLOS"
    Exit Sub
End If
If Left(combofiltro.Text, 1) = "3" Then
    LOAD_LISTA_PAR " WHERE POL=1 OR POL=2 OR POL=3 ", "POL,TITLOS"
    Exit Sub
End If





End Sub

Private Sub combometasx_Click()
  txtFields(8).Text = txtFields(8).Text + Left(combometasx.Text, 1)
End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     BOHU31.Label1.Caption = txtFields(14).Text

110     BOHU31.SHOW

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     If Len(F_ALLAGH) = 0 Then   'DEN ALLAXE KATI

            Exit Sub

        End If

110     If Left(POS_APOU.Text, 1) = "1" Then
120         Gdb.Execute "UPDATE EGGTIM SET XRE=ABS(POSO),PIS=0 WHERE LEFT(ATIM,1)='" + F_ALLAGH + "'"
        End If

130     If Left(POS_APOU.Text, 1) = "2" Then
140         Gdb.Execute "UPDATE EGGTIM SET XRE=-ABS(POSO),PIS=0 WHERE LEFT(ATIM,1)='" + F_ALLAGH + "'"
        End If

150     If Left(POS_APOU.Text, 1) = "3" Then
160         Gdb.Execute "UPDATE EGGTIM SET PIS=ABS(POSO),XRE=0 WHERE LEFT(ATIM,1)='" + F_ALLAGH + "'"
        End If

170     If Left(POS_APOU.Text, 1) = "4" Then
180         Gdb.Execute "UPDATE EGGTIM SET PIS=-ABS(POSO),XRE=0 WHERE LEFT(ATIM,1)='" + F_ALLAGH + "'"
        End If

190     If Left(POS_APOU.Text, 1) = "0" Then
200         Gdb.Execute "UPDATE EGGTIM SET PIS=0,XRE=0 WHERE LEFT(ATIM,1)='" + F_ALLAGH + "'"
        End If

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()
   txtFields(0).Text = ""
Dim ANS As Integer
ANS = MsgBox("To αντιγραφο θα χρησιμοποιηθει ως ΑΚΥΡΩΤΙΚΟ ;", vbYesNo)


'ψαχνω να βρω ελευθερο γράμμα για παρασταττικό
Dim LETTERS As String
LETTERS = "QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnmερτυθιοπασδφγηξκλζχψωβνμΕΡΤΥΘΙΟΠΑΣΔΦΓΗΞΚΛΖΧΨΩΒΝΜ"
Dim R As New ADODB.Recordset
 R.Open "select EIDOS from PARASTAT order by EIDOS ", Gdb, adOpenDynamic, adLockOptimistic

      
      
      Dim k As Integer
      Dim c As String
      Dim FOUND As Boolean
     
      For k = 1 To Len(LETTERS)
        c = mID$(LETTERS, k, 1)
        
        FOUND = False
        R.MoveFirst
        Do While Not R.EOF
           If R(0) = c Then
              FOUND = True
              Exit Do
           End If
           R.MoveNext
        Loop
        If FOUND = False Then
            txtFields(0).Text = c
            txtFields(1).Text = "*" + txtFields(1).Text
            Exit For
        End If
      Next
      Gdb.Execute "insert into PARASTAT ( EIDOS,TITLOS) VALUES ('" + txtFields(0).Text + "','*" + txtFields(1).Text + "')"
     R.Close
     
     R.Open "SELECT MAX(ID) FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
     mID2 = R(0)
     
     If ANS = vbYes Then
        ISMYDATA.Value = vbUnchecked
        mycateg.ListIndex = 11  'keno
        myTypEsod.ListIndex = 32
        typos_par.Text = "215"
     
     
     End If
      txtFields(16).Text = Val(txtFields(16)) + 20
      
      MsgBox "ΟΚ. ΑΝΤΙΓΡΑΦΗΚΕ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ ,O ΑΡΙΘΜΗΤΗΡΑΣ ΕΓΙΝΕ+20, ΔΙΟΡΘΩΣΤΕ ΚΑΙ ΠΑΤΗΣΤΕ ΕΝΗΜΕΡΩΣΗ"
      
      
      
      
      
      

End Sub

Private Sub defPar_Change()

        Dim R As New ADODB.Recordset

        Dim k

300     R.Open "select *from PARASTAT order by POL,TITLOS ", Gdb, adOpenDynamic, adLockOptimistic

        Dim defparas As String, Ndefparas As Integer

        defparas = ""
        Ndefparas = -1

305     If Len(defPar.Text) = 1 Then
308         defparas = defPar.Text
        End If

        'ΦΟΡΤΩΝΩ TA ΠΑΡΑΣΤΑΤΙΚΑ
310     R.MoveFirst
320     k = 0

330     Do While Not R.EOF

340         If Not IsNull(R("TITLOS")) And Left(R("TITLOS"), 1) <> "-" Then
                If R("eidos") = defparas Then
                    Ndefparas = k + 1
                End If

380             k = k + 1
            End If

390         R.MoveNext
        Loop

        If Ndefparas > -1 Then
            Combo2.Text = Combo2.ComboItems(Ndefparas)
            Combo2_Click
        End If

        txtFields(17).SetFocus
  
End Sub

'Private Sub editform1_Click()
 '   Shell "EXPLORER.EXE " + FORMA1.Text, vbMaximizedFocus
'End Sub

Private Sub ELINES_Click()
'typos_par.Text = Split(ELINES.Text, ";", 1)
 Dim txtLines() As String
         txtLines = Split(ELINES.Text, ";")
         
         typos_par.Text = txtLines(0)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then cmdClose_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()


Dim r3 As New ADODB.Recordset
r3.Open "SELECT COUNT(*), ASCII(EIDOS) FROM PARASTAT GROUP BY ASCII(EIDOS) HAVING COUNT(*)>1 ", Gdb, adOpenDynamic, adLockOptimistic
If Not r3.EOF Then
   If r3(0) > 1 Then
     ' MsgBox "Προσοχή υπάρχει δύο φορές το παραστατικό " + Chr(r3(1))
     ' MsgBox "διορθώστε το στις παραμέτρους παραστατικών"
      ' ΟΤΑΝ ΥΠΑΡΧΕΙ 2 ΦΟΡΕΣ ΤΟ ΙΔΙΟ ΠΑΡΑΣΤΑΤΙΚΟ ΝΑ ΣΒΗΝΕΙ  ΤΟ ΝΕΩΤΕΡΟ
      Gdb.Execute "delete from PARASTAT WHERE ID>0 AND ID=(SELECT TOP 1 ID FROM PARASTAT WHERE  ASCII(EIDOS)='" + str(r3(1)) + "' ORDER BY ID DESC)"
      
     ' Exit Sub
   End If
End If
r3.Close


f_Timol2_SeiresTimologioy = Val(FINDPARAMETROI(4, "PAR1", "f_TIMOL2_SEIRESTIMOLOGIOY", 25, "Β. 4 Σειρές τιμ.Πώλησης φορμας TIMOL2"))
f_SynSeiresTimologioy = Val(FINDPARAMETROI(4, "PAR1", "f_SynSeiresTimologioy", "25", "Β. 3 Σειρές Τιμ.Πωλ Κενές+Γραμμ"))

SEIRES.Caption = "0=" + LTrim(str(f_SynSeiresTimologioy)) + " 2=" + LTrim(str(f_Timol2_SeiresTimologioy))




Gdb.Execute "UPDATE PARASTAT SET FORM_EKTYP=2 ,FORMA1='C:\MERCVB\reports\TIMOL1.RPT',FORMA2='c:\mercvb\reports\TIMOL1.RPT' WHERE FORM_EKTYP=11"
Gdb.Execute "UPDATE PARASTAT SET FORM_EKTYP=2 ,FORMA1='c:\mercvb\reports\TIMOL1.RPT',FORMA2='' WHERE FORM_EKTYP=10"
Gdb.Execute "UPDATE PARASTAT SET FORM_EKTYP=2 ,FORMA1='c:\mercvb\reports\TIMOL1.RPT',FORMA2='c:\mercvb\reports\TIMOL1.RPT',FORMA3='c:\mercvb\reports\TIMOL1.RPT' WHERE FORM_EKTYP=18"

Gdb.Execute "UPDATE PARASTAT SET FORM_EKTYP=2 ,FORMA1='c:\mercvb\reports\TIMOL2.RPT',FORMA2='' WHERE FORM_EKTYP=14"
Gdb.Execute "UPDATE PARASTAT SET FORM_EKTYP=2 ,FORMA1='c:\mercvb\reports\TIMOL2.RPT',FORMA2='c:\mercvb\reports\TIMOL2.RPT' WHERE FORM_EKTYP=12"
Gdb.Execute "UPDATE PARASTAT SET FORM_EKTYP=2 ,FORMA1='c:\mercvb\reports\TIMOL3.RPT',FORMA2='' WHERE FORM_EKTYP=13"


'  2 Με όνομα φόρμας(Εναλλακτικά)
'''''11 TIMOL1 (2εντυπα)
''''10 TIMOL1 (1έντυπο)


'''''18 TIMOL1 (3έντυπα)
'14 TIMOL2 (1εντυπο )
'12 ΤΙΜΟL2 (2εντυπα)

'13 TIMOL3 (1 έντυπο)







  'SSTab1.Width = 0.8 * Me.Width
     Frame1.Left = 0 ' Frame4.Left
     Frame2.Left = 0 ' Frame4.Left
     Frame4.Left = 0
fmeArxWidt = Me.Width

        cmdUpdate.Enabled = False
       

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        Dim o As ComboItem
    
        Dim R As New ADODB.Recordset

        ' mForm_Load Me, fh, fw, ft, fl
100     Me.KeyPreview = True
        'Set db = New Connection
        'db.CursorLocation = adUseClient
        'db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=C:\MERCVB\Emp.mdb;"

110     F_ALLAGH = ""

120     Me.Picture = LoadPicture(gPicture)

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        ' Dim dum2
        'dum2 = ADD_FIELD("PARASTAT", "ID", "INT IDENTITY(1,1)")
        'dum2 = ADD_FIELD("PARASTAT", "YP1", "INT")
        'dum2 = ADD_FIELD("PARASTAT", "YP2", "INT")
        'dum2 = ADD_FIELD("PARASTAT", "YP3", "INT")
        'dum2 = ADD_FIELD("PARASTAT", "YP4", "INT")
        'dum2 = ADD_FIELD("PARASTAT", "YP5", "INT")
        'dum2 = ADD_FIELD("PARASTAT", "C1", "CHAR(20)")
        'dum2 = ADD_FIELD("PARASTAT", "C2", "CHAR(20)")
        '
        ' R.Open "select  COUNT(*) from PARASTAT WHERE YP1=1 ", Gdb, adOpenDynamic, adLockOptimistic
        ' If R(0) = 0 Then
        '    Gdb.Execute "UPDATE PARASTAT SET YP1=1"
        ' End If
        ' R.Close
        '
        ' R.Open "select  COUNT(*) from PARASTAT WHERE YP2=1 ", Gdb, adOpenDynamic, adLockOptimistic
        ' If R(0) = 0 Then
        '    Gdb.Execute "UPDATE PARASTAT SET YP2=1"
        ' End If
        ' R.Close
        '
        '
        '

        '
        'New Recordset
        'Adodc1.Open "select EIDOS,TITLOS,SYNT_TITL,POL,PEL,POS_APOU,AJIA_APOU,STADIO,METASX,SHM1,ETIK,EKTYP,FORM_EKTYP,GEF_M,GEF_P,XONDR,DESMIA from Parastat", db, adOpenStatic, adLockOptimistic

        Dim oTEXT As TextBox

        'Bind the text boxes to the data provider

        'Adodc1.DatabaseName = gDir
130     Adodc1.RecordSource = "SELECT ISNULL(DESMIA,0) AS DESMIA2,*FROM PARASTAT"
140     Adodc1.ConnectionString = gConnect
150     Adodc1.Refresh



        txtFields(0) = ""  'l
        txtFields(1) = ""  'αποδειξη λιανικης
        txtFields(2) = ""  'αποδειξη λιανικης
        
        'Label11.Caption = "Προεπιλ.Κατηγ και Τύπος Φόρου π.χ. 3;6" + Chr(13) + "Για λοιποί φόροι και ξενοδοχ.1-2 αστέρια"

       'ΦΟΡΤΩΜΑ ELINES
160     R.Open "SELECT * FROM PINAKES WHERE TYPOS=42 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
        Dim i As Integer
        i = 0
170     Do While Not R.EOF
            i = i + 1
180         ELINES.AddItem str(R!AYJON) + ";" + R!PERIGRAFH
            
190         R.MoveNext
        Loop
        If i = 0 Then
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,215,'ΕΙΔΙΚΟ ΑΚΥΡΩΤΙΚΟ ΣΤΟΙΧΕΙΟ',-1 )"
         Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,221,'ΔΕΛΤΙΟ ΑΠΟΣΤΟΛ.ΤΙΜΟΛΟΓΙΟ',1)"
          Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,222,'Δ. Αποστολής - Τιμολόγιο',1)"
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,223,'ΤΙΜΟΛΟΓΙΟ (Παροχης Υπηρεσιων) - ΔΕΛΤΙΟ ΑΠΟΣΤΟΛΗΣ',1)"
            Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,226,'ΤΙΜΟΛΟΓΙΟ (Παροχή Υπηρεσιων) - ΤΙΜΟΛΟΓΙΟ (Πώληση Αγαθών)',1)"
             Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,227,'ΤΙΜΟΛΟΓΙΟ(Παροχή Υπηρεσιων)-ΤΙΜΟΛΟΓΙΟ(Πώληση Αγαθών)-ΔΕΛ.ΑΠ.',1)"
              Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,229,'ΔΕΛΤΙΟ ΠΟΣΟΤΙΚΗΣ ΠΑΡΑΛΑΒΗΣ - ΠΙΣΤΩΤΙΚΟ ΤΙΜΟΛΟΓΙΟ',-1)"
               Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,231,'Απόδειξη Λιανικής Πώλησης  ΔΑ',1)"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,232,'ΑΠΟΔΕΙΞΗ ΠΑΡΟΧΗΣ ΥΠΗΡΕΣΙΩΝ - ΔΕΛΤΙΟ ΑΠΟΣΤΟΛΗΣ',1)"
                 Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,40,'ΔΕΛΤΙΟ ΑΠΟΣΤΟΛΗΣ ΑΠΟ ΠΕΛΑΤΗ',0)"
                  Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,158,'Δελτίο Αποστολής Λιανικής',0)"
                   Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,161,'Tιμολόγιο',1)"
                    Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,162,'ΤΙΜΟΛΟΓΙΟ ΠΑΡΟΧ. ΥΠΗΡΕΣΙΩΝ',1)"
                     Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,165,'ΤΙΜΟΛΟΓΙΟ',1)"
                      Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,169,'Πιστωτικό Τιμ.Παροχής Εκπτωσης',-1)"
                       Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,174,'Απόδειξη Παροχής Υπηρεσιών',1)"
                        Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,173,'Απόδειξη Λιανικής Πώλησης',1)"
                        Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,175,'Δελτίο Επιστρ.Λιανικώς Πωληθ.',-1)"
                         Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,233,'ΑΠΟΔΕΙΞΗ ΠΑΡΟΧΗΣ ΥΠΗΡΕΣΙΩΝ - ΑΠΟΔΕΙΞΗ ΛΙΑΝΙΚΗΣ ΠΩΛΗΣΗΣ',1)"
                          Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,235,'ΑΠΟΔΕΙΞΗ ΕΠΙΣΤΡΟΦΗΣ - ΔΕΛΤΙΟ ΑΠΟΣΤΟΛΗΣ',-1)"
        Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,236,'Πιστωτικό Τιμολόγιο(ΔΑ)',-1)"
        Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (42,502,'ΑΝΤΙΤΥΠΟ',0)"
        
        
       
        
        End If
        

'
'
'
'
' fortoma aade
    R.Close

    R.Open "SELECT * FROM PINAKES WHERE TYPOS=43 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
       
        i = 0
     Do While Not R.EOF
            i = i + 1
         mydata.AddItem str(R!AYJON) + ";" + R!PERIGRAFH
         R.MoveNext
        Loop
        If i = 0 Then
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,1,'ΤΙΜΟΛΟΓΙΟ ΠΩΛΗΣΗΣ        ;1.1;E3_561_001;category1_1',0 )"
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,2,'ΤΙΜΟΛΟΓΙΟ ΠΩΛΗΣΗΣ ΕΝΔΟΚ  ;1.2;E3_561_005;category1_1',0 )"
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,3,'ΤΙΜΟΛΟΓΙΟ ΠΩΛΗΣΗΣ ΤΡΙΤΩΝ ;1.3;E3_561_006;category1_1',0 )"
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,4,'ΣΥΜΠΛ.ΤΙΜΟΛΟΓΙΟ ΠΩΛΗΣΗΣ  ;1.6;E3_561_001;category1_1',0 )"
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,5,'ΤΙΜ.ΠΑΡΟΧΗΣ ΥΠΗΡ.        ;2.1;E3_561_001;category1_3',0 )"
           
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,6,'ΤΙΜ.ΠΑΡ.ΥΠΗΡ.ΕΝΔΟΚ       ;2.2;E3_561_005;category1_3',0 )"
           
           
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,7,'ΤΙΜ.ΠΑΡ.ΥΠΗΡ.ΤΡΙΤΗΣ      ;2.3;E3_561_006;category1_3',0 )"
           
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,8,'ΠΙΣΤ.ΤΙΜΟΛΟΓΙΟ ΣΥΣΧ      ;5.1;E3_561_001;category1_1',0 )"
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,9,'ΠΙΣΤ.ΤΙΜΟΛΟΓΙΟ ΜΗ ΣΥΣΧ   ;5.2;E3_561_001;category1_1',0 )"
           
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,10,'ΣΤΟΙΧΕΙΟ ΑΥΤΟΠΑΡΑΔΟΣΗΣ  ;6.1;E3_561_001;category1_1',0 )"
           
           
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,11,'ΑΠΟΔ.ΛΙΑΝ.ΠΩΛ           ;11.1;E3_561_003;category1_1',0 )"
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,12,'ΑΠΟΔ.ΠΑΡ.ΥΠΗΡ           ;11.2;E3_561_003;category1_3',0 )"
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,13,'ΑΠΟΔ.ΕΠΙΣΤ.ΛΙΑΝ         ;11.4;E3_561_003;category1_1',0 )"
           
           
           Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (43,14,'ΕΝΔΟΚ.ΑΠΟΚΤ             ;14.1;E3_102_001;category2_1',0 )"
           
         
           
        End If
        
        
        

    R.Close
    
    
        R.Open "SELECT * FROM PINAKES WHERE TYPOS=44 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
        'Dim i As Integer
        i = 0
        Do While Not R.EOF
            i = i + 1
            ApallFpa.AddItem Format(i, "00") + ";" + str(R!AYJON) + ";" + R!PERIGRAFH
            
            R.MoveNext
        Loop
       R.Close
       
    
    
    
    
    

     R.Open "SELECT *FROM PINAKES WHERE TYPOS=10 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

     Do While Not R.EOF
         AYTOKINHTO.AddItem R("PERIGRAFH")
         R.MoveNext
        Loop

200     If AYTOKINHTO.ListCount = 0 Then
210        ' Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'   ',10)"
        End If

220     R.Close

        ' For Each oText In Me.txtFields

        '   oText.DataSource = ADODC1
        ' Next

230     mbDataChanged = False

        Dim k As Integer

240     For k = 0 To 3
250         F(k) = txtFields(k).DataField
260         txtFields(k).DataField = ""
        Next

270     For k = 8 To 16
280         F(k) = txtFields(k).DataField
290         txtFields(k).DataField = ""
        Next

        load_paras_aritm
        
        
  ' If Option4.Value = True Then
           gDirection = cHorizontal
      '  Else
        '   gDirection = cVertical
       ' End If
        gColor1 = &HFFFFC0  'PRASINO &HC0FFC0    'vbBlue '&HFFFFC0     ' vbBlue
        gColor2 = vbCyan      '&HFF0000
        SetStyle SSTab1.hwnd, cGradient  '//--- Set The Style of The SSTab
        SetGradientDir SSTab1.hwnd, gDirection '//--- Set The Gradient Direction
        SetGradientColor1 SSTab1.hwnd, gColor1  '//--- Asing new Gradient Color Start
        SetGradientColor2 SSTab1.hwnd, gColor2  '//--- Asing new Gradient Color End
 
     SSTabSubclass SSTab1.hwnd '//--- Begin SubClassing
     RedrawWindow SSTab1.hwnd, ByVal 0&, ByVal 0&, &H1

   
     
        
'       dum2 = ADD_FIELD("PARASTAT", "MYEID", "nvarCHAR(10)")
'dum2 = ADD_FIELD("PARASTAT", "MYTYPESOD", "nvarCHAR(20)")
'dum2 = ADD_FIELD("PARASTAT", "MYCATEG", "nvarCHAR(20)")
'dum2 = ADD_FIELD("PARASTAT", "MYISAYTOTIM", "INT")
'
         
        

'
    
        
        
        
        
        
        
        
        
        
        
        
        
        

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub load_paras_aritm()
'
'LOAD_LISTA_PAR ""


If combofiltro = "" Then
    LOAD_LISTA_PAR "", "POL,TITLOS"
    Exit Sub
End If

If Left(combofiltro.Text, 1) = "1" Then
    LOAD_LISTA_PAR "", "POL,TITLOS"
    Exit Sub
End If


If Left(combofiltro.Text, 1) = "2" Then
    LOAD_LISTA_PAR " WHERE ISMYDATA=1 ", "POL,TITLOS"
    Exit Sub
End If
If Left(combofiltro.Text, 1) = "3" Then
    LOAD_LISTA_PAR " WHERE POL=1 OR POL=2 OR POL=3 ", "POL,TITLOS"
    Exit Sub
End If











Dim R As New ADODB.Recordset
Dim k As Integer
'
'Combo2.ComboItems.Clear
'
'ARITMISI.Clear
'
'
'
'
'300     R.Open "select *from PARASTAT order by POL,TITLOS ", Gdb, adOpenDynamic, adLockOptimistic
'        'Dim defparas As String, Ndefparas As Integer
'        'defparas = ""
'        'Ndefparas = -1
'
'        '305 If Len(defPar.Text) = 1 Then
'        '308    defparas = defPar.Text
'        '    End If
'
'        'ΦΟΡΤΩΝΩ TA ΠΑΡΑΣΤΑΤΙΚΑ
'310     R.MoveFirst
'320     k = 0
'
'330     Do While Not R.EOF
'
'340         If Not IsNull(R("TITLOS")) And Left(R("TITLOS"), 1) <> "-" Then
'
'                'If r("eidos") = defparas Then
'                '   Ndefparas = k + 1
'                'End If
'350             If R("POL") = 2 Then  ' αγορες
'360                 Combo2.ComboItems.ADD k + 1, "", Format(k + 1, "##") + "-" + R("eidos") + "  " + R("TITLOS") + Space(1) + R("eidos"), 6
'                    ' Combo2.ComboItems(k).Image = 6
'                Else
'
'370                 Combo2.ComboItems.ADD k + 1, "", Format(k + 1, "##") + "-" + R("eidos") + "  " + R("TITLOS") + Space(1) + R("eidos"), 3
'                    ' Combo2.ComboItems(k).Image = 3
'                End If
'
'380             k = k + 1
'            End If
'
'390         R.MoveNext
'        Loop
'
'        '    Combo2.Text = Combo2.ComboItems(0)
'
'        'If Ndefparas > -1 Then
'        '   Combo2.Text = Combo2.ComboItems(Ndefparas)
'        '   Combo2_Click
'        'End If
'        'ΑΡΙΘΜΗΤΗΡΕΣ ΠΑΡΑΣΤΑΤΙΚΩΝ
'400     R.Close
410     R.Open "select *from PARASTAT  WHERE ARITMISI>0   order by ARITMISI,TITLOS ", Gdb, adOpenDynamic, adLockOptimistic
420     R.MoveFirst
430     k = 0

440     Do While Not R.EOF
450         ARITMISI.AddItem Format(R("ARITMISI"), "##") + "---" + " " + R("TITLOS") + " " + R("eidos")
       
460         R.MoveNext
        Loop
R.Close


End Sub

Sub LOAD_LISTA_PAR(ByVal WHEREQ As String, ORDERBY As String)  '1=POL,TITLOS
'================================

Dim R As New ADODB.Recordset
Dim k As Integer

Combo2.ComboItems.Clear

ARITMISI.Clear




300     R.Open "select *from PARASTAT " + WHEREQ + " order by " + ORDERBY, Gdb, adOpenDynamic, adLockOptimistic
        'Dim defparas As String, Ndefparas As Integer
        'defparas = ""
        'Ndefparas = -1

        '305 If Len(defPar.Text) = 1 Then
        '308    defparas = defPar.Text
        '    End If

        'ΦΟΡΤΩΝΩ TA ΠΑΡΑΣΤΑΤΙΚΑ
310     R.MoveFirst
320     k = 0

330     Do While Not R.EOF

340         If Not IsNull(R("TITLOS")) And Left(R("TITLOS"), 1) <> "-" Then

                'If r("eidos") = defparas Then
                '   Ndefparas = k + 1
                'End If
350             If R("POL") = 2 Then  ' αγορες
360                 Combo2.ComboItems.Add k + 1, "", Format(k + 1, "##") + "-" + R("eidos") + "  " + R("TITLOS") + Space(1) + R("eidos"), 6
                    ' Combo2.ComboItems(k).Image = 6
                Else
             
370                 Combo2.ComboItems.Add k + 1, "", Format(k + 1, "##") + "-" + R("eidos") + "  " + R("TITLOS") + Space(1) + R("eidos"), 3
                    ' Combo2.ComboItems(k).Image = 3
                End If

380             k = k + 1
            End If

390         R.MoveNext
        Loop

        '    Combo2.Text = Combo2.ComboItems(0)
     
        'If Ndefparas > -1 Then
        '   Combo2.Text = Combo2.ComboItems(Ndefparas)
        '   Combo2_Click
        'End If
        'ΑΡΙΘΜΗΤΗΡΕΣ ΠΑΡΑΣΤΑΤΙΚΩΝ
400     R.Close





End Sub



Private Sub Form_Resize()
Dim log As Double
log = Me.Width / fmeArxWidt
     Frame4.Width = Me.Width
     Frame1.Width = Frame4.Width
     
     
     Frame4.Height = Me.Height - Frame4.Top
     Frame2.Width = Frame4.Width
     SSTab1.Height = Frame4.Height
     SSTab1.Width = Me.Width
     
    On Error Resume Next
  'Frame2.Top = SSTab1.Top + SSTab1.Height ' e
   If Me.Height - Frame2.Top > 0 Then
      'Frame2.Height = Me.Height - Frame2.Top
   End If
   
     
    ' mForm_Resize Me, 8, 8580, 9930, fh, fw, ft, fl
End Sub

Private Sub Form_Unload(Cancel As Integer)

        '<EhHeader>
        On Error GoTo Form_Unload_Err

        '</EhHeader>

100     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.Form_Unload " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.Form_Unload " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub cmdAdd_Click()
'    On Error GoTo AddErr
'
'
'    Adodc1.Recordset.AddNew
'    Adodc1.Recordset.Update
'
'
'    Dim k As Integer
'    For k = 0 To 16
'        On Error Resume Next
'        txtFields(k).Text = Adodc1.Recordset(F(k))
'    Next
'
'
'
'
'    Exit Sub
'AddErr:
'    MsgBox Err.Description
'End Sub

'Private Sub cmdDelete_Click()
'    On Error GoTo DeleteErr
'    With adodc1.Recordset
'        .Delete
'        .MoveNext
'        If .EOF Then .MoveLast
'    End With
'    Exit Sub
'DeleteErr:
'    MsgBox Err.Description
'End Sub

Private Sub cmdEdit_Click()

        '<EhHeader>
        On Error GoTo cmdEdit_Click_Err

        '</EhHeader>
        On Error GoTo EditErr

100     mbEditFlag = True
110     SetButtons False

        Exit Sub

EditErr:
120     MsgBox Err.Description

        '<EhFooter>
        Exit Sub

cmdEdit_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.cmdEdit_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.cmdEdit_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdCancel_Click()

        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err

        '</EhHeader>
        On Error Resume Next

100     SetButtons True
110     mbEditFlag = False
120     mbAddNewFlag = False
130     Adodc1.Recordset.CancelUpdate

140     If mvBookMark > 0 Then
150         Adodc1.Recordset.Bookmark = mvBookMark
        Else
160         Adodc1.Recordset.MoveFirst
        End If

170     mbDataChanged = False

        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.cmdCancel_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.cmdCancel_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdUpdate_Click()

Command3.Enabled = False

cmdUpdate.Enabled = False

        '<EhHeader>
        On Error GoTo cmdUpdate_Click_Err

        '</EhHeader>
        On Error GoTo UpdateErr

        Dim k As Integer

        Dim B

        Dim nn As Long

100     If Len(Trim(txtFields(15).Text)) = 0 Then
110         MsgBox "συμπληρωθηκε το πεδίο Χονδρική / Λιανική με Χονδρική"
             txtFields(15).Text = "1"

            'Exit Sub

        End If

120     B = mID2.Caption '  txtFields(0).Text    'EIDOS

130     For k = 0 To 3
140        ' Gdb.Execute "UPDATE PARASTAT SET " + F(k) + "='" + txtFields(k).Text + "' where ID='" + B + "'"
        Next

150     For k = 8 To 16
160        ' Gdb.Execute "UPDATE PARASTAT SET " + F(k) + "='" + txtFields(k).Text + "' where ID='" + B + "'"
        Next

170     Gdb.Execute "UPDATE PARASTAT SET PEL='" + Left(PEL.Text, 1) + "' where ID='" + B + "'"
180     Gdb.Execute "UPDATE PARASTAT SET POS_APOU ='" + Left(POS_APOU.Text, 1) + "' where ID='" + B + "'"
190     Gdb.Execute "UPDATE PARASTAT SET AJIA_APOU='" + Left(AJIA_APOU.Text, 1) + "' where ID='" + B + "'"
200    ' Gdb.Execute "UPDATE PARASTAT SET STADIO='" + Left(STADIO.Text, 1) + "' where ID='" + B + "'"
210
   
220     Gdb.Execute "UPDATE PARASTAT SET LAJ1='" + log1(1) + "' where ID='" + B + "'"
230     Gdb.Execute "UPDATE PARASTAT SET LAJ2='" + log1(2) + "' where ID='" + B + "'"
240     Gdb.Execute "UPDATE PARASTAT SET LAJ3='" + log1(3) + "' where ID='" + B + "'"
250     Gdb.Execute "UPDATE PARASTAT SET LAJ4='" + log1(4) + "' where ID='" + B + "'"
260     Gdb.Execute "UPDATE PARASTAT SET LAJ5='" + log1(5) + "' where ID='" + B + "'"
   
270     Gdb.Execute "UPDATE PARASTAT SET LAJ6='" + log1(6) + "' where ID='" + B + "'"
280     Gdb.Execute "UPDATE PARASTAT SET LAJ7='" + log1(7) + "' where ID='" + B + "'"


Gdb.Execute "UPDATE PARASTAT SET LPARAKRAT='" + log1(8) + "' where ID='" + B + "'"

Gdb.Execute "UPDATE PARASTAT SET ISDIAKIN=" + Left(isDiakin.Text, 1) + " where ID='" + B + "'"

   
        Gdb.Execute "UPDATE PARASTAT SET PARAKRATISI='" + Replace(str(Val(Parakratisi)), ",", ".") + "' where ID='" + B + "'"
        'PARAKRATISI
   
290     Gdb.Execute "UPDATE PARASTAT SET LFPA1='" + log2(1) + "' where ID='" + B + "'"
300     Gdb.Execute "UPDATE PARASTAT SET LFPA2='" + log2(2) + "' where ID='" + B + "'"
310     Gdb.Execute "UPDATE PARASTAT SET LFPA3='" + log2(3) + "' where ID='" + B + "'"
320     Gdb.Execute "UPDATE PARASTAT SET LFPA4='" + log2(4) + "' where ID='" + B + "'"
330     Gdb.Execute "UPDATE PARASTAT SET LFPA5='" + log2(5) + "' where ID='" + B + "'"
        
340     Gdb.Execute "UPDATE PARASTAT SET LFPA6='" + log2(6) + "' where ID='" + B + "'"
350     Gdb.Execute "UPDATE PARASTAT SET LFPA7='" + log2(7) + "' where ID='" + B + "'"
        
360     Gdb.Execute "UPDATE PARASTAT SET LFPA16='" + log2(16) + "' where ID='" + B + "'"
370     Gdb.Execute "UPDATE PARASTAT SET LFPA17='" + log2(17) + "' where ID='" + B + "'"
        
        Gdb.Execute "UPDATE PARASTAT SET POLITES='" + POLITES.Text + "' where ID='" + B + "'"
        
        
         Gdb.Execute "UPDATE PARASTAT SET DESMIA='" + DESMIA.Text + "' where ID='" + B + "'"
        
         Gdb.Execute "UPDATE PARASTAT SET FORMA1='" + FORMA1.Text + "' where ID='" + B + "'"
        
         Gdb.Execute "UPDATE PARASTAT SET FORMA2='" + FORMA2.Text + "' where ID='" + B + "'"
         
         Gdb.Execute "UPDATE PARASTAT SET FORMA3='" + FORMA3.Text + "' where ID='" + B + "'"
         
         
         'ΥΠΟΧΡΕΩΤΙΚΟ ΑΦΜ
       If chKvAT.Value = vbChecked Then
           Gdb.Execute "UPDATE PARASTAT SET GEF_M=99 where ID=" + B + ""
       Else
           Gdb.Execute "UPDATE PARASTAT SET GEF_M=0 where ID='" + B + "'"
       End If
        
        
        
        Gdb.Execute "UPDATE PARASTAT SET C1='" + typos_par.Text + "' where ID='" + B + "'"
        
      '  typos_par.Text = adodc1.Recordset("C1")
        
        
380     Gdb.Execute "UPDATE PARASTAT SET XREPEL='" + log1(0) + "' where ID='" + B + "'"
390     Gdb.Execute "UPDATE PARASTAT SET PISPEL='" + log2(0) + "' where ID='" + B + "'"
        
        If YP1.Value = vbUnchecked Then
            Gdb.Execute "UPDATE PARASTAT SET YP1=0 where ID='" + B + "'"
        Else
            Gdb.Execute "UPDATE PARASTAT SET YP1=1 where ID='" + B + "'"
        End If
        
        If yp2.Value = vbUnchecked Then
            Gdb.Execute "UPDATE PARASTAT SET YP2=0 where ID='" + B + "'"
        Else
            Gdb.Execute "UPDATE PARASTAT SET YP2=1 where ID='" + B + "'"
        End If

        If yp3.Value = vbUnchecked Then
            Gdb.Execute "UPDATE PARASTAT SET YP3=0 where ID='" + B + "'"
        Else
            Gdb.Execute "UPDATE PARASTAT SET YP3=1 where ID='" + B + "'"
        End If
        
        If yp4.Value = vbUnchecked Then
            Gdb.Execute "UPDATE PARASTAT SET YP4=0 where ID='" + B + "'"
        Else
            Gdb.Execute "UPDATE PARASTAT SET YP4=1 where ID='" + B + "'"
        End If
        
        
        If ISMYDATA.Value = vbUnchecked Then
            Gdb.Execute "UPDATE PARASTAT SET ISMYDATA=0 where ID='" + B + "'"
        Else
            Gdb.Execute "UPDATE PARASTAT SET ISMYDATA=1 where ID='" + B + "'"
        End If
        
        
          If Checkb2g.Value = vbUnchecked Then
            Gdb.Execute "UPDATE PARASTAT SET B2G=0 where ID='" + B + "'"
          Else
            Gdb.Execute "UPDATE PARASTAT SET B2G=1 where ID='" + B + "'"
            If Len(Trim(FORMA1.Text)) < 2 Then
                MsgBox ("ΠΡΕΠΕΙ ΝΑ ΥΠΑΡΧΕΙ ΟΝΟΜΑΣΤΙΚΗ ΦΟΡΜΑ Π.Χ.  c:\mercvb\timol1.rpt")
            End If
        End If
        
        
        
    '    Print Split("aa 4567", " ")(1)
    ' 4567
    
    
    
    Dim cvUBL, CV1, CV2, CV3 As String
    
    If Len(ComboUBL.Text) = 0 Then
       cvUBL = ""
    Else
       cvUBL = Split(ComboUBL.Text, " ")(0)
    End If


    If Len(MYEID.Text) = 0 Then
       CV1 = ""
    Else
       CV1 = Split(MYEID.Text, " ")(0)
    End If








    If Len(myTypEsod.Text) = 0 Then
       CV2 = ""
    Else
       CV2 = Split(myTypEsod.Text, " ")(0)
    End If
    
    If Len(mycateg.Text) = 0 Then
       CV3 = ""
    Else
       CV3 = Split(mycateg.Text, " ")(0)
    End If


Dim CV11, CV12, CV13 As String
    If Len(myTypEXod.Text) = 0 Then
       CV12 = ""
    Else
       CV12 = Split(myTypEXod.Text, " ")(0)
    End If
    
    If Len(myCatExod.Text) = 0 Then
       CV13 = ""
    Else
       CV13 = Split(myCatExod.Text, " ")(0)
    End If




'
   ' c2 = IIf(Len(myTypEsod.Text) = 0, "", Split(myTypEsod.Text, " ")(0))
 ''   c3 = IIf(Len(myTypEsod.Text) = 0, "", Split(mycateg.Text, " ")(0))
 
         Gdb.Execute "UPDATE PARASTAT SET UBL='" + cvUBL + "' where ID='" + B + "'" ' = adodc1.Recordset("MYEID")
        
         Gdb.Execute "UPDATE PARASTAT SET MYEID='" + CV1 + "' where ID='" + B + "'" ' = adodc1.Recordset("MYEID")
         Gdb.Execute "UPDATE PARASTAT SET MYTYPESOD='" + CV2 + "' where ID='" + B + "'"
         Gdb.Execute "UPDATE PARASTAT SET MYCATEG='" + CV3 + "' where ID='" + B + "'"
         Gdb.Execute "UPDATE PARASTAT SET MYISAYTOTIM=" + IIf(MYISAYTOTIM.Value = vbChecked, "1", "0") + " where ID='" + B + "'"
    

         Gdb.Execute "UPDATE PARASTAT SET MYTYPEXOD='" + CV12 + "' where ID='" + B + "'"
         Gdb.Execute "UPDATE PARASTAT SET MYCATEXOD='" + CV13 + "' where ID='" + B + "'"
    
         Gdb.Execute "UPDATE PARASTAT SET DEFAULTAPALL=" + IIf(Len(Trim(DEFAULTAPALL.Text)) = 0, "0", DEFAULTAPALL.Text) + " where ID='" + B + "'"
         
         Gdb.Execute "UPDATE PARASTAT SET PROEPILFPA=" + IIf(Len(Trim(proepilFPA.Text)) = 0, "0", proepilFPA.Text) + " where ID='" + B + "'"
  'HEADER
         Gdb.Execute "UPDATE PARASTAT SET TITLOS='" + txtFields(1) + "' where ID=" + B
         Gdb.Execute "UPDATE PARASTAT SET SYNT_TITL='" + txtFields(2) + "' where ID=" + B
         Gdb.Execute "UPDATE PARASTAT SET EIDOS='" + txtFields(0) + "' where ID=" + B
         Gdb.Execute "UPDATE PARASTAT SET SHM1='" + txtFields(9) + "' where ID= " + B
         
        'tab κινησεις-ενημερωσεις
         Gdb.Execute "UPDATE PARASTAT SET POL=" + RTrim(LTrim(str(Val(txtFields(3))))) + " where ID='" + B + "'"
         Gdb.Execute "UPDATE PARASTAT SET XONDR=" + RTrim(LTrim(str(Val(txtFields(15))))) + " where ID='" + B + "'"
          
         'TAB Μετασχηματισμοί Ακυρώσεις
         Gdb.Execute "UPDATE PARASTAT SET METASX='" + RTrim(LTrim(txtFields(8))) + "' where ID='" + B + "'"
         '  Gdb.Execute "UPDATE PARASTAT SET GEF_M=" + RTrim(LTrim(str(Val(txtFields(13))))) + " where ID=" + B
         
         Gdb.Execute "UPDATE PARASTAT SET STADIO=" + IIf(Val(Left(STADIO.Text, 1)) = 0, "0", Left$(STADIO.Text, 1)) + " where ID='" + B + "'"
         Gdb.Execute "UPDATE PARASTAT SET AKYR='" + Left(akyr.Text, 1) + "' where ID='" + B + "'"
        'ΦΟΡΜΕΣ-ΑΡΙΘΜΗΣΗ
        Gdb.Execute "UPDATE PARASTAT SET EKTYP='" + RTrim(LTrim(txtFields(11))) + "' where ID='" + B + "'"
        Gdb.Execute "UPDATE PARASTAT SET FORM_EKTYP='" + RTrim(LTrim(txtFields(12))) + "' where ID='" + B + "'"
        Gdb.Execute "UPDATE PARASTAT SET ARITMISI=" + RTrim(LTrim(str(Val(txtFields(16))))) + " where ID='" + B + "'"
        






  
        
        '    log1(1).Text = adodc1.Recordset("LAJ1")
        '   log1(2).Text = adodc1.Recordset("LAJ2")
        '  log1(3).Text = adodc1.Recordset("LAJ3")
        ' log1(5).Text = adodc1.Recordset("LAJ5")
    
        'log2(1).Text = adodc1.Recordset("LFPA1")
        '   log2(2).Text = adodc1.Recordset("LFPA2")
        '  log2(3).Text = adodc1.Recordset("LFPA3")
        ' log2(5).Text = adodc1.Recordset("LFPA5")

400     If Val(txtFields(16).Text) > 0 Then
410         Gdb.Execute "UPDATE ARITMISI SET ARITMISI=" + txtFields(17).Text + " WHERE ID=" + txtFields(16).Text, nn

420         If nn = 0 Then
430             MsgBox "ΠΡΟΣΟΧΗ Ο ΑΡΙΘΜΗΤΗΡΑΣ ΔΕΝ ΥΠΑΡΧΕΙ,ΔΙΑΛΕΞΤΕ ΜΙΚΡΟΤΕΡΟ"

                Exit Sub

            End If
        
        End If

440     Gdb.Execute "UPDATE PARASTAT SET AYTOK='" + str(AYTOKINHTO.ListIndex) + "' where ID='" + B + "'"
        Gdb.Execute "UPDATE PARASTAT SET TAMEIAKI='" + tameiaki.Text + "' where ID='" + B + "'"

442     Gdb.Execute "UPDATE PARASTAT SET MYF=" + str(myf.ListIndex) + " where ID='" + B + "'"

450     If Adodc1.Recordset("pos_apou") <> Left(POS_APOU.Text, 1) Or Adodc1.Recordset("ajia_apou") <> Left(AJIA_APOU.Text, 1) Then
460         Command2.Enabled = True
470         F_ALLAGH = txtFields(0).Text
        Else
480         F_ALLAGH = ""
        End If

490     For k = 0 To 3
500         txtFields(k).Text = "     "
        Next

510     For k = 8 To 17
520         txtFields(k).Text = "     "
        Next

        '  Form_Load
530     If Command2.Enabled = True Then
            'ΓΙΑ ΝΑ ΠΡΟΛΑΒΩ ΝΑ ΠΑΤΗΣΩ ΕΝΗΜΕΡΩΣΗ
        Else
            ' Unload Me
        End If
        
      
          
          
          
        
        load_paras_aritm  'ενημερωνει τους αριθμητηρες
        
        Combo2.Enabled = True
        Combo2.Text = ""
        

        Exit Sub

UpdateErr:
540     MsgBox Err.Description

550     Resume Next

        '<EhFooter>
        Exit Sub

cmdUpdate_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.cmdUpdate_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.cmdUpdate_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdClose_Click()

        '<EhHeader>
        On Error GoTo cmdClose_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

cmdClose_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.cmdClose_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.cmdClose_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub cmdFirst_Click()
'    On Error GoTo GoFirstError
'
'    adodc1.Recordset.MoveFirst
'    mbDataChanged = False
'
'    Exit Sub
'
'GoFirstError:
'    MsgBox Err.Description
'End Sub

'Private Sub cmdLast_Click()
'    On Error GoTo GoLastError
'
'    adodc1.Recordset.MoveLast
'    mbDataChanged = False
'
'    Exit Sub
'
'GoLastError:
'    MsgBox Err.Description
'End Sub

'Private Sub cmdNext_Click()
'
'    Dim k As Integer
'    Dim b
'
'
'
'    Adodc1.Recordset.MoveNext
'    If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
'    For k = 0 To 16
'        On Error Resume Next
'        txtFields(k).Text = Adodc1.Recordset(F(k))
'    Next
'
'
'
'End Sub

'Private Sub cmdPrevious_Click()
'
'    Dim k As Integer
'    Dim b
'
'
'
'    adodc1.Recordset.MovePrevious
'
'    If adodc1.Recordset.BOF Then adodc1.Recordset.MoveFirst
'
'    For k = 0 To 16
'        On Error Resume Next
'        txtFields(k).Text = adodc1.Recordset(F(k))
'    Next
'
'
'
'End Sub

Private Sub SetButtons(bVal As Boolean)

        ' cmdAdd.Visible = bVal
        '<EhHeader>
        On Error GoTo SetButtons_Err

        '</EhHeader>

100     cmdEdit.Visible = bVal
110     cmdUpdate.Visible = Not bVal
120     cmdCancel.Visible = Not bVal
        '  cmdDelete.Visible = bVal
130     cmdClose.Visible = bVal

        '<EhFooter>
        Exit Sub

SetButtons_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.SetButtons " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.SetButtons " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub



Private Sub foros_Click()
katforou.Clear

Dim N As Integer
N = 0
Dim n2 As Integer
N = (Val(Left(foros.Text, 2)))
If N > 0 And N < 5 Then
    If N = 1 Then n2 = 45
    If N = 2 Then n2 = 54
    If N = 3 Then n2 = 55
    If N = 4 Then n2 = 56
Dim i As Integer

    Dim R As New ADODB.Recordset
        R.Open "SELECT * FROM PINAKES WHERE TYPOS=" + str(n2) + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
        'Dim i As Integer
        i = 0
        Do While Not R.EOF
            i = i + 1
            katforou.AddItem Format(i, "00") + ";" + str(R!AYJON) + ";" + R!PERIGRAFH
            
            R.MoveNext
        Loop
       R.Close


End If



End Sub

Private Sub Image1_Click(index As Integer)
    Shell "EXPLORER.EXE " + Me.Controls("FORMA" + Format(index, "0")).Text, vbMaximizedFocus
End Sub

Private Sub mCopy_Click()

        '<EhHeader>
        On Error GoTo mCopy_Click_Err

        '</EhHeader>
        Dim k As Long

100     For k = 0 To 7
110         mLog1(k) = log1(k).Text
120         mLog2(k) = log2(k).Text
        Next

130     mLog2(16) = log2(16).Text
140     mLog2(17) = log2(17).Text
   
        '<EhFooter>
        Exit Sub

mCopy_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.mCopy_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.mCopy_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mPaste_Click()

        '<EhHeader>
        On Error GoTo mPaste_Click_Err

        '</EhHeader>
        Dim k As Long

100     For k = 0 To 7
110         log1(k).Text = mLog1(k)
120         log2(k).Text = mLog2(k)
        Next

130     log2(16).Text = mLog2(16)
140     log2(17).Text = mLog2(17)

        '<EhFooter>
        Exit Sub

mPaste_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.mPaste_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.mPaste_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mydata_Click()
    Dim txtLines() As String
         txtLines = Split(mydata.Text, ";")
         
         txtFields(10).Text = Left(txtLines(2), 15) + ";" + Left(txtLines(3), 10) + ";" + Left(txtLines(4), 15)

End Sub

Private Sub txtFields_GotFocus(index As Integer)

        '<EhHeader>
        On Error GoTo txtFields_GotFocus_Err

        '</EhHeader>

100     txtFields(index).BackColor = vbYellow

        '<EhFooter>
        Exit Sub

txtFields_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.txtFields_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.txtFields_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtFields_KeyPress(index As Integer, KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtFields_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

txtFields_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.txtFields_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.txtFields_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtFields_LostFocus(index As Integer)

        '<EhHeader>
        On Error GoTo txtFields_LostFocus_Err

        '</EhHeader>

100     txtFields(index).BackColor = vbWhite

         If index = 16 And Val(txtFields(16).Text) > 0 Then
           Adodc2.ConnectionString = gConnect
           Adodc2.RecordSource = "SELECT *FROM ARITMISI WHERE ID=" + txtFields(16).Text
           Adodc2.Refresh
           txtFields(17).Text = Adodc2.Recordset("aritmisi")
           txtFields(17).SetFocus
         End If
        '<EhFooter>
        Exit Sub

txtFields_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu3.txtFields_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu3.txtFields_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

