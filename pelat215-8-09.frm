VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form pelat2 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form2"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15135
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15135
   WindowState     =   2  'Maximized
   Begin VB.ComboBox KATHGORIA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   94
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Ημερ/νία"
      Height          =   345
      Left            =   12090
      TabIndex        =   93
      Top             =   990
      Width           =   570
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   8040
      TabIndex        =   92
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   615
      Left            =   11040
      TabIndex        =   89
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"pelat2.frx":0000
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   10905
      TabIndex        =   88
      Top             =   6015
      Width           =   2760
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Αποστολή Αρχείου"
      Height          =   345
      Left            =   10905
      TabIndex        =   87
      Top             =   5595
      Width           =   2760
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   11085
      Picture         =   "pelat2.frx":008D
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   86
      Top             =   975
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   11625
      Picture         =   "pelat2.frx":06AB
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   85
      Top             =   975
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3675
      TabIndex        =   84
      Top             =   1125
      Width           =   330
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12120
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   4320
      TabIndex        =   49
      Top             =   6000
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5530
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   16776960
      TabCaption(0)   =   "ΠΕΔΙΑ ΗΜΕΡ/ΝΙΑΣ"
      TabPicture(0)   =   "pelat2.frx":0F89
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lhm1"
      Tab(0).Control(1)=   "lhm2"
      Tab(0).Control(2)=   "lhm3"
      Tab(0).Control(3)=   "lhm4"
      Tab(0).Control(4)=   "lhm5"
      Tab(0).Control(5)=   "lhm6"
      Tab(0).Control(6)=   "dtpicker1"
      Tab(0).Control(7)=   "dtpicker2"
      Tab(0).Control(8)=   "dtpicker3"
      Tab(0).Control(9)=   "dtpicker4"
      Tab(0).Control(10)=   "dtpicker5"
      Tab(0).Control(11)=   "dtpicker6"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "ΚΕΙΜΕΝΟΥ"
      TabPicture(1)   =   "pelat2.frx":0FA5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lab1"
      Tab(1).Control(1)=   "lab2"
      Tab(1).Control(2)=   "lab3"
      Tab(1).Control(3)=   "lab4"
      Tab(1).Control(4)=   "lab5"
      Tab(1).Control(5)=   "lab6"
      Tab(1).Control(6)=   "keim1"
      Tab(1).Control(7)=   "keim2"
      Tab(1).Control(8)=   "keim3"
      Tab(1).Control(9)=   "keim4"
      Tab(1).Control(10)=   "keim5"
      Tab(1).Control(11)=   "keim6"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "ΑΡΙΘΜΗΤΙΚΑ "
      TabPicture(2)   =   "pelat2.frx":0FC1
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Lnum4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Lnum3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Lnum2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Lnum1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Check1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text5"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text6"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text7"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text8"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.TextBox Text8 
         DataField       =   "NUM1"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   3210
         TabIndex        =   79
         Top             =   1095
         Width           =   2000
      End
      Begin VB.TextBox Text7 
         DataField       =   "NUM2"
         DataSource      =   "DATA1"
         Height          =   315
         Left            =   3210
         TabIndex        =   78
         Top             =   1530
         Width           =   2000
      End
      Begin VB.TextBox Text6 
         DataField       =   "NUM3"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   3210
         TabIndex        =   77
         Top             =   1980
         Width           =   2000
      End
      Begin VB.TextBox Text5 
         DataField       =   "NUM4"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   3210
         TabIndex        =   76
         Top             =   2460
         Width           =   2000
      End
      Begin VB.TextBox dtpicker6 
         DataField       =   "HM6"
         DataSource      =   "DATA1"
         Height          =   300
         Left            =   -71475
         TabIndex        =   74
         Top             =   2775
         Width           =   2000
      End
      Begin VB.TextBox dtpicker5 
         DataField       =   "HM5"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71475
         TabIndex        =   73
         Top             =   2385
         Width           =   2000
      End
      Begin VB.TextBox dtpicker4 
         DataField       =   "HM4"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71475
         TabIndex        =   72
         Top             =   1965
         Width           =   2000
      End
      Begin VB.TextBox dtpicker3 
         DataField       =   "HM3"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   -71475
         TabIndex        =   71
         Top             =   1485
         Width           =   2000
      End
      Begin VB.TextBox dtpicker2 
         DataField       =   "HM2"
         DataSource      =   "DATA1"
         Height          =   315
         Left            =   -71475
         TabIndex        =   70
         Top             =   1035
         Width           =   2000
      End
      Begin VB.TextBox dtpicker1 
         DataField       =   "HM1"
         DataSource      =   "DATA1"
         Height          =   330
         Left            =   -71475
         TabIndex        =   69
         Top             =   600
         Width           =   2000
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ενεργός"
         DataField       =   "ENERGOS"
         DataSource      =   "DATA1"
         Height          =   255
         Left            =   2160
         TabIndex        =   68
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox keim6 
         DataField       =   "CH6"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   62
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox keim5 
         DataField       =   "CH5"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   61
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox keim4 
         DataField       =   "CH4"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   60
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox keim3 
         DataField       =   "CH3"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   59
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox keim2 
         DataField       =   "CH2"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   58
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox keim1 
         DataField       =   "CH1"
         DataSource      =   "DATA1"
         Height          =   285
         Left            =   -71880
         TabIndex        =   57
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Lnum1 
         Caption         =   "Aριθμ1"
         Height          =   255
         Left            =   165
         TabIndex        =   83
         Top             =   1095
         Width           =   2055
      End
      Begin VB.Label Lnum2 
         Height          =   255
         Left            =   165
         TabIndex        =   82
         Top             =   1605
         Width           =   2055
      End
      Begin VB.Label Lnum3 
         Height          =   255
         Left            =   165
         TabIndex        =   81
         Top             =   2070
         Width           =   2055
      End
      Begin VB.Label Lnum4 
         Height          =   255
         Left            =   165
         TabIndex        =   80
         Top             =   2490
         Width           =   2055
      End
      Begin VB.Label lab6 
         Height          =   255
         Left            =   -74640
         TabIndex        =   67
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label lab5 
         Height          =   255
         Left            =   -74640
         TabIndex        =   66
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label lab4 
         Height          =   255
         Left            =   -74640
         TabIndex        =   65
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label lab3 
         Height          =   255
         Left            =   -74640
         TabIndex        =   64
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lab2 
         Height          =   255
         Left            =   -74640
         TabIndex        =   63
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lab1 
         Caption         =   "lab1 - keim1"
         Height          =   255
         Left            =   -74640
         TabIndex        =   56
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lhm6 
         Height          =   255
         Left            =   -74520
         TabIndex        =   55
         Top             =   2805
         Width           =   2055
      End
      Begin VB.Label lhm5 
         Height          =   255
         Left            =   -74520
         TabIndex        =   54
         Top             =   2430
         Width           =   2055
      End
      Begin VB.Label lhm4 
         Height          =   255
         Left            =   -74520
         TabIndex        =   53
         Top             =   1995
         Width           =   2055
      End
      Begin VB.Label lhm3 
         Height          =   255
         Left            =   -74520
         TabIndex        =   52
         Top             =   1575
         Width           =   2055
      End
      Begin VB.Label lhm2 
         Height          =   255
         Left            =   -74520
         TabIndex        =   51
         Top             =   1110
         Width           =   2055
      End
      Begin VB.Label lhm1 
         Caption         =   "Lhm1"
         Height          =   255
         Left            =   -74520
         TabIndex        =   50
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ετικεττα"
      Height          =   375
      Left            =   6585
      TabIndex        =   48
      Top             =   5490
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Καρτέλλα"
      Height          =   375
      Left            =   4650
      TabIndex        =   47
      Top             =   5505
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "XRVMA"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   3675
      TabIndex        =   4
      Top             =   1560
      Width           =   840
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   41
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataSource      =   "DATA1"
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
      Left            =   1200
      TabIndex        =   40
      Top             =   6840
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "pelat2.frx":0FDD
      Left            =   1200
      List            =   "pelat2.frx":0FED
      TabIndex        =   39
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   3720
      TabIndex        =   38
      ToolTipText     =   "Aλλαγή κωδικού"
      Top             =   120
      Width           =   135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "pelat2.frx":102B
      Height          =   3450
      Left            =   4650
      TabIndex        =   37
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
      Left            =   480
      Top             =   8400
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   1140
      TabIndex        =   34
      Top             =   5400
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   1140
      TabIndex        =   33
      Top             =   5880
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   1140
      TabIndex        =   11
      Top             =   4920
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   1140
      TabIndex        =   10
      Top             =   4440
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   1140
      TabIndex        =   9
      Top             =   3960
      Width           =   2200
   End
   Begin VB.CommandButton Command3 
      Height          =   400
      Left            =   4650
      TabIndex        =   12
      Top             =   4935
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εξοδος"
      Height          =   400
      Left            =   8610
      TabIndex        =   13
      Top             =   4935
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Διαγραφή"
      Height          =   400
      Left            =   6570
      TabIndex        =   29
      Top             =   4935
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "kod"
      DataSource      =   "Data1"
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   1140
      TabIndex        =   20
      Top             =   120
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1140
      TabIndex        =   1
      Top             =   600
      Width           =   3285
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1140
      TabIndex        =   2
      Top             =   1080
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   1140
      TabIndex        =   3
      Top             =   1560
      Width           =   1680
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   1140
      TabIndex        =   5
      Top             =   2040
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   1140
      TabIndex        =   6
      Top             =   2520
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   1140
      TabIndex        =   7
      Top             =   3000
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   1140
      TabIndex        =   8
      Top             =   3480
      Width           =   2200
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   16
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   14
      Top             =   0
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.ComboBox PELPROM 
         Height          =   315
         ItemData        =   "pelat2.frx":103F
         Left            =   0
         List            =   "pelat2.frx":1049
         TabIndex        =   90
         Top             =   0
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   1320
      Top             =   8160
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
   Begin RichTextLib.RichTextBox TEXT3 
      DataField       =   "MEMO"
      DataSource      =   "DATA1"
      Height          =   3450
      Left            =   11070
      TabIndex        =   75
      Top             =   1425
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   6085
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"pelat2.frx":1067
   End
   Begin VB.Label ffbuff 
      Height          =   195
      Left            =   8430
      TabIndex        =   91
      Top             =   1155
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
      TabIndex        =   46
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "T.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      TabIndex        =   45
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
      TabIndex        =   44
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
      TabIndex        =   43
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
      TabIndex        =   42
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "pelat2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FappWord As Word.Application
Dim FWRDdoc As Word.Document


Dim f_TIMOKAT As Integer

Dim F_LHM1 As String
Dim F_LHM2 As String
Dim f_lHM3 As String
Dim f_lHM4 As String
Dim f_lHM5 As String
Dim f_lHM6 As String

Dim f_lab1 As String
Dim f_lab2 As String
Dim f_lab3 As String
Dim f_lab4 As String
Dim f_lab5 As String
Dim f_lab6 As String

Dim f_num1 As String
Dim f_num2 As String
Dim f_num3 As String
Dim f_num4 As String

Dim f_New_Record As Boolean
Dim f_FirstTab As Integer

Dim PELPROM_STRING As String

Dim f_files  As String ' arxeia office

   Dim FSELSTART
   Dim FSELLENGTH


Dim PROTH

Private Sub SHOW_FIELDS()

Text1(0).Text = Data1.Recordset("KOD")
Text1(1).Text = Data1.Recordset("EPO")
Text1(2).Text = Data1.Recordset("DIE")
Text1(3).Text = Data1.Recordset("POL")
Text1(13).Text = Data1.Recordset("XRVMA") '
Text1(4).Text = Data1.Recordset("EPA")
Text1(5).Text = Data1.Recordset("AFM")
Text1(6).Text = Data1.Recordset("THL")
Text1(7).Text = Data1.Recordset("DOY")
Text1(8).Text = Data1.Recordset("PEK")
Text1(9).Text = Data1.Recordset("PLAFON")
Text1(10).Text = Data1.Recordset("SHM2")
Text1(11).Text = Data1.Recordset("AYP")
Text1(12).Text = Data1.Recordset("TYP")

Text1(14).Text = Data1.Recordset("PLAISIO")


Dim n As Integer
For n = 0 To Combo1.ListCount - 1
   If Val(Left(Combo1.List(n), 3)) = Val(Data1.Recordset("TYPOS")) Then
      Combo1.Text = Combo1.List(n) ' KATHESTOS  FPA
      Exit For
   End If
Next

For n = 0 To Combo2.ListCount - 1
   If Val(Left(Combo2.List(n), 3)) = Val(Data1.Recordset("ARPARAG")) Then
      Combo2.Text = Combo2.List(n) 'TROPOS PLIROMIS
      Exit For
   End If
Next


For n = 0 To KATHGORIA.ListCount - 1
   If Val(Left(Combo2.List(n), 3)) = Val(Data1.Recordset("KODGAL")) Then
      Combo2.Text = Combo2.List(n) 'TROPOS PLIROMIS
      Exit For
   End If
Next






' KATHGORIA.Text = Data1.Recordset("KODGAL")
'KATHGORIA
















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
  If Text1(1).Enabled Then 'ΑΝ ΕΙΣΑΙ ΣΤΗΝ ΔΙΟΡΘΩΣΗ
     'On Error Resume Next
     Data1.Recordset.UpdateBatch
    
      
     
  End If
  Unload Me
  
  
If gApoMenu = True Then
  SendKeys "%"
  SendKeys "~" 'ENTER
  SendKeys "{DOWN}"
Else
   Unload pelat2
   
End If

End Sub

Private Sub Command2_Click()

 Dim ANS As Integer
Dim neos As String, RECS As Integer
Dim BUFF As String
BUFF = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)
'  neos = InputBox("Δώσε τον νέο κωδικό ")
  Dim R As New ADODB.Recordset
  
  R.Open "select count(*) from EGG WHERE KOD='" + Data1.Recordset("KOD") + "' and EIDOS='" + BUFF + "'", Gdb, adOpenForwardOnly, adLockReadOnly
  If R(0) > 0 Then
      MsgBox "Αδύνατη η διαγραφή γιατί υπάρχουν κινήσεις"
      R.Close
      Exit Sub
  End If
 
On Error Resume Next
ANS = MsgBox("Προσοχή θα διαγραφεί ο λογαριασμός " + Data1.Recordset("epo"), vbYesNo)
    If ANS = vbYes Then
       Data1.Recordset.Delete
       
                  
       MHNYMA2.Timer1.Interval = 2000
       MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " διεγράφη."
       MHNYMA2.Label1.Caption = "Διαγραφή εγγραφής"
       MHNYMA2.SHOW 1

       
       
       
       'MsgBox "Διεγράφη"
    End If
End Sub

Private Sub Command3_Click()
Dim k As Integer

On Error GoTo ISLOCKED

If Command3.Caption = "Διόρθωση" Then
       Command3.Caption = "Ενημέρωση"
       For k = 1 To 13
          Text1(k).Enabled = True
          Text1(k).ForeColor = vbBlack
       '   Text1(k).MaxLength = Data1.Recordset(Text1(k).DataField).DefinedSize
       Next
       Text1(1).SetFocus
       Text3.Enabled = True
       Command7.Enabled = True
Else
       Command3.Caption = "Διόρθωση"
     Data1.Recordset.MoveFirst
     
     '  Data1.Recordset.UpdateBatch adAffectAll
     
     
       Dim sql As String
       
       sql = "UPDATE PEL SET "    ' MEMO='" + Text3.Text + "'"
       sql = sql + "EPO='" + Replace(Text1(1).Text, "'", "~") + "'"
       sql = sql + ",TYPOS='" + Left(Combo1.Text, 2) + "'"
       sql = sql + ",DIE='" + Replace(Text1(2).Text, "'", "~") + "'"
       sql = sql + ",POL='" + Replace(Text1(3).Text, "'", "~") + "'"
       sql = sql + ",XRVMA='" + Replace(Text1(13).Text, "'", "~") + "'"
       sql = sql + ",EPA='" + Replace(Text1(4).Text, "'", "~") + "'"
       sql = sql + ",AFM='" + Replace(Text1(5).Text, "'", "~") + "'"
       sql = sql + ",THL='" + Replace(Text1(6).Text, "'", "~") + "'"
       sql = sql + ",DOY='" + Replace(Text1(7).Text, "'", "~") + "'"
       sql = sql + ",PEK='" + Replace(Text1(8).Text, "'", "~") + "'"
       sql = sql + ",PLAFON='" + Replace(Text1(9).Text, "'", "~") + "'"
       sql = sql + ",SHM2='" + Replace(Text1(10).Text, "'", "~") + "'"
       
       sql = sql + ",AYP=" + Replace(Text1(11).Text, "'", "~") + ""
       sql = sql + ",PLAISIO='" + Replace(Text1(12).Text, "'", "~") + "'"
       sql = sql + ",ARPARAG='" + Left(Combo2.Text, 2) + "'"
       sql = sql + ",KODGAL='" + Left(KATHGORIA.Text, 2) + "'"
       
       sql = sql + " WHERE KOD='" + Text1(0) + "' AND EIDOS='" + mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1) + "'"
       
       Gdb.Execute sql
      ' Data1.Refresh
       
       
       
       MHNYMA2.Timer1.Interval = 2000
       MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " αποθηκεύτηκε με επιτυχία"
       MHNYMA2.SHOW 1
       Text2(0).SetFocus
End If


Exit Sub

ISLOCKED:
   MsgBox "Η εγγραφή ενημερώνεται από άλλο χρήστη"
   Resume Next
   Exit Sub

End Sub

Private Sub Command4_Click()
Dim neos As String, RECS As Integer
Dim BUFF As String
BUFF = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)
'BUFF = IIf(PELPROM.ListIndex = 0, "e", "r")
  neos = InputBox("Δώσε τον νέο κωδικό ")
  Dim R As New ADODB.Recordset
  R.Open "select count(*) from PEL WHERE KOD='" + neos + "' and EIDOS='" + BUFF + "'", Gdb, adOpenForwardOnly, adLockReadOnly
  If R(0) = 0 Then
     
     
     Gdb.Execute "UPDATE PEL SET KOD='" + neos + "' WHERE EIDOS='" + BUFF + "' AND KOD='" + Text1(0).Text + "'", RECS
     MsgBox "ΑΡΧΕΙΟ ΠΕΛΑΤΩΝ " + Chr(13) + "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")
  
     Gdb.Execute "UPDATE EGG SET KOD='" + neos + "' WHERE EIDOS='" + BUFF + "' AND KOD='" + Text1(0).Text + "'", RECS
     MsgBox "ΑΡΧΕΙΟ ΚΙΝΗΣΕΩΝ ΑΠΟΘΗΚΗΣ " + Chr(13) + "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")
  
  
  
     'Gdb.Execute "UPDATE TIM SET PELKOD='" + neos + "' WHERE EIDOS='" + BUFF + "' AND PELKOD='" + Text1(0).Text + "'", RECS
     MsgBox "ΑΡΧΕΙΟ TIMOΛΟΓΙΩΝ " + Chr(13) + "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")
  
     Gdb.Execute "UPDATE TIM SET KPE='" + neos + "' WHERE EIDOS='" + BUFF + "' AND KPE='" + Text1(0).Text + "'", RECS
     
  MsgBox "ΑΡΧΕΙΟ TIMOΛΟΓΙΩΝ " + Chr(13) + "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")
  
     Gdb.Execute "UPDATE EGGTIM SET PELKOD='" + neos + "' WHERE EIDOS='" + BUFF + "' AND PELKOD='" + Text1(0).Text + "'", RECS
     MsgBox "ΑΡΧΕΙΟ ΚΙΝΗΣΕΩΝ ΑΠΟΘΗΚΗΣ " + Chr(13) + "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")
  
  
  
  
  Else
     MsgBox "Υπάρχει ο κωδικός " + neos
  End If
  End Sub


Private Sub Command5_Click()
  Dim fff As New pelat4
  ' fff.PELPROM.Text = PELPROM.Text
  fff.ffbuff = IIf(Left(PELPROM.Text, 1) = "1", "e", "r") + Text1(0).Text
  'fff.ffbuff = IIf(Left(PELPROM.Text, 1) = "1", "e", "r") + Text1(0).Text
  fff.Command7.Enabled = False
  fff.APOALLO = "1"
  
  fff.SHOW
  
  Exit Sub
  
  
  
  
  
  
  
  
  pelat4.PELPROM.Text = PELPROM.Text
  pelat4.SHOW
  
  
 ' pelat4.TEXT2(0).Text = Text1(0)
  
  pelat4.APOALLO = "1"
  pelat4.Data1.ConnectionString = gConnect
  
  pelat4.Text2(0).Text = Text1(0)
    
'  pelat4.Data1.Refresh
  
 SendKeys "~" 'ENTER
  SendKeys "~" 'enter  {DOWN}"
 
  
End Sub

Private Sub Command6_Click()
' ETIKETTA ME KEIMENO

Dim ll, apoSCANNER As Long
Dim db
Dim PSIFIA As Integer
Dim fores As Long, k As Integer
Dim X As Printer
Dim barc(200)
Dim mONOMA, mBarcode, mERG, mLTI5
Dim R As New ADODB.Recordset
Dim RD As DAO.Recordset



On Error GoTo 0

mONOMA = Trim(Text1(1).Text) 'ONOMA
mBarcode = Text1(0).Text '  ' BARCODE
mERG = Text1(6).Text ' erg
mLTI5 = Format(Val(Text1(13).Text), "###0.00") ' LTI5


On Error Resume Next
For Each X In Printers
    If InStr(X.DeviceName, "SATO") > 0 Then
        ' Set printer as system default.
        Set Printer = X
        ' Stop looking for a printer  5200016000239
        
        Exit For
    End If
Next

apoSCANNER = 1
'If Scanner Then

 ' μπουγουδης
'Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
'Set RD = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM= 1234")
'If RD.RecordCount = 0 Then
'   MsgBox "δεν υπάρχουν εγγραφές"
'   Exit Sub
'End If

  
'  RD.MoveFirst
  k = 0
  apoSCANNER = k
' End If


fores = Val(InputBox("Πόσες εττικέτες; ", "", 1))
Dim mFORES As Long
For ll = 1 To fores

mFORES = fores

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


Dim Z As Integer
Z = 1 ' InputBox("METATOPISH")

For k = 1 To mFORES



Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992

If Len(mONOMA) > 16 Then
    Printer.FONTSIZE = 14
Else
    Printer.FONTSIZE = 16
End If
Printer.CurrentY = 0

Printer.FontBold = True
Printer.Print Tab(4 + Z); Left(mONOMA, 16) 'Text1(1).Text 'ONOMA
'Printer.Print Tab(4 + Z); Mid$(mONOMA, 17, 16) 'Text1(1).Text 'ONOMA
'Printer.Print Tab(4 + Z); Mid$(mONOMA, 33, 16) 'Text1(1).Text 'ONOMA
Printer.FONTSIZE = 12
Printer.Print Tab(4 + Z); Text1(2).Text;   ' DIE
Printer.Print Tab(4 + Z); Text1(3).Text;   ' POL
Printer.Print Tab(4 + Z); "THΛ."; Text1(6).Text; ' THL





Printer.EndDoc
Next k




Next ll








End Sub





Private Sub Command7_Click()
Dim C As String
C = get_next_pel("e")
       
Gdb.Execute "insert into PEL (EIDOS,KOD) VALUES ('e','" + C + "')"

Text2(0).Text = C
Text1(0).Text = C
Text2(0).SetFocus
 
 Data1.RecordSource = "select *from PEL where  KOD='" + C + "'  AND EIDOS='e' ;"
Data1.Refresh

       
       Command7.Enabled = False
'       f_New_Record = True
'
'
'       For K = 0 To 11
'          Text1(K).Enabled = True
'          Text1(K).ForeColor = vbBlack
'       Next
   
   
   
   Command3.Caption = "Διόρθωση"
   'Command3_Click
   
   DataGrid1.SetFocus
   SendKeys "{enter}{enter}"
       
   
   
End Sub

Private Sub Command8_Click()
  Dim mF As String
  
  'ΑΝ ΔΕΝ ΔΙΑΛΕΞΑ ΠΕΛΑΤΗ ΝΑ ΦΕΥΓΕΙ
  If Len(Trim(Text1(0))) = 0 Then
     Exit Sub
  End If
  
  '\\Server\data
  
  
  
  
  CommonDialog1.ShowOpen
  'If PELPROM.ListIndex = 0 Then
     mF = f_files + "\MERCVB\files\" + mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1) + Text1(0)
 ' Else
  '   mF = f_files + "\MERCVB\files\r" + Text1(0)
 ' End If
  If Len(Dir(f_files + "\MERCVB\files", vbDirectory)) = 0 Then
      MkDir f_files + "\MERCVB\files"
  End If
  
  If Len(Dir(mF, vbDirectory)) = 0 Then
      MkDir mF
  End If
  
  Dim mf2 As Integer, files2 As String
  
  On Error Resume Next
  
  mf2 = InStrRev(CommonDialog1.FILENAME, "\")
  files2 = mID(CommonDialog1.FILENAME, mf2 + 1, Len(CommonDialog1.FILENAME) - mf2)
     
  FileCopy CommonDialog1.FILENAME, mF + "\" + files2
  
  
  Dim OLDPATH As String
  OLDPATH = File1.Path
  
  File1.Visible = False
  File1.Path = "C:\MERCVB\PEL"
  
  File1.Refresh
  File1.Path = OLDPATH
  DoEvents
  File1.Visible = True
  
  
  File1.Refresh
  
  
  
End Sub

Private Sub Command9_Click()
Text3.SetFocus
SendKeys Format(Now, "dd/MM/YYYY")
Text3.SetFocus
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
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0 'suppress the beep
      'keybd_event VK_TAB, 0, 0, 0 'send a tab
      Command3.SetFocus
      
   End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)


Dim mF As String

  'If Option1(1).value = True Then
 On Error Resume Next
 Dim n As Integer
 n = PELPROM.ListIndex
 If n < 0 Then n = 0
 mF = UCase(f_files + "\MERCVB\files\" + mID(PELPROM_STRING, n + 1, 1) + Text1(0))
  'Else
  '   mF = UCase(f_files + "\MERCVB\files\r" + Text1(0))
  'End If
On Error Resume Next

If Len(Dir(mF, vbDirectory)) > 0 Then
   File1.Path = mF
Else
   File1.Path = "c:\mercvb\files\dum"
End If

   
   File1.Refresh


SHOW_FIELDS




End Sub

Private Sub File1_Click()
   Dim sFilename As String
'
'
' sFileName = File1.Path + "\" + File1.FILENAME
sFilename = File1.Path + "\" + File1.FILENAME
If UCase(Right(sFilename, 3)) = "DOC" Then
   Set FappWord = New Word.Application
   Set FWRDdoc = FappWord.Documents.Open(sFilename)
   FappWord.Visible = True
End If




End Sub

Private Sub Form_DblClick()

 PARAMETROI.PARAM.Caption = "PELAT2"
  PARAMETROI.SHOW 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then Command1_Click

If KeyCode = 121 Then
     ut_rantebou.Label1 = Text1(1).Text
     
     ut_rantebou.SHOW 1
  End If
End Sub

Private Sub Form_Load()
  Me.KeyPreview = True
  
     'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
   ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
    'SkinFramework1.ApplyWindow Me.hWnd
    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

  
 'f_checkPos = Val(FindParametroi("PAR1", "F_CHECKPOS", "DELETE", "1=ΕΛΕΓΧOΣ 2=ΑΠΑΓΟΡΕΥΣΗ 0=ΤΙΠΟΤΑ"))
  
 PELPROM_STRING = "er"
 PELPROM.ListIndex = 0
 
 f_TIMOKAT = Val(FindParametroi("PELAT2", "F_TIMOKAT", "0", "Aριθμός Τιμοκαταλόγων"))
 
 
 
 
 
 Dim DUM
 
 DUM = FindParametroi("PELAT2", "F_KODPROM", "DELETE", "")
 DUM = FindParametroi("PELAT2", "F_RAFI", "DELETE", "")
 DUM = FindParametroi("PELAT2", "F_OIKO", "DELETE", "")
 DUM = FindParametroi("PELAT2", "F_YPOOIK", "DELETE", "")
 DUM = FindParametroi("PELAT2", "F_XTI", "DELETE", "")
 
 DUM = FindParametroi("PELAT2", "F_BOHU", "DELETE", "")
 DUM = FindParametroi("PELAT2", "F_LTI", "DELETE", "")
 DUM = FindParametroi("PELAT2", "F_PROM", "DELETE", "")
 DUM = FindParametroi("PELAT2", "F_STOK", "DELETE", "")
 
 
 
 
 
 '\\Server\data
 
 
 f_FirstTab = FindParametroi("PELAT2", "FirstTab", "1", "Πρoεπιλεγμένο Τab ")
 
 SSTab1.Tab = IIf(f_FirstTab > SSTab1.Tabs - 1 Or f_FirstTab < 0, SSTab1.Tabs - 1, f_FirstTab)
 
 
 
 f_files = Trim(FindParametroi("PELAT2", "F_FILES", "C:", "Δίσκος με αρχεία office"))
 
 F_LHM1 = FindParametroi("PELAT2", "F_lhm1", "Ημερ1", "Ετικέτα ημερ.1")
 F_LHM2 = FindParametroi("PELAT2", "F_lhm2", "Ημερ2", "Ετικέτα ημερ.2")
 f_lHM3 = FindParametroi("PELAT2", "F_lhm3", "Ημερ3", "Ετικέτα ημερ.3")
 f_lHM4 = FindParametroi("PELAT2", "F_lhm4", "Ημερ4", "Ετικέτα ημερ.4")
 f_lHM5 = FindParametroi("PELAT2", "F_lhm5", "Ημερ5", "Ετικέτα ημερ.5")
 f_lHM6 = FindParametroi("PELAT2", "F_lhm6", "Ημερ6", "Ετικέτα ημερ.6")
 
 
 
 
 lhm1.Caption = F_LHM1
 lhm2.Caption = F_LHM2
 lhm3.Caption = f_lHM3
 lhm4.Caption = f_lHM4
 lhm5.Caption = f_lHM5
 lhm6.Caption = f_lHM6
 
 ' Gdb.Execute "ALTER TABLE PEL ADD ENERGOS B"
 
 f_lab1 = FindParametroi("PELAT2", "F_lab1", "Κειμ1", "Ετικέτα κειμ.1")
 f_lab2 = FindParametroi("PELAT2", "F_lab2", "Κειμ2", "Ετικέτα κειμ.2")
 f_lab3 = FindParametroi("PELAT2", "F_lab3", "Κειμ3", "Ετικέτα κειμ.3")
 f_lab4 = FindParametroi("PELAT2", "F_lab4", "Κειμ4", "Ετικέτα κειμ.4")
 f_lab5 = FindParametroi("PELAT2", "F_lab5", "Κειμ5", "Ετικέτα κειμ.5")
 f_lab6 = FindParametroi("PELAT2", "F_lab6", "Κειμ6", "Ετικέτα κειμ.6")
 LaB1.Caption = f_lab1
 Lab2.Caption = f_lab2
 LAB3.Caption = f_lab3
 Lab4.Caption = f_lab4
 Lab5.Caption = f_lab5
 Lab6.Caption = f_lab6
 
 
 f_num1 = FindParametroi("PELAT2", "F_num1", "Αριθ1", "Ετικέτα κειμ.1")
 f_num2 = FindParametroi("PELAT2", "F_num2", "Αριθ2", "Ετικέτα κειμ.2")
 f_num3 = FindParametroi("PELAT2", "F_num3", "Αριθ3", "Ετικέτα κειμ.3")
 f_num4 = FindParametroi("PELAT2", "F_num4", "Αριθ4", "Ετικέτα κειμ.4")
 LNUM1.Caption = f_num1
 LNUM2.Caption = f_num2
 LNUM3.Caption = f_num3
 Lnum4.Caption = f_num4
 
 
 
 
 
 
 Dim R As New ADODB.Recordset
 On Error GoTo UPDATE_STR
 
 R.Open "SELECT TOP 1 HM1 FROM PEL", Gdb, adOpenForwardOnly, adLockReadOnly
 On Error Resume Next
 
 
 R.Close

 
'SXOLIA CHAR(50)

 
 'DTPicker1.Value = Null
 'DTPicker2.Value = Null
 'dtpicker3.Value = Null
 'dtpicker4.Value = Null
 'dtpicker5.Value = Null
 'dtpicker6.Value = Null
 
 
 Me.Picture = LoadPicture(gPicture)
 PROTH = 1
 Text2(1) = "" 'goo
Command3.Caption = "Διόρθωση"   'ΣΔΦΔΣΦ

pelat2.Picture = LoadPicture(gPicture)
 

'tropos plhromhs
'Dim R As New ADODB.Recordset

 Adodc2.ConnectionString = gConnect
 
 Adodc2.RecordSource = "SELECT STR(AYJON,2,0) AS AYJ,PERIGRAFH FROM PINAKES WHERE TYPOS=13 ORDER BY PERIGRAFH"
 Adodc2.Refresh

 
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=12 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
   
Do While Not R.EOF
   If R("typos") = 12 Then
      Combo2.AddItem Str(R("AYJON")) + "." + R("PERIGRAFH")
   End If
   R.MoveNext
Loop
 
 
'KATHGORIA
  R.Close
  R.Open "SELECT *FROM PINAKES WHERE TYPOS=13 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
   If R.EOF Then
     Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH) VALUES (13,1,'ΕΜΠΟΡΟΣ')"
   End If




Do While Not R.EOF
   If R("typos") = 13 Then
      KATHGORIA.AddItem Str(R("AYJON")) + "." + R("PERIGRAFH")
   End If
   R.MoveNext
Loop
 R.Close
 
 
 
 
 
' R.Close








If f_TIMOKAT = 1 Then
  Label1(11).Caption = "Τιμοκατάλογος"
End If


 Dim mf2
  
If Len(f_files) = 0 Then
Else
  mf2 = UCase(f_files + "\MERCVB\files\")
 Dir1.Path = mf2
 Dir1.Refresh
End If



Exit Sub


UPDATE_STR:
Gdb.Execute "ALTER TABLE PEL ADD HM1 DATETIME"
Gdb.Execute "ALTER TABLE PEL ADD HM2 DATETIME"
Gdb.Execute "ALTER TABLE PEL ADD HM3 DATETIME"
Gdb.Execute "ALTER TABLE PEL ADD HM4 DATETIME"
Gdb.Execute "ALTER TABLE PEL ADD HM5 DATETIME"
Gdb.Execute "ALTER TABLE PEL ADD HM6 DATETIME"
 
Gdb.Execute "ALTER TABLE PEL ADD CH1 CHAR(30)"
Gdb.Execute "ALTER TABLE PEL ADD CH2 CHAR(30)"
Gdb.Execute "ALTER TABLE PEL ADD CH3 CHAR(30)"
Gdb.Execute "ALTER TABLE PEL ADD CH4 CHAR(30)"
Gdb.Execute "ALTER TABLE PEL ADD CH5 CHAR(30)"
Gdb.Execute "ALTER TABLE PEL ADD CH6 CHAR(30)"
Gdb.Execute "ALTER TABLE PEL ADD ENERGOS BIT"

Resume Next







End Sub




Private Sub Form_Paint()
If PROTH = 1 Then
   PROTH = 2
   'data1.DatabaseName = gDir
   
   Data1.RecordSource = "SELECT TOP 10 *FROM PEL WHERE EIDOS='X'"
   Data1.ConnectionString = gConnect
   Data1.Refresh
   'Text2(0).SetFocus
  PELPROM.SetFocus '  Option1(1).SetFocus
  
  
  If Len(Trim(ffbuff.Caption)) > 0 Then
    If Left(ffbuff.Caption, 1) = "e" Then
         PELPROM.ListIndex = 0
    Else
         PELPROM.ListIndex = 1
    End If
    
    Text2(0).Text = Trim(mID(ffbuff, 2, 10))
    Text2_LostFocus 0
    
   
   End If
  
  
  
  
  
  
  
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'goo = Text2(1).Text
End Sub


Private Sub Label2_Click()
   

End Sub

Private Sub Label1_Click(Index As Integer)
    
 If Index = 18 Then
    UPDATE_PINAKES 12, Label1(Index), Combo2
 End If
 
End Sub

'Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'       KeyAscii = 0 'suppress the beep
'       Text2(0).SetFocus
'       'keybd_event VK_TAB, 0, 0, 0 'send a tab
'    End If
'End Sub



Private Sub Picture1_Click()
  'αλλαγή γραμματοσειράς
  
    CommonDialog1.flags = cdlCFBoth   ' Flags property must be set
    CommonDialog1.ShowFont
    With Text3
'        .SELSTART = FSELSTART
'        .SELLENGTH = FSELLENGTH
        
       .SelFontName = CommonDialog1.FontName
       .SelFontSize = CommonDialog1.FONTSIZE
       .SelBold = CommonDialog1.FontBold
       .SelItalic = CommonDialog1.FontItalic
       .SelStrikeThru = CommonDialog1.FontStrikethru
       .SelUnderline = CommonDialog1.FontUnderline
    
    End With

End Sub

Private Sub Picture2_Click()
  'αλλαγή γραμματοσειράς
   CommonDialog1.ShowColor
   
'   Text3.SELSTART = FSELSTART
'        Text3.SELLENGTH = FSELLENGTH
   
   Text3.SelColor = CommonDialog1.Color
   
   
   
End Sub
Private Sub Text1_GotFocus(Index As Integer)
  Text1(Index).BackColor = vbYellow
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
  Dim b, A
 Text1(Index).BackColor = vbWhite
  If Index = 5 Then
    b = Text1(5).Text
    A = check_afm(b)
  End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   Text2(Index).BackColor = vbYellow
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0 'suppress the beep
      keybd_event VK_TAB, 0, 0, 0 'send a tab
   End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)

Dim X As String
Text2(Index).BackColor = vbWhite
If Len(Trim(Text2(Index))) = 0 Then Exit Sub
'If Option1(1) Then X = "e" Else X = "r"

X = "e"
On Error Resume Next
X = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)

'If gUserId = 2 Then  ' TAMIAS KAGIOYDIS
'
'   If Index = 1 Then  'eponumo
'      DATA1.RecordSource = "select EPO,THL,*from PEL where KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%' AND EIDOS='" + x + "' and EPO like '" + UCase(SameLetters(Text2(1).Text)) + "%';"
'   End If
'
'   If Index = 0 Then 'kodikos
'      DATA1.RecordSource = "select EPO,THL,*from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND  EIDOS='" + x + "' and KOD like '" + Text2(0).Text + "%';"
'   End If
'
'   If Index = 2 Then 'afm
'       DATA1.RecordSource = "select EPO,THL.* from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND  EIDOS='" + x + "' and AFM like '" + Text2(2).Text + "%';"
'   End If











If gUserId = 2 Then


If Index = 1 Then  'eponumo
    Data1.RecordSource = "select *from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND EIDOS='" + X + "' and EPO like '" + UCase(SameLetters(Text2(1).Text)) + "%';"
End If
If Index = 0 Then 'kodikos
    Data1.RecordSource = "select *from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND EIDOS='" + X + "' and KOD like '" + UCase(SameLetters(Text2(0).Text)) + "%';"
End If
If Index = 2 Then 'afm
    Data1.RecordSource = "select *from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND EIDOS='" + X + "' and AFM like '" + Text2(2).Text + "%';"
End If


Else
Dim FF
'If Option1(1) Then X = "e" Else X = "r"
If Index = 1 Then  'eponumo

   If Asc(Left(Text2(1).Text, 1)) > 127 Then
     FF = Text2(1).Text
   Else
     FF = UCase(SameLetters(Text2(1).Text))
   End If
   
    Data1.RecordSource = "select *from PEL where EIDOS='" + X + "' and EPO like '" + FF + "%';"
End If
If Index = 0 Then 'kodikos
    Data1.RecordSource = "select *from PEL where EIDOS='" + X + "' and KOD like '" + UCase(SameLetters(Text2(0).Text)) + "%';"
End If
If Index = 2 Then 'afm
    Data1.RecordSource = "select *from PEL where EIDOS='" + X + "' and AFM like '" + Text2(2).Text + "%';"
End If

End If


'DataGrid1.Enabled = False

Data1.Refresh
DataGrid1.Columns(0).width = TextWidth("A") * 30


'DataGrid1.Enabled = True


Dim mF As String

'  If Option1(1).value = True Then
     mF = UCase(f_files + "\MERCVB\files\" + mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1) + Text1(0))
'  Else
'     mF = UCase(f_files + "\MERCVB\files\r" + Text1(0))
'  End If
On Error Resume Next


If Text1(0).Text = Data1.Recordset("KOD") Then  'GIA NA MHN KATHISTEREI
 If Len(Dir(mF, vbDirectory)) > 0 Then
   File1.Path = mF
 Else
   If Len(Dir("c:\mercvb\files\dum", vbDirectory)) = 0 Then
     If Len(Dir("c:\mercvb\files", vbDirectory)) = 0 Then
        MkDir "c:\mercvb\files"
     End If
        MkDir "c:\mercvb\files\dum"
   End If
  
   File1.Path = "c:\mercvb\files\dum"
 End If

   
   File1.Refresh

End If











On Error Resume Next
DataGrid1.SetFocus


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
Dim db As Database
Dim R As New ADODB.Recordset
Dim R0 As New ADODB.Recordset

Dim k As Long

On Error Resume Next


'Set db = OpenDatabase(gDir)
'Set db = OpenDatabase(gDir, False, False)
'Set db = OpenDatabase(gDir, False, False, gConnect)


If True Then 'eyresi kenoy aritmoy
  R.Open "select KOD FROM PEL WHERE LEN(KOD)=4 AND EIDOS='" + meidos + "' order by KOD", Gdb, adOpenForwardOnly, adLockReadOnly
  If R.EOF Then
     get_next_pel = "0000"
    R.Close: Exit Function
  End If
  
  Me.MousePointer = vbHourglass
  

  R.MoveFirst
  k = 0
  Do While Not R.EOF
   If Val(R(0)) = 0 Then
      R.MoveNext
   Else
     k = k + 1
     If Left(R(0), 4) = Format(k, "0000") Then
        R.MoveNext
     Else
        On Error GoTo 0
       R0.Open "select COUNT(*) FROM PEL WHERE KOD='" + Format(k, "0000") + "' AND EIDOS='" + meidos + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
       If R0(0) > 0 Then 'ΗΔΗ ΥΠΑΡΧΕΙ
          R0.Close
          Exit Do
       Else
         get_next_pel = Format(k, "0000")
         R.Close
         Me.MousePointer = vbNormal
         Exit Function
       End If
     
     
     End If
   End If
  Loop
  R.Close

Dim ll, L
'

For L = k + 1 To 9999
      R.Open "select COUNT(*) FROM PEL WHERE KOD='" + Format(L, "0000") + "' AND EIDOS='" + meidos + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
      ll = R(0)
      R.Close
      DoEvents
      Me.Caption = k
      If ll = 0 Then
        Exit For
      End If
Next
get_next_pel = Format(L, "0000")
Me.MousePointer = vbNormal
Exit Function
  
  
  
  
  
  
  
End If





R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic

If meidos = "e" Then
   get_next_pel = R("epel") + 1
Else
   get_next_pel = R("rpel") + 1
End If
End Function

Private Sub Text3_LostFocus()
   
   Text3.SelStart = FSELSTART
        Text3.SelLength = FSELLENGTH
   
   





End Sub

Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    FSELSTART = Text3.SelStart
    FSELLENGTH = Text3.SelLength


End Sub

