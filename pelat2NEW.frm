VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form pelat2 
   BackColor       =   &H00FF0000&
   Caption         =   "ƒÈ¸ÒË˘ÛÁ"
   ClientHeight    =   12525
   ClientLeft      =   195
   ClientTop       =   510
   ClientWidth     =   21780
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12525
   ScaleWidth      =   21780
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5640
      Top             =   11760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      Destination     =   1
      DiscardSavedData=   -1  'True
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      TabIndex        =   111
      Top             =   0
      Width           =   21780
      _ExtentX        =   38418
      _ExtentY        =   1799
      ButtonWidth     =   2593
      ButtonHeight    =   1640
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Õ›· ≈„„Ò·ˆﬁ"
            Key             =   "cmdNew"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ƒÈ¸ÒË˘ÛÁ"
            Key             =   "cmdDiortosi"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ƒÈ·„Ò·ˆﬁ"
            Key             =   "cmdDelete"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "¡ÔËﬁÍÂıÛÁ"
            Key             =   "cmdUpdate"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " ·ÒÙ›ÎÎ·"
            Key             =   "cmdKartela"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "≈ÙÈÍ›ÙÙ·"
            Key             =   "cmdEtiketa"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "¡ÎÎ·„ﬁ  ˘‰ÈÍÔ˝"
            Key             =   "cmdChange"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "≈ÓÔ‰ÔÚ"
            Key             =   "cmdExit"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Height          =   2532
      Left            =   0
      TabIndex        =   105
      Top             =   720
      Width           =   17715
      Begin VB.CheckBox chkANENERGOS 
         BackColor       =   &H00C0FFC0&
         Caption         =   " ·È ·ÌÂÌÂÒ„Ô˝Ú"
         Height          =   255
         Left            =   1320
         TabIndex        =   128
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text2 
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
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Index           =   2
         Left            =   1320
         TabIndex        =   107
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox PELPROM 
         Height          =   288
         ItemData        =   "pelat2NEW.frx":0000
         Left            =   1320
         List            =   "pelat2NEW.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   1800
         Width           =   1692
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "pelat2NEW.frx":0028
         Height          =   1896
         Left            =   3240
         TabIndex        =   2
         Top             =   360
         Width           =   13092
         _ExtentX        =   23098
         _ExtentY        =   3334
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   24
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
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "EPO"
            Caption         =   "≈–ŸÕ’Ãœ"
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
            Caption         =   "ƒ…≈’»’Õ”«"
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
            Caption         =   "‘«À≈÷ŸÕœ"
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
            Caption         =   " Ÿƒ… œ”"
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
            DataField       =   "EPA"
            Caption         =   "E–¡√√≈ÀÃ¡"
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
         BeginProperty Column05 
            DataField       =   "POL"
            Caption         =   "–œÀ«"
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
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ˘‰ÈÍ¸Ú"
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
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   110
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈˘ÌıÏﬂ·"
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
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   109
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A.÷.Ã."
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
         Height          =   252
         Index           =   2
         Left            =   360
         TabIndex        =   108
         Top             =   1320
         Width           =   1332
      End
   End
   Begin VB.ComboBox kentroAdyn 
      Height          =   288
      ItemData        =   "pelat2NEW.frx":003C
      Left            =   17400
      List            =   "pelat2NEW.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2700
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11640
      Top             =   10920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   336
      Left            =   14220
      Top             =   11880
      Visible         =   0   'False
      Width           =   2772
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   15
      Top             =   11010
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   7932
      Left            =   -30
      TabIndex        =   4
      Top             =   3450
      Width           =   17715
      Begin VB.TextBox BRANCH 
         DataField       =   "BRANCH"
         DataSource      =   "DATA1"
         Height          =   375
         Left            =   13320
         TabIndex        =   140
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtText3 
         DataField       =   "STREETNUMBER"
         DataSource      =   "DATA1"
         Height          =   375
         Left            =   5460
         TabIndex        =   139
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox chkEnergos 
         BackColor       =   &H00C0FFC0&
         Caption         =   "≈ÌÂÒ„¸Ú"
         Height          =   375
         Left            =   11160
         TabIndex        =   129
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdAllaghKodikou 
         Caption         =   "¡ÎÎ·„ﬁ  ˘‰ÈÍÔ˝"
         Height          =   375
         Left            =   2640
         TabIndex        =   127
         Top             =   240
         Width           =   990
      End
      Begin VB.TextBox Text1 
         DataField       =   "ADT"
         DataSource      =   "DATA1"
         Height          =   375
         Index           =   16
         Left            =   14040
         TabIndex        =   124
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox COUNTRY 
         DataField       =   "COUNTRY"
         DataSource      =   "DATA1"
         Height          =   375
         Left            =   13140
         TabIndex        =   122
         Top             =   1200
         Width           =   852
      End
      Begin VB.CommandButton cmd√√ 
         Caption         =   "√.√.–.”."
         Height          =   360
         Left            =   3360
         TabIndex        =   118
         Top             =   1440
         Width           =   990
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "¡ÔÛÙÔÎﬁ ¡Ò˜ÂﬂÔı"
         Height          =   690
         Left            =   16605
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   3120
         Width           =   855
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   6360
         TabIndex        =   103
         Top             =   7560
         Visible         =   0   'False
         Width           =   2772
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ƒÈ·„Ò·ˆﬁ ¡Ò˜ÂﬂÔı"
         Height          =   795
         Left            =   16605
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   3960
         Width           =   870
      End
      Begin VB.CommandButton Command14 
         Caption         =   "≈˝ÒÂÛÁ ·Ò˜ÂﬂÔı ÂÈÍ¸Ì·Ú"
         Height          =   810
         Left            =   16605
         TabIndex        =   101
         Top             =   2010
         Width           =   810
      End
      Begin VB.TextBox Text1 
         DataField       =   "DOY"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   5460
         TabIndex        =   77
         Top             =   1440
         Width           =   2200
      End
      Begin VB.TextBox Text1 
         DataField       =   "THL"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   10020
         TabIndex        =   78
         Top             =   1200
         Width           =   2200
      End
      Begin VB.TextBox Text1 
         DataField       =   "AFM"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   1140
         TabIndex        =   76
         Top             =   1440
         Width           =   2200
      End
      Begin VB.TextBox Text1 
         DataField       =   "EPA"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   10020
         TabIndex        =   72
         Top             =   240
         Width           =   2200
      End
      Begin VB.TextBox Text1 
         DataField       =   "POL"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   6720
         TabIndex        =   74
         Top             =   720
         Width           =   2040
      End
      Begin VB.TextBox Text1 
         DataField       =   "DIE"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   2
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         DataField       =   "EPO"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   5460
         TabIndex        =   71
         Top             =   240
         Width           =   3285
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
         TabIndex        =   70
         Top             =   240
         Width           =   1392
      End
      Begin VB.TextBox Text1 
         DataField       =   "XRVMA"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   13
         Left            =   10020
         TabIndex        =   75
         Top             =   720
         Width           =   840
      End
      Begin VB.DirListBox Dir1 
         Height          =   315
         Left            =   3960
         TabIndex        =   69
         Top             =   7680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox SYMFPOSO 
         DataField       =   "NUM1"
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
         Height          =   285
         Left            =   16740
         TabIndex        =   68
         ToolTipText     =   "¡ÔÛÙÔÎﬁ ÛÂ ”ÁÏÂÈ˛ÛÂÈÚ: F7 ”ıÏˆ˘Ìﬂ·  F5 EÙÔÈÏÔ ÍÂﬂÏÂÌÔ  "
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox mhxanhmata 
         DataField       =   "NUM2"
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
         Height          =   285
         Left            =   18660
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox HMESYMF 
         DataField       =   "HM1"
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
         Height          =   285
         Left            =   17400
         TabIndex        =   66
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox foto2 
         Height          =   2895
         Left            =   13563
         ScaleHeight     =   2835
         ScaleWidth      =   2715
         TabIndex        =   65
         Top             =   1980
         Width           =   2772
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5412
         Left            =   180
         TabIndex        =   5
         Top             =   1980
         Width           =   8892
         _ExtentX        =   15690
         _ExtentY        =   9551
         _Version        =   393216
         Tabs            =   6
         Tab             =   5
         TabHeight       =   520
         BackColor       =   16777088
         TabCaption(0)   =   "–≈ƒ…¡ «Ã≈—/Õ…¡”"
         TabPicture(0)   =   "pelat2NEW.frx":0059
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lhm6"
         Tab(0).Control(1)=   "lhm4"
         Tab(0).Control(2)=   "lhm3"
         Tab(0).Control(3)=   "lhm2"
         Tab(0).Control(4)=   "lhm1"
         Tab(0).Control(5)=   "lhm5"
         Tab(0).Control(6)=   "dtpicker6"
         Tab(0).Control(7)=   "dtpicker4"
         Tab(0).Control(8)=   "dtpicker3"
         Tab(0).Control(9)=   "dtpicker2"
         Tab(0).Control(10)=   "dtpicker1"
         Tab(0).Control(11)=   "dtpicker5"
         Tab(0).ControlCount=   12
         TabCaption(1)   =   " ≈…Ã≈Õœ’"
         TabPicture(1)   =   "pelat2NEW.frx":0075
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lab7"
         Tab(1).Control(1)=   "lab6"
         Tab(1).Control(2)=   "lab5"
         Tab(1).Control(3)=   "lab4"
         Tab(1).Control(4)=   "lab3"
         Tab(1).Control(5)=   "lab2"
         Tab(1).Control(6)=   "lab1"
         Tab(1).Control(7)=   "keim(1)"
         Tab(1).Control(8)=   "keim(7)"
         Tab(1).Control(9)=   "keim(6)"
         Tab(1).Control(10)=   "keim(5)"
         Tab(1).Control(11)=   "keim(4)"
         Tab(1).Control(12)=   "keim(3)"
         Tab(1).Control(13)=   "keim(2)"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   "¡—…»Ã«‘… ¡ "
         TabPicture(2)   =   "pelat2NEW.frx":0091
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Lnum1"
         Tab(2).Control(1)=   "Lnum2"
         Tab(2).Control(2)=   "Lnum3"
         Tab(2).Control(3)=   "Lnum4"
         Tab(2).Control(4)=   "Text8"
         Tab(2).Control(5)=   "Text7"
         Tab(2).Control(6)=   "Text6"
         Tab(2).Control(7)=   "Text5"
         Tab(2).ControlCount=   8
         TabCaption(3)   =   "–ÔÎ.≈ÈÎÔ„˛Ì"
         TabPicture(3)   =   "pelat2NEW.frx":00AD
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lbcb1"
         Tab(3).Control(1)=   "lbcb2"
         Tab(3).Control(2)=   "lbcb3"
         Tab(3).Control(3)=   "lbcb4"
         Tab(3).Control(4)=   "lbcb5"
         Tab(3).Control(5)=   "tComb1"
         Tab(3).Control(6)=   "tComb2"
         Tab(3).Control(7)=   "tComb3"
         Tab(3).Control(8)=   "tComb4"
         Tab(3).Control(9)=   "tComb5"
         Tab(3).Control(10)=   "Comb1"
         Tab(3).Control(11)=   "Comb2"
         Tab(3).Control(12)=   "Comb3"
         Tab(3).Control(13)=   "Comb4"
         Tab(3).Control(14)=   "Comb5"
         Tab(3).ControlCount=   15
         TabCaption(4)   =   "œ… œÕœÃ… ¡ ”‘œ…◊≈…¡"
         TabPicture(4)   =   "pelat2NEW.frx":00C9
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdTIMOKAT"
         Tab(4).Control(1)=   "Text1(8)"
         Tab(4).Control(2)=   "Combo2"
         Tab(4).Control(3)=   "Text4"
         Tab(4).Control(4)=   "Combo1"
         Tab(4).Control(5)=   "Text1(11)"
         Tab(4).Control(6)=   "Text1(12)"
         Tab(4).Control(7)=   "Text1(10)"
         Tab(4).Control(8)=   "Text1(9)"
         Tab(4).Control(9)=   "Combo3"
         Tab(4).Control(10)=   "DataCombo2"
         Tab(4).Control(11)=   "Label1(11)"
         Tab(4).Control(12)=   "Label1(25)"
         Tab(4).Control(13)=   "Label1(17)"
         Tab(4).Control(14)=   "Label1(16)"
         Tab(4).Control(15)=   "Label1(15)"
         Tab(4).Control(16)=   "Label1(14)"
         Tab(4).Control(17)=   "Label1(13)"
         Tab(4).Control(18)=   "Label1(12)"
         Tab(4).Control(19)=   "Label1(18)"
         Tab(4).Control(20)=   "Label1(20)"
         Tab(4).Control(21)=   "Label1(21)"
         Tab(4).ControlCount=   22
         TabCaption(5)   =   "B2G-email- …Õ«‘œ"
         TabPicture(5)   =   "pelat2NEW.frx":00E5
         Tab(5).ControlEnabled=   -1  'True
         Tab(5).Control(0)=   "Label1(23)"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "Label1(22)"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).Control(2)=   "Label3"
         Tab(5).Control(2).Enabled=   0   'False
         Tab(5).Control(3)=   "Label4"
         Tab(5).Control(3).Enabled=   0   'False
         Tab(5).Control(4)=   "Label5"
         Tab(5).Control(4).Enabled=   0   'False
         Tab(5).Control(5)=   "Label6"
         Tab(5).Control(5).Enabled=   0   'False
         Tab(5).Control(6)=   "TRAINING"
         Tab(5).Control(6).Enabled=   0   'False
         Tab(5).Control(7)=   "Text1(15)"
         Tab(5).Control(7).Enabled=   0   'False
         Tab(5).Control(8)=   "Text1(14)"
         Tab(5).Control(8).Enabled=   0   'False
         Tab(5).Control(9)=   "CONTRACTIDENTIFIER"
         Tab(5).Control(9).Enabled=   0   'False
         Tab(5).Control(10)=   "BUDTYPE"
         Tab(5).Control(10).Enabled=   0   'False
         Tab(5).Control(11)=   "CONTRACTINGAUTHORITYID"
         Tab(5).Control(11).Enabled=   0   'False
         Tab(5).Control(12)=   "BUDIDENTIFIER"
         Tab(5).Control(12).Enabled=   0   'False
         Tab(5).ControlCount=   13
         Begin VB.TextBox BUDIDENTIFIER 
            BackColor       =   &H0080C0FF&
            DataField       =   "BUDIDENTIFIER"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4080
            TabIndex        =   133
            Top             =   3000
            Width           =   3285
         End
         Begin VB.TextBox CONTRACTINGAUTHORITYID 
            BackColor       =   &H0080C0FF&
            DataField       =   "CONTRACTINGAUTHORITYID"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4080
            TabIndex        =   132
            Top             =   1080
            Width           =   3285
         End
         Begin VB.TextBox BUDTYPE 
            BackColor       =   &H0080C0FF&
            DataField       =   "BUDTYPE"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4080
            TabIndex        =   131
            Top             =   2100
            Width           =   3285
         End
         Begin VB.TextBox CONTRACTIDENTIFIER 
            BackColor       =   &H0080C0FF&
            DataField       =   "CONTRACTIDENTIFIER"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4080
            TabIndex        =   130
            Top             =   1560
            Width           =   3285
         End
         Begin VB.CommandButton cmdTIMOKAT 
            Caption         =   "TIMOKATAÀO√œ”"
            Height          =   360
            Left            =   -69000
            TabIndex        =   126
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox dtpicker5 
            BackColor       =   &H0080C0FF&
            DataField       =   "HM5"
            DataSource      =   "DATA1"
            Height          =   330
            Left            =   -71475
            TabIndex        =   119
            Top             =   2685
            Width           =   4404
         End
         Begin VB.TextBox Text1 
            DataField       =   "EMAIL"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   14
            Left            =   4080
            TabIndex        =   113
            Top             =   4380
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            DataField       =   "KINHTO"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   15
            Left            =   4080
            TabIndex        =   112
            Top             =   4920
            Width           =   3285
         End
         Begin VB.TextBox keim 
            BackColor       =   &H0080C0FF&
            DataField       =   "CH2"
            DataSource      =   "DATA1"
            Height          =   330
            Index           =   2
            Left            =   -71880
            TabIndex        =   18
            Top             =   1200
            Width           =   5000
         End
         Begin VB.TextBox keim 
            BackColor       =   &H0080C0FF&
            DataField       =   "CH3"
            DataSource      =   "DATA1"
            Height          =   330
            Index           =   3
            Left            =   -71880
            TabIndex        =   19
            Top             =   1680
            Width           =   5000
         End
         Begin VB.TextBox keim 
            BackColor       =   &H0080C0FF&
            DataField       =   "CH4"
            DataSource      =   "DATA1"
            Height          =   330
            Index           =   4
            Left            =   -71880
            TabIndex        =   20
            Top             =   2160
            Width           =   5000
         End
         Begin VB.TextBox keim 
            BackColor       =   &H0080C0FF&
            DataField       =   "CH5"
            DataSource      =   "DATA1"
            Height          =   330
            Index           =   5
            Left            =   -71880
            TabIndex        =   21
            Top             =   2640
            Width           =   5000
         End
         Begin VB.TextBox keim 
            BackColor       =   &H0080C0FF&
            DataField       =   "CH6"
            DataSource      =   "DATA1"
            Height          =   330
            Index           =   6
            Left            =   -71880
            TabIndex        =   22
            Top             =   3120
            Width           =   5000
         End
         Begin VB.TextBox dtpicker1 
            BackColor       =   &H0080C0FF&
            DataField       =   "HM1"
            DataSource      =   "DATA1"
            Height          =   330
            Left            =   -71475
            TabIndex        =   32
            Top             =   900
            Width           =   4404
         End
         Begin VB.TextBox dtpicker2 
            BackColor       =   &H0080C0FF&
            DataField       =   "HM2"
            DataSource      =   "DATA1"
            Height          =   330
            Left            =   -71475
            TabIndex        =   31
            Top             =   1335
            Width           =   4404
         End
         Begin VB.TextBox dtpicker3 
            BackColor       =   &H0080C0FF&
            DataField       =   "HM3"
            DataSource      =   "DATA1"
            Height          =   330
            Left            =   -71475
            TabIndex        =   30
            Top             =   1785
            Width           =   4404
         End
         Begin VB.TextBox dtpicker4 
            BackColor       =   &H0080C0FF&
            DataField       =   "HM4"
            DataSource      =   "DATA1"
            Height          =   330
            Left            =   -71475
            TabIndex        =   29
            Top             =   2265
            Width           =   4404
         End
         Begin VB.TextBox dtpicker6 
            BackColor       =   &H0080C0FF&
            DataField       =   "HM6"
            DataSource      =   "DATA1"
            Height          =   330
            Left            =   -71475
            TabIndex        =   28
            Top             =   3075
            Width           =   4404
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H0080C0FF&
            DataField       =   "NUM4"
            DataSource      =   "DATA1"
            Height          =   285
            Left            =   -71790
            TabIndex        =   27
            Top             =   2760
            Width           =   5000
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H0080C0FF&
            DataField       =   "NUM3"
            DataSource      =   "DATA1"
            Height          =   330
            Left            =   -71790
            TabIndex        =   26
            Top             =   2280
            Width           =   5000
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H0080C0FF&
            DataField       =   "NUM2"
            DataSource      =   "DATA1"
            Height          =   315
            Left            =   -71790
            TabIndex        =   25
            Top             =   1830
            Width           =   5000
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H0080C0FF&
            DataField       =   "NUM1"
            DataSource      =   "DATA1"
            Height          =   330
            Left            =   -71790
            TabIndex        =   24
            Top             =   1395
            Width           =   5000
         End
         Begin VB.TextBox keim 
            BackColor       =   &H0080C0FF&
            DataField       =   "CH7"
            DataSource      =   "DATA1"
            Height          =   330
            Index           =   7
            Left            =   -71880
            TabIndex        =   23
            Top             =   3600
            Width           =   5000
         End
         Begin VB.TextBox keim 
            BackColor       =   &H0080C0FF&
            DataField       =   "CH1"
            DataSource      =   "DATA1"
            Height          =   330
            Index           =   1
            Left            =   -71880
            TabIndex        =   17
            Top             =   687
            Width           =   5000
         End
         Begin VB.ComboBox Comb5 
            Height          =   315
            Left            =   -72480
            TabIndex        =   16
            Top             =   3240
            Width           =   3135
         End
         Begin VB.ComboBox Comb4 
            Height          =   315
            Left            =   -72480
            TabIndex        =   15
            Top             =   2760
            Width           =   3135
         End
         Begin VB.ComboBox Comb3 
            Height          =   315
            Left            =   -72480
            TabIndex        =   14
            Top             =   2280
            Width           =   3135
         End
         Begin VB.ComboBox Comb2 
            Height          =   315
            Left            =   -72480
            TabIndex        =   13
            Top             =   1800
            Width           =   3135
         End
         Begin VB.ComboBox Comb1 
            Height          =   315
            Left            =   -72480
            TabIndex        =   12
            Top             =   1320
            Width           =   3135
         End
         Begin VB.TextBox tComb5 
            DataField       =   "COMB5"
            DataSource      =   "DATA1"
            Height          =   285
            Left            =   -69360
            TabIndex        =   11
            Top             =   3240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox tComb4 
            DataField       =   "COMB4"
            DataSource      =   "DATA1"
            Height          =   285
            Left            =   -69360
            TabIndex        =   10
            Top             =   2760
            Visible         =   0   'False
            Width           =   168
         End
         Begin VB.TextBox tComb3 
            DataField       =   "COMB3"
            DataSource      =   "DATA1"
            Height          =   285
            Left            =   -69360
            TabIndex        =   9
            Top             =   2280
            Visible         =   0   'False
            Width           =   168
         End
         Begin VB.TextBox tComb2 
            DataField       =   "COMB2"
            DataSource      =   "DATA1"
            Height          =   285
            Left            =   -69360
            TabIndex        =   8
            Top             =   1800
            Visible         =   0   'False
            Width           =   168
         End
         Begin VB.TextBox tComb1 
            DataField       =   "COMB1"
            DataSource      =   "DATA1"
            Height          =   375
            Left            =   -69360
            TabIndex        =   7
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text1 
            DataField       =   "PEK"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   8
            Left            =   -71520
            TabIndex        =   81
            Top             =   720
            Width           =   2200
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "ARPARAG"
            DataSource      =   "DATA1"
            Height          =   315
            Left            =   -71520
            TabIndex        =   92
            Top             =   4080
            Width           =   2055
         End
         Begin VB.TextBox Text4 
            DataField       =   "PLAISIO"
            DataSource      =   "DATA1"
            Height          =   375
            Left            =   -71520
            TabIndex        =   90
            Top             =   3600
            Width           =   2655
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "TYPOS"
            DataSource      =   "DATA1"
            Height          =   315
            ItemData        =   "pelat2NEW.frx":0101
            Left            =   -71520
            List            =   "pelat2NEW.frx":0111
            TabIndex        =   88
            Top             =   3120
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#####0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   11
            Left            =   -71520
            TabIndex        =   86
            Top             =   2160
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
            Left            =   -71520
            TabIndex        =   6
            Top             =   2640
            Width           =   2200
         End
         Begin VB.TextBox Text1 
            DataField       =   "SHM2"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   10
            Left            =   -71520
            TabIndex        =   84
            Top             =   1680
            Width           =   2200
         End
         Begin VB.TextBox Text1 
            DataField       =   "PLAFON"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   9
            Left            =   -71520
            TabIndex        =   82
            Top             =   1200
            Width           =   2200
         End
         Begin VB.ComboBox Combo3 
            DataField       =   "PVLHTHS"
            DataSource      =   "DATA1"
            Height          =   315
            Left            =   -71520
            TabIndex        =   96
            Top             =   4860
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "pelat2NEW.frx":0150
            DataField       =   "KODGAL"
            DataSource      =   "DATA1"
            Height          =   315
            Left            =   -71520
            TabIndex        =   94
            Top             =   4500
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "PERIGRAFH"
            BoundColumn     =   "AYJ"
            Text            =   "DataCombo1"
         End
         Begin MSDataGridLib.DataGrid TRAINING 
            Bindings        =   "pelat2NEW.frx":0165
            Height          =   495
            Left            =   720
            TabIndex        =   121
            Top             =   4200
            Visible         =   0   'False
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   873
            _Version        =   393216
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
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "CENARXH"
               Caption         =   "¡–œ"
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
               DataField       =   "CLHXH"
               Caption         =   "≈Ÿ”"
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
               DataField       =   "MHNAS"
               Caption         =   "Ã«Õ¡”"
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
               DataField       =   "SYNEDRIES"
               Caption         =   "”’Õ≈ƒ—…≈”"
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "¡ƒ¡/≈Õ¡—…»Ãœ” ”’Ã¬¡”«”"
            Height          =   195
            Left            =   1800
            TabIndex        =   137
            Top             =   3000
            Width           =   2160
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "‘’–œ” –—œ’–œÀœ√…”Ãœ’1-3"
            Height          =   195
            Left            =   1800
            TabIndex        =   136
            Top             =   2160
            Width           =   2220
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "¡ƒ¡Ã ¡—.”’Ã¬¡”«”"
            Height          =   195
            Left            =   1800
            TabIndex        =   135
            Top             =   1680
            Width           =   1590
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "¡Õ¡»≈‘œ’”¡ ¡—◊«"
            Height          =   195
            Left            =   1800
            TabIndex        =   134
            Top             =   1080
            Width           =   1500
         End
         Begin VB.Label lhm5 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   -74520
            TabIndex        =   120
            Top             =   2685
            Width           =   3000
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "≈ÍÙ˘ÛÁ %/‘ÈÏÔÍ·Ù‹ÎÔ„ÔÚ"
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
            Height          =   375
            Index           =   11
            Left            =   -74370
            TabIndex        =   116
            Top             =   720
            Width           =   2400
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   22
            Left            =   2100
            TabIndex        =   115
            Top             =   4380
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   23
            Left            =   2160
            TabIndex        =   114
            Top             =   4920
            Width           =   975
         End
         Begin VB.Label lhm1 
            BackColor       =   &H00C0E0FF&
            Caption         =   ".................................................................................."
            Height          =   330
            Left            =   -74520
            TabIndex        =   63
            Top             =   900
            Width           =   3000
         End
         Begin VB.Label lhm2 
            BackColor       =   &H00C0E0FF&
            Height          =   336
            Left            =   -74520
            TabIndex        =   62
            Top             =   1320
            Width           =   3000
         End
         Begin VB.Label lhm3 
            BackColor       =   &H00C0E0FF&
            Height          =   336
            Left            =   -74520
            TabIndex        =   61
            Top             =   1800
            Width           =   3000
         End
         Begin VB.Label lhm4 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   -74520
            TabIndex        =   60
            Top             =   2295
            Width           =   3000
         End
         Begin VB.Label lhm6 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   -74520
            TabIndex        =   59
            Top             =   3075
            Width           =   3000
         End
         Begin VB.Label lab1 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   -74520
            TabIndex        =   58
            Top             =   687
            Width           =   2532
         End
         Begin VB.Label lab2 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   -74520
            TabIndex        =   57
            Top             =   1200
            Width           =   2532
         End
         Begin VB.Label lab3 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   -74520
            TabIndex        =   56
            Top             =   1680
            Width           =   2532
         End
         Begin VB.Label lab4 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   -74520
            TabIndex        =   55
            Top             =   2160
            Width           =   2532
         End
         Begin VB.Label lab5 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   -74520
            TabIndex        =   54
            Top             =   2640
            Width           =   2532
         End
         Begin VB.Label lab6 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   -74520
            TabIndex        =   53
            Top             =   3120
            Width           =   2532
         End
         Begin VB.Label Lnum4 
            BackColor       =   &H0080C0FF&
            Caption         =   "."
            Height          =   255
            Left            =   -74835
            TabIndex        =   52
            Top             =   2790
            Width           =   2500
         End
         Begin VB.Label Lnum3 
            BackColor       =   &H0080C0FF&
            Caption         =   "."
            Height          =   255
            Left            =   -74835
            TabIndex        =   51
            Top             =   2370
            Width           =   2500
         End
         Begin VB.Label Lnum2 
            BackColor       =   &H0080C0FF&
            Caption         =   "."
            Height          =   255
            Left            =   -74835
            TabIndex        =   50
            Top             =   1905
            Width           =   2500
         End
         Begin VB.Label Lnum1 
            BackColor       =   &H0080C0FF&
            Caption         =   "AÒÈËÏ1"
            Height          =   255
            Left            =   -74835
            TabIndex        =   49
            Top             =   1395
            Width           =   2500
         End
         Begin VB.Label lbcb5 
            Caption         =   " ÂÈÏ.–ÔÎ 5"
            Height          =   255
            Left            =   -74760
            TabIndex        =   48
            Top             =   3240
            Width           =   2055
         End
         Begin VB.Label lbcb4 
            Caption         =   " ÂÈÏ.–ÔÎ 4"
            Height          =   255
            Left            =   -74760
            TabIndex        =   47
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label lbcb3 
            Caption         =   " ÂÈÏ.–ÔÎ 3"
            Height          =   255
            Left            =   -74760
            TabIndex        =   46
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label lbcb2 
            Caption         =   " ÂÈÏ.–ÔÎ 2"
            Height          =   375
            Left            =   -74760
            TabIndex        =   45
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lbcb1 
            Caption         =   " ÂÈÏ.–ÔÎ 1"
            Height          =   255
            Left            =   -74760
            TabIndex        =   44
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "≈ÍÙ˘ÛÁ % TÈÏÔÍ·Ù‹ÎÔ„ÔÚ"
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
            Height          =   252
            Index           =   25
            Left            =   -74280
            TabIndex        =   43
            Top             =   11520
            Width           =   2652
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Ã≈‘¡÷œ—… π"
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
            Height          =   252
            Index           =   17
            Left            =   -74280
            TabIndex        =   42
            Top             =   3600
            Width           =   1452
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   " ·ËÂÛÙ˛Ú ÷–¡"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   372
            Index           =   16
            Left            =   -74280
            TabIndex        =   41
            Top             =   3120
            Width           =   1452
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "¡Ò˜.’¸ÎÔÈÔ"
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
            Height          =   372
            Index           =   15
            Left            =   -74280
            TabIndex        =   40
            Top             =   2160
            Width           =   1452
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "‘ÂÎÈÍ¸ ’¸ÎÔÈÔ"
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
            Height          =   372
            Index           =   14
            Left            =   -74280
            TabIndex        =   39
            Top             =   2640
            Width           =   1452
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   " ˘‰ÈÍ¸Ú ÀÔ„ÈÛÙÈÍﬁÚ"
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
            Height          =   372
            Index           =   13
            Left            =   -74280
            TabIndex        =   38
            Top             =   1680
            Width           =   1452
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "–Î·ˆ¸Ì"
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
            Height          =   372
            Index           =   12
            Left            =   -74280
            TabIndex        =   37
            Top             =   1200
            Width           =   1452
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "‘Ò¸ÔÚ –ÎÁÒ˘ÏﬁÚ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   18
            Left            =   -74280
            TabIndex        =   36
            Top             =   4080
            Width           =   2412
         End
         Begin VB.Label Lab7 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   -74520
            TabIndex        =   35
            Top             =   3600
            Width           =   2532
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   " ·ÙÁ„ÔÒﬂ·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   192
            Index           =   20
            Left            =   -74280
            TabIndex        =   34
            Top             =   4500
            Width           =   1452
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "–˘ÎÁÙÁÚ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   192
            Index           =   21
            Left            =   -74280
            TabIndex        =   33
            Top             =   4860
            Width           =   1452
         End
      End
      Begin ADOMERCNEW.msEditor2 msEditor1 
         Height          =   5475
         Left            =   9180
         TabIndex        =   64
         Top             =   1980
         Width           =   4335
         _extentx        =   7646
         _extenty        =   9551
         backcolor       =   12648447
         font            =   "pelat2NEW.frx":017A
         open_visible    =   0   'False
         save_visible    =   0   'False
         backcolor       =   12648447
         enabled         =   0   'False
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2520
         Left            =   13560
         TabIndex        =   117
         Top             =   4920
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4445
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "œÌÔÏ·"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "«ÏÂÒ/Ìﬂ·"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ã›„ÂËÔÚ"
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ˘‰.’ÔÍ"
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
         Left            =   12360
         TabIndex        =   141
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lbl¡ÒÈË 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¡ÒÈË."
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
         Left            =   4560
         TabIndex        =   138
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¡.ƒ.‘."
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
         Index           =   24
         Left            =   14640
         TabIndex        =   125
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "◊˛Ò·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   12240
         TabIndex        =   123
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ƒ.œ.’."
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
         Height          =   375
         Index           =   7
         Left            =   4500
         TabIndex        =   100
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "‘ÁÎ›ˆ˘ÌÔ"
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
         Height          =   372
         Index           =   6
         Left            =   8820
         TabIndex        =   99
         Top             =   1200
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "¡.÷.Ã."
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
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   98
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "◊˛Ò·/≈‹„„ÂÎ"
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
         Height          =   372
         Index           =   4
         Left            =   8820
         TabIndex        =   97
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "–¸ÎÁ"
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
         Height          =   375
         Index           =   3
         Left            =   6240
         TabIndex        =   95
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ƒÈÂ˝ËıÌÛÁ   œ‰¸Ú"
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
         Height          =   615
         Index           =   8
         Left            =   180
         TabIndex        =   93
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "≈˘ÌıÏﬂ·"
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
         Height          =   372
         Index           =   9
         Left            =   4500
         TabIndex        =   91
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " ˘‰ÈÍ¸Ú"
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
         Height          =   372
         Index           =   10
         Left            =   180
         TabIndex        =   89
         Top             =   240
         Width           =   1452
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
         ForeColor       =   &H00000000&
         Height          =   372
         Index           =   19
         Left            =   8820
         TabIndex        =   87
         Top             =   720
         Width           =   1452
      End
      Begin VB.Label ffbuff 
         Height          =   192
         Left            =   15084
         TabIndex        =   85
         Top             =   1680
         Visible         =   0   'False
         Width           =   1068
      End
      Begin VB.Label symf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”ıÏˆ.–ÔÛ¸"
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
         Index           =   1
         Left            =   15780
         TabIndex        =   83
         Top             =   480
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label symf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”ıÏˆ.ÏÁ˜·Ì"
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
         Index           =   0
         Left            =   17460
         TabIndex        =   80
         Top             =   480
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label symf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«ÏÂÒ.”ıÏˆ˘Ìﬂ·Ú"
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
         Index           =   2
         Left            =   16080
         TabIndex        =   79
         Top             =   960
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFC0&
         Height          =   1695
         Left            =   0
         Top             =   120
         Width           =   16515
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   3960
      _ExtentX        =   979
      _ExtentY        =   979
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":01A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":0A80
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":0ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":1324
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":1776
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":1BC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":2242
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":28BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":2F36
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":35B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":3C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":42A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat2NEW.frx":491E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADODC1 
      Height          =   336
      Left            =   0
      Top             =   0
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
End
Attribute VB_Name = "pelat2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ELEGXOS AFM FResize

Const F_FRAME2HEIGHT = 2532

Dim f_printer
Dim F_FORMAT_PEL_KOD      As String  '= FINDPARAMETROI(1, "PELAT1", "F_FORMAT_PEL_KOD", "0000", "ÿ«÷…¡  Ÿƒ… œ’ –≈À¡‘ŸÕ")

'-------------------------------------------------------------------------------------------------------------
'SUBCLASSING THE SSTab Control  By Mario Alberto Flores Gonzalez
'version 1.1
'February 10, 2004
'Feel free to use this source code as you wish in your projects
'                        sistec_de_juarez@hotmail.com
'Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
 Dim xWidth  As Long
 Dim xHeight As Long
 Dim xColor1 As Long
 Dim xColor2 As Long
 
 Dim gColor1 As Long
 Dim gColor2 As Long
 Dim gColor3 As Long
 Dim gColor4 As Long
 Dim gDirection    As Direction
'---------------------------------------------------------------------------------------------------------------

Dim ffW As Long, ffH As Long


Dim f_KENTRA_ADYNATISMATOS As Integer
Dim F_CAN_SYGXONEYSI As Integer

Dim F_MONOENERGOS As Integer '= Val(FINDPARAMETROI(1, "PAR1", "F_MONOENERGOS", "0", "ƒ≈…◊Õ≈… ÃœÕœ ‘œ’” ≈Õ≈—√œ’”=1 œÀœ’”=0")) '


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

Dim f_lab1                 As String

Dim f_lab2                 As String

Dim f_lab3                 As String

Dim f_lab4                 As String

Dim f_lab5                 As String

Dim f_lab6                 As String
Dim f_lab7                 As String



Dim f_pedioAFM                 As String
Dim f_LabAFM                 As String

Dim f_keim_Back_color As Long


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

Dim f_EKPTCAPTION          As String '≈ÍÙ˘ÛÁ % TÈÏÔÍ·Ù‹ÎÔ„ÔÚ


'Dim F_TRAINPERIODS As Integer

Dim F_LABEL_ADT As String

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



' Note that this example does not scale controls within
' other controls. For example, if you place a TextBox
' inside a PictureBox, the TextBox's position and
' dimensions are relative to the PictureBox not the
' form. This example does not handle this case.
'
' It also does not adjust fonts, stretch pictures, etc.
'
' For a more complete example, see the Stretchable
' control in my book "Custom Controls Library".
'
'     http://www.vb-helper.com/ccl.htm

Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single

Private Sub ResizeControls()
' Arrange the controls for the new size.

Dim I As Integer
Dim Ctl As Control
Dim X_Scale As Single
Dim Y_Scale As Single

' Don't bother if we are minimized.
If WindowState = vbMinimized Then Exit Sub

' Get the form's current scale factors.
X_Scale = ScaleWidth / m_FormWid
Y_Scale = ScaleHeight / m_FormHgt

' Position the controls.
I = 1
For Each Ctl In Controls
    With m_ControlPositions(I)
        If TypeOf Ctl Is LINE Then
            Ctl.X1 = X_Scale * .Left
            Ctl.Y1 = Y_Scale * .Top
            Ctl.X2 = Ctl.X1 + X_Scale * .Width
            Ctl.Y2 = Ctl.Y1 + Y_Scale * .Height
        Else
            Ctl.Left = X_Scale * .Left
            Ctl.Top = Y_Scale * .Top
            Ctl.Width = X_Scale * .Width
            If Not (TypeOf Ctl Is ComboBox) Then
                ' Cannot change height of ComboBoxes.
                Ctl.Height = Y_Scale * .Height
            End If
            On Error Resume Next
            Ctl.Font.Size = Y_Scale * .FontSize
            On Error GoTo 0
        End If
    End With
    I = I + 1
Next

End Sub





Private Sub SaveSizes()
' Save the form's and controls' dimensions.

Dim I As Integer
Dim Ctl As Control



m_FormWid = ScaleWidth
m_FormHgt = ScaleHeight

' Save the controls' positions and sizes.
ReDim m_ControlPositions(1 To Controls.Count)
I = 1
For Each Ctl In Controls
    With m_ControlPositions(I)
        If TypeOf Ctl Is LINE Then
            .Left = Ctl.X1
            .Top = Ctl.Y1
            .Width = Ctl.X2 - Ctl.X1
            .Height = Ctl.Y2 - Ctl.Y1
        Else
            .Left = Ctl.Left
            .Top = Ctl.Top
            .Width = Ctl.Width
            .Height = Ctl.Height
            On Error Resume Next
            .FontSize = Ctl.Font.Size
            On Error GoTo 0
        End If
    End With
    I = I + 1
Next

' Save the form's size.
m_FormWid = ScaleWidth
m_FormHgt = ScaleHeight

End Sub



















Private Sub CmdAllaghKodikou_Click()


        '<EhHeader>
        On Error Resume Next 'GoTo Command4_Click_Err

        '</EhHeader>
        Dim neos As String, RECS As Integer

        Dim buff As String

100     buff = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)
        'BUFF = IIf(PELPROM.ListIndex = 0, "e", "r")
110     neos = InputBox("ƒ˛ÛÂ ÙÔÌ Ì›Ô Í˘‰ÈÍ¸ ")

        If Len(Trim(neos)) = 0 Then
        
            MsgBox "ƒ≈Õ ’–¡—◊≈… Õ≈œ”  Ÿƒ… œ”. ƒ…¡ œ–« ƒ…¡ƒ… ¡”…¡”"
            Exit Sub
        End If
        


        Dim R As New ADODB.Recordset

120     R.Open "select count(*),EPO from PEL WHERE KOD='" + neos + "' and EIDOS='" + buff + "'  GROUP BY EPO", Gdb, adOpenForwardOnly, adLockReadOnly

        Dim ANS As Integer, OKSYGX As Integer

130     OKSYGX = 0

        If R.EOF Then
            OKSYGX = 0
        Else
        
        
        If F_CAN_SYGXONEYSI = 0 Then
           ANS = MsgBox("’–¡—◊≈… «ƒ« O  Ÿƒ… œ” " + neos + " " + R(1))
           Exit Sub
        End If
        
        
        

140         If R(0) = 1 Then    ' ¬—≈»« ≈  ¡… ¡ÀÀœ”
150             ANS = MsgBox("’–¡—◊≈… «ƒ« O " + R(1) + " Õ¡ ”’√◊ŸÕ≈’»œ’Õ; ", vbYesNo)

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
        
         On Error GoTo Command4_Click_Err
        ' Gdb.BeginTrans
200      Gdb.BeginTrans

On Error GoTo CloseT
If OKSYGX = 0 Then
210         Gdb.Execute "UPDATE PEL SET KOD='" + neos + "' WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'", RECS
220         MsgBox "¡—◊≈…œ –≈À¡‘ŸÕ " + Chr(13) + "≈Õ«Ã≈—ø»« ¡Õ   " + Format(RECS, "#####")
        Else
230         Gdb.Execute "delete from  PEL  WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'", RECS

240         If RECS > 0 Then
250             MsgBox "œÀ≈” œ…  …Õ«”≈…” ≈◊œ’Õ Ã≈‘¡÷≈—»≈… ”‘œÕ  Ÿƒ… œ " + neos
            Else
260             MsgBox "« ƒ…¡ƒ… ¡”…¡ ƒ…≈ œ–«"
                Gdb.RollbackTrans
                Exit Sub

            End If
        End If

270     Gdb.Execute "UPDATE EGG SET KOD='" + neos + "' WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'", RECS
280     MsgBox "¡—◊≈…œ  …Õ«”≈ŸÕ ¡–œ»« «” " + Chr(13) + "≈Õ«Ã≈—ø»« ¡Õ   " + Format(RECS, "#####")
Dim dum As Integer
'dum = "a"
        'Gdb.Execute "UPDATE TIM SET PELKOD='" + neos + "' WHERE EIDOS='" + BUFF + "' AND PELKOD='" + Text1(0).Text + "'", RECS
290     MsgBox "¡—◊≈…œ TIMOÀœ√…ŸÕ " + Chr(13) + "≈Õ«Ã≈—ø»« ¡Õ   " + Format(RECS, "#####")

300     Gdb.Execute "UPDATE TIM SET KPE='" + neos + "' WHERE EIDOS='" + buff + "' AND KPE='" + Text1(0).Text + "'", RECS

310     MsgBox "¡—◊≈…œ TIMOÀœ√…ŸÕ " + Chr(13) + "≈Õ«Ã≈—ø»« ¡Õ   " + Format(RECS, "#####")

320     Gdb.Execute "UPDATE EGGTIM SET PELKOD='" + neos + "' WHERE EIDOS='" + buff + "' AND PELKOD='" + Text1(0).Text + "'", RECS
330     MsgBox "¡—◊≈…œ  …Õ«”≈ŸÕ ¡–œ»« «” " + Chr(13) + "≈Õ«Ã≈—ø»« ¡Õ   " + Format(RECS, "#####")

        Gdb.CommitTrans
        


        ' Else
        '   MsgBox "’‹Ò˜ÂÈ Ô Í˘‰ÈÍ¸Ú " + neos
        'End If
        '<EhFooter>
        Exit Sub
        
CloseT:
        Gdb.RollbackTrans
        
        Exit Sub
        
        
        

Command4_Click_Err:
         ' ƒ≈Õ Ã–œ—≈… Õ¡ Œ≈ …Õ«”≈… ‘œ ‘—¡Õ”¡ÿ‘…œÕ
         
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command4_Click " & "at line " & Erl

         Gdb.Close
         Gdb.Open gConnect
         Gdb.BeginTrans
        Resume Next

        '</EhFooter>



End Sub

Private Sub cmdTIMOKAT_Click()

Dim F As New apot4
F.TIMOK.Text = Text1(8).Text
F.TIMOK_LostFocus1
F.SHOW






End Sub

Private Sub cmd√√_Click()

On Error Resume Next

   If Len(Trim(Text1(5).Text)) = 9 Then
   
   
     ' pelat1.neo_FIND_AFM Trim(Text1(5).Text), 1
      'Exit Sub
   
   
   
       'OK       Form1.parametros.Caption = "http://lagakis.gr/greek-vat/web-vat-details.php?myafm=028783755&afm=" + Trim(Text1(5).Text) ' 028783755"
      Dim u As String, p As String, AFM As String
      
      
     Dim f_EYRESHAFM_USER, f_EYRESHAFM_PWD, f_EYRESHAFM_AFM As String
      
      
    f_EYRESHAFM_USER = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_USER", "LAGAKIS23041960", "◊—«”‘«” ≈’—≈”«” ¡÷Ã")
    f_EYRESHAFM_PWD = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_PWD", "d389833921", " Ÿƒ… œ” ≈’—≈”«” ¡÷Ã")
    f_EYRESHAFM_AFM = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_AFM", "028783755", "¡÷Ã ≈–…◊≈…—«”«")
       
    Dim ANS As Integer
    ANS = MsgBox("Õ· ·ÌÙÈÍ·Ù·ÛÙ·ËÔ˝Ì Ù· ÛÙÔÈ˜Âﬂ· ·¸ ÙÁÌ √ÂÌÈÍﬁ √Ò·ÏÏ·ÙÂﬂ· –ÎÁÒÔˆ.”ıÛÙÁÏ‹Ù˘Ì;", vbYesNo)
    If ANS = vbNo Then
       Exit Sub
    End If
      
      
      
    u = Trim(f_EYRESHAFM_USER)   ' = "LAGAKIS23041960"  ' Ww367868u916
    p = Trim(f_EYRESHAFM_PWD)  '  = "d389833921"   '080353
    AFM = Trim(f_EYRESHAFM_AFM)    ' = "028783755"   '051103283
      
       

      PutGGPS u
      
 
      Dim EPO As String, DIE As String, EPA As String, pol As String, tk As String, doy As String
      Dim c As String
      c = Text1(5).Text
      Dim ENERGOS As Integer
      ENERGOS = 1
      Par7MyData.requestafm2 u, p, c, EPO, DIE, EPA, pol, tk, doy, ENERGOS
      
      
    Dim sos As Integer
    sos = 0
      
     If sos = 1 Or Len(Trim(EPO)) = 0 Then
       ' pelat1.Text1(1) = "ÏÁ Â„ÍıÒÔ ¡÷Ã"
     Else
    
       Me.Text1(1) = Replace(EPO, """", "")
       Me.Text1(2) = Replace(DIE, """", "")
       Me.Text1(3) = Replace(pol, """", "")
       Me.Text1(4) = Replace(EPA, """", "")
       Me.Text1(7) = Replace(doy, """", "")
       Me.Text1(7) = Replace(Me.Text1(7), "'", "`")
       Me.Text1(13) = Replace(tk, """", "")
   
    
     End If
      
      
      
      
      
      
      
      
      
      
      
      
      
    
   Else
       MsgBox "”’Ã–À«—Ÿ”‘≈ ‘œ ¡÷Ã "
       Text1(5).SetFocus
       
   
   End If




























End Sub

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


Sub EXODOS()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
100     If Text1(1).Enabled Then    '¡Õ ≈…”¡… ”‘«Õ ƒ…œ—»Ÿ”«
            'On Error Resume Next
            '     Data1.Recordset.UpdateBatch

        End If

        ' If Data1.Recordset.RecordCount = 1 Then

        On Error Resume Next

If Frame2.Visible = False Then
    Frame1.Visible = False
    Frame2.Visible = True
    Exit Sub
    
    
End If






'Dim ANS As Integer: ANS = MsgBox("Õ· Û˘ËÔ˝Ì ÔÈ ·ÎÎ·„›Ú;", vbYesNo)
'If ANS = vbYes Then
110     Data1.Recordset.Move 0  'ayto einai gia to text kitrino

'End If


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

100     sFileName = File1.Path + "\" + Listview1.SelectedItem
110     Kill sFileName
   ' mF = f_files + "\MERCVB\files\" + mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1) + Text1(0)
          SORT_LISTBOX File1.Path
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

Private Sub cmdDelete_Click()

        '<EhHeader>
        On Error GoTo cmdDelete_Click_Err

        '</EhHeader>

        Dim ANS  As Integer

        Dim neos As String, RECS As Integer

        Dim buff As String

100     buff = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)

        '  neos = InputBox("ƒ˛ÛÂ ÙÔÌ Ì›Ô Í˘‰ÈÍ¸ ")
        Dim R As New ADODB.Recordset

110     R.Open "select count(*) from EGG WHERE KOD='" + Data1.Recordset("KOD") + "' and EIDOS='" + buff + "'", Gdb, adOpenForwardOnly, adLockReadOnly

120     If R(0) > 0 Then
130         MsgBox "¡‰˝Ì·ÙÁ Á ‰È·„Ò·ˆﬁ „È·Ùﬂ ı‹Ò˜ÔıÌ ÍÈÌﬁÛÂÈÚ"
140         R.Close

            Exit Sub

        End If

        On Error Resume Next

150     ANS = MsgBox("–ÒÔÛÔ˜ﬁ Ë· ‰È·„Ò·ˆÂﬂ Ô ÎÔ„·ÒÈ·ÛÏ¸Ú " + Data1.Recordset("epo"), vbYesNo)

160     If ANS = vbYes Then
170         Data1.Recordset.delete

180         MHNYMA2.Timer1.Interval = 2000
190         MHNYMA2.Label2.Caption = "« Â„„Ò·ˆﬁ " + Text1(1).Text + " ‰ÈÂ„Ò‹ˆÁ."
200         MHNYMA2.Label1.Caption = "ƒÈ·„Ò·ˆﬁ Â„„Ò·ˆﬁÚ"
210         MHNYMA2.SHOW 1

            'MsgBox "ƒÈÂ„Ò‹ˆÁ"
        End If

        '<EhFooter>
        Exit Sub

cmdDelete_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub cmdDiortHosi()

        '<EhHeader>
       ' On Error GoTo cmdDiortosi_Click_Err

        '</EhHeader>
        Dim k As Integer
        
'        cmdDelete.Enabled = True
'        cmdKartella.Enabled = True
'        cmdEtiketta.Enabled = True
'        cmdChange.Enabled = True
        
        Text2(0).Enabled = True
        Text2(1).Enabled = True
        Text2(2).Enabled = True
        DataGrid1.Enabled = True
        
        
        
        
        
        
120         For k = 1 To 15
               
130             Text1(k).Enabled = True
140             Text1(k).ForeColor = vbBlack
150             If k <> 11 Then
                     Text1(k).MaxLength = Data1.Recordset(Text1(k).DataField).DefinedSize
                End If
            Next
            
            If Data1.Recordset("ENERGOS") = 1 Then ' Or Data1.Recordset("ENERGOS") = -1 Then
                      ' Check1.Value = vbChecked
                       chkEnergos.Value = vbChecked
            Else
                       'Check1.Value = vbUnchecked
                       chkEnergos.Value = vbUnchecked
            End If
            
            
             Text1(11).Text = nNull(Data1.Recordset("AYP"))

160       Frame1.Visible = True
            Text1(1).SetFocus
170         msEditor1.Enabled = True
180       '  cmdNew.Enabled = True
        
        
        
        
        
        
         TRAINING.Visible = False
             'sql = "CREATE TABLE TRAINPERIODS(KOD VARCHAR(12),[ENARXH] [date] NULL,[LHXH] [date] NULL,MHNAS VARCHAR(20),[SYNEDRIES] [int] ) "
'If F_TRAINPERIODS = 1 Then
'ADODC1.ConnectionString = gConnect
'
'          ADODC1.RecordSource = "select CENARXH,CLHXH,MHNAS,SYNEDRIES,ID from TRAINPERIODS where  KOD='" + Text1(0).Text + "'  ;"
'          ADODC1.Refresh
'          If ADODC1.Recordset.RecordCount = 0 Then
'              Dim KSYN As Integer
'              For k = 1 To 12
'                  Gdb.Execute "INSERT INTO TRAINPERIODS (KOD) VALUES ('" + Text1(0).Text + "')"
'              Next
'
'          End If
'          TRAINING.Visible = True
'
'          ADODC1.Refresh
'
'
'
'End If





        
        
        
        
        
        
        
        
        
        
        
        
End Sub
      
        
Sub cmdUpdate()
        On Error GoTo ISLOCKED
        Dim k As Long
        Dim MYP As String: MYP = Text1(11).Text

100    ' If cmdDiortosi.Caption = "ƒÈ¸ÒË˘ÛÁ" Then
'110         cmdDiortosi.Caption = "≈ÌÁÏ›Ò˘ÛÁ"
'
'120         For k = 1 To 15
'130             Text1(k).Enabled = True
'140             Text1(k).ForeColor = vbBlack
'150             Text1(k).MaxLength = DATA1.Recordset(Text1(k).DataField).DefinedSize
'            Next
'
'160         Text1(1).SetFocus
'170         msEditor1.Enabled = True
'180         cmdNew.Enabled = True
       ' Else
190      '   cmdDiortosi.Caption = "ƒÈ¸ÒË˘ÛÁ"
            '  DATA1.Recordset.UpdateBatch adAffectAll

            tComb1.Text = Comb1.Text
            tComb2.Text = Comb2.Text
            tComb3.Text = Comb3.Text
            tComb4.Text = Comb4.Text
            tComb5.Text = Comb5.Text
Dim mmID As String
mmID = str(Data1.Recordset("id"))
200         Data1.Recordset.Move 0

Gdb.Execute "update PEL SET ENERGOS=" + IIf(chkEnergos.Value = vbChecked, "1", "0") + " WHERE ID=" + mmID

Dim buff As String

     buff = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)

If InStr(MYP, ",") > 0 And InStr(MYP, ".") > 0 Then
     If InStr(MYP, ",") < InStr(MYP, ".") Then    'XILIADES ME ,
          MYP = Replace(MYP, ",", "")
     Else
          MYP = Replace(MYP, ".", "")
     End If
   

End If
MYP = Replace(MYP, ",", ".")

 Gdb.Execute "UPDATE PEL SET AYP=" + MYP + " WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'"
Dim CCB As String
CCB = "0"
'If Check1.Value = vbChecked Then CCB = "1"
'Gdb.Execute "UPDATE PEL SET ENERGOS=" + CCB + " WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'"

210         MHNYMA2.Timer1.Interval = 2000
220         MHNYMA2.Label2.Caption = "« Â„„Ò·ˆﬁ " + Text1(1).Text + " ·ÔËÁÍÂ˝ÙÁÍÂ ÏÂ ÂÈÙı˜ﬂ·"
230         MHNYMA2.SHOW 1

            If f_KENTRA_ADYNATISMATOS = 1 Then

                On Error Resume Next

                Dim R5   As New ADODB.Recordset

                Dim ANS5 As Integer

                ' ANOIGO PROSOPO
                If kentroAdyn.ListIndex = 0 Then ' AN SVMA
                    R5.Open "SELECT * FROM PEL WHERE KOD='10" + Format(Val(Text1(0).Text), "0000") + "'", Gdb, adOpenDynamic, adLockOptimistic

                    If R5.EOF Then
                       
                        ANS5 = MsgBox("Ì· ·ÌÔÈ˜ËÂﬂ Í·È ÛÙÔ Ò¸Û˘Ô;", vbYesNo)

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
                       
                        ANS5 = MsgBox("Ì· ·ÌÔÈ˜ËÂﬂ Í·È ÛÙÔ ”øÃ¡;", vbYesNo)

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
       ' End If

        ' KEIM1.BackColor = vbGreen
        ' keim2.BackColor = vbGreen
    
250     keim(2).Refresh





'  TRAINING.Visible = False
'             'sql = "CREATE TABLE TRAINPERIODS(KOD VARCHAR(12),[ENARXH] [date] NULL,[LHXH] [date] NULL,MHNAS VARCHAR(20),[SYNEDRIES] [int] ) "
'If F_TRAINPERIODS = 1 Then
'Adodc1.ConnectionString = gConnect
'
'          Adodc1.RecordSource = "select ENARXH,LHXH,MHNAS,SYNEDRIES from TRAINPERIODS where  KOD='" + Text1(0).Text + "'  AND EIDOS='" + "e" + "' ;"
'          Adodc1.Refresh
'          If Adodc1.Recordset.RecordCount = 0 Then
'              Dim KSYN As Integer
'              For K = 1 To 12
'                  Gdb.Execute "INSERT INTO TRAINPERIODS (KOD) VALUES ('" + Text1(0).Text + "')"
'              Next
'
'          End If
'          TRAINING.Visible = True
'
'          Adodc1.Refresh
'
'
'
'End If
'
'
'
'

























        Exit Sub

ISLOCKED:
260    ' MsgBox "« Â„„Ò·ˆﬁ ÂÌÁÏÂÒ˛ÌÂÙ·È ·¸ ‹ÎÎÔ ˜ÒﬁÛÙÁ"

270     Resume Next

        Exit Sub

        '<EhFooter>
        Exit Sub

cmdDiortosi_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub cmdChange_Click()----ÔÎ‰

'        '<EhHeader>
'        On Error GoTo cmdChange_Click_Err
'
'        '</EhHeader>
'        Dim neos As String, RECS As Integer
'
'        Dim buff As String
'
'100     buff = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)
'        'BUFF = IIf(PELPROM.ListIndex = 0, "e", "r")
'110     neos = InputBox("ƒ˛ÛÂ ÙÔÌ Ì›Ô Í˘‰ÈÍ¸ ")
'
'        Dim R As New ADODB.Recordset
'
'120     R.Open "select count(*),EPO from PEL WHERE KOD='" + neos + "' and EIDOS='" + buff + "'  GROUP BY EPO", Gdb, adOpenForwardOnly, adLockReadOnly
'
'        Dim ANS As Integer, OKSYGX As Integer
'
'130     OKSYGX = 0
'
'        If R.EOF Then
'            OKSYGX = 0
'        Else
'
'
'        If F_CAN_SYGXONEYSI = 0 Then
'           ANS = MsgBox("’–¡—◊≈… «ƒ« O  Ÿƒ… œ” " + neos + " " + R(1))
'           Exit Sub
'        End If
'
'
'
'
'140         If R(0) = 1 Then    ' ¬—≈»« ≈  ¡… ¡ÀÀœ”
'150             ANS = MsgBox("’–¡—◊≈… «ƒ« O " + R(1) + " Õ¡ ”’√◊ŸÕ≈’»œ’Õ; ", vbYesNo)
'
'160             If ANS = vbYes Then
'170                 OKSYGX = 1
'                Else
'180                 OKSYGX = 0
'
'                    Exit Sub
'
'                End If
'
'            Else
'190             OKSYGX = 0
'            End If
'        End If
'
'200     If OKSYGX = 0 Then
'210         Gdb.Execute "UPDATE PEL SET KOD='" + neos + "' WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'", RECS
'220         MsgBox "¡—◊≈…œ –≈À¡‘ŸÕ " + Chr(13) + "≈Õ«Ã≈—ø»« ¡Õ   " + Format(RECS, "#####")
'        Else
'230         Gdb.Execute "delete from  PEL  WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'", RECS
'
'240         If RECS > 0 Then
'250             MsgBox "œÀ≈” œ…  …Õ«”≈…” ≈◊œ’Õ Ã≈‘¡÷≈—»≈… ”‘œÕ  Ÿƒ… œ " + neos
'            Else
'260             MsgBox "« ƒ…¡ƒ… ¡”…¡ ƒ…≈ œ–«"
'
'                Exit Sub
'
'            End If
'        End If
'
'270     Gdb.Execute "UPDATE EGG SET KOD='" + neos + "' WHERE EIDOS='" + buff + "' AND KOD='" + Text1(0).Text + "'", RECS
'280     MsgBox "¡—◊≈…œ  …Õ«”≈ŸÕ ¡–œ»« «” " + Chr(13) + "≈Õ«Ã≈—ø»« ¡Õ   " + Format(RECS, "#####")
'
'        'Gdb.Execute "UPDATE TIM SET PELKOD='" + neos + "' WHERE EIDOS='" + BUFF + "' AND PELKOD='" + Text1(0).Text + "'", RECS
'290     MsgBox "¡—◊≈…œ TIMOÀœ√…ŸÕ " + Chr(13) + "≈Õ«Ã≈—ø»« ¡Õ   " + Format(RECS, "#####")
'
'300     Gdb.Execute "UPDATE TIM SET KPE='" + neos + "' WHERE EIDOS='" + buff + "' AND KPE='" + Text1(0).Text + "'", RECS
'
'310     MsgBox "¡—◊≈…œ TIMOÀœ√…ŸÕ " + Chr(13) + "≈Õ«Ã≈—ø»« ¡Õ   " + Format(RECS, "#####")
'
'320     Gdb.Execute "UPDATE EGGTIM SET PELKOD='" + neos + "' WHERE EIDOS='" + buff + "' AND PELKOD='" + Text1(0).Text + "'", RECS
'330     MsgBox "¡—◊≈…œ  …Õ«”≈ŸÕ ¡–œ»« «” " + Chr(13) + "≈Õ«Ã≈—ø»« ¡Õ   " + Format(RECS, "#####")
'
'        ' Else
'        '   MsgBox "’‹Ò˜ÂÈ Ô Í˘‰ÈÍ¸Ú " + neos
'        'End If
'        '<EhFooter>
'        Exit Sub
'
'cmdChange_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.pelat2.Command4_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command4_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>

'End Sub

Private Sub cmdKartela_Click()

        '<EhHeader>
        On Error GoTo cmdKartela_Click_Err
        
        
        
        
        'ÂÎÂ„˜˘ ·Ì Â˜ÂÈ ÙÔ ‰ÈÍ·È˘Ï· Ì· ·ÌÔÈÓÂÈ ÙÁÌ Í·ÒÙÂÎÎ·
         Dim r1 As New ADODB.Recordset, M_ENABLE As Boolean
         r1.Open "SELECT * FROM ACTIVE_MENUS WHERE ID=5 and  USER_ID=" + str(gUserId), Gdb, adOpenDynamic, adLockOptimistic
         If r1("ENABLED") = 1 Then
             M_ENABLE = True
         Else
             M_ENABLE = False
             Exit Sub
         End If
         
                    
        

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

cmdKartela_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>
        
        
        
        
        
' Ù· ·Ò·Í·Ù˘ ÂﬂÌ·È ·¸ ÙÁÌ ÙÈÏÔÎ¸„ÁÛÁ
        
'        <EhHeader>
'        On Error GoTo Command3_Click_Err
'
'        </EhHeader>
'        Dim fpelat4 As New pelat4
'
'        If Index = 0 Then
'100     fpelat4.SHOW
'110     fpelat4.Text2(0).Text = Text2(0).Text
'
'120     If POLAGOR.ListIndex = 0 Then
'130         fpelat4.PELPROM.ListIndex = 1
'        Else
'
'            If f_KENTRA_ADYNATISMATOS = 1 Then
'                If KentroAdyn.ListIndex = 1 Then
'                    fpelat4.KentroAdyn.ListIndex = 1
'                End If
'            End If
'
'140         fpelat4.PELPROM.ListIndex = 0
'            fpelat4.Option1(1).Value = True
'        End If
'
'150     fpelat4.DATA1.ConnectionString = gConnect
'160     fpelat4.Refresh
'170     gApoMenu = False
'
'        pelat4.Text2(0).SetFocus
'
'        fpelat4.SHOW
'180     fpelat4.Text2(0).SetFocus
'190     SendKeys "+{ENTER}"
'
'        Dim DUM
'
'200     DUM = fpelat4.KARTELLA(gDir)
'        End If
'210     Set fpelat4 = Nothing
'
'
'

End Sub

Private Sub cmdEtiketta_Click()
Dim DB2 As Database

   Set DB2 = OpenDatabase("c:\mercvb\reports\reports.mdb")



DB2.Execute "UPDATE MEM SET FPO1='" + Trim(Data1.Recordset("EPO")) + "'"
DB2.Execute "UPDATE MEM SET FPO2='" + Trim(Data1.Recordset("DIE")) + "'"
DB2.Execute "UPDATE MEM SET FPO3='" + Trim(Data1.Recordset("POL")) + "'"
DB2.Execute "UPDATE MEM SET FPO4='" + Trim(Data1.Recordset("EPA")) + "'"
DB2.Execute "UPDATE MEM SET FPO5='" + Trim(Data1.Recordset("THL")) + "'"


'CrystalReport1.Refresh


' CrystalReport1.ReportFileName = F_FORMA1
 '               CrystalReport1.Action = 1
DB2.Close
DoEvents



CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\ETIKPEL.RPT"
CrystalReport1.Action = 1

Exit Sub




        ' ETIKETTA ME KEIMENO
        '<EhHeader>
        On Error GoTo cmdEtiketta_Click_Err

        '</EhHeader>

        Dim ll, apoSCANNER As Long

        Dim DB

        Dim PSIFIA As Integer

        Dim fores  As Long, k As Integer

        Dim X      As Printer

        Dim barc(200)

        Dim mONOMA, mBarcode, mERG, mLTI5

        Dim R  As New ADODB.Recordset

        Dim RD As DAO.Recordset

        On Error GoTo cmdEtiketta_Click_Err

100     mONOMA = Trim(Text1(1).Text)    'ONOMA
110     mBarcode = Text1(0).Text    '  ' BARCODE
120     mERG = Text1(6).Text    ' erg
130     mLTI5 = Format(Val(Text1(13).Text), "###0.00")    ' LTI5

        On Error Resume Next

140     For Each X In Printers

150         If InStr(X.DeviceName, f_printer) > 0 Then

                ' Set printer as system default.
160             Set Printer = X
                ' Stop looking for a printer  5200016000239

                Exit For

            End If

        Next

170     apoSCANNER = 1
        'If Scanner Then

        ' ÏÔı„Ôı‰ÁÚ
        'Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
        'Set RD = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM= 1234")
        'If RD.RecordCount = 0 Then
        '   MsgBox "‰ÂÌ ı‹Ò˜ÔıÌ Â„„Ò·ˆ›Ú"
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

240     fores = Val(InputBox("–¸ÛÂÚ ÂÙÙÈÍ›ÙÂÚ; ", "", 1))

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

280             If OSSA = 0 Then ' Õœ—Ã¡À ≈ ‘’–Ÿ”« ”≈ Ã… —≈” ≈‘… ≈‘≈”
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
390                 Printer.Print Tab(4 + Z); "THÀ."; Text1(6).Text;    ' THL
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
530                 Printer.Print Tab(4 + Z); Left(keim(2), 50)    'Text1(1).Text 'ONOMA

                     Dim jk As Integer: jk = 70
                     Dim aseira As String, bseira As String
                     aseira = keim(3).Text
                    If Len(Trim(keim(3).Text)) > 70 Then
                      jk = InStr(58, keim(3).Text, " ")
                      If jk = 0 Then jk = 70 'ÛÁÏ·ÈÌÂÈ ‰ÂÌ Â˜ÂÈ ÍÂÌ· ÛÙÔ ÙÂÎÔÚ
                      aseira = Left(keim(3).Text, jk)
                      bseira = mID$(keim(3).Text, 71, 70)
                      Printer.Print Tab(4 + Z); aseira;   ' DIE DATA1.Recordset("die")
                      Printer.Print Tab(4 + Z); bseira;
                    Else
                      Printer.Print Tab(4 + Z); aseira;   ' DIE DATA1.Recordset("die")
                    End If
                     
540                 'Printer.Print Tab(4 + Z); Left(keim3.Text, jk);   ' DIE DATA1.Recordset("die")
550                 Printer.Print Tab(4 + Z); keim(4).Text;   ' DIE
560                 Printer.Print Tab(4 + Z); keim(5).Text;   ' POL
570                 Printer.Print Tab(4 + Z); keim(6).Text;   ' POL
580                 Printer.Print Tab(4 + Z); "TEL."; keim(7).Text;   ' POL
590           '      Printer.Print Tab(4 + Z); "TEL."; Text1(6).Text;    ' THL
600                 Printer.EndDoc
                End If

610         Next k

620     Next ll

        '<EhFooter>
        Exit Sub

cmdEtiketta_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdNew_Click()

   
   

        '<EhHeader>
        On Error GoTo cmdNew_Click_Err

        '</EhHeader>
        Dim c As String

100    Dim buff As String

buff = IIf(Left(PELPROM, 1) = "1", "e", "r")

       c = pelat1.get_next_pel(buff, 0, F_FORMAT_PEL_KOD) ' get_next_pel(BUFF)
       
       c = Format(c, F_FORMAT_PEL_KOD)
       
       
       Dim R As New ADODB.Recordset
       R.Open "select COUNT(*) from PEL where  KOD='" + c + "'  AND EIDOS='" + buff + "' ;", Gdb, adOpenDynamic, adLockOptimistic
       
       If R(0) > 0 Then
            MsgBox "’–¡—◊≈… «ƒ« œ  Ÿƒ… œ”"
            Exit Sub
       End If
       
      ' Check1.Value = vbChecked
       
       
       

110     Gdb.Execute "insert into PEL (EIDOS,KOD) VALUES ('" + buff + "','" + c + "')"

120     Text2(0).Text = c
130     Text1(0).Text = c
140     Text2(0).SetFocus

150     Data1.RecordSource = "select *from PEL where  KOD='" + c + "'  AND EIDOS='" + buff + "' ;"
160     Data1.Refresh

170    ' cmdNew.Enabled = False
        '       f_New_Record = True
        '
        '
        '       For K = 0 To 11
        '          Text1(K).Enabled = True
        '          Text1(K).ForeColor = vbBlack
        '       Next

180   '  cmdDiortosi.Caption = "ƒÈ¸ÒË˘ÛÁ"
        'Command3_Click

190     DataGrid1.SetFocus
200     SendKeys "{enter}{enter}"

'        cmdDelete.Enabled = False
'        cmdKartela.Enabled = False
'        cmdEtiketta.Enabled = False
'        cmdChange.Enabled = False
        
        Text2(0).Enabled = False
        Text2(1).Enabled = False
        Text2(2).Enabled = False
        DataGrid1.Enabled = False

        '<EhFooter>
        Exit Sub

cmdNew_Click_Err:
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
        Dim mF As String

        '¡Õ ƒ≈Õ ƒ…¡À≈Œ¡ –≈À¡‘« Õ¡ ÷≈’√≈…
100     If Len(Trim(Text1(0))) = 0 Then

            Exit Sub

        End If

        '\\Server\data

110     CommonDialog1.ShowOpen
        'If PELPROM.ListIndex = 0 Then
120     mF = f_files + "\MERCVB\files\" + mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1) + Text1(0)
        ' Else
        '   mF = f_files + "\MERCVB\files\r" + Text1(0)
        ' End If

        On Error Resume Next

130     If Len(Dir(f_files, vbDirectory)) = 0 Then
140         MkDir f_files + "\MERCVB\files"
        End If

150     If Len(Dir(mF, vbDirectory)) = 0 Then
160         MkDir mF
        End If

        Dim mf2 As Integer, files2 As String

        On Error Resume Next

170     mf2 = InStrRev(CommonDialog1.FILENAME, "\")
180     files2 = mID(CommonDialog1.FILENAME, mf2 + 1, Len(CommonDialog1.FILENAME) - mf2)

190     FileCopy CommonDialog1.FILENAME, mF + "\" + files2
        'MILSEC 1000
        '       File1.Refresh

      '  REFRESH_FILES
        ' mF = f_files + "\MERCVB\files\" + mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1) + Text1(0)
          SORT_LISTBOX mF

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

        Dim mF As String
        
100     N = PELPROM.ListIndex

110     If N < 0 Then N = 0
120     mF = UCase(f_files + "\MERCVB\files\" + mID(PELPROM_STRING, N + 1, 1) + Text1(0))

        'Else
        '   mF = UCase(f_files + "\MERCVB\files\r" + Text1(0))
        'End If
        On Error Resume Next

130     If Len(Dir(mF, vbDirectory)) > 0 Then
140         File1.Path = mF
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
'  '¡Õ ƒ≈Õ ƒ…¡À≈Œ¡ –≈À¡‘« Õ¡ ÷≈’√≈…
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
           ' cmdDiortHosi
           ' MDataGrid1_RowColChange
120        ' cmdDiortosi.SetFocus

       
       cmdDiortHosi
       Frame2.Visible = False
       Frame1.Top = Toolbar1.Height
       
       Toolbar1.Buttons.Item("cmdDiortosi").Value = tbrPressed
       Toolbar1.Buttons.Item("cmdNew").Enabled = False ' vbRed
      
       Toolbar1.Buttons.Item("cmdKartela").Enabled = False
        Toolbar1.Buttons.Item("cmdDelete").Enabled = False
        Toolbar1.Buttons.Item("cmdEtiketa").Enabled = False
       ' cmdEtiketa
        Toolbar1.Buttons.Item("cmdChange").Enabled = False
        
       Toolbar1.Buttons.Item("cmdUpdate").Enabled = True
       Toolbar1.Buttons.Item("cmdKartela").Enabled = True
        'cmdKartela





















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
       
       
       
       
       Toolbar1.Buttons.Item("cmdDiortosi").Enabled = True
       Toolbar1.Buttons.Item("cmdKartela").Enabled = True
       Toolbar1.Buttons.Item("cmdDelete").Enabled = True
       Toolbar1.Buttons.Item("cmdEtiketa").Enabled = True
       Toolbar1.Buttons.Item("cmdChange").Enabled = True


    
        '<EhHeader>
        On Error GoTo DataGrid1_RowColChange_Err

        '</EhHeader>

        Dim mF As String

        'If Option1(1).value = True Then
        On Error Resume Next

        Dim N As Integer

100     N = PELPROM.ListIndex

110     If N < 0 Then N = 0
120     mF = UCase(f_files + "\MERCVB\files\" + mID(PELPROM_STRING, N + 1, 1) + Text1(0))

        'Else
        '   mF = UCase(f_files + "\MERCVB\files\r" + Text1(0))
        'End If
        On Error Resume Next

130     If Len(Dir(mF, vbDirectory)) > 0 Then
140         File1.Path = mF
        Else
150         File1.Path = "c:\mercvb\files\dum"
        End If

160     File1.Refresh


        SORT_LISTBOX mF



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


 
     'If gXEIRISTHS = 1 Then
      ' If f_KENTRA_ADYNATISMATOS = 1 Then
               '  msEditor1.Visible = False
      ' End If
    ' End If
 
    
                 
             
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



Sub locForm_Resize(frm As Form, _
                        fontHeight As Single, _
                        WW As Long, _
                        HH As Long, _
                        fh, _
                        fw, _
                        ft, _
                        fl)

    '<EhHeader>
    On Error Resume Next


Dim x_size As Double



    '</EhHeader>
    Dim k As Long

    Dim nW, nH

    nW = frm.Width
    nH = frm.Height

    On Error Resume Next

    'If ffproto < 2 Then
    '   ffproto = ffproto + 1
    'If Me.width > ffW Then
    'Else
    '    ffW = Me.width
    '     ffH = Me.height
    'End If
    'Else
    Dim ccc As String
    
    
    
   For k = 0 To frm.Controls.Count - 1
    
     ccc = UCase(frm.Controls(k).Container.Name)
    
    If ccc = "CCCCCCCCCFRAME1" Then
   ' ElseIf ccc = "FRAME2" Then
    ElseIf ccc = "SSTAB1" Then   ' If x.Container.Name = "SSTab1" Then
        ccc = ccc
    Else
    
        frm.Controls(k).Width = fw(k) * (nW / WW)
        ' On Error Resume Next
        frm.Controls(k).Height = fh(k) * (nH / HH)
        frm.Controls(k).Top = ft(k) * (nH / HH)    ' ORIG10830)
        frm.Controls(k).Left = fl(k) * (nW / WW)    '13860)

        x_size = nW / WW

        '   x_size = ((nH / HH) + (nW / WW)) / 2
        If x_size > 0 Then
            'On Error GoTo 0
            If x_size > 1.5 Then
               frm.Controls(k).Font.Size = 10 ' SetFontSize(fontHeight)
            ElseIf x_size < 1 Then
                frm.Controls(k).Font.Size = 7 ' SetFontSize(fontHeight)
            End If
            'On Error Resume Next
            
        End If
        
        
     End If

        'frm.Controls(k).Name
   Next
   
   
   If Frame2.Visible = True Then
      Frame2.Top = 720
   Else
      Frame1.Top = 720
   End If
   

    'End If

End Sub










Sub SORT_LISTBOX(ByVal mF As String)
'============================================================================================

Dim fdate, fsize, SearchName, F, itmx, thisExtn, ddate

Dim GFOLDER As String

GFOLDER = mF '  "c:\mercvb"
SearchName = GFOLDER & "\*.*"
F = Dir(SearchName, vbReadOnly)
Listview1.ListItems.Clear

Do While F <> ""
            
            
            
            fdate = FileDateTime(GFOLDER & "\" & F)
            
    fsize = FileLen(GFOLDER & "\" & F)
    Set itmx = Listview1.ListItems.ADD(, , F)
     ' itmx.SubItems(1) = UCase(thisExtn)
      itmx.SubItems(1) = Format(fdate, "YYYY/MM/DD") ' HhNn")  'fdate
      itmx.SubItems(2) = Round(fsize / 1024, 1)
      'itmx.SubItems(3) = Format(fsize, "0000000000")
      'itmx.SubItems(4) = Format(fdate, "YYYYMMDDHhNnSs")
    Set itmx = Nothing
    
    F = Dir()
            If F = "" Then
               'Exit Do
            End If
               
Loop


'If FIRST_TIME = 1 Then
 '   FIRST_TIME = 0
 
  
 
 
 

'„È· Ì· ‚·ÎÂÈ Í·Ù· ˆËﬂÌÔıÛ· ÁÏÂÒÔÏÁÌÈ·
   Listview1.SortKey = 1

    Listview1.SortOrder = 1 'Abs(Not ListView1.SortOrder = 1)
    Listview1.Sorted = True



End Sub








Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then EXODOS

110     If KeyCode = vbKeyF10 Then
120         ut_rantebou.Label1 = Text1(1).Text

130         ut_rantebou.SHOW 1
        End If
        
        If f_KENTRA_ADYNATISMATOS = 1 Then

            If KeyCode = vbKeyF5 Then

                Dim a

                a = Chr(13) + "¡—…”.√œÕ¡‘œ=" + Chr(13) + "¡—…”.Ã«—œ”="
                a = a + Chr(13) + "–≈—…÷≈—≈…¡="
                a = a + Chr(13) + "ƒ≈Œ.√œÕ¡‘œ= " + Chr(13) + "ƒ≈Œ.Ã«—œ”="
                a = a + Chr(13) + " œ…À…¡=" + Chr(13) + "KGR="

                If f_KENTRA_ADYNATISMATOS = 1 Then
                    msEditor1.Text = msEditor1.Text + a
                End If
            End If
        
            If KeyCode = vbKeyF7 Then
                a = "”ıÏˆ.–ÔÛ¸ " + Format(Val(SYMFPOSO.Text), "####.00") + Chr(13)
                a = a + "”ıÏˆ.MÁ˜·ÌﬁÏ·Ù· " + Format(Val(mhxanhmata.Text), "####") + Chr(13)
                a = a + "«ÏÂÒ.”ıÏˆ." + Format(HMESYMF, "DD/MM/YYYY") + Chr(13)
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

F_MONOENERGOS = Val(FINDPARAMETROI(1, "PAR1", "F_MONOENERGOS", "0", "ƒ≈…◊Õ≈… ÃœÕœ ‘œ’” ≈Õ≈—√œ’”=1 œÀœ’”=0")) '




        '<EhHeader>
        On Error GoTo Form_Load_Err

     ffW = Me.Width
     ffH = Me.Height

        '</EhHeader>
100     If gApoMenu Then  ' ¡Õ ‘œ  ¡ÀŸ ¡–œ ‘«Õ  ¡—‘≈ÀÀ¡ Õ¡ Ã«Õ ¡ÀÀ¡∆≈…ƒ…¡”‘¡”≈…”
110        ' mForm_Load Me, fh, fw, ft, fl
          ' SaveSizes
        End If
 
        f_keim_Back_color = keim(1).BackColor
        
 
 
 
        Toolbar1.Buttons.Item("cmdDiortosi").Enabled = False
        Toolbar1.Buttons.Item("cmdKartela").Enabled = False
        Toolbar1.Buttons.Item("cmdDelete").Enabled = False
        Toolbar1.Buttons.Item("cmdEtiketa").Enabled = False
        Toolbar1.Buttons.Item("cmdUpdate").Enabled = False
        Toolbar1.Buttons.Item("cmdChange").Enabled = False
 
   
   f_printer = FINDPARAMETROI(1, "APOT2", "F_PRINTER", "SATO", "Ã≈—œ” ‘«” œÕœÃ¡”…¡” ‘œ’ ≈ ‘’–.≈‘… ≈‘‘ŸÕ")
   
   
'   Dim f_pedioAFM                 As String
'Dim f_LabAFM                 As String

      f_pedioAFM = FINDPARAMETROI(1, "PELAT2", "f_pedioAFM", "CH1", "–ÔÈ¸ Â‰ﬂÔ Ë· ·Ì·ÊÁÙ‹Ù·È Ï·Êﬂ ÏÂ ÙÔ ¡÷Ã")
      f_LabAFM = FINDPARAMETROI(1, "PELAT2", "f_LabAFM", "¡÷Ã /ÍÂÈÏ1", "TﬂÙÎÔÚ ÂÙÈÍ›Ù·Ú ¡÷Ã")
      'Label1(2).Caption = f_LabAFM
     F_FORMAT_PEL_KOD = FINDPARAMETROI(1, "PELAT1", "F_FORMAT_PEL_KOD", "0000", "ÿ«÷…¡  Ÿƒ… œ’ –≈À¡‘ŸÕ")
     
 Frame1.Visible = False
 
' Frame2.Height = Me.Height - Toolbar1.Height
 'DataGrid1.Height = Frame2.Height - 400
 
 
 
 'GoTo 222
 
 'sstab1 xroma
    'GradientColor Style
    '--------------------------------------------------------------------------------

    'ElseIf Option3.Value = True Then
             
       ' If Option4.Value = True Then
           gDirection = cHorizontal
      '  Else
        '   gDirection = cVertical
       ' End If
        gColor1 = &HC0FFC0     'vbBlue '&HFFFFC0     ' vbBlue
        gColor2 = vbCyan      '&HFF0000
        SetStyle SSTab1.hwnd, cGradient  '//--- Set The Style of The SSTab
        SetGradientDir SSTab1.hwnd, gDirection '//--- Set The Gradient Direction
        SetGradientColor1 SSTab1.hwnd, gColor1  '//--- Asing new Gradient Color Start
        SetGradientColor2 SSTab1.hwnd, gColor2  '//--- Asing new Gradient Color End
 
     SSTabSubclass SSTab1.hwnd '//--- Begin SubClassing
     RedrawWindow SSTab1.hwnd, ByVal 0&, ByVal 0&, &H1
 
222
 
 
 
 
 
 
 
 
 
 
        ' El.mForm_Load Me

        F_CAN_SYGXONEYSI = Val(FINDPARAMETROI(1, "PELAT2", "F_CAN_SYGXONEYSI", "0", "ƒıÌ·Ù¸ÙÁÙ· Ûı„˜˛ÌÂıÛÁÚ Í˘‰ÈÍ˛Ì=1 œ◊…=0"))

        f_KENTRA_ADYNATISMATOS = Val(FINDPARAMETROI(99, "PELAT2", "f_KENTRA_ADYNATISMATOS", "0", "¡Ì ÂﬂÌ·È Í›ÌÙÒÔ ·‰ıÌ·ÙﬂÛÏ·ÙÔÚ=1"))

        If f_KENTRA_ADYNATISMATOS = 1 Then
            
            
        
 
           If gXEIRISTHS = 1 Then
     
                 msEditor1.Visible = False
      
           End If
     
            
            
            
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

        'f_checkPos = Val(FindParametroi(1,"PAR1", "F_CHECKPOS", "DELETE", "1=≈À≈√◊O” 2=¡–¡√œ—≈’”« 0=‘…–œ‘¡"))

130     PELPROM_STRING = "er"
140     PELPROM.ListIndex = 0

        Dim AA As Long
    
150     AA = GetCurrentTime()
    
160     f_TIMOKAT = Val(FINDPARAMETROI(1, "PELAT2", "F_TIMOKAT", "0", "0=ÂÍÙ.ÂÎ‹ÙÁ ‘ÈÏÔÍ·Ù. -1=ÙÈÏﬁÚ 1=ÂÍÙ˘ÛÁÚ 333=ÏÂ 3 ˜ÔÌ‰ÒÈÍ›Ú 9999=ÏÂ ÂÎ·ÙÁ & Âﬂ‰ÔÚ"))

        ' f_tabAP = Val(FindParametroi(1,"PELAT2", "F_TABAP", "3", "–ÒÔÂÈÎÂ„Ï›ÌÔ TAB "))
        ' SSTab1.Tab = IIf(f_tabAP > SSTab1.Tabs - 1 Or f_tabAP < 0, SSTab1.Tabs - 1, f_tabAP)

        Dim dum

170     dum = FINDPARAMETROI(1, "PELAT2", "F_KODPROM", "DELETE", "")
180     dum = FINDPARAMETROI(1, "PELAT2", "F_RAFI", "DELETE", "")
190     dum = FINDPARAMETROI(1, "PELAT2", "F_OIKO", "DELETE", "")
200     dum = FINDPARAMETROI(1, "PELAT2", "F_YPOOIK", "DELETE", "")
210     dum = FINDPARAMETROI(1, "PELAT2", "F_XTI", "DELETE", "")

220     dum = FINDPARAMETROI(1, "PELAT2", "F_BOHU", "DELETE", "")
230     dum = FINDPARAMETROI(1, "PELAT2", "F_LTI", "DELETE", "")
240     dum = FINDPARAMETROI(1, "PELAT2", "F_PROM", "DELETE", "")
250     dum = FINDPARAMETROI(1, "PELAT2", "F_STOK", "DELETE", "")

260     F_DIGpel = Val(FINDPARAMETROI(1, "PELAT2", "F_DIGPEL", "4", "¡—.ÿ«÷…ŸÕ  Ÿƒ… œ’ –≈À¡‘ŸÕ/–—œÃ«»≈’‘ŸÕ"))

        '\\Server\data

270     f_FirstTab = FINDPARAMETROI(1, "PELAT2", "FirstTab", "1", "–ÒoÂÈÎÂ„Ï›ÌÔ ‘ab ")

280     SSTab1.Tab = IIf(f_FirstTab > SSTab1.Tabs - 1 Or f_FirstTab < 0, SSTab1.Tabs - 1, f_FirstTab)

290     f_files = Trim(FINDPARAMETROI(1, "PELAT2", "F_FILES", "C:", "ƒﬂÛÍÔÚ ÏÂ ·Ò˜Âﬂ· office"))



Label1(2).Caption = f_LabAFM





Dim E2 As String
E2 = FINDPARAMETROI(2, "PELAT2", "e2", "----", String(80, "-"))
300     F_LHM1 = FINDPARAMETROI(2, "PELAT2", "F_lhm1", "«ÏÂÒ1", "≈ÙÈÍ›Ù· ÁÏÂÒ.1")
310     F_LHM2 = FINDPARAMETROI(2, "PELAT2", "F_lhm2", "«ÏÂÒ2", "≈ÙÈÍ›Ù· ÁÏÂÒ.2")
320     f_lHM3 = FINDPARAMETROI(2, "PELAT2", "F_lhm3", "«ÏÂÒ3", "≈ÙÈÍ›Ù· ÁÏÂÒ.3")
330     f_lHM4 = FINDPARAMETROI(2, "PELAT2", "F_lhm4", "«ÏÂÒ4", "≈ÙÈÍ›Ù· ÁÏÂÒ.4")
340     f_lHM5 = FINDPARAMETROI(2, "PELAT2", "F_lhm5", "«ÏÂÒ5", "≈ÙÈÍ›Ù· ÁÏÂÒ.5")
350     f_lHM6 = FINDPARAMETROI(2, "PELAT2", "F_lhm6", "«ÏÂÒ6", "≈ÙÈÍ›Ù· ÁÏÂÒ.6")

        Dim f_EKPTCAPTION As String '≈ÍÙ˘ÛÁ % TÈÏÔÍ·Ù‹ÎÔ„ÔÚ

360     f_EKPTCAPTION = FINDPARAMETROI(1, "PELAT2", "f_EKPTCAPTION", "≈ÍÙ˘ÛÁ %", "EÙÈÍÂÙ· ÂÍÙ˘ÛÁÚ  ")




F_LABEL_ADT = FINDPARAMETROI(1, "PELAT2", "F_LABEL_ADT", "¡ƒ‘", "EÙÈÍÂÙ· ¡ÒÈË.‘·ıÙ¸ÙÁÙ·Ú  ")

Label1(24).Caption = F_LABEL_ADT

     '   F_TRAINPERIODS = Val(FINDPARAMETROI(1, "PELAT2", "F_TRAINPERIODS", "0", "1=÷—œÕ‘…”‘«—…¡ Ã≈ –≈—…œƒœ’” 0=œ◊…  "))


370     Label1(11).Caption = f_EKPTCAPTION

380     lhm1.Caption = F_LHM1
390     lhm2.Caption = F_LHM2
400     lhm3.Caption = f_lHM3
410     lhm4.Caption = f_lHM4
420     lhm5.Caption = f_lHM5
430     lhm6.Caption = f_lHM6

        ' Gdb.Execute "ALTER TABLE PEL ADD ENERGOS B"


Dim e3 As String
e3 = FINDPARAMETROI(3, "PELAT2", "e3", "----", String(80, "-"))

440     f_lab1 = FINDPARAMETROI(3, "PELAT2", "F_lab1", " ÂÈÏ1", "≈ÙÈÍ›Ù· ÍÂÈÏ.1")
450     f_lab2 = FINDPARAMETROI(3, "PELAT2", "F_lab2", " ÂÈÏ2", "≈ÙÈÍ›Ù· ÍÂÈÏ.2")
460     f_lab3 = FINDPARAMETROI(3, "PELAT2", "F_lab3", " ÂÈÏ3", "≈ÙÈÍ›Ù· ÍÂÈÏ.3")
470     f_lab4 = FINDPARAMETROI(3, "PELAT2", "F_lab4", " ÂÈÏ4", "≈ÙÈÍ›Ù· ÍÂÈÏ.4")
480     f_lab5 = FINDPARAMETROI(3, "PELAT2", "F_lab5", " ÂÈÏ5", "≈ÙÈÍ›Ù· ÍÂÈÏ.5")
490     f_lab6 = FINDPARAMETROI(3, "PELAT2", "F_lab6", " ÂÈÏ6", "≈ÙÈÍ›Ù· ÍÂÈÏ.6")

492     f_lab7 = FINDPARAMETROI(3, "PELAT2", "F_lab7", " ÂÈÏ7", "≈ÙÈÍ›Ù· ÍÂÈÏ.7")

500     lab1.Caption = f_lab1
510     lab2.Caption = f_lab2
520     lab3.Caption = f_lab3
530     lab4.Caption = f_lab4
540     lab5.Caption = f_lab5
550     lab6.Caption = f_lab6
        Lab7.Caption = f_lab7
        
        Dim e4 As String
e4 = FINDPARAMETROI(4, "PELAT2", "e4", "----", String(80, "-"))
        
560     f_num1 = FINDPARAMETROI(4, "PELAT2", "F_num1", "¡ÒÈË1", "≈ÙÈÍ›Ù· ¡ÒÈËÏ.1")
570     f_num2 = FINDPARAMETROI(4, "PELAT2", "F_num2", "¡ÒÈË2", "≈ÙÈÍ›Ù· ¡ÒÈËÏ.2")
580     f_num3 = FINDPARAMETROI(4, "PELAT2", "F_num3", "¡ÒÈË3", "≈ÙÈÍ›Ù· ¡ÒÈËÏ.3")
590     f_num4 = FINDPARAMETROI(4, "PELAT2", "F_num4", "¡ÒÈË4", "≈ÙÈÍ›Ù· ¡ÒÈËÏ.4")
600     Lnum1.Caption = f_num1
610     Lnum2.Caption = f_num2
620     Lnum3.Caption = f_num3
630     Lnum4.Caption = f_num4

        'Dim f_lbcb1 As String, f_comb1 As String
        'Dim f_lbcb2 As String, f_comb2 As String
        'Dim f_lbcb3 As String, f_comb3 As String
        'Dim f_lbcb4 As String, f_comb4 As String
        'Dim f_lbcb5 As String, f_comb5 As String
        
        
        
        
        Dim e50 As String
        e50 = FINDPARAMETROI(50, "PELAT2", "e50", "----", String(80, "-"))
        
        

        f_lbcb1 = FINDPARAMETROI(51, "PELAT2", "f_lbcb1", "¡ÒÈË1", "–ÔÎ.EÙÈÍ 1")
        f_lbcb2 = FINDPARAMETROI(53, "PELAT2", "f_lbcb2", "¡ÒÈË2", "–ÔÎ.EÙÈÍ 2")
        f_lbcb3 = FINDPARAMETROI(55, "PELAT2", "f_lbcb3", "¡ÒÈË3", "–ÔÎ.EÙÈÍ 3")
        f_lbcb4 = FINDPARAMETROI(57, "PELAT2", "f_lbcb4", "¡ÒÈË4", "–ÔÎ.EÙÈÍ 4")
        f_lbcb5 = FINDPARAMETROI(59, "PELAT2", "f_lbcb5", "¡ÒÈË4", "–ÔÎ.EÙÈÍ 5")
        lbcb1.Caption = f_lbcb1
        lbcb2.Caption = f_lbcb2
        lbcb3.Caption = f_lbcb3
        lbcb4.Caption = f_lbcb4
        lbcb5.Caption = f_lbcb5

        f_comb1 = FINDPARAMETROI(52, "PELAT2", "f_comb1", "", "≈ÈÎÔ„›Ú .˜. ≈ÏÔÒÔÚ;…‰È˛ÙÁÚ;’‹ÎÎÁÎÔÚ")
        f_comb2 = FINDPARAMETROI(54, "PELAT2", "f_comb2", "", "≈ÈÎÔ„›Ú .˜. ≈ÏÔÒÔÚ;…‰È˛ÙÁÚ;’‹ÎÎÁÎÔÚ")
        f_comb3 = FINDPARAMETROI(56, "PELAT2", "f_comb3", "", "≈ÈÎÔ„›Ú .˜. ≈ÏÔÒÔÚ;…‰È˛ÙÁÚ;’‹ÎÎÁÎÔÚ")
        f_comb4 = FINDPARAMETROI(58, "PELAT2", "f_comb4", "", "≈ÈÎÔ„›Ú .˜. ≈ÏÔÒÔÚ;…‰È˛ÙÁÚ;’‹ÎÎÁÎÔÚ")
        f_comb5 = FINDPARAMETROI(60, "PELAT2", "f_comb5", "", "≈ÈÎÔ„›Ú .˜. ≈ÏÔÒÔÚ;…‰È˛ÙÁÚ;’‹ÎÎÁÎÔÚ")

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
690    ' cmdDiortosi.Caption = "ƒÈ¸ÒË˘ÛÁ"   '”ƒ÷ƒ”÷

700     Me.Picture = LoadPicture(gPicture)

        'tropos plhromhs
        'Dim R As New ADODB.Recordset

710     Adodc2.ConnectionString = gConnect

720     Adodc2.RecordSource = "SELECT STR(AYJON,2,0) AS AYJ,PERIGRAFH FROM PINAKES WHERE TYPOS=13 ORDER BY AYJON"
730     Adodc2.Refresh

740     F_ArPerEtik = Val(FINDPARAMETROI(1, "PELAT2", "F_ArPerEtik", "2", "¡ÒÈÛÙÂÒ¸ ÂÒÈË˛ÒÈÔ ÂÙÈÍ›ÙÙ·Ú")) 'posa psifia tha exei h kathe seira
    
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
830         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH) VALUES (13,1,'≈Ã–œ—œ”')"
        End If

        'Do While Not R.EOF
        '   If R("typos") = 13 Then
        '      KATHGORIA.A Str(R("AYJON")) + "." '+ R(0) '"PERIGRAFH")
        '   End If
        '   R.MoveNext
        'Loop

840     R.Close

        ' R.Close

        '÷œ—‘ŸÃ¡ –ŸÀ«‘ŸÕ
850     R.Open "SELECT *FROM PINAKES WHERE TYPOS=5 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

860     Do While Not R.EOF
870         Combo3.AddItem (str(R("AYJON"))) + "." + R("PERIGRAFH")
880         R.MoveNext
        Loop

890     R.Close

900     If f_TIMOKAT = 1 Then
910         Label1(11).Caption = "‘ÈÏÔÍ·Ù‹ÎÔ„ÔÚ"
        End If

        Dim mf2

920     If Len(f_files) = 0 Then
        Else
930         mf2 = UCase(f_files + "\MERCVB\files\")
940         Dir1.Path = mf2
950         Dir1.Refresh
        End If

        '    mForm_Load Me, fh, fw, ft, fl
       
960   '  Me.Caption = str(GetCurrentTime() - AA)
        Text1(0).SetFocus
        Exit Sub

UPDATE_STR:

        On Error Resume Next

'970     Gdb.Execute "ALTER TABLE PEL ADD HM1 DATETIME"
'980     Gdb.Execute "ALTER TABLE PEL ADD HM2 DATETIME"
'990     Gdb.Execute "ALTER TABLE PEL ADD HM3 DATETIME"
'1000    Gdb.Execute "ALTER TABLE PEL ADD HM4 DATETIME"
'1010    Gdb.Execute "ALTER TABLE PEL ADD HM5 DATETIME"
'1020    Gdb.Execute "ALTER TABLE PEL ADD HM6 DATETIME"
'
'1030    Gdb.Execute "ALTER TABLE PEL ADD CH1 CHAR(30)"
'1040    Gdb.Execute "ALTER TABLE PEL ADD CH2 CHAR(30)"
'1050    Gdb.Execute "ALTER TABLE PEL ADD CH3 CHAR(30)"
'1060    Gdb.Execute "ALTER TABLE PEL ADD CH4 CHAR(30)"
'1070    Gdb.Execute "ALTER TABLE PEL ADD CH5 CHAR(30)"
'1080    Gdb.Execute "ALTER TABLE PEL ADD CH6 CHAR(30)"
'1090    Gdb.Execute "ALTER TABLE PEL ADD CH7 CHAR(30)"
'
'1100    Gdb.Execute "ALTER TABLE PEL ADD ENERGOS INTEGER"

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
            Text2(1).SetFocus
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
    If gApoMenu Then  ' ¡Õ ‘œ  ¡ÀŸ ¡–œ ‘«Õ  ¡—‘≈ÀÀ¡ Õ¡ Ã«Õ ¡ÀÀ¡∆≈…ƒ…¡”‘¡”≈…”
           'locForm_Resize Me, 8, ffW, ffH, fh, fw, ft, fl  ' 8 ,PLATOS,YCOS
           'ResizeControls
    End If '

 Frame2.Height = MDIForm1.Height - Toolbar1.Height
 DataGrid1.Height = Frame2.Height - 2000

 Frame2.Width = MDIForm1.Width

Frame1.Width = MDIForm1.Width

Frame1.Height = MDIForm1.Height












' cmbStyle.AddItem "Horizontal"
  '  cmbStyle.AddItem "Vertical"
   ' DrawFormGradient Me, vbBlue, vbCyan, 0 ' "-1=Vertical  0=horizontal"  ' cmbStyle.ListIndex    frmGradient
 'DrawFormGradient Me, vbBlue, vbCyan, 0






    ' El.mForm_Resize Me, 15000, 10000  ', fh, fw, ft, fl
  '  ResizeFormFor Me
    '3. In Form Resize, use:
    ' If Not m_IsLoading Then El.FormResize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'goo = Text2(1).Text
End Sub

Private Sub Frame1_DblClick()

100     PARAMETROI.PARAM.Caption = "PELAT2"
110     PARAMETROI.SHOW 1
End Sub

Private Sub hmesymf_LostFocus()

    Dim a As String
  
    If IsDate(HMESYMF) Then
   
    Else
        HMESYMF = ""
   
    End If

End Sub

Private Sub keim_GotFocus(index As Integer)

    keim(index).BackColor = vbYellow
    keim(index).SelStart = 0
    keim(index).SelLength = Len(keim(index).Text)




End Sub

Private Sub keim_KeyPress(index As Integer, KeyAscii As Integer)


     '</EhHeader>
       If KeyAscii = vbKeyReturn Then
           KeyAscii = 0    'suppress the beep
           keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

End Sub

Private Sub keim_LostFocus(index As Integer)
  keim(index).BackColor = f_keim_Back_color


End Sub

Private Sub kentroAdyn_Click()

    On Error Resume Next

    If f_KENTRA_ADYNATISMATOS = 1 Then
        If Data1.Recordset("NUM3") = 0 Then ' HTAN STO SVMA
            If kentroAdyn.Text = "”ŸÃ¡" Then  ' PARAMENEI STO SOMA
                'MHN KANEIS TIPOTA
                Text2(0).Text = Data1.Recordset("kod")
            Else  ' TO GYRIZEI STO PROSOPO
                Text2(0).Text = "10" + Data1.Recordset("KOD")
            End If
        End If
   
        If Data1.Recordset("NUM3") = 1 Then ' HTAN STO PROSOPO
            If kentroAdyn.Text <> "”ŸÃ¡" Then  ' PARAMENEI STO PROSOPO
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

Private Sub Label1_Click(index As Integer)

        '<EhHeader>
        On Error GoTo Label1_Click_Err

        '</EhHeader>

100     If index = 18 Then
110         UPDATE_PINAKES 12, Label1(index), Combo2
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
'    '·ÎÎ·„ﬁ „Ò·ÏÏ·ÙÔÛÂÈÒ‹Ú
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
'    '·ÎÎ·„ﬁ „Ò·ÏÏ·ÙÔÛÂÈÒ‹Ú
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

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


'Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ' If ColumnHeader.Text = "Size" Then
    '    Listview1.SortKey = 4
    'ElseIf ColumnHeader.Text = "Date" Then
     '       Listview1.SortKey = 5
    'Else
        Listview1.SortKey = ColumnHeader.index - 1
    'End If
    Listview1.SortOrder = Abs(Not Listview1.SortOrder = 1)
    Listview1.Sorted = True
'End Sub







End Sub

Private Sub ListView1_DblClick()

 '</EhHeader>

        Dim sFileName As String

        Dim xl        As New excel.Application

        Dim xlsheet   As excel.Worksheet

        Dim xlsheet3  As excel.Worksheet

        Dim xlwbook   As excel.workbook

        '
        ' sFileName = File1.Path + "\" + File1.FILENAME
100     sFileName = File1.Path + "\" + Listview1.SelectedItem  ' File1.FILENAME

        'Exit Sub
' Listview1.SelectedItem
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
























End Sub

Private Sub SSTab1_DblClick()
  'DrawFormGradient Me, vbBlue, vbCyan, 0 ' "-1=Vertical  0=horizontal"  ' cmbStyle.ListIndex    frmGradient
End Sub

Private Sub SSTab1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' DrawFormGradient Me, vbBlue, vbCyan, 0 ' "-1=Vertical  0=horizontal"  ' cmbStyle.ListIndex    frmGradient
End Sub

'        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat2.Picture2_Click " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
Private Sub Text1_GotFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text1_GotFocus_Err
       ' Text1(Index).BackColor = vbYellow

        '</EhHeader>

100     Text1(index).BackColor = vbYellow

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

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)

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

Private Sub Text1_LostFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text1_LostFocus_Err

        '</EhHeader>
        Dim B, a

100     Text1(index).BackColor = vbWhite

110     If index = 5 Then
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

Private Sub Text2_GotFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text2_GotFocus_Err

        '</EhHeader>

100     Text2(index).BackColor = vbYellow

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

Private Sub Text2_KeyPress(index As Integer, KeyAscii As Integer)

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

Public Sub Text2_LostFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text2_LostFocus_Err

        '</EhHeader>

        Dim X As String

100     Text2(index).BackColor = vbWhite

110     If Len(Trim(Text2(index))) = 0 Then Exit Sub
        'If Option1(1) Then X = "e" Else X = "r"

120     X = "e"

        On Error Resume Next

130     X = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)
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
            If X = "e" Then
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
220     If index = 1 Then  'eponumo

230         'If Asc(Left(Text2(1).Text, 1)) > 127 Then
240          '   FF = Text2(1).Text
            'Else
250             FF = (SameLetters(Text2(1).Text))
            'End If
            If F_MONOENERGOS = 0 Then '  Val(FINDPARAMETROI(1, "PAR1", "F_MONOENERGOS", "0", "ƒ≈…◊Õ≈… ÃœÕœ ‘œ’” ≈Õ≈—√œ’”=1 œÀœ’”=0")) '
                Data1.RecordSource = "select *from PEL where " + MSYNT + " EIDOS='" + X + "'  and EPO like '" + FF + "%';"
            Else
              If chkANENERGOS.Value = vbChecked Then
                 Data1.RecordSource = "select *from PEL where " + MSYNT + " EIDOS='" + X + "'  and EPO like '" + FF + "%';"
              Else
260              Data1.RecordSource = "select *from PEL where " + MSYNT + " EIDOS='" + X + "' AND ISNULL(ENERGOS,1)=1 and EPO like '" + FF + "%';"
              End If
            End If
        End If

270     If index = 0 Then    'kodikos
280         Data1.RecordSource = "select *from PEL where " + MSYNT + " EIDOS='" + X + "' and KOD like '" + (SameLetters(Text2(0).Text)) + "%';"
        End If

290     If index = 2 Then    'afm
           
           If Len(Trim(f_pedioAFM)) = 0 Then
                Data1.RecordSource = "select *from PEL where " + MSYNT + " EIDOS='" + X + "' and AFM like '" + Text2(2).Text + "%';"
           Else
                Data1.RecordSource = "select *from PEL where " + MSYNT + " EIDOS='" + X + "' and   ( AFM like '%" + Text2(2).Text + "%' or " + Trim(f_pedioAFM) + " like '%" + Text2(2).Text + "%') "
           
           End If
           
                
        End If

        ' End If

        'DataGrid1.Enabled = False

310     Data1.Refresh
320     DataGrid1.columns(0).Width = TextWidth("A") * 30
          
        MDataGrid1_RowColChange
        'DataGrid1.Enabled = True

        Dim mF As String

        '  If Option1(1).value = True Then
330     mF = UCase(f_files + "\MERCVB\files\" + mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1) + Text1(0))

        '  Else
        '     mF = UCase(f_files + "\MERCVB\files\r" + Text1(0))
        '  End If
        On Error Resume Next

340     If Text1(0).Text = Data1.Recordset("KOD") Then  'GIA NA MHN KATHISTEREI
350         If Len(Dir(mF, vbDirectory)) > 0 Then
360             File1.Path = mF
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

260                     If r0(0) > 0 Then    '«ƒ« ’–¡—◊≈…
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
370            ' Me.Caption = k

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
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'  Private Sub xxToolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.key
    Case "cmdNew"
       cmdNew_Click
       cmdDiortHosi
       
       Frame2.Visible = False
       Frame1.Top = Toolbar1.Height
       Text1(0).Enabled = True  'ÂÌÂÒ„ÔÔÈ˛ ÙÔÌ Í˘‰ÈÍ¸ „È· Ì· ÏÔÒÂﬂ Ì· ÙÔÌ ·ÎÎ‹ÓÂÈ
      
       
       Toolbar1.Buttons.Item("cmdNew").Value = tbrPressed
       'Toolbar1.Buttons.Item("cmdNew").Enabled = False ' vbRed
       Toolbar1.Buttons.Item("cmdDiortosi").Enabled = False
       Toolbar1.Buttons.Item("cmdKartela").Enabled = False
        Toolbar1.Buttons.Item("cmdDelete").Enabled = False
        Toolbar1.Buttons.Item("cmdEtiketa").Enabled = False
       ' cmdEtiketa
        Toolbar1.Buttons.Item("cmdChange").Enabled = False

       'ÂÌÂÒ„ÔÔÈ˛ ÙÔ UPDATE
       Toolbar1.Buttons.Item("cmdUpdate").Enabled = True


        Text1(0).SetFocus
       
       
       
       
    Case "cmdDiortosi"
       cmdDiortHosi
       
       
       Frame2.Visible = False
       Frame1.Top = Toolbar1.Height
       
       Toolbar1.Buttons.Item("cmdDiortosi").Value = tbrPressed
       Toolbar1.Buttons.Item("cmdNew").Enabled = False ' vbRed
      
       Toolbar1.Buttons.Item("cmdKartela").Enabled = False
        Toolbar1.Buttons.Item("cmdDelete").Enabled = False
        Toolbar1.Buttons.Item("cmdEtiketa").Enabled = False
       ' cmdEtiketa
        Toolbar1.Buttons.Item("cmdChange").Enabled = False
       
       Toolbar1.Buttons.Item("cmdUpdate").Enabled = True
       
       
       
       
       
 Case "cmdUpdate"
       cmdUpdate
       Toolbar1.Buttons.Item("cmdNew").Value = tbrUnpressed
       
       Toolbar1.Buttons.Item("cmdDiortosi").Enabled = True
       Toolbar1.Buttons.Item("cmdKartela").Enabled = True
       Toolbar1.Buttons.Item("cmdDelete").Enabled = True
       Toolbar1.Buttons.Item("cmdEtiketa").Enabled = True
       Toolbar1.Buttons.Item("cmdChange").Enabled = True
       Toolbar1.Buttons.Item("cmdNew").Enabled = True
       Toolbar1.Buttons.Item("cmdUpdate").Enabled = False
       
       
       Text1(0).Enabled = False  '·ÂÌÂÒ„ÔÔÈ˛ ÙÔÌ Í˘‰ÈÍ¸ „È· Ì· ÏÁÌ  ÏÔÒÂﬂ Ì· ÙÔÌ ·ÎÎ‹ÓÂÈ
       
            
       Frame2.Visible = True
       Frame1.Top = Frame2.Height + 100
       
     Case "cmdDelete"
       cmdDelete_Click
       
       
     Case "cmdKartela"
     
       cmdKartela_Click
       
     Case "cmdExit"
       EXODOS
       
       
    Case "cmdEtiketa"
        cmdEtiketta_Click
     
    Case "cmdChange"
      ' cmdChange_Click
       CmdAllaghKodikou_Click
       
       
End Select

End Sub



'End Sub
