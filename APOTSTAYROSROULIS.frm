VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form apot2ROULISnew 
   Caption         =   "Form1"
   ClientHeight    =   10905
   ClientLeft      =   225
   ClientTop       =   -2385
   ClientWidth     =   18960
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
   ScaleHeight     =   10905
   ScaleWidth      =   18960
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox EIDxalasmeno 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   18480
      ScaleHeight     =   15
      ScaleWidth      =   1200
      TabIndex        =   109
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   18315
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   18375
      Begin VB.CommandButton cmdDELETE 
         BackColor       =   &H80000004&
         Caption         =   "Διαγραφή Εμπόρευμα"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   8160
         MaskColor       =   &H00FFFFFF&
         Picture         =   "APOTSTAYROSROULIS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CMDmovenext 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12360
         Picture         =   "APOTSTAYROSROULIS.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton CMDmoveprevious 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11640
         Picture         =   "APOTSTAYROSROULIS.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000005&
         Caption         =   "Αναζήτηση Τιμολόγιο"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   10080
         Picture         =   "APOTSTAYROSROULIS.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1320
         TabIndex        =   53
         Top             =   600
         Width           =   3405
      End
      Begin VB.CommandButton neo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Νέα Εγγραφή"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "APOTSTAYROSROULIS.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton upload1 
         BackColor       =   &H80000005&
         Caption         =   "Upload"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   13320
         Picture         =   "APOTSTAYROSROULIS.frx":104A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdΕΝΗΜΕΡΩΣΗ 
         BackColor       =   &H0000FF00&
         Caption         =   "ΑΠΟΘΗΚΕΥΣΗ (ΕΜΠΟΡΙΚΟ)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   0
         Left            =   14880
         MaskColor       =   &H00FFFFFF&
         Picture         =   "APOTSTAYROSROULIS.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdANAZHTHSH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ΑΚΥΡΩΣΗ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6600
         Picture         =   "APOTSTAYROSROULIS.frx":16CE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdΕξοδος 
         BackColor       =   &H000000FF&
         Caption         =   "ΕΞΟΔΟΣ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   16440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "APOTSTAYROSROULIS.frx":1A10
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox ONO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   3405
      End
      Begin VB.TextBox KOD 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   1320
         MaxLength       =   16
         TabIndex        =   58
         Top             =   0
         Width           =   3405
      End
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   5040
         X2              =   15120
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   5040
         X2              =   15120
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label LIST1_Label 
         BackColor       =   &H00FF8080&
         Caption         =   "..."
         Height          =   15
         Left            =   0
         TabIndex        =   107
         Top             =   1800
         Visible         =   0   'False
         Width           =   18000
      End
      Begin VB.Label mERG 
         BackColor       =   &H0080C0FF&
         Height          =   360
         Left            =   13320
         TabIndex        =   60
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Γνήσιος  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   54
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label NEA 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "ΝΕΑ ΕΓΓΡΑΦΗ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15480
         TabIndex        =   48
         Top             =   1200
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label kodlabel 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         DataField       =   "KOD"
         DataSource      =   "EID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   5040
         TabIndex        =   14
         Top             =   1200
         Width           =   1845
      End
      Begin VB.Label ONOLABEL 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         DataField       =   "ONO"
         DataSource      =   "EID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   6960
         TabIndex        =   12
         Top             =   1200
         Width           =   6315
      End
      Begin VB.Label lblΠεριγραφή 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Περιγραφή :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblΚωδικός 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδικός :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9375
      Left            =   18120
      TabIndex        =   1
      Top             =   1800
      Width           =   252
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   13095
      Left            =   0
      ScaleHeight     =   13065
      ScaleWidth      =   18105
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1800
      Width           =   18135
      Begin MSAdodcLib.Adodc eggtim2 
         Height          =   330
         Left            =   16200
         Top             =   7680
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   15840
         Top             =   8760
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   15960
         Top             =   8160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   255
         Left            =   8760
         TabIndex        =   115
         Top             =   8880
         Width           =   6015
         ExtentX         =   10610
         ExtentY         =   450
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin RichTextLib.RichTextBox Rich1 
         Height          =   495
         Left            =   17400
         TabIndex        =   114
         Top             =   8040
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"APOTSTAYROSROULIS.frx":1D52
      End
      Begin MSComctlLib.TreeView T1 
         Height          =   615
         Left            =   17400
         TabIndex        =   113
         Top             =   8520
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin ADOMERCNEW.msEditor2 msEditor21 
         Height          =   2500
         Left            =   1560
         TabIndex        =   112
         Top             =   7560
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2778
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
      Begin MSAdodcLib.Adodc eid 
         Height          =   330
         Left            =   17160
         Top             =   8880
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
      Begin VB.CommandButton cmdΕΝΗΜΕΡΩΣΗ 
         BackColor       =   &H0000FF00&
         Caption         =   "ΑΠΟΘΗΚΕΥΣΗ (ΕΜΠΟΡΙΚΟ)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   1
         Left            =   16440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "APOTSTAYROSROULIS.frx":1DD8
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H80000005&
         Caption         =   "Καρτέλλα Είδους"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   16440
         Picture         =   "APOTSTAYROSROULIS.frx":211A
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   840
         Width           =   1455
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   106
         Top             =   0
         Visible         =   0   'False
         Width           =   18000
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H80000005&
         Caption         =   "Διόρθωση Είδους"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   16440
         Picture         =   "APOTSTAYROSROULIS.frx":245C
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H80000005&
         Caption         =   "Εκτύπωση Ετικέτας"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   16440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "APOTSTAYROSROULIS.frx":279E
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   5760
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "APOTSTAYROSROULIS.frx":2AE0
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7800
         Picture         =   "APOTSTAYROSROULIS.frx":2E22
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         DataField       =   "PROM"
         DataSource      =   "EID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11880
         TabIndex        =   100
         Top             =   360
         Width           =   6015
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "APOTSTAYROSROULIS.frx":3164
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   2280
         Width           =   375
      End
      Begin VB.ComboBox MARKA 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   97
         Text            =   "Combo3"
         Top             =   3600
         Width           =   5895
      End
      Begin VB.ComboBox EIDOS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   96
         Text            =   "Είδος"
         Top             =   2280
         Width           =   5175
      End
      Begin VB.ComboBox KATHGORIA 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "APOTSTAYROSROULIS.frx":34A6
         Left            =   1560
         List            =   "APOTSTAYROSROULIS.frx":34A8
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton delete 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   8040
         Picture         =   "APOTSTAYROSROULIS.frx":34AA
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   5760
         Width           =   375
      End
      Begin VB.CommandButton delete 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   8040
         Picture         =   "APOTSTAYROSROULIS.frx":37EC
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox POS 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   11880
         TabIndex        =   91
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton ADD 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   7560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "APOTSTAYROSROULIS.frx":3B2E
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   5760
         Width           =   375
      End
      Begin VB.CommandButton ADD 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   7560
         Picture         =   "APOTSTAYROSROULIS.frx":3E70
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   4440
         Width           =   375
      End
      Begin VB.CommandButton delete 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   8040
         Picture         =   "APOTSTAYROSROULIS.frx":41B2
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   3630
         Width           =   375
      End
      Begin VB.CommandButton ADD 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   7560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "APOTSTAYROSROULIS.frx":44F4
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3600
         Width           =   375
      End
      Begin VB.ListBox List 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   2
         Left            =   1560
         TabIndex        =   86
         Top             =   5040
         Width           =   5895
      End
      Begin VB.ListBox List 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   1560
         TabIndex        =   85
         Top             =   4080
         Width           =   5895
      End
      Begin VB.CheckBox used 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Μεταχειρισμένο"
         DataField       =   "USED"
         DataSource      =   "EID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8280
         TabIndex        =   84
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox XTI 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   14400
         TabIndex        =   82
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         Picture         =   "APOTSTAYROSROULIS.frx":4836
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1800
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11880
         TabIndex        =   78
         Text            =   "Combo1"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox UES 
         Appearance      =   0  'Flat
         DataField       =   "UES"
         DataSource      =   "EID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11880
         TabIndex        =   76
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox KODERG 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   74
         Text            =   "Κωδικός Κατασκευαστή"
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8280
         TabIndex        =   73
         Top             =   360
         Width           =   2535
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   8280
         TabIndex        =   72
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox ERGOSTASIO 
         Appearance      =   0  'Flat
         DataField       =   "TEXT_PRODUCTS_ZIPCODE"
         DataSource      =   "EID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   70
         Text            =   "Εργοστάσιο Κατασκευής (Μάρκα)"
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox ONO2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   68
         Text            =   "Κύρια Περιγραφή"
         Top             =   1800
         Width           =   6615
      End
      Begin VB.CommandButton cmdDIPLA 
         Caption         =   "DIPLA"
         Height          =   360
         Left            =   13080
         TabIndex        =   67
         Top             =   9840
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   1560
         TabIndex        =   66
         Top             =   6360
         Width           =   5895
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Eνημέρωση site από εμπορικό"
         Height          =   615
         Left            =   13680
         TabIndex        =   65
         Top             =   20000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "preview"
         Height          =   360
         Left            =   11640
         TabIndex        =   64
         Top             =   9840
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.ComboBox ftp 
         Height          =   300
         ItemData        =   "APOTSTAYROSROULIS.frx":4B78
         Left            =   17760
         List            =   "APOTSTAYROSROULIS.frx":4B82
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   10440
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CheckBox reload 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Επαναφόρτωση"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8280
         TabIndex        =   62
         Top             =   2640
         Width           =   1455
      End
      Begin VB.PictureBox WebBrowser1xalasmeno 
         Height          =   6045
         Left            =   -600
         ScaleHeight     =   5985
         ScaleWidth      =   14445
         TabIndex        =   61
         Top             =   11040
         Visible         =   0   'False
         Width           =   14505
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   960
         TabIndex        =   52
         Top             =   10440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "DOWNLOAD TABLE PRODUCTS"
         Height          =   375
         Left            =   3000
         TabIndex        =   51
         Top             =   10440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox WEBPRICE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   14400
         TabIndex        =   50
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   14400
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   45
         Text            =   "APOTSTAYROSROULIS.frx":4B95
         Top             =   9600
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.ListBox hlikies 
         Height          =   1680
         Left            =   5640
         MultiSelect     =   2  'Extended
         TabIndex        =   43
         Top             =   20000
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Διαθέσιμο"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9840
         TabIndex        =   41
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox kataskeyasths 
         Height          =   300
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   20000
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.ComboBox xaraktiras 
         Height          =   300
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   20000
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.ComboBox fylo 
         Height          =   300
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   20000
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox WEIGHT 
         Appearance      =   0  'Flat
         DataField       =   "WEIGHT"
         DataSource      =   "EID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11880
         TabIndex        =   33
         Top             =   840
         Width           =   1455
      End
      Begin VB.PictureBox foto2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   6
         Left            =   14880
         ScaleHeight     =   1665
         ScaleWidth      =   1905
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1935
      End
      Begin VB.PictureBox foto2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   5
         Left            =   12840
         ScaleHeight     =   1545
         ScaleWidth      =   1905
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   7200
         Width           =   1935
      End
      Begin VB.PictureBox foto2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   4
         Left            =   10800
         ScaleHeight     =   1545
         ScaleWidth      =   1905
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   7200
         Width           =   1935
      End
      Begin VB.PictureBox foto2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   3
         Left            =   8760
         ScaleHeight     =   1545
         ScaleWidth      =   1905
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   7200
         Width           =   1935
      End
      Begin VB.PictureBox foto2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   2
         Left            =   12840
         ScaleHeight     =   1665
         ScaleWidth      =   1905
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   5400
         Width           =   1935
      End
      Begin VB.PictureBox foto2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   1
         Left            =   12840
         ScaleHeight     =   1665
         ScaleWidth      =   1905
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox LTI5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   14400
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox FONO 
         Appearance      =   0  'Flat
         DataField       =   "ono"
         DataSource      =   "EID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   16
         Text            =   "Περιγραφή από Τιμολόγιο"
         Top             =   1320
         Width           =   6615
      End
      Begin VB.PictureBox CommonDialog1xalasmeno 
         Height          =   465
         Left            =   2880
         ScaleHeight     =   405
         ScaleWidth      =   1500
         TabIndex        =   110
         Top             =   9600
         Width           =   1560
      End
      Begin VB.PictureBox foto2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3495
         Index           =   0
         Left            =   8760
         ScaleHeight     =   3465
         ScaleWidth      =   3945
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   360
         Left            =   0
         TabIndex        =   2
         Top             =   13200
         Width           =   990
      End
      Begin VB.PictureBox msEditor22 
         BackColor       =   &H00C0FFFF&
         Height          =   975
         Left            =   5640
         ScaleHeight     =   915
         ScaleWidth      =   1995
         TabIndex        =   7
         Top             =   13680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Προμηθευτής"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   101
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Κύρια Φωτογραφία"
         Height          =   255
         Left            =   10080
         TabIndex        =   99
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Υπόλοιπο"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11160
         TabIndex        =   92
         Top             =   2520
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   17520
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Τιμ.Αγοράς"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13560
         TabIndex        =   83
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Κωδ.Γνήσιου"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   80
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Μον.Μέτρ."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11160
         TabIndex        =   79
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Ραφι"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   77
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδ.Κατασκευαστή"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   75
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Εργ.Κατασκευής"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   71
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Κύρια Περιγραφή"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblΤιμήWEB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Τιμή WEB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   13680
         TabIndex        =   49
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label lblΚατηγΦΠΑ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κατηγ.ΦΠΑ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   13440
         TabIndex        =   47
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label cΗλικίες 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Ηλικίες"
         Height          =   330
         Left            =   5640
         TabIndex        =   42
         Top             =   20000
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Χαρακτήρας/Θέμα"
         Height          =   180
         Left            =   480
         TabIndex        =   40
         Top             =   20000
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αγόρι/Κορίτσι"
         Height          =   180
         Left            =   480
         TabIndex        =   39
         Top             =   20000
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κατασκευαστής"
         Height          =   180
         Left            =   480
         TabIndex        =   38
         Top             =   20000
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Βάρος"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11400
         TabIndex        =   34
         Top             =   960
         Width           =   435
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   10305
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Τιμή Καταστ."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   13440
         TabIndex        =   19
         Top             =   2040
         Width           =   1020
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Περιγραφή Τιμολόγιο"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   1440
         Width           =   1515
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
         TabIndex        =   15
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
         TabIndex        =   8
         Top             =   13440
         Visible         =   0   'False
         Width           =   3000
      End
   End
End
Attribute VB_Name = "apot2ROULISnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================   ftp  ================================================================
'custom fields

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
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
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
    Y As Long

End Type

Private Type Msg

    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI

End Type

Private Declare Function PeekMessage _
                Lib "user32" _
                Alias "PeekMessageA" (lpMsg As Msg, _
                                      ByVal hwnd As Long, _
                                      ByVal wMsgFilterMin As Long, _
                                      ByVal wMsgFilterMax As Long, _
                                      ByVal wRemoveMsg As Long) As Long

Private Declare Function WaitMessage Lib "user32" () As Long

Private bCancel As Boolean

Private Const WM_MOUSEWHEEL = 522

'αποθηκευει τα στοιχεια των κατηγοριων
Dim fc_max          As Long

Dim FC_id(5000)     As Long

Dim FC_parenT(5000) As Long

Dim fc_names(5000)  As String

Dim fc_aa(5000)     As Integer

Dim names(5000)     As String

Private Sub ProcessMessages()

    Dim Message As Msg
 
    Dim s

    s = 200

    Do While Not bCancel

        If f_exit = True Then
            Unload Me
        End If
 
        WaitMessage 'Wait For message and...

        If PeekMessage(Message, Me.hwnd, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then '...when the mousewheel is used...
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

    Dim R As New ADODB.Recordset

    Dim SS, N As Integer, kor As String, B As String, manuf As String
   
    On Error Resume Next
    
    List1.Clear
    List2.Clear
    List3.Clear
    KODERG.Text = ""
    
    KATHGORIA.Text = EID.Recordset("TEXT_PRODUCTS_ZIPCODE")
    
    ONOLABEL.Caption = EID.Recordset("ono")
    WEBPRICE.Text = ""

    If IsNull(EID.Recordset("WEBPRICE")) Then
        WEBPRICE.Text = EID.Recordset("lti5")
    Else
        WEBPRICE.Text = EID.Recordset("WEBPRICE")
    End If

    mERG.Caption = EID.Recordset("ERG")
    
    kodlabel.Caption = EID.Recordset("kod")
    
    Set msEditor21.mDataSource = EID  ' for code you have to use the 'adoNotes' instead of the Adodc1
    msEditor21.MaxLength = 8000          ' Optional maximum length.

    msEditor21.mDataField = "MEMO" '  "nNotes"     ' The field of your text
    msEditor21.Enabled = True
    EID.Recordset.Move 0

    '  Set msEditor22.mDataSource = EID  ' for code you have to use the 'adoNotes' instead of the Adodc1
    '  msEditor22.MaxLength = 8000          ' Optional maximum length.

    '  msEditor22.mDataField = "MEMOENG" '  "nNotes"     ' The field of your text
    '  msEditor22.Enabled = True

    'Set msEditor1.mDataSource = DATA1  ' for code you have to use the 'adoNotes' instead of the Adodc1
    ' msEditor1.MaxLength = 8000          ' Optional maximum length.
    '
    ' msEditor1.mDataField = "MEMO" '  "nNotes"     ' The field of your text
    ' msEditor1.Enabled = True
 
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
            StretchSourcePictureFromPicture apot2ROULIS.foto2(N).Picture, apot2ROULIS.foto2(N)
        Else
            foto2(N).Picture = LoadPicture()
        End If

    Next
   
    Combo2.ListIndex = EID.Recordset("FPA") - 1
    kataskeyasths.ListIndex = EID.Recordset("MANUFACTURERS_ID")
   
    FONO.Text = EID.Recordset("ONO")
    KODERG.Text = EID.Recordset("KODERG")
    ERGOSTASIO = ERGOSTASIO.Text
    ONO2.Text = EID.Recordset("ONO2")
    WEIGHT.Text = EID.Recordset("WEIGHT")
    UES.Text = EID.Recordset("UES")
    POS.Text = EID.Recordset("POS")
    XTI.Text = EID.Recordset("XTI")
    lti5.Text = EID.Recordset("LTI5")
    WEBPRICE.Text = EID.Recordset("WEBPRICE")
  
    If EID.Recordset("diathesimo") = 1 Then
        Check1.Value = vbChecked
    Else
        Check1.Value = vbUnchecked
    End If

    Combo2.ListIndex = EID.Recordset("fpa") - 1

    Combo1.Text = EID.Recordset("MON")

    'T1.Nodes.Clear
    '
    'For n = 1 To fc_max
    '196    error = 0
    '       'paterades
    '        If FC_parenT(n) = 0 Then
    '202       T1.Nodes.ADD , , "r" + LTrim(str(FC_id(n))), fc_names(n)
    '       Else
    '
    '204     T1.Nodes.ADD "r" + LTrim(str(FC_parenT(n))), tvwChild, "r" + LTrim(str(FC_id(n))), (fc_names(n))
    '       End If
    'Next
    '
    If f_site = 333 Then
        R.Open "SELECT * FROM GNISIOI_KODIKOI WHERE KOD='" + kod + "'", Gdb, adOpenDynamic, adLockOptimistic
        'Gdb.Execute "delete from GNISIOI_KODIKOI where KOD='" + KOD + "'"

        Do While Not R.EOF

            List3.AddItem R!GNISIOS, 0
            'Gdb.Execute "INSERT INTO GNISIOI_KODIKOI (KOD,GNISIOS) VALUES ('" + KOD + "'," + str(F_id_LOCAL) + "," + m_categ + ");"
            'End If

            R.MoveNext

        Loop

        R.Close
    End If

    '

    R.Open "SELECT * FROM PRODUCT_TO_CATEGORIES WHERE KOD='" + kod + "'", Gdb, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF

        ' T1.Nodes.Item(20).Expanded = True
        ' T1.Nodes.Item(7).Selected = True
        For k = 1 To UBound(names) - 1

            If Val(mID(names(k), 2, 5)) = R!categories_id Then
                T1.Nodes.Item(k).Selected = True
                T1.Nodes.Item(k).BackColor = vbCyan
                T1.SetFocus
                ' T1.ShowWhatsThis
                List2.AddItem names(k), 0
            End If

        Next
   
        If f_site = 222 Then

            Exit Do  ' GIA MIA KATHGORIA

        End If
   
        R.MoveNext
    Loop

    R.Close

    Dim surl As String

    Dim m_ID As String

    If f_site = 111 Then

        'gdblag.Close
        'gdblag.Open f_connect  '"DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131 ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
        If Not R.EOF Then
            R.Open "select * from " + f_vasi + "_virtuemart_products where product_sku='" + kod + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
            m_ID = str(R("virtuemart_product_id"))
            surl = "http://www.lagakis.gr/administrator/index.php?option=com_virtuemart&view=product&task=edit&virtuemart_product_id=" + m_ID + "&product_parent_id=0"
            WebBrowser1.navigate surl
        End If
    End If

    If f_site = 333 Then

        'gdblag.Close
        'gdblag.Open f_connect  '"DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131 ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
        If Not R.EOF Then
            R.Open "select * from " + f_vasi + "_virtuemart_products where product_sku='" + kod + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
            m_ID = str(R("virtuemart_product_id"))
            surl = "http://www.auto-roulis.gr/site/administrator/index.php?option=com_virtuemart&view=product&task=edit&virtuemart_product_id=" + m_ID + "&product_parent_id=0"
            WebBrowser1.navigate surl
        End If
    
        EIDOS.Text = EID.Recordset("YOU_TUBE")
        ERGOSTASIO.Text = EID.Recordset("ERGOSTASIO")
    
    End If

    If f_site = 222 Then
        R.Open "select * from products where products_model='" + kod + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

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

Private Sub add0_Click()

    Dim Y

    Dim nam

    Dim c    As String

    Dim sql  As String

    Dim k    As Long  '240 agglia 131 isp 64    '

    Dim m_ID As String

    Dim R    As New ADODB.Recordset

    If T1.Nodes.Count > 0 And KeyCode = 13 Then 'prevents an error if TV is empty
        Y = T1.SelectedItem.index
        nam = T1.SelectedItem.Text
        Me.Caption = nam & " - " & str(Y) + " names=" + names(Y)
        Set T1.DropHighlight = T1.SelectedItem

        '   m_categ = mID(names(Y), 2, 5)
        List2.AddItem names(Y)
    End If

    If T1.Nodes.Count > 0 And KeyCode = 46 Then 'DELETE SBHNΩ ΚΑΤΗΓΟΡΙΑ
        Y = T1.SelectedItem.index
        'Label1.Caption = Str(y)
        nam = T1.SelectedItem.Text
        Me.Caption = nam & " - " & str(Y) + " names=" + names(Y) + str(FC_id(Y))

        ' FC_id (Y)

        'ψαχνω να δω αν υπάρχουν κατηγοριες με πατερα αυτο που θα σβήσω
        R.Open "SELECT count(*) from " + f_vasi + "_virtuemart_category_categories where category_parent_id=" + str(FC_id(Y)), gdblag, adOpenDynamic, adLockOptimistic

        If R(0) > 0 Then
            MsgBox "αδυνατη η διαγραφή γιατί υποδιαιρείται η κατηγορία"

            Exit Sub

        End If

        R.Close
 
        'ψαχνω να δω αν υπάρχουν ειδη με κατηγορια αυτο που θα σβήσω
        R.Open "SELECT count(*) from " + f_vasi + "_virtuemart_product_categories where virtuemart_category_id=" + str(FC_id(Y)), gdblag, adOpenDynamic, adLockOptimistic

        If R(0) > 0 Then
            MsgBox "αδυνατη η διαγραφή γιατί υπάρχουν είδη με αυτή την κατηγορία"

            Exit Sub

        End If

        R.Close
 
        sql = "DELETE FROM  " + f_vasi + "_virtuemart_category_categories  where id=" + str(FC_id(Y))
        gdblag.Execute sql, k

        '    sql = "DELETE FROM  " + f_vasi + "_virtuemart_categories_el_gr  where virtuemart_category_id=" + str(FC_id(Y))
        '   gdblag.Execute sql, k

        sql = "DELETE FROM  " + f_vasi + "_virtuemart_categories_el_gr  where virtuemart_category_id=" + str(FC_id(Y))
        gdblag.Execute sql, k

        categories

        'MsgBox "ok"
        'T1.Refresh
        ' επιλεγω το καινουριο
        For k = 1 To UBound(names) - 1

            If Val(mID(names(k), 2, 5)) = Y Then
                T1.Nodes.Item(Y).Selected = True
                T1.Nodes.Item(Y).BackColor = vbCyan
                T1.SetFocus
                T1.Nodes.Item(Y).Expanded = True
                T1.Refresh
           
                Exit For

            End If
        
        Next

    End If

    If T1.Nodes.Count > 0 And KeyCode = 45 Then 'insert ΠΡΟΣΘΕΤΩ ΚΑΤΗΓΟΡΙΑ
        Y = T1.SelectedItem.index
        'Label1.Caption = Str(y)
        nam = T1.SelectedItem.Text
        Me.Caption = nam & " - " & str(Y) + " names=" + names(Y)
        Set T1.DropHighlight = T1.SelectedItem

        '   m_categ = mID(names(Y), 2, 5)
        List2.AddItem names(Y)
   
        '       FC_id(N) = r!ID
        '    FC_parenT(N) = r!Parent
        On Error GoTo T1_KeyUp_err
   
        If fc_aa(Y) = 1 Then
            ' T1.Nodes.Add , , "r" + LTrim(str(r!ID)), r!Name
            MsgBox "ΑΔΥΝΑΤΗ ΕΙΣΑΓΩΓΗ ΣΕ ΑΥΤΟ ΤΟ ΕΠΙΠΕΔΟ 1"
        Else
            'T1.Nodes.Clear
            c = InputBox("Δώσε τον τίτλο της κατηγορίας", "Εισαγωγή τύπου")
            sql = "INSERT INTO " + f_vasi + "_virtuemart_category_categories (category_parent_id,ordering)"
            sql = sql + " VALUES (" + str(FC_id(Y)) + ",0);"
            gdblag.Execute sql, k

            If k > 0 Then
                      
                R.Open "SELECT MAX(id) from " + f_vasi + "_virtuemart_category_categories", gdblag, adOpenDynamic, adLockOptimistic
                m_ID = str(R(0))
                      
                sql = "INSERT INTO " + f_vasi + "_virtuemart_categories_el_gr (virtuemart_category_id,category_name,category_description,metadesc,metakey,customtitle,slug) VALUES"
                sql = sql + "( " + m_ID + ",'" + c + "','','','','','" + c + "');"
                gdblag.Execute sql
                      
                sql = "UPDATE " + f_vasi + "_virtuemart_category_categories set category_child_id=" + m_ID + " where id=" + m_ID
                gdblag.Execute sql, k
            End If
                    
        End If

        On Error GoTo 0

        categories

        'MsgBox "ok"
        'T1.Refresh
        ' επιλεγω το καινουριο
        For k = 1 To UBound(names) - 1

            If Val(mID(names(k), 2, 5)) = Y Then
                T1.Nodes.Item(Y).Selected = True
                T1.Nodes.Item(Y).BackColor = vbCyan
                T1.SetFocus
                T1.Nodes.Item(Y).Expanded = True
                T1.Refresh
           
                Exit For

            End If
        
        Next

    End If

    Exit Sub

T1_KeyUp_err:

    'err.Description
    '[MySQL][ODBC 5.1 Driver][mysqld-5.5.34-MariaDB]MySQL server has gone away
    If Err.Number = -2147467259 Then
        gdblag.Close
        gdblag.Open f_connect

        Resume Next

    End If

    Resume Next

End Sub

Private Sub add_Click(index As Integer)

    Dim Y

    Dim nam

    Dim c    As String

    Dim sql  As String

    Dim k    As Long  '240 agglia 131 isp 64    '

    Dim m_ID As String

    Dim R    As New ADODB.Recordset

    'If T1.Nodes.Count > 0 And KeyCode = 13 Then 'prevents an error if TV is empty
    '   Y = T1.SelectedItem.Index
    '   nam = T1.SelectedItem.Text
    '   Me.Caption = nam & " - " & str(Y) + " names=" + names(Y)
    '   Set T1.DropHighlight = T1.SelectedItem
    '
    ''   m_categ = mID(names(Y), 2, 5)
    '   List2.AddItem names(Y)
    'End If
    '

    'If T1.Nodes.Count > 0 And KeyCode = 46 Then 'DELETE SBHNΩ ΚΑΤΗΓΟΡΙΑ
    '   Y = T1.SelectedItem.Index
    '   'Label1.Caption = Str(y)
    '   nam = T1.SelectedItem.Text
    '   Me.Caption = nam & " - " & str(Y) + " names=" + names(Y) + str(FC_id(Y))
    '
    '   ' FC_id (Y)
    ''ψαχνω να δω αν υπάρχουν κατηγοριες με πατερα αυτο που θα σβήσω
    ' R.Open "SELECT count(*) from " + f_vasi + "_virtuemart_category_categories where category_parent_id=" + str(FC_id(Y)), gdblag, adOpenDynamic, adLockOptimistic
    ' If R(0) > 0 Then
    '    MsgBox "αδυνατη η διαγραφή γιατί υποδιαιρείται η κατηγορία"
    '    Exit Sub
    ' End If
    ' R.Close
    '
    ' 'ψαχνω να δω αν υπάρχουν ειδη με κατηγορια αυτο που θα σβήσω
    ' R.Open "SELECT count(*) from " + f_vasi + "_virtuemart_product_categories where virtuemart_category_id=" + str(FC_id(Y)), gdblag, adOpenDynamic, adLockOptimistic
    ' If R(0) > 0 Then
    '    MsgBox "αδυνατη η διαγραφή γιατί υπάρχουν είδη με αυτή την κατηγορία"
    '    Exit Sub
    ' End If
    ' R.Close
    '    sql = "DELETE FROM  " + f_vasi + "_virtuemart_category_categories  where id=" + str(FC_id(Y))
    '                      gdblag.Execute sql, k
    '
    ''    sql = "DELETE FROM  " + f_vasi + "_virtuemart_categories_el_gr  where virtuemart_category_id=" + str(FC_id(Y))
    '                   '   gdblag.Execute sql, k
    '
    '    sql = "DELETE FROM  " + f_vasi + "_virtuemart_categories_el_gr  where virtuemart_category_id=" + str(FC_id(Y))
    '                      gdblag.Execute sql, k
    '  categories
    '       'MsgBox "ok"
    '       'T1.Refresh
    '        ' επιλεγω το καινουριο
    '        For k = 1 To UBound(names) - 1
    '        If Val(mID(names(k), 2, 5)) = Y Then
    '           T1.Nodes.Item(Y).Selected = True
    '           T1.Nodes.Item(Y).BackColor = vbCyan
    '           T1.SetFocus
    '           T1.Nodes.Item(Y).Expanded = True
    '           T1.Refresh
    '
    '           Exit For
    '        End If
    '        Next
    'End If

    'If T1.Nodes.Count > 0 And KeyCode = 45 Then

    '---------------------------------  insert ΠΡΟΣΘΕΤΩ ΚΑΤΗΓΟΡΙΑ
    Dim m_ID0

    If index = 1 Then
        ' βρισκω το ID του προηγούμενου δέντρου
        m_ID0 = MARKA.ItemData(MARKA.ListIndex)
    ElseIf index = 2 Then
        m_ID0 = List(1).ItemData(List(1).ListIndex)
    Else
        m_ID0 = 0
    End If

    Dim C1 As String

    Dim c2 As String

    Dim c3 As String

    Dim c4 As String

    Dim n2 As Integer

    On Error GoTo aa_click_err
   
    'T1.Nodes.Clear
    If index = 2 Then
        C1 = InputBox("Δώσε μοντελο/τύπο(30x) π.χ. Agila(H00) 1.2 16V ", "Εισαγωγή τύπου")
        c2 = InputBox("Δώσε έτος κατασκευής(20x) π.χ. 2000-2004 ")
        c3 = InputBox("Δώσε κωδ.κινητήρα(15x) π.χ. Ζ 12 ΧΕ ")
        c4 = InputBox("Δώσε χωρητικότητα(12x) π.χ. 1199cc")
        c = Left(C1 + Space(30), 30) + Left(c2 + Space(20), 20) + Left(c3 + Space(15), 15) + Left(c4 + Space(12), 12)
    Else
        c = InputBox("Δώσε τον τίτλο της κατηγορίας", "Εισαγωγή τύπου")
    End If
                   
    sql = "INSERT INTO " + f_vasi + "_virtuemart_categories (virtuemart_vendor_id,category_template,category_layout,category_product_layout,products_per_row,limit_list_start,limit_list_step,limit_list_max,limit_list_initial,hits,metarobot,metaauthor,ordering,shared,published,created_on,created_by,modified_on,modified_by,locked_on,locked_by) VALUES (1,'0','0','0',0,NULL,'0',NULL,0,0,'','',0,0,1,'2013-11-10 21:50:16',42,'2013-11-10 21:50:16',42,'2013-11-10 21:50:16',42);"
    gdblag.Execute sql, n2

    If n2 = 0 Then
        MsgBox "δεν έγινε η εισαγωγή"
    End If

    DoEvents
                                        
    R.Open "SELECT MAX(virtuemart_category_id) from " + f_vasi + "_virtuemart_categories", gdblag, adOpenDynamic, adLockOptimistic
    m_ID = str(R(0))
                   
    sql = "INSERT INTO " + f_vasi + "_virtuemart_category_categories (category_parent_id,category_child_id,ordering)"
    sql = sql + " VALUES (" + str(m_ID0) + "," + str(m_ID) + ",0);"
    gdblag.Execute sql, k

    If k > 0 Then
                      
        ' R.Open "SELECT MAX(id) from " + f_vasi + "_virtuemart_category_categories", gdblag, adOpenDynamic, adLockOptimistic
        ' m_ID = str(R(0))
                      
        sql = "INSERT INTO " + f_vasi + "_virtuemart_categories_el_gr (virtuemart_category_id,category_name,category_description,metadesc,metakey,customtitle,slug) VALUES"
        sql = sql + "( " + m_ID + ",'" + c + "','','','','','" + c + "');"
        gdblag.Execute sql
                      
        'sql = "UPDATE " + f_vasi + "_virtuemart_category_categories set category_child_id=" + m_ID + " where id=" + m_ID
        'gdblag.Execute sql, k
    End If

    On Error GoTo 0

    categories
        
    ' επιλεγω το καινουριο
    '        For k = 1 To UBound(names) - 1
    '            If Val(mID(names(k), 2, 5)) = Y Then
    '               T1.Nodes.Item(Y).Selected = True
    '               T1.Nodes.Item(Y).BackColor = vbCyan
    '               T1.SetFocus
    '               T1.Nodes.Item(Y).Expanded = True
    '               T1.Refresh
    '               Exit For
    '             End If
    '        Next

    Exit Sub

aa_click_err:

    'err.Description
    '[MySQL][ODBC 5.1 Driver][mysqld-5.5.34-MariaDB]MySQL server has gone away
    If Err.Number = -2147467259 Then
        gdblag.Close
        gdblag.Open f_connect

        Resume Next

    End If

    Resume Next

End Sub

Private Sub cmdANAZHTHSH_Click()
    NEA.Visible = False
    kod.Enabled = True
    ono.Enabled = True
    neo.Enabled = True
 
End Sub

Private Sub cmdDelete_Click()

    Dim ANS As Integer

    ANS = MsgBox("Να διαγραφεί το είδος " + kod.Text + " " + FONO.Text, vbYesNo)

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

    Gdb.Execute "delete FROM EID WHERE KOD='" + kod.Text + "'", k

    Gdb.Execute "delete from  BARCODES WHERE KOD='" + kod.Text + "'", k

    If k > 0 Then
  
        Gdb.Execute "delete PRODUCT_TO_CATEGORIES where KOD='" + kod + "'", k
        Gdb.Execute "DELETE FROM PRODUCT_ATTRIBUTES where KOD='" + kod + "' ", k
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
    gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=www.toys-shop.gr;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
    R.Open "select * from products where products_model='" + kod + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

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

Function dele_site_lagakis() As Integer  '1=ok  0=not ok

    Dim R    As New ADODB.Recordset

    Dim m_ID As String

    R.Open "select * from " + f_vasi + "_virtuemart_products where product_sku='" + kod.Text + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

    If R.EOF Then
        MsgBox "δεν βρέθηκε στο site ο κωδικός " + kod
        dele_site_lagakis = 0

        Exit Function

    End If

    m_ID = str(R("virtuemart_product_id"))
    R.Close

    gdblag.Execute "delete from " + f_vasi + "_virtuemart_products_el_gr        where virtuemart_product_id=" + m_ID + " limit 1"
    gdblag.Execute "delete from " + f_vasi + "_virtuemart_product_categories    where virtuemart_product_id=" + m_ID + " limit 20"
   
    Dim sql           As String

    Dim site_foto(16) As String

    Dim N             As Integer

    N = 0
    sql = "select * from " + f_vasi + "_virtuemart_medias   where  virtuemart_media_id in "
    sql = sql + " (select virtuemart_media_id  FROM " + f_vasi + "_virtuemart_product_medias where virtuemart_product_id=" + m_ID + ") "

    R.Open sql, gdblag, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        N = N + 1
        site_foto(N) = CNull(R!file_title)
        R.MoveNext
    Loop

    R.Close

    'σβησιμο απο sql τις αναφορές εικόνων
    sql = "delete from " + f_vasi + "_virtuemart_medias   where  virtuemart_media_id in "
    sql = sql + " (select virtuemart_media_id  FROM " + f_vasi + "_virtuemart_product_medias where virtuemart_product_id=" + m_ID + ") "
    gdblag.Execute sql

    gdblag.Execute "delete from " + f_vasi + "_virtuemart_product_medias where virtuemart_product_id=" + m_ID

    gdblag.Execute "delete from " + f_vasi + "_virtuemart_products    where virtuemart_product_id=" + m_ID + " limit 1"

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

Private Sub cmdDIPLA_Click()

    sql = "INSERT INTO " + f_vasi + "_virtuemart_categories (virtuemart_vendor_id,category_template,category_layout,category_product_layout,products_per_row,limit_list_start,limit_list_step,limit_list_max,limit_list_initial,hits,metarobot,metaauthor,ordering,shared,published,created_on,created_by,modified_on,modified_by,locked_on,locked_by) VALUES (1,'0','0','0',0,NULL,'0',NULL,0,0,'','',0,0,1,'2013-11-10 21:50:16',42,'2013-11-10 21:50:16',42,'2013-11-10 21:50:16',42);"

    For k = 3542 To 3599
        gdblag.Execute sql
    Next

    Exit Sub

    Dim R  As New ADODB.Recordset

    Dim r2 As New ADODB.Recordset

    R.Open "select * from DOKCATEGTEL WHERE AA=3 ORDER BY PARENT", Gdb, adOpenDynamic, adLockOptimistic

    Dim oldParent As Long

    oldParent = 99999

    Dim NAMEPARENT As String

    Dim PAPOYS     As Long

    Dim m_ID       As Long, M_PARENT As Long

    'Dim sql As String
    PAPOYS = 0

    Dim c As String
   
    Gdb.Execute "UPDATE DOKCATEGTEL SET UPD=0 "
   
    Do While Not R.EOF
      
        ' sql = "SELECT virtuemart_category_id as ID,category_name as NAME FROM " + f_vasi + "_virtuemart_categories_el_gr  where virtuemart_category_id=" + str(r!ID)
        r2.Open "select top 1 * from DOKCATEGTEL WHERE UPD=0 AND AA=4 AND PARENT=" + str(R!ID), Gdb, adOpenDynamic, adLockOptimistic

        If r2.EOF Then
            c = R!Name
            r2.Close
        Else
            c = R!Name + " " + r2!Name
            Gdb.Execute "UPDATE DOKCATEGTEL SET UPD=1 WHERE ID=" + str(r2!ID)
            r2.Close
            gdblag.Execute "UPDATE  " + f_vasi + "_virtuemart_categories_el_gr SET category_name='" + c + "' WHERE virtuemart_category_id=" + str(R!ID)
        End If
     
        Me.Caption = str(R!ID)
        R.MoveNext
      
        DoEvents
      
    Loop

    MsgBox "OK"

    Exit Sub

    'Dim r As New ADODB.Recordset
    '
    'Dim r2 As New ADODB.Recordset
    '
    '   r.Open "select * from DOKCATEGTEL WHERE AA=2 ORDER BY PARENT", Gdb, adOpenDynamic, adLockOptimistic
    '   Dim oldParent As Long
    '   oldParent = 99999
    '   Dim NAMEPARENT  As String
    '   Dim PAPOYS As Long
    '   Dim m_ID As Long, M_PARENT As Long
    '   'Dim sql As String
    '   PAPOYS = 0
    '   Dim C As String
    '
    '   Do While Not r.EOF
    '
    '      sql = "SELECT virtuemart_category_id as ID,category_name as NAME FROM " + f_vasi + "_virtuemart_categories_el_gr  where virtuemart_category_id=" + str(r!ID)
    ''      r2.Open sql, gdblag, adOpenDynamic, adLockOptimistic
    ''      C = r2!Name
    ''       r2.Close
    '      gdblag.Execute "UPDATE  " + f_vasi + "_virtuemart_categories_el_gr SET category_name='" + r!Name + "' WHERE virtuemart_category_id=" + str(r!ID)
    '
    '     Me.Caption = str(r!ID)
    '    r.MoveNext
    '
    '      DoEvents
    '
    '
    '
    '   Loop
    'MsgBox "OK"
    '
    'Exit Sub

    'Dim r As New adodb.Recordset
    '
    'Dim r2 As New adodb.Recordset
    '
    '    sql = "SELECT a.virtuemart_category_id as ID,b.category_parent_id as PARENT,a.category_name as NAME FROM " + f_vasi + "_virtuemart_categories_el_gr a inner join " + f_vasi + "_virtuemart_category_categories b on a.virtuemart_category_id=b.category_child_id"
    '   r.Open "select * from CATEGTEL WHERE AA=4 ORDER BY PARENT", Gdb, adOpenDynamic, adLockOptimistic
    '
    '   Dim oldParent As Long
    '   oldParent = 99999
    '
    '   Dim NAMEPARENT  As String
    '   Dim PAPOYS As Long
    '   Dim M_ID As Long, M_PARENT As Long
    '   'Dim sql As String
    '   PAPOYS = 0
    '
    '   Do While Not r.EOF
    '    If r!Parent = oldParent Then
    '       ' αφηνω την ιδια εγγραφη αλλα την αλλαζω ονομα και σαν πατέρα βάζω τον παππου
    '
    '     'ΒΑΖΩ ΣΤΟ ΟΝΟΜΑ ΤΟΥ yioy   =  ΠΑΤΕΡΑ + ΤΟ ΟΝΟΜΑ ΤΟΥ ΓΙΟΥ
    '      gdblag.Execute "UPDATE  " + f_vasi + "_virtuemart_categories_el_gr SET category_name='" + NAMEPARENT + " " + r!Name + "' WHERE virtuemart_category_id=" + str(r!ID)
    '
    '      'bazo san patera ton pappoy
    '      gdblag.Execute "update   " + f_vasi + "_virtuemart_category_categories set category_parent_id=" + str(PAPOYS) + " where id=" + str(r!ID)
    '    Else
    '      ' ΒΡΙΣΚΩ ΤΗΝ ΠΕΡΙΓΡΑΦΗ ΤΟΥ ΠΑΤΕΡΑ kai ton  kodiko toy papoy
    '      sql = "SELECT a.virtuemart_category_id as ID,b.category_parent_id as PARENT,a.category_name as NAME FROM " + f_vasi + "_virtuemart_categories_el_gr a inner join " + f_vasi + "_virtuemart_category_categories b on a.virtuemart_category_id=b.category_child_id where virtuemart_category_id=" + str(r!Parent)
    '      r2.Open sql, gdblag, adOpenDynamic, adLockOptimistic
    '      NAMEPARENT = r2!Name
    '      PAPOYS = r2!Parent
    '      r2.Close
    '      'ΒΑΖΩ ΣΤΟ ΟΝΟΜΑ ΤΟΥ ΠΑΤΕΡΑ + ΤΟ ΟΝΟΜΑ ΤΟΥ ΓΙΟΥ
    '      gdblag.Execute "UPDATE  " + f_vasi + "_virtuemart_categories_el_gr SET category_name=category_name+'" + r!Name + "' WHERE virtuemart_category_id=" + str(r!Parent)
    '
    '      'σβηνω το γιο
    '      gdblag.Execute "delete from   " + f_vasi + "_virtuemart_categories_el_gr  where virtuemart_category_id=" + str(r!ID)
    '      gdblag.Execute "delete from   " + f_vasi + "_virtuemart_category_categories  where id=" + str(r!ID)
    '    End If
    '      oldParent = r!Parent ' για να ελεγξω μηπως και ο επόμενος έχει τον ίδιο πατέρα
    '      r.MoveNext
    '      DoEvents
    '      Me.Caption = str(oldParent)
    '   Loop
    'MsgBox "OK"

    'SELECT * FROM `vasi3`.`abie3_virtuemart_categories_el_gr`;INSERT INTO `vasi3`.`abie3_virtuemart_categories_el_gr`
    '(`virtuemart_category_id`,
    '`category_name`,
    '`category_description`,
    '`metadesc`,
    '`metakey`,
    '`customtitle`,
    '`slug`)
    'Values
    '(
    '{virtuemart_category_id: INT UNSIGNED},
    '{category_name: CHAR},
    '{category_description: VARCHAR},
    '{metadesc: VARCHAR},
    '{metakey: VARCHAR},
    '{customtitle: CHAR},
    '{slug: CHAR}
    ');

    'SELECT * FROM `vasi3`.`abie3_virtuemart_category_categories`;INSERT INTO `vasi3`.`abie3_virtuemart_category_categories`
    '(`id`,
    '`category_parent_id`,
    '`category_child_id`,
    '`ordering`)
    'Values
    '(
    '{id: INT UNSIGNED},
    '{category_parent_id: INT UNSIGNED},
    '{category_child_id: INT UNSIGNED},
    '{ordering: INT}
    ');

End Sub

Private Sub cmdPreview_Click()
    WebBrowser1.SetFocus

    SendKeys ("admin")
    SendKeys ("{TAB}")
    SendKeys ("lagakis")
    SendKeys ("{Enter}")

End Sub

'Private Sub cmd00000ΕΝΗΜΕΡΩΣΗ_Click()

Private Sub cmdΕΝΗΜΕΡΩΣΗ_Click(index As Integer)
    Dim a       As String

    Dim m_categ As String

    Dim Y       As Long

    Dim m_new   As Integer
 
    On Error GoTo cmd_enhmer_Err
 
    'If neo.Enabled = True Then
    
10  If FONO = "" Then Exit Sub
 
    'Else
    
12  kod.Enabled = True
 
    'On Error Resume Next
  
    '30 If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty
    '   On Error Resume Next
    '   Y = T1.SelectedItem.Index
    '
    '   Y = T1.SelectedItem.Index
    '   m_categ = mID(names(Y), 2, 4)
    '   Set T1.DropHighlight = T1.SelectedItem
    'Else
    '   MsgBox "Δεν επιλέξατε κατηγορία "
    '   Exit Sub
    'End If

14  If Val(Left(Combo2.Text, 2)) < 1 Then
        MsgBox "Δεν επιλέξατε κατηγορία ΦΠΑ"

        Exit Sub

16  End If
 
    Dim energo As String

18  energo = "0"

20  If Check1.Value = vbChecked Then
22      energo = "1"
24  End If
 
        If Val(WEIGHT.Text) = 0 Then
            WEIGHT.Text = "1"
        End If
 
        Dim kor   As String, B As String

        Dim manuf As String

        'kor = str(f_fylo(fylo.ListIndex))
        'manuf = str(f_kataskeyasths(kataskeyasths.ListIndex))
 
        'End If
        Dim SS

        Gdb.Execute "UPDATE EID SET FPA=" + Left(Combo2.Text, 2) + " WHERE KOD='" + kod + "'", SS
 
        If Check1.Value = vbChecked Then DIATH = "1" Else DIATH = "0"
 
        Gdb.Execute "UPDATE EID SET DIATHESIMO=" + DIATH + " WHERE KOD='" + kod + "'", SS

        Dim R As New ADODB.Recordset

        Gdb.Execute "delete from PRODUCT_TO_CATEGORIES where KOD='" + kod + "'"

        For k0 = 0 To List2.ListCount - 1
            ' m_categ = str(Val(mID(List2.List(k0), 2, 5)))
            m_categ = str(Val(mID(List2.List(k0), 2, 5)))
            Gdb.Execute "INSERT INTO PRODUCT_TO_CATEGORIES (KOD,PRODUCT_ID,CATEGORIES_ID) VALUES ('" + kod + "'," + str(F_id_LOCAL) + "," + m_categ + ");"
        Next

        '   Gdb.Execute "INSERT INTO PRODUCT_TO_CATEGORIES (KOD,PRODUCT_ID,CATEGORIES_ID) VALUES ('" + KOD + "'," + str(F_id_LOCAL) + "," + m_categ + ");"
   
        Gdb.Execute "delete from GNISIOI_KODIKOI where KOD='" + kod + "'", SS

        For k0 = 0 To List3.ListCount - 1

            If Len(List3.List(k0)) > 2 Then
                Gdb.Execute "INSERT INTO GNISIOI_KODIKOI (KOD,GNISIOS) VALUES ('" + kod + "','" + List3.List(k0) + "');", SS
            End If

        Next

        'PRODUCT_TO_CATEGORIES ( KOD VARCHAR(16) NOT NULL,PRODUCT_ID int NOT NULL,CATEGORIES_ID INT NOT NULL )

        '201    DUM = ADD_FIELD("EID", "DIATHESIMO", "BOOLEAN")
        '202 DUM = ADD_FIELD("EID", "HME_DIATH", "DATE")
        '203 DUM = ADD_FIELD("EID", "MANUFACTURERS_ID", "INT")
        '204 DUM = ADD_FIELD("EID", "ERGOSTASIO", "VARCHAR(30)")
        '205 DUM = ADD_FIELD("EID", "ADDITIONALPRICE", "DECIMAL(10,2)")
        '206 DUM = ADD_FIELD("EID", "WEIGHT", "DECIMAL(10,2)")
        '
        '207 DUM = ADD_FIELD("EID", "FYLO", "INT")
        '208 DUM = ADD_FIELD("EID", "XARAKTIRAS", "INT")
    
        'Gdb.Execute "DELETE FROM PRODUCT_ATTRIBUTES where KOD='" + KOD + "' ", k
        'Dim sort As Integer
        'For k = 0 To hlikies.ListCount - 1
        '  If hlikies.Selected(k) Then
        '      Gdb.Execute "INSERT INTO PRODUCT_ATTRIBUTES " _
        '      & "(KOD,OPTIONS_ID,OPTIONS_VALUES_ID,PRODUCT_ID) " _
        '      & " VALUES ('" + KOD + "',1," + str(f_hlikies(k)) + "," + str(EID.Recordset("ID")) + ");"
        '  End If
        'Next k
    
        'Gdb.Execute "INSERT INTO BARCODES (KOD,ERG) VALUES ('" + KOD + "','" + KOD + "')"
 
        Dim ANS As Integer

        Dim mc  As Integer
 
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
 
        ' DUM = ADD_FIELD("EID", "ERGOSTASIO", "VARCHAR(30)")
        ' DUM = ADD_FIELD("EID", "YOU_TUBE", "VARCHAR(100)")
 
        On Error GoTo 0

        ' EID.Recordset.Move 0
  
        'ΤΟΠΟΘΕΤΩ ΤΗΝ ΚΥΡΙΑ ΚΑΤΗΓΟΡΙΑ(ΑΝΤΑΛΛΑΚΤΙΚΟ) ΚΑΙ ΤΟ ΜΕΡΟΣ-ΑΥΟΚΙΝΗΤΟΥ
 
        Dim APOS       As String

        'APOS=STR(VAL(TEXT9.TEXT))
  
        Dim meidos     As String

        Dim mMON       As String

        Dim mKATHGORIA As String
  
100     mMON = Replace("" + Combo1.Text, "'", "''")
  
200     meidos = Replace("" + EIDOS.Text, "'", "''")
300     mKATHGORIA = Replace("" + KATHGORIA.Text, "'", "''")
  
        If NEA.Visible = True Then
500         UPD_EID "APOS", POS.Text, "N"
600         UPD_EID "POS", POS.Text, "N"
        End If
     
        ' Gdb.Execute "UPDATE EID  SET ONO=" + APOST(FONO.Text) + " WHERE KOD='" + KOD + "'"
700     UPD_EID "ONO", FONO.Text, "C"
800     UPD_EID "KODERG", KODERG.Text, "C"
900     UPD_EID "ERGOSTASIO", ERGOSTASIO.Text, "C"
920     UPD_EID "ONO2", ONO2.Text, "C"
930     UPD_EID "WEIGHT", str(WEIGHT.Text), "N"
950     UPD_EID "UES", UES.Text, "C"
970     UPD_EID "XTI", XTI.Text, "N"
980     UPD_EID "LTI5", lti5.Text, "N"
990     UPD_EID "WEBPRICE", WEBPRICE.Text, "N"
  
992     UPD_EID "TEXT_PRODUCTS_ZIPCODE", mKATHGORIA, "C"
  
994     UPD_EID "YOU_TUBE", meidos, "C"
  
996     UPD_EID "MON", mMON, "C"
        UPD_EID "MEMO", msEditor21.Text, "C"
        'UPDATE_CONTROLS

        UPD_EID "PIC0", PIC(0).Caption, "C"
        UPD_EID "PIC1", PIC(1).Caption, "C"
        UPD_EID "PIC2", PIC(2).Caption, "C"
        UPD_EID "PIC3", PIC(3).Caption, "C"
        UPD_EID "PIC4", PIC(4).Caption, "C"

        UPD_EID "PIC5", PIC(5).Caption, "C"
        UPD_EID "PIC6", PIC(6).Caption, "C"
  
        'ERGOSTASIO
    
        NEA.Visible = False
    
        kod.Enabled = True
        ono.Enabled = True
        upload1.Enabled = True
     
        neo.Enabled = True
        Me.Caption = "Η ΑΠΟΘΗΚΕΥΣΗ ΕΓΙΝΕ"
        'ΝΑ MHN ΜΠΟΡΕΙ ΝΑ ΒΑΛΕΙ ΑΡΧΙΚΟ ΥΠΟΛΟΙΠΟ
        POS.Enabled = False

        Exit Sub

cmd_enhmer_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.APOT2ROULIS.cmdenhmer " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        f_err = 1

        Resume Next

        '</EhFooter>

End Sub



'End Sub

Private Sub cmdΕΞΟΔΟΣ_Click()
    'bCancel = True
    'DoEvents
    'gdblag.Close
 
    Unload Me
    'f_exit = True
    'Me.Hide
 
End Sub

Private Sub Command14_Click(index As Integer)

        '<EhHeader>
        On Error GoTo Command14_Click_Err

        '</EhHeader>
        Dim F_EIKONA As Long

100     F_EIKONA = 11
110     CommonDialog1.FILENAME = "*.jpg"
120     CommonDialog1.ShowOpen

        Dim m As String

        On Error Resume Next

130     m = "a" + kodlabel.Caption + "-" + Format(index, "0") ' Text1(0).Text

140     If F_EIKONA = 11 Then
150         m = m + ".JPG"
        End If

160     FileCopy CommonDialog1.FILENAME, F_FAKEL_KOINOS + "\images\" + m
    
170     foto2(index).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + m)
            
180     StretchSourcePictureFromPicture apot2ROULIS.foto2(index).Picture, apot2ROULIS.foto2(index)
        PIC(index).Caption = m
    
        '<EhFooter>
        Exit Sub

Command14_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command14_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2ROULIS.Command14_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command10_Click()

    '-----------------------------ετικετα -----------------------------------------------------------------------
    Dim DB   As Database

    Dim R    As New ADODB.Recordset

    Dim PROM As String

    Dim hme  As String

    On Error GoTo err_db

    R.Open "select TOP 1 PELKOD,HME FROM EGGTIM WHERE KODE='" + kod.Text + "' AND UPPER(LEFT(ATIM,1))='G' ORDER BY HME DESC", Gdb, adOpenDynamic, adLockOptimistic
   
    PROM = "    "

    If R.EOF Then
        hme = "    "
    Else
        hme = Format(R(1), "DDMMYY")

        If Not R.EOF Then
            If Not IsNull(R(0)) Then
                PROM = R(0)
            End If
        End If
    End If
   
    Set DB = OpenDatabase("c:\mercvb\reports\reports.mdb")
    DB.Execute "UPDATE EIDETIK SET LTI5=" + lti5.Text + ",KODERG='" + KODERG.Text + "-" + PROM + "-" + hme + "',ONO='" + FONO.Text + "',KOD='" + kod.Text + "'"
    DB.Close
   
    Dim POSO As Long

    POSO = InputBox("ΠΟΣΕΣ ΕΤΙΚΕΤΕΣ ΘΕΛΕΤΕ;", "", 1)

    Dim k As Long

    For k = 1 To POSO
        CrystalReport1.Action = 1  ' ETIK.RPT
    Next
   
    Exit Sub
   
err_db:
    MsgBox "λαθος στην εκτύπωση. Σειρά " + str(Erl)
   
End Sub

Private Sub Command11_Click()
    APOT2.kod.Text = Trim(kod.Text)
    APOT2.SHOW

    On Error Resume Next

    SendKeys "~"    'ENTER
End Sub

Private Sub Command12_Click()

        '<EhHeader>
        On Error GoTo Command12_Click_Err

        '</EhHeader>
        Dim F33 As New apot21

110     F33.Text1.Text = Trim(kod.Text) ' Grid1.TextMatrix(Grid1.row, 1)

120     F33.Text2.Text = FONO.Text ' Grid1.TextMatrix(Grid1.row, 2)
130     F33.Label1.Caption = "1" ' Left(APOTH1.Text, 1)
140     F33.Label2.Caption = "0"

150     F33.SHOW

160     APOT2.Hide

        '<EhFooter>
        Exit Sub

Command12_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.Grid1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot2roulis.Command12_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

    Dim R    As New ADODB.Recordset

    Dim rLOC As New ADODB.Recordset

    gdblag.Close  'SERVER=188.40.85.131
    gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=www.toys-shop.gr;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
    R.Open "select * from products WHERE products_model<>'' order by products_model", gdblag, adOpenDynamic, adLockOptimistic

    Dim N    As Long

    Dim KODE As String

    Do While Not R.EOF()
        KODE = Replace(R!PRODUCTS_MODEL, "'", "~")
        'rLOC.Open "select * from EID WHERE KOD='" + KODE + "'", Gdb, adOpenDynamic, adLockOptimistic
        '   If rLOC.EOF Then
        '       Gdb.Execute "INSERT INTO EID (KOD) VALUES ('" + KODE + "')"
        '   End If
        '   rLOC.Close
   
        If Not IsNull(R!products_youtube) Then
            Gdb.Execute "UPDATE EID SET YOU_TUBE='" + Replace(R!products_youtube, "'", "''") + "' WHERE KOD='" + KODE + "'"
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

    R.Open "select * from products_attributes ", gdblag, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        rLOC.Open "select * from EID WHERE ID_WEB=" + str(R!PRODUCTS_ID), Gdb, adOpenDynamic, adLockOptimistic

        If Not rLOC.EOF Then
            Gdb.Execute "INSERT INTO PRODUCT_ATTRIBUTES " & "(KOD,OPTIONS_ID,OPTIONS_VALUES_ID,PRODUCT_ID) " & " VALUES ('" + rLOC!kod + "',1," + str(R!options_values_id) + ",0);"
        End If

        rLOC.Close
        N = N + 1

        DoEvents
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

    Dim a As String, B As String

    a = InputBox("ΑΡΙΘΜΟΣ ΤΙΜΟΛ.ΑΓΟΡΑΣ")
    B = InputBox("ΗΜΕΡ/ΝΙΑ ΤΙΜΟΛ.ΑΓΟΡΑΣ", , Format(Now, "DD/MM/YYYY"))

    EGGTIM2.ConnectionString = gConnect
    EGGTIM2.RecordSource = "SELECT KODE FROM EGGTIM WHERE ATIM LIKE 'G%" + a + "%' AND CONVERT(CHAR(10),HME,103)='" + B + "' ORDER BY PROOD_AJ"
    EGGTIM2.Refresh

    If Not EGGTIM2.Recordset.EOF Then
        CMDmovenext.BackColor = vbGreen
        CMDmoveprevious.BackColor = vbGreen
        kod.Text = EGGTIM2.Recordset("KODE")
        kod.SetFocus
        FONO.SetFocus
    
    Else
    
        CMDmovenext.BackColor = vbRed
        CMDmoveprevious.BackColor = vbRed
    End If

End Sub

Private Sub Command5_Click()

    Dim rs As New ADODB.Recordset

    Dim R  As New ADODB.Recordset

    Dim energo, mq, mTimh

    Dim k As Integer

    Dim TIM

    TIM = GetCurrentTime()

    On Error GoTo command5_out

    R.Open "select * from EID WHERE LEN(PIC0)>2", Gdb, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
   
        If IsNull(R!WEBPRICE) Then
            If IsNull(R!lti5) Then
                mTimh = "0"
            Else
       
                mTimh = Format(R!lti5 / ((100 + g_Fpa(R!FPA)) / 100), "####.00")
       
            End If

        Else
            mTimh = Format(R!WEBPRICE / ((100 + g_Fpa(R!FPA)) / 100), "####.00")
        End If
   
        If IsNull(R!POS) Then
            mq = "0"
            energo = "0"
        Else
            mq = Format(R!POS, "####.00")
            energo = "1"

            If Val(mTimh) < 20 Then
                If Val(mq) < 5 Then
                    energo = "0"
                End If
            End If
        End If
   
        List2.AddItem R!kod, 0
   
        DoEvents
   
        gdblag.Execute "update products set products_price=" + mTimh + ",products_quantity=" + mq + ",products_status=" + energo + " where products_model='" + R!kod + "' limit 1"
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

    MsgBox "λαθος στο ειδος " + R!kod

End Sub

Private Sub Command6_Click()

    On Error Resume Next

    List3.RemoveItem List3.ListIndex
End Sub

Private Sub Command7_Click()

    Dim gg2

    Dim m_new As String
  
    m_new = InputBox("Δώσε νέο τύπο (έιδος) ανταλλακτικού")

    If Len(Trim(m_new)) < 2 Then
        MsgBox "Κενή εγγραφή. Δεν κατεχωρήθη"

        Exit Sub

    End If
  
    If InStr(m_new, ";") > 0 Then
        MsgBox "Μη επιτρεπτός χαρακτήρας: ;" + Chr(13) + "Δεν κατεχωρήθη"

        Exit Sub

    End If
  
    Dim yparxon As String

    yparxon = ""

    If typos_exist(m_new, yparxon) Then
        MsgBox "Υπάρχει ήδη τύπος με το όνομα " + m_new

        Exit Sub

    End If
  
    yparxon = yparxon + m_new + ";"

    Dim k As Long
  
    gdblag.Execute "update " + f_vasi + "_virtuemart_customs set custom_value='" + yparxon + "' where virtuemart_custom_id=23", k
  
    If k > 0 Then
        MsgBox "ok "
    Else
        MsgBox "δεν προστέθηκε ο τύπος"
    End If
    
    refresh_typos
    EIDOS.ListIndex = 0

End Sub

Function typos_exist(ByVal MTYPOS As String, ByRef yparxon As String) As Boolean

    Dim sql As String

    Dim R   As New ADODB.Recordset

    Dim F   As String

    sql = "SELECT * FROM " + f_vasi + "_virtuemart_customs where virtuemart_custom_id=23"
    R.Open sql, gdblag, adOpenDynamic, adLockOptimistic
    F = R!custom_value
    gg2 = Split(F, ";")
    typos_exist = False

    For k = LBound(gg2) To UBound(gg2)

        If MTYPOS = gg2(k) Then
            typos_exist = True
        End If

    Next

    R.Close
    yparxon = F

End Function

Private Sub Command8_Click()

    Dim R As New ADODB.Recordset

    R.Open "SELECT count(*) FROM " + f_vasi + "_virtuemart_product_customfields where virtuemart_custom_id=23 and custom_value='" + EIDOS.Text + "'", gdblag, adOpenDynamic, adLockOptimistic

    If R(0) > 0 Then
        MsgBox "Αδύνατη η διαγραφή γιατί υπάρχουν προϊόντα με αυτό τον τύπο"
        R.Close

        Exit Sub

    End If

    R.Close

    Dim yparxon As String

    Dim ANS3    As Integer

    ANS3 = MsgBox("Να διαγραφεί το : " + Chr(13) + EIDOS.Text, vbYesNo)

    If ANS3 = vbNo Then

        Exit Sub

    End If

    yparxon = ""

    If (Len(Trim(EIDOS.Text)) > 0) And (typos_exist(EIDOS.Text, yparxon)) Then
   
        ' einai stin arxi
        If EIDOS.Text = Left(yparxon, Len(EIDOS.Text)) Then
            yparxon = Replace(yparxon, EIDOS.Text + ";", "")
        Else  'einai sth mesi  η στο τελος
            yparxon = Replace(yparxon, ";" + EIDOS.Text + ";", ";")
        End If
   
        ' EIDOS.RemoveItem EIDOS.ListIndex
   
        gdblag.Execute "update " + f_vasi + "_virtuemart_customs set custom_value='" + yparxon + "' where virtuemart_custom_id=23", k
  
        If k > 0 Then
            MsgBox "ok διεγράφη "
        Else
            MsgBox "δεν έγινε η διαγραφή"
        End If

    Else
        MsgBox "Δεν διεγράφη "

        Exit Sub

    End If

    refresh_typos

    EIDOS.ListIndex = 0

End Sub

Private Sub Command9_Click()

    Dim gg2

    Dim m_allagi As String
  
    m_allagi = InputBox("Δώσε νέο όνομα για τον τύπο:" + EIDOS.Text)

    If Len(Trim(m_allagi)) < 2 Then
        MsgBox "Κενή εγγραφή. Δεν έγινε η αλλαγή"

        Exit Sub

    End If
  
    If InStr(m_allagi, ";") > 0 Then
        MsgBox "Μη επιτρεπτός χαρακτήρας: ;" + Chr(13) + "Δεν έγινε η αλλαγή"

        Exit Sub

    End If
  
    Dim yparxon As String

    yparxon = ""

    If (Len(Trim(EIDOS.Text)) > 0) And (typos_exist(EIDOS.Text, yparxon)) Then
   
        ' einai stin arxi
        If EIDOS.Text = Left(yparxon, Len(EIDOS.Text)) Then
            yparxon = Replace(yparxon, EIDOS.Text + ";", m_allagi + ";")
        Else  'einai sth mesi  η στο τελος
            yparxon = Replace(yparxon, ";" + EIDOS.Text + ";", ";" + m_allagi + ";")
        End If
   
        ' EIDOS.RemoveItem EIDOS.ListIndex
   
        gdblag.Execute "update " + f_vasi + "_virtuemart_customs set custom_value='" + yparxon + "' where virtuemart_custom_id=23", k
        gdblag.Execute "update " + f_vasi + "_virtuemart_product_customfields set custom_value='" + m_allagi + "' where virtuemart_custom_id=23 and custom_value='" + EIDOS.Text + "'", proionta_pou_allaksan
    
        If k > 0 Then
            MsgBox "ok η αλλαγή του τύπου."
        Else
            MsgBox "δεν έγινε η αλλαγή του τύπου."
        End If

    Else
        MsgBox "Δεν έγινε η τροποποίηση."

        Exit Sub

    End If

    refresh_typos
    EIDOS.ListIndex = 0

End Sub

Private Sub delete_Click(index As Integer)

    Dim m_ID0

    ' βρισκω το ID του προηγούμενου δέντρου
    m_ID0 = List(index).ItemData(List(index).ListIndex)
    'Else
    '  m_ID0 = 0
    'End If

    Dim ANS3 As Integer

    ANS3 = MsgBox("Να διαγραφεί  ", vbYesNo)

    If ANS3 = vbNo Then

        Exit Sub

    End If

    'ψαχνω να δω αν υπάρχουν κατηγοριες με πατερα αυτο που θα σβήσω
    Dim R As New ADODB.Recordset

    R.Open "SELECT count(*) from " + f_vasi + "_virtuemart_category_categories where category_parent_id=" + str(m_ID0), gdblag, adOpenDynamic, adLockOptimistic

    If R(0) > 0 Then
        MsgBox "αδυνατη η διαγραφή γιατί υποδιαιρείται η κατηγορία"

        Exit Sub

    End If

    R.Close
 
    'ψαχνω να δω αν υπάρχουν ειδη με κατηγορια αυτο που θα σβήσω
    R.Open "SELECT count(*) from " + f_vasi + "_virtuemart_product_categories where virtuemart_category_id=" + str(m_ID0), gdblag, adOpenDynamic, adLockOptimistic

    If R(0) > 0 Then
        MsgBox "αδυνατη η διαγραφή γιατί υπάρχουν είδη με αυτή την κατηγορία"

        Exit Sub

    End If

    R.Close
 
    sql = "DELETE FROM  " + f_vasi + "_virtuemart_category_categories  where id=" + str(m_ID0)
    gdblag.Execute sql, k

    sql = "DELETE FROM  " + f_vasi + "_virtuemart_categories_el_gr  where virtuemart_category_id=" + str(m_ID0)
    gdblag.Execute sql, k

    List(index).RemoveItem List(index).ListIndex

    categories
    'MsgBox "ok"
    'T1.Refresh
    ' επιλεγω το καινουριο
    '        For k = 1 To UBound(names) - 1
    '        If Val(mID(names(k), 2, 5)) = Y Then
    '           T1.Nodes.Item(Y).Selected = True
    '           T1.Nodes.Item(Y).BackColor = vbCyan
    '           T1.SetFocus
    '           T1.Nodes.Item(Y).Expanded = True
    '           T1.Refresh
    '
    '           Exit For
    '        End If
    '
    '
    '        Next

End Sub

Private Sub FONO_KeyUp(KeyCode As Integer, Shift As Integer)

100     If KeyCode = vbKeyReturn Then
110         KeyCode = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

End Sub

Private Sub FONO_LostFocus()

    If Len(Trim(ONO2.Text)) = 0 Then
        ONO2.Text = FONO.Text
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        'cmdΕξοδος_Click
        List1.Visible = False
        LIST1_Label.Visible = False
    End If

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     Me.AutoRedraw = True
        KeyPreview = True
102     f_pel = "mpoyg" ' "lagak"
 
        'syndesh xori dsn me connection string  ' ip toys 188.40.85.131
        ' gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=db5.papaki.gr;PORT =3306;DATABASE=vasi;UID=lagakis;Password=3898339214231"
104     On ERRROR GoTo NOCONNECT
106     f_site = Val(FINDPARAMETROI(1, "MDIFORM1", "F_SITE", "0", "ΚΩΔΙΚΟΣ SITE 1=VM 2=TOYS"))
        'f_site = 222

108     If f_site = 222 Then
110         f_connect = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
112     ElseIf f_site = 333 Then
114         f_connect = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=db9.papaki.gr;PORT =3306;DATABASE=vasi3;UID=autoroulis;Password=autoroulis2013"
        Else
116         f_connect = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=db9.papaki.gr;PORT =3306;DATABASE=lagakis_new;UID=lagakis;Password=39213921"
        End If
        
        
 On Error GoTo Form_Load_Err_fatal
118     gdblag.Open f_connect

        '104 gdblag.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=toys-shop.gr;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"

       
        
     



        On Error GoTo Form_Load_Err
120     EID.ConnectionString = gConnect

        Dim DUM

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
    
164     F_FAKEL_KOINOS = Trim(FINDPARAMETROI(1, "APOT2", "F_FAKEL_KOINOS", "C:\MERCVB", "KOINOXPHΣTOΣ ΦΑΚΕΛΟΣ ΕΙΚΟΝΩΝ Π.Χ. \\PC\MERCVB"))
    
166     If f_site = 111 Then
168         f_vasi = "" + f_vasi + ""
170     ElseIf f_site = 333 Then
172         f_vasi = "abie3"
        Else
174         f_vasi = ""
        End If
    
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

        'ΕΛΕΓΧΩ ΑΝ ΕΝΗΜΕΡΩΘΗΚΕ ΣΗΜΕΡΑ
224     R.Open "SELECT OBJECT_NAME(OBJECT_ID) AS DatabaseName, last_user_update AS HME  FROM sys.dm_db_index_usage_stats  WHERE database_id = DB_ID( 'MERCURY')  AND OBJECT_ID=OBJECT_ID('CATEGORIESANT')", Gdb, adOpenDynamic, adLockOptimistic

226     If Not R.EOF Then
228         If Not IsNull(R!hme) Then
230             If Format(R!hme, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
232                 'r.Close
                    'GoTo PARAKAMCH
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

        '---------------------- γεμισμα του treeview control ----------------------------

270     R.Open "SELECT * FROM CATEGTEL ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic
        ' 192 r.MoveFirst

        'r.Open "SELECT * FROM CATEG1", Gdb
272     Do While Not R.EOF
       
274         error = 0
            'paterades
276         mess = "id=" + LTrim(str(R!ID)) + "Parent:" + LTrim(str(R!Parent)) + " " + R!Name

278         If R!Parent = 0 Then
280             T1.Nodes.ADD , , "r" + LTrim(str(R!ID)), R!Name
            Else
    
282             T1.Nodes.ADD "r" + LTrim(str(R!Parent)), tvwChild, "r" + LTrim(str(R!ID)), (R!Name)
            End If
       
284         If error = 0 Then
286             N = N + 1
288             names(N) = "r" + Format(R!ID, "00000") + "        " + R!Name
290             FC_id(N) = R!ID
292             FC_parenT(N) = R!Parent
294             fc_names(N) = R!Name
296             fc_aa(nn) = R!AA
            End If

298         R.MoveNext
        Loop
     
300     R.Close
302     fc_max = N

        '-----------------------------  κομματι ρουλη ------------------------------------------------------------------
        ' gia ta listbox mono
304     R.Open "SELECT * FROM CATEGTEL where AA=1 ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic

        'Dim roulis(5000, 4)
        Dim i As Long

306     i = 0

308     Do While Not R.EOF
310         MARKA.AddItem R!Name, i
312         MARKA.ItemData(i) = R!ID
314         i = i + 1
316         R.MoveNext
        Loop

318     R.Close

        '---------------------- μοναδες μέτρησης ---------------------------------
320     R.Open "SELECT *FROM PINAKES WHERE TYPOS=2 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

322     Do While Not R.EOF

324         If R("typos") = 2 Then
326             Combo1.AddItem R("PERIGRAFH")
            Else

                Exit Do

            End If

328         R.MoveNext
        Loop

330     R.Close
332     Combo1.ListIndex = 0

        '============================= ΚΑΤΗΓΟΡΙΕΣ =====================================================
        ' an einai VM  ΝΑ ΠΑΡΑΚΑΜΠΤΕΙ ΧΑΡΑΚΤΗΡΕΣ,ΦΥΛΟ,ΗΛΙΚΙΕΣ
334     If f_site = 111 Or f_site = 333 Then
336         kataskeyasths.Visible = False
338         xaraktiras.Visible = False
340         fylo.Visible = False
   
342         GoTo parakamch2
        End If
     
        '=========================================χαρακτηρας ============================================
     
344     sql = "select *from extra_field_values where epf_id=1 and languages_id=4 order by  epf_value"
        '  (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (7,3,4,0,1,'?????');
    
        Dim n7 As Long

346     n7 = 0
348     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic

350     Do While Not R.EOF
352         xaraktiras.AddItem R!epf_value
354         f_xaraktiras(n7) = R!Value_ID
356         R.MoveNext
358         n7 = n7 + 1

360         If f_err = 1 Then Exit Do
        Loop

362     xaraktiras.AddItem "                "
364     f_xaraktiras(n7) = 0
366     R.Close
368     xaraktiras.ListIndex = -1
     
        '========================================= FYLO ============================================
370     sql = "select *from extra_field_values where epf_id=3 and languages_id=4 order by epf_value"
        '  (value_id,epf_id,languages_id,parent_id,sort_order,epf_value) VALUES (7,3,4,0,1,'?????');
372     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic
374     n7 = 0

376     Do While Not R.EOF
378         fylo.AddItem R!epf_value
380         f_fylo(n7) = R!Value_ID
382         n7 = n7 + 1
384         R.MoveNext

386         If f_err = 1 Then Exit Do
        Loop

388     R.Close
390     fylo.ListIndex = 0
     
        '========================================= KATASKEYASTHS ============================================
392     sql = "SELECT * FROM manufacturers order by manufacturers_name"
     
394     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic
396     n7 = 0

398     Do While Not R.EOF
400         kataskeyasths.AddItem R!manufacturers_name
402         f_kataskeyasths(n7) = R!manufacturers_id
404         n7 = n7 + 1
406         R.MoveNext

408         If f_err = 1 Then Exit Do
        Loop

410     R.Close
412     kataskeyasths.ListIndex = 0
     
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
438     Do While Not R.EOF

440         If R("typos") = 1 Then
442             Combo2.AddItem str(R("AYJON")) + " -> " + str(R("TIMH"))
            End If

444         R.MoveNext

446         If f_err = 1 Then Exit Do
        Loop

        ' mon.metrhshs
448     R.Close
     
450     Combo2.ListIndex = 1
   
        ' sql = "SELECT * FROM abie3_virtuemart_customs where virtuemart_custom_id=22  "
 
        Dim F As String

        Dim gg
 
452     sql = "SELECT * FROM " + f_vasi + "_virtuemart_customs where virtuemart_custom_id=22  "

454     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic
 
456     F = R!custom_value
458     gg = Split(F, ";")

460     For k = LBound(gg) To UBound(gg)
462         KATHGORIA.AddItem gg(k)
            'KATHGORIA.ItemData(k) = gg(k)
        Next

464     KATHGORIA.ListIndex = 0
466     R.Close

        refresh_typos
        EIDOS.ListIndex = 0
     
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
486         MsgBox "Δεν φορτώθηκαν σωστά οι πίνακες."
        End If
     
        MARKA.ListIndex = 0
488     ftp.ListIndex = 0
     
        ' Me.Print "Please use now mouse wheel to move this form."
490     Me.SHOW
        ' ProcessMessages

492     Text2.SetFocus

        '<EhFooter>
        Exit Sub
        
NOCONNECT:
        MsgBox "ΑΔΥΝΑΤΗ Η ΣΥΝΔΕΣΗ ΜΕ ΤΟ SITE. ΔΟΚΙΜΑΣΤΕ ΑΡΓΟΤΕΡΑ"
        Unload Me

        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.APOT2ROULIS.Form_Load " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        f_err = 1

        Resume Next

Form_Load_Err_fatal:
    f_err = 99
           MsgBox " αδυνατη η συνδεση με το site"
           Unload Me

        '</EhFooter>

End Sub
       
Sub refresh_typos()
        EIDOS.Clear

        Dim R As New ADODB.Recordset

        Dim F As String

        Dim k As Integer

        Dim gg2

468     sql = "SELECT * FROM " + f_vasi + "_virtuemart_customs where virtuemart_custom_id=23"
470     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic
 
472     F = R!custom_value
474     gg2 = Split(F, ";")

476     For k = LBound(gg2) To UBound(gg2)

            If Len(Trim(gg2(k))) > 0 Then
478             EIDOS.AddItem gg2(k)
            End If
              
        Next

480     'EIDOS.ListIndex = 0
482     R.Close
     
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
    WebBrowser1.Width = Picture2.Width
 
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

Private Sub foto2_DblClick(index As Integer)

        '<EhHeader>
        On Error GoTo FOTO2_DBLClick_Err

        '</EhHeader>
        Dim F_EIKONA As Long

100     F_EIKONA = 11
110     CommonDialog1.FILENAME = "*.jpg"
120     CommonDialog1.ShowOpen

        Dim m As String

        On Error Resume Next
        
130     m = Trim(mID(CommonDialog1.FILENAME, InStrRev(CommonDialog1.FILENAME, "\") + 1, 30))
        '+ kodlabel.Caption
        m = "a" + "-" + m  ' CommonDialog1.FILENAME ' Text1(0).Text

140     If F_EIKONA = 11 Then
150         ' m = m + ".JPG"
        End If
        
        If f_site = 111 Or f_site = 333 Then
            ' m = Replace(UCase(m), ".JPG", ".jpeg")
        End If
        
160     FileCopy CommonDialog1.FILENAME, F_FAKEL_KOINOS + "\images\" + m
    
170     foto2(index).Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + m)
            
180     StretchSourcePictureFromPicture apot2ROULIS.foto2(index).Picture, apot2ROULIS.foto2(index)
        PIC(index).Caption = m
    
        '<EhFooter>
        Exit Sub

FOTO2_DBLClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat2.Command14_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2ROULIS.FOTO2_DBLClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub foto2_MouseUp(index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          x As Single, _
                          Y As Single)

    Dim ANS As Integer

    If Button = 2 Then
        ANS = MsgBox("Να καταργηθεί η εικόνα;", vbYesNo)

        If ANS = vbYes Then
            foto2(index).Picture = LoadPicture()
            PIC(index).Caption = ""
        End If
    End If
   
End Sub

Private Sub kod_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then

        KeyAscii = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If
 
End Sub

Private Sub kod_LostFocus()

    Dim R As New ADODB.Recordset

    Dim zz

    Dim lti5, POS

    kod.Text = Trim(kod.Text)

    If Len(kod) > 0 Then
  
        R.Open "SELECT COUNT(*) FROM EID WHERE KOD LIKE '%" + kod.Text + "%'", Gdb, adOpenDynamic, adLockOptimistic

        If R(0) = 0 Then
            MsgBox "ΔΕΝ ΒΡΕΘΗΚΑΝ ΕΓΓΡΑΦΕΣ"
            R.Close

            Exit Sub

        End If

        If R(0) > 1 Then
            R.Close
            R.Open "SELECT * FROM EID WHERE KOD LIKE '%" + kod.Text + "%'  oRder by ONO", Gdb, adOpenDynamic, adLockOptimistic
            List1.Clear
            List1.Height = 3900
            List1.Visible = True
            LIST1_Label.Visible = True

            Dim rrono As String
   
            Do While Not R.EOF
      
                If IsNull(R!ono) Then rron = "" Else rrono = Left(R("ONO") + Space(50), 50)
      
                If IsNull(R!lti5) Then lti5 = 0 Else lti5 = R!lti5
                If IsNull(R!POS) Then POS = 0 Else POS = R!POS
      
                If IsNull(R!ERGOSTASIO) Then zz = Space(20) Else zz = Left(R!ERGOSTASIO + Space(20), 20)
      
                List1.AddItem Left(R("kod") + Space(20), 20) + rrono + zz + Right(Space(8) + Format(lti5, "####0.00"), 8) + "€" + " " + Right(Space(8) + Format(POS, "####0"), 8)
                ' List1.AddItem Left(R("KOD") + Space(20), 20) + CNull(R("ONO"))
                R.MoveNext
            Loop

            List1.SetFocus

            Exit Sub

        End If

        R.Close
        R.Open "SELECT * FROM EID WHERE KOD LIKE '%" + kod.Text + "%'", Gdb, adOpenDynamic, adLockOptimistic
        kod.Text = R("KOD")

        R.Close

        'EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + KOD.Text + "'"
        '  EID.Refresh

        '    If Len(EID.RecordSource) > 5 Then
        '
        '    If Not EID.Recordset.EOF Then
        '        Exit Sub
        '        EID.Recordset.Move 0
        '    End If
        '
        '    End If

        EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + kod + "'"
        EID.Refresh
   
        If Not EID.Recordset.EOF Then
            UPDATE_CONTROLS
        Else

            Exit Sub

        End If
   
        F_id_LOCAL = EID.Recordset("ID")
    End If

End Sub

Private Sub list_Click(index As Integer)
    MOVELIST index
   
    '124    If list(INDEX + 1).ListCount > 0 Then
    '126      list(INDEX + 1).SetFocus
    '128      list(INDEX + 1).ListIndex = 0
    '    End If
End Sub

Private Sub list_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)

100     If KeyCode <> 13 Then

            Exit Sub

        End If

        If List(index).ListCount = 0 Then

            Exit Sub

        End If

        Dim R     As New ADODB.Recordset

        Dim i     As Long

        Dim lItem As ListItem
    
        MOVELIST index

        Dim YPARXEI:

        YPARXEI = 0

        If index = 2 Then ' ΠΡΟΣΘΕΤΩ ΤΗΝ ΚΑΤΗΓΟΡΙΑ  (ειμαι στην τελευταία)

            For i = 0 To List2.ListCount - 1

                If Val(Left(List2.List(i), 7)) = List(index).ItemData(List(index).ListIndex) Then
                    YPARXEI = 1
                End If

            Next

            If YPARXEI = 0 Then
                List2.AddItem " " + str(List(index).ItemData(List(index).ListIndex)) + "-" + MARKA.List(MARKA.ListIndex) + "/" + List(1).List(List(1).ListIndex) + "/" + List(index).List(List(index).ListIndex)
            End If
       
            Exit Sub

        End If
    
        'αν δεν υποδιαιρείται η προτελευταία να μπορώ να την προσθέσω
        If index = 1 And List(index + 1).ListCount = 0 Then   ' ΠΡΟΣΘΕΤΩ ΤΗΝ ΚΑΤΗΓΟΡΙΑ (ειμαι στην προ τελευταία)

            For i = 0 To List2.ListCount - 1

                If Val(Left(List2.List(i), 7)) = List(index).ItemData(List(index).ListIndex) Then
                    YPARXEI = 1
                End If

            Next

            If YPARXEI = 0 Then
                List2.AddItem " " + str(List(index).ItemData(List(index).ListIndex)) + "-" + MARKA.List(MARKA.ListIndex) + "/" + List(1).List(List(1).ListIndex) + "/" + List(index).List(List(index).ListIndex)
            End If
       
            Exit Sub

        End If
    
        If List(index + 1).ListCount > 0 Then
            List(index + 1).SetFocus
            List(index + 1).ListIndex = 0
        Else
            R.Open "select * from EID where KOD in (select KOD from PRODUCT_TO_CATEGORIES where CATEGORIES_ID=" + str(List(index).ItemData(List(index).ListIndex)) + ")", Gdb, adOpenDynamic, adLockOptimistic
            'PRODUCT_TO_CATEGORIES ( KOD VARCHAR(16) NOT NULL,PRODUCT_ID int NOT NULL,CATEGORIES_ID INT NOT NULL )
        
            Exit Sub

            Listview1.ListItems.Clear

            Do While Not R.EOF
                Set lItem = Listview1.ListItems.ADD(, , R!kod)  ' Rnd(100))
                lItem.SubItems(1) = R!ono
                lItem.SubItems(2) = str(R!lti5)
                i = i + 1
            Loop
    
        End If

        Exit Sub

End Sub

Sub MOVELIST(index As Integer)

        Dim i As Long

        Dim R As New ADODB.Recordset

        On Error GoTo MOVELIST

102     If index <= 3 And List(index).ListCount > 0 Then
104         R.Open "SELECT * FROM CATEGTEL where AA=" + str(index + 2) + " AND PARENT=" + str(List(index).ItemData(List(index).ListIndex)) + " ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic
            'Dim roulis(5000, 4)
            'Dim I As Long
106         i = 0

108         If index + 1 <= 2 Then List(index + 1).Clear
110         If (index + 2) <= 2 Then List(index + 2).Clear
112         'If (Index + 3) <= 3 Then List(Index + 3).Clear
 
114         Do While Not R.EOF
116             List(index + 1).AddItem R!Name, i
118             List(index + 1).ItemData(i) = R!ID
120             R.MoveNext
            Loop

122         R.Close
   
124         If List(index).ListCount > 0 Then
126             ' list(INDEX + 1).SetFocus
128             ' list(INDEX + 1).ListIndex = 0
            End If

        End If

        Exit Sub '

MOVELIST:
130     MsgBox "MOVELIST" + str(Erl)

        Resume Next

End Sub

Private Sub List1_DblClick()

    Dim a

    a = Left(List1.List(List1.ListIndex), 20)
   
    kod.Text = a
    EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + Trim(a) + "'"
    EID.Refresh
   
    '   List2.Clear
    '   List3.Clear
    '   KODERG.Text = ""
   
    List1.Visible = False
    LIST1_Label.Visible = False
   
    If Not EID.Recordset.EOF Then
        UPDATE_CONTROLS
    Else

        Exit Sub

    End If

End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then List1_DblClick

End Sub

Private Sub List2_Click()

    Dim a

    a = 1
End Sub

Private Sub List2_DblClick()

    Dim a

    a = 1
    List2.RemoveItem (List2.ListIndex)

End Sub

Private Sub List2_DragDrop(Source As Control, x As Single, Y As Single)

    Dim a

    a = 1
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

    Dim a

    a = 1
End Sub

Private Sub LTI5_Change()
    lti5.Text = Replace$(lti5.Text, ",", ".")
End Sub

Private Sub lti5_KeyUp(KeyCode As Integer, Shift As Integer)

100     If KeyCode = vbKeyReturn Then
110         KeyCode = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

End Sub

Private Sub lti5_LostFocus()
    lti5.Text = Replace$(lti5.Text, ",", ".")
    'If Val(WEBPRICE.Text) = 0 Then
    WEBPRICE.Text = lti5.Text
    'End If

End Sub

Private Sub cmdmovenext_Click()

    If Not EGGTIM2.Recordset.EOF Then
        EGGTIM2.Recordset.MoveNext

        If Not EGGTIM2.Recordset.EOF Then
            CMDmovenext.BackColor = vbGreen
            CMDmoveprevious.BackColor = vbGreen
            kod.Text = EGGTIM2.Recordset("KODE")
            kod.SetFocus
            kod_LostFocus
            UPDATE_CONTROLS
            'FONO.SetFocus
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
            kod.Text = EGGTIM2.Recordset("KODE")
            kod.SetFocus
            kod_LostFocus
            UPDATE_CONTROLS
            'FONO.SetFocus
        Else
            CMDmoveprevious.BackColor = vbYellow  ' vbgree
        End If
    
    End If

End Sub

Private Sub MARKA_Click()

    Dim R     As New ADODB.Recordset

    Dim i     As Long

    Dim lItem As ListItem
 
    If MARKA.ListCount > 0 Then
        R.Open "SELECT * FROM CATEGTEL where AA=2 AND PARENT=" + str(MARKA.ItemData(MARKA.ListIndex)) + " ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic
        i = 0
        List(1).Clear
        List(2).Clear
 
        Do While Not R.EOF
            List(1).AddItem R!Name, i
            List(1).ItemData(i) = R!ID
            R.MoveNext
        Loop

        R.Close
    End If

    On Error Resume Next
  
    If List(1).ListCount > 0 Then
        List(1).SetFocus
        List(1).ListIndex = 0
    End If

End Sub

Private Sub MARKA_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 13 Then

        Exit Sub

    End If

    MARKA_Click
End Sub

Private Sub neo_Click()
    NEO_EIDOS
End Sub

Sub NEO_EIDOS()

    On Error Resume Next

    NEA.Visible = True

    Dim NEOEIDOS As String

    NEOEIDOS = InputBox("ΔΩΣΕ ΤΟΝ ΚΩΔΙΚΟ ΤΟΥ ΝΕΟΥ ΕΙΔΟΥΣ", , kod.Text)
    kod.Text = NEOEIDOS

    List2.Clear
    List3.Clear

    'ΝΑ ΜΠΟΡΕΙ ΝΑ ΒΑΛΕΙ ΑΡΧΙΚΟ ΥΠΟΛΟΙΠΟ
    'Text9.Enabled = True

    a = "insert into EID (KOD,WEIGHT) VALUES ('" + kod.Text + "',1)"
 
    ' neo.Enabled = True
    ' Exit Sub
    Dim R As New ADODB.Recordset

    '"   στα νέα είδη αυτόματα να βάζει default βάρος=1κιλό & status=διαθέσιμο

    R.Open "SELECT * FROM EID WHERE KOD='" + kod.Text + "'", Gdb, adOpenDynamic, adLockOptimistic

    If Not R.EOF() Then
        MsgBox "Υπάρχει ο κωδικός με το όνομα " + "" & R!ono
        neo.Enabled = True
        kod.SetFocus
    Else
        Gdb.Execute a
        Gdb.Execute "INSERT INTO BARCODES (KOD,ERG) VALUES ('" + kod.Text + "','" + kod.Text + "')"
        'WEIGHT.Text = "1"
        Check1.Value = vbChecked

    End If

    EID.RecordSource = "SELECT * FROM EID WHERE KOD='" + kod.Text + "'"
    EID.Refresh
   
    KODERG.Text = kod.Text ' ""
    List3.AddItem kod.Text
   
    ONO2.Text = ""
    XTI.Text = ""
    lti5.Text = ""
    WEBPRICE.Text = ""
    WEIGHT.Text = "1"

    Dim k As Integer

    For k = 0 To 6
        foto2(k).Picture = LoadPicture()
    Next
   
    F_id_LOCAL = EID.Recordset("ID")
   
    neo.Enabled = False
    kod.Enabled = False
    ono.Enabled = False
    upload1.Enabled = False

End Sub

Private Sub ono_LostFocus()
    ono.Text = Trim(ono.Text)

    If Len(ono.Text) = 0 Then

        Exit Sub

    End If

    Dim zz

    Dim lti5, POS

    Dim R As New ADODB.Recordset

    On Error GoTo OUT 'Resume Next

    R.Open "select TOP 1000 * from EID WHERE ONO LIKE '" + SameLetters(ono.Text) + "%'  order by ONO ", Gdb, adOpenDynamic, adLockOptimistic
    List1.Clear
    List1.Height = 3900
    List1.Visible = True
    LIST1_Label.Visible = True
  
    Do While Not R.EOF
  
        If IsNull(R!lti5) Then lti5 = 0 Else lti5 = R!lti5
        If IsNull(R!POS) Then POS = 0 Else POS = R!POS
   
        If IsNull(R!ERGOSTASIO) Then zz = Space(20) Else zz = Left(R!ERGOSTASIO + Space(20), 20)
      
        List1.AddItem Left(R("kod") + Space(20), 20) + Left(R("ONO") + Space(50), 50) + zz + Right(Space(8) + Format(lti5, "####0.00"), 8) + "€" + " " + Right(Space(8) + Format(POS, "####0"), 8)

        '  List1.AddItem Left(R("KOD") + Space(20), 20) + R("ONO")
        R.MoveNext
        ' Exit Do
    Loop

    R.Close

    Exit Sub
 
OUT:
  
End Sub

Private Sub T1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim Y

    Dim nam

    Dim c    As String

    Dim sql  As String

    Dim k    As Long  '240 agglia 131 isp 64    '

    Dim m_ID As String

    Dim R    As New ADODB.Recordset

    If T1.Nodes.Count > 0 And KeyCode = 13 Then 'prevents an error if TV is empty
        Y = T1.SelectedItem.index
        'Label1.Caption = Str(y)
        nam = T1.SelectedItem.Text
        Me.Caption = nam & " - " & str(Y) + " names=" + names(Y)
        Set T1.DropHighlight = T1.SelectedItem

        '   m_categ = mID(names(Y), 2, 5)
        List2.AddItem names(Y)
    End If

    If T1.Nodes.Count > 0 And KeyCode = 46 Then 'DELETE SBHNΩ ΚΑΤΗΓΟΡΙΑ
        Y = T1.SelectedItem.index
        'Label1.Caption = Str(y)
        nam = T1.SelectedItem.Text
        Me.Caption = nam & " - " & str(Y) + " names=" + names(Y) + str(FC_id(Y))

        ' FC_id (Y)

        'ψαχνω να δω αν υπάρχουν κατηγοριες με πατερα αυτο που θα σβήσω
        R.Open "SELECT count(*) from " + f_vasi + "_virtuemart_category_categories where category_parent_id=" + str(FC_id(Y)), gdblag, adOpenDynamic, adLockOptimistic

        If R(0) > 0 Then
            MsgBox "αδυνατη η διαγραφή γιατί υποδιαιρείται η κατηγορία"

            Exit Sub

        End If

        R.Close
 
        'ψαχνω να δω αν υπάρχουν ειδη με κατηγορια αυτο που θα σβήσω
        R.Open "SELECT count(*) from " + f_vasi + "_virtuemart_product_categories where virtuemart_category_id=" + str(FC_id(Y)), gdblag, adOpenDynamic, adLockOptimistic

        If R(0) > 0 Then
            MsgBox "αδυνατη η διαγραφή γιατί υπάρχουν είδη με αυτή την κατηγορία"

            Exit Sub

        End If

        R.Close
 
        sql = "DELETE FROM  " + f_vasi + "_virtuemart_category_categories  where id=" + str(FC_id(Y))
        gdblag.Execute sql, k

        '    sql = "DELETE FROM  " + f_vasi + "_virtuemart_categories_el_gr  where virtuemart_category_id=" + str(FC_id(Y))
        '   gdblag.Execute sql, k

        sql = "DELETE FROM  " + f_vasi + "_virtuemart_categories_el_gr  where virtuemart_category_id=" + str(FC_id(Y))
        gdblag.Execute sql, k

        categories

        'MsgBox "ok"
        'T1.Refresh
        ' επιλεγω το καινουριο
        For k = 1 To UBound(names) - 1

            If Val(mID(names(k), 2, 5)) = Y Then
                T1.Nodes.Item(Y).Selected = True
                T1.Nodes.Item(Y).BackColor = vbCyan
                T1.SetFocus
                T1.Nodes.Item(Y).Expanded = True
                T1.Refresh
           
                Exit For

            End If
        
        Next

    End If

    If T1.Nodes.Count > 0 And KeyCode = 45 Then 'insert ΠΡΟΣΘΕΤΩ ΚΑΤΗΓΟΡΙΑ
        Y = T1.SelectedItem.index
        'Label1.Caption = Str(y)
        nam = T1.SelectedItem.Text
        Me.Caption = nam & " - " & str(Y) + " names=" + names(Y)
        Set T1.DropHighlight = T1.SelectedItem

        '   m_categ = mID(names(Y), 2, 5)
        List2.AddItem names(Y)
   
        '       FC_id(N) = r!ID
        '    FC_parenT(N) = r!Parent
        On Error GoTo T1_KeyUp_err
   
        If fc_aa(Y) = 1 Then
            ' T1.Nodes.Add , , "r" + LTrim(str(r!ID)), r!Name
            MsgBox "ΑΔΥΝΑΤΗ ΕΙΣΑΓΩΓΗ ΣΕ ΑΥΤΟ ΤΟ ΕΠΙΠΕΔΟ 1"
        Else
            'T1.Nodes.Clear
            c = InputBox("Δώσε τον τίτλο της κατηγορίας", "Εισαγωγή τύπου")
            sql = "INSERT INTO " + f_vasi + "_virtuemart_category_categories (category_parent_id,ordering)"
            sql = sql + " VALUES (" + str(FC_id(Y)) + ",0);"
            gdblag.Execute sql, k

            If k > 0 Then
                      
                R.Open "SELECT MAX(id) from " + f_vasi + "_virtuemart_category_categories", gdblag, adOpenDynamic, adLockOptimistic
                m_ID = str(R(0))
                      
                sql = "INSERT INTO " + f_vasi + "_virtuemart_categories_el_gr (virtuemart_category_id,category_name,category_description,metadesc,metakey,customtitle,slug) VALUES"
                sql = sql + "( " + m_ID + ",'" + c + "','','','','','" + c + "');"
                gdblag.Execute sql
                      
                sql = "UPDATE " + f_vasi + "_virtuemart_category_categories set category_child_id=" + m_ID + " where id=" + m_ID
                gdblag.Execute sql, k
            End If
                    
        End If

        On Error GoTo 0

        categories

        'MsgBox "ok"
        'T1.Refresh
        ' επιλεγω το καινουριο
        For k = 1 To UBound(names) - 1

            If Val(mID(names(k), 2, 5)) = Y Then
                T1.Nodes.Item(Y).Selected = True
                T1.Nodes.Item(Y).BackColor = vbCyan
                T1.SetFocus
                T1.Nodes.Item(Y).Expanded = True
                T1.Refresh
           
                Exit For

            End If
        
        Next

    End If

    Exit Sub

T1_KeyUp_err:

    'err.Description
    '[MySQL][ODBC 5.1 Driver][mysqld-5.5.34-MariaDB]MySQL server has gone away
    If Err.Number = -2147467259 Then
        gdblag.Close
        gdblag.Open f_connect

        Resume Next

    End If

    Resume Next

End Sub

Private Sub T1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim Y

    Dim nam

    If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty
        Y = T1.SelectedItem.index
        'Label1.Caption = Str(y)
        nam = T1.SelectedItem.Text
        Me.Caption = nam & " - " & str(Y) + " names=" + names(Y)
        Set T1.DropHighlight = T1.SelectedItem

        '  m_categ = mID(names(Y), 2, 4)
        '  List2.AddItem names(Y)

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

    'If Len(Text2.Text) = 0 Then
    '   Exit Sub
    'End If
    '
    'If Left(Text2.Text, 1) = "/" Then
    '   Text2.Text = Trim(mID(Text2.Text, 2, 30))
    '   R.Open "select TOP 100 * from EID WHERE ERG LIKE '" + Text2.Text + "%'", Gdb, adOpenDynamic, adLockOptimistic
    '   List1.Clear
    '   Do While Not R.EOF
    '      List1.AddItem Left(R("KOD") + Space(15), 15) + R("ONO")
    '      R.MoveNext
    '     ' Exit Do
    '   Loop
    '   R.Close
    '   Exit Sub
    'Else
    '    R.Open "select * from EID WHERE KOD IN ( SELECT TOP 1 KOD FROM BARCODES WHERE ERG='" + Text2.Text + "')", Gdb, adOpenDynamic, adLockOptimistic
    'End If
    '
    '
    'If Not R.EOF Then
    '   KOD.Text = R(0)
    '   KOD.SetFocus
    '   FONO.SetFocus
    '   'UPDATE_CONTROLS
    'Else
    '   Exit Sub
    'End If
    '
    '   F_id_LOCAL = R("ID")
    If Len(Text2.Text) = 0 Then

        Exit Sub

    End If

    R.Open "SELECT DISTINCT E.KOD, E.ONO,E.LTI5,E.POS,E.ERGOSTASIO FROM GNISIOI_KODIKOI G  JOIN EID E ON G.KOD=E.KOD WHERE GNISIOS LIKE '%" + Text2.Text + "%' ORDER BY ONO", Gdb, adOpenDynamic, adLockOptimistic
    List1.Clear
    List1.Height = 3900
    List1.Visible = True
    LIST1_Label.Visible = True

    Dim lti5, POS

    Dim zz

    ' List1.Width = 1500
    Do While Not R.EOF
   
        If IsNull(R!lti5) Then lti5 = 0 Else lti5 = R!lti5
        If IsNull(R!POS) Then POS = 0 Else POS = R!POS
        If IsNull(R!ERGOSTASIO) Then zz = Space(20) Else zz = Left(R!ERGOSTASIO + Space(20), 20)
      
        List1.AddItem Left(R("kod") + Space(20), 20) + Left(R("ONO") + Space(50), 50) + zz + Right(Space(8) + Format(lti5, "####0.00"), 8) + "€" + " " + Right(Space(8) + Format(POS, "####0"), 8)
        R.MoveNext
    Loop

    List1.SetFocus

End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        List3.AddItem Text5.Text, 0
    End If

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
                kod.Text = EGGTIM2.Recordset("KODE")
                kod.SetFocus
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
        Y = T1.SelectedItem.index
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
    R.Open "select * from products where products_model='" + kod + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

    M_DATE = Format(Now, "YYYY-MM-DD")

    If R.EOF Then
        m_new = 1
        sql = "INSERT INTO products (products_youtube,products_quantity,products_model,products_image,products_subimage1,"
        sql = sql + " products_subimage2,products_subimage3,products_subimage4,products_subimage5,products_subimage6,products_price,"
        sql = sql + " products_date_added,products_last_modified,products_date_available,products_weight,products_status,products_tax_class_id,"
        sql = sql + " manufacturers_id,products_ordered) VALUES ('',6,'" + kod + "','', "
        sql = sql + " NULL,NULL,NULL,NULL,NULL,NULL,'','" + M_DATE + "','" + M_DATE + "',NULL,'0.20',1,1,109,0); "

        ' Exit Sub
        gdblag.Execute sql, N
        R.Close
        R.Open "select * from products where products_model='" + kod + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
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

    Dim B As String

    ' m_ID = "30917"

    ' extra_value_id1 χαρακτηρας θεμα
    'b = xaraktiras.List(xaraktiras.ListIndex)

    If xaraktiras.ListIndex = -1 Then
        B = "0"
    Else
        B = str(f_xaraktiras(xaraktiras.ListIndex))
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
        m_erg = kataskeyasths.Text + " "
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

        gdblag.Execute "insert into products_description(language_id,products_id,extra_value_id3,extra_value_id1,products_name,products_description) values (4," + m_ID + "," + kor + "," + B + ",'" + m_erg + "','" + Replace(a, "'", "''") + "')", N
    Else

        R.Open "select * from  products_description  where products_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic

        If R.EOF Then
            gdblag.Execute "insert into products_description(language_id,products_id,extra_value_id3,extra_value_id1,products_name,products_description) values (4," + m_ID + "," + kor + "," + B + ",'" + Replace(FONO.Text, "'", "''") + m_erg + "','" + Replace(a, "'", "''") + "')", N
        Else
            gdblag.Execute "update products_description set  extra_value_id3=" + kor + ",extra_value_id1=" + B + " where products_id=" + m_ID, N
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
    'ADDITIONALPRICE              SHIPPINGPRICE                         ERGOSTASIO      YOU_TUBE  WEBPRICE
    'ADDIT_PRICE.text              INDV_PRICE.text                      ZIP_CODE.text
    R.Open "select * from products_shipping where products_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic

    If R.EOF Then
        gdblag.Execute "INSERT INTO products_shipping (products_id,products_ship_methods_id,products_ship_zip,products_ship_price,products_ship_price_two) VALUES (" + m_ID + ",0,'" + ZIP_CODE.Text + "','" + INDV_PRICE.Text + "','" + ADDIT_PRICE.Text + "');", N
    Else
        gdblag.Execute "update products_shipping set  products_ship_price_two='" + ADDIT_PRICE.Text + "',products_ship_price='" + INDV_PRICE.Text + "',products_ship_zip='" + ZIP_CODE.Text + "' where products_id=" + m_ID, N
    End If

    R.Close
    R.Open "select * from products_to_categories where products_id=" + m_ID + " and  categories_id=" + m_categ, gdblag, adOpenDynamic, adLockOptimistic

    If R.EOF Then
        gdblag.Execute "INSERT INTO products_to_categories (products_id,categories_id) VALUES (" + m_ID + "," + m_categ + ");", N
    End If

    R.Close

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

End Sub 'P@ssw0rd#!

Sub ftp_Eikones(strLocal As String, strRemote As String)

    Dim lTimer As Long

    'Dim strRemote As String

    'Dim strLocal  As String
    ' strRemote = "/httpdocs/images"  '             Text4.Text & "/" & Text5.Text
    ' strLocal = "c:\mercvb\images"   ' App.Path & "\Updates\" & Text5.Text

    Dim m2ftp  As New cFTP
    
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
 
    'If T1.Nodes.Count > 0 Then 'prevents an error if TV is empty
    '   Y = T1.SelectedItem.INDEX
    '   nam = T1.SelectedItem.Text
    '   m_categ = mID(names(Y), 2, 4)
    '   Set T1.DropHighlight = T1.SelectedItem
    'Else
    '
    '
    If List2.ListCount = 0 Then
        MsgBox "Δεν επιλέξατε κατηγορία "

        Exit Sub

    End If

    'id=7  category=9
    'INSERT INTO `"+f_vasi+"_virtuemart_products` (virtuemart_vendor_id,product_parent_id,product_sku,product_weight,  ' (1,0,'12346',NULL,
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
 
    'INSERT INTO `"+f_vasi+"_virtuemart_product_categories` (virtuemart_product_id,virtuemart_category_id,ordering) VALUES (7,9,0);
    'προσοχη στο πεδιο slug πρεπει να ειναι μοναδικό π.χ. str(product_id)
 
    'INSERT INTO `"+f_vasi+"_virtuemart_products_el_gr` (virtuemart_product_id,product_s_desc,product_desc,product_name,metadesc,metakey,customtitle,slug)
    '                                                     VALUES (9,        'HP No 21',    '',           'DASSDAS','','','',"9");
 
    ' On Error Resume Next
    'EID.Recordset.Move 0  ' gia na ginei to rich1 opos to mseditor21
 
    a = RichToHTML(Rich1, 0&, Len(Rich1.Text))  '(.TextRTF)
    a = Replace(a, "'", "''")

    gdblag.Close

    gdblag.Open f_connect  '"DRIVER={MySQL ODBC 5.1 Driver};SERVER=188.40.85.131 ;PORT =3306;DATABASE=toyshop;UID=toys;Password=doxadramas78"
    R.Open "select * from " + f_vasi + "_virtuemart_products where product_sku='" + kod + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic

    M_DATE = Format(Now, "YYYY-MM-DD")

    If R.EOF Then
        m_new = 1
   
        sql = "INSERT INTO `" + f_vasi + "_virtuemart_products` (virtuemart_vendor_id,product_parent_id,product_sku,product_weight,product_weight_uom,product_length,product_width,product_height,product_lwh_uom,product_url,product_in_stock,product_ordered,low_stock_notification,product_available_date,product_availability,product_special,product_sales,product_unit,product_packaging,product_params,hits,intnotes,metarobot,metaauthor,layout,published,pordering,created_on,created_by,modified_on,modified_by,locked_on,locked_by) "
        sql = sql + " VALUES (1,0,'" + kod + "',NULL,'TEM',NULL,NULL,NULL,'M','',0,0,0,'2013-11-03 00:00:00','',0,0,'KG',NULL,"
        sql = sql + " '',NULL,'','','','0',1,0,'2013-11-03 16:59:59',42,'2013-11-03 16:59:59',42,'0000-00-00 00:00:00',0);"
        gdblag.Execute sql
        R.Close
        R.Open "select * from " + f_vasi + "_virtuemart_products where product_sku='" + kod + "' LIMIT 1", gdblag, adOpenDynamic, adLockOptimistic
    End If

    m_ID = str(R("virtuemart_product_id"))
    R.Close

    Dim B As String
 
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

    Dim k0 As Integer

    'm_new = 1

    ' μέρη αυτοκινήτου
    sql = "INSERT INTO " + f_vasi + "_virtuemart_product_customfields (virtuemart_product_id,virtuemart_custom_id,custom_value,custom_price,custom_param,published,created_on,created_by,modified_on,modified_by,locked_on,locked_by,ordering)"
    sql = sql + " VALUES (" + m_ID + ",23,'" + EIDOS.Text + "',NULL,'',0,'0000-00-00 00:00:00',0,'2013-12-26 16:47:01',42,'0000-00-00 00:00:00',0,0);"

    Dim SQL2 As String

    'κατηγοριες /
    SQL2 = "INSERT INTO  " + f_vasi + "_virtuemart_product_customfields (virtuemart_product_id,virtuemart_custom_id,custom_value,custom_price,custom_param,published,created_on,created_by,modified_on,modified_by,locked_on,locked_by,ordering) "
    SQL2 = SQL2 + " VALUES (" + m_ID + ",22,'" + KATHGORIA.Text + "',NULL,'',0,'0000-00-00 00:00:00',0,'2013-12-26 16:47:01',42,'0000-00-00 00:00:00',0,1);"

    Dim NK As Long

    If m_new = 1 Then
        ' Exit Sub
        ' κατηγοριες μέρη αυτοκινήτου
   
        gdblag.Execute "INSERT INTO " + f_vasi + "_virtuemart_products_el_gr (virtuemart_product_id,product_s_desc,product_desc,product_name,metadesc,metakey,customtitle,slug) VALUES (" + m_ID + ",'" + FONO.Text + "','" + a + "','" + FONO.Text + "','','','','" + m_ID + "');"
    
        gdblag.Execute "INSERT INTO " + f_vasi + "_virtuemart_product_prices (virtuemart_product_id,virtuemart_shoppergroup_id,product_price,override,product_override_price,product_tax_id,product_discount_id,product_currency,product_price_publish_up,product_price_publish_down,price_quantity_start,price_quantity_end,created_on,created_by,modified_on,modified_by,locked_on,locked_by) VALUES " & "(" + m_ID + ",0,'" + Format(Val(WEBPRICE.Text), "####.00000") + "',0,'0.00000'," + FPA + ",0,47,'2013-11-01 00:00:00','0000-00-00 00:00:00',0,0,'2013-11-03 16:35:41',42,'2013-11-03 16:44:17',42,'0000-00-00 00:00:00',0);"
    
        For k0 = 0 To List2.ListCount - 1
            m_categ = mID(List2.List(k0), 2, 5)  '  mID(names(Y), 2, 5)
            gdblag.Execute "INSERT INTO " + f_vasi + "_virtuemart_product_categories (virtuemart_product_id,virtuemart_category_id,ordering) VALUES (" + m_ID + "," + m_categ + ",0);"
        Next

    Else
  
        '  gdblag.Execute "update " + f_vasi + "_virtuemart_product_customfields set custom_value='" + KATHGORIA.Text + "' where virtuemart_product_id=" + m_ID + " and virtuemart_custom_id=22 limit 1", NK
        '  If NK = 0 Then
        '     gdblag.Execute Sql2
        '  End If
        '
        '  gdblag.Execute "update " + f_vasi + "_virtuemart_product_customfields set custom_value='" + EIDOS.Text + "' where virtuemart_product_id=" + m_ID + " and virtuemart_custom_id=23 limit 1", NK
        '  If NK = 0 Then
        '     gdblag.Execute sql
        '  End If

        gdblag.Execute "update " + f_vasi + "_virtuemart_products_el_gr set product_name='" + FONO.Text + "' ,product_s_desc='" + FONO.Text + "',product_desc='" + a + "'  where virtuemart_product_id=" + m_ID
   
        gdblag.Execute "delete from  " + f_vasi + "_virtuemart_product_categories where virtuemart_product_id=" + m_ID

        For k0 = 0 To List2.ListCount - 1
            m_categ = mID(List2.List(k0), 2, 5) ' mID(names(Y), 2, 5)
            gdblag.Execute "INSERT INTO " + f_vasi + "_virtuemart_product_categories (virtuemart_product_id,virtuemart_category_id,ordering) VALUES (" + m_ID + "," + m_categ + ",0);"
        Next

        ' σβηνω τις παλιές κατηγορίες / μέρη αυτοκινήτου
        gdblag.Execute "delete from " + f_vasi + "_virtuemart_product_customfields  where virtuemart_product_id=" + m_ID + " and (virtuemart_custom_id=23 or virtuemart_custom_id=22) limit 10", NK
    End If

    gdblag.Execute sql
    gdblag.Execute SQL2

    If used.Value = vbChecked Then
    
        gdblag.Execute "DELETE FROM " + f_vasi + "_virtuemart_product_customfields WHERE virtuemart_custom_id=24 AND virtuemart_product_id=" + m_ID + " LIMIT 2"
        '21 ειναι ο κωδικος "μεταχειρισμενο     1=ναι μεταχ   0=οχι καινουριο
        gdblag.Execute "INSERT INTO " + f_vasi + "_virtuemart_product_customfields (virtuemart_product_id,virtuemart_custom_id,custom_value) VALUES (" + m_ID + ",21,'Μεταχειρισμένο');"
    ElseIf Condition Then
        gdblag.Execute "DELETE FROM " + f_vasi + "_virtuemart_product_customfields WHERE virtuemart_custom_id=24 AND virtuemart_product_id=" + m_ID + " LIMIT 2"
        gdblag.Execute "INSERT INTO " + f_vasi + "_virtuemart_product_customfields (virtuemart_product_id,virtuemart_custom_id,custom_value) VALUES (" + m_ID + ",21,'Καινούργιο');"
    End If

    'INSERT INTO `"+f_vasi+"_virtuemart_product_prices`
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

    gdblag.Execute "update " + f_vasi + "_virtuemart_products  set published = " + energo + "   where virtuemart_product_id=" + m_ID

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
    R.Open "SELECT * FROM " + f_vasi + "_virtuemart_product_medias  inner join  " + f_vasi + "_virtuemart_medias  on " + f_vasi + "_virtuemart_medias.virtuemart_media_id=" + f_vasi + "_virtuemart_product_medias.virtuemart_media_id where virtuemart_product_id=" + m_ID, gdblag, adOpenDynamic, adLockOptimistic

    'gdblag.Execute "delete from  "+f_vasi+"_virtuemart_product_medias where virtuemart_product_id = " + m_ID
    n_fotos = 0

    Do While Not R.EOF
        site_foto(n_fotos) = CNull(R!file_title)
        n_fotos = n_fotos + 1
        R.MoveNext
    Loop
  
    ' ΠΩΣ ΑΠΟΘΗΚΕΥΟΝΤΑΙ ΟΙ ΦΩΤΟΓΡΑΦΙΕΣ
    ' 1. FTP THN FOTOGRAFIA
    ' 2. INSERT INTO `"+f_vasi+"_virtuemart_medias` (ΑΥΤΟΜΑΤΟ virtuemart_media_id      ,virtuemart_vendor_id,file_title,file_description,file_meta,file_mimetype,file_type,file_url,file_url_thumb,file_is_product_image,file_is_downloadable,file_is_forSale,file_params,file_lang,shared,published,created_on,created_by,modified_on,modified_by,locked_on,locked_by) VALUES (2,1,'background1.bmp','','','','product','images/stories/virtuemart/product/background1.bmp','',0,0,0,'','',0,1,'2013-11-05 18:55:35',42,'2013-11-06 05:34:29',42,'0000-00-00 00:00:00',0);
    ' ΠΑΙΡΝΩ ΤΟ ID TOY MEDIA ( virtuemart_media_id) ΚΑΙ ΤΟ ΒΑΖΩ ΣΤΟ 3
    ' 3. INSERT INTO `"+f_vasi+"_virtuemart_product_medias` (id,virtuemart_product_id,virtuemart_media_id,ordering) VALUES (1,7,1,1);
 
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
   
        If f_site = 111 Then
            Print #file, "open www.lagakis.gr"
            Print #file, "lagakis"
            Print #file, "39213921"
        Else
            Print #file, "open www.auto-roulis.gr"
            Print #file, "auto-roulis"
            Print #file, "autoroulis2013"
            Print #file, "cd site"
        End If
   
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
    ' 2. INSERT INTO `"+f_vasi+"_virtuemart_medias` (ΑΥΤΟΜΑΤΟ virtuemart_media_id      ,virtuemart_vendor_id,file_title,file_description,file_meta,file_mimetype,file_type,file_url,file_url_thumb,file_is_product_image,file_is_downloadable,file_is_forSale,file_params,file_lang,shared,published,created_on,created_by,modified_on,modified_by,locked_on,locked_by) VALUES (2,1,'background1.bmp','','','','product','images/stories/virtuemart/product/background1.bmp','',0,0,0,'','',0,1,'2013-11-05 18:55:35',42,'2013-11-06 05:34:29',42,'0000-00-00 00:00:00',0);
    ' ΠΑΙΡΝΩ ΤΟ ID TOY MEDIA ( virtuemart_media_id) ΚΑΙ ΤΟ ΒΑΖΩ ΣΤΟ 3
    ' 3. INSERT INTO `"+f_vasi+"_virtuemart_product_medias` (id,virtuemart_product_id,virtuemart_media_id,ordering) VALUES (1,7,1,1);
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
        gdblag.Execute "delete from  " + f_vasi + "_virtuemart_product_medias where virtuemart_product_id = " + m_ID
    End If
 
    For k = 0 To 6

        If Len(PIC(k).Caption) > 0 Then
            gdblag.Execute "INSERT INTO " + f_vasi + "_virtuemart_medias (virtuemart_vendor_id,file_title,file_description,file_meta,file_mimetype,file_type,file_url,file_url_thumb,file_is_product_image,file_is_downloadable,file_is_forSale,file_params,file_lang,shared,published,created_on,created_by,modified_on,modified_by,locked_on,locked_by) VALUES (1,'" + PIC(k).Caption + "','','','image/jpeg','product','images/stories/virtuemart/product/" + PIC(k).Caption + "','',0,0,0,'','',0,1,'" + Format(Now, "YYYY-MM-DD") + "',42,'" + Format(Now, "YYYY-MM-DD") + "',42,'0000-00-00 00:00:00',0);          "
            nc = nc + 1
            
            gdblag.Execute "INSERT INTO " + f_vasi + "_virtuemart_product_medias (virtuemart_product_id,virtuemart_media_id,ordering) VALUES (" + m_ID + ",(SELECT MAX(virtuemart_media_id) FROM " + f_vasi + "_virtuemart_medias )," + str(nc) + ");"
        End If

    Next

    Dim surl As String

    Dim F

    If f_site = 111 Then
        surl = "http://www.lagakis.gr/administrator/index.php?option=com_virtuemart&view=product&task=edit&virtuemart_product_id=" + m_ID + "&product_parent_id=0"
        'BARMAIN.Value = 100
        WebBrowser1.SetFocus
        WebBrowser1.navigate surl

        While Not WebBrowser1.readyState = READYSTATE_COMPLETE

            DoEvents

        Wend

        WebBrowser1.SetFocus
        SendKeys ("admin")
        SendKeys ("{TAB}")
        SendKeys ("lagakis")
        F = WebBrowser1.Document.Forms(0)
    End If

    MsgBox "Ok"

End Sub

Private Sub VScroll1_Change()
    VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
    Picture2.Top = -VScroll1.Value + Picture1.Height + Picture1.Top
End Sub

Sub categories() ' JANASXEDIAZEI TIS KATHGORIES

        '====================================================================
        Dim R As New ADODB.Recordset

        Dim mess

        Dim N

        Dim sql

        If f_site = 222 Then
            sql = "SELECT a.categories_id as ID,a.parent_id as PARENT,b.CATEGORIES_NAME AS NAME from categories a inner join categories_description b on a.categories_id=b.categories_id order by parent_id,a.categories_id"
        Else  'vm
            sql = "SELECT a.virtuemart_category_id as ID,b.category_parent_id as PARENT,a.category_name as NAME FROM " + f_vasi + "_virtuemart_categories_el_gr a inner join " + f_vasi + "_virtuemart_category_categories b on a.virtuemart_category_id=b.category_child_id"
        End If

        f_err = 0
161     Gdb.Execute "DELETE FROM CATEGORIESANT"
162     Gdb.Execute "DELETE FROM CATEGTEL"

164     R.Open sql, gdblag, adOpenDynamic, adLockOptimistic

166     Do While Not R.EOF
168         Gdb.Execute "INSERT INTO CATEGORIESANT (ID,PARENT,NAME) VALUES (" + str(R!ID) + "," + str(R!Parent) + ",'" + Replace(Left(R!Name, 50), "'", "''") + "')"
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

        '---------------------- γεμισμα του treeview control ----------------------------
        'T1.Visible = False
        '
        'T1.Nodes.Clear
        'T1.Refresh
        'T1.Visible = True
        'DoEvents
        '
        '
        'On Error GoTo 0
        '190 R.Open "SELECT * FROM CATEGTEL ORDER BY AA", Gdb, adOpenDynamic, adLockOptimistic
        '' 192 r.MoveFirst
        '
        '    'r.Open "SELECT * FROM CATEG1", Gdb
        '194 Do While Not R.EOF
        '
        '196    'error = 0
        '       'paterades
        '198    MESS = "id=" + LTrim(str(R!ID)) + "Parent:" + LTrim(str(R!Parent)) + " " + R!Name
        '200    If R!Parent = 0 Then
        '202       T1.Nodes.add , , "r" + LTrim(str(R!ID)), R!Name
        '       Else
        '
        '204     T1.Nodes.add "r" + LTrim(str(R!Parent)), tvwChild, "r" + LTrim(str(R!ID)), (R!Name)
        '       End If
        '
        '206    If error = 0 Then
        '208         n = n + 1
        '210         names(n) = "r" + Format(R!ID, "00000") + "        " + R!Name
        '            FC_id(n) = R!ID
        '            FC_parenT(n) = R!Parent
        '            fc_names(n) = R!Name
        '            fc_aa(nn) = R!AA
        '       End If
        '212    R.MoveNext
        '    Loop
        '
        '
        '214  R.Close
        'fc_max = n
        'T1.Refresh

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
'ERGOSTASIO
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

Private Sub UPD_EID(ByVal FIELD As String, ByVal TIMH As String, ByVal TYPOS As String)

    On Error GoTo mess

    Dim L

    If UCase(TYPOS) = "C" Then
        TIMH = Replace(TIMH, "'", "''")
        TIMH = "'" + TIMH + "'"
    End If
  
    If UCase(TYPOS) = "N" Then
        TIMH = Replace(TIMH, ",", ".")
    End If
  
    Gdb.Execute "UPDATE EID  SET " + FIELD + "=" + TIMH + " WHERE KOD='" + kod.Text + "'", L

    If L = 0 Then
        MsgBox "ΠΡΟΒΛΗΜΑ ENHMEΡΩΣΗΣ ΣΤΟ ΠΕΔΙΟ " + FIELD
    End If

    Exit Sub
  
mess:
    MsgBox "ΠΡΟΒΛΗΜΑ ΣΥΝΤΑΞΗΣ ΣΤΟ ΠΕΔΙΟ " + FIELD
End Sub
