VERSION 5.00
Object = "{CCA214C0-DFEB-4C91-9F0D-2665F77F6E23}#1.2#0"; "IDAutomationLinear.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form APOT2 
   BackColor       =   &H00C0C000&
   Caption         =   "Διόρθωση Είδους"
   ClientHeight    =   12384
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   17820
   Icon            =   "APOT2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12384
   ScaleWidth      =   17820
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdWEB 
      Caption         =   "WEB"
      Height          =   360
      Left            =   12360
      TabIndex        =   151
      Top             =   9840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   0
      TabIndex        =   137
      Top             =   2880
      Width           =   6975
      Begin VB.TextBox pos_kerd2 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1920
         TabIndex        =   144
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox lti 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2880
         TabIndex        =   143
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox XondrmeFPA 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   4200
         TabIndex        =   142
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox lti5 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   3000
         TabIndex        =   141
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox pos_kerd 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1920
         TabIndex        =   140
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox xti0 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         HideSelection   =   0   'False
         Left            =   960
         TabIndex        =   139
         Top             =   480
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid times 
         Height          =   1215
         Left            =   0
         TabIndex        =   138
         Top             =   0
         Width           =   6615
         _ExtentX        =   11663
         _ExtentY        =   2138
         _Version        =   393216
         Rows            =   3
         Cols            =   7
         RowHeightMin    =   400
         BackColor       =   14737632
         BackColorFixed  =   14737632
         ScrollBars      =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   0
      TabIndex        =   119
      Top             =   0
      Width           =   6972
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Πραγματική Απογραφή"
         Height          =   255
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CheckBox chkEnergos 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ενεργό"
         Height          =   255
         Left            =   5115
         TabIndex        =   153
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Height          =   288
         Index           =   16
         Left            =   1680
         TabIndex        =   147
         Top             =   2640
         Width           =   1815
      End
      Begin VB.ComboBox Combo5 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   146
         Top             =   2280
         Width           =   1800
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   288
         Index           =   4
         Left            =   1680
         TabIndex        =   145
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   128
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   127
         Top             =   720
         Width           =   4890
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   126
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   125
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         TabIndex        =   124
         Top             =   1920
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   1680
         TabIndex        =   123
         Top             =   1155
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   3720
         TabIndex        =   122
         ToolTipText     =   "Αλλαγή κωδικού"
         Top             =   240
         Width           =   135
      End
      Begin VB.ComboBox Combo4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         TabIndex        =   121
         Top             =   1560
         Width           =   1785
      End
      Begin VB.TextBox lastupd 
         Height          =   285
         Left            =   4800
         TabIndex        =   120
         Top             =   1200
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Υπόλοιπο"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   150
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κατηγορία"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Index           =   26
         Left            =   3240
         TabIndex        =   149
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδ.Προμηθ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   148
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδικός"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   136
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Περιγραφή"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   135
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Μονάδα μέτρησης"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   134
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Κατηγ.ΦΠΑ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   133
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Υποοικογένεια"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   180
         Index           =   4
         Left            =   3240
         TabIndex        =   132
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Βοηθ.Κλειδί(BARCODE)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   131
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Οικογένεια"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   25
         Left            =   3480
         TabIndex        =   130
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Τελ.Αγορά"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   180
         Index           =   5
         Left            =   3600
         TabIndex        =   129
         Top             =   1200
         Width           =   1695
      End
   End
   Begin TrueOleDBGrid80.TDBGrid DBGrid1 
      Bindings        =   "APOT2.frx":000C
      Height          =   2625
      Left            =   7080
      TabIndex        =   101
      Top             =   2880
      Width           =   7635
      _ExtentX        =   13462
      _ExtentY        =   4636
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
      Splits(0).RecordSelectorWidth=   974
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2731"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2731"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=7.8,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=7.8,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
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
   Begin ADOMERCNEW.msEditor2 msEditor1 
      Height          =   2415
      Left            =   7080
      TabIndex        =   90
      Top             =   6240
      Width           =   3855
      _ExtentX        =   6795
      _ExtentY        =   4255
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      Height          =   1335
      Left            =   15000
      ScaleHeight     =   1284
      ScaleWidth      =   6924
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   10320
      Width           =   6972
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H80000010&
      Caption         =   "Εξοδος"
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Αναζήτηση απο ρυθμιζόμενες"
      Height          =   255
      Left            =   10980
      TabIndex        =   79
      Top             =   9180
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc data22 
      Height          =   330
      Left            =   11160
      Top             =   10800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   572
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
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000D&
      Caption         =   "Υπολογισμός Mηνιαίων "
      Height          =   255
      Left            =   480
      TabIndex        =   78
      Top             =   8760
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Νέα Εγγραφή"
      Height          =   495
      Left            =   9840
      TabIndex        =   76
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      TabIndex        =   73
      Top             =   11040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command15 
      Caption         =   "ΕΤΙΚΕΤΤΑ"
      Height          =   195
      Left            =   6480
      TabIndex        =   69
      Top             =   10920
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Εύρεση αρχείου εικόνας"
      Height          =   300
      Left            =   11040
      TabIndex        =   44
      Top             =   8820
      Width           =   3675
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   10680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   11040
      ScaleHeight     =   2484
      ScaleWidth      =   3624
      TabIndex        =   35
      Top             =   6210
      Width           =   3675
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   450
      Left            =   12960
      ScaleHeight     =   20.4
      ScaleMode       =   2  'Point
      ScaleWidth      =   40.2
      TabIndex        =   34
      Top             =   10920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ΑΝΑΖΗΤΗΣΗ"
      Height          =   300
      Left            =   7080
      TabIndex        =   25
      Top             =   8820
      Width           =   2580
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   9870
      TabIndex        =   24
      Top             =   8820
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Διαγραφή"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5640
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   9240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   572
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
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   4485
      Top             =   9480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   572
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
      Caption         =   "DATA2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   6720
      Top             =   9480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   572
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
      Caption         =   "DATA1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Καρτέλλα"
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Διόρθωση"
      Height          =   495
      Left            =   8565
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4896
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Width           =   7092
      _ExtentX        =   12510
      _ExtentY        =   8615
      _Version        =   393216
      MousePointer    =   99
      Tabs            =   9
      Tab             =   4
      TabHeight       =   529
      BackColor       =   12648384
      TabCaption(0)   =   "0.Προμηθ./ελαχ.στόκ"
      TabPicture(0)   =   "APOT2.frx":0021
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "efk"
      Tab(0).Control(1)=   "litra"
      Tab(0).Control(2)=   "LTI2"
      Tab(0).Control(3)=   "lti3"
      Tab(0).Control(4)=   "PROMPOL"
      Tab(0).Control(5)=   "Text1(18)"
      Tab(0).Control(6)=   "Text1(17)"
      Tab(0).Control(7)=   "exoda"
      Tab(0).Control(8)=   "Text1(15)"
      Tab(0).Control(9)=   "Text1(10)"
      Tab(0).Control(10)=   "Text1(7)"
      Tab(0).Control(11)=   "Text1(8)"
      Tab(0).Control(12)=   "Text1(9)"
      Tab(0).Control(13)=   "Label9"
      Tab(0).Control(14)=   "Label8"
      Tab(0).Control(15)=   "lblLTI3"
      Tab(0).Control(16)=   "lbLTI2"
      Tab(0).Control(17)=   "Label7"
      Tab(0).Control(18)=   "Label1(29)"
      Tab(0).Control(19)=   "Label1(28)"
      Tab(0).Control(20)=   "Label3"
      Tab(0).Control(21)=   "Label1(24)"
      Tab(0).Control(22)=   "Label2"
      Tab(0).Control(23)=   "Label1(18)"
      Tab(0).Control(24)=   "Label1(17)"
      Tab(0).Control(25)=   "Label1(16)"
      Tab(0).Control(26)=   "Label1(15)"
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "1.Γραφική παράσταση"
      TabPicture(1)   =   "APOT2.frx":003D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "graph1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "2.Παράμετροι Λογιστ."
      TabPicture(2)   =   "APOT2.frx":0059
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "discountOption"
      Tab(2).Control(1)=   "apallagesFPA"
      Tab(2).Control(2)=   "ONO2"
      Tab(2).Control(3)=   "timokat"
      Tab(2).Control(4)=   "DataGrid2"
      Tab(2).Control(5)=   "Text1(12)"
      Tab(2).Control(6)=   "Text1(11)"
      Tab(2).Control(7)=   "Label11"
      Tab(2).Control(8)=   "lblΑπαλλαγήΦΠΑ"
      Tab(2).Control(9)=   "Label4"
      Tab(2).Control(10)=   "Label1(20)"
      Tab(2).Control(11)=   "Label1(19)"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "3.Κινήσεις κατα μήνα"
      TabPicture(3)   =   "APOT2.frx":0075
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSFlexGrid1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "4.Υπόλοιπα κατά Αποθ."
      TabPicture(4)   =   "APOT2.frx":0091
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label6"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "MSHFlexGrid1"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "5.BARCODES"
      TabPicture(5)   =   "APOT2.frx":00AD
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Adodc2"
      Tab(5).Control(1)=   "Command11"
      Tab(5).Control(2)=   "Command10"
      Tab(5).Control(3)=   "Text3"
      Tab(5).Control(4)=   "DataGrid1"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "6.Κίνηση περισυνή"
      TabPicture(6)   =   "APOT2.frx":00C9
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "MSFlexGrid2"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "7.Γραφημα περισυνό"
      TabPicture(7)   =   "APOT2.frx":00E5
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "graph2"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Πεδία Χρήστη/CPV"
      TabPicture(8)   =   "APOT2.frx":0101
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lhm1"
      Tab(8).Control(1)=   "lhm2"
      Tab(8).Control(2)=   "lhm3"
      Tab(8).Control(3)=   "LaB1"
      Tab(8).Control(4)=   "Lab2"
      Tab(8).Control(5)=   "LAB3"
      Tab(8).Control(6)=   "LNUM1"
      Tab(8).Control(7)=   "LNUM2"
      Tab(8).Control(8)=   "LNUM3"
      Tab(8).Control(9)=   "Lab4"
      Tab(8).Control(10)=   "Lab5"
      Tab(8).Control(11)=   "Lab6"
      Tab(8).Control(12)=   "Label10"
      Tab(8).Control(13)=   "HM1"
      Tab(8).Control(14)=   "HM2"
      Tab(8).Control(15)=   "HM3"
      Tab(8).Control(16)=   "CH1"
      Tab(8).Control(17)=   "CH2"
      Tab(8).Control(18)=   "CH3"
      Tab(8).Control(19)=   "NUM1"
      Tab(8).Control(20)=   "NUM2"
      Tab(8).Control(21)=   "NUM3"
      Tab(8).Control(22)=   "ch6"
      Tab(8).Control(23)=   "ch5"
      Tab(8).Control(24)=   "ch4"
      Tab(8).Control(25)=   "CPV"
      Tab(8).ControlCount=   26
      Begin VB.ComboBox discountOption 
         Height          =   315
         ItemData        =   "APOT2.frx":011D
         Left            =   -72960
         List            =   "APOT2.frx":012A
         Style           =   2  'Dropdown List
         TabIndex        =   157
         Top             =   2760
         Width           =   1572
      End
      Begin VB.TextBox CPV 
         BackColor       =   &H0080C0FF&
         Height          =   375
         Left            =   -70440
         TabIndex        =   155
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox apallagesFPA 
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72975
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox ONO2 
         Height          =   375
         Left            =   -72945
         MultiLine       =   -1  'True
         TabIndex        =   100
         Top             =   1248
         Width           =   4545
      End
      Begin VB.TextBox efk 
         Height          =   324
         Left            =   -69384
         TabIndex        =   97
         Top             =   3120
         Width           =   636
      End
      Begin VB.TextBox litra 
         Height          =   288
         Left            =   -69384
         TabIndex        =   96
         Top             =   2652
         Width           =   636
      End
      Begin VB.TextBox LTI2 
         Height          =   375
         Left            =   -69360
         TabIndex        =   95
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox lti3 
         Height          =   375
         Left            =   -69360
         TabIndex        =   93
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox PROMPOL 
         Height          =   375
         Left            =   -69360
         TabIndex        =   88
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox ch4 
         Height          =   330
         Left            =   -70440
         TabIndex        =   65
         Text            =   " "
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox ch5 
         Height          =   330
         Left            =   -70440
         TabIndex        =   64
         Top             =   2730
         Width           =   1935
      End
      Begin VB.TextBox ch6 
         Height          =   1050
         Left            =   -70440
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox NUM3 
         Height          =   330
         Left            =   -72825
         TabIndex        =   53
         Top             =   4215
         Width           =   2250
      End
      Begin VB.TextBox NUM2 
         Height          =   330
         Left            =   -72825
         TabIndex        =   52
         Top             =   3855
         Width           =   2250
      End
      Begin VB.TextBox NUM1 
         Height          =   330
         Left            =   -72825
         TabIndex        =   51
         Top             =   3465
         Width           =   2250
      End
      Begin VB.TextBox CH3 
         Height          =   330
         Left            =   -72825
         TabIndex        =   50
         Top             =   3105
         Width           =   2250
      End
      Begin VB.TextBox CH2 
         Height          =   330
         Left            =   -72825
         TabIndex        =   49
         Top             =   2685
         Width           =   2250
      End
      Begin VB.TextBox CH1 
         Height          =   330
         Left            =   -72825
         TabIndex        =   48
         Text            =   " "
         Top             =   2295
         Width           =   2250
      End
      Begin VB.TextBox HM3 
         Height          =   330
         Left            =   -72825
         TabIndex        =   47
         Top             =   1935
         Width           =   2250
      End
      Begin VB.TextBox HM2 
         Height          =   330
         Left            =   -72825
         TabIndex        =   46
         Top             =   1530
         Width           =   2250
      End
      Begin VB.TextBox HM1 
         Height          =   330
         Left            =   -72825
         TabIndex        =   45
         Top             =   1095
         Width           =   2250
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   18
         Left            =   -72720
         TabIndex        =   41
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         DataField       =   " "
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   17
         Left            =   -72720
         TabIndex        =   40
         Top             =   3648
         Width           =   1815
      End
      Begin VB.TextBox exoda 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72720
         TabIndex        =   32
         Top             =   4500
         Width           =   1830
      End
      Begin MSAdodcLib.Adodc timokat 
         Height          =   330
         Left            =   -70200
         Top             =   5640
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   572
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
         Caption         =   "Adodc3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "APOT2.frx":014E
         Height          =   1308
         Left            =   -74508
         TabIndex        =   30
         Top             =   3420
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   2307
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   -70545
         Top             =   3660
         Visible         =   0   'False
         Width           =   2070
         _ExtentX        =   3641
         _ExtentY        =   572
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
         Caption         =   "BARCODES"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command11 
         Caption         =   "ΔΙΑΓΡΑΦΗ BARCODE"
         Height          =   285
         Left            =   -70575
         TabIndex        =   29
         Top             =   3000
         Width           =   2100
      End
      Begin VB.CommandButton Command10 
         Caption         =   "ΝΕΟ BARCODE"
         Height          =   285
         Left            =   -70575
         TabIndex        =   28
         Top             =   2310
         Width           =   2100
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   -70590
         TabIndex        =   27
         Top             =   1725
         Width           =   2070
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "APOT2.frx":0164
         Height          =   2940
         Left            =   -73965
         TabIndex        =   26
         Top             =   1755
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   5186
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         DataField       =   " "
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   15
         Left            =   -72720
         TabIndex        =   22
         Top             =   2910
         Width           =   1815
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Bindings        =   "APOT2.frx":0179
         Height          =   2775
         Left            =   240
         TabIndex        =   18
         Top             =   1500
         Width           =   6015
         _ExtentX        =   10605
         _ExtentY        =   4890
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSChart20Lib.MSChart graph1 
         Height          =   3495
         Left            =   -74880
         OleObjectBlob   =   "APOT2.frx":018E
         TabIndex        =   17
         Top             =   1020
         Width           =   6375
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   -69495
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   -72960
         TabIndex        =   13
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   -72720
         TabIndex        =   8
         Top             =   1500
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   -72720
         TabIndex        =   7
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   -72720
         TabIndex        =   6
         Top             =   2460
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   -72720
         TabIndex        =   5
         Top             =   1980
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   4
         Top             =   1140
         Width           =   5295
         _ExtentX        =   9335
         _ExtentY        =   6160
         _Version        =   393216
         Rows            =   13
         Cols            =   5
         GridLines       =   2
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3450
         Left            =   -74340
         TabIndex        =   37
         Top             =   1035
         Width           =   5115
         _ExtentX        =   9017
         _ExtentY        =   6075
         _Version        =   393216
         Rows            =   13
         Cols            =   5
         GridLines       =   2
      End
      Begin MSChart20Lib.MSChart graph2 
         Height          =   3495
         Left            =   -74970
         OleObjectBlob   =   "APOT2.frx":1F97
         TabIndex        =   38
         Top             =   1050
         Width           =   6375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Δικαίωμα Εκπτωσης"
         Height          =   192
         Left            =   -74640
         TabIndex        =   158
         Top             =   2760
         Width           =   1380
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPV"
         Height          =   195
         Left            =   -70440
         TabIndex        =   156
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label lblΑπαλλαγήΦΠΑ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Απαλλαγή ΦΠΑ"
         Height          =   195
         Left            =   -74640
         TabIndex        =   103
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label Label9 
         Caption         =   "Ε.Φ.Κ."
         Height          =   324
         Left            =   -70476
         TabIndex        =   99
         Top             =   3120
         Width           =   948
      End
      Begin VB.Label Label8 
         Caption         =   "Λίτρα"
         Height          =   324
         Left            =   -70476
         TabIndex        =   98
         Top             =   2652
         Width           =   792
      End
      Begin VB.Label lblLTI3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Χονδρική 3η"
         Height          =   195
         Left            =   -70560
         TabIndex        =   94
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label lbLTI2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Χονδρική 2η"
         Height          =   195
         Left            =   -70680
         TabIndex        =   92
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label7 
         Caption         =   "Προμήθεια Πωλητή"
         Height          =   375
         Left            =   -70800
         TabIndex        =   89
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   1800
         TabIndex        =   75
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Δεσμευμένα"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Περιγραφή 2η είδους"
         Height          =   390
         Left            =   -74640
         TabIndex        =   70
         Top             =   1320
         Width           =   1650
      End
      Begin VB.Label Lab6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70485
         TabIndex        =   68
         Top             =   3195
         Width           =   1980
      End
      Begin VB.Label Lab5 
         Height          =   330
         Left            =   -70500
         TabIndex        =   67
         Top             =   2310
         Width           =   1995
      End
      Begin VB.Label Lab4 
         Height          =   210
         Left            =   -70560
         TabIndex        =   66
         Top             =   1800
         Width           =   1950
      End
      Begin VB.Label LNUM3 
         Height          =   285
         Left            =   -74970
         TabIndex        =   62
         Top             =   4215
         Width           =   2205
      End
      Begin VB.Label LNUM2 
         Height          =   285
         Left            =   -74925
         TabIndex        =   61
         Top             =   3855
         Width           =   2100
      End
      Begin VB.Label LNUM1 
         Height          =   255
         Left            =   -74955
         TabIndex        =   60
         Top             =   3480
         Width           =   2130
      End
      Begin VB.Label LAB3 
         Height          =   300
         Left            =   -74955
         TabIndex        =   59
         Top             =   3120
         Width           =   2100
      End
      Begin VB.Label Lab2 
         Height          =   285
         Left            =   -74940
         TabIndex        =   58
         Top             =   2715
         Width           =   2070
      End
      Begin VB.Label LaB1 
         Height          =   315
         Left            =   -74955
         TabIndex        =   57
         Top             =   2295
         Width           =   2130
      End
      Begin VB.Label lhm3 
         Height          =   285
         Left            =   -74940
         TabIndex        =   56
         Top             =   1950
         Width           =   2130
      End
      Begin VB.Label lhm2 
         Height          =   270
         Left            =   -74940
         TabIndex        =   55
         Top             =   1545
         Width           =   2115
      End
      Begin VB.Label lhm1 
         Height          =   315
         Left            =   -74955
         TabIndex        =   54
         Top             =   1110
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Συσκευασία"
         Height          =   375
         Index           =   29
         Left            =   -74280
         TabIndex        =   43
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Υποσυσκευασία"
         Height          =   375
         Index           =   28
         Left            =   -74280
         TabIndex        =   42
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Εξοδα"
         Height          =   216
         Left            =   -74208
         TabIndex        =   33
         Top             =   4500
         Width           =   1368
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus σε κάρτα"
         Height          =   375
         Index           =   24
         Left            =   -74280
         TabIndex        =   23
         Top             =   2910
         Width           =   1455
      End
      Begin VB.Label Label2 
         Height          =   495
         Left            =   -74880
         TabIndex        =   20
         Top             =   3960
         Width           =   6375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδ.Λογ.Αγορών"
         Height          =   375
         Index           =   20
         Left            =   -71040
         TabIndex        =   16
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδ.Λογ.Πωλήσεων"
         Height          =   372
         Index           =   19
         Left            =   -74640
         TabIndex        =   15
         Top             =   2280
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ράφι-Θέση"
         Height          =   375
         Index           =   18
         Left            =   -74280
         TabIndex        =   12
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Προμηθευτής"
         Height          =   375
         Index           =   17
         Left            =   -74280
         TabIndex        =   11
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Επιθυμητό Υπόλοιπο"
         Height          =   375
         Index           =   16
         Left            =   -74280
         TabIndex        =   10
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ελάχιστο στόκ"
         Height          =   375
         Index           =   15
         Left            =   -74280
         TabIndex        =   9
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Διεύθυνση Επιχείρησης"
         Height          =   285
         Index           =   14
         Left            =   -74760
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
      End
   End
   Begin VB.Frame search 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Αναζήτηση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   7080
      TabIndex        =   82
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox ENERGA 
         Height          =   315
         ItemData        =   "APOT2.frx":3DA0
         Left            =   6480
         List            =   "APOT2.frx":3DAD
         Style           =   2  'Dropdown List
         TabIndex        =   152
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox ono 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   118
         ToolTipText     =   "Περιγραφή $ Κωδικός  Συνδυασμένη αναζήτηση π.χ. LL$456 η περιγραφή αρχίζει από LL και ο κωδικός αρχίζει με 456"
         Top             =   800
         Width           =   1575
      End
      Begin VB.TextBox kod 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   117
         ToolTipText     =   "Aναζήτηση"
         Top             =   380
         Width           =   1575
      End
      Begin VB.CheckBox ypoloipo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "με υπόλοιπο"
         Height          =   375
         Left            =   5280
         TabIndex        =   116
         Top             =   960
         Width           =   2000
      End
      Begin VB.TextBox erg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   115
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox barc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   114
         Top             =   1620
         Width           =   1575
      End
      Begin VB.CommandButton cmdAnazit 
         Caption         =   "αναζήτηση"
         Height          =   375
         Left            =   3720
         TabIndex        =   113
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Εττικέτα"
         Height          =   255
         Left            =   3720
         TabIndex        =   111
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Εττικέτα ράφι"
         Height          =   225
         Left            =   3720
         TabIndex        =   110
         Top             =   1770
         Width           =   1455
      End
      Begin VB.CheckBox Scanner 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Από Scanner"
         Height          =   195
         Left            =   5280
         TabIndex        =   109
         ToolTipText     =   "Aπό αριθμό τιμ. 1234"
         Top             =   1320
         Width           =   2000
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Εττικέτα με περιγραφή"
         Height          =   405
         Left            =   5280
         TabIndex        =   108
         Top             =   1560
         Width           =   1600
      End
      Begin VB.CommandButton Command13 
         Height          =   360
         Left            =   6720
         TabIndex        =   107
         Top             =   1560
         Width           =   120
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Left            =   3840
         Picture         =   "APOT2.frx":3DC8
         ScaleHeight     =   444
         ScaleWidth      =   564
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox XRONIES 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   600
         Width           =   1212
      End
      Begin VB.CheckBox cXRONIES 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Παλιές Χρονιές"
         Height          =   372
         Left            =   5280
         TabIndex        =   104
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Περιγραφή"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   45
         TabIndex        =   87
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Κωδικός"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   45
         TabIndex        =   86
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Βοηθ.Κωδικός(Εργ)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   45
         TabIndex        =   85
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Barcode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   45
         TabIndex        =   84
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Αποθήκη"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   83
         Top             =   2040
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   9240
      Top             =   11880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblLTI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Προμήθεια Πωλητή"
      Height          =   375
      Left            =   4200
      TabIndex        =   91
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label flag 
      Caption         =   "flag"
      Height          =   255
      Left            =   240
      TabIndex        =   77
      Top             =   8760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label proteinX 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4935
      TabIndex        =   72
      Top             =   2265
      Width           =   795
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode4 
      Height          =   900
      Left            =   8880
      TabIndex        =   71
      Top             =   10920
      Visible         =   0   'False
      Width           =   1995
      _cx             =   3519
      _cy             =   1587
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      SymbologyId     =   6
      DataToEncode    =   ""
      Orientation     =   0
      BarHeight       =   1
      NarrowBarWidth  =   0.03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   0
      Code128CharSet  =   1
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   0
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0.2
      TopMarginCM     =   0.2
      SupplementToEncode=   ""
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode3 
      Height          =   900
      Left            =   1440
      TabIndex        =   39
      Top             =   10200
      Visible         =   0   'False
      Width           =   1995
      _cx             =   3519
      _cy             =   1587
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      SymbologyId     =   6
      DataToEncode    =   "123456789012"
      Orientation     =   0
      BarHeight       =   1
      NarrowBarWidth  =   0.03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   0
      Code128CharSet  =   1
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   1
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0.2
      TopMarginCM     =   0.2
      SupplementToEncode=   ""
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode2 
      Height          =   1065
      Left            =   3960
      TabIndex        =   36
      Top             =   10440
      Visible         =   0   'False
      Width           =   1995
      _cx             =   3519
      _cy             =   1879
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      SymbologyId     =   6
      DataToEncode    =   "123456789012"
      Orientation     =   0
      BarHeight       =   1
      NarrowBarWidth  =   0.03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   0
      Code128CharSet  =   1
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   1
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0.2
      TopMarginCM     =   0.2
      SupplementToEncode=   ""
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode1 
      Height          =   1065
      Left            =   0
      TabIndex        =   31
      Top             =   10080
      Visible         =   0   'False
      Width           =   1500
      _cx             =   2646
      _cy             =   1879
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.4
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      SymbologyId     =   13
      DataToEncode    =   "0940"
      Orientation     =   0
      BarHeight       =   1
      NarrowBarWidth  =   0.03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   0
      Code128CharSet  =   2
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   1
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0.2
      TopMarginCM     =   0.2
      SupplementToEncode=   ""
   End
   Begin VB.Label PROTEIN 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3795
      TabIndex        =   21
      Top             =   3075
      Width           =   720
   End
End
Attribute VB_Name = "APOT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim f_tab    'TO TAB CONTROL ΣΕ ΠΟΙΟ  ΣΤΑΜΑΤΑΕΙ

Dim F_ET_EPIX_FONTSIZE As Integer
Dim F_ET_EPIX_X As Long
Dim F_ET_EPIX_Y As Long

            Dim F_ET_PER_X, F_ET_PER_Y, F_ET_BAR_H
  
   Dim F_ET_KOD_X, F_ET_KOD_Y
            Dim F_ET_BAR_X, F_ET_BAR_Y
            Dim F_ET_TIM_X, F_ET_TIM_Y

Dim F_ARXEIO_ETIK   As String

Dim f_wresize, f_Hresize As Long

Dim F_DEK_LIANIKIS As Long

Dim F_EuroPerLiter As Single

Dim fmEN As String

'--  This example creates a backup device called MyNwind_1 that contains the full database backup of the MyNwind database.
'
'-- Create the backup device for the full MyNwind backup.
'
'USE emp
'
'EXEC sp_addumpdevice 'disk', 'e2005',     'c:\mssql7\backup\2005.dat'
'
'
'
'-- Back up the full MyNwind database.
'
'BACKUP DATABASE emp TO e2005
'
'

'Αντιγραφω το αρχείο της τρέχουσας χρονιάς 2004 με άλλο όνομα μέσα
'στον φάκελλο ..\sql7\data και τον κάνω attach με το όνομα της παλιάς χρονίας
' ετσι έχω τα αρχεία της τρέχουσας σε δύο databases
'CREATE DATABASE e2004
'
'ON PRIMARY (FILENAME = 'c:\mssql7\data\e2004.mdf'),
'(FILENAME = 'c:\mssql7\data\e2004.ldf')
'FOR ATTACH
'
'GO
'
''

Dim F_KAT_FPA As Integer

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
Dim f_printer As String

'Private El As New class_Elastic
Dim f_TIMOKAT

Dim F_add_barc As Integer

Dim F_YBARCODE, F_XBARCODE
Attribute F_XBARCODE.VB_VarUserMemId = 1073938435

Dim F_CHANGETIMH
Attribute F_CHANGETIMH.VB_VarUserMemId = 1073938437

Dim F_0_BARCODE As String


Dim F_CHANGE_MARKUP As Integer



Dim F_Epix      As String

Dim F_MIK       As Integer

Dim f_dek_xondr As Integer

Dim f_site      As Integer

Dim f_XOND_MEFPA_DISABLE As Integer ' f_XOND_MEFPA_DISABLE

Dim f_YPOL_MHN  As Integer

Dim FGDB2       As New ADODB.Connection  ' fanaropoylos singular

Dim fConnect2   As String   ' fanaropoylos singular

'Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

'Dim f_Refresh As Boolean

Option Explicit

Dim F_ArPerEtik    ' αριστερο περιθωριο ετικεττας
Attribute F_ArPerEtik.VB_VarUserMemId = 1073938445

Dim F_XPONTOI As Long, F_YPONTOI As Long ' 200

Dim F_2ArPerEtik    '2o αριστερο περιθωριο ετικεττας
Attribute F_2ArPerEtik.VB_VarUserMemId = 1073938446

Dim F_PIChEIGHT    ' Val(FindParametroi(1,"APOT2", "F_PIChEIGHT", "2535", "Υψος εικόνας")) 'posa psifia tha exei h kathe seira
Attribute F_PIChEIGHT.VB_VarUserMemId = 1073938447

Dim F_PICwidth   ' = Val(FindParametroi(1,"APOT2", "F_PICwidth", "2535", "Υψος εικόνας")) 'posa psifia tha exei h kathe seira
Attribute F_PICwidth.VB_VarUserMemId = 1073938448

Dim F_sygx    ' SYGXONEYSH KODIKON=1
Attribute F_sygx.VB_VarUserMemId = 1073938449

Dim F_EIKONA  ' 1=KVDIKOS   2=ERG  3=TIPOTA DEN EXEI EIKONES
Attribute F_EIKONA.VB_VarUserMemId = 1073938450

Dim F_KOD As Integer, F_ONO As Integer, F_ERG As Integer, F_XTI As Integer

Dim f_FontName   '"Arial"
Attribute f_FontName.VB_VarUserMemId = 1073938455

Dim F_LHM1 As String

Dim F_LHM2 As String

Dim f_lHM3 As String

Dim f_lab1 As String

Dim f_lab2 As String

Dim f_lab3 As String

Dim f_lab4 As String

Dim f_lab5 As String

Dim f_lab6 As String

Dim F_EXO_XROMATA
Attribute F_EXO_XROMATA.VB_VarUserMemId = 1073938465

Dim F_UPPER    ' ΜΕΤΑΤΡΕΠΕΙ ΣΕ ΚΕΦΑΛΑΙΑ
Attribute F_UPPER.VB_VarUserMemId = 1073938466

Dim f_num1         As String

Dim f_num2         As String

Dim f_num3         As String

Dim El             As New FORMA

Dim F_FAKEL_KOINOS As String

Dim F_PONTOI
Attribute F_PONTOI.VB_VarUserMemId = 1073938470

Dim PROTH_FORA

'
' Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
 Dim xWidth  As Long
 Dim xHeight As Long
 Dim xColor1 As Long
 Dim xColor2 As Long
 
 
Dim gColor1 As Long
 Dim gColor2 As Long
 Dim gColor3 As Long
 Dim gColor4 As Long
 Dim gDirection    As Direction






Private Type BITMAP '14 bytes

    bmType                      As Long
    bmWidth                     As Long
    bmHeight                    As Long
    bmWidthBytes                As Long
    bmPlanes                    As Integer
    bmBitsPixel                 As Integer
    bmBits                      As Long

End Type



Private Sub barc_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
         KeyAscii = 0    'suppress the beep
         keybd_event VK_TAB, 0, 0, 0    'send a tab
        
         
     End If
End Sub

Private Sub barc_LostFocus()
     

  Dim xti_or_pos As Integer

        xti_or_pos = Val(FINDPARAMETROI(1, "APOT2", "xti_or_pos", "1", "ΣΤΟ GRID 1=ΠΟΣΟΤΗΤΑ 2=ΤΙΜ.ΑΓΟΡΑΣ +10 ΧΟΝΔΡ ΑΝΤΙ ΛΙΑΝΙΚΗ"))

 If F_0_BARCODE = 1 Then
     If Left(barc.Text, 1) = "0" Then
         barc.Text = Trim(mID(barc.Text, 2, 18))
     End If
 End If

 Dim PED2 As String
 Dim PED1 As String
 
 
 
        If xti_or_pos Mod 10 = 2 Then
            PED2 = " CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] "
        Else
            PED2 = " CONVERT(decimal(10,2),POS) AS [ΥΠΟΛΟΙΠΟ] "
        End If
     
 If xti_or_pos > 10 Then
     PED1 = "CONVERT(decimal(10,2),LTI) AS [XONΔΡ.ΤΙΜΗ],"
     
 Else
     PED1 = "CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],"
 End If
 
     
     If Len(barc) > 0 Then

          '  On Error GoTo Command5_Click_Err

         data22.RecordSource = "select  KOD,ONO," + PED1 + PED2 + " from EID where KOD IN (SELECT KOD FROM BARCODES WHERE ERG like '" + barc.Text + "%' ESCAPE '$' )"
         data22.Refresh

         If data22.Recordset.EOF Then
             Text1(0).Text = " "
             Text1(1).Text = " "
            End If

         DBGrid1.SetFocus
         
          cmdAnazit_Click
     End If
End Sub

Private Sub cmdWEB_Click()

    If f_site = 111 Or f_site = 222 Or f_site = 444 Then
       ' apot2vmn.SHOW
    ElseIf f_site = 333 Then
   
      '  apot2ROULIS.SHOW
   
        '  apot2scroll.kod.Text = Text1(0).Text
        ' apot2scroll.kod.SetFocus
        ' apot2scroll.ono.SetFocus
        ' apot2scroll.SHOW
   
    End If

End Sub

'Function Update_Next_Pel(meidos As String) As Integer
'
'    Dim db As Database
'    Dim R As Recordset
'    '12- 9252  141
'
'End Function

'Function get_next_pel(meidos As String) As Integer
'    Dim db As Database
'    Dim R As Recordset
'
'
'    Set db = OpenDatabase(gDir)
'    Set db = OpenDatabase(gDir, False, False)
'    Set db = OpenDatabase(gDir, False, False, gConnect)
'
'
'
'
'    Set R = db.OpenRecordset("mem")
'    If meidos = "e" Then
'        get_next_pel = R("epel") + 1
'    Else
'        get_next_pel = R("rpel") + 1
'    End If
'End Function

Private Sub Combo1_GotFocus()

        '<EhHeader>
        On Error GoTo Combo1_GotFocus_Err

        '</EhHeader>

100     Combo1.BackColor = vbYellow

        '<EhFooter>
        Exit Sub

Combo1_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo1_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo1_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Combo1_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Combo1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo1_LostFocus()

        '<EhHeader>
        On Error GoTo Combo1_LostFocus_Err

        '</EhHeader>

100     Combo1.BackColor = vbWhite

        '<EhFooter>
        Exit Sub

Combo1_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo1_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo1_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo2_GotFocus()

        '<EhHeader>
        On Error GoTo Combo2_GotFocus_Err

        '</EhHeader>

100     Combo2.BackColor = vbYellow

        '<EhFooter>
        Exit Sub

Combo2_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo2_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo2_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Combo2_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Combo2_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo2_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo2_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo2_LostFocus()

        '<EhHeader>
        On Error GoTo Combo2_LostFocus_Err
        
        
        If Val(Left$(Combo2, 2)) = 5 Then
           
              apallagesFPA.Enabled = True
              MsgBox "ΔΗΛΩΣΤΕ ΛΟΓΟ ΑΠΑΛΛΑΓΗΣ"
        
        End If
        
        
        

        '</EhHeader>

100     Combo2.BackColor = vbWhite

        '<EhFooter>
        Exit Sub

Combo2_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo2_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo2_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo3_GotFocus()

        '<EhHeader>
        On Error GoTo Combo3_GotFocus_Err

        '</EhHeader>

100     Combo3.BackColor = vbYellow

        '<EhFooter>
        Exit Sub

Combo3_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo3_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo3_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Combo3_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Combo3_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo3_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo3_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo3_LostFocus()

        '<EhHeader>
        On Error GoTo Combo3_LostFocus_Err

        '</EhHeader>

100     Combo3.BackColor = vbWhite

        '<EhFooter>
        Exit Sub

Combo3_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo3_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo3_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo4_GotFocus()

        '<EhHeader>
        On Error GoTo Combo4_GotFocus_Err

        '</EhHeader>

100     Call SendMessage(Combo4.hwnd, CB_SHOWDROPDOWN, 1, 0)

        '<EhFooter>
        Exit Sub

Combo4_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo4_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo4_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo4_LostFocus()

        ' YPOOIKOGENEIES
        '<EhHeader>
        On Error GoTo Combo4_LostFocus_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

100     If Len(Combo4.Text) > 0 Then

110         BALE_TSONTES "3", Left(Combo4.Text, 2)
120         R.Open "SELECT *FROM PINAKES WHERE TYPOS=3 AND TIMH=" + Left(Combo4.Text, 2) + "ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

            Dim N As Integer

130         N = Combo3.ListIndex
140         Combo3.Clear

150         Do While Not R.EOF

160             If R("typos") = 3 Then
170                 Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
                Else

                    Exit Do

                End If

180             R.MoveNext
            Loop

190         R.Close

            On Error Resume Next

            ' Call SendMessage(Combo3.hWnd, CB_SHOWDROPDOWN, 1, 0)
200         If Combo3.ListCount < N Then

210             Combo3.ListIndex = 0
            Else
220             Combo3.ListIndex = N   ' - 1 Else Combo3.ListIndex = 0
            End If

        End If

        '<EhFooter>
        Exit Sub

Combo4_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Combo4_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Combo4_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub CMDEXIT_Click()

        'exodos
        '<EhHeader>
        On Error GoTo CMDEXIT_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

CMDEXIT_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command1_Click()
On Error Resume Next
  
  Dim CC As String
  CC = Text1(4).Text
  Dim cHME As String
  Dim cday, cmonth, cYear As String
  If Year(Now) = Year(gLhjh) Then
     cHME = "GETDATE()"
  Else
     cHME = "12/31/" + Format(Year(gLhjh), "0000")
  End If
  Dim cc22 As String
  
   cc22 = InputBox("Δωσε Αποθήκη:", , 1)
  Dim cc2 As String
Dim cc3 As String



  Dim R As New ADODB.Recordset
  
  R.Open "SELECT (select ROUND(SUM(isnull(XRE,0)),0) -ROUND(SUM(isnull(PIS,0) ),0) )  AS [ypol] FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' AND APOT=" + cc22, Gdb, adOpenDynamic, adLockOptimistic
 If R.EOF Then
    CC = "0"
  Else
     CC = str(R(0))
     CC = Replace(CC, ",", ".")
  End If
 cc2 = InputBox("δωσε το πραγματικό υπόλοιπο. Υπάρχον υπόλοιπο σε Η/Υ:" + CC, CC)
  
cc3 = str(Val(Replace(cc2, ",", ".")) - Val(CC))
  
  
  Gdb.Execute "INSERT INTO EGGTIM (APOT,EIDOS,ATIM,HME,KODE,POSO,XRE,PIS) VALUES (" + cc22 + ",'r','s000100'," + cHME + ",'" + Text1(0) + "'," + cc3 + "," + cc3 + ",0)"
  
  
  Dim synt
     synt = " HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND HME<'" + Format(gLhjh, "MM/DD/YYYY") + "' "
  R.Close
  R.Open "SELECT KODE,APOT,(select PERIGRAFH FROM PINAKES WHERE TYPOS=4 AND AYJON=APOT) AS [ΑΠΟΘ],ROUND(SUM(XRE),0) AS [ΕΙΣ],ROUND(SUM(PIS),0) AS [ΕΞΑΓ]," & " round( (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" & "  (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ),1) AS YPOL " & " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' AND " + synt + "  GROUP BY KODE,APOT", Gdb, adOpenDynamic, adLockOptimistic
  Text1(4).Text = R!YPOL
  
  
  
  
End Sub

Private Sub Command10_Click()

        '<EhHeader>
        On Error GoTo Command10_Click_Err

        '</EhHeader>

        Dim a

        Dim r0 As New ADODB.Recordset

100     a = SSTab1.Tab

110     If a = 5 And F_EXO_XROMATA = 1 Then
120         UGridEdit.Update.Visible = True

130         UGridEdit.Text1.Visible = True
140         UGridEdit.kod.Caption = Text1(0).Text
150         UGridEdit.ono.Caption = Text1(1).Text
160         UGridEdit.CH1.Caption = CH1.Text
170         UGridEdit.lti5.Caption = lti5.Text
180         UGridEdit.XTI.Caption = xti0.Text

            Dim DUM

190         DUM = ADD_FIELD("BARCODES", "PROELEYSH", "CHAR(15)")

            Dim ColS As Integer, rows As Integer

200         ColS = 0: rows = 0

            'ΦΟΡΤΩΝΩ ΤΙΣ ΕΠΙΚΕΦΑΛΙΔΕΣ ΤΟΥ GRID ΧΡΩΜΑΤΩΝ
            Dim MEGETHOS As String, XROMA As String

210         XROMA = Format(29 + (2 * Val(Left(Combo5.Text, 2)) - 1), "00")
220         MEGETHOS = Format(29 + (2 * Val(Left(Combo5.Text, 2))), "00")

            'ΦΟΡΤΩΝΩ ΤΙΣ ΕΠΙΚΕΦΑΛΙΔΕΣ ΤΟΥ GRID ΧΡΩΜΑΤΩΝ
230         r0.Open "SELECT *FROM PINAKES WHERE TYPOS=" + XROMA + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

240         Do While Not r0.EOF
250             UGridEdit.Grid1.TextMatrix(r0("AYJON"), 0) = r0("PERIGRAFH")
260             rows = rows + 1
270             r0.MoveNext
            Loop

280         r0.Close

            'ΦΟΡΤΩΝΩ ΤΙΣ ΕΠΙΚΕΦΑΛΙΔΕΣ ΤΟΥ GRID ΜΕΓΕΘΩΝ

290         r0.Open "SELECT *FROM PINAKES WHERE TYPOS=" + MEGETHOS + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

300         Do While Not r0.EOF
310             UGridEdit.Grid1.TextMatrix(0, r0("AYJON")) = r0("PERIGRAFH")
320             ColS = ColS + 1
330             r0.MoveNext
            Loop

340         r0.Close

350         UGridEdit.Grid1.rows = rows + 1
360         UGridEdit.Grid1.ColS = ColS + 1

            'ΤΑ ΥΠΟΛΟΙΠΑ ANA XROMA ΑΠΟ ΤΟ MSFLEXGRID1 TA DEIXNO ΣΤΟ  GRID1 TOY UGRIDEDIT
            Dim k As Integer

370         UGridEdit.Label1.Caption = "NoPaint"    ' flag gia na mhn gemisei to grid1 apo to gm_str(i)

            On Error Resume Next

            Dim X  As Long, Y As Long

            Dim n2 As Long

            Dim N3 As Long

            Dim n4

380         For k = 1 To MSHFlexGrid1.rows - 1

390             For n2 = 1 To UGridEdit.Grid1.rows - 1  ' ΨΑΧΝΩ ΣΤΙΣ ΣΕΙΡΕΣ ΝΑ ΒΡΩ ΤΟ ΧΡΩΜΑ

400                 If Left(MSHFlexGrid1.TextMatrix(k, 1), 3) = Left(UGridEdit.Grid1.TextMatrix(n2, 0), 3) Then
410                     X = n2

                        Exit For

                    End If

                Next

                'ΨΑΧΝΩ ΣΤΙΣ ΣΤΗΛΕΣ ΝΑ ΒΡΩ ΤΟ ΜΕΓΕΘΟΣ
420             N3 = InStr(MSHFlexGrid1.TextMatrix(k, 1), "*")
430             n4 = Max(1, N3 - 4)

440             For n2 = 1 To UGridEdit.Grid1.ColS - 1  ' ΨΑΧΝΩ ΣΤΙΣ ΣΕΙΡΕΣ ΝΑ ΒΡΩ ΤΟ ΧΡΩΜΑ

450                 If mID(MSHFlexGrid1.TextMatrix(k, 1), 4, n4) = Left(UGridEdit.Grid1.TextMatrix(0, n2), n4) Then
460                     Y = n2

                        Exit For

                    End If

                Next

470             UGridEdit.Grid1.TextMatrix(X, Y) = Val(UGridEdit.Grid1.TextMatrix(X, Y)) + Val(MSHFlexGrid1.TextMatrix(k, 6))

            Next

480         UGridEdit.SHOW 1

            Exit Sub

        End If

        On Error Resume Next

        Dim D

        'On Error GoTo Command10_Click_Err
490     Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + Text1(0).Text + "','" + Text3.Text + "')", D
500     Adodc2.Refresh

510     If D = 0 Then
520         MsgBox "Δεν καταχωρήθηκε"
        End If

        '<EhFooter>
        Exit Sub

Command10_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command10_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command10_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command11_Click()

        '<EhHeader>
        On Error GoTo Command11_Click_Err

        '</EhHeader>

100     Gdb.Execute "DELETE  FROM BARCODES WHERE ERG='" + Text3.Text + "' AND KOD='" + Text1(0).Text + "'"
110     Adodc2.Refresh

        '<EhFooter>
        Exit Sub

Command11_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command11_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command12_Click()

        ' ETIKETTA ME KEIMENO
        '<EhHeader>
        On Error GoTo Command12_Click_Err

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

100     mONOMA = Text1(1).Text    'ONOMA
110     mBarcode = Text1(0).Text    '  ' BARCODE
120     mERG = Text1(6).Text    ' erg
130     mLTI5 = Format(Val(lti5.Text), "###0." + String(F_DEK_LIANIKIS, "0"))

        Dim Z As Integer

140     Z = InputBox("ΠΟΣΕΣ ΕΤΙΚΕΤΤΕΣ? ", , 1)

150     For k = 1 To Z

160         If F_MIK = 1 Or F_MIK = 11 Then ' Len(Dir("C:\MPOYGZEB.EXE", vbNormal)) > 0 Then
170             GoSub ZEBRAmonoBARCODE
            Else
180             GoSub SATO
            End If

190     Next k

200     Printer.EndDoc

        Exit Sub

SATO:

        On Error Resume Next

210     For Each X In Printers

220         If InStr(X.DeviceName, "ETIK") > 0 Then
                ' Set printer as system default.
230             Set Printer = X
                ' Stop looking for a printer  5200016000239

                Exit For

            End If

        Next

240     apoSCANNER = 1

250     If Scanner Then

            ' μπουγουδης
260         Set DB = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
270         Set RD = DB.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM= 1234")

280         If RD.RecordCount = 0 Then
290             MsgBox "δεν υπάρχουν εγγραφές"

                Exit Sub

            End If

300         RD.MoveFirst
310         k = 0

320         Do While Not RD.EOF
330             k = k + 1

340             If k > 80 Then Exit Do
350             If IsNull(RD("BARCODE")) Then
360                 barc(k) = " "    'Grid1.TextMatrix(K, f_k) = " "
                Else
370                 barc(k) = RD("barcode")    'Grid1.TextMatrix(K, f_k) = R("BARCODE")
                End If

380             RD.MoveNext
            Loop

390         RD.Close
400         DB.Close
410         apoSCANNER = k
        End If

420     fores = Z ' Val(InputBox("Πόσες εττικέτες; ", "", 1))

        Dim mFORES As Long

430     For ll = 1 To apoSCANNER

440         mFORES = fores

450         If Scanner Then
460             R.Close
470             R.Open "select * FROM EID WHERE KOD='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

480             If R.EOF Or Len(Trim(barc(ll))) = 0 Then
490                 mFORES = 0
                Else
500                 mONOMA = R("ONO")    'ONOMA
510                 mBarcode = R("KOD")    '  ' BARCODE
520                 mERG = R("ERG")    ' erg
530                 mLTI5 = Format(R("LTI5"), "###0." + String(F_DEK_LIANIKIS, "0"))   ' LTI5
                End If
            End If

540         Z = -5    ' InputBox("METATOPISH")


Dim TEL As Integer
TEL = InStrRev(Left(mONOMA, 25), " ")



550         For k = 1 To mFORES

560             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992

570             Printer.FontSize = 12
580             Printer.FontBold = True

                Dim a As Integer

590             a = InStr(1, mONOMA, ";")

600             If a = 0 Then
610                 Printer.Print Tab(0); Left(mONOMA, TEL)    'Text1(1).Text 'ONOMA
                    Printer.Print Tab(0); mID(mONOMA, TEL + 1, 20) 'Text1(1).Text 'ONOMA
                Else
620                 Printer.Print Tab(5); Left(mONOMA, a - 1)    'Text1(1).Text 'ONOMA
630                 Printer.Print Tab(5); mID$(mONOMA, a + 1, 17)    'Text1(1).Text 'ONOMA
                End If

640             Printer.EndDoc
650         Next k

660     Next ll

670     Return

ZEBRAmonoBARCODE:

        On Error Resume Next

680     For Each X In Printers

690         If InStr(X.DeviceName, "MIKR") > 0 Then
700             Set Printer = X

                Exit For

            End If

        Next

710     Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName454967325476-5200016000239

720     Printer.FontSize = 8

730     Printer.CurrentY = 0
740     Printer.CurrentX = 0
        'Printer.Print Left(Text1(6).Text, 12)
If F_MIK = 11 Then
       BarCode2.DataToEncode = DataGrid1.Text
Else
       BarCode2.DataToEncode = Text1(0).Text
End If
760     BarCode2.Code128CharSet = Set_A

770     If Len(DataGrid1.Text) = 12 Then
780         BarCode2.SymbologyID = UPCA
        Else
790         BarCode2.SymbologyID = EAN13
        End If

800     Printer.PaintPicture BarCode2.Picture, 200, -400
810     Printer.EndDoc

820     Return

        'Dim fores As Long, K As Integer
        'Dim x As Printer
        'On Error Resume Next
        'For Each x In Printers
        '    If InStr(x.DeviceName, "SATO") > 0 Then
        '        ' Set printer as system default.
        '        Set Printer = x
        '        ' Stop looking for a printer
        '        Exit For
        '    End If
        'Next
        '
        '
        '
        'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text))
        '
        '
        '
        'For K = 1 To fores
        '
        '
        '
        'Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
        'Printer.FontSize = 8
        '
        '
        '
        'Printer.Print Tab(20); Text1(1).Text 'ONOMA
        'Printer.Print Tab(30); Text1(6).Text ' mlabel
        '' Printer.Print
        'If Len(Trim(Text1(0).Text)) = 12 Then
        '   Printer.FontName = "UPCA" ' "128"
        'Else
        '   Printer.FontName = "EAN13SMALL" ' "128"
        'End If

        'Printer.FontSize = 5
        'Printer.Print Tab(9); Text1(0).Text 'eggtim.Recordset("kode") ' mlabel
        'Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
        'Printer.FontSize = 24
        '
        'If Len(mlti5) = 6 Then ' 208.35
        '     Printer.Print Tab(12); " € " + mlti5; ""
        'ElseIf Len(mlti5) = 5 Then '18.35
        '     Printer.Print Tab(13); "€ " + mlti5; ""
        'Else
        '     Printer.Print Tab(15); "€ " + mlti5; ""
        'End If
        '
        'Printer.Print
        'Printer.Print
        ' Printer.Print
        'Printer.EndDoc
        'Next

        '<EhFooter>
        Exit Sub

Command12_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command12_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        
        
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command12_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command13_Click()

        '<EhHeader>
        On Error GoTo Command13_Click_Err

        '</EhHeader>

        Dim X

        Dim rec As ADODB.Recordset

        Dim ANS

        Dim k

        Dim fores

        Dim FF, ARX

        Dim REC2 As New ADODB.Recordset

100     Text4.Visible = True

110     If Len(Dir("C:\FANAR.EXE", vbNormal)) > 0 Then

120         For Each X In Printers

130             If InStr(X.DeviceName, "TLP2844") > 0 Then
                    ' Set printer as system default.

140                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

150         GoSub FIND_N_NUMBER

160         ARX = Text4.Text

170         FF = InputBox(ARX, , ARX)

180         Text4.Text = Right("000000000000" + LTrim(str(FF)), 12)

            '   5212345678935

190         BarCode2.DataToEncode = FF

200         fores = Val(InputBox("Πόσα κομμάτια  "))

210         For k = 1 To fores
220             Printer.CurrentY = 400    '000000020404
230             Printer.CurrentX = 600 + F_ArPerEtik * 100
240             BarCode2.SymbologyID = EAN13
250             Printer.PaintPicture BarCode2.Picture, F_XBARCODE, 100, , 1300
260             Printer.EndDoc
270         Next k

280         If FF <> ARX Then

290             ANS = MsgBox("Να αποθηκευτεί ο αριθμός " + FF, vbYesNo)

300             If ANS = vbYes Then
310                 GoSub update_N_NUMBER
                End If
            End If

320     ElseIf Len(Dir("C:\mercvb\zebra.txt", vbNormal)) > 0 Then

330         For Each X In Printers

340             If InStr(X.DeviceName, "TLP2844") > 0 Then
                    ' Set printer as system default.

350                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

360         GoSub FIND_N_NUMBER

370         ARX = Text4.Text

380         FF = InputBox(ARX, , ARX)

390         Text4.Text = Right("000000000000" + LTrim(str(FF)), 12)

            '   5212345678935

400         BarCode2.DataToEncode = FF

            'For k = 1 To fores
410         Printer.CurrentY = 400    '000000020404
420         Printer.CurrentX = 600 + F_ArPerEtik * 100

430         BarCode2.SymbologyID = EAN13
440         Printer.PaintPicture BarCode2.Picture, F_XBARCODE, F_YBARCODE, , 1300

450         Printer.EndDoc

460         If FF <> ARX Then

470             ANS = MsgBox("Να αποθηκευτεί ο αριθμός " + FF, vbYesNo)

480             If ANS = vbYes Then
490                 GoSub update_N_NUMBER
                End If
            End If

        Else

500         fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

            Dim PARAM

510         For Each X In Printers

520             If InStr(X.DeviceName, "SATO") > 0 Then
                    ' Set printer as system default.

530                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

540         BarCode2.DataToEncode = DataGrid1.Text

550         For k = 1 To fores
560             Printer.CurrentY = 400
570             Printer.CurrentX = 600 + F_ArPerEtik * 100
580             Printer.FontSize = 12
590             Printer.Print Text1(0).Text
600             Printer.EndDoc

            Next

        End If

        Exit Sub

FIND_N_NUMBER:
        '************************

        'rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

610     REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
620     REC2.MoveFirst
630     REC2.Move 10

        'rec.Edit
640     If IsNull(REC2("epel")) Then
650         REC2("epel") = 1
        Else
660         REC2("epel") = REC2("epel") + 1
        End If

670     REC2.Update

680     Text4.Text = Right("000000000000" + LTrim(str(REC2("EPEL"))), 12)

690     Return

update_N_NUMBER:
        '************************

700     REC2.Close
710     REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
720     REC2.MoveFirst
730     REC2.Move 10
740     REC2("epel") = Val(FF)
750     REC2.Update

760     Text4.Text = Right("000000000000" + LTrim(str(REC2("EPEL"))), 12)

770     Return

        '<EhFooter>
        Exit Sub

Command13_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command13_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command13_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command14_Click()

        '<EhHeader>
        On Error GoTo Command14_Click_Err

        '</EhHeader>

100     CommonDialog1.ShowOpen

        Dim m As String

        On Error Resume Next

110     m = Text1(0).Text

120     If F_EIKONA = 11 Then
130         m = m + ".JPG"
        End If
    
140     FileCopy CommonDialog1.FILENAME, F_FAKEL_KOINOS + "\images\" + m

        If f_site = "111" Or f_site = "222" Then
            Gdb.Execute "UPDATE EID SET PIC0='" + m + "' WHERE KOD='" + Text1(0).Text + "'"
            Picture2.Picture = LoadPicture()
            Picture2.Refresh
            Picture2.Picture = LoadPicture(F_FAKEL_KOINOS + "\images\" + m)
            StretchSourcePictureFromPicture Me.Picture2.Picture, Me.Picture2
            Picture2.Refresh
        End If
        
        '    foto2(Index).Picture = LoadPicture("c:\mercvb\images\" + m)
        '    StretchSourcePictureFromPicture apot2vm.foto2(Index).Picture, apot2vm.foto2(Index)

        '<EhFooter>
        Exit Sub

Command14_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command14_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command14_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command15_Click()

        '<EhHeader>
        On Error GoTo Command15_Click_Err

        '</EhHeader>

        Dim STR_EKT As String

100     STR_EKT = DATA2.Recordset("kod") + " " + Format(DATA2.Recordset("hm2"), "DD/MM/YYYY") + " "

110     STR_EKT = STR_EKT + Left(DATA2.Recordset("ch2") + Space(25), 25) + " " + Left(DATA2.Recordset("ch4") + Space(20), 20)

120     STR_EKT = STR_EKT + " " + Left(DATA2.Recordset("ch3") + Space(22), 22) + " " + Left(DATA2.Recordset("ono") + Space(25), 25) + " " + Left(DATA2.Recordset("mon") + Space(4), 3) + " " + Format(1, "#0") + " " + Left(DATA2.Recordset("ch1") + Space(15), 15)

        'Printer.FontName = "Courier New"
        'Printer.FontSize = 8

130     Open "LPT1" For Output As 1
140     Print #1, Chr(15) + to437(STR_EKT)
150     Close 1

        '<EhFooter>
        Exit Sub

Command15_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command15_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command15_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command16_Click()

        '<EhHeader>
        On Error GoTo Command16_Click_Err

        '</EhHeader>

        Dim k      As Integer, DUM As Integer, DB As Database, R As New ADODB.Recordset

        Dim X      As String

        Dim mlabel As String

        Dim SQL3   As String

        Dim RECS   As Long

100     If Command16.Caption = "Νέα Εγγραφή" Then
110         Command2.Enabled = False
120         Command3.Enabled = False
130         Command4.Enabled = False
             F_KAT_FPA = Val(FINDPARAMETROI(1, "APOT1", "F_KAT_FPA", "6", "Προεπιλεγμένη κατηγ.ΦΠΑ"))
             
               If F_KAT_FPA - 1 < 0 Then F_KAT_FPA = 1

               Combo2.Text = Combo2.List(F_KAT_FPA - 1)
             
             
             
            

140         Command16.Caption = "Ενημέρωση"

150         For k = 1 To 3
160             Text1(k).Enabled = True
170             Text1(k).ForeColor = vbBlack
            Next

180         For k = 6 To 7
190             Text1(k).Enabled = True
200             Text1(k).ForeColor = vbBlack
            Next

210         For k = 11 To 13
220             Text1(k).Enabled = True
230             Text1(k).ForeColor = vbBlack
            Next

240         Text1(9).Enabled = True
250         pos_kerd2.Enabled = True
260         pos_kerd2.ForeColor = vbBlack

270         Combo1.Enabled = True
280         Combo2.Enabled = True
290         Combo3.Enabled = True

300         Text1(0).Text = ""
310         Text1(1).Text = ""
320         Text1(6).Text = ""
330         lti5.Text = ""
340         lti.Text = ""
350         xti0.Text = ""

360         Combo2.Text = Combo2.List(1)
370         Text1(0).Enabled = True

            If Len(barc.Text) > 0 Then
               Text1(0).Text = barc.Text
            End If
            barc.Text = ""

 If F_KAT_FPA - 1 < 0 Then F_KAT_FPA = 1

               Combo2.Text = Combo2.List(F_KAT_FPA - 1)



380         Text1(0).SetFocus





        Else

            'ελεγχος μην ξαναυπάρχει ο κωδικός
            'Set db = OpenDatabase(gDir, False, False, gConnect)
390         R.Open "select count(*) as ar from EID where  KOD='" + Text1(0).Text + "';", Gdb, adOpenDynamic, adLockOptimistic

400         If R("ar") > 0 And Len(Text1(0).Text) > 0 Then
410             MsgBox "Υπάρχει ο κωδικός"

                Exit Sub

            End If

420         If Val(Left(Combo2.Text, 2)) = 0 Then
430             MsgBox "ΦΠΑ ;"

                Exit Sub

            End If

            Gdb.Execute "insert into BARCODES ( KOD,ERG) VALUES ('" + Text1(0).Text + "','" + Text1(0).Text + "')"



            ' ενημέρωση μετρητή

            '   dum = Update_Next_Pel(x)
            On Error GoTo NoUpdate

440         Data1.Recordset.AddNew

            On Error Resume Next

450         Data1.Recordset("kod") = UCase(Text1(0))  ', "0000")
460         Data1.Recordset("ONO") = Left(Text1(1).Text, Data1.Recordset("ONO").DefinedSize)

470         Data1.Recordset("FPA") = Val(Left(Combo2.Text, 2))

480         Data1.Recordset("mon") = Combo1.Text

            '       DATA2.Recordset("kodlog") = Val(left(Combo3.Text, 3)) ' Combo3.ListIndex  'Right(Combo3.Text, 3)

490         Data1.Recordset("aeg") = Val(Left(Combo4.Text, 3))    ' Combo3.ListIndex  'Right(Combo3.Text, 3)
            'Data1.Recordset("kodlog") = Right(Combo3.Text, 3)

            '       'αν εχω προσθέσει υπο οικογένεια την καταχωρεί αυτόματα
            '       If Val(Combo3.Text) = 0 Then
            '           SQL3 = "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (3," + Str(Combo3.ListCount + 1) + ",'" + Combo3.Text + "'," + Left(Combo4.Text, 2) + ")"
            '           Gdb.Execute SQL3
            '       End If

500         Data1.Recordset("kodlog") = Val(Left(Combo3.Text, 3))

510         Data1.Recordset("xti") = Val(Replace(xti0.Text, ",", "."))
520         Data1.Recordset("lti") = Val(Replace(lti.Text, ",", "."))    'Val(lti.Text)
530         Data1.Recordset("lti5") = Val(Replace(lti5.Text, ",", "."))    ' Val(Text1(10).Text)
540         Data1.Recordset("pos_kerD") = Val(Replace(pos_kerd2.Text, ",", "."))    ' Val(pos_KERD2.Text)

550         Data1.Recordset("prom") = Text1(7).Text
560         Data1.Recordset("spa") = Val(Text1(9).Text)
            'DATA1.Recordset("erg") = Val(Text1(6).Text)
570         Data1.Recordset("ues") = Text1(6).Text

580         Data1.Recordset("kodsynod") = Text1(12).Text
590         Data1.Recordset("kodlogag") = Text1(11).Text

600         Data1.Recordset("CH1") = CH1.Text

610         Data1.Recordset("ERG") = Text1(6).Text    'ΕΡΓ
620         Data1.Recordset("KODERG") = lti5.Text    ' ΚΩΔ.ΠΡΟΜΗΘΕΥΤΗ
630         'DATA1.Recordset("MEMO") = Text14.Text
640         Data1.Recordset("KATHGORIA") = Val(Left(Combo5.Text, 3))    ' Combo3.ListIndex  'Right(Combo3.Text, 3)

            On Error GoTo NoUpdate

650         Data1.Recordset.Update

            On Error Resume Next

660         If F_add_barc = 1 Then   ' file("C:\LAGEURO\MPOYG.TXT") Or
                'ΤΟ ΒΑRCODE ΕΙΝΑΙ Ο ΚΩΔΙΚΟΣ
670             Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + UCase(Text1(0).Text) + "','" + UCase(Text1(0).Text) + "')"
680         ElseIf F_add_barc = 2 Then

                'ΤΟ ΒΑRCODE ΕΙΝΑΙ TO BOHΘ.ΚΛΕΙΔΙ
690             If Val(Text1(6).Text) = 0 Then
700                 GoSub FIND_N_NUMBER
                End If

710             RECS = 0

720             Do While RECS = 0
                    ' Exit Do
730                 Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + UCase(Text1(0).Text) + "','" + UCase(Text1(6).Text) + "')", RECS

740                 If RECS = 0 Then GoSub FIND_N_NUMBER
                Loop

            End If

            'για να ενημερωσει ολα τα πεδια
750         data22.RecordSource = "select  KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where KOD = '" + Text1(0).Text + "';"
760         data22.Refresh

770         mDBGrid1_RowColChange

780         f_TIMOKAT = Val(FINDPARAMETROI(1, "PELAT2", "F_TIMOKAT", "0", "Aριθμός Τιμοκαταλόγων"))

790         If f_TIMOKAT > 0 Then

800             For k = 1 To f_TIMOKAT
810                 Gdb.Execute "INSERT INTO TIMOKAT SELECT KOD,0 AS EKPT, " + str(k) + " AS TIMOK,ONO FROM EID WHERE KOD='" + UCase(Text1(0)) + "'"
                Next

            End If

820         If flag.Caption = "1" Then  ' ΕΡΧΕΤΑΙ ΑΠΟ ΤΑ ΤΙΜΟΛΟΓΙΑ
830             MDIForm1.StatusBar1.Panels(6).Text = Text1(0).Text
840             Unload Me
            End If

850         MHNYMA2.Timer1.Interval = 1000
860         MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " αποθηκεύτηκε με επιτυχία"
870         MHNYMA2.SHOW 1

880         Command16.Caption = "Νέα Εγγραφή"
890         Command2.Enabled = True
900         Command3.Enabled = True
910         Command4.Enabled = True

        End If

        ' Command2.SetFocus

        Exit Sub

NoUpdate:
920     MsgBox "Δεν έγινε η αποθήκευση"

930     Resume Next

        Exit Sub

FIND_N_NUMBER:

        '************************
        Dim rec  As ADODB.Recordset

        'rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
        Dim REC2 As New ADODB.Recordset

940     REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
950     REC2.MoveFirst
960     REC2.Move 10

        'rec.Edit
970     If IsNull(REC2("epel")) Then
980         REC2("epel") = 1
        Else
990         REC2("epel") = REC2("epel") + 1
        End If

1000    REC2.Update

1010    Text1(6).Text = Right("00000000" + LTrim(str(REC2("EPEL"))), 8)
1020    Gdb.Execute "UPDATE EID SET ERG='" + Text1(6).Text + "' WHERE KOD='" + Text1(0).Text + "'"
1030    Return

        '<EhFooter>
        Exit Sub

Command16_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command16_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command16_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        'diortosi
        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim k As Integer, DUM As Integer, DB As Database, R As Recordset

        Dim X As String

100     If Command2.Caption = "Διόρθωση" Then

110         Command2.Caption = "Ενημέρωση"

120         For k = 1 To 3
130             Text1(k).Enabled = True
140             Text1(k).ForeColor = vbBlack
            Next

150         pos_kerd2.Enabled = True
160         pos_kerd2.ForeColor = vbBlack

170         Text1(6).Enabled = True
180         Text1(6).ForeColor = vbBlack

190         For k = 7 To 13
200             Text1(k).Enabled = True
210             Text1(k).ForeColor = vbBlack
            Next

220         ' Text14.Enabled = True

230         ' Text14.ForeColor = vbBlack

240         Combo1.Enabled = True
250         Combo2.Enabled = True
260         Combo3.Enabled = True
270         Combo4.Enabled = True
280         Combo5.Enabled = True
290         pos_kerd.Enabled = True

300         Text1(1).SetFocus

        Else

            On Error Resume Next
            
            
            DATA2.Recordset("APALLFPA") = Val(Left(apallagesFPA.Text, 2))
            apallagesFPA.Enabled = False
            

310         DATA2.Recordset("ONO") = Left(Text1(1).Text, DATA2.Recordset("ONO").DefinedSize)
320         DATA2.Recordset("ONO2") = Left(ONO2.Text, DATA2.Recordset("ONO2").DefinedSize)
330         DATA2.Recordset("FPA") = Val(Left(Combo2.Text, 2))
340         DATA2.Recordset("mon") = Combo1.Text
            'YPOOIKOGENEIA
350         DATA2.Recordset("kodlog") = Left(Combo3.Text, 3)    ' Combo3.ListIndex  'Right(Combo3.Text, 3)
            'OIKOGENEIA
360         DATA2.Recordset("aeg") = Val(Left(Combo4.Text, 3))    ' Combo3.ListIndex  'Right(Combo3.Text, 3)
370         DATA2.Recordset("KATHGORIA") = Val(Left(Combo5.Text, 3))    ' Combo3.ListIndex  'Right(Combo3.Text, 3)
380         DATA2.Recordset("kodSYNOD") = Text1(11).Text     'Right(Combo3.Text, 3)
390         DATA2.Recordset("kodlogAG") = Text1(12).Text     'Right(Combo3.Text, 3)
400         DATA2.Recordset("PROM") = Text1(7).Text     'Right(Combo3.Text, 3)
410         DATA2.Recordset("SPA") = Val(Text1(9).Text)     'Right(Combo3.Text, 3)
420         DATA2.Recordset("PROMHU") = Val(Text1(15).Text)     ' BONUS
430         DATA2.Recordset("UES") = Text1(10).Text     'Right(Combo3.Text, 3)
440         DATA2.Recordset("EPIUYP") = Val(Text1(8).Text)     'Right(Combo3.Text, 3)
450         DATA2.Recordset("koderg") = Text1(16).Text
460         DATA2.Recordset("SYSKMAX") = Text1(18).Text    'SYSKEYASIA
470         DATA2.Recordset("SYSKMIN") = Text1(17).Text    ' YPOSYSKEYASIA
480         DATA2.Recordset("pos_kerd2") = Val(Replace(pos_kerd2.Text, ",", "."))
490         DATA2.Recordset("pos_kerd") = Val(Replace(pos_kerd.Text, ",", "."))    ' Val(pos_KERD.Text)
500         DATA2.Recordset("xti") = Val(Replace(xti0.Text, ",", "."))    ' Val(xti.Text)
510         DATA2.Recordset("lti") = Val(Replace(lti.Text, ",", "."))    'Val(lti.Text)
520         DATA2.Recordset("lti5") = Val(Replace(lti5.Text, ",", "."))    ' Val(lti5.Text)

            DATA2.Recordset("lti2") = Val(Replace(LTI2.Text, ",", "."))    ' Val(lti5.Text)
            DATA2.Recordset("lti3") = Val(Replace(LTI3.Text, ",", "."))    ' Val(lti5.Text)

            DATA2.Recordset("energo") = IIf(chkEnergos.Value = vbChecked, 1, 0)

            If discountOption.Text = discountOption.List(0) Then
                DATA2.Recordset("discountOption") = Null
            ElseIf discountOption.Text = discountOption.List(1) Then
                DATA2.Recordset("discountOption") = 0
            Else
                DATA2.Recordset("discountOption") = 1
            End If
             



530         DATA2.Recordset("ERG") = Text1(6).Text
540         ' DATA2.Recordset("MEMO") = Text14.Text
550         DATA2.Recordset("NUM1") = NUM1
560         DATA2.Recordset("NUM2") = NUM2
570         DATA2.Recordset("NUM3") = NUM3
            DATA2.Recordset("prompol") = Val(Replace(PROMPOL.Text, ",", "."))

580         DATA2.Recordset("CH1") = Trim(CH1)
590         DATA2.Recordset("CH2") = Trim(CH2)
600         DATA2.Recordset("CH3") = Trim(CH3)

610         DATA2.Recordset("CH4") = Trim(ch4)
620         DATA2.Recordset("CH5") = Trim(ch5)
630         DATA2.Recordset("CH6") = Trim(ch6)

            DATA2.Recordset("CPV") = Trim(CPV.Text)

640         DATA2.Recordset("HM1") = HM1
650         DATA2.Recordset("HM2") = HM2
660         DATA2.Recordset("HM3") = HM3
670         DATA2.Recordset("lastupd") = lastupd

            DATA2.Recordset("litra") = Val(Replace(litra.Text, ",", "."))
    
            DATA2.Recordset("efk") = Val(Replace(efk.Text, ",", "."))



            On Error GoTo no_update

680         DATA2.Recordset.Update

            On Error Resume Next

            DATA2.Recordset.Move 0

690         For k = 1 To 3
700             Text1(k).Enabled = False
710             Text1(k).ForeColor = vbBlack

            Next

720         pos_kerd2.Enabled = False
730         Text1(6).Enabled = False
740         pos_kerd2.ForeColor = vbBlack

750         Combo1.Enabled = False
760         Combo2.Enabled = False
770         Combo3.Enabled = False
780         Combo4.Enabled = False
790         Combo5.Enabled = False
800         pos_kerd.Enabled = False

810         For k = 7 To 13
820             Text1(k).Enabled = False
830             Text1(k).ForeColor = vbBlack
            Next

            Dim R22 As New ADODB.Recordset

            Dim CC  As String

840         If Len(Dir("C:\ILIADIS.EXE", vbNormal)) > 0 Then




          




           If Len(Dir("C:\OLATAEIDH.EXE", vbNormal)) > 0 Then
             ' ola ta eidh
              R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE ( CASE WHEN FPA=4 THEN 1 ELSE FPA+1 END ) END)  AS TMHMA  FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
              Open "C:\MERCVB\POSFILE.TXT" For Output As #1
              
              MsgBox " ΠΡΟΣΟΧΗ ΘΑ ΣΤΑΛΟΥΝ ΟΛΑ ΤΑ ΕΙΔΗ ΣΤΟ C:\MERCVB\POSFILE.TXT" + Chr(13) + "META PREPEI NA STALOYN \\st6500-01\tec_pos\Data\POSFILE.TXT"
              MsgBox " ΓΙΑ ΝΑ ΜΗΝ ΞΑΝΑΣΤΑΛΟΥΝ ΟΛΑ ΜΑΖΙ ΑΦΑΙΡΕΣΤΕ ΤΟ  C:\OLATAEIDH.EXE"
           
           Else
              
             'DEBUG  Open "C:\MERCVB\POSFILE.TXT" For Output As #1
              Open "\\st6500-01\tec_pos\Data\POSFILE.TXT" For Output As #1
             ' ena eido
              R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON, ( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE  (CASE when FPA=4 THEN 1 ELSE 23 END )   END  ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE ( CASE WHEN FPA=4 THEN 1 ELSE FPA+1 END ) END)  AS TMHMA FROM EID  WHERE KOD='" + Text1(0) + "'", Gdb, adOpenDynamic, adLockOptimistic
             
           
           End If
           

           
             ' ola ta eidh
             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
              
             'DEBUG  On Error GoTo 0
              On Error GoTo iliadis_error

            Do While Not R22.EOF
                ' Open "C:\POSFILE.TXT" For Output As #1
870             CC = Left(R22(0) + Space(20), 13)
880             CC = CC + Space(18) + Left(R22("ONO") + Space(30), 30)
890             CC = CC + Space(9) + Format(R22("LTI5"), "0000." + String(F_DEK_LIANIKIS, "0"))
     
                '  If IsNull(R22("mon")) Then
                '     CC = CC + Space(3)
                '   Else
                '      CC = CC + R22("mon") '+ Space(3), 3)
                ' End If
     
900             CC = CC + Space(32) + Format(R22("fpa1"), "00")
910             CC = CC + Space(18) + Format(R22("tmhma"), "0.00")
920             Print #1, CC

                  
                R22.MoveNext
            Loop






930             Close #1
940             R22.Close

                On Error Resume Next

            End If






         If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then

'OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
'OK----01       036         30                                    ΠΕΡΙΓΡΑΦΗ                                                              G
'OK----03       087         08                                    ΤΙΜΗ                                                                D,###.00
'OK ----10       095         01                                    ΦΠΑ 1=6,2=13,3=24,5=0                                                                A
'OK'' ----12       107         02                                    ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ 1=ΤΕΜ 2=ΚΙΛΑ                                                    A
'OK----17       095         01                                    ΤΜΗΜΑ                                                              I
'----20       140         01                                    ΕΛΕΥΘΕΡΗ ΤΙΜΗ 0=ΟΧΙ 1=ΝΑΙ                                                        I,S0~1
'----21       137         01        0                           ΖΥΓΙΖΟΜΕΝΟ ΕΙΔΟΣ 0=ΟΧΙ 1=ΝΑΙ                                                              I
'----24       131         04        0                           ΠΟΝΤΟΙ ΕΙΔΟΥΣ                                                           L
'OK  ----25       001         15                                    Β ΚΩΔΙΚΟΣ

            Open "C:\MERCVB\ERRORSPOSFILE.TXT" For Output As #2

            Open "C:\MERCVB\POSFILE.TXT" For Output As #1
             ' ena eidos
              R22.Open "SELECT TOP 1000  EID.KOD,BARCODES.ERG,ONO,LTI5,( CASE WHEN MON IS NULL THEN  '1' ELSE '1' END )  AS MON,( CASE WHEN FPA=4 THEN 1 ELSE ( CASE WHEN FPA=1 THEN 2 ELSE 3 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID RIGHT JOIN BARCODES ON EID.KOD=BARCODES.KOD where  NOT ( EID.KOD LIKE '913%'  ORDER BY EID.KOD) ", Gdb, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
             
             ' ola ta eidh
             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
                On Error GoTo iliadis_error


Dim mLTI5 As Single
Dim MONO As String
Dim MKOD As String
Dim mbar As String
Dim MFPA As Integer
Dim MBARC As String


Dim nc As Long


            Do While Not R22.EOF
                ' Open "C:\POSFILE.TXT" For Output As #1
                
             If IsNull(R22(0)) Then
                MKOD = "."
             Else
                MKOD = R22(0)
             End If
                
                
                
                
             CC = Left(MKOD + Space(21), 20)  ' KVDIKOS
             
             
             
             'R22 ("ONO")
             If IsNull(R22("ono")) Then
                MONO = "."
             Else
                MONO = R22("ono")
             End If
             
             
             'R22 ("ONO")
             If IsNull(R22(1)) Then
                MBARC = "0000"
             Else
                MBARC = R22(1)
             End If
             
             
             
             
             CC = CC + Left(MBARC + Space(21), 15) ' BARCODE
             
             
             
             
             
             CC = CC + Left(MONO + Space(30), 30) + Space(21)
             If IsNull(R22("LTI5")) Then
                mLTI5 = 0
             Else
                mLTI5 = R22("LTI5")
             End If
             
             CC = CC + Replace(Format(mLTI5, "00000.00"), ".", ",") '94   '+ String(F_DEK_LIANIKIS, "0"))
     
                '  If IsNull(R22("mon")) Then
                '     CC = CC + Space(3)
                '   Else
                '      CC = CC + R22("mon") '+ Space(3), 3)
                ' End If
     
             
             
             If IsNull(R22("FPA1")) Then
                MFPA = 3
             Else
                MFPA = R22("FPA1")
             End If
             
             
             DoEvents
             
             Me.Caption = nc
             nc = nc + 1
             
             CC = CC + Format(MFPA, "0") + Space(10)
             CC = CC + " 1"    ' monada Format(R22("mon")
             
             
   '          On Error Resume Next
             
            ' CC = CC + Space(18) + Format(R22("tmhma"), "0.00")
             
             CC = CC + "                       00    0  0      "   ' ποντοι ζυγιζομενα
             CC = Replace(CC, Chr(13), "")
             If Len(CC) < 140 Then
                 Print #2, CC
             Else
                 Print #1, CC
             End If
'123456789012340078900230567890
                  
                R22.MoveNext
            Loop






             Close #1
             
             Close #2
             
             
             
             R22.Close

                On Error Resume Next
                
                Exit Sub

            End If









'
'840         If Len(Dir("C:\ILIADIS.EXE", vbNormal)) > 0 Then
'850             R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  WHERE KOD='" + Text1(0) + "'", Gdb, adOpenDynamic, adLockOptimistic
'
'                On Error GoTo iliadis_error
'
'860             Open "\\st6500-01\tec_pos\Data\POSFILE.TXT" For Output As #1
'                ' Open "C:\POSFILE.TXT" For Output As #1
'870             CC = Left(R22(0) + Space(20), 13)
'880             CC = CC + Space(18) + Left(R22("ONO") + Space(30), 30)
'890             CC = CC + Space(9) + Format(R22("LTI5"), "0000." + String(F_DEK_LIANIKIS, "0"))
'
'                '  If IsNull(R22("mon")) Then
'                '     CC = CC + Space(3)
'                '   Else
'                '      CC = CC + R22("mon") '+ Space(3), 3)
'                ' End If
'
'900             CC = CC + Space(32) + Format(R22("fpa1"), "00")
'910             CC = CC + Space(18) + Format(R22("tmhma"), "0.00")
'920             Print #1, CC
'930             Close #1
'940             R22.Close
'
'                On Error Resume Next
'
'            End If










950         MHNYMA2.Timer1.Interval = 1000
960         MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " αποθηκεύτηκε με επιτυχία"

970         MHNYMA2.SHOW 1

980         kod.SetFocus
990         Command2.Caption = "Διόρθωση"
        End If

        Exit Sub

no_update:

1000    MHNYMA2.Timer1.Interval = 1000
1010    MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " ΔΕΝ αποθηκεύτηκε"
1020    MHNYMA2.BackColor = vbRed
1030    MHNYMA2.SHOW 1
        Exit Sub
        
iliadis_error:
1040    MsgBox "ΠΡΟΣΟΧΗ ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ ΤΟ POS(ΤΑΜΕΙΑΚΗ)"

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Command2_KeyUp_Err

        '</EhHeader>
        Dim a

100     a = 0

        '<EhFooter>
        Exit Sub

Command2_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command2_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command2_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        ' kartella
        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>

100     apot21.Text1.Text = Text1(0).Text
110     apot21.Text2.Text = Text1(1).Text
120     apot21.Text3.Text = " Αποθ." + Combo9.Text
130   '  apot21.Label1.Caption = str(Val(Left(Me.Combo9.Text, 2)))
140     apot21.Label2.Caption = IIf(IsNull(Data1.Recordset("apos")), "", Data1.Recordset("apos"))
      If Val(Left(Me.Combo9.Text, 2)) = 99 Then
         apot21.filtroPEL.Caption = " APOT>0 "
      Else
        apot21.filtroPEL.Caption = " APOT=" + str(Val(Left(Me.Combo9.Text, 2)))
      End If
     If Me.cXRONIES.Value = vbChecked Then
         apot21.cXRONIES.Value = vbChecked
        ' DATA1.ConnectionString = gConnect + ";DATABASE=" + Trim(mID(APOT2.XRONIES.Text, 21, 20))
     Else
         'DATA1.ConnectionString = gConnect
     End If






        'apot2.Hide
'filtroPEL.Caption


150     If cXRONIES.Value = vbChecked Then
160         Gdb.Close
            'f_PROPERSI = Trim(mID(XRONIES.Text, 21, 20))
170         Gdb.Open gConnect + ";DATABASE=" + Trim(mID(XRONIES.Text, 21, 20))
180         apot21.DTPicker1.Value = CDate("01/01/2000")
            ' synt = " YEAR(HME)>1900  "
        End If

        On Error Resume Next

190     apot21.SHOW 1

200     If cXRONIES.Value = vbChecked Then Gdb.Close: Gdb.Open gConnect

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
Sub diastaseis_Grid()

       DBGrid1.columns(1).Width = 3300
       DBGrid1.columns(2).Width = 1000
       DBGrid1.columns(3).Width = 1000
       DBGrid1.columns(2).alignment = dbgLeft
       DBGrid1.columns(2).alignment = dbgRight
End Sub
Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>
        Dim R    As New ADODB.Recordset

        Dim RECS As Long

100     R.Open "SELECT COUNT(*) FROM EGGTIM WHERE KODE='" + Text1(0).Text + "'", Gdb, adOpenForwardOnly, adLockReadOnly

110     If R(0) > 0 Then
120         MsgBox "Αδύνατη η διαγραφή επειδή υπάρχουν κινησεις"
130         R.Close

            Exit Sub

        End If

140     Select Case MsgBox(Text1(1).Text + vbCrLf + "Είσαι σίγουρος για την διαγραφή;", vbYesNo Or vbExclamation Or vbSystemModal Or vbDefaultButton1, "Θα διαγραφεί το " + Text1(0).Text)

            Case vbYes

150         Case vbNo

                Exit Sub

        End Select

160     Gdb.Execute "DELETE FROM EID WHERE KOD='" + Text1(0).Text + "'", RECS

170     If RECS > 0 Then
180         Gdb.Execute "DELETE FROM BARCODES WHERE KOD='" + Text1(0).Text + "'"

190         MHNYMA2.Timer1.Interval = 2000
200         MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " διεγράφη."
210         MHNYMA2.Label1.Caption = "Διαγραφή εγγραφής"
220         MHNYMA2.SHOW 1

            'Call MsgBox("Διεγράφησαν " + Str(RECS) + "" _
             & vbCrLf & " εγγραφές " _
             , vbExclamation Or vbSystemModal, App.title)

        Else
230         Call MsgBox("Δεν ευρέθη η εγγραφή", vbCritical Or vbSystemModal, App.Title)
        End If

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdAnazit_Click()

   anazitisi 1

End Sub


Public Sub anazitisi(esoterika As Integer)  ' 1=apo thn idia forma  0=apo allh forma to kalo
        ' anazitisi
        '<EhHeader>
        On Error GoTo cmdAnazit_Click_Err

        '</EhHeader>
        Dim mk As String

        Dim mO As String

        On Error Resume Next

      
       
       
       
       

100    ' DBGrid1.ColWidth(1) = 3300

110    ' DBGrid1.ColWidth(2) = 1250
120    ' DBGrid1.ColAlignment(2) = 7

130   '  DBGrid1.ColWidth(3) = 1250
140     'DBGrid1.ColAlignment(3) = 7



  Dim xti_or_pos As Integer

        xti_or_pos = Val(FINDPARAMETROI(1, "APOT2", "xti_or_pos", "1", "ΣΤΟ GRID 1=ΠΟΣΟΤΗΤΑ 2=ΤΙΜ.ΑΓΟΡΑΣ"))



 Dim PED2 As String
        If xti_or_pos = 2 Then
            PED2 = " CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] "
        Else
            
          'round(  ISNULL(SUM(XRE),0)-ISNULL(SUM(PIS),0)  ,1) AS [ΥΠΟΛ]" & " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' AND " + synt + "  GROUP BY KODE,APOT"
            PED2 = "CONVERT(decimal(10,2),  ( SELECT ISNULL(SUM(XRE),0)-ISNULL(SUM(PIS),0) FROM EGGTIM WHERE KODE=EID.KOD),1)  AS [ΥΠΟΛΟΙΠΟ] "
        End If
        
        
        
    Dim PED1 As String
 If xti_or_pos > 10 Then
     PED1 = "CONVERT(decimal(10,2),LTI) AS [XONΔΡ.ΤΙΜΗ],"
     
 Else
     PED1 = "CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],"
 End If
        
        
        
        
        

150     If Len(ono.Text) > 0 Then
160         If ypoloipo Then

170             If F_UPPER = 1 Then
180                 mO = SameLetters(ono.Text)
                Else
190                 mO = ono.Text
                End If

                '     mO = SameLetters(ono.Text)
200             If Val(barc) > 0 Then
210                 data22.RecordSource = "select  KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH] from EID where ENERGO IN ('" + fmEN + "') and  POS>0 and CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ]>" + str(Val(Round(barc, 2))) + " and ONO like '" + mO + "%' order by ONO;"
                Else
220                 data22.RecordSource = "select   KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH] from EID where ENERGO IN ('" + fmEN + "') and  POS>0 and ONO like '" + mO + "%'  order by ONO ;"
                End If

            Else  'If ypoloipo Then

                Dim ONOM As String

                Dim KODI As String

                Dim DOL  As Long

                'SYNDYASMENH ANAZHTHSH ME $  DHLADH   ΣΚΙΑ$99 ΟΛΕΣ ΤΙΣ "ΣΚΙΕΣ" ΜΕ ΚΩΔΙΚΟ ΠΟΥ ΑΡΧΙΖΕΙ ΑΠΟ "99"
230             ONOM = ono.Text

240             DOL = InStr(1, ONOM, "$")
250             KODI = ""

260             If DOL > 1 Then
270                 KODI = Trim(mID(ONOM, DOL + 1, 14))
280                 ONOM = Left(ONOM, DOL - 1)

290                 If F_UPPER = 1 Then
300                     mO = SameLetters(ONOM)
                    Else
310                     mO = ONOM
                    End If

320                 data22.RecordSource = "select  KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH]  from EID where ENERGO IN ('" + fmEN + "') and  KOD LIKE '" + KODI + "%' AND ONO like '" + UCase(mO) + "%'  order by ONO;"
                Else

330                 If F_UPPER = 1 Then
340                     mO = SameLetters(ono.Text)
                    Else
350                     mO = ono.Text
                    End If

                    '                   mO = SameLetters(ono.Text)
360                 If Val(barc) > 0 Then
370                     data22.RecordSource = "Select   KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH]  from EID where ENERGO IN ('" + fmEN + "') and  CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ]>" + str(Val(Round(barc, 2))) + " and ONO like '" + UCase(mO) + "%'  order by ONO ;"
                    Else
380                     data22.RecordSource = "select  KOD,ONO," + PED1 + PED2 + "," & " (SELECT  (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END ) FROM EGGTIM WHERE EGGTIM.KODE=EID.KOD ) AS [ΑΓΟΡΕΣ] " & "  from EID where ENERGO IN ('" + fmEN + "') and   ONO like '" + mO + "%'  order by ONO;"
                    End If
                End If
            End If    'If ypoloipo Then

            ' f_Refresh = False
390         data22.Refresh
            ' f_Refresh = True

400         If data22.Recordset.EOF Then
410             Text1(0).Text = " "
420             Text1(1).Text = " "

            End If

            diastaseis_Grid
            




430         DBGrid1.SetFocus
440         mDBGrid1_RowColChange

            Exit Sub

        End If  ' If Len(ono.Text) > 0 Then

450     If Len(kod) > 0 Then

460         mk = SameLetters(kod.Text)

470         If ypoloipo Then
480             If Val(barc) > 0 Then
490                 data22.RecordSource = "select  KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH]  from EID where ENERGO IN ('" + fmEN + "') and  POS>0 and CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ]>" + str(Val(Round(barc.Text, 2))) + " and KOD like '" + mk + "%'   ESCAPE '$' order by KOD ;"
                Else
500                 data22.RecordSource = "select   KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH]  from EID where ENERGO IN ('" + fmEN + "') and  POS>0 and KOD like '" + mk + "%' ESCAPE '$'  order by KOD ;"
                End If

            Else

510             If Val(barc) > 0 Then
520                 data22.RecordSource = "select  KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH]   from EID where ENERGO IN ('" + fmEN + "') and   CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ]>" + str(Val(Round(barc, 2))) + " and KOD like '" + mk + "%'   ESCAPE '$'  order by KOD;"
                Else
530                 data22.RecordSource = "select  KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH]  from EID where ENERGO IN ('" + fmEN + "') and  KOD like '" + mk + "%'  ESCAPE '$' order by KOD ;"
                End If
            End If

            '  f_Refresh = False
540         data22.Refresh
            '  f_Refresh = True

            '     DATA2.Refresh
            '       CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ].SetFocus
            '       Data2.Refresh

550         If data22.Recordset.EOF Then    ' δεν βρηκε κωδικούς
560             Text1(0).Text = " "
570             Text1(1).Text = " "
580             kod.SetFocus
            Else
590             DBGrid1.SetFocus
            End If

        End If

600     If Len(erg) > 0 Then

            Dim mERG

610         mERG = SameLetters(erg.Text)

620         If Left(mERG, 1) = "+" Then
630             data22.RecordSource = UCase("select  KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH]  from EID where ENERGO IN ('" + fmEN + "') and  KODERG = '" + Trim(mID(erg.Text, 2, 10)) + "'   ;")
            Else
640             data22.RecordSource = UCase("select  KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH]  from EID where ENERGO IN ('" + fmEN + "') and  ERG like '" + mERG + "%'  ESCAPE '$' ;")
            End If

650         data22.Refresh

660         If data22.Recordset.EOF Then
670             Text1(0).Text = " "
680             Text1(1).Text = " "
            End If

690         DBGrid1.SetFocus

        End If

700     If Len(barc) > 0 Then

            On Error GoTo cmdAnazit_Click_Err

710         data22.RecordSource = "select  KOD,ONO," + PED1 + PED2 + ",XTI AS [ΤΙΜ.ΑΓΟΡΑΣ],LTI5 AS [LIAN.TIMH]  from EID where ENERGO IN ('" + fmEN + "') and  KOD IN (SELECT KOD FROM BARCODES WHERE ERG like '" + barc.Text + "%' ESCAPE '$' )"
720         data22.Refresh

730         If data22.Recordset.EOF Then
740             Text1(0).Text = " "
750             Text1(1).Text = " "
            End If

760         DBGrid1.SetFocus

        End If

770     mDBGrid1_RowColChange

diastaseis_Grid

        '<EhFooter>
        Exit Sub

cmdAnazit_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>
        Dim neos As String, RECS As Integer

100     neos = InputBox("Δώσε τον νέο κωδικό ")

        Dim R As New ADODB.Recordset

110     If Len(neos) = 0 Then
120         Call MsgBox("Δεν δώσατε νέο κωδικό", vbCritical Or vbDefaultButton1, App.Title)

            Exit Sub

        End If

130     If neos = Text1(0).Text Then
140         MsgBox "Ο νέος κωδικός είναι ίδιος με τον παλιό"

            Exit Sub

        End If

150     R.Open "select count(*) from EID WHERE KOD='" + neos + "'", Gdb, adOpenForwardOnly, adLockReadOnly

160     If R(0) = 0 Then
170         Gdb.Execute "UPDATE EID SET KOD='" + neos + "' WHERE KOD='" + Text1(0).Text + "'", RECS
180         MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")

190         Gdb.Execute "UPDATE EGGTIM SET KODE='" + neos + "' WHERE KODE='" + Text1(0).Text + "'", RECS
200         MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ ΣΕ ΤΙΜΟΛΟΓΙΑ  " + Format(RECS, "#####")

210         Gdb.Execute "UPDATE BARCODES SET KOD='" + neos + "' WHERE KOD='" + Text1(0).Text + "'", RECS
220         MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ ΣΕ BARCODES  " + Format(RECS, "#####")
        Else

            Dim ANS

230         If F_sygx = 0 Then
240             MsgBox "Υπάρχει ήδη ο κωδικός "

                Exit Sub

            End If

250         ANS = MsgBox("Υπάρχει ήδη ο κωδικός " + neos + Chr(13) + " Nα συγχωνευθεί ο παλιός με τον νέο κωδικό;", vbYesNo)

260         If ANS = vbYes Then

270             Gdb.Execute "delete FROM  EID  WHERE KOD='" + Text1(0).Text + "'", RECS
                'MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")

280             Gdb.Execute "UPDATE EGGTIM SET KODE='" + neos + "' WHERE KODE='" + Text1(0).Text + "'", RECS
290             MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ ΣΕ ΤΙΜΟΛΟΓΙΑ  " + Format(RECS, "#####")

300             Gdb.Execute "UPDATE BARCODES SET KOD='" + neos + "' WHERE KOD='" + Text1(0).Text + "'", RECS
310             MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ ΣΕ BARCODES  " + Format(RECS, "#####")

            End If

        End If


On Error Resume Next
        'αλλαζει τον κωδικο στις κατηγορίες
        Gdb.Execute "update PRODUCT_TO_CATEGORIES set KOD='" + neos + "' WHERE KOD='" + Text1(0).Text + "';"
        Gdb.Execute "update PRODUCT_ATTRIBUTES set KOD='" + neos + "' WHERE KOD='" + Text1(0).Text + "';"
        
        'Gdb.Execute "INSERT INTO PRODUCT_ATTRIBUTES " & "(KOD,OPTIONS_ID,OPTIONS_VALUES_ID,PRODUCT_ID) " & " VALUES ('" + kod + "',1," + str(f_hlikies(K)) + "," + str(EID.Recordset("ID")) + ");"









        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub PAYLOS()

        '========================etiketes ========================================================
        '<EhHeader>
        On Error GoTo PAYLOS_Err

        '</EhHeader>
        Dim d1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        Dim mlabel As String

        Dim MPER As String

        Dim mtimes As Single

        Dim TIMM As Single

        Dim timologio As String

        Dim XEIR As String

        'On Error Resume NextN  N

        Dim X As Printer

100     For Each X In Printers

110         If InStr(X.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
120             Set Printer = X

                ' Stop looking for a printer.
                Exit For

            End If

        Next

        On Error GoTo PAYLOS_Err

130     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

        Dim r1 As New ADODB.Recordset

140     r1.Open "SELECT * FROM BARCODES WHERE KOD='" + DATA2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

150     If r1.EOF Then
160         GoSub FIND_N_NUMBER
        Else

170         If Val(r1("ERG")) = 0 Then
180             GoSub FIND_N_NUMBER
            Else
190             mlabel = r1("erg")
            End If
        End If

200     MPER = DATA2.Recordset("ono2")
210     mtimes = InputBox("Πόσα κουτιά;", MPER, 1)

        Dim nTIMES

220     nTIMES = 2
230     nTIMES = InputBox("ΕΤΙΚΕΤΕΣ/ΚΟΥΤΙ;", , nTIMES, 1)

240     XEIR = Format(Now, "DD/MM/YYYY")
250     XEIR = InputBox("ΗΜΕΡΟΜΗΝΙΑ-ΧΧ-ΧΧ ΧΕΙΡΙΣΤΕΣ;", , XEIR, 1)

        'TIMM = " " ' Format(DATA2.Recordset("lti5"), "#####.00")

        'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

260     For k = 1 To Val(mtimes) * nTIMES

270         If True Then   ' Len(mlabel) <> 13 Then
280             Printer.FontBold = True
290             Printer.FontSize = 12
300             Printer.CurrentX = 100
310             Printer.CurrentY = -50
320             Printer.Print Tab(2); MPER
330             Printer.Print Tab(2); XEIR
340             Printer.Print Tab(2); mlabel
350             BarCode1.ShowText = NO
360             BarCode1.NarrowBarWidth = 0.025  '0.03
370             BarCode1.DataToEncode = mlabel
380             Printer.PaintPicture BarCode1.Picture, 900, 500  ' '-88,124
390             Printer.EndDoc
            Else
400             Printer.Print Tab(F_ArPerEtik + 10); Left(MPER, 25)
410             Printer.FontName = "EAN13"    ' "128"
420             Printer.Print Tab(F_ArPerEtik + 10); mlabel
430             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
440             Printer.FontSize = 6
450             Printer.Print Tab(F_ArPerEtik + 8); DATA2.Recordset("kod")    ' mlabel
460             Printer.FontSize = 8
470             Printer.EndDoc
            End If

        Next
    
        Dim aaparast As Long, ANS As Integer

480     ANS = MsgBox("Να αποθηκευτούν;", vbYesNo)

490     If ANS = vbNo Then

            Exit Sub

        End If

500     aaparast = gfind_parastat("S", 0)
        'ReadWrite=0   Βρίσκει τον αριθμο παραστατικού          -------------------------
        'ReadWrite=1   Aποθηκεύει τον αριθμο παραστατικού + 1   -------------------------
        'par ειναι το είδος του παραστατικου  π.χ. par="T"
    
        Dim RB As New ADODB.Recordset

        Dim re As New ADODB.Recordset

        Dim rt As New ADODB.Recordset
    
510     re.Open "SELECT top 10 *FROM EGGTIM", Gdb, adOpenDynamic, adLockOptimistic
520     rt.Open "SELECT top 10 *FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
530     re.AddNew

540     If IsDate(Left(XEIR, 10)) Then
550         re("hme") = CDate(Left(XEIR, 10))
        Else
560         re("hme") = Now
        End If

570     re("atim") = "S" + Format(aaparast, "000000")
580     re("poso") = DATA2.Recordset("SYSKMAX") * mtimes
590     re("timm") = 0
600     re("kode") = Text1(0)
610     re("onoma") = Text1(1)
620     re("mona") = Combo1.Text
630     re("kau_ajia") = 0
640     re("MIK_AJIA") = 0
650     re("fpa") = 2
660     re("ekpt") = 0
670     re("pelkod") = "0001"
680     re("eidos") = "r"
690     re("APOT") = 1
700     re("PIS") = 0
710     re("XRE") = DATA2.Recordset("SYSKMAX") * mtimes
720     re.Update
    
        ' ΑΦΑΙΡΕΡΣΗ Α ΥΛΩΝ

730     If DATA2.Recordset("NUM1") > 0 Then
740         re.AddNew

750         If IsDate(Left(XEIR, 10)) Then
760             re("hme") = CDate(Left(XEIR, 10))
            Else
770             re("hme") = Now
            End If

780         re("atim") = "S" + Format(aaparast, "000000")
790         re("poso") = Round(DATA2.Recordset("SYSKMAX") * mtimes * DATA2.Recordset("NUM1") / 1000 * 1.02, 5)
800         re("timm") = 0
810         re("kode") = Trim(DATA2.Recordset("CH1"))  'ΣΥΣΧΕΤΙΣΜΕΝΟΣ ΚΩΔΙΚΟΣ Α ΥΛΗΣ
820         re("onoma") = ""
830         re("mona") = "ΚΙΛ"
840         re("kau_ajia") = 0
850         re("MIK_AJIA") = 0
860         re("fpa") = 2
870         re("ekpt") = 0
880         re("pelkod") = "0001"
890         re("eidos") = "r"
900         re("APOT") = 1
910         re("PIS") = Round(DATA2.Recordset("SYSKMAX") * mtimes * DATA2.Recordset("NUM1") / 1000 * 1.02, 5)
920         re("XRE") = 0
930         re.Update
        End If

940     rt.AddNew
    
950     If IsDate(Left(XEIR, 10)) Then
960         rt("hme") = CDate(Left(XEIR, 10))
        Else
970         rt("hme") = Now
        End If

980     rt("atim") = "S" + Format(aaparast, "000000")
        ' rt("kleidi") = "φ" +
990     rt("AJ1") = 0
1000    rt("AJ3") = 0
1010    rt("AJ4") = 0
1020    rt("AJ5") = 0

        'rt("KLEIDI") = "φ" + m_tim

1030    rt("aj2") = 0
1040    rt("fpa2") = 0
1050    rt("kpe") = "0001"
1060    rt("aji") = 0
1070    rt("eidos") = "r"
1080    rt("EIDPAR") = ""
1090    rt("B_N1") = 1
1100    rt("PARAT") = XEIR
       
1110    rt.Update
    
1120    aaparast = gfind_parastat("S", 1)
    
1130    MsgBox "αποθηκευτηκαν"
    
        Exit Sub

FIND_N_NUMBER:
        '************************
1140    rec.MoveFirst
1150    rec.Move 10

        'rec.Edit
1160    If IsNull(rec("epel")) Then
1170        rec("epel") = 1
        Else
1180        rec("epel") = rec("epel") + 1
        End If

1190    rec.Update
1200    mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)
1210    Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + DATA2.Recordset("KODE") + "'"
1220    Return

        '<EhFooter>
        Exit Sub

PAYLOS_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.PAYLOS " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.PAYLOS " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub LAG_ETIK()

        '========================etiketes ========================================================
        '<EhHeader>
        On Error GoTo LAG_ETIK_Err

        '</EhHeader>
        Dim d1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        Dim mlabel As String

        Dim MPER As String

        Dim mtimes As Single

        Dim TIMM As Single

        Dim timologio As String

        'On Error Resume NextN  N

        Dim X As Printer

100     For Each X In Printers

110         If InStr(X.DeviceName, f_printer) > 0 Then
                ' Set printer as system default.
120             Set Printer = X

                ' Stop looking for a printer.
                Exit For

            End If

        Next

        On Error GoTo LAG_ETIK_Err

130     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

        Dim r1 As New ADODB.Recordset

140     r1.Open "SELECT * FROM BARCODES WHERE KOD='" + DATA2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

150     If r1.EOF Then
160         GoSub FIND_N_NUMBER
        Else

170         If Val(r1("ERG")) = 0 Then
180             GoSub FIND_N_NUMBER
            Else
190             mlabel = r1("erg")
            End If
        End If

200     MPER = DATA2.Recordset("ono")
210     mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)

220     TIMM = Format(DATA2.Recordset("lti5"), "#####." + String(F_DEK_LIANIKIS, "0"))

        'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

230     For k = 1 To Val(mtimes)

240         If True Then   ' Len(mlabel) <> 13 Then
250             Printer.FontName = "128"
260             Printer.FontSize = 6
270             Printer.Print Tab(F_ArPerEtik + 5); mlabel
280             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
290             Printer.FontSize = 8
300             Printer.Print
                '      Printer.Print
                '      Printer.Print
                '      Printer.Print

310             Printer.Print Tab(F_ArPerEtik + 10); mlabel + "  Kωδ." + Text1(0).Text
320             Printer.Print Tab(F_ArPerEtik + 10); Left(MPER, 25)
330             Printer.Print Tab(F_ArPerEtik + 10); timologio
340             Printer.FontSize = 12
350             Printer.FontBold = True
360             Printer.Print Tab(F_ArPerEtik + 10); Format(Val(Trim(LTrim(TIMM))), "00.00") + " €"
370             Printer.FontBold = False

380             Printer.EndDoc
            Else
390             Printer.Print Tab(F_ArPerEtik + 10); Left(MPER, 25)

400             Printer.FontName = "EAN13"    ' "128"
410             Printer.Print Tab(F_ArPerEtik + 10); mlabel
420             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
430             Printer.FontSize = 6
440             Printer.Print Tab(F_ArPerEtik + 8); DATA2.Recordset("kod")    ' mlabel

450             Printer.FontSize = 8
460             Printer.Print
470             Printer.Print
480             Printer.Print
490             Printer.Print
500             Printer.EndDoc
            End If

        Next

        Exit Sub

FIND_N_NUMBER:
        '************************
510     rec.MoveFirst
520     rec.Move 10

        'rec.Edit
530     If IsNull(rec("epel")) Then
540         rec("epel") = 1
        Else
550         rec("epel") = rec("epel") + 1
        End If

560     rec.Update
570     mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)
580     Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + DATA2.Recordset("KODE") + "'"
590     Return

        '<EhFooter>
        Exit Sub

LAG_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.LAG_ETIK " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.LAG_ETIK " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub diesel_ETIK()

        '========================etiketes ========================================================
        '<EhHeader>
        On Error GoTo diesel_ETIK_Err

        '</EhHeader>
        Dim d1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        Dim mlabel As String

        Dim MPER As String

        Dim mtimes As Single

        Dim TIMM As Single

        Dim timologio As String

        'On Error Resume NextN  N

        Dim X As Printer

100     For Each X In Printers

110         If InStr(X.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
120             Set Printer = X

                ' Stop looking for a printer.
                Exit For

            End If

        Next

        On Error GoTo diesel_ETIK_Err

130     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

        Dim r1 As New ADODB.Recordset

140     r1.Open "SELECT * FROM BARCODES WHERE KOD='" + DATA2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

150     If r1.EOF Then
160         GoSub FIND_N_NUMBER
        Else

170         If Val(r1("ERG")) = 0 Then
180             GoSub FIND_N_NUMBER
            Else
190             mlabel = r1("erg")
            End If
        End If

200     MPER = DATA2.Recordset("ono")
210     mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)

220     TIMM = Format(DATA2.Recordset("lti5"), "#####." + String(F_DEK_LIANIKIS, "0"))

        'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

230     For k = 1 To Val(mtimes)

240         If True Then

                ' BarCode2.SymbologyID = CODE128
                ' BarCode2.DataToEncode = mlabel

                ' Picture2.height = F_PIChEIGHT
                'Picture2.width = F_PICwidth

                'Printer.PaintPicture BarCode2.Picture, F_XBARCODE, F_YBARCODE

250             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
260             Printer.FontSize = 8
270             Printer.Print
                '      Printer.Print
                '      Printer.Print10002357

                '      Printer.Print

280             Printer.Print Tab(F_ArPerEtik + 10); mlabel + "  Kωδ." + Text1(0).Text
290             Printer.Print Tab(F_ArPerEtik + 10); Left(MPER, 25)
300             Printer.Print Tab(F_ArPerEtik + 10); timologio
310             Printer.FontSize = 12
320             Printer.FontBold = True
330             Printer.Print Tab(F_ArPerEtik + 10); Format(Val(Trim(LTrim(TIMM))), "00.00") + " €"
340             Printer.FontBold = False

350             Printer.CurrentY = F_YBARCODE

360             Printer.FontName = "128"
370             Printer.FontSize = 6
380             Printer.Print Tab(12); mlabel

390             Printer.EndDoc
            Else
400             Printer.Print Tab(F_ArPerEtik + 10); Left(MPER, 25)
410             Printer.Print Tab(F_ArPerEtik + 20); mlabel
420             Printer.FontName = "ean13"    ' "128"
430             Printer.FontSize = 6
440             Printer.Print Tab(F_ArPerEtik + 8); DATA2.Recordset("kod")    ' mlabel
450             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
460             Printer.FontSize = 8
470             Printer.Print
480             Printer.Print
490             Printer.Print
500             Printer.Print
510             Printer.EndDoc
            End If

        Next

        Exit Sub

FIND_N_NUMBER:
        '************************
520     rec.MoveFirst
530     rec.Move 10

        'rec.Edit
540     If IsNull(rec("epel")) Then
550         rec("epel") = 1
        Else
560         rec("epel") = rec("epel") + 1
        End If

570     rec.Update
580     mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)
590     Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + DATA2.Recordset("KOD") + "'"

600     Gdb.Execute "INSERT INTO BARCODES (ERG,KOD) VALUES('" + mlabel + "','" + DATA2.Recordset("KOD") + "')"

610     Return

        '<EhFooter>
        Exit Sub

diesel_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.diesel_ETIK " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.diesel_ETIK " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub LAG2_ETIK()

        '========================etiketes ========================================================
        '<EhHeader>
        On Error GoTo LAG2_ETIK_Err

        '</EhHeader>
        Dim d1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        Dim mlabel As String

        Dim MPER As String

        Dim mtimes As Single

        Dim TIMM As Single

        Dim timologio As String

        'On Error Resume NextN  N

        Dim X As Printer

100     For Each X In Printers

110         If InStr(X.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
120             Set Printer = X

                ' Stop looking for a printer.
                Exit For

            End If

        Next

        On Error GoTo LAG2_ETIK_Err

130     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

        Dim r1 As New ADODB.Recordset

140     r1.Open "SELECT * FROM BARCODES WHERE KOD='" + DATA2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

150     If r1.EOF Then
160         GoSub FIND_N_NUMBER
        Else

170         If Val(r1("ERG")) = 0 Then
180             GoSub FIND_N_NUMBER
            Else
190             mlabel = r1("erg")
            End If
        End If

200     MPER = DATA2.Recordset("ono")
210     mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)

220     TIMM = Format(DATA2.Recordset("lti5"), "#####." + String(F_DEK_LIANIKIS, "0"))

        'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

        '---------------------------------------------------------------------------
        '    If Len(DataGrid1.Text) = 12 Then
        '         BarCode2.SymbologyID = UPCA
        '    Else
        '         BarCode2.SymbologyID = EAN13
        '    End If
        '
        '    For Each X In Printers
        '        If InStr(X.DeviceName, "2844") > 0 Then
        '          ' Set printer as system default.
        '
        '          Set Printer = X
        '          ' Stop looking for a printer
        '          Exit For
        '        End If
        '    Next
        '
        '    BarCode2.DataToEncode = DataGrid1.Text
        '
        '
        '   For K = 1 To fores
        '    Printer.CurrentY = 400
        '   ' Printer.CurrentX = 2048
        '    'Printer.Print s
        '    Printer.CurrentX = 600 + F_ArPerEtik * 100
        '    Printer.FontSize = 12
        '
        '    Printer.Print Text1(0).Text
        '    Printer.PaintPicture BarCode2.Picture, 348, 824
        '    Printer.EndDoc
        '   Next

230     F_ArPerEtik = 5

        '------------------------------------------------------------------------
240     For k = 1 To Val(mtimes)

250         If True Then
260             Printer.FontName = "128"
270             Printer.FontSize = 6
280             Printer.Print Tab(12); mlabel
290             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
300             Printer.FontSize = 8
310             Printer.Print
320             Printer.Print
330             Printer.Print
340             Printer.Print

350             Printer.Print Tab(F_ArPerEtik + 10); mlabel
360             Printer.Print Tab(F_ArPerEtik + 10); MPER
370             Printer.Print Tab(F_ArPerEtik + 10); timologio
380             Printer.FontSize = 14
390             Printer.FontBold = True
400             Printer.Print Tab(F_ArPerEtik + 10); Format(Val(Trim(LTrim(TIMM))), "00.00") + " €"
410             Printer.FontBold = False

420             Printer.EndDoc
            Else
430             Printer.Print Tab(F_ArPerEtik + 10); MPER
440             Printer.Print Tab(F_ArPerEtik + 20); mlabel

450             BarCode2.SymbologyID = 128
460             BarCode2.DataToEncode = DataGrid1.Text
470             Printer.PaintPicture BarCode2.Picture, 348, 824

                '
                '
                '
                '      Printer.FontName = "ean13" ' "128"
                '      Printer.FontSize = 6
                '
                '
                '
                '      Printer.Print Tab(F_ArPerEtik + 8); DATA2.Recordset("kod") ' mlabel
                '      Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
480             Printer.FontSize = 8
490             Printer.Print
500             Printer.Print
510             Printer.Print
520             Printer.Print
530             Printer.EndDoc
            End If

        Next

        Exit Sub

FIND_N_NUMBER:
        '************************
540     rec.MoveFirst
550     rec.Move 10

        'rec.Edit
560     If IsNull(rec("epel")) Then
570         rec("epel") = 1
        Else
580         rec("epel") = rec("epel") + 1
        End If

590     rec.Update
600     mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)
610     Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + DATA2.Recordset("KODE") + "'"
620     Return

        '<EhFooter>
        Exit Sub

LAG2_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.LAG2_ETIK " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.LAG2_ETIK " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub patrick_etik()

        '========================etiketes ========================================================
        '<EhHeader>
        On Error GoTo patrick_etik_Err

        '</EhHeader>
        Dim d1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        Dim mlabel As String

        Dim MPER As String

        Dim mtimes As Single

        Dim TIMM As Single

        Dim timologio As String

        'On Error Resume NextN  N

        Dim nn As Long

        Dim X As Printer

100     For Each X In Printers

110         If InStr(X.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
120             Set Printer = X

                ' Stop looking for a printer.
                Exit For

            End If

        Next

        On Error GoTo patrick_etik_Err

130     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

        Dim r1 As New ADODB.Recordset

140     r1.Open "SELECT * FROM BARCODES WHERE KOD='" + DATA2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

150     If r1.EOF Then
160         GoSub FIND_N_NUMBER  ' EPISTREFEI TO mlabel
        Else

170         If Val(r1("ERG")) = 0 Then
180             GoSub FIND_N_NUMBER
            Else
190             mlabel = r1("ERG")
            End If
        End If

200     MPER = DATA2.Recordset("ono")
210     mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)

220     TIMM = Format(DATA2.Recordset("lti5"), "#####." + String(F_DEK_LIANIKIS, "0"))

        'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

230     For k = 1 To Val(mtimes)

240         Printer.FontSize = 8
250         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName

260         Printer.Print Tab(0); MPER

            'Printer.FontName = "EAN8"
270         Printer.FontName = "EAN8"
280         Printer.FontSize = 8
290         Printer.Print Tab(0); mlabel

300         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
310         Printer.Print
320         Printer.Print
330         Printer.Print

            'If Val(TIMM) < 10 Then
            '    Printer.FontSize = 14
            'Else
340         Printer.FontSize = 12
            'End If

350         Printer.FontBold = True

360         Printer.Print Tab(1); Format(Val(TIMM), "###0.00") + " €"

370         Printer.FontSize = 8
380         Printer.FontBold = False

390         Printer.EndDoc

        Next

        Exit Sub

FIND_N_NUMBER:
        '************************
400     rec.MoveFirst
410     rec.Move 10

        'rec.Edit
420     If IsNull(rec("epel")) Then
430         rec("epel") = 1
        Else
440         rec("epel") = rec("epel") + 1
        End If

450     rec.Update
460     mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)

470     Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + DATA2.Recordset("KOD") + "','" + mlabel + "')", nn

480     If nn = 0 Then
490         MsgBox "ΔΕΝ ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ BARCODE"
        End If

        '  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Text1(0).Text + "'"
500     Return

        '<EhFooter>
        Exit Sub

patrick_etik_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.patrick_etik " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.patrick_etik " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub SOKRATIS_etik()

        '========================etiketes ========================================================
        '<EhHeader>
        On Error GoTo SOKRATIS_etik_Err

        '</EhHeader>
        Dim d1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        Dim mlabel As String

        Dim MPER As String

        Dim mtimes As Single

        Dim TIMM As Single

        Dim timologio As String

        'On Error Resume NextN  N

        Dim nn As Long

        Dim X As Printer

100     For Each X In Printers

110         If InStr(X.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
120             Set Printer = X

                ' Stop looking for a printer.
                Exit For

            End If

        Next

        On Error GoTo SOKRATIS_etik_Err

130     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

        Dim r1 As New ADODB.Recordset

140     r1.Open "SELECT * FROM BARCODES WHERE KOD='" + DATA2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

150     If r1.EOF Then
160         GoSub FIND_N_NUMBER  ' EPISTREFEI TO mlabel
        Else

170         If Val(r1("ERG")) = 0 Then
180             GoSub FIND_N_NUMBER
            Else
190             mlabel = r1("ERG")
            End If
        End If

200     MPER = DATA2.Recordset("ono")
210     mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)

220     TIMM = Format(DATA2.Recordset("lti5"), "#####." + String(F_DEK_LIANIKIS, "0"))

        'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

230     For k = 1 To Val(mtimes)

240         Printer.FontSize = 6
250         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
260         Printer.FontBold = True

270         Printer.CurrentY = 12

280         Printer.Print Tab(0); MPER

            'If Len(DataGrid1.Text) = 12 Then
            ' BarCode1.SymbologyID = CODE128
            'Else
            '     BarCode1.SymbologyID = EAN13
            'End If
290         BarCode1.ShowText = NO
300         BarCode1.NarrowBarWidth = 0.025  '0.03
310         BarCode1.DataToEncode = DataGrid1.Text
320         Printer.PaintPicture BarCode1.Picture, -18, F_YBARCODE  ' '-88,124

            'Printer.FontName = "EAN8"
            'Printer.FontName = "EAN8"
            'Printer.FontSize = 8
330         Printer.CurrentY = 900
340         Printer.Print ; Text1(0).Text; " "; mlabel
350         Printer.FontSize = 8
360         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
            'Printer.Print
            'Printer.Print
            'Printer.Print
            'Printer.Print

370         Printer.CurrentY = 900
            'Printer.Print Tab(5); Format(Val(TIMM), "###0.00") + " ΕΥΡΩ "
380         Printer.EndDoc

        Next

        Exit Sub

FIND_N_NUMBER:
        '************************
390     rec.MoveFirst
400     rec.Move 10

        'rec.Edit
410     If IsNull(rec("epel")) Then
420         rec("epel") = 1
        Else
430         rec("epel") = rec("epel") + 1
        End If

440     rec.Update
450     mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)

460     Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + DATA2.Recordset("KOD") + "','" + mlabel + "')", nn

470     If nn = 0 Then
480         MsgBox "ΔΕΝ ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ BARCODE"
        End If

        '  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Text1(0).Text + "'"
490     Return

        '<EhFooter>
        Exit Sub

SOKRATIS_etik_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.SOKRATIS_etik " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.SOKRATIS_etik " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub GATIDIS_ETIK()

        '<EhHeader>
        On Error GoTo GATIDIS_ETIK_Err

        '</EhHeader>

        '========================etiketes ========================================================
        Dim d1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        Dim mlabel As String

        Dim MPER As String

        Dim mtimes As Single

        Dim TIMM As Single

        Dim timologio As String

        'On Error Resume NextN  N

        Dim nn As Long

        Dim X As Printer

100     For Each X In Printers

110         If InStr(X.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
120             Set Printer = X

                ' Stop looking for a printer.
                Exit For

            End If

        Next

        On Error GoTo GATIDIS_ETIK_Err

130     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

        Dim r1 As New ADODB.Recordset

140     r1.Open "SELECT * FROM BARCODES WHERE KOD='" + DATA2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

150     If r1.EOF Then
160         GoSub FIND_N_NUMBER  ' EPISTREFEI TO mlabel
        Else

170         If Val(r1("ERG")) = 0 Then
180             GoSub FIND_N_NUMBER
            Else
190             mlabel = r1("ERG")
            End If
        End If

200     MPER = DATA2.Recordset("ono")
210     mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)

220     TIMM = Format(DATA2.Recordset("lti5"), "#####." + String(F_DEK_LIANIKIS, "0"))

        'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

230     For k = 1 To Val(mtimes)

240         Printer.FontSize = 6
250         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
260         Printer.FontBold = True

270         Printer.CurrentY = 12

            'Printer.Print Tab(0); MPER

            'If Len(DataGrid1.Text) = 12 Then
            ' BarCode1.SymbologyID = CODE128
            'Else
            '     BarCode1.SymbologyID = EAN13
            'End If
280         BarCode1.ShowText = NO
290         BarCode1.NarrowBarWidth = 0.025  '0.03
300         BarCode1.DataToEncode = DataGrid1.Text
310         Printer.PaintPicture BarCode1.Picture, F_XBARCODE, F_YBARCODE  ' '-88,124

            'Printer.FontName = "EAN8"
            'Printer.FontName = "EAN8"
            'Printer.FontSize = 8
320         Printer.CurrentY = 850
330         Printer.Print Tab(10); Text1(0).Text
340         Printer.Print Tab(10); mlabel
350         Printer.FontSize = 8
360         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
            'Printer.Print
            'Printer.Print
            'Printer.Print
            'Printer.Print

370         Printer.CurrentY = 900
            'Printer.Print Tab(5); Format(Val(TIMM), "###0.00") + " ΕΥΡΩ "
380         Printer.EndDoc

        Next

        Exit Sub

FIND_N_NUMBER:
        '************************
390     rec.MoveFirst
400     rec.Move 10

        'rec.Edit
410     If IsNull(rec("epel")) Then
420         rec("epel") = 1
        Else
430         rec("epel") = rec("epel") + 1
        End If

440     rec.Update
450     mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)

460     Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + DATA2.Recordset("KOD") + "','" + mlabel + "')", nn

470     If nn = 0 Then
480         MsgBox "ΔΕΝ ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ BARCODE"
        End If

        '  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Text1(0).Text + "'"
490     Return

        '<EhFooter>
        Exit Sub

GATIDIS_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.GATIDIS_ETIK " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.GATIDIS_ETIK " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub gazoros()

        '---------------------------------------------------------------
        '<EhHeader>
        On Error GoTo gazoros_Err

        '</EhHeader>
        Dim kod, per, MON, POS, TIMH

        Dim F As String, RECS As Integer

        Dim k, X

        Dim barc(200)

        Dim apoSCANNER

        Dim fores

        Dim a$

        'yupie
100     fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

110     If Scanner Then
            '
120         Open "C:\APOG.TXT" For Input As #1
130         X = 0
140         k = 0

150         Do While Not EOF(1)    'True

                On Error Resume Next

160             Line Input #1, a$
170             X = X + 1
180             kod = mID(a$, 1, 13)
190             POS = Val(mID$(a$, 15, 20))
200             barc(X) = kod
            Loop

210         Close #1

220         If X = 0 Then
230             MsgBox "δεν υπάρχουν εγγραφές"

                Exit Sub

            End If

240         apoSCANNER = X
        End If

        ' ΑΝ ΔΕΝ ΕΙΝΑΙ ΑΠΟ ΣΚΑΝΕΡ
250     If X = 0 Then
260         apoSCANNER = 1
270         barc(1) = Text1(6).Text
        End If

        Dim ll

        Dim mFORES As Long

        Dim R      As New ADODB.Recordset

        Dim mONOMA, mBarcode, mERG, mLTI5

280     For ll = 1 To apoSCANNER
            '
            'mFORES = fores
            '
            'If Scanner Then
290         R.Open "select * FROM EID WHERE ERG='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

300         If R.EOF Or Len(Trim(barc(ll))) = 0 Then
310             mFORES = 0
            Else
320             mONOMA = R("ONO")    'ONOMA
330             mBarcode = R("KOD")    '  ' BARCODE
340             mERG = R("ERG")    ' erg
350             mLTI5 = Format(R("LTI5"), "###0." + String(F_DEK_LIANIKIS, "0"))   ' LTI5
            End If

360         R.Close
            'End If

370         If Len(Text1(6).Text) < 13 Then
380             BarCode3.SymbologyID = CODE128    ' UPCA
            Else
390             BarCode3.SymbologyID = EAN13
            End If

400         For Each X In Printers

410             If InStr(X.DeviceName, "2-8-4-4") Or InStr(X.DeviceName, "SATO") > 0 Then
                    ' Set printer as system default.

420                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

430         BarCode3.DataToEncode = mERG    '  Text1(6).Text 'DataGrid1.Text

440         For k = 1 To fores

450             Printer.CurrentX = F_ArPerEtik * 100
460             Printer.CurrentY = 0
470             Printer.FontSize = 10
480             Printer.FontBold = False
490             Printer.FontName = "Arial Greek"

500             Printer.Print mONOMA    'Text1(1).Text 'perigrafh

510             Printer.CurrentX = F_ArPerEtik * 100

520             Printer.Print " "    ' Text1(6).Text ' barcode Text1(0).Text 'kodikos

                ' Printer.CurrentX = 2048
                'Printer.Print s
530             Printer.CurrentX = 400 + F_ArPerEtik * 100
                'Printer.FontSize = 12
540             Printer.FontSize = 10

550             Printer.CurrentX = 200 + F_ArPerEtik * 100
560             Printer.CurrentY = 50

570             Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 300  '500
                'Printer.CurrentY = 1200

580             Printer.CurrentY = 1000
                'Printer.Print MERG  ' Text1(6).Text ' barcode Text1(0).Text 'kodikos

590             Printer.EndDoc

            Next

        Next

        'If Scanner Then
        '
        ' ' μπουγουδης
        'Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
        'Set RD = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM= 1234")
        'If RD.RecordCount = 0 Then
        '   MsgBox "δεν υπάρχουν εγγραφές"
        '   Exit Sub
        'End If
        '
        '
        '  RD.MoveFirst
        '  k = 0
        '  Do While Not RD.EOF
        '     k = k + 1
        '     If k > 80 Then Exit Do
        '     If IsNull(RD("BARCODE")) Then
        '        barc(k) = " " 'Grid1.TextMatrix(K, f_k) = " "
        '     Else
        '        barc(k) = RD("barcode") 'Grid1.TextMatrix(K, f_k) = R("BARCODE")
        '     End If
        '     RD.MoveNext
        '  Loop
        '  RD.Close
        '  db.Close
        '  apoSCANNER = k
        'End If
        ''----------------------
        '
        'fores = Val(InputBox("Πόσες εττικέτες; ", "", 1))
        'Dim mFORES As Long
        'For ll = 1 To apoSCANNER
        '
        'mFORES = fores
        '
        'If Scanner Then
        '   R.Close
        '   R.Open "select * FROM EID WHERE KOD='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        '
        '  If R.EOF Or Len(Trim(barc(ll))) = 0 Then
        '     mFORES = 0
        '  Else
        '   mONOMA = R("ONO") 'ONOMA
        '   mBarcode = R("KOD") '  ' BARCODE
        '   mERG = R("ERG") ' erg
        '   mLTI5 = Format(R("LTI5"), "###0.00") ' LTI5
        '  End If
        'End If
        '

        '-----------ANTIGRAFO GAZOROY
        '   'yupie
        '    fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
        '
        '    'If Len(DataGrid1.Text) < 12 Then
        '    '   MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
        '    '   Exit Sub
        '    'End If
        '
        '    If Len(Text1(6).Text) < 13 Then
        '         BarCode3.SymbologyID = CODE128 ' UPCA
        '    Else
        '         BarCode3.SymbologyID = EAN13
        '    End If
        '
        '    For Each X In Printers
        '        If InStr(X.DeviceName, "2-8-4-4") Or InStr(X.DeviceName, "SATO") > 0 Then
        '          ' Set printer as system default.
        '
        '          Set Printer = X
        '          ' Stop looking for a printer
        '          Exit For
        '        End If
        '    Next
        '
        '    BarCode3.DataToEncode = Text1(6).Text 'DataGrid1.Text
        '
        '
        '  For k = 1 To fores
        '
        '    Printer.CurrentX = F_ArPerEtik * 100
        '    Printer.CurrentY = 0
        '    Printer.FONTSIZE = 10
        '     Printer.FontBold = False
        '    Printer.FontName = "Arial Greek"
        '
        '
        '    Printer.Print Text1(1).Text 'perigrafh
        '
        '    Printer.CurrentX = F_ArPerEtik * 100
        '
        '    Printer.Print " " ' Text1(6).Text ' barcode Text1(0).Text 'kodikos
        '
        '   ' Printer.CurrentX = 2048
        '    'Printer.Print s
        '    Printer.CurrentX = 400 + F_ArPerEtik * 100
        '    'Printer.FontSize = 12
        '    Printer.FONTSIZE = 10
        '
        '
        '    Printer.CurrentX = 200 + F_ArPerEtik * 100
        '        Printer.CurrentY = 50
        '
        '    Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 300  '500
        '    'Printer.CurrentY = 1200
        '
        '     Printer.CurrentY = 1000
        '   Printer.Print Text1(6).Text ' barcode Text1(0).Text 'kodikos
        '    'Printer.CurrentX = 1900 + F_ArPerEtik * 100
        '    'Printer.FontSize = 24
        '    'Printer.FontBold = True
        '    'Printer.Print Format(lti5.Text, "##0.00") + " €" ' timh
        '    ' Printer.Print "4.55 €" ' timh
        '    Printer.EndDoc
        '
        '  Next
        '

        '<EhFooter>
        Exit Sub

gazoros_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.gazoros " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.gazoros " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub gazorosTIMES()

        '---------------------------------------------------------------
        '<EhHeader>
        On Error GoTo gazorosTIMES_Err

        '</EhHeader>
        Dim kod, per, MON, POS, TIMH

        Dim F As String, RECS As Integer

        Dim k, X

        Dim barc(200)

        Dim apoSCANNER

        Dim fores

        Dim a$

        'yupie
100     fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

110     If Scanner Then
            '
120         Open "C:\APOG.TXT" For Input As #1
130         X = 0
140         k = 0

150         Do While Not EOF(1)    'True

                On Error Resume Next

160             Line Input #1, a$
170             X = X + 1
180             kod = mID(a$, 1, 13)
190             POS = Val(mID$(a$, 15, 20))
200             barc(X) = kod
            Loop

210         Close #1

220         If X = 0 Then
230             MsgBox "δεν υπάρχουν εγγραφές"

                Exit Sub

            End If

240         apoSCANNER = X
        End If

        ' ΑΝ ΔΕΝ ΕΙΝΑΙ ΑΠΟ ΣΚΑΝΕΡ
250     If X = 0 Then
260         apoSCANNER = 1
270         barc(1) = Text1(6).Text
        End If

        Dim ll

        Dim mFORES As Long

        Dim R      As New ADODB.Recordset

        Dim mONOMA, mBarcode, mERG, mLTI5

280     For ll = 1 To apoSCANNER
            '
            'mFORES = fores
            '
            'If Scanner Then
290         R.Open "select * FROM EID WHERE ERG='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

300         If R.EOF Or Len(Trim(barc(ll))) = 0 Then
310             mFORES = 0
            Else
320             mONOMA = R("ONO")    'ONOMA
330             mBarcode = R("KOD")    '  ' BARCODE
340             mERG = R("ERG")    ' erg
350             mLTI5 = Format(R("LTI5"), "###0." + String(F_DEK_LIANIKIS, "0"))   ' LTI5
            End If

360         R.Close
            'End If

370         If Len(Text1(6).Text) < 13 Then
380             BarCode3.SymbologyID = CODE128    ' UPCA
            Else
390             BarCode3.SymbologyID = EAN13
            End If

400         For Each X In Printers

410             If InStr(X.DeviceName, "2844") > 0 Then
                    ' Set printer as system default.

420                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

430         BarCode3.DataToEncode = mERG    '  Text1(6).Text 'DataGrid1.Text

            '   BarCode3.DataToEncode = Text1(6).Text 'DataGrid1.Text

440         For k = 1 To fores

450             Printer.CurrentX = F_ArPerEtik * 100
460             Printer.CurrentY = 200
470             Printer.FontSize = 12
480             Printer.FontBold = False
490             Printer.FontName = "Arial Greek"

500             Printer.Print mONOMA    ' Text1(1).Text 'perigrafh

510             Printer.CurrentX = F_ArPerEtik * 100

520             Printer.CurrentY = 0
530             Printer.CurrentX = 800
540             Printer.FontSize = 8
                ' Printer.Print Tab(10); MERG ' Text1(6).Text ' barcode    kodikos

                ' Printer.CurrentX = 2048
                'Printer.Print s
550             Printer.CurrentX = 400 + F_ArPerEtik * 100
                'Printer.FontSize = 12
560             Printer.FontSize = 10

570             Printer.CurrentX = 200 + F_ArPerEtik * 100
580             Printer.CurrentY = 100

590             Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 500    '1124
600             Printer.CurrentY = 1200

610             Printer.CurrentX = 1900 + F_ArPerEtik * 100
620             Printer.FontSize = 24
630             Printer.FontBold = True
640             Printer.Print mLTI5 + " €"    ' timh
                ' Printer.Print "4.55 €" ' timh
650             Printer.EndDoc

            Next

        Next

        'If Scanner Then
        '
        ' ' μπουγουδης
        'Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
        'Set RD = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM= 1234")
        'If RD.RecordCount = 0 Then
        '   MsgBox "δεν υπάρχουν εγγραφές"
        '   Exit Sub
        'End If
        '
        '
        '  RD.MoveFirst
        '  k = 0
        '  Do While Not RD.EOF
        '     k = k + 1
        '     If k > 80 Then Exit Do
        '     If IsNull(RD("BARCODE")) Then
        '        barc(k) = " " 'Grid1.TextMatrix(K, f_k) = " "
        '     Else
        '        barc(k) = RD("barcode") 'Grid1.TextMatrix(K, f_k) = R("BARCODE")
        '     End If
        '     RD.MoveNext
        '  Loop
        '  RD.Close
        '  db.Close
        '  apoSCANNER = k
        'End If
        ''----------------------
        '
        'fores = Val(InputBox("Πόσες εττικέτες; ", "", 1))
        'Dim mFORES As Long
        'For ll = 1 To apoSCANNER
        '
        'mFORES = fores
        '
        'If Scanner Then
        '   R.Close
        '   R.Open "select * FROM EID WHERE KOD='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        '
        '  If R.EOF Or Len(Trim(barc(ll))) = 0 Then
        '     mFORES = 0
        '  Else
        '   mONOMA = R("ONO") 'ONOMA
        '   mBarcode = R("KOD") '  ' BARCODE
        '   mERG = R("ERG") ' erg
        '   mLTI5 = Format(R("LTI5"), "###0.00") ' LTI5
        '  End If
        'End If
        '

        '-----------ANTIGRAFO GAZOROY
        '   'yupie
        '    fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
        '
        '    'If Len(DataGrid1.Text) < 12 Then
        '    '   MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
        '    '   Exit Sub
        '    'End If
        '
        '    If Len(Text1(6).Text) < 13 Then
        '         BarCode3.SymbologyID = CODE128 ' UPCA
        '    Else
        '         BarCode3.SymbologyID = EAN13
        '    End If
        '
        '    For Each X In Printers
        '        If InStr(X.DeviceName, "2-8-4-4") Or InStr(X.DeviceName, "SATO") > 0 Then
        '          ' Set printer as system default.
        '
        '          Set Printer = X
        '          ' Stop looking for a printer
        '          Exit For
        '        End If
        '    Next
        '
        '    BarCode3.DataToEncode = Text1(6).Text 'DataGrid1.Text
        '
        '
        '  For k = 1 To fores
        '
        '    Printer.CurrentX = F_ArPerEtik * 100
        '    Printer.CurrentY = 0
        '    Printer.FONTSIZE = 10
        '     Printer.FontBold = False
        '    Printer.FontName = "Arial Greek"
        '
        '
        '    Printer.Print Text1(1).Text 'perigrafh
        '
        '    Printer.CurrentX = F_ArPerEtik * 100
        '
        '    Printer.Print " " ' Text1(6).Text ' barcode Text1(0).Text 'kodikos
        '
        '   ' Printer.CurrentX = 2048
        '    'Printer.Print s
        '    Printer.CurrentX = 400 + F_ArPerEtik * 100
        '    'Printer.FontSize = 12
        '    Printer.FONTSIZE = 10
        '
        '
        '    Printer.CurrentX = 200 + F_ArPerEtik * 100
        '        Printer.CurrentY = 50
        '
        '    Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 300  '500
        '    'Printer.CurrentY = 1200
        '
        '     Printer.CurrentY = 1000
        '   Printer.Print Text1(6).Text ' barcode Text1(0).Text 'kodikos
        '    'Printer.CurrentX = 1900 + F_ArPerEtik * 100
        '    'Printer.FontSize = 24
        '    'Printer.FontBold = True
        '    'Printer.Print Format(lti5.Text, "##0.00") + " €" ' timh
        '    ' Printer.Print "4.55 €" ' timh
        '    Printer.EndDoc
        '
        '  Next
        '

        '<EhFooter>
        Exit Sub

gazorosTIMES_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.gazorosTIMES " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.gazorosTIMES " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command7_Click()

        '<EhHeader>
        On Error GoTo Command7_Click_Err

        '</EhHeader>

        Dim fores As Long, k As Integer

        Dim X     As Printer

        Dim DUM
        
         
         Dim F_ET2_PER_X, F_ET2_PER_Y
         Dim F_ET2_KOD_X, F_ET2_KOD_Y
         Dim F_ET2_BAR_X, F_ET2_BAR_Y
         Dim F_ET2_TIM_X, F_ET2_TIM_Y
       
       
       If Len(Dir("C:\ALL2ETIK.TXT", vbNormal)) > 0 Then
            Me.Caption = "C:\ALL2ETIK.TXT"

            'yupie
            fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

           ' Dim F_ET2_PER_X, F_ET2_PER_Y
  
            F_ET2_PER_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_PER_X", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΠΕΡΙΓΡΑΦΗΣ"))
            F_ET2_PER_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_PER_Y", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΠΕΡΙΓΡΑΦΗΣ"))
      
         

            F_ET2_KOD_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_KOD_X", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΚΩΔΙΚΟΥ"))
            F_ET2_KOD_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_KOD_Y", "200", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Υ ΚΩΔΙΚΟΥ"))

         

            F_ET2_BAR_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_BAR_X", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ BARCODE"))
            F_ET2_BAR_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_BAR_Y", "600", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y BARCODE"))

         

            F_ET2_TIM_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_TIM_X", "200", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ TIMΗΣ"))
            F_ET2_TIM_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_TIM_Y", "800", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y TIMΗΣ"))

            If Len(DataGrid1.Text) = 12 Then
                BarCode3.SymbologyID = UPCA
            ElseIf Len(DataGrid1.Text) = 13 Then
                BarCode3.SymbologyID = EAN13
            Else
                BarCode3.SymbologyID = CODE128
            End If

            For Each X In Printers

                If InStr(X.DeviceName, "SATO") > 0 Then
                    ' Set printer as system default.

                    Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

            For k = 1 To fores
        
                BarCode3.DataToEncode = DataGrid1.Text

                Printer.CurrentX = F_ArPerEtik * 100
                Printer.CurrentY = 200
                Printer.FontSize = 10
                Printer.FontBold = False
                Printer.FontName = "Arial Greek"
        
                Printer.CurrentX = F_ET2_PER_X  '  200 + F_ArPerEtik * 100
                Printer.CurrentY = F_ET2_PER_Y '600
                Printer.Print Text1(1).Text    'perigrafh

                Printer.CurrentX = F_ET2_KOD_X '  F_ArPerEtik * 100
                Printer.CurrentY = F_ET2_KOD_Y '600
                Printer.Print Text1(0).Text    'kodikos

                ' Printer.CurrentX = 2048
                'Printer.Print s
        
                Printer.CurrentX = 400 + F_ArPerEtik * 100
                'Printer.FontSize = 12
                Printer.FontSize = 10

                Printer.CurrentX = F_ET2_BAR_X  ' 200 + F_ArPerEtik * 100
                Printer.CurrentY = F_ET2_BAR_Y ' 600

                Printer.PaintPicture BarCode3.Picture, F_ET2_BAR_X, F_ET2_BAR_Y

                Printer.CurrentY = F_ET2_TIM_Y '1000 + F_ArPerEtik * 100
        
                Printer.FontSize = 18
                Printer.FontBold = True
        
                Printer.CurrentX = F_ET2_PER_X
                Printer.Print F_Epix;
        
                Printer.FontSize = 24
                Printer.CurrentX = F_ET2_TIM_X '1000 + F_ArPerEtik * 100
                Printer.Print Format(lti5.Text, "##0.00") + " €"    ' timh
                ' Printer.Print "4.55 €" ' timh
                Printer.EndDoc
            Next

            Exit Sub

        End If


'προσοχη για εκθεση διονυση   βγαζει χοντρικη  lti5=>"2014"+lti+"2015"  και μικρα γραμματα στην τιμη 24αρια=>10 αρια
        If Len(Dir("C:\ALL2xondETIK.TXT", vbNormal)) > 0 Then
            Me.Caption = "C:\ALL2ETIK.TXT"

            'yupie
            fores = Val(InputBox("(Xoνδρ) Πόσα .κομμάτια από το " + Text1(1).Text, , 1))

           
  
            F_ET2_PER_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_PER_X", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΠΕΡΙΓΡΑΦΗΣ"))
            F_ET2_PER_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_PER_Y", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΠΕΡΙΓΡΑΦΗΣ"))
      
            

            F_ET2_KOD_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_KOD_X", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΚΩΔΙΚΟΥ"))
            F_ET2_KOD_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_KOD_Y", "200", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Υ ΚΩΔΙΚΟΥ"))

       
            F_ET2_BAR_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_BAR_X", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ BARCODE"))
            F_ET2_BAR_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_BAR_Y", "600", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y BARCODE"))

       
            F_ET2_TIM_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_TIM_X", "200", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ TIMΗΣ"))
            F_ET2_TIM_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET2_TIM_Y", "800", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y TIMΗΣ"))

            If Len(DataGrid1.Text) = 12 Then
                BarCode3.SymbologyID = UPCA
            ElseIf Len(DataGrid1.Text) = 13 Then
                BarCode3.SymbologyID = EAN13
            Else
                BarCode3.SymbologyID = CODE128
            End If

            For Each X In Printers

                If InStr(X.DeviceName, "SATO") > 0 Then
                    ' Set printer as system default.

                    Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

            For k = 1 To fores
        
                BarCode3.DataToEncode = DataGrid1.Text

                Printer.CurrentX = F_ArPerEtik * 100
                Printer.CurrentY = 200
                Printer.FontSize = 10
                Printer.FontBold = False
                Printer.FontName = "Arial Greek"
        
                Printer.CurrentX = F_ET2_PER_X  '  200 + F_ArPerEtik * 100
                Printer.CurrentY = F_ET2_PER_Y '600
                Printer.Print Text1(1).Text    'perigrafh

                Printer.CurrentX = F_ET2_KOD_X '  F_ArPerEtik * 100
                Printer.CurrentY = F_ET2_KOD_Y '600
                Printer.Print Text1(0).Text    'kodikos

                ' Printer.CurrentX = 2048
                'Printer.Print s
        
                Printer.CurrentX = 400 + F_ArPerEtik * 100
                'Printer.FontSize = 12
                Printer.FontSize = 10

                Printer.CurrentX = F_ET2_BAR_X  ' 200 + F_ArPerEtik * 100
                Printer.CurrentY = F_ET2_BAR_Y ' 600

                Printer.PaintPicture BarCode3.Picture, F_ET2_BAR_X, F_ET2_BAR_Y

                Printer.CurrentY = F_ET2_TIM_Y '1000 + F_ArPerEtik * 100
        
                Printer.FontSize = 10
                Printer.FontBold = True
        
                Printer.CurrentX = F_ET2_PER_X
                Printer.Print F_Epix;
        
                Printer.FontSize = 10
                Printer.CurrentX = F_ET2_TIM_X '1000 + F_ArPerEtik * 100
                Printer.Print "2014" + LTrim(Format(lti.Text, "##0.00")) + "2015"  ' timh
                ' Printer.Print "4.55 €" ' timh
                Printer.EndDoc
            Next

            Exit Sub

        End If

        '========================================================================00002357

100     If Len(Dir("C:\GAZOROS2.EXE", vbNormal)) > 0 Then
110         gazoros

            Exit Sub

        End If

120     If Len(Dir("C:\PAYLOS.EXE", vbNormal)) > 0 Then
            Me.Caption = "C:\PAYLOS.EXE"
130         PAYLOS

            Exit Sub

        End If

        On Error Resume Next

140     If Len(Dir("C:\LAGEURO\LAGAKIS.TXT", vbNormal)) > 0 Then
150         LAG_ETIK

            Exit Sub

        End If

160     If Len(Dir("C:\LAGEURO\diesel.TXT", vbNormal)) > 0 Then
170         diesel_ETIK

            Exit Sub

        End If

180     If Len(Dir("C:\LAGEURO\LAGAKIS2.TXT", vbNormal)) > 0 Then
190         LAG2_ETIK

            Exit Sub

        End If

200     If Len(Dir("C:\LAGEURO\patrick.TXT", vbNormal)) > 0 Then
210         patrick_etik

            Exit Sub

        End If

220     If Len(Dir("C:\LAGEURO\SOKRATIS.TXT", vbNormal)) > 0 Then
230         SOKRATIS_etik

            Exit Sub

        End If

240     If Len(Dir("C:\LAGEURO\GATIDIS.TXT", vbNormal)) > 0 Then
250         GATIDIS_ETIK

            Exit Sub

        End If

260     If Len(Dir("C:\LAGEURO\MAKISGET.EXE", vbNormal)) > 0 Then
            'DUM = Shell("c:\lageuro\makisget.exe " + DATA2.Recordset("KOD") + " " + Replace(DATA2.Recordset("ONO"), " ", "_") + " " + Format(DATA2.Recordset("LTI5"), "##0.00") + " " + Format(DATA2.Recordset("POS_KERD"), "##0") + " " + DATA2.Recordset("ERG"))
270         MAKIS_ETIK

            Exit Sub

        End If

280     If Len(Dir("C:\DIONISIS.EXE", vbNormal)) > 0 Then

290         fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

            Dim PARAM

300         If Len(DataGrid1.Text) < 12 Then
310             MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"

                Exit Sub

            End If

320         If Len(DataGrid1.Text) = 12 Then
330             BarCode2.SymbologyID = UPCA
            Else
340             BarCode2.SymbologyID = EAN13
            End If

350         For Each X In Printers

360             If InStr(X.DeviceName, "2844") > 0 Then
                    ' Set printer as system default.

370                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

380         BarCode2.DataToEncode = DataGrid1.Text

390         For k = 1 To fores
400             Printer.CurrentY = 400
                ' Printer.CurrentX = 2048
                'Printer.Print s
410             Printer.CurrentX = 600 + F_ArPerEtik * 100
420             Printer.FontSize = 12

430             Printer.Print "*" + Text1(0).Text
440             Printer.PaintPicture BarCode2.Picture, F_XBARCODE, F_YBARCODE
450             Printer.EndDoc
            Next

            Exit Sub

        End If

460     For Each X In Printers

470         If InStr(X.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
480             Set Printer = X

                ' Stop looking for a printer
                Exit For

            End If

        Next

490     fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text))

500     If Len(Dir("C:\LAGEURO\MPOYGET.EXE", vbNormal)) > 0 Then
            'προγραμμα dos για εκτυπωτη epson tml-60
510         Shell ("C:\LAGEURO\MPOYGET.EXE " + Text1(0).Text + " " + str(fores) + " " + Replace(Text1(1).Text, " ", ".") + " " + "1")

            Exit Sub

        End If

        ' MPOYGET.EXE 5201219046154 2 PERIGRAFH 3.56

520     For k = 1 To fores

530         If Len(Dir("C:\mercvb\zebra.txt", vbNormal)) > 0 Then

                '    Printer.Print Tab(20); Text1(1).Text 'ONOMA
                '   Printer.Print Tab(30); Text1(6).Text ' mlabel  5200016000277

                'Printer.FontName = "Arial" ' "128"
540             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName454967325476-5200016000239

550             Printer.FontSize = 8

560             Printer.CurrentY = 0
570             Printer.CurrentX = 0
                'Printer.Print Left(Text1(6).Text, 12)

580             BarCode2.DataToEncode = Text1(0).Text
590             BarCode2.Code128CharSet = Set_A
600             BarCode2.SymbologyID = EAN13
610             Printer.PaintPicture BarCode2.Picture, 200, -500

                '    If Len(Trim(Text1(0).Text)) = 12 Then
                '       Printer.FontName = "UPCA" ' "128"
                '    Else
                '       Printer.FontName = "C128" '  "EAN13" ' "128"
                '    End If
                '   ' Printer.FontSize = 6
                '    Printer.CurrentY = 0
                '    Printer.CurrentX = 800
                '    Printer.Print Left(Text1(6).Text, 13) '  "123456789012" '  Text1(6).Text 'eggtim.Recordset("kode") ' mlabel

620             Printer.EndDoc

            Else
            
            
             If Len(Dir("C:\mercvb\loisir.txt", vbNormal)) > 0 Then
                 Printer.Print Tab(2); Text1(0).Text
                 
                 Printer.Print Tab(2); Format(gVal(lti5.Text), "###0.00") + "€" ' Text1(0).Text  ' kvdikow Text1(1).Text    'ONOMA
640             'Printer.Print Tab(30); Text1(6).Text    ' mlabel

'650             If Len(Trim(Text1(0).Text)) = 12 Then
'660                 Printer.FontName = "UPCA"    ' "128"
'                Else
'670                 Printer.FontName = "ean13"    ' "128"
'                End If
'
'680             Printer.FontSize = 6
'690             Printer.Print Tab(18); Text1(0).Text    'eggtim.Recordset("kode") ' mlabel
'700             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
'710             Printer.FontSize = 8
'720             Printer.Print
'730             Printer.Print
'740             Printer.Print
'750             Printer.Print
                Printer.EndDoc
              Else
            
                Printer.Print Tab(20); Text1(1).Text    'ONOMA
              Printer.Print Tab(30); Text1(6).Text    ' mlabel

650             If Len(Trim(Text1(0).Text)) = 12 Then
660                 Printer.FontName = "UPCA"    ' "128"
                Else
670                 Printer.FontName = "ean13"    ' "128"
                End If

680             Printer.FontSize = 6
690             Printer.Print Tab(18); Text1(0).Text    'eggtim.Recordset("kode") ' mlabel
700             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
710             Printer.FontSize = 8
720             Printer.Print
730             Printer.Print
740             Printer.Print
750             Printer.Print
760             Printer.EndDoc

               End If
            End If

        Next

        '<EhFooter>
        Exit Sub

Command7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command8_Click()

        '<EhHeader>
        On Error GoTo Command8_Click_Err

  
  
  
150         F_ET_PER_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET_PER_X", "0", "12. 1.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΠΕΡΙΓΡΑΦΗΣ"))
160         F_ET_PER_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET_PER_Y", "0", "12. 2.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y ΠΕΡΙΓΡΑΦΗΣ"))
      
           

170         F_ET_KOD_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET_KOD_X", "0", "12. 3.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΚΩΔΙΚΟΥ"))
180         F_ET_KOD_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET_KOD_Y", "200", "12. 4.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Υ ΚΩΔΙΚΟΥ"))


190         F_ET_BAR_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET_BAR_X", "0", "12. 5.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ BARCODE"))
200         F_ET_BAR_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET_BAR_Y", "600", "12. 6.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y BARCODE"))
            F_ET_BAR_H = Val(FINDPARAMETROI(1, "APOT2", "F_ET_BAR_H", "500", "12. 7.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  ΥΨΟΣ BARCODE"))


210         F_ET_TIM_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET_TIM_X", "200", "12. 8.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ TIMΗΣ"))
220         F_ET_TIM_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET_TIM_Y", "800", "12. 9.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y TIMΗΣ"))










        '</EhHeader>
        Dim nn As Long

        Dim ll, apoSCANNER As Long

        Dim DB

        Dim PSIFIA As Integer

        Dim fores  As Long, k As Integer

        Dim X      As Printer

        Dim barc(9000)

        Dim mONOMA, mBarcode, mERG, mLTI5

        Dim R  As New ADODB.Recordset

        Dim RD As DAO.Recordset

100     mONOMA = Text1(1).Text    'ONOMA
110     mBarcode = Text1(0).Text    '  ' BARCODE
120     mERG = Text1(6).Text    ' erg
130     mLTI5 = Format(Val(lti5.Text), "###0.00")    ' LTI5

        Dim PARAM

        Dim mFORES As Long

        Dim r1     As New ADODB.Recordset

        Dim C0     As String

        '   If Len(Dir("C:\FANAR.EXE", vbNormal)) > 0 Then
        '       fanar_etik
        '       Exit Sub
        '   End If

        'If Len(Dir("C:\FANAR.EXE", vbNormal)) > 0 Then
        '    'yupie
        '    'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
        '
        '    If Len(DataGrid1.Text) < 12 Then
        '       MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
        '       Exit Sub
        '    End If
        '
        '    If Len(DataGrid1.Text) = 12 Then
        '         BarCode3.SymbologyID = UPCA
        '    Else
        '         BarCode3.SymbologyID = EAN13
        '    End If
        '
        '    For Each X In Printers
        '        If InStr(X.DeviceName, "2844") > 0 Then
        '          ' Set printer as system default.
        '
        '          Set Printer = X
        '          ' Stop looking for a printer
        '          Exit For
        '        End If
        '    Next
        '
        '    BarCode3.DataToEncode = DataGrid1.Text
        '
        '
        '   'For K = 1 To fores
        '
        '    Printer.CurrentX = F_ArPerEtik * 100
        '    Printer.CurrentY = -150
        '    Printer.FONTSIZE = 10
        '     Printer.FontBold = True
        '    Printer.FontName = "COURIER" ' "Arial Greek"
        '
        '
        ' '   Printer.Print ""
        '    Printer.Print Tab(3); Left(UCase(Text1(1).Text), 24) 'perigrafh
        '    Printer.Print Tab(3); mID(UCase(Text1(1).Text), 25, 25) 'perigrafh
        '
        '
        '     Printer.CurrentY = 1600
        '    Printer.CurrentX = 2000
        '    Printer.FONTSIZE = 8
        '    Printer.Print Text1(0).Text 'kodikos
        '
        '    FGDB2.Execute "update Material  set retailprice=" + lti5.Text + "  WHERE  Code='" + Text1(0).Text + "'", nn
        '
        '
        '    If nn = 0 Then
        '       MsgBox "δεν δενημερωθηκε το singular"
        '    End If
        '
        '   ' Printer.CurrentX = 2048
        '    'Printer.Print s
        '    Printer.CurrentX = 400 + F_ArPerEtik * 100
        '    'Printer.FontSize = 12
        '    Printer.FONTSIZE = 10
        '
        '
        '    Printer.CurrentX = 200 + F_ArPerEtik * 100
        '        Printer.CurrentY = 700 ' 600
        '
        '    Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 20, 1200, , 500
        '
        '      Printer.CurrentY = 600 ' 600
        '    Printer.CurrentX = 1000 + F_ArPerEtik * 100
        '    Printer.FONTSIZE = 30
        '    Printer.FontBold = True
        '    Printer.Print Format(lti5.Text, "##0.00") + " €" ' timh
        '    ' Printer.Print "4.55 €" ' timh
        '    Printer.EndDoc
        '   'Next
        '
        '
        '   Exit Sub
        'End If

140     If Len(Dir("C:\ALLETIK.TXT", vbNormal)) > 0 Then
             Me.Caption = "C:\ALLETIK.TXT"
            'yupie
            'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

'            Dim F_ET_PER_X, F_ET_PER_Y
'
'150         F_ET_PER_X = Val(FindParametroi(1,"APOT2", "F_ET_PER_X", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΠΕΡΙΓΡΑΦΗΣ"))
'160         F_ET_PER_Y = Val(FindParametroi(1,"APOT2", "F_ET_PER_Y", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΠΕΡΙΓΡΑΦΗΣ"))
'
'            Dim F_ET_KOD_X, F_ET_KOD_Y
'
'170         F_ET_KOD_X = Val(FindParametroi(1,"APOT2", "F_ET_KOD_X", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΚΩΔΙΚΟΥ"))
'180         F_ET_KOD_Y = Val(FindParametroi(1,"APOT2", "F_ET_KOD_Y", "200", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Υ ΚΩΔΙΚΟΥ"))
'
'            Dim F_ET_BAR_X, F_ET_BAR_Y
'
'190         F_ET_BAR_X = Val(FindParametroi(1,"APOT2", "F_ET_BAR_X", "0", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ BARCODE"))
'200         F_ET_BAR_Y = Val(FindParametroi(1,"APOT2", "F_ET_BAR_Y", "600", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y BARCODE"))
'
'            Dim F_ET_TIM_X, F_ET_TIM_Y
'
'210         F_ET_TIM_X = Val(FindParametroi(1,"APOT2", "F_ET_TIM_X", "200", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ TIMΗΣ"))
'220         F_ET_TIM_Y = Val(FindParametroi(1,"APOT2", "F_ET_TIM_Y", "800", "ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y TIMΗΣ"))


               If Scanner.Value = vbChecked Then
                     polla_alletik
                     Exit Sub
                End If
                     











230         If Len(DataGrid1.Text) = 12 Then
240             BarCode3.SymbologyID = UPCA
250         ElseIf Len(DataGrid1.Text) = 13 Then
260             BarCode3.SymbologyID = EAN13
            Else
270             BarCode3.SymbologyID = CODE128
            End If

280         For Each X In Printers

290             If InStr(X.DeviceName, "SATO") > 0 Then
                    ' Set printer as system default.

300                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next
        
310         BarCode3.DataToEncode = DataGrid1.Text

320         Printer.CurrentX = F_ArPerEtik * 100
330         Printer.CurrentY = 200
340         Printer.FontSize = 10
350         Printer.FontBold = False
360         Printer.FontName = "Arial Greek"
        
370         Printer.CurrentX = F_ET_PER_X  '  200 + F_ArPerEtik * 100
380         Printer.CurrentY = F_ET_PER_Y '600
390         Printer.Print Text1(1).Text    'perigrafh

400         Printer.CurrentX = F_ET_KOD_X '  F_ArPerEtik * 100
410         Printer.CurrentY = F_ET_KOD_Y '600
420         Printer.Print Text1(0).Text    'kodikos

            ' Printer.CurrentX = 2048
            'Printer.Print s
        
430         Printer.CurrentX = 400 + F_ArPerEtik * 100
            'Printer.FontSize = 12
440         Printer.FontSize = 10

450         Printer.CurrentX = F_ET_BAR_X  ' 200 + F_ArPerEtik * 100
460         Printer.CurrentY = F_ET_BAR_Y ' 600

470         Printer.PaintPicture BarCode3.Picture, F_ET_BAR_X, F_ET_BAR_Y, , F_ET_BAR_H


            Printer.FontSize = F_ET_EPIX_FONTSIZE
            Printer.CurrentX = F_ET_EPIX_X
            Printer.CurrentY = F_ET_EPIX_Y
            Printer.Print F_Epix;




480         Printer.CurrentY = F_ET_TIM_Y '1000 + F_ArPerEtik * 100
        
490         Printer.FontSize = 18
500         Printer.FontBold = True
        
510         Printer.CurrentX = F_ET_PER_X

              
        
530         Printer.FontSize = 24
540         Printer.CurrentX = F_ET_TIM_X '1000 + F_ArPerEtik * 100
550         Printer.Print Format(lti5.Text, "##0.00") + " €"    ' timh
            ' Printer.Print "4.55 €" ' timh
552         Printer.EndDoc
            'Next

            Exit Sub

        End If

554     If Len(Dir("C:\DIONXONDR.EXE", vbNormal)) > 0 Then
                 Me.Caption = "C:\DIONXONDR.EXE"
            'yupie
            'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

556         If Len(DataGrid1.Text) < 12 Then
558             MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"

                Exit Sub

            End If


            Me.Caption = "DIONXONDR.EXE"


560         If Len(DataGrid1.Text) = 12 Then
562             BarCode3.SymbologyID = UPCA
            Else
564             BarCode3.SymbologyID = EAN13
            End If

566         For Each X In Printers

568             If InStr(X.DeviceName, "2844") > 0 Then
                    ' Set printer as system default.

570                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

572         BarCode3.DataToEncode = DataGrid1.Text

            'For K = 1 To fores

574       '  Printer.CurrentX = F_ArPerEtik * 100
576        ' Printer.CurrentY = 200
578         Printer.FontSize = 10
580         Printer.FontBold = False
582         Printer.FontName = "Arial Greek"

            
            Printer.CurrentX = F_ET_KOD_X
            Printer.CurrentY = F_ET_KOD_Y
            Printer.Print Text1(0).Text    'kodikos
            
            
            
            Printer.CurrentX = F_ET_PER_X
            Printer.CurrentY = F_ET_PER_Y
584         Printer.Print Text1(1).Text    'perigrafh

'586         Printer.CurrentX = F_ArPerEtik * 100

'588         Printer.Print Text1(0).Text    'kodikos

            ' Printer.CurrentX = 2048
            'Printer.Print s
'590         Printer.CurrentX = 400 + F_ArPerEtik * 100
            'Printer.FontSize = 12
'592         Printer.FontSize = 10

'594         Printer.CurrentX = 200 + F_ArPerEtik * 100
'596         Printer.CurrentY = 600

598         Printer.PaintPicture BarCode3.Picture, F_ET_BAR_X, F_ET_BAR_Y, , F_ET_BAR_H

602         Printer.FontSize = 14
604         Printer.FontBold = True
606        ' Printer.Print Tab(F_ArPerEtik); Format(lti.Text, "##0.00") + " €";    ' timh

            'Printer.FontSize = 14
            
            Printer.CurrentX = F_ET_TIM_X
            Printer.CurrentY = F_ET_TIM_Y
            
            Printer.Print Format(lti.Text, "##0.00") + " €"; " ΜΕ ΦΠΑ" + "(" + Format(XondrmeFPA.Text, "##0.00") + ") €" ' timh ' timh

            'Printer.FontSize = 14
            'Printer.Print Tab(15); "(" + Format(XondrmeFPA.Text, "##0.00") + ") €" ' timh
            ' Printer.Print "4.55 €" ' timh
608         Printer.EndDoc
            'Next

            Exit Sub

        End If

        If Len(Dir("C:\DION22.EXE", vbNormal)) > 0 Then
            'yupie
            'fores = Val(InputBo
On Error GoTo 0
            
            dion22
            Exit Sub
        End If
            


610     If Len(Dir("C:\DION2.EXE", vbNormal)) > 0 Then
            'yupie
            'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

612         If Len(DataGrid1.Text) < 12 Then
624             MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"

                Exit Sub

            End If

626         If Len(DataGrid1.Text) = 12 Then
628             BarCode3.SymbologyID = UPCA
            Else
629             BarCode3.SymbologyID = EAN13
            End If

630         For Each X In Printers

640             If InStr(X.DeviceName, "2844") > 0 Then
                    ' Set printer as system default.

650                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

660         BarCode3.DataToEncode = DataGrid1.Text

            'For K = 1 To fores

670         Printer.CurrentX = F_ArPerEtik * 100
680         Printer.CurrentY = 200
690         Printer.FontSize = 10
700         Printer.FontBold = False
710         Printer.FontName = "Arial Greek"

720         Printer.Print Text1(1).Text    'perigrafh

730         Printer.CurrentX = F_ArPerEtik * 100

740         Printer.Print Text1(0).Text    'kodikos

            ' Printer.CurrentX = 2048
            'Printer.Print s
750         Printer.CurrentX = 400 + F_ArPerEtik * 100
            'Printer.FontSize = 12
760         Printer.FontSize = 10

770         Printer.CurrentX = 200 + F_ArPerEtik * 100
780         Printer.CurrentY = 600

790         Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 1124

800         Printer.CurrentX = 1000 + F_ArPerEtik * 100
810         Printer.FontSize = 24
820         Printer.FontBold = True
830         Printer.Print Format(lti5.Text, "##0.00") + " €"    ' timh
            ' Printer.Print "4.55 €" ' timh
831         Printer.EndDoc
            'Next

            Exit Sub

        End If



' NEA ZEBRA   420T
     If Len(Dir("C:\DION3.EXE", vbNormal)) > 0 Then
     

            'yupie
            'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

832         If Len(DataGrid1.Text) < 12 Then
833             MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"

                Exit Sub

            End If

834         If Len(DataGrid1.Text) = 12 Then
835             BarCode3.SymbologyID = UPCA
            Else
             BarCode3.SymbologyID = EAN13
            End If

         For Each X In Printers

             If InStr(X.DeviceName, "2844") > 0 Then
                    ' Set printer as system default.

                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next
    ' Printer.Print "": GoTo 848
        
        
        BarCode3.DataToEncode = DataGrid1.Text

            'For K = 1 To fores

         Printer.CurrentX = F_ArPerEtik * 100
        Printer.CurrentY = 200
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.FontName = "Arial Greek"

         Printer.Print Text1(1).Text    'perigrafh

        Printer.CurrentX = F_ArPerEtik * 100

         Printer.Print Text1(0).Text    'kodikos



 'Printer.Print "": GoTo 848


            ' Printer.CurrentX = 2048
            'Printer.Print s
         Printer.CurrentX = 400 + F_ArPerEtik * 100
            'Printer.FontSize = 12
         Printer.FontSize = 10

' Printer.Print "": GoTo 848


         Printer.CurrentX = 200 + F_ArPerEtik * 100
        Printer.CurrentY = 800
        
    '     Printer.Print "": GoTo 848
       ' Printer.Font = "EAN 13"
       ' Printer.FontSize = 18
      '  Printer.Print DATAGRID1.Text
    '    Printer.PaintPicture BarCode3.Picture, 300, 824

  '  Printer.Print "": GoTo 848
Printer.FontName = "Zebra 1"

         Printer.CurrentX = 500  '1000 + F_ArPerEtik * 100
         Printer.FontSize = 24
         Printer.FontBold = True
         Printer.Print Format(lti5.Text, "##0.00") + " €"    ' timh
            
848        Printer.EndDoc
            'Next

            Exit Sub

        End If































        '========================================================================
850     If Len(Dir("C:\GAZOROS.EXE", vbNormal)) > 0 Then
860         gazorosTIMES

            Exit Sub

            'yupie
870         fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

            'If Len(DataGrid1.Text) < 12 Then
            '   MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
            '   Exit Sub
            'End If

880         If Len(Text1(6).Text) < 13 Then
890             BarCode3.SymbologyID = CODE128    ' UPCA
            Else
900             BarCode3.SymbologyID = EAN13
            End If

910         For Each X In Printers

920             If InStr(X.DeviceName, "2844") > 0 Then
                    ' Set printer as system default.

930                 Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

940         BarCode3.DataToEncode = Text1(6).Text    'DataGrid1.Text

950         For k = 1 To fores

960             Printer.CurrentX = F_ArPerEtik * 100
970             Printer.CurrentY = 200
980             Printer.FontSize = 12
990             Printer.FontBold = False
1000            Printer.FontName = "Arial Greek"

1010            Printer.Print Text1(1).Text    'perigrafh

1020            Printer.CurrentX = F_ArPerEtik * 100

1030            Printer.CurrentY = 450
1040            Printer.CurrentX = 200
1050            Printer.FontSize = 8
1060            Printer.Print Tab(10); Text1(6).Text    ' barcode    kodikos

                ' Printer.CurrentX = 2048
                'Printer.Print s
1070            Printer.CurrentX = 400 + F_ArPerEtik * 100
                'Printer.FontSize = 12
1080            Printer.FontSize = 10

1090            Printer.CurrentX = 200 + F_ArPerEtik * 100
1100            Printer.CurrentY = 100

1110            Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 500    '1124
1120            Printer.CurrentY = 1200

1130            Printer.CurrentX = 1900 + F_ArPerEtik * 100
1140            Printer.FontSize = 24
1150            Printer.FontBold = True
1160            Printer.Print Format(lti5.Text, "##0.00") + " €"    ' timh
                ' Printer.Print "4.55 €" ' timh
1170            Printer.EndDoc

            Next

            Exit Sub

        End If

1180    If Len(Dir("C:\DION3.EXE", vbNormal)) > 0 Then
            'yupie
            'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

1190        If Len(DataGrid1.Text) < 12 Then
1200            MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"

                Exit Sub

            End If

1210        If Len(DataGrid1.Text) = 12 Then
1220            BarCode3.SymbologyID = UPCA
            Else
1230            BarCode3.SymbologyID = EAN13
            End If

1240        For Each X In Printers

1250            If InStr(X.DeviceName, "2844") > 0 Then
                    ' Set printer as system default.

1260                Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

1270        BarCode3.DataToEncode = DataGrid1.Text

            'For K = 1 To fores

1280        Printer.CurrentX = F_ArPerEtik * 100

1290        Printer.FontSize = 10
1300        Printer.FontBold = False
1310        Printer.FontName = "Arial Greek"

            'Printer.CurrentY = 1100
            'Printer.Print DataGrid1.Text

1320        Printer.CurrentY = 200

1330        Printer.Print Text1(1).Text    'perigrafh

1340        Printer.CurrentX = F_ArPerEtik * 100

1350        Printer.Print Text1(0).Text    'kodikos

            ' Printer.CurrentX = 2048
            'Printer.Print s
1360        Printer.CurrentX = 400 + F_ArPerEtik * 100
            'Printer.FontSize = 12
1370        Printer.FontSize = 10

1380        Printer.CurrentX = 200 + F_ArPerEtik * 100
1390        Printer.CurrentY = 800

1400        Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 1224

1410        Printer.CurrentY = 600
1420        Printer.CurrentX = 1000 + F_ArPerEtik * 100
1430        Printer.FontSize = 24
1440        Printer.FontBold = True
1450        Printer.Print Format(lti5.Text, "##0.00") + " €"    ' timh
            ' Printer.Print "4.55 €" ' timh
1460        Printer.EndDoc
            'Next

            Exit Sub

        End If

        '--------------------------------------------------------------------------------
1470    If Len(Dir("C:\DIONXONDR.TXT", vbNormal)) > 0 Then

            'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

1480        If Len(DataGrid1.Text) < 12 Then
1490            MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"

                Exit Sub

            End If

1500        If Len(DataGrid1.Text) = 12 Then
1510            BarCode2.SymbologyID = UPCA
            Else
1520            BarCode2.SymbologyID = EAN13
            End If

1530        For Each X In Printers

1540            If InStr(X.DeviceName, "2844") > 0 Then
                    ' Set printer as system default.

1550                Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

1560        BarCode2.DataToEncode = DataGrid1.Text

            'For K = 1 To fores

1570        Printer.FontBold = False
1580        Printer.FontName = Text1(1).FontName  '  "Courier New Greek"

1590        Printer.CurrentX = F_2ArPerEtik * 100
1600        Printer.CurrentY = 200
1610        Printer.FontSize = 10
1620        Printer.Print Text1(1).Text    'perigrafh

1630        Printer.CurrentX = F_2ArPerEtik * 100

1640        Printer.Print Text1(0).Text    'kodikos

            ' Printer.CurrentX = 2048
            'Printer.Print s
1650        Printer.CurrentX = 400 + F_2ArPerEtik * 100
            'Printer.FontSize = 12
1660        Printer.FontSize = 10

1670        Printer.CurrentX = 200 + F_2ArPerEtik * 100
1680        Printer.CurrentY = 600

1690        Printer.PaintPicture BarCode2.Picture, F_2ArPerEtik * 100, 1124

1700        Printer.CurrentX = 1000 + F_2ArPerEtik * 100
1710        Printer.FontSize = 16
1720        Printer.FontBold = True
1730        Printer.FontName = "Courier New"

1740        Printer.Print Format(lti.Text, "##0.00") + " € (" + Format((100 + g_Fpa(Combo2.Text)) / 100 * lti.Text, "##0.00") + ")"    ' timh

            ' Printer.Print "4.55 €" ' timh
1750        Printer.EndDoc
            'Next

            Exit Sub

        End If

        '-----------------------------------------------------------If Len(Dir("C:\DIONISIS.EXE", vbNormal)) > 0 Then

        '--------------------------------------------------------------------------------
1760    If Len(Dir("C:\DIONISIS.EXE", vbNormal)) > 0 Then

            'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))

1770        If Len(DataGrid1.Text) < 12 Then
1780            MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"

                Exit Sub

            End If

1790        If Len(DataGrid1.Text) = 12 Then
1800            BarCode1.SymbologyID = UPCA
            Else
1810            BarCode1.SymbologyID = EAN13
            End If

1820        For Each X In Printers

1830            If InStr(X.DeviceName, "2844") > 0 Then
                    ' Set printer as system default.

1840                Set Printer = X

                    ' Stop looking for a printer
                    Exit For

                End If

            Next

1850        BarCode1.DataToEncode = DataGrid1.Text

            'For K = 1 To fores

1860        Printer.FontBold = False
1870        Printer.FontName = "Arial"

1880        Printer.CurrentX = F_ArPerEtik * 100
1890        Printer.CurrentY = 200
1900        Printer.FontSize = 10
1910        Printer.Print Text1(1).Text    'perigrafh

1920        Printer.CurrentX = F_ArPerEtik * 100

1930        Printer.Print Text1(0).Text    'kodikos

            ' Printer.CurrentX = 2048
            'Printer.Print s
1940        Printer.CurrentX = 400 + F_ArPerEtik * 100
            'Printer.FontSize = 12
1950        Printer.FontSize = 10

1960        Printer.CurrentX = 200 + F_ArPerEtik * 100
1970        Printer.CurrentY = 600

1980        Printer.PaintPicture BarCode1.Picture, F_ArPerEtik * 100, 1124

1990        Printer.CurrentX = 1000 + F_ArPerEtik * 100
2000        Printer.FontSize = 24
2010        Printer.FontBold = True
2020        Printer.Print Format(lti5.Text, "##0.00") + " €"    ' timh
            ' Printer.Print "4.55 €" ' timh
2030        Printer.EndDoc
            'Next

            Exit Sub

        End If

        '-----------------------------------------------------------If Len(Dir("C:\DIONISIS.EXE", vbNormal)) > 0 Then

2040    If Len(Dir("C:\MERCVB\ILIADIS2.EXE", vbNormal)) > 0 Then

            On Error Resume Next

2050        For Each X In Printers
                'f_printer
2060            If InStr(X.DeviceName, f_printer) > 0 Then
                    ' Set printer as system default.
2070                Set Printer = X

                    ' Stop looking for a printer  5200016000239=== 5200016000192
                    Exit For

                End If

            Next

2080        apoSCANNER = 1

2090        If Scanner Then

                ' μπουγουδης
2100            Set DB = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
2110            Set RD = DB.OpenRecordset("SELECT * FROM TIMAGOR ")

2120            If RD.RecordCount = 0 Then
2130                MsgBox "δεν υπάρχουν εγγραφές"

                    Exit Sub

                End If

2140            RD.MoveFirst
2150            k = 0

2160            Do While Not RD.EOF
2170                k = k + 1

2180                'If k > 2000 Then Exit Do
2190                If IsNull(RD("BARCODE")) Then
2200                    barc(k) = " "    'Grid1.TextMatrix(K, f_k) = " "
                    Else
2210                    barc(k) = RD("barcode")    'Grid1.TextMatrix(K, f_k) = R("BARCODE")
                    End If

2220                RD.MoveNext
                Loop

2230            RD.Close
2240            DB.Close
2250            apoSCANNER = k
            End If

            '----------------------
2260        fores = 1

2270        fores = Val(InputBox("Πόσες εττικέτες; ", "", 1))

2280        For ll = 1 To apoSCANNER

2290            mFORES = fores

2300            If Scanner Then
2310                R.Close
                    ' R.Open "SELECT * FROM BARCODES WHERE ERG = '" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                    'C0 = R("KOD")
                    'R.Close

2320                R.Open "select * FROM EID WHERE KOD LIKE '%" + barc(ll) + "%'", Gdb, adOpenForwardOnly, adLockReadOnly

2330                If R.EOF Or Len(Trim(barc(ll))) = 0 Then
2340                    mFORES = 0
                    Else
                        'το r1 xrhsimopoihuhke για να τυπωνει to barcode 15/02/2010
                        'R1.Open "SELECT top 1 * FROM BARCODES WHERE KOD='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
2350                    mBarcode = R("ERG")    '  R("KOD") '  ' BARCO
                        'R1.Close
2360                    mONOMA = R("ONO")    'ONOMA
                        'DE
2370                    Text1(6).Text = R("ERG")

2380                    mERG = R("ERG")    ' erg
2390                    mLTI5 = Format(R("LTI5"), "###0." + String(F_DEK_LIANIKIS, "0")) ' LTI5
                    End If

                Else
2400                r1.Open "SELECT top 1 ERG FROM BARCODES WHERE KOD='" + Text1(0).Text + "'", Gdb, adOpenForwardOnly, adLockReadOnly
2410                mBarcode = r1("ERG")
2420                r1.Close

                End If

                Dim Z As Integer

2430            Z = F_ArPerEtik    ' InputBox("METATOPISH")

2440            For k = 1 To mFORES

2450                If Len(Dir("C:\CMPOYGZEB.EXE", vbNormal)) > 0 Then

                    Else

                        'GoSub SATO
                    End If  ' 078257586417

2460                GoSub ZEBRAILIADIS

2470            Next k

2480        Next ll

            Exit Sub

        End If    ''''''''''''''''''''''''''''''''''''''''' Len(Dir("C:\ILIADIS.EXE", vbNormal)) > 0 Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''' ILIADIS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
        
        
        On Error Resume Next

2490    For Each X In Printers

2500        If InStr(X.DeviceName, "SATONEW") > 0 Then
                ' Set printer as system default.
2510            Set Printer = X

                ' Stop looking for a printer  5200016000239=== 5200016000192
                Exit For

            End If

        Next

2520    apoSCANNER = 1

2530    If Scanner Then
               

           If Len(Dir("C:\MERCVB\EGGTIM2.TXT", vbNormal)) > 0 Then
             Open "c:\mercvb\eggtim2.txt" For Input As #6

                Dim rSQL As New ADODB.Recordset
                Dim rSQL2 As New ADODB.Recordset
                Dim rSQL3 As New ADODB.Recordset
                Dim RR() As String, a As String
                Dim M_CODE As String

                k = 0
                Do While True ' Not EOF(6)
                   k = k + 1
                   Line Input #6, a
                   RR = Split(a, ";")
                   barc(k) = RR(7)
                   If EOF(6) Then
                      Exit Do
                   End If
                 Loop
              Close #6
              apoSCANNER = k

           Else
            ' μπουγουδης
2540        Set DB = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
2550        Set RD = DB.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM= 1234")

2560        If RD.RecordCount = 0 Then
2570            MsgBox "δεν υπάρχουν εγγραφές"

                Exit Sub

            End If

2580        RD.MoveFirst
2590        k = 0

2600        Do While Not RD.EOF
2610            k = k + 1

2620            If k > 80 Then Exit Do
2630            If IsNull(RD("BARCODE")) Then
2640                barc(k) = " "    'Grid1.TextMatrix(K, f_k) = " "
                Else
2650                barc(k) = RD("barcode")    'Grid1.TextMatrix(K, f_k) = R("BARCODE")
                End If

2660            RD.MoveNext
            Loop

2670        RD.Close
2680        DB.Close
2690        apoSCANNER = k
        
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        End If

        '----------------------
2700    fores = 1

2710    fores = Val(InputBox("Πόσες εττικέτες; ", "", 1))

2720    For ll = 1 To apoSCANNER

2730        mFORES = fores

2740        If Scanner Then
2750            R.Close
2760            R.Open "SELECT * FROM BARCODES WHERE ERG='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
2770            C0 = R("KOD")
2780            R.Close

2790            R.Open "select * FROM EID WHERE KOD='" + C0 + "'", Gdb, adOpenForwardOnly, adLockReadOnly

2800            If R.EOF Or Len(Trim(barc(ll))) = 0 Then
2810                mFORES = 0
                Else
                    'το r1 xrhsimopoihuhke για να τυπωνει to barcode 15/02/2010
2820                r1.Open "SELECT top 1 * FROM BARCODES WHERE KOD='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
2830                mBarcode = r1("ERG")    '  R("KOD") '  ' BARCO
2840                r1.Close
2850                mONOMA = R("ONO")    'ONOMA
                    'DE
2860                Text1(6).Text = R("ERG")

2870                mERG = R("ERG")    ' erg
2880                mLTI5 = Format(R("LTI5"), "###0." + String(F_DEK_LIANIKIS, "0"))   ' LTI5
                End If

            Else
2890            r1.Open "SELECT top 1 ERG FROM BARCODES WHERE KOD='" + Text1(0).Text + "'", Gdb, adOpenForwardOnly, adLockReadOnly
2900            mBarcode = r1("ERG")
2910            r1.Close

                mONOMA = R("ONO")    'ONOMA
                    'DE
                Text1(6).Text = R("ERG")
                mERG = R("ERG")    ' erg
                mLTI5 = Format(R("LTI5"), "###0." + String(F_DEK_LIANIKIS, "0"))   ' LTI5







            End If

2920        Z = F_ArPerEtik    ' InputBox("METATOPISH")

2930        For k = 1 To mFORES

2940            If Len(Dir("C:\CMPOYGZEB.EXE", vbNormal)) > 0 Then
2950                GoSub ZEBRA
                Else
2960                GoSub SATO
                End If  ' 078257586417

2970        Next k

2980    Next ll

        Exit Sub

        '------------------- ZEBRA-------------------------------------
ZEBRA:
2990    Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992  123456789012

        'Printer.CurrentY = 0

3000    Printer.FontSize = 8
3010    Printer.FontBold = True
3020    Printer.Print Tab(4); mONOMA    'Text1(1).Text 'ONOMA
3030    Printer.FontSize = 8
3040    Printer.Print Tab(4); mBarcode    '  Text1(0).Text;  ' BARCODE
3050    Printer.CurrentY = 350
        'boithitikos kodikos
3060    Printer.Print Tab(4); Text1(6).Text    ' mlabel
3070    Printer.Print
3080    Printer.CurrentY = 200

3090    If Len(DataGrid1.Text) = 12 Then
3100        BarCode4.SymbologyID = UPCA
        Else
3110        BarCode4.SymbologyID = EAN13
        End If

        '    BarCode4.ShowText = NO
3120    BarCode4.DataToEncode = mBarcode

3130    Printer.PaintPicture BarCode4.Picture, 1900, 150, , 400    ' PICTURE ,X,Y,PLATOS,YPSOS  500ΥΨΟΣ

        'If Len(Trim(Text1(0).Text)) = 12 Then
        '   Printer.FontName = "UPCA" ' "128"
        'Else
        '   Printer.FontName = "EAN13SMALL" ' "128"
        'End If
        'Printer.Print Tab(14 + Z); mBarcode '  Text1(0).Text;  ' BARCODE

3140    If F_PONTOI = 1 Then
3150        Printer.FontSize = 5
            '   Printer.Print Tab(15 + Z); mBarcode 'Text1(0).Text 'eggtim.Recordset("kode") ' mlabel
        End If

3160    Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
3170    Printer.FontBold = True
3180    Printer.FontSize = 8
3190    Printer.CurrentY = 570
3200    Printer.FontSize = 24

3210    If Len(mLTI5) = 6 Then    ' 208.35
3220        Printer.Print Tab(1 + 7); "" + mLTI5; " €"   '-7 EIXE STIS 17-8-2006
3230    ElseIf Len(mLTI5) = 5 Then    '18.35
3240        Printer.Print Tab(2 + 7); "" + mLTI5; " €"
        Else
3250        Printer.Print Tab(2 + 8); "" + mLTI5; " €"
        End If

3260    Printer.FontSize = 6
3270    Printer.Print
3280    Printer.FontSize = 8
3290    Printer.Print

3300    If F_PONTOI = 0 Then    'Len(Dir("C:\LAGEURO\NOPONTOI.TXT", vbNormal)) > 0 Then
        Else
             
            ' ΜΠΟΥΓΟΥΔΗΣ
3310        Printer.CurrentX = F_XPONTOI ' 200
3320        Printer.CurrentY = 550
3330        Printer.FontSize = 14

3340        If Val(mLTI5) > 2.99 Then
3350            Printer.Print Tab(F_YPONTOI); "ΠΟΝΤΟΙ"
3360            PSIFIA = Len(Format(Int(Val(mLTI5) / 3), "##0"))

3370            If PSIFIA = 3 Then Printer.Print ; Tab(5 + Z - 7); Format(Int(Val(mLTI5) / 3), "##0")
3380            If PSIFIA = 1 Then Printer.Print ; Tab(6 + Z - 6); Format(Int(Val(mLTI5) / 3), "##0")
3390            If PSIFIA = 2 Then Printer.Print ; Tab(5 + Z - 6); Format(Int(Val(mLTI5) / 3), "##0")
            End If
        End If

3400    Printer.EndDoc
3410    Return

        '--------------------------------------------SATO ----------------------
SATO:
3420    Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992
3430    Printer.FontSize = 8
3440    Printer.FontBold = True
3450    Printer.Print Tab(4 + Z); mONOMA    'Text1(1).Text 'ONOMA
3460    Printer.FontSize = 8
3470    Printer.Print Tab(4 + Z); mBarcode    '  Text1(0).Text;  ' BARCODE
3480    Printer.CurrentY = 350
        'boithitikos kodikos
3490    Printer.Print Tab(4 + Z); Text1(6).Text    ' mlabel
3500    Printer.Print
3510    Printer.CurrentY = 200

3520    If Len(mBarcode) = 12 Then    ' If Len(Trim(Text1(0).Text)) = 12 Then
3530        Printer.FontName = "UPCA"    ' "128"
        Else
3540        Printer.FontName = "EAN13SMALL"    ' "128"
        End If

3550    Printer.Print Tab(14 + Z); mBarcode    '  Text1(0).Text;  ' BARCODE

3560    If F_PONTOI = 1 Then
3570        Printer.FontSize = 5
            '   Printer.Print Tab(15 + Z); mBarcode 'Text1(0).Text 'eggtim.Recordset("kode") ' mlabel
        End If

3580    Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
3590    Printer.FontBold = True
3600    Printer.FontSize = 8
3610    Printer.CurrentY = 500
3620    Printer.FontSize = 24

3630    If Len(mLTI5) = 6 Then    ' 208.35
3640        Printer.Print Tab(1 + Z); "" + mLTI5; " €"   '-7 EIXE STIS 17-8-2006
3650    ElseIf Len(mLTI5) = 5 Then    '18.35
3660        Printer.Print Tab(2 + Z); "" + mLTI5; " €"
        Else
3670        Printer.Print Tab(2 + Z); "" + mLTI5; " €"
        End If

3680    Printer.FontSize = 6
3690    Printer.Print
3700    Printer.FontSize = 8
3710    Printer.Print

3720    If F_PONTOI = 0 Then    'Len(Dir("C:\LAGEURO\NOPONTOI.TXT", vbNormal)) > 0 Then
        Else
            ' ΜΠΟΥΓΟΥΔΗΣ
3730        Printer.CurrentX = 200
3740        Printer.CurrentY = 500
3750        Printer.FontSize = 14

3760        If Val(mLTI5) > 2.99 Then
3770            Printer.Print Tab(Z - 2); "ΠΟΝΤΟΙ"
3780            PSIFIA = Len(Format(Int(Val(mLTI5) / 3), "##0"))

3790            If PSIFIA = 3 Then Printer.Print ; Tab(5 + Z - 5); Format(Int(Val(mLTI5) / 3), "##0")
3800            If PSIFIA = 1 Then Printer.Print ; Tab(6 + Z - 4); Format(Int(Val(mLTI5) / 3), "##0")
3810            If PSIFIA = 2 Then Printer.Print ; Tab(5 + Z - 4); Format(Int(Val(mLTI5) / 3), "##0")
            End If
        End If

3820    Printer.EndDoc
3830    Return

        '------------------- ZEBRA-------------------------------------
ZEBRAILIADIS:
3840    Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992  123456789012

        'Printer.CurrentY = 0

3850    Printer.FontSize = 8
3860    Printer.FontBold = True
3870    Printer.Print Tab(4); mONOMA    'Text1(1).Text 'ONOMA
3880    Printer.FontSize = 8
3890    Printer.Print Tab(1); mBarcode    '  Text1(0).Text;  ' BARCODE
3900    Printer.CurrentY = 350
        'boithitikos kodikos
3910    Printer.Print Tab(1); Text1(6).Text    ' mlabel
3920    Printer.Print
3930    Printer.CurrentY = 200

3940    If Len(DataGrid1.Text) = 12 Then
3950        BarCode4.SymbologyID = UPCA
        Else
3960        BarCode4.SymbologyID = EAN13
        End If

        '    BarCode4.ShowText = NO
3970    BarCode4.DataToEncode = mBarcode

3980    Printer.PaintPicture BarCode4.Picture, 1200, 150, , 400    ' PICTURE ,X,Y,PLATOS,YPSOS  500ΥΨΟΣ

        'If Len(Trim(Text1(0).Text)) = 12 Then
        '   Printer.FontName = "UPCA" ' "128"
        'Else
        '   Printer.FontName = "EAN13SMALL" ' "128"
        'End If
        'Printer.Print Tab(14 + Z); mBarcode '  Text1(0).Text;  ' BARCODE

3990    If F_PONTOI = 1 Then
4000        Printer.FontSize = 5
            '   Printer.Print Tab(15 + Z); mBarcode 'Text1(0).Text 'eggtim.Recordset("kode") ' mlabel
        End If

4010    Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
4020    Printer.FontBold = True
4030    Printer.FontSize = 8
4040    Printer.CurrentY = 570
4050    Printer.FontSize = 24

4060    If Len(mLTI5) = 6 Then    ' 208.35
4070        Printer.Print Tab(1 + 3); "" + mLTI5; " €"   '-7 EIXE STIS 17-8-2006
4080    ElseIf Len(mLTI5) = 5 Then    '18.35
4090        Printer.Print Tab(2 + 3); "" + mLTI5; " €"
        Else
4100        Printer.Print Tab(2 + 4); "" + mLTI5; " €"
        End If

4110    Printer.FontSize = 6
        'Printer.Print
4120    Printer.FontSize = 8
        'Printer.Print

4130    Printer.EndDoc
4140    Return

        '<EhFooter>
        Exit Sub

Command8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command9_Click()

        '<EhHeader>
        On Error GoTo Command9_Click_Err

        '</EhHeader>

100     data22.RecordSource = UCase(" SELECT * FROM EID WHERE MEMO LIKE '%" + Text2.Text + "%';")
110     data22.Refresh
120     DBGrid1.SetFocus

        '   SELECT * FROM EID WHERE MEMO LIKE '%DOKIM%'
        '<EhFooter>
        Exit Sub

Command9_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Command9_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Command9_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        ' On Error Resume Next
        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        '</EhHeader>

100     Text3.Text = DataGrid1.Text

        '<EhFooter>
        Exit Sub

DataGrid1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.DataGrid1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.DataGrid1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DBGrid1_Click()
     mDBGrid1_RowColChange
    ' MILSEC 3000
    ' mDBGrid1_RowColChange
End Sub

Private Sub DBGRID1_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo DBGRID1_KeyDown_Err

        '</EhHeader>

100     mDBGrid1_RowColChange

        '<EhFooter>
        Exit Sub

DBGRID1_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.DBGRID1_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.DBGRID1_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DBGrid1_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = 13 Then

110         KeyAscii = 0
120         mDBGrid1_RowColChange
130         Command2.SetFocus
        End If

        '<EhFooter>
        Exit Sub

DBGrid1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.DBGrid1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.DBGrid1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo DBGrid1_KeyUp_Err

        '</EhHeader>
100     If KeyCode = 38 Or KeyCode = 40 Then    'ano=38 kai kato=40 belos
110         mDBGrid1_RowColChange
        End If

        '<EhFooter>
        Exit Sub

DBGrid1_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.DBGrid1_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.DBGrid1_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mDBGrid1_RowColChange()    '(LastRow As Variant, ByVal LastCol As Integer)

        '<EhHeader>
        On Error GoTo mDBGrid1_RowColChange_Err

        '</EhHeader>
        Dim G, m, a, GA, MA, p, k, SM, SG, SM2, SG2

        Dim mPos01 As Single, mPos02 As Single

        Dim mPos03 As Single

        Dim mPos04 As Single

        On Error Resume Next

        Dim SQL2 As String
        
        DBGrid1.Refresh
        Dim mmkod As String
        
mmkod = DBGrid1.columns(0).Text

100    ' DBGrid1.BackColorSel = vbBlue
        ' apot2.MousePointer = vbHourglass
110     DATA2.RecordSource = " select * from EID WHERE KOD='" + mmkod + "'"
120     DATA2.Refresh

        If DATA2.Recordset.EOF Then

            Exit Sub

        End If


        If DATA2.Recordset("energo") = 1 Then chkEnergos.Value = vbChecked Else chkEnergos.Value = vbUnchecked


        Set msEditor1.mDataSource = DATA2   ' for code you have to use the 'adoNotes' instead of the Adodc1
        msEditor1.MaxLength = 8000          ' Optional maximum length.
        '
        msEditor1.mDataField = "MEMO" '  "nNotes"     ' The field of your text
        msEditor1.Enabled = True

        '  apot2.Caption = Now
130     Text1(0).Text = DATA2.Recordset("kod")
140     Text1(1).Text = DATA2.Recordset("ONO")

150     Text1(16).Text = DATA2.Recordset("koderg")

160     Text1(4).Text = IIf(IsNull(DATA2.Recordset("pos")), " ", DATA2.Recordset("pos"))

170     Combo2.Text = DATA2.Recordset("FPA")
        
        Dim N As Integer
        If IsNull(DATA2.Recordset("discountOption")) Then
            discountOption.Text = discountOption.List(0)
         ElseIf DATA2.Recordset("discountOption") = 0 Then
            discountOption.Text = discountOption.List(0)
        Else
            discountOption.Text = discountOption.List(0)
         End If
         
         
        Combo5.ListIndex = -1
180     For N = 0 To Combo5.ListCount - 1
190         If Val(Left(Combo5.List(N), 3)) = Val(nNull(DATA2.Recordset("kathgoria"))) Then
200             Combo5.Text = Combo5.List(N)    'OIKOGENEIA
                Exit For
            End If
        Next

        'Combo5.Text = Combo5.List(Val(DATA2.Recordset("KATHGORIA")) - 1)

210     Combo1.Text = DATA2.Recordset("mon")

        ' Combo4.Text = Combo4.List(Val(Data2.Recordset("AEG")) - 1) 'OIKOGENEIA
        'Dim n As Integer


        Combo4.ListIndex = -1: Combo4.Text = ""
220     For N = 0 To Combo4.ListCount - 1
230         If Val(Left(Combo4.List(N), 3)) = Val(nNull(DATA2.Recordset("AEG"))) Then
240             Combo4.Text = Combo4.List(N)    'OIKOGENEIA
                Exit For
            End If
        Next

        Dim R   As New ADODB.Recordset

        Dim NYP As Integer

250     NYP = 0

        Dim NYP0 As Integer

         NYP0 = -1


          Combo3.Text = ""
260     If Len(Combo4.Text) > 0 Then

            ' On Error GoTo LATOSEXo
            On Error Resume Next

270         R.Open "SELECT *FROM PINAKES WHERE TYPOS=3 AND TIMH=" + Left(Combo4.Text, 2) + "ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
280         Combo3.Clear
           
290         Do While Not R.EOF

300             Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)

310             If R("AYJON") = Val(DATA2.Recordset("kodlog")) Then
320                 NYP0 = NYP
                End If

330             R.MoveNext
340             NYP = NYP + 1
            Loop

350         R.Close

        End If

        On Error Resume Next

360     Combo3.Text = Combo3.List(NYP0)

        'Combo3.Text = Combo3.List(Val(Data2.Recordset("kodlog")) - 1)

370     xti0.Text = DATA2.Recordset("xti")

        'If IsNull(Data2.Recordset("pos_kerd2")) Then
380     pos_kerd2.Text = nNull(DATA2.Recordset("pos_kerd2"))
        'End If

        'If IsNull(Data2.Recordset("pos_kerd")) Then
390     pos_kerd.Text = nNull(DATA2.Recordset("pos_kerd"))
        'End If

400     Text1(6).Text = CNull(DATA2.Recordset("ERG"))
410     ONO2.Text = CNull(DATA2.Recordset("ono2"))

420     lti.Text = DATA2.Recordset("lti")
430     lti5.Text = DATA2.Recordset("lti5")

        LTI2.Text = nNull(DATA2.Recordset("lti2"))
        LTI3.Text = nNull(DATA2.Recordset("lti3"))
 

        ' DATA2.Recordset("APALLFPA") = Val(Left(apallagesFPA.Text, 2))
        
        If IsNull(DATA2.Recordset("APALLFPA")) Then
             apallagesFPA.Text = ""
        Else
             apallagesFPA.Text = apallagesFPA.List(DATA2.Recordset("APALLFPA") - 1)
        End If
         

        PROMPOL.Text = nNull(DATA2.Recordset("PROMPOL"))
440     XondrmeFPA.Text = DATA2.Recordset("lti") * (1 + g_Fpa(DATA2.Recordset("fpa")) / 100)

450     If IsNull(DATA2.Recordset("MEMO")) Then
460         ' Text14.Text = ""
        Else
470         ' Text14.Text = DATA2.Recordset("MEMO")
        End If

480     Text1(15).Text = DATA2.Recordset("PROMHU")    ' BONUS

490     Text1(11).Text = DATA2.Recordset("kodSYNOD")
500     Text1(12).Text = DATA2.Recordset("kodlogAG")

510     Text1(7).Text = CNull(DATA2.Recordset("PROM"))    'Right(Combo3.Text, 3)
520     Text1(9).Text = DATA2.Recordset("SPA")   'Right(Combo3.Text, 3)

530     Text1(10).Text = CNull(DATA2.Recordset("UES"))    'Right(Combo3.Text, 3)
540     Text1(8).Text = DATA2.Recordset("EPIUYP")

550     exoda.Text = DATA2.Recordset("pos_kerd3")


            litra.Text = nNull(DATA2.Recordset("litra"))
           efk.Text = nNull(DATA2.Recordset("efk"))




560     If IsNull(DATA2.Recordset("POS01")) Then mPos01 = 0 Else mPos01 = DATA2.Recordset("POS01")

570     If IsNull(DATA2.Recordset("POS02")) Then mPos02 = 0 Else mPos02 = DATA2.Recordset("POS02")

580     If IsNull(DATA2.Recordset("POS03")) Then mPos03 = 0 Else mPos03 = DATA2.Recordset("POS03")
590     If IsNull(DATA2.Recordset("POS04")) Then mPos04 = 0 Else mPos04 = DATA2.Recordset("POS04")

        ' Label2.Caption = "ΤΕΛ.ΕΝΗΜΕΡΩΣΗ: KENΤΡ :" + Format(mPos01, "###0") + " ΥΠ.1: " + Format(mPos02, "###0") + " ΥΠ.3: " + Format(mPos03, "###0")

600     If Combo9.ListCount = 2 + 1 Then
610         Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " + Trim(mID(Combo9.List(0), 3, 10)) + " :" + Format(mPos01, "###0") + "   " + Trim(mID(Combo9.List(1), 3, 10)) + " :" + Format(mPos02, "###0")
620     ElseIf Combo9.ListCount = 1 + 1 Then
630         Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " + Trim(mID(Combo9.List(0), 3, 10)) + " :" + Format(mPos01, "###0")
640     ElseIf Combo9.ListCount = 3 + 1 Then
650         Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " + Trim(mID(Combo9.List(0), 3, 10)) + " :" + Format(mPos01, "###0") + " - " + Trim(mID(Combo9.List(1), 3, 10)) + " :" + "   " + Format(mPos02, "###0") + " - " + Trim(mID(Combo9.List(2), 3, 10)) + " :" + Format(mPos03, "###0")

660     ElseIf Combo9.ListCount = 4 + 1 Then
670         Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " + Trim(mID(Combo9.List(0), 3, 10)) + " :" + Format(mPos01, "###0") + "   " + Trim(mID(Combo9.List(1), 3, 10)) + " :  " + Format(mPos02, "###0") + "   " + Trim(mID(Combo9.List(2), 3, 10)) + " :" + Format(mPos03, "###0") + "   " + Trim(mID(Combo9.List(3), 3, 10)) + " :" + Format(mPos04, "###0")

        End If
















680     Text1(18).Text = DATA2.Recordset("SYSKMAX")    'SYSKEYASIA
690     Text1(17).Text = DATA2.Recordset("SYSKMIN")    ' YPOSYSKEYASIA
700     MSFlexGrid1.row = 0
710     MSFlexGrid1.Col = 1
720     MSFlexGrid1.Text = "Αγορές"
730     MSFlexGrid1.Col = 2
740     MSFlexGrid1.Text = "Αξία Αγορ."

750     MSFlexGrid1.Col = 3
760     MSFlexGrid1.Text = "Πωλήσεις"

770     MSFlexGrid1.Col = 4
780     MSFlexGrid1.Text = "Αξία πωλήσεων"

        '' Γραφική παράσταση
        's = 0

        'For l = 1 To 12
        ' Exit Sub
        '  G.row = 1 + addit: G.Column = l: G.Data = Stat(l, addit2 + 1) - Stat(l, addit2 + 2)
        ' G.row = 4 + addit: G.Column = l: G.Data = Stat(l, addit2 + 4) - Stat(l, addit2 + 5)
        'Next
790     If DATA2.Recordset.RecordCount = 0 Then
800         Me.MousePointer = vbNormal

810         If cXRONIES.Value = vbChecked Then Gdb.Close: Gdb.Open gConnect

            Exit Sub

        End If

820     SM = 0: SG = 0
830     SM2 = 0: SG2 = 0
        ' Dim R As New ADODB.Recordset

        On Error GoTo LATOS

        Dim synt

840     synt = " HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND HME<'" + Format(gLhjh, "MM/DD/YYYY") + "' "

850     If cXRONIES.Value = vbChecked Then
860         Gdb.Close
            'f_PROPERSI = Trim(mID(XRONIES.Text, 21, 20))
870         Gdb.Open gConnect + ";DATABASE=" + Trim(mID(XRONIES.Text, 21, 20))
880         synt = " YEAR(HME)>1900  "
        End If

890     If Check1.Value = vbUnchecked Then GoTo OLD700

900     R.Open "SELECT MONTH(HME),SUM(XRE),SUM(PIS),SUM(XRE*TIMM*(100-EKPT)/100),SUM(PIS*TIMM*(100-EKPT)/100) FROM EGGTIM WHERE " + synt + " AND KODE='" + DATA2.Recordset("kod") + "' GROUP BY MONTH(HME) ", Gdb, adOpenForwardOnly, adLockReadOnly

        Dim A_XRE(12)   As Single

        Dim A_PIS(12)   As Single

        Dim A_XREAJ(12) As Single

        Dim A_PISAJ(12) As Single

910     For k = 1 To 12: A_XRE(k) = 0: A_PIS(k) = 0: A_XREAJ(k) = 0: A_PISAJ(k) = 0: Next

        ' Exit Sub
920     If Not R.EOF Then

            On Error Resume Next    'GoTo NOCONNECT

930         R.MoveFirst

940         Do While Not R.EOF
                'On Error GoTo NOCONNECT
950             A_XRE(R(0)) = R(1): A_XREAJ(R(0)) = R(3)
960             A_PIS(R(0)) = R(2): A_PISAJ(R(0)) = R(4)
970             R.MoveNext
            Loop

980         R.Close

            On Error Resume Next

        End If

        'graph1.top = 4680  'arxika
        '  graph1.height = 4485

        ' graph1.top = 46
        'graph1.height = 8485

990     MSFlexGrid1.Visible = False

1000    For k = 1 To 12
            '  G = "G" + Format(k, "00")
            '  M = "M" + Format(k, "00")
            '  GA = "g" + Format(k, "00")
            '  MA = "M" + Format(k, "00")

1010        graph1.row = 1: graph1.Column = k: graph1.Data = A_XRE(k)    ' DATA2.Recordset(G)
1020        graph1.row = 2: graph1.Column = k: graph1.Data = A_PIS(k)    'DATA2.Recordset(M)
1030        graph1.ColumnLabel = k
1040        graph1.Refresh

1050        MSFlexGrid1.row = k
1060        MSFlexGrid1.Col = 0
1070        MSFlexGrid1.Text = k

1080        MSFlexGrid1.Col = 1

1090        SM = SM + A_PIS(k)    'DATA2.Recordset(M)
1100        SG = SG + A_XRE(k)    ' DATA2.Recordset(G)

1110        SM2 = SM2 + A_PISAJ(k)    'DATA2.Recordset(M)
1120        SG2 = SG2 + A_XREAJ(k)    ' DATA2.Recordset(G)

1130        MSFlexGrid1.Text = Format(A_XRE(k), "#####.##")    ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset(G)), "####.##")

1140        MSFlexGrid1.Col = 2
1150        MSFlexGrid1.Text = Format(A_XREAJ(k), "#####.##")    ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset("XTI") * DATA2.Recordset(G)), "#####,##")

1160        MSFlexGrid1.Col = 3
1170        MSFlexGrid1.Text = Format(A_PIS(k), "#####.##")    'Format(IIf(IsNull(DATA2.Recordset(M)), 0, DATA2.Recordset(M)), "#####,##")

1180        MSFlexGrid1.Col = 4
1190        MSFlexGrid1.Text = Format(A_PISAJ(k), "#####.##")    'Format(IIf(IsNull(DATA2.Recordset(MA)), 0, DATA2.Recordset("LTI") * DATA2.Recordset(MA)), "#####,##")

        Next

1200    If MSFlexGrid1.rows >= 14 Then

1210        MSFlexGrid1.rows = 13
        End If

        'MSFlexGrid1.AddItem ("Σύνολα" & Chr(9) & Format(SG, "####.##") & Chr(9) & Format(0, "####.##") & Chr(9) & Format(SM, "####.##"))
1220    MSFlexGrid1.AddItem ("Σύνολα" & Chr(9) & Format(SG, "####.##") & Chr(9) & Format(SG2, "####.##") & Chr(9) & Format(SM, "####.##") & Chr(9) & Format(SM2, "####.##"))
1230    MSFlexGrid1.Visible = True

OLD700:

1240    Adodc1.ConnectionString = gConnect

1250    If F_EXO_XROMATA = 0 Then

            'ΥΠΟΛΟΓΙΖΕΙ ΤΟ ΥΠΟΛΟΙΠΟ / ΑΝΤΙΑΚΑΤΑΣΤΑΘΗΚΕ ΜΕ ΕΝΑ LOOP ΠΑΡΑΚΑΤΩ ΓΙΑ ΝΑ ΜΗΝ ΑΡΓΕΙ
            'Adodc1.RecordSource = "SELECT  round( (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" _
            '& "  (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ),1) AS [ΥΠΟΛ]" _
            '& " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' AND " + synt
            'Adodc1.Refresh
            ' Text1(4).Text = Adodc1.Recordset(0)

1260        Adodc1.RecordSource = "SELECT KODE,APOT,(select PERIGRAFH FROM PINAKES WHERE TYPOS=4 AND AYJON=APOT) AS [ΑΠΟΘ],ROUND(SUM(XRE),0) AS [ΕΙΣ],ROUND(SUM(PIS),0) AS [ΕΞΑΓ]," & " round( (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" & "  (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ),1) AS [ΥΠΟΛ]" & " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' AND " + synt + "  GROUP BY KODE,APOT"

        Else

1270        Adodc1.RecordSource = "SELECT KODE,LEFT(PROELEYSH,10) AS [ΧΡ-ΜΕΓ],APOT,(select PERIGRAFH FROM PINAKES WHERE TYPOS=4 AND AYJON=APOT) AS [ΑΠΟΘ],ROUND(SUM(XRE),0) AS [ΕΙΣ],ROUND(SUM(PIS),0) AS [ΕΞΑΓ]," & " round( (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" & "  (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ),1) AS [ΥΠΟΛ]" & " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' AND " + synt + " GROUP BY APOT,KODE,LEFT(PROELEYSH,10) ORDER BY APOT,KODE,LEFT(PROELEYSH,10)"
        End If

        On Error GoTo mDBGrid1_RowColChange_Err

1280    Adodc1.Refresh
1290    MSHFlexGrid1.ColWidth(0) = 1000

1300    Text1(4).Text = 0

1310    Do While Not Adodc1.Recordset.EOF
1320        Text1(4).Text = Val(Text1(4).Text) + Adodc1.Recordset("ΥΠΟΛ")
1330        Adodc1.Recordset.MoveNext
        Loop

1340    SQL2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" & "  (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Text1(0).Text + "'"

        On Error GoTo LATOS

        On Error Resume Next

1350    R.Close

1360    R.Open SQL2, Gdb, adOpenForwardOnly, adLockReadOnly
1370    Label6.Caption = CNull(R(0))
1380    R.Close

1390    MSHFlexGrid1.ColWidth(0) = 1

1400    Adodc2.ConnectionString = gConnect

1410    If F_EXO_XROMATA = 0 Then
1420        Adodc2.RecordSource = "select ERG from BARCODES where KOD='" + Trim(Text1(0).Text) + "'"
        Else
1430        Adodc2.RecordSource = "select ERG,PROELEYSH from BARCODES where  KOD='" + Text1(0).Text + "'"
        End If

1440    Adodc2.Refresh

        'DataGrid2.DataSource = timokat

        ' timokat.ConnectionString = gConnect
        ' timokat.RecordSource = "select * from TIMOKAT where  LEFT(KOD,14)='" + Text1(0).Text + "'"
        ' timokat.Refresh

        Dim PI, pi2

1450    If F_EIKONA > 0 Then

            On Error Resume Next

1460        If F_EIKONA = 2 Then    'βοηθητικοσ κωδικοσ
1470            PI = F_FAKEL_KOINOS + "\IMAGES\" + Text1(6).Text
1480        ElseIf F_EIKONA = 1 Then    ' kvdikos
1490            PI = F_FAKEL_KOINOS + "\IMAGES\" + Text1(0).Text
1500        ElseIf F_EIKONA = 12 Then    'βοηθητικοσ κωδικοσ +.jpg
1510            PI = F_FAKEL_KOINOS + "\IMAGES\" + Text1(6).Text + ".JPG"
1520        ElseIf F_EIKONA = 11 Then    ' kodikos .jpg
1530            PI = F_FAKEL_KOINOS + "\IMAGES\" + Text1(0).Text + ".JPG"
1540        ElseIf F_EIKONA = 111 Then    ' kodikos .jpg
1550            PI = F_FAKEL_KOINOS + "\IMAGES\" + Trim(Text1(0).Text) + "*" + ".JPG"
            End If
            
            PI = Replace(PI, " ", "")
            
         
            If f_site > 0 Then ' = Val(FindParametroi(1,"MDIFORM1", "F_SITE", "0", "ΚΩΔΙΚΟΣ SITE 1=VM 2=TOYS"))
                PI = DATA2.Recordset("pic0")

                If f_site = 222 Then
                    pi2 = F_FAKEL_KOINOS + "\images\" + PI
                Else
                    pi2 = F_FAKEL_KOINOS + "\images\" + PI
                End If
  
                If Not IsNull(pi2) Then
                    If Len(Dir(pi2)) > 0 Then
                        Picture2.Picture = LoadPicture(pi2)
                        StretchSourcePictureFromPicture Me.Picture2.Picture, Me.Picture2
                    Else
                        Picture2.Picture = LoadPicture()
                    End If

                Else
                    Picture2.Picture = LoadPicture()
                End If

                ' Picture2.PaintPicture Picture1.Picture, 0, 0, Picture2.width, Picture2.height, 0, 0, Picture1.width, Picture1.height
                 
            Else

1560            If Len(f_site) > 0 Then
1570                Picture2.Picture = LoadPicture(PI)  ' F_FAKEL_KOINOS + "\images\" + Dir(PI))
1580                Picture2.PaintPicture Picture2.Picture, 0, 0, Picture2.Width, Picture2.Height, 0, 0, Picture2.Width, Picture1.Height

                     StretchSourcePictureFromPicture Me.Picture2.Picture, Me.Picture2

                Else
1590                'Picture1.Picture = LoadPicture()
1600                Picture2.Picture = LoadPicture()
                End If
            End If

        End If

        'DBGRID1.Columns(1).width = 4000

1610    NUM1.BackColor = IIf(IsNull(DATA2.Recordset("NUM1")), vbWhite, vbYellow)
1620    NUM2.BackColor = IIf(IsNull(DATA2.Recordset("NUM2")), vbWhite, vbYellow)
1630    NUM3.BackColor = IIf(IsNull(DATA2.Recordset("NUM3")), vbWhite, vbYellow)
1640    CH1.BackColor = IIf(IsNull(DATA2.Recordset("ch1")), vbWhite, vbYellow)
1650    CH2.BackColor = IIf(IsNull(DATA2.Recordset("ch2")), vbWhite, vbYellow)
1660    CH3.BackColor = IIf(IsNull(DATA2.Recordset("ch3")), vbWhite, vbYellow)
1670    ch4.BackColor = IIf(IsNull(DATA2.Recordset("ch4")), vbWhite, vbYellow)
1680    ch5.BackColor = IIf(IsNull(DATA2.Recordset("ch5")), vbWhite, vbYellow)

1690    NUM1 = IIf(IsNull(DATA2.Recordset("NUM1")), "", DATA2.Recordset("NUM1"))
1700    NUM2 = IIf(IsNull(DATA2.Recordset("NUM2")), "", DATA2.Recordset("NUM2"))
1710    NUM3 = IIf(IsNull(DATA2.Recordset("NUM3")), "", DATA2.Recordset("NUM3"))

1720    CH1 = IIf(IsNull(DATA2.Recordset("CH1")), "", DATA2.Recordset("CH1"))
1730    CH2 = IIf(IsNull(DATA2.Recordset("CH2")), "", DATA2.Recordset("CH2"))
1740    CH3 = IIf(IsNull(DATA2.Recordset("CH3")), "", DATA2.Recordset("CH3"))

1750    ch4 = IIf(IsNull(DATA2.Recordset("CH4")), "", DATA2.Recordset("CH4"))
1760    ch5 = IIf(IsNull(DATA2.Recordset("CH5")), "", DATA2.Recordset("CH5"))

1770    If IsNull(DATA2.Recordset("CH6")) Then ch6 = "" Else ch6 = DATA2.Recordset("CH6")

        If IsNull(DATA2.Recordset("CPV")) Then CPV = "" Else CPV = DATA2.Recordset("CPV")

1780    HM1 = IIf(IsNull(DATA2.Recordset("HM1")), "", DATA2.Recordset("HM1"))
1790    HM2 = IIf(IsNull(DATA2.Recordset("HM2")), "", DATA2.Recordset("HM2"))
1800    HM3 = IIf(IsNull(DATA2.Recordset("HM3")), "", DATA2.Recordset("HM3"))
        lastupd = IIf(IsNull(DATA2.Recordset("LASTUPD")), "", DATA2.Recordset("LASTUPD"))

        '===============================================================================
        '==================== PALIA XRONIA =============================================
        '========================================================================

1810    If Check1.Value = vbUnchecked Then
1820        Me.MousePointer = vbNormal

1830        If cXRONIES.Value = vbChecked Then Gdb.Close: Gdb.Open gConnect

            Exit Sub

        End If

        Dim DB2         As New ADODB.Connection

        Dim PaLIAXRONIA As String

        'On Error GoTo mDBGrid1_RowColChange_Err
1840    PaLIAXRONIA = Trim(FINDPARAMETROI(1, "PAR1", "F_PALIAXRONIA", "2005", "Γ.18 DSN ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ"))

1850    If Len(PaLIAXRONIA) = 0 Then
1860        Me.MousePointer = vbNormal

1870        If cXRONIES.Value = vbChecked Then Gdb.Close: Gdb.Open gConnect

            Exit Sub

        End If

        On Error GoTo oxi_palia

1880    DB2.Open gConnect + ";DATABASE=" + PaLIAXRONIA

        On Error Resume Next

1890    R.Close
1900    R.Open "SELECT MONTH(HME),SUM(XRE),SUM(PIS),SUM(XRE*TIMM),SUM(PIS*TIMM) FROM EGGTIM WHERE KODE='" + DATA2.Recordset("kod") + "' GROUP BY MONTH(HME) ", DB2, adOpenForwardOnly, adLockReadOnly

1910    For k = 1 To 12
1920        A_XRE(k) = 0: A_PIS(k) = 0: A_XREAJ(k) = 0: A_PISAJ(k) = 0
        Next

        ' Exit Sub
1930    If Not R.EOF Then

            On Error Resume Next

1940        R.MoveFirst

1950        Do While Not R.EOF
1960            A_XRE(R(0)) = R(1): A_XREAJ(R(0)) = R(3)
1970            A_PIS(R(0)) = R(2): A_PISAJ(R(0)) = R(4)
1980            R.MoveNext
            Loop

        End If

1990    MSFlexGrid2.Visible = False

2000    SM2 = 0: SG2 = 0
2010    SM = 0: SG = 0

2020    For k = 1 To 12
2030        graph2.row = 1: graph2.Column = k: graph2.Data = A_XRE(k)    ' DATA2.Recordset(G)
2040        graph2.row = 2: graph2.Column = k: graph2.Data = A_PIS(k)    'DATA2.Recordset(M)
2050        graph2.ColumnLabel = k
2060        graph2.Refresh

2070        MSFlexGrid2.row = k
2080        MSFlexGrid2.Col = 0
2090        MSFlexGrid2.Text = k

2100        MSFlexGrid2.Col = 1

2110        SM = SM + A_PIS(k)    'DATA2.Recordset(M)
2120        SG = SG + A_XRE(k)    ' DATA2.Recordset(G)

2130        SM2 = SM2 + A_PISAJ(k)    'DATA2.Recordset(M)
2140        SG2 = SG2 + A_XREAJ(k)    ' DATA2.Recordset(G)

2150        MSFlexGrid2.Text = Format(A_XRE(k), "#####.##")    ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset(G)), "####.##")

2160        MSFlexGrid2.Col = 2
2170        MSFlexGrid2.Text = Format(A_XREAJ(k), "#####.##")    ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset("XTI") * DATA2.Recordset(G)), "#####,##")

2180        MSFlexGrid2.Col = 3
2190        MSFlexGrid2.Text = Format(A_PIS(k), "#####.##")    'Format(IIf(IsNull(DATA2.Recordset(M)), 0, DATA2.Recordset(M)), "#####,##")

2200        MSFlexGrid2.Col = 4
2210        MSFlexGrid2.Text = Format(A_PISAJ(k), "#####.##")    'Format(IIf(IsNull(DATA2.Recordset(MA)), 0, DATA2.Recordset("LTI") * DATA2.Recordset(MA)), "#####,##")

        Next

2220    If MSFlexGrid2.rows >= 14 Then

2230        MSFlexGrid2.rows = 13
        End If

2240    MSFlexGrid2.AddItem ("Σύνολα" & Chr(9) & Format(SG, "####.##") & Chr(9) & Format(SG2, "####.##") & Chr(9) & Format(SM, "####.##") & Chr(9) & Format(SM2, "####.##"))
2250    MSFlexGrid2.Visible = True

2260    Me.MousePointer = vbNormal

2270    If cXRONIES.Value = vbChecked Then Gdb.Close: Gdb.Open gConnect


     diastaseis_Grid



        Exit Sub

NOCONNECT:

2280    If cXRONIES.Value = vbChecked Then Gdb.Close: Gdb.Open gConnect

        Exit Sub

LATOSEXo:

2290    If cXRONIES.Value = vbChecked Then Gdb.Close: Gdb.Open gConnect

        Exit Sub

LATOS:
2300    MsgBox Err.Description

2310    Resume Next

oxi_palia:

2320    If cXRONIES.Value = vbChecked Then Gdb.Close: Gdb.Open gConnect

        Exit Sub

        '<EhFooter>
        Exit Sub

mDBGrid1_RowColChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.mDBGrid1_RowColChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.mDBGrid1_RowColChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo DBGrid1_MouseUp_Err

        '</EhHeader>

100    ' mDBGrid1_RowColChange
       
        '<EhFooter>
        Exit Sub

DBGrid1_MouseUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.DBGrid1_MouseUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.DBGrid1_MouseUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   Dim a As String
   a = DBGrid1.columns(0).Text
   
   mDBGrid1_RowColChange
   


End Sub

Private Sub ENERGA_Click()

fmEN = "1"
'ENERGA.Text = ENERGA.List(0)
Select Case ENERGA.ListIndex
    Case 0
       fmEN = "1"
    Case 1
       fmEN = "0"
    Case 2
       fmEN = "1','0"
    
    
    

End Select





End Sub

'Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'  Dim d
'   d = DBGrid1.Text
' If f_Refresh Then
' mDBGrid1_RowColChange
' End If

'End Sub

Private Sub erg_GotFocus()

        '<EhHeader>
        On Error GoTo erg_GotFocus_Err

        '</EhHeader>

100     erg.BackColor = vbYellow

        '<EhFooter>
        Exit Sub

erg_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.erg_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.erg_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub erg_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo erg_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

erg_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.erg_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.erg_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub erg_LostFocus()

        '<EhHeader>
        On Error GoTo erg_LostFocus_Err

        '</EhHeader>

100     erg.BackColor = vbWhite

110     ono.Text = ""
120     kod.Text = ""
130     barc.Text = ""

140     cmdAnazit_Click

        '<EhFooter>
        Exit Sub

erg_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.erg_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.erg_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "APOT2"
110     PARAMETROI.SHOW 1
        Form_Resize
        

        'F_TAB = Val(FindParametroi(1,"APOT2", "F_TAB", "3", "Σε ποιό TAB είναι σταματημένο"))
        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
        Dim a

100     a = 1

110     If KeyCode = 27 Then
120         CMDEXIT_Click
        End If

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
    
 '  Me.Hide
    
        ' El.mForm_Load Me
         mForm_Load Me, fh, fw, ft, fl

fmEN = "1"
ENERGA.Text = ENERGA.List(0)



'  mForm_Load Me, fh, fw, ft, fl

        Dim DB, R As New ADODB.Recordset

100     KeyPreview = True
110     Me.Picture = LoadPicture(gPicture)
        ' Command3.Picture = LoadPicture("c:\mercvb\page.gif")
120     alignGridTimon

        ' Picture4.Picture = LoadPicture(gPicture)

        'APOT2.SHOW
        times.RowHeight(0) = 200
        ' f_Refresh = True

        ''SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", "NormalOffice2007.ini"

        ''SkinFramework1.ApplyWindow apot2.hWnd
        ''SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
        On Error Resume Next
        
        
        F_ET_EPIX_FONTSIZE = Val(FINDPARAMETROI(1, "APOT2", "F_ET_EPIX_FONTSIZE", "12", "FONT SIZE ΕΠΙΧΕΙΡΗΣΗΣ"))
        F_ET_EPIX_X = Val(FINDPARAMETROI(1, "APOT2", "F_ET_EPIX_X", "100", "Χ ΕΠΙΧΕΙΡΗΣΗΣ"))
        F_ET_EPIX_Y = Val(FINDPARAMETROI(1, "APOT2", "F_ET_EPIX_Y", "200", "Y ΕΠΙΧΕΙΡΗΣΗΣ"))
        
        
        
        F_ARXEIO_ETIK = FINDPARAMETROI(1, "APOT2", "F_ARXEIO_ETIK", "C:\MERCVB\EGGTIM2.TXT", "APXEIO ΕΤΙΚΕΤΤΩΝ ΑΠΟ ΚΙΝΗΤΟ")
        f_printer = FINDPARAMETROI(1, "APOT2", "F_PRINTER", "SATO", "ΜΕΡΟΣ ΤΗΣ ΟΝΟΜΑΣΙΑΣ ΤΟΥ ΕΚΤΥΠ.ΕΤΙΚΕΤΤΩΝ")
        
        
        'f_wresize, f_Hresize
        f_wresize = Val(FINDPARAMETROI(1, "APOT2", "f_wresize", "15000", "πλατος resize(15000-18500)"))
        f_Hresize = Val(FINDPARAMETROI(1, "APOT2", "f_Hresize", "9945", "υψος resize(9500-10500) "))
        
        
        
        
        F_EuroPerLiter = Val(FINDPARAMETROI(1, "APOT2", "F_EuroPerLiter", "0.2", "ΤΙΜΗ ΕΙΔ.ΦΟΡ. ΑΝΑ ΛΙΤΡΟ"))
 
        F_DEK_LIANIKIS = Val(FINDPARAMETROI(1, "APOT2", "F_DEK_LIANIKIS", "2", "ΔΕΚΑΔΙΚΑ ΛΙΑΝΙΚΗΣ"))

        F_0_BARCODE = Val(FINDPARAMETROI(1, "APOT2", "F_0_BARCODE", "0", "1=ΑΓΝΟΩ ΤΟ 1ο ΨΗΦΙΟ ΑΝ ΕΙΝΑΙ 0  0=ΤΟ ΒΑRCODE ΟΠΩΣ ΕΙΝΑΙ"))
        

        F_XPONTOI = Val(FINDPARAMETROI(1, "APOT2", "F_XPONTOI", "2", "TAB(X) METATOΠIΣH ΠΟΝΤΩΝ"))
        F_YPONTOI = Val(FINDPARAMETROI(1, "APOT2", "F_YPONTOI", "550", "Y METATOΠIΣH ΠΟΝΤΩΝ"))
        'As Long, F_YPONTOI

        F_MIK = Val(FINDPARAMETROI(1, "APOT2", "F_MIK", "0", "11=τυπωνω barcode 1=ΤΥΠΩΝΩ kod σαν BARCODE 0=ΠΕΡΙΓ. ΣΤΟ ΚΟΥΜΠΙ ΠΕΡΙΓ"))

        f_dek_xondr = Val(FINDPARAMETROI(1, "APOT2", "f_dek_xondr", "4", "Αριθμός δεκαδικών χονδρικής τιμής"))

        f_site = Val(FINDPARAMETROI(1, "MDIFORM1", "F_SITE", "0", "ΚΩΔΙΚΟΣ SITE 1=VM 2=TOYS"))

        F_FAKEL_KOINOS = FINDPARAMETROI(1, "APOT2", "F_FAKEL_KOINOS", "C:\MERCVB", "KOINOXPHΣTOΣ ΦΑΚΕΛΟΣ ΕΙΚΟΝΩΝ Π.Χ. \\PC\MERCVB")
F_FAKEL_KOINOS = Trim$(F_FAKEL_KOINOS)
        If f_site > 0 Then
            cmdWEB.Visible = True
        End If

130     If Len(Dir("c:\mercfan.txt")) > 0 Then
140         Open "c:\mercfan.txt" For Input As #1
150         Input #1, fConnect2    'DUMMY
160         Input #1, fConnect2
170         Close #1
180         FGDB2.Open fConnect2

        End If

190     R.Open "select DSN,TITLOS FROM ETAIREIES ORDER BY TITLOS DESC", Gdb, adOpenDynamic, adLockOptimistic
200     R.MoveFirst
210     XRONIES.Clear

220     Do While Not R.EOF

230         If Not IsNull(R("TITLOS")) And (Not IsNull(R("DSN"))) Then
240             XRONIES.AddItem Left(R("TITLOS") + Space(20), 20) + R("DSN")
            End If

250         R.MoveNext
        Loop

260     R.Close

270     If Len(Dir("C:\LAGEURO\LAGAKIS.TXT", vbNormal)) > 0 Then
280         Command15.Visible = True

            'Exit Sub
        End If

290     F_CHANGETIMH = Val(FINDPARAMETROI(1, "APOT2", "F_CHANGETIMH", "0", "ΑΛΛΑΖΕΙ ΑΥΤΟΜΑΤΑ ΤΗΝ ΧΟΝΔ.ΤΙΜΗ=1/0"))
300     F_UPPER = Val(FINDPARAMETROI(1, "APOT2", "F_UPPER", "1", "ΜΕΤΑΤΡΟΠΗ ΠΕΡΙΓΡΑΦΩΝ ΣΕ ΚΕΦΑΛΑΙΑ ΕΛΛΗΝΙΚΑ=1/0"))

310     F_EIKONA = Val(FINDPARAMETROI(1, "APOT2", "F_EIKONA", "1", "1=EIKONA=>ΚΩΔΙΚΟΣ 2=>ΜΕ ΒΟΗΘ.ΚΛΕΙΔΙ 11=ΚVDIKOS.JPG 12=BOH.JPG 0=ΟΧΙ ΕΙΚΟΝΕΣ"))

320     F_add_barc = Val(FINDPARAMETROI(1, "PAR1", "add_barc", "1", "Γ.16 Barcode έιναι 1.Κωδ 2.Βοηθ.Κλειδί"))

330     f_YPOL_MHN = Val(FINDPARAMETROI(1, "APOT2", "f_YPOL_MHN", "1", "Γ.1 ΠΡΟΕΠΙΛΟΓΗ ΥΠΟΛ.ΜΗΝΙΑΙΩΝ=1 / 0"))
        'As Integer



'F_CHANGE_MARKUP
     F_CHANGE_MARKUP = Val(FINDPARAMETROI(1, "APOT2", "F_CHANGE_MARKUP", "1", "Γ.9 ME ΑΛΛΑΓΗ ΧΟΝΔΡΙΚΗΣ-ΛΙΑΝΙΚΗΣ ΑΛΛΑΖΕΙ MARKUP=1 / OXI=0"))

        f_XOND_MEFPA_DISABLE = Val(FINDPARAMETROI(1, "APOT2", "f_XOND_MEFPA_DISABLE", "0", "Γ.8 ΑΠΕΝΕΡΓΟΠΟΙΜΕΝΗ Η ΧΟΝΔ. ΜΕ ΦΠΑ=1 / 0"))

        If f_XOND_MEFPA_DISABLE = 1 Then
             XondrmeFPA.Enabled = False
             
        
        End If
        
340     If f_YPOL_MHN = 1 Then
350         Check1.Value = vbChecked
        Else
360         Check1.Value = vbUnchecked

        End If

370     F_EXO_XROMATA = Val(FINDPARAMETROI(1, "PAR1", "F_EXO_XROMATA", "0", "Γ.1 ΧΡΗΣΙΜΟΠΟΙΩ ΜΕΓΕΘΗ-ΧΡΩΜΑΤΑ=1"))   'posa psifia tha exei h kathe seira

380     f_FontName = FINDPARAMETROI(1, "APOT2", "F_FONTNAME", "Arial", "FONT-NAME BARCODE")   '"Arial"

390     f_tab = Val(FINDPARAMETROI(1, "APOT2", "F_TAB", "3", "Προεπιλεγμένο TAB "))
400     SSTab1.Tab = IIf(f_tab > SSTab1.Tabs - 1 Or f_tab < 0, SSTab1.Tabs - 1, f_tab)

410     Me.Picture = LoadPicture(gPicture)

420     F_ArPerEtik = Val(FINDPARAMETROI(1, "APOT2", "F_ArPerEtik", "2", "Αριστερό περιθώριο ετικέττας")) 'posa psifia tha exei h kathe seira

430     F_2ArPerEtik = Val(FINDPARAMETROI(1, "APOT2", "F_2ArPerEtik", "2", "2o Αριστερό περιθώριο ετικέττας")) 'posa psifia tha exei h kathe seira

440     F_PIChEIGHT = Val(FINDPARAMETROI(1, "APOT2", "F_PIChEIGHT", "2535", "Υψος εικόνας"))   'posa psifia tha exei h kathe seira
450     F_PICwidth = Val(FINDPARAMETROI(1, "APOT2", "F_PICwidth", "3675", "Πλάτος εικόνας"))   'posa psifia tha exei h kathe seira

460     F_XBARCODE = Val(FINDPARAMETROI(1, "APOT2", "F_XBARCODE", "348", "barcode Χ"))   'posa psifia tha exei h kathe seira
470     F_YBARCODE = Val(FINDPARAMETROI(1, "APOT2", "F_YBARCODE", "824", "barcode Y"))   'posa psifia tha exei h kathe seira

480     F_Epix = FINDPARAMETROI(1, "APOT2", "F_EPIX", "--", "TITΛOS EΠIX ΣΤHN ETIKETA")   'posa psifia tha exei h kathe seira

490     Picture2.Height = F_PIChEIGHT
500     Picture2.Width = F_PICwidth

510     F_sygx = Val(FINDPARAMETROI(1, "APOT2", "F_SYGX", "1", "ΕΠΙΤΡΕΠΕΤΑΙ ΣΥΓΧΩΝΕΥΣΗ ΚΩΔΙΚΩΝ=1 0=ΟΧΙ"))   'posa psifia tha exei h kathe seira

520     times.TextMatrix(0, 1) = "Tιμή Αγοράς"
530     times.TextMatrix(0, 2) = "Mark Up"
540     times.TextMatrix(0, 3) = "Χωρίς ΦΠΑ"
550     times.TextMatrix(0, 4) = "Με ΦΠΑ"

560     times.TextMatrix(1, 0) = "Χονδρ.Τιμή"
570     times.TextMatrix(2, 0) = "Λιαν.Τιμή"

        times.TextMatrix(0, 6) = "Προτεινόμενη"

580     F_KOD = Val(FINDPARAMETROI(1, "APOT2", "F_KOD", "1", "Εμφανίζεται η αναζήτηση με κωδικό=1"))
590     F_ONO = Val(FINDPARAMETROI(1, "APOT2", "F_ONO", "1", "Εμφανίζεται η αναζήτηση με ΟΝΟΜΑ=1"))
600     F_ERG = Val(FINDPARAMETROI(1, "APOT2", "F_ERG", "1", "Εμφανίζεται η αναζήτηση με ΒΟΗΘ.ΚΛΕΙΔΙ=1"))
610     F_XTI = Val(FINDPARAMETROI(1, "APOT2", "F_XTI", "1", "Εμφανίζεται η αναζήτηση με BARCODE=1"))

620     If F_KOD = 0 Then kod.Visible = False
630     If F_ONO = 0 Then ono.Visible = False
640     If F_ERG = 0 Then erg.Visible = False
650     If F_XTI = 0 Then barc.Visible = False

660     F_LHM1 = FINDPARAMETROI(1, "APOT2", "F_lhm1", "Ημερ1", "Ετικέτα ημερ.1")
670     F_LHM2 = FINDPARAMETROI(1, "APOT2", "F_lhm2", "Ημερ2", "Ετικέτα ημερ.2")
680     f_lHM3 = FINDPARAMETROI(1, "APOT2", "F_lhm3", "Ημερ3", "Ετικέτα ημερ.3")
690     lhm1.Caption = F_LHM1
700     lhm2.Caption = F_LHM2
710     lhm3.Caption = f_lHM3
720     f_lab1 = FINDPARAMETROI(1, "APOT2", "F_lab1", "Κειμ1", "Ετικέτα κειμ.1")
730     f_lab2 = FINDPARAMETROI(1, "APOT2", "F_lab2", "Κειμ2", "Ετικέτα κειμ.2")
740     f_lab3 = FINDPARAMETROI(1, "APOT2", "F_lab3", "Κειμ3", "Ετικέτα κειμ.3")

750     f_lab4 = FINDPARAMETROI(1, "APOT2", "F_lab4", "Κειμ4", "Ετικέτα κειμ.4")
760     f_lab5 = FINDPARAMETROI(1, "APOT2", "F_lab5", "Κειμ5", "Ετικέτα κειμ.5")
770     f_lab6 = FINDPARAMETROI(1, "APOT2", "F_lab6", "Κειμ6", "Ετικέτα κειμ.6")

780     lab1.Caption = f_lab1
790     lab2.Caption = f_lab2
800     lab3.Caption = f_lab3

810     lab4.Caption = f_lab4
820     lab5.Caption = f_lab5
830     lab6.Caption = f_lab6

840     f_num1 = FINDPARAMETROI(1, "APOT2", "F_num1", "Αριθ1", "Ετικέτα κειμ.1")
850     f_num2 = FINDPARAMETROI(1, "APOT2", "F_num2", "Αριθ2", "Ετικέτα κειμ.2")
860     f_num3 = FINDPARAMETROI(1, "APOT2", "F_num3", "Αριθ3", "Ετικέτα κειμ.3")
870     Lnum1.Caption = f_num1
880     Lnum2.Caption = f_num2
890     Lnum3.Caption = f_num3

        ' Data1.DatabaseName = gDir
        'Data1.Connect = gConnect

        'Data2.DatabaseName = gDir
        'Data2.Connect = gConnect

        ' Set db = OpenDatabase(gDir, False, False, gConnect)
        'If Len(Dir(gDir + "\pinakes.ndx")) < 2 Then
        '   db.Execute "CREATE INDEX PINAKES ON PINAKES(TYPOS,ayjon)"
        'End If

        '  Set R = db.OpenRecordset("pinakes")

        'This example creates an index consisting of the fields Home Phone and Extension in the Employees table.
        '   Dim dbs As Database
        ' Modify this line to include the path to Northwind
        ' on your computer.
        '  Set dbs = OpenDatabase("Northwind.mdb")
        ' Create the NewIndex index on the Employees table.
        ' dbs.Execute "CREATE INDEX NewIndex ON Employees " _
        '    & "(HomePhone, Extension);"
        'dbs.Close
        'db.Execute "CREATE INDEX PINAKES ON PINAKES(TYPOS,"

900     If gXEIRISTHS < 8 Then
910         Command3.Visible = False
920         xti0.Visible = False
930         pos_kerd.Visible = False
940         pos_kerd2.Visible = False
950         lti.Visible = False
960         Label1(5).Visible = False
970         Label1(12).Visible = False
980         Label1(6).Visible = False
990         Label1(23).Visible = False
        End If

        'Set db = OpenDatabase(gDir, False, False, gConnect)
1000    R.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

        'R.Index = "pinakes"

        'FPA
        'R.Seek "=", 1, 1
1010    Do While Not R.EOF

1020        If R("typos") = 1 Then
1030            Combo2.AddItem str(R("AYJON")) + " -> " + str(R("TIMH"))
            End If

1040        R.MoveNext
        Loop

        ' mon.metrhshs
1050    R.Close

1060    R.Open "SELECT *FROM PINAKES WHERE TYPOS=2 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

1070    Do While Not R.EOF

1080        If R("typos") = 2 Then
1090            Combo1.AddItem R("PERIGRAFH")
            Else

                Exit Do

            End If

1100        R.MoveNext
        Loop

        ' kathgories

1110    R.Close

        'αποθηκη
1120    R.Open "SELECT *FROM PINAKES WHERE TYPOS=4 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

1130    Do While Not R.EOF

1140        If R("typos") = 4 Then
1150            Combo9.AddItem Format(R("ayjon"), "##") + "." + R("PERIGRAFH")
            Else

                Exit Do

            End If

1160        R.MoveNext
        Loop

1170    Combo9.Text = Combo9.List(0)

1180    R.Close

        ' OIKOGENEIES

1190    R.Open "SELECT *FROM PINAKES WHERE TYPOS=11 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

        'R.MoveFirst
1200    Do While Not R.EOF

1210        If R("typos") = 11 Then
1220            Combo4.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
            Else

                Exit Do

            End If

1230        R.MoveNext
        Loop

1240    R.Close

        ' ΚΑΤΗΓΟΡΙΕΣ

1250    R.Open "SELECT *FROM PINAKES WHERE TYPOS=15 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
        'R.MoveFirst

1260    If R.EOF Then
1270        Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH) VALUES (15,1,'ΕΜΠΟΡΕΥΜΑ')"
        End If

1280    Do While Not R.EOF

1290        If R("typos") = 15 Then
1300            Combo5.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
            Else

                Exit Do

            End If

1310        R.MoveNext
        Loop

1320    R.Close

1330    F_PONTOI = Val(FINDPARAMETROI(1, "APOT2", "F_PONTOI", "1", "ΣΤΗΝ ΕΤΙΚΕΤΤΑ ΠΟΝΤΟΥΣ=1 ΟΧΙ=0"))

1340    'DBGrid1.ColWidth(1) = 3300
        DBGrid1.columns(1).Width = 3800
        DBGrid1.columns(2).Width = 800
        DBGrid1.columns(3).Width = 900

DBGrid1.columns(1).alignment = dbgLeft
DBGrid1.columns(2).alignment = dbgRight
DBGrid1.columns(3).alignment = dbgRight



1350    'DBGrid1.ColWidth(3) = 1250
1360   'DBGrid1.ColAlignment(3) = 7

1370    'DBGrid1.ColWidth(2) = 1250
1380   ' DBGrid1.ColAlignment(2) = 7



1390    Data1.RecordSource = "SELECT TOP 100 *FROM EID"
1400    Data1.ConnectionString = gConnect
1410    Data1.Refresh

1420    DATA2.ConnectionString = gConnect

1430    data22.ConnectionString = gConnect
1440    data22.RecordSource = "select TOP 30  KOD,ONO,CONVERT(decimal(10,2),LTI) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ]  from EID where POS>0;"
1450    data22.Refresh

1460    PROTH_FORA = 0

diastaseis_Grid


'If False Then  ' αποχρωσεις στο ΤΑΒ αλλα στην μεγεθυνση χανει τα ΤΕΧΤ πεδια

  ' If Option4.Value = True Then
           gDirection = cHorizontal
      '  Else
        '   gDirection = cVertical
       ' End If
        gColor1 = &HC0FFC0    'vbBlue '&HFFFFC0     ' vbBlue
        gColor2 = vbCyan      '&HFF0000
        SetStyle SSTab1.hwnd, cGradient  '//--- Set The Style of The SSTab
        SetGradientDir SSTab1.hwnd, gDirection '//--- Set The Gradient Direction
        SetGradientColor1 SSTab1.hwnd, gColor1  '//--- Asing new Gradient Color Start
        SetGradientColor2 SSTab1.hwnd, gColor2  '//--- Asing new Gradient Color End
 
     SSTabSubclass SSTab1.hwnd '//--- Begin SubClassing
     RedrawWindow SSTab1.hwnd, ByVal 0&, ByVal 0&, &H1
'End If



  'ΦΟΡΤΩΜΑ ΑΠΑΛΛΑΓΩΝ ΦΠΑ
    R.Open "SELECT *FROM PINAKES WHERE TYPOS=44 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
apallagesFPA.Clear

    Do While Not R.EOF
        apallagesFPA.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH")
        R.MoveNext
    Loop

R.Close

















        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Option1_Click(Index As Integer)
'
'End Sub

Private Sub Form_Paint()

        '<EhHeader>
        On Error GoTo Form_Paint_Err

        '</EhHeader>

100     If PROTH_FORA = 0 Then
110         PROTH_FORA = 1

            On Error Resume Next

120         kod.SetFocus
        End If

        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Form_Paint " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Form_Paint " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()
  
  
    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
 
   ' ResizeFormFor Me
  
'  ResizeForm Me 'frm As Form)




  
 locForm_Resize Me, 8, f_wresize, f_Hresize, fh, fw, ft, fl   ' h=10545
  
  
  
  
  
'    If F_PLATOS_FORMAS < 0 Then
'        ResizeFormFor Me
'    Else
 '       mForm_Resize Me, 8, F_PLATOS_FORMAS, F_YCOS_FORMAS, fh, fw, ft, fl
'    End If
'
  
    search.Left = DBGrid1.Left
   
   DBGrid1.Width = Picture2.Left + Picture2.Width - DBGrid1.Left
    search.Width = DBGrid1.Width
    
    
  times.Left = Frame1.Left
  
 
 ' El.mForm_Load Me
         'mForm_Load Me, fh, fw, ft, fl
' Exit Sub

times.Width = Frame2.Width

search.Width = DBGrid1.Width

 
 
 search.Height = DBGrid1.Top - search.Top - 30

    times.row = 1
    times.Col = 1
    xti0.Top = times.Top + times.CellTop
    xti0.Left = times.Left + times.CellLeft
    xti0.Width = times.CellWidth - 15
    xti0.Height = times.CellHeight - 25

    times.row = 1
    times.Col = 2
    pos_kerd2.Top = times.Top + times.CellTop
    pos_kerd2.Left = times.Left + times.CellLeft
    pos_kerd2.Width = times.CellWidth - 15
    pos_kerd2.Height = times.CellHeight - 25

    times.row = 1
    times.Col = 3
    lti.Top = times.Top + times.CellTop
    lti.Left = times.Left + times.CellLeft
    lti.Width = times.CellWidth - 15
    lti.Height = times.CellHeight - 25

    times.row = 1
    times.Col = 4
    XondrmeFPA.Top = times.Top + times.CellTop
    XondrmeFPA.Left = times.Left + times.CellLeft
    XondrmeFPA.Width = times.CellWidth - 15
    XondrmeFPA.Height = times.CellHeight - 25

    times.row = 2
    times.Col = 2
    pos_kerd.Top = times.Top + times.CellTop - 25
    pos_kerd.Left = times.Left + times.CellLeft
    pos_kerd.Width = times.CellWidth - 15
    pos_kerd.Height = times.CellHeight - 25

    times.row = 2
    times.Col = 4
    lti5.Top = times.Top + times.CellTop - 25
    lti5.Left = times.Left + times.CellLeft
    lti5.Width = times.CellWidth - 15
    lti5.Height = times.CellHeight - 25
        
    times.Width = (3 * XondrmeFPA.Width + XondrmeFPA.Left - times.Left) + 200
    times.Height = times.CellHeight * 3
    Picture4.Width = times.Width ' + 10
    Picture4.Height = times.Height + 10

End Sub

Sub locForm_Resize(frm As Form, _
                        fontHeight As Single, _
                        ByVal WW As Long, _
                        ByVal HH As Long, _
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
    For k = 0 To frm.Controls.Count - 1
    
    
    
    If UCase(frm.Controls(k).Container.Name) = "SSTAB1" Then  ' If x.Container.Name = "SSTab1" Then
    
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

    'End If

End Sub







Private Sub kod_GotFocus()

        '<EhHeader>
        On Error GoTo kod_GotFocus_Err

        '</EhHeader>

100     kod.BackColor = vbYellow

        '<EhFooter>
        Exit Sub

kod_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.kod_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.kod_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub kod_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo kod_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

kod_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.kod_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.kod_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub kod_LostFocus()

        '<EhHeader>
        On Error GoTo kod_LostFocus_Err

        '</EhHeader>

100     kod.BackColor = vbWhite

        ' If Len(Trim(kod.Text)) > 0 Then

110     ono.Text = ""
120     erg.Text = ""
130     barc.Text = ""

140     cmdAnazit_Click

        'End If

        '<EhFooter>
        Exit Sub

kod_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.kod_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.kod_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Lab6_Click()

        '<EhHeader>
        On Error GoTo Lab6_Click_Err

        '</EhHeader>

        Dim l_lngRetVal As Long

        Const SW_SHOWNORMAL = 1

100     l_lngRetVal = ShellExecute(0&, vbNullString, ch6.Text, vbNullString, "C:\", SW_SHOWNORMAL)

        '<EhFooter>
        Exit Sub

Lab6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Lab6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Lab6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label1_Click(index As Integer)

        '----------- OIKOGENEIES ---------------------------
        '<EhHeader>
        On Error GoTo Label1_Click_Err

        '</EhHeader>
        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        'Sub UPDATE_PINAKES(MTYPOS As Integer, OO As Object, MCOMBO As Object)
        'mtypos  αριθμος πίνακα
        'οο: to object που θέλω για το κεντράρισμα της οθόνης DataEntry
        'mcombo : το combo που θα γεμίσει με τις τιμές

100     If index = 2 Then    'monades metrhshs

110         UPDATE_PINAKES 2, Label1(index), Combo1
        End If

120     If index = 25 Then    ' oikogeneies
130         UPDATE_PINAKES 11, Label1(index), Combo4
        End If

        '-----------------------------------------------------------------------------------------
        'If Index = 2 Then 'ΜΟΝΑΔΕΣ ΜΕΤΡΗΣΗΣ
        '  'sql οπου θα γίνεται το update ή το addnew
        ' '  f.SHOW
        '  Load f
        '
        '  Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=2"
        '  f.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=2 " '    sql
        ''τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
        'f.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " _
        '& " FROM PINAKES WHERE TYPOS=2 ORDER BY AYJON " 'ParamGrid   sql2
        '
        ''με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
        'f.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=2 AND " 'ParamGrid              DelSQL
        '
        ''ποιό update query θα δημιουργείται (προαιρετικά )
        ''αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
        ''                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
        ' DD = " UPDATE PINAKES SET TYPOS=2 where TYPOS IS NULL AND AYJON="" + t1   "
        '' AddNewScript
        '  DD = "Sub Main()" & vbCrLf & _
        '         " DIM mCon  " & vbCrLf & _
        '         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
        '         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
        '         " mCon.execute """ + DD + " " & vbCrLf & _
        '        "Set r=Nothing" & vbCrLf & _
        '        "Set mCon=Nothing" & vbCrLf & _
        '        "End Sub"
        'f.ParamGrid.TextMatrix(0, 4) = DD
        'f.ParamGrid.TextMatrix(1, 1) = "A/A"
        'f.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "
        ''ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
        'f.Labels(0).width = f.Labels(0).width + 1000 '* TextWidth("A")
        'f.Textboxes(0).Left = f.Textboxes(0).Left + 1000 ' * TextWidth("A")
        'f.Left = Label1(Index).Left
        'f.Top = Label1(Index).Top + MDIForm1.Top
        '
        '
        '
        'f.SHOW 1
        'Combo1.Clear
        ' R.Open "SELECT *FROM PINAKES WHERE TYPOS=2 ORDER BY AYJON", Gdb, adOpenDynamic, _
        '     adLockOptimistic
        'Do While Not R.EOF
        '  If R("typos") = 2 Then
        '     Combo1.AddItem R("PERIGRAFH")
        '  Else
        '     Exit Do
        '  End If
        '  R.MoveNext
        'Loop
        'R.Close
        '
        '
        '
        '
        '
        '
        'End If
        '

        ''-----------------------------------------------------------------------------------------
        'If Index = 25 Then 'OIKOGENEIES
        '  'sql οπου θα γίνεται το update ή το addnew
        ' '  f.SHOW
        '  Load f
        '
        '  Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=11"
        '  f.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=11 " '    sql
        ''τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
        'f.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " _
        '& " FROM PINAKES WHERE TYPOS=11 ORDER BY AYJON " 'ParamGrid   sql2
        '
        ''με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
        'f.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=11 AND " 'ParamGrid              DelSQL
        '
        ''ποιό update query θα δημιουργείται (προαιρετικά )
        ''αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
        ''                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
        ' DD = " UPDATE PINAKES SET TYPOS=11 where TYPOS IS NULL AND AYJON="" + t1   "
        '' AddNewScript
        '  DD = "Sub Main()" & vbCrLf & _
        '         " DIM mCon  " & vbCrLf & _
        '         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
        '         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
        '         " mCon.execute """ + DD + " " & vbCrLf & _
        '        "Set r=Nothing" & vbCrLf & _
        '        "Set mCon=Nothing" & vbCrLf & _
        '        "End Sub"
        'f.ParamGrid.TextMatrix(0, 4) = DD
        'f.ParamGrid.TextMatrix(1, 1) = "A/A"
        'f.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "
        ''ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
        'f.Labels(0).width = f.Labels(0).width + 1000 '* TextWidth("A")
        'f.Textboxes(0).Left = f.Textboxes(0).Left + 1000 ' * TextWidth("A")
        '
        '
        'f.Left = Label1(Index).Left
        'f.Top = Label1(Index).Top + MDIForm1.Top
        '
        ''f.height = 18000
        '
        'f.SHOW 1
        '
        'Combo4.Clear
        ' R.Open "SELECT *FROM PINAKES WHERE TYPOS=11 ORDER BY AYJON", Gdb, adOpenDynamic, _
        '     adLockOptimistic
        'Do While Not R.EOF
        '       Combo4.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
        '  R.MoveNext
        'Loop
        'R.Close
        '
        'End If
        '

        '----------------------------------------------------------------------------------------
140     If index = 4 Then    'YPO - OIKOGENEIES
150         If Val(Left(Combo4.Text, 2)) = 0 Then

                Exit Sub

            End If

            'sql οπου θα γίνεται το update ή το addnew
160         F.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=3 AND TIMH= " + Left(Combo4.Text, 2)    '    sql
            'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
170         F.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " & " FROM PINAKES WHERE TYPOS=3 AND TIMH=" + Left(Combo4.Text, 2) + " ORDER BY AYJON "    'ParamGrid   sql2

            'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
180         F.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=3  AND TIMH=" + Left(Combo4.Text, 2) + " AND "    'ParamGrid              DelSQL

            'ποιό update query θα δημιουργείται (προαιρετικά )
            'αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
            '                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL

190         F.ParamGrid.TextMatrix(0, 6) = "TYPOS"    'EPIPLEON FIELD
200         F.ParamGrid.TextMatrix(0, 7) = 3    ' TIMH TOY EPIPLEON FIELD

210         F.ParamGrid.TextMatrix(0, 8) = "TIMH"    '2 EPIPLEON FIELD
220         F.ParamGrid.TextMatrix(0, 9) = Left(Combo4.Text, 2)    ' 2 TIMH TOY EPIPLEON FIELD

            ' DD = " UPDATE PINAKES SET TYPOS=3,TIMH=" + Left(Combo4.Text, 2) + " where TYPOS IS NULL AND AYJON="" + t1   "
            '' AddNewScript
            '  DD = "Sub Main()" & vbCrLf & _
            '         " DIM mCon  " & vbCrLf & _
            '         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
            '         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
            '         " mCon.execute """ + DD + " " & vbCrLf & _
            '        "Set r=Nothing" & vbCrLf & _
            '        "Set mCon=Nothing" & vbCrLf & _
            '        "End Sub"
230         F.ParamGrid.TextMatrix(0, 4) = ""    ' DD
240         F.ParamGrid.TextMatrix(1, 1) = "A/A"
250         F.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "

            'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
260         F.Labels(0).Width = F.Labels(0).Width + 1000    '* TextWidth("A")
270         F.Textboxes(0).Left = F.Textboxes(0).Left + 1000    ' * TextWidth("A")
            'f.Left = 0
            'f.Top = 0
            'f.height = 18000

280         F.Left = Label1(index).Left
290         F.Top = Label1(index).Top + MDIForm1.Top
300         F.SHOW 1

310         Combo3.Clear
320         R.Open "SELECT * FROM PINAKES WHERE TYPOS=3 AND TIMH= " + Left(Combo4.Text, 2) + " ORDER BY AYJON ", Gdb, adOpenDynamic, adLockOptimistic

330         Do While Not R.EOF
340             Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
350             R.MoveNext
            Loop

360         R.Close

        End If

        '  apot2.WindowState = 2

        '<EhFooter>
        Exit Sub

Label1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Label1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Label1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub lti_GotFocus()
    '         lti.BackColor = vbYellow
    '         lti.SelStart = 0
    '         lti.SelLength = Len(lti.Text)
    OBJECT_GOTFOCUS lti
End Sub

Private Sub lti_KeyPress(KeyAscii As Integer)
    enter_to_tab KeyAscii
End Sub

Private Sub lti_LostFocus()
    YPOL_TIMON 3, lti
End Sub

Sub OBJECT_GOTFOCUS(ByRef D As TextBox)
    D.BackColor = vbYellow
    D.SelStart = 0
    D.SelLength = Len(D.Text)
         
End Sub

Private Sub lti5_GotFocus()
    OBJECT_GOTFOCUS lti5
End Sub

Private Sub lti5_KeyPress(KeyAscii As Integer)
    enter_to_tab KeyAscii
End Sub

Private Sub lti5_LostFocus()
    YPOL_TIMON 13, lti5
End Sub

Private Sub ono_GotFocus()

        '<EhHeader>
        On Error GoTo ono_GotFocus_Err

        '</EhHeader>

100     ono.BackColor = vbYellow

        '<EhFooter>
        Exit Sub

ono_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.ono_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.ono_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub ono_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo ono_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

ono_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.ono_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.ono_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub ono_LostFocus()

        '<EhHeader>
        On Error GoTo ono_LostFocus_Err

        '</EhHeader>

100     ono.BackColor = vbWhite

110     kod.Text = ""
120     erg.Text = ""
130     barc.Text = ""

140     cmdAnazit_Click

        'DBGrid1.SetFocus
        '<EhFooter>
        Exit Sub

ono_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.ono_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.ono_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub pos_kerd_GotFocus()
    OBJECT_GOTFOCUS pos_kerd
End Sub

Private Sub pos_kerd_KeyPress(KeyAscii As Integer)
    enter_to_tab KeyAscii
End Sub

Private Sub pos_kerd_LostFocus()
    YPOL_TIMON 14, pos_kerd
End Sub

Private Sub pos_kerd2_GotFocus()
    OBJECT_GOTFOCUS pos_kerd2
End Sub

Private Sub pos_kerd2_KeyPress(KeyAscii As Integer)
    enter_to_tab KeyAscii
End Sub

Private Sub pos_kerd2_LostFocus()
    YPOL_TIMON 5, pos_kerd2
End Sub

'Private Sub TabStrip1_Click()
'    Dim a
'    a = SSTab1.TabIndex
'End Sub

Private Sub Text1_GotFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text1_GotFocus_Err

        '</EhHeader>

100     Text1(index).BackColor = vbYellow

        Text1(index).SelStart = 0
120     Text1(index).SelLength = Len(Text1(index).Text)

        '<EhFooter>
        Exit Sub

Text1_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Text1_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Text1_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Text1_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then

110         KeyAscii = 0    'suppress the beep

120         If index = 13 Then
130             Command2.SetFocus

            Else
140             keybd_event VK_TAB, 0, 0, 0    'send a tab
            End If
        End If

        '<EhFooter>
        Exit Sub

Text1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Text1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Text1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_Change(index As Integer)
    ' If Index = 5 Or Index = 4 Then
    '      lti.Text = Val(xti.Text) * (100 + Val(pos_KERD2.Text)) / 100
    ' End If

    '  If Index = 3 And Val(xti.Text) > 0 Then 'ΟΤΑΝ ΑΛΛΑΖΩ ΤΙΜΗ XONTRIKHS (3)  ΑΛΛΑΖΕΙ ΤΟ ΜΑΡΚΑΠ (5)      Η ΑΓΟΡΑ ΕΙΝΑΙ ΤΟ (2)
    '      pos_KERD2.Text = Val(lti.Text) / Val(xti.Text) * 100
    ' End If

    '  If Index = 13 And Val(xti.Text) > 0 Then 'ΟΤΑΝ ΑΛΛΑΖΩ ΤΙΜΗ ΛΙΑΝΙΚΗΣ (13)  ΑΛΛΑΖΕΙ ΤΟ ΜΑΡΚΑΠ (14)      Η ΑΓΟΡΑ ΕΙΝΑΙ ΤΟ (2)
    '      pos_KERD.Text = Val(lti5.Text) / (1 + g_Fpa(Val(Combo2.Text))) / Val(xti.Text) * 100
    ' End If

End Sub

Sub YPOL_TIMON(ByVal index As Integer, ByRef D As TextBox)

        ' 13= lti5
100     If index = 13 And Val(xti0.Text) > 0 Then    'ΟΤΑΝ ΑΛΛΑΖΩ ΤΙΜΗ ΛΙΑΝΙΚΗΣ (13)  ΑΛΛΑΖΕΙ ΤΟ ΜΑΡΚΑΠ (14)      Η ΑΓΟΡΑ ΕΙΝΑΙ ΤΟ (2)
            'pos_KERD.Text = (Val(lti5.Text) / Val(XTI0.Text) - 1) * 100
            If F_CHANGE_MARKUP = 1 Then
110             pos_kerd.Text = Round(gVal(lti5.Text) / (1 + g_Fpa(gVal(Combo2.Text)) / 100) / gVal(xti0.Text) * 100 - 100, 3)
            End If
            XondrmeFPA.Text = Format(Val(Replace(XondrmeFPA.Text, ",", ".")), "#####.00")
            lti.Text = Format(Val(Replace(lti.Text, ",", ".")), "#####." + String(f_dek_xondr, "0"))
            lti5.Text = Format(Val(Replace(lti5.Text, ",", ".")), "#####." + String(F_DEK_LIANIKIS, "0"))
        End If   ' Format(Val(lti5.Text), "###0." + String(F_DEK_LIANIKIS, "0"))

        '3 =lti
120     If index = 3 And Val(xti0.Text) > 0 Then    'ΟΤΑΝ ΑΛΛΑΖΩ ΤΙΜΗ XONTRIKHS (3)  ΑΛΛΑΖΕΙ ΤΟ ΜΑΡΚΑΠ (5)      Η ΑΓΟΡΑ ΕΙΝΑΙ ΤΟ (2)


            If F_CHANGE_MARKUP = 1 Then
130               pos_kerd2.Text = Round((gVal(lti.Text) / gVal(xti0.Text) - 1) * 100, 3)
            End If

140         XondrmeFPA.Text = Round(gVal(lti.Text) * (1 + g_Fpa(gVal(Combo2.Text)) / 100), 2)
            lti.Text = Format(Val(Replace(lti.Text, ",", ".")), "#####." + String(f_dek_xondr, "0"))
            lti5.Text = Format(Val(Replace(lti5.Text, ",", ".")), "#####." + String(F_DEK_LIANIKIS, "0"))
        End If

150     If index = 5 Then  'pos_kerd2
160         If Val(Combo2.Text) > 0 Then
170             ' προτεινόμενη τιμη   proteinX.Caption = Val(XTI0.Text) * (100 + Val(pos_KERD2.Text)) / 100
                times.TextMatrix(1, 6) = Round(gVal(xti0.Text) * (100 + gVal(pos_kerd2.Text)) / 100, 2)
                XondrmeFPA.Text = Format(Val(Replace(XondrmeFPA.Text, ",", ".")), "#####.00")
                lti.Text = Format(Val(Replace(lti.Text, ",", ".")), "#####." + String(f_dek_xondr, "0"))
                lti5.Text = Format(Val(Replace(lti5.Text, ",", ".")), "#####.00")
            End If

180         If F_CHANGETIMH = 1 Then
190             lti.Text = Format(gVal(xti0.Text) * (100 + gVal(pos_kerd2.Text)) / 100, "###0." + String(f_dek_xondr, "0"))
200             XondrmeFPA.Text = Format(gVal(lti.Text) * (1 + g_Fpa(Val(Combo2.Text)) / 100), "###0.00")
                lti.Text = Format(Val(Replace(lti.Text, ",", ".")), "#####." + String(f_dek_xondr, "0"))
                lti5.Text = Format(Val(Replace(lti5.Text, ",", ".")), "#####.00")
            End If

        End If

210     ' Text1(Index).BackColor = vbWhite

        ' 14 = pos_kerd
220     If index = 14 Then
230         If Val(Combo2.Text) > 0 Then
240             ' προτεινομενη τιμη  '  proteinX.Caption = Val(XTI0.Text) * (1 + g_Fpa(Val(Combo2.Text)) / 100) * (100 + Val(pos_KERD.Text)) / 100
                times.TextMatrix(2, 6) = Round(gVal(xti0.Text) * (1 + g_Fpa(gVal(Combo2.Text)) / 100) * (100 + Val(pos_kerd.Text)) / 100, 2)
                XondrmeFPA.Text = Format(Val(Replace(XondrmeFPA.Text, ",", ".")), "#####.00")
                lti.Text = Format(Val(Replace(lti.Text, ",", ".")), "#####." + String(f_dek_xondr, "0"))
                lti5.Text = Format(Val(Replace(lti5.Text, ",", ".")), "#####.00")
            End If
        End If
        
        D.BackColor = vbWhite

'        XondrmeFPA.Text = Format(Val(Replace(XondrmeFPA.Text, ",", ".")), "#####.00")
'        lti.Text = Format(Val(Replace(lti.Text, ",", ".")), "#####." + String(f_dek_xondr, "0"))
'        LTI5.Text = Format(Val(Replace(LTI5.Text, ",", ".")), "#####.00")

        'lti.Text

        '<EhFooter>
        Exit Sub

Text1_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.Text1_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Text1_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_LostFocus(index As Integer)

    '<EhHeader>
    ' On Error GoTo Text1_LostFocus_Err
    '</EhHeader>

    Text1(index).BackColor = vbWhite

End Sub

Sub alignGridTimon()

        'timh agoras
        '<EhHeader>
        On Error GoTo alignGridTimon_Err

        '</EhHeader>

100     xti0.Left = times.Left + times.CellWidth + 60
110     xti0.Top = times.Top + times.CellHeight + 70
120     xti0.Width = times.CellWidth

        'markup xontrikhs
130     pos_kerd2.Left = times.Left + 2 * times.CellWidth + 60
140     pos_kerd2.Top = times.Top + times.CellHeight + 70
150     pos_kerd2.Width = times.CellWidth
        'timh xontrikhs
160     lti.Left = times.Left + 3 * times.CellWidth + 80
170     lti.Top = times.Top + times.CellHeight + 70
180     lti.Width = times.CellWidth
        'timh xontrikhs me fpa
190     XondrmeFPA.Left = times.Left + 4 * times.CellWidth + 90
200     XondrmeFPA.Top = times.Top + times.CellHeight + 70
210     XondrmeFPA.Width = times.CellWidth

        'markup lianikhs
220     pos_kerd.Left = times.Left + 2 * times.CellWidth + 80
230     pos_kerd.Top = times.Top + 2 * times.CellHeight + 70
240     pos_kerd.Width = times.CellWidth
        'timh lianikhs me fpa
250     lti5.Left = times.Left + 4 * times.CellWidth + 90
260     lti5.Top = times.Top + 2 * times.CellHeight + 50
270     lti5.Width = times.CellWidth

280     lti5.Height = times.CellHeight

        '<EhFooter>
        Exit Sub

alignGridTimon_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.alignGridTimon " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.alignGridTimon " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
'Private Sub Text14_GotFocus()
'        '<EhHeader>
'        On Error GoTo Text14_GotFocus_Err
'        '</EhHeader>
'100    ' Text14.BackColor = vbYellow
'110     'Text14.ForeColor = vbBlack
'        '<EhFooter>
'        Exit Sub
'
'Text14_GotFocus_Err:
'        'MsgBox Err.Description & vbCrLf & _
'               "in ADOMERCNEW.APOT2.Text14_GotFocus " & _
'               "at line " & Erl, _
'               vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Text14_GotFocus " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub

'Private Sub Text14_LostFocus()
    '<EhHeader>
    '       On Error GoTo Text14_LostFocus_Err
    '</EhHeader>
    ' Text14.BackColor = vbWhite
    '
    '<EhFooter>

    '      Exit Sub

'Text14_LostFocus_Err:
    'MsgBox Err.Description & vbCrLf & _
    '        "in ADOMERCNEW.APOT2.Text14_LostFocus " & _
    '        "at line " & Erl, _
    '        vbExclamation + vbOKOnly, "Application Error"
    '     SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.Text14_LostFocus " & "at line " & Erl
    '    Resume Next
    '</EhFooter>'

'End Sub

Private Sub XondrmeFPA_Change()

    On Error Resume Next

    If XondrmeFPA.DataChanged = True Then
        If F_CHANGE_MARKUP = 1 Then
            pos_kerd2.Text = Round(gVal(XondrmeFPA.Text) / (1 + g_Fpa(gVal(Combo2.Text)) / 100) / gVal(xti0.Text) * 100 - 100, 3)
        End If
        
       If f_XOND_MEFPA_DISABLE = 1 Then
       
       Else
          lti.Text = Round(gVal(XondrmeFPA.Text) / (1 + g_Fpa(gVal(Combo2.Text)) / 100), 3)
       End If
    End If

End Sub

Private Sub XondrmeFPA_GotFocus()

    XondrmeFPA.SelStart = 0
    XondrmeFPA.SelLength = Len(XondrmeFPA.Text)

End Sub

Private Sub XondrmeFPA_KeyPress(KeyAscii As Integer)

100     If KeyAscii = vbKeyReturn Then

110         KeyAscii = 0    'suppress the beep
120         'If Index = 13 Then
130         '   pos_KERD.SetFocus

            ' Else
140         keybd_event VK_TAB, 0, 0, 0    'send a tab
            'End If
        End If

End Sub

Private Sub XondrmeFPA_LostFocus()

    'YPOLOGIZEI TO POSOSTO OTAN ALLAZO THN TIMH ME FPA
    '<EhHeader>
    On Error GoTo XondrmeFPA_LostFocus_Err

    '</EhHeader>

    '        On Error Resume Next
    '        If XondrmeFPA.DataChanged = True Then
    '100        pos_KERD2.Text = Val(XondrmeFPA.Text) / (1 + g_Fpa(Val(Combo2.Text)) / 100) / Val(xti.Text) * 100 - 100
    '110        lti.Text = Round(Val(XondrmeFPA.Text) / (1 + g_Fpa(Val(Combo2.Text)) / 100), 3)
    '        End If
        
    '<EhFooter>
    Exit Sub

XondrmeFPA_LostFocus_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.APOT2.XondrmeFPA_LostFocus " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.XondrmeFPA_LostFocus " & "at line " & Erl

    Resume Next

    '</EhFooter>

End Sub

Private Sub xti_Change()
    ' data22.recordsource = "select *from EID where pos>0 and xti>" + Str(Val(xti)) + " order by xti;"
    ' Data2.Refresh
End Sub

Private Sub xti_GotFocus()

        '<EhHeader>
        On Error GoTo xti_GotFocus_Err

        '</EhHeader>
100     ' XTI.BackColor = vbYellow
        '<EhFooter>
        Exit Sub

xti_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.xti_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.xti_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub xti_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo xti_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

xti_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.xti_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.xti_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub xti_LostFocus()

        '<EhHeader>
        On Error GoTo xti_LostFocus_Err

        '</EhHeader>

100     ' XTI.BackColor = vbWhite

110     ono.Text = ""
120     erg.Text = ""
130     kod.Text = ""

        On Error Resume Next

140     cmdAnazit_Click

        '<EhFooter>
        Exit Sub

xti_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.xti_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.xti_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub MAKIS_ETIK()

        '<EhHeader>
        On Error GoTo MAKIS_ETIK_Err

        '</EhHeader>

        Dim k, timologio

        Dim rec As New ADODB.Recordset

        't = "c:\lageuro\makisget.exe "
        't = t + par1.EID.Recordset("KOD") + " "
        't = t + Replace(par1.EID.Recordset("ONO"), " ", "_") + " "
        't = t + Format(Val(Text1(4).Text), "##0.00") + " "   ' LTI5  LIANIKI
        't = t + Format(Val(xti.Text), "##0") + " "  ' POS_KERD
        'If Len(Trim(mERG)) > 0 Then
        '   t = t + mERG
        'Else
        '   t = t + "00"
        'End If

        Dim MPER

100     MPER = DATA2.Recordset("ONO")

        Dim mtimes, TIMM, pos_kerd

110     mtimes = 1
120     TIMM = Format(Val(lti5.Text), "##0.00")
130     MPER = Text1(1).Text

140     pos_kerd = Format(Val(pos_kerd.Text), "##0") + " -"    ' POS_KERD

        '   timologio = par1.EID.Recordset("KOD") + Left(par1.Text1.Text, 6) + "=" + par1.PEL.Recordset("KOD")
        '  timologio = timologio + Format(Val(Label11.Caption) * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(par1.DTPicker1, "dd/mm/yyyy")
        Dim mlabel

150     Printer.FontBold = True

160     Printer.CurrentY = 30

170     Printer.FontSize = 10
180     Printer.Print Tab(2); MPER

190     Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
200     Printer.FontSize = 26

210     Printer.CurrentY = 600
220     Printer.Print "      " + TIMM + " €"
230     Printer.FontSize = 10
240     Printer.CurrentY = 1450
250     Printer.Print "  " + Text1(6).Text + "-" + Format(500 + Val(pos_kerd.Text), "####0") + " -" + Text1(0).Text

260     Printer.EndDoc

        Exit Sub

FIND_N_NUMBER:
        '************************

        On Error GoTo MAKIS_ETIK_Err

270     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

280     rec.MoveFirst
290     rec.Move 10

        'rec.Edit
300     If IsNull(rec("epel")) Then
310         rec("epel") = 1
        Else
320         rec("epel") = rec("epel") + 1
        End If

330     rec.Update
340     rec.Close

350     mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)
360     Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Data1.Recordset("KOD") + "'"
370     Return

        '<EhFooter>
        Exit Sub

MAKIS_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT2.MAKIS_ETIK " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT2.MAKIS_ETIK " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub xti0_GotFocus()
    OBJECT_GOTFOCUS xti0
End Sub

Private Sub xti0_KeyPress(KeyAscii As Integer)
    enter_to_tab KeyAscii
End Sub

'If Len(Dir("C:\FANAR.EXE", vbNormal)) > 0 Then
'   fanar_etik
'End If

'Sub fanar_etik()
'    Dim X, X2
'    Dim barc(500)
'    Dim apoSCANNER
'
'
'
'    'yupie
'    'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
'    Dim a$
'
'    X = 0
'
'    If Scanner Then
'        '
'        Open "C:\MERCVB\APOG.TXT" For Input As #1
'        X = 0
'
'        Do While Not EOF(1)    'True
'            On Error Resume Next
'            Line Input #1, a$
'
'            kod = mID(a$, 1, 13)
'            ' pos = Val(mID$(a$, 15, 20))
'            If Len(Trim(Left(kod, 3))) = 0 Then
'                ' TIPOTA
'            Else
'                X = X + 1
'                barc(X) = kod
'            End If
'
'
'        Loop
'        Close #1
'
'        If X = 0 Then
'            MsgBox "δεν υπάρχουν εγγραφές"
'            Exit Sub
'        End If
'        apoSCANNER = X
'    End If
'
'
'    ' ΑΝ ΔΕΝ ΕΙΝΑΙ ΑΠΟ ΣΚΑΝΕΡ
'    If X = 0 Then
'        apoSCANNER = 1
'        barc(1) = DataGrid1.Text  ' Text1(6).Text
'    End If
'
'
'
'
'    Dim mkod
'
'
'    Dim mFORES As Long
'
'
'    Dim R As New ADODB.Recordset
'    Dim nn
'    Dim ll
'    Dim mONOMA, mERG, mLTI5
'    For ll = 1 To apoSCANNER
'
'
'        R.Open "select BARCODES.ERG,EID.* FROM BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD  WHERE BARCODES.ERG='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'
'        If R.EOF Or Len(Trim(barc(ll))) = 0 Then
'            mFORES = 0
'        Else
'            mONOMA = R("ONO")    'ONOMA
'            mkod = R("KOD")    '  ' BARCODE
'
'            mLTI5 = Format(R("LTI5"), "###0.00")    ' LTI5
'            mERG = R(0)    ' erg
'
'        End If
'        R.Close
'
'
'
'
'
'
'
'
'        If Len(mERG) < 12 Then
'            ' MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
'            ' Exit Sub
'        End If
'
'        If Len(mERG) = 12 Then
'            BarCode3.SymbologyID = UPCA
'
'        ElseIf Len(mERG) = 13 Then
'            BarCode3.SymbologyID = EAN13
'        Else
'            BarCode3.SymbologyID = CODE128    ' UCC128 '  EAN8 '  ' "128"  ' UPCA
'            BarCode3.Code128CharSet = Set_C
'        End If
'
'        For Each X2 In Printers
'            If InStr(X2.DeviceName, "2844") > 0 Then
'                ' Set printer as system default.
'
'                Set Printer = X2
'                ' Stop looking for a printer
'                Exit For
'            End If
'        Next
'
'        BarCode3.DataToEncode = mERG
'
'
'        'For K = 1 To fores
'
'        Printer.CurrentX = F_ArPerEtik * 100
'        Printer.CurrentY = -150
'        Printer.FONTSIZE = 10
'        Printer.FontBold = True
'        Printer.FontName = "COURIER"    ' "Arial Greek"
'
'
'        '   Printer.Print ""
'        Printer.Print Tab(3); Left(UCase(mONOMA), 24)    'perigrafh
'        Printer.Print Tab(3); mID(UCase(mONOMA), 25, 25)    'perigrafh
'
'
'        Printer.CurrentY = 1300    ' 1600 ΟΚ 10-4-2010
'        Printer.CurrentX = 2000
'        Printer.FONTSIZE = 8
'        Printer.Print mERG    'kodikos
'
'
'        Printer.CurrentY = 1600    ' 1600 ΟΚ 10-4-2010
'        Printer.CurrentX = 2000
'        Printer.FONTSIZE = 8
'
'        Printer.Print mkod  'kodikos
'
'
'        If Scanner = False And ll = 1 Then
'            FGDB2.Execute "update Material  set retailprice=" + lti5.Text + "  WHERE  Code='" + Text1(0).Text + "'", nn
'            'FGDB2.Execute "update Material  set retailprice=" + mLti5 + "  WHERE  Code='" + mKod + "'", nn
'            If nn = 0 Then
'                MsgBox "δεν δενημερωθηκε το singular"
'            End If
'        End If
'
'
'        ' Printer.CurrentX = 2048
'        'Printer.Print s
'        Printer.CurrentX = 400 + F_ArPerEtik * 100
'        'Printer.FontSize = 12
'        Printer.FONTSIZE = 10
'
'
'        Printer.CurrentX = 200 + F_ArPerEtik * 100
'        Printer.CurrentY = 700    ' 600
'
'        Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 20, 1200, , 500
'
'        Printer.CurrentY = 600    ' 600
'        Printer.CurrentX = 1000 + F_ArPerEtik * 100
'        Printer.FONTSIZE = 30
'        Printer.FontBold = True
'        Printer.Print mLTI5 + " €"    ' timh  ' Format(lti5.Text, "##0.00")
'        ' Printer.Print "4.55 €" ' timh
'        Printer.EndDoc
'    Next
'    If Scanner Then
'        Kill "C:\MERCVB\APOG.TXT"
'    End If
'
'
'
'    Exit Sub
'
'End Sub

Private Sub xti0_LostFocus()
    YPOL_TIMON 2, xti0

    xti0.SelLength = 0
End Sub
Private Sub LITRA_LostFocus()
   On Error Resume Next
   efk.Text = Format(litra.Text * F_EuroPerLiter, "#0.0000")
    
End Sub

Sub dion22()

        'opos to dion 2 me polla barcode
        'yupie
        'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
            
        Dim barc(9000)
            
        Dim apoSCANNER As Integer

        apoSCANNER = 1

        If Scanner Then
            If Len(Dir("C:\MERCVB\EGGTIM2.TXT", vbNormal)) > 0 Then
                Open "c:\mercvb\eggtim2.txt" For Input As #6

                Dim rSQL   As New ADODB.Recordset

                Dim rSQL2  As New ADODB.Recordset

                Dim rSQL3  As New ADODB.Recordset

                Dim RR()   As String, a As String

                Dim M_CODE As String

                Dim k      As Integer

                k = 0

                Do While True ' Not EOF(6)
                    k = k + 1
                    Line Input #6, a
                    RR = Split(a, ";")
                    barc(k) = RR(7)

                    If EOF(6) Then

                        Exit Do

                    End If

                Loop

                Me.Caption = str(k) + "  ΕΤΙΚΕΤΤΕΣ"
                Close #6
                apoSCANNER = k
            End If
        End If

        Dim ll     As Long

        Dim mFORES As Integer

        Dim C0     As String

        Dim mONOMA, mERG As String

        Dim mLTI5    As String

        Dim R        As New ADODB.Recordset

        Dim r1       As New ADODB.Recordset

        Dim mBarcode As String

        Dim X      As Printer

2720    For ll = 1 To apoSCANNER

2730        mFORES = 1

2740        If Scanner Then
2750            'R.Close
2760            R.Open "SELECT * FROM BARCODES WHERE ERG='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

                If R.EOF Then
                    C0 = "0"
                Else
2770                C0 = R("KOD")
                End If

2780            R.Close

2790            R.Open "select * FROM EID WHERE KOD='" + C0 + "'", Gdb, adOpenForwardOnly, adLockReadOnly

2800            If R.EOF Or Len(Trim(barc(ll))) = 0 Then
2810                mFORES = 0
                Else
                    'το r1 xrhsimopoihuhke για να τυπωνει to barcode 15/02/2010
2820                r1.Open "SELECT top 1 * FROM BARCODES WHERE KOD='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
2830                mBarcode = r1("ERG")    '  R("KOD") '  ' BARCO
2840                r1.Close
2850                mONOMA = R("ONO")    'ONOMA
                    'DE
2860                Text1(6).Text = R("ERG")

2870                mERG = R("ERG")    ' erg
2880                mLTI5 = Format(R("LTI5"), "###0." + String(F_DEK_LIANIKIS, "0"))   ' LTI5
                End If

            Else
                r1.Open "SELECT top 1 ERG FROM BARCODES WHERE KOD='" + Text1(0).Text + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                mBarcode = r1("ERG")
                r1.Close
                C0 = Text1(0).Text
                mONOMA = Text1(1).Text
                mLTI5 = lti5.Text

            End If

            R.Close

            For k = 1 To mFORES

612             If Len(mBarcode) < 12 Then
624                 MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"

                    Exit Sub

                End If

626             If Len(mBarcode) = 12 Then
628                 BarCode3.SymbologyID = UPCA
                Else
629                 BarCode3.SymbologyID = EAN13
                End If

630             For Each X In Printers

640                 If InStr(X.DeviceName, "2844") > 0 Then
                        ' Set printer as system default.

650                     Set Printer = X

                        ' Stop looking for a printer
                        Exit For

                    End If

                Next

660             BarCode3.DataToEncode = mBarcode

                'For K = 1 To fores

670             Printer.CurrentX = F_ArPerEtik * 100
680             Printer.CurrentY = 200
690             Printer.FontSize = 10
700             Printer.FontBold = False
710             Printer.FontName = "Arial Greek"

720             Printer.Print mONOMA   'perigrafh

730             Printer.CurrentX = F_ArPerEtik * 100

740             Printer.Print C0  'kodikos

                ' Printer.CurrentX = 2048
                'Printer.Print s
750             Printer.CurrentX = 400 + F_ArPerEtik * 100
                'Printer.FontSize = 12
760             Printer.FontSize = 10

770             Printer.CurrentX = 200 + F_ArPerEtik * 100
780             Printer.CurrentY = 600

790             Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 1124

800             Printer.CurrentX = 1000 + F_ArPerEtik * 100
810             Printer.FontSize = 24
820             Printer.FontBold = True
830             Printer.Print Format(mLTI5, "##0.00") + " €"    ' timh
                ' Printer.Print "4.55 €" ' timh
831             Printer.EndDoc
                'Next

            Next  ' MFORES

        Next

End Sub

Sub polla_alletik()
'=====================================================================
   



Dim CC As String



CC = F_ARXEIO_ETIK
If Len(CC) = 0 Then
     cd1.InitDir = "c:\mercvb"
     cd1.ShowOpen
     CC = cd1.FILENAME
End If

On Error Resume Next


Dim MKOD As String
Dim MONO As String
Dim mLTI5 As Single



Dim a As String


Open CC For Input As #1
Dim RR() As String
Dim MFPA As String
Dim R As New ADODB.Recordset
Dim sql As String

Dim a7 As Integer:
a = vbNo ' a7 = MsgBox("ενημερωνω βαση δεδομένων με τιμες? ", vbYesNo)



'Do While Not EOF(1)

    Line Input #1, a
   
   RR = Split(a, ";")
   MKOD = RR(0)
   'MONO = RR(1)
   'mLTI5 = Val(Replace(RR(2), ",", "."))
   'MFPA = RR(3)
   'If Len(MFPA) = 0 Then
   '   MFPA = "2"
   'End If
 
   
 'Loop
   Close #1
   
   
Dim nn As Long
For nn = 0 To UBound(RR) - 1
   MKOD = RR(nn)
   
   R.Open "select * from EID WHERE KOD LIKE '" + LTrim(Trim(MKOD)) + "%'", Gdb, adOpenDynamic, adLockOptimistic
   
 If a7 = vbYes Then
   If R.EOF Then
      sql = "INSERT INTO EID (KOD,ONO,LTI5,FPA) VALUES ('" + MKOD + "','" + MONO + "'," + Replace(Format(mLTI5, "###.00"), ",", ".") + "," + MFPA + ")"
      Gdb.Execute sql
      sql = "INSERT INTO BARCODES (KOD,ERG) VALUES ('" + MKOD + "','" + MKOD + "' )"
      Gdb.Execute sql
      
      
   Else
     
     Gdb.Execute "UPDATE  EID SET LTI5=" + Replace(Format(mLTI5, "###.00"), ",", ".") + " WHERE KOD='" + MKOD + "'"
     Gdb.Execute "UPDATE  EID SET ONO='" + Replace(MONO, "'", "`") + "' WHERE KOD='" + MKOD + "'"
   End If
      
  
   
   Else  ' ΝΑ ΜΗΝ ΕΝΗΜΕΡΨΝΕΙ ΤΗΝ ΒΑΣΗ
         
       If Not R.EOF Then
           MONO = R("ONO")
           mLTI5 = R("LTI5")
       End If
   End If

 R.Close

230    If Len(MKOD) = 12 Then
240     BarCode3.SymbologyID = UPCA
250 ElseIf Len(MKOD) = 13 Then
260     BarCode3.SymbologyID = EAN13
    Else
270     BarCode3.SymbologyID = CODE128
    End If
Dim X
280 For Each X In Printers

290     If InStr(X.DeviceName, "SATO") > 0 Then
            ' Set printer as system default.

300         Set Printer = X

            ' Stop looking for a printer
            Exit For

        End If

    Next
        
310 BarCode3.DataToEncode = MKOD

320 Printer.CurrentX = F_ArPerEtik * 100
330 Printer.CurrentY = 200
340 Printer.FontSize = 10
350 Printer.FontBold = False
360 Printer.FontName = "Arial Greek"
        
370 Printer.CurrentX = F_ET_PER_X  '  200 + F_ArPerEtik * 100
380 Printer.CurrentY = F_ET_PER_Y '600
390 Printer.Print MONO   'perigrafh

400 Printer.CurrentX = F_ET_KOD_X '  F_ArPerEtik * 100
410 Printer.CurrentY = F_ET_KOD_Y '600
420 Printer.Print MKOD ' Text1(0).Text    'kodikos

    ' Printer.CurrentX = 2048
    'Printer.Print s
        
430 Printer.CurrentX = 400 + F_ArPerEtik * 100
    'Printer.FontSize = 12
440 Printer.FontSize = 10

450 Printer.CurrentX = F_ET_BAR_X  ' 200 + F_ArPerEtik * 100
460 Printer.CurrentY = F_ET_BAR_Y ' 600

470 Printer.PaintPicture BarCode3.Picture, F_ET_BAR_X, F_ET_BAR_Y

    Printer.FontSize = F_ET_EPIX_FONTSIZE
    Printer.CurrentX = F_ET_EPIX_X
    Printer.CurrentY = F_ET_EPIX_Y
    Printer.Print F_Epix;

480 Printer.CurrentY = F_ET_TIM_Y '1000 + F_ArPerEtik * 100
        
490 Printer.FontSize = 18
500 Printer.FontBold = True
        
510 Printer.CurrentX = F_ET_PER_X
        
530 Printer.FontSize = 24
540 Printer.CurrentX = F_ET_TIM_X '1000 + F_ArPerEtik * 100
550 Printer.Print Format(mLTI5, "##0.00") + " €"    ' timh
    ' Printer.Print "4.55 €" ' timh
552 Printer.EndDoc
    'Next




Next









'Close #1





    Exit Sub

End Sub  ' polla

