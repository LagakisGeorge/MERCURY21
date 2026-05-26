VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FF0000&
   Caption         =   "oe"
   ClientHeight    =   8736
   ClientLeft      =   168
   ClientTop       =   -3972
   ClientWidth     =   15144
   Icon            =   "MDIFORM13.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFORM13.frx":0442
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   8088
      Left            =   0
      ScaleHeight     =   8088
      ScaleWidth      =   3000
      TabIndex        =   2
      Top             =   336
      Width           =   3000
      Begin VB.CommandButton cmdset 
         BackColor       =   &H80000010&
         Caption         =   "Ρυθμίσεις Συστήματος"
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5640
         Width           =   3015
      End
      Begin VB.CommandButton cmdfile 
         BackColor       =   &H80000010&
         Caption         =   "Γρήγορη Πρόσβαση"
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   3015
      End
      Begin VB.PictureBox picSet 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   34
         ScaleHeight     =   3492
         ScaleWidth      =   2940
         TabIndex        =   6
         Top             =   6015
         Width           =   2940
         Begin MSComctlLib.ListView ListView2 
            Height          =   4500
            Left            =   120
            TabIndex        =   12
            Top             =   0
            Width           =   2748
            _ExtentX        =   4847
            _ExtentY        =   7938
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDragMode     =   1
            _Version        =   393217
            Icons           =   "i32x32"
            SmallIcons      =   "i32x32"
            ForeColor       =   -2147483641
            BackColor       =   -2147483634
            Appearance      =   0
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "MDIFORM13.frx":1F6FC8
            OLEDragMode     =   1
            NumItems        =   0
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   33
            TabIndex        =   7
            Top             =   180
            Width           =   2055
         End
      End
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   34
         ScaleHeight     =   4572
         ScaleWidth      =   2940
         TabIndex        =   3
         Top             =   1080
         Width           =   2940
         Begin MSComctlLib.ListView Listview1 
            Height          =   4500
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2988
            _ExtentX        =   5271
            _ExtentY        =   7938
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDragMode     =   1
            _Version        =   393217
            Icons           =   "i32x32"
            SmallIcons      =   "i32x32"
            ForeColor       =   -2147483641
            BackColor       =   -2147483634
            Appearance      =   0
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "MDIFORM13.frx":1F78A2
            OLEDragMode     =   1
            NumItems        =   0
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   33
            TabIndex        =   5
            Top             =   180
            Width           =   2055
         End
      End
      Begin VB.Image Image11 
         Height          =   23145
         Left            =   3000
         Picture         =   "MDIFORM13.frx":1F817C
         Stretch         =   -1  'True
         Top             =   -1560
         Width           =   570
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   0
         Picture         =   "MDIFORM13.frx":1F8374
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Σήμερα είναι :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   168
         Left            =   48
         TabIndex        =   11
         Top             =   468
         Width           =   888
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   750
         TabIndex        =   10
         Top             =   450
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Καλώς ήρθατε ,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   168
         Index           =   0
         Left            =   48
         TabIndex        =   9
         Top             =   252
         Width           =   1008
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   750
         TabIndex        =   8
         Top             =   240
         Width           =   180
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   0
      Top             =   8424
      Width           =   15144
      _ExtentX        =   26712
      _ExtentY        =   550
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   9
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2893
            MinWidth        =   2893
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "GDIR"
            TextSave        =   "GDIR"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3175
            MinWidth        =   3175
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3492
            MinWidth        =   3492
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   ".."
            TextSave        =   ".."
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15144
      _ExtentX        =   26712
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      HelpContextID   =   65278
      ImageList       =   "imlToolbarPics"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Table"
            Object.ToolTipText     =   "Table type Recordset"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dynaset"
            Object.ToolTipText     =   "Dynaset type Recordset"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Schedule"
            Object.ToolTipText     =   "Ημερολόγιο"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BIBLIO"
            ImageIndex      =   8
         EndProperty
      EndProperty
      Begin MSAdodcLib.Adodc TELEFON 
         Height          =   330
         Left            =   2880
         Top             =   30
         Visible         =   0   'False
         Width           =   2640
         _ExtentX        =   4657
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
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   7620
         Top             =   -30
      End
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   -210
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FilterIndex     =   1144
   End
   Begin MSComctlLib.ImageList imlToolbarPics 
      Left            =   1605
      Top             =   2295
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F856C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F867E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F8790
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F88A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F89B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F8AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F8BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F8CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F8DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F8F0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   0
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F9020
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1F9CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1FA9D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1FB6AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1FC388
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1FD062
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1FDD3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1FEA16
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":1FF6F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":2003CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":2010A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":201D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":202A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":203732
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":20440C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":2050E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":205DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":206A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2160
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":207774
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":20844E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":209128
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":209E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":20AADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":20B7B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":20C490
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":20D16A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":20DE44
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":20EB1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":20F7F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":2104D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":2111AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":211E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":212B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":21383A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":214514
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":2151EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":21547D
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":215B2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":216886
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":2168E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":216942
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":219E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM13.frx":21A6F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu pelat 
      Caption         =   "&Πελάτες/Προμηθευτές"
      Begin VB.Menu mpelat1 
         Caption         =   "&1.Εισαγωγή νέου"
         Shortcut        =   ^P
      End
      Begin VB.Menu mpelat2 
         Caption         =   "&2.Διόρθωση/Διαγραφή"
      End
      Begin VB.Menu mpelat3 
         Caption         =   "&3.Παρουσίαση"
      End
      Begin VB.Menu mpelat4 
         Caption         =   "&4.Καρτέλλα"
         Shortcut        =   ^K
      End
      Begin VB.Menu mpelat5 
         Caption         =   "&5.Παρουσίαση Κινήσεων"
      End
      Begin VB.Menu exit 
         Caption         =   "&6.Εξοδος"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu apot 
      Caption         =   "&Αποθήκη"
      Begin VB.Menu mapot1 
         Caption         =   "&1.Εισαγωγή νέου είδους"
         Index           =   21
         Shortcut        =   ^E
      End
      Begin VB.Menu mapot2 
         Caption         =   "&2.Διόρθωση είδους"
         Index           =   22
         Shortcut        =   ^D
      End
      Begin VB.Menu pareid 
         Caption         =   "&3.Παρουσίαση ειδών"
         Index           =   23
         Begin VB.Menu mapot31 
            Caption         =   "Είδη Αποθήκης"
            Index           =   310
         End
         Begin VB.Menu mapot32 
            Caption         =   "Υπόλοιπα Αποθηκών"
         End
         Begin VB.Menu mapot33 
            Caption         =   "Ισοζύγιο Αποθήκης"
         End
         Begin VB.Menu mapot34_Partides 
            Caption         =   "Παρτίδες"
         End
      End
      Begin VB.Menu apo4 
         Caption         =   "&4.Τιμοκατάλογος Ειδών"
         Index           =   24
      End
      Begin VB.Menu apo5 
         Caption         =   "&5.Κίνηση Ειδών"
         Index           =   25
      End
      Begin VB.Menu mapot6 
         Caption         =   "&6.Οικογένειες-Υποοικογένειες"
         Index           =   26
      End
      Begin VB.Menu mapot7 
         Caption         =   "&7.Ρυθμιζόμενες εκτυπώσεις"
         Index           =   27
      End
      Begin VB.Menu apo2scroll 
         Caption         =   "&8.Επεξεργασία Είδους (web)"
         Index           =   28
         Visible         =   0   'False
      End
      Begin VB.Menu apot_anaz 
         Caption         =   "&9.Αναζήτηση Είδους"
         Index           =   29
         Visible         =   0   'False
      End
   End
   Begin VB.Menu epit 
      Caption         =   "&Επιταγές"
      Begin VB.Menu eisepit 
         Caption         =   "&1.Eισαγωγή Επιταγής"
      End
      Begin VB.Menu diorepit 
         Caption         =   "&2.Στατιστικά"
      End
      Begin VB.Menu parepit 
         Caption         =   "&3.Παρουσίαση Επιταγών"
      End
   End
   Begin VB.Menu Par 
      Caption         =   "Παρασ&τατικά"
      Begin VB.Menu mpar1 
         Caption         =   "&1.Εισαγωγή Παρασ/κού"
         Index           =   41
         Shortcut        =   ^T
      End
      Begin VB.Menu diagrpar 
         Caption         =   "&2.Προβολή Παραστατικών κατά ημέρα"
         Index           =   42
      End
      Begin VB.Menu parenos 
         Caption         =   "&3.Παρουσίαση ενός Παραστατικού"
         Index           =   43
      End
      Begin VB.Menu timspar 
         Caption         =   "&4.Παρουσίαση Παραστατικών"
         Index           =   44
      End
      Begin VB.Menu mpar5 
         Caption         =   "&5.Εικόνα Επιχείρησης"
      End
      Begin VB.Menu mPar6 
         Caption         =   "&6.Ταμείο ημέρας"
      End
      Begin VB.Menu mpar7 
         Caption         =   "&7.Μαζικός μετασχηματισμός"
      End
      Begin VB.Menu mpar8 
         Caption         =   "&8.Εισαγωγή Δελτίων Ποσοτικής Παραλαβής"
      End
      Begin VB.Menu mpar9 
         Caption         =   "9.Σύνολα φόρτωσης"
      End
   End
   Begin VB.Menu bohu 
      Caption         =   "&Βοηθητικά"
      Begin VB.Menu mBohu1 
         Caption         =   "&1.Πίνακες"
         Index           =   51
      End
      Begin VB.Menu mbohu2 
         Caption         =   "&2.Παραμετροι Προγράμματος"
         Index           =   52
      End
      Begin VB.Menu par3 
         Caption         =   "&3.Παράμετροι Παραστατικών"
         Index           =   53
      End
      Begin VB.Menu parkin 
         Caption         =   "&4.Παράμετροι κινήσεων πελατών"
         Index           =   54
      End
      Begin VB.Menu pros 
         Caption         =   "&5.Import-Export"
         Index           =   55
      End
      Begin VB.Menu genlog 
         Caption         =   "&6.Ημερολόγιο Γεν.Λογιστικής"
         Index           =   56
      End
      Begin VB.Menu DBF 
         Caption         =   "&7.Αποστολή σε Φορητά Τερματικά"
         Index           =   57
      End
      Begin VB.Menu pontoif 
         Caption         =   "&8.Μεταφορές Αρχείων-Πόντοι"
         Index           =   58
      End
      Begin VB.Menu apoemporiko 
         Caption         =   "&9.Aντίγραφα / Επαναφορά"
         Index           =   59
      End
      Begin VB.Menu calculator 
         Caption         =   "10.Αριθμομηχανή"
         Index           =   510
         Shortcut        =   {F2}
      End
      Begin VB.Menu mbohu11 
         Caption         =   "11.Αλλαγή κωδικών χρήστη"
      End
      Begin VB.Menu mbohu12 
         Caption         =   "12.Συγκεντρωτική Τιμολογίων"
      End
      Begin VB.Menu mBohu13 
         Caption         =   "13.Ελεγχοι πληρότητας αρχείων"
      End
   End
   Begin VB.Menu helpform 
      Caption         =   "Βοήθεια"
      Begin VB.Menu man 
         Caption         =   "Βιβλίο οδηγιών"
      End
      Begin VB.Menu ekdoseis 
         Caption         =   "Εκδόσεις"
      End
   End
   Begin VB.Menu RUNTIME 
      Caption         =   "RUNTIMEMENU"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu runt 
         Caption         =   "RUNT"
         Index           =   0
      End
      Begin VB.Menu runt2 
         Caption         =   "runt2"
      End
   End
   Begin VB.Menu par1menu 
      Caption         =   "Par1"
      Visible         =   0   'False
      Begin VB.Menu par2show 
         Caption         =   "Παρουσίαση Παρ/κού"
      End
      Begin VB.Menu diortosis 
         Caption         =   "Διόρθωση"
      End
      Begin VB.Menu ektyp_par 
         Caption         =   "Εκτύπωση Παρ/κού"
      End
      Begin VB.Menu metasx_par 
         Caption         =   "Μετασχηματισμός Παρ/κού"
      End
      Begin VB.Menu metasx_lian 
         Caption         =   "Μετασχηματισμός σε ΑΠ.ΛΙΑΝ"
      End
      Begin VB.Menu METASX_TIM 
         Caption         =   "Μετασχηματισμός σε ΤΙΜ.ΔΑ"
      End
      Begin VB.Menu METASX_DA 
         Caption         =   "Μετασχηματισμός σε ΔΕΛ.ΑΠ"
      End
   End
   Begin VB.Menu EPITAG311 
      Caption         =   "EPITAG311"
      Visible         =   0   'False
      Begin VB.Menu epit311 
         Caption         =   "epit311"
         Index           =   0
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'c5450
Option Explicit
Public fkrithria As New KRITHRIA
Dim ETAIR(20), mc As Integer    ' POPUP MENU
Attribute mc.VB_VarUserMemId = 1073938432

Dim F_TOOLBAR As Long

Dim f_Ypenth  As Long

Dim f_supervisor_code
Attribute f_supervisor_code.VB_VarUserMemId = 1073938436

Dim f_telephone  As Integer

Dim F_BACKGROUND As Integer

Dim f_FARDOSIMAGE

Dim bhide     As Boolean

Dim getAction As String

Dim f_site    As Long

Private Sub apo2scroll_Click(Index As Integer)

    On Error Resume Next

    If f_site = 333 Then
        apot2ROULISnew.SHOW
    ElseIf f_site = 334 Then
        apot2ROULIS.SHOW
    Else
        apot2vmn.SHOW
    End If

    'ElseIf f_site = 222 Then
    ' apot2scroll.SHOW
    'End If

End Sub

'Private Sub antigrafa_Click(Index As Integer)
''   BOHU8.Show
'
'End Sub

'Private Sub Adodc1_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
'End Sub

'Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
'End Sub

Private Sub apo4_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo apo4_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     apot4.SHOW

        '<EhFooter>
        Exit Sub

apo4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.apo4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.apo4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub apo5_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo apo5_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     APOT5.SHOW

        '<EhFooter>
        Exit Sub

apo5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.apo5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.apo5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub apoemporiko_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo apoemporiko_Click_Err

        '</EhHeader>

100     gApoMenu = True

110     bohu8.SHOW

        '<EhFooter>
        Exit Sub

apoemporiko_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.apoemporiko_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.apoemporiko_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub apot2VirtueMart_Click()
    '    apot2vm.SHOW
  
End Sub

Private Sub apot_anaz_Click(Index As Integer)
    apot2ROULISSEARCH.SHOW
  
    'apot_searching.SHOW

End Sub

Private Sub calculator_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo calculator_Click_Err

        '</EhHeader>
        Dim DUM

100     DUM = Shell("calc.exe", vbMaximizedFocus)

        '<EhFooter>
        Exit Sub

calculator_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.calculator_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.calculator_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Data1_Validate(Action As Integer, Save As Integer)
'
'End Sub

Private Sub cmdFile_Click()

        '<EhHeader>
        On Error GoTo cmdFile_Click_Err

        '</EhHeader>
       
100     getAction = "picFile"
    
110     If picMenu.Visible = True Then
120         Timer1.Interval = 10
130         bhide = True
        Else
140         Timer1.Interval = 10
150         bhide = False
160         picMenu.Visible = True
        End If

        '<EhFooter>
        Exit Sub

cmdFile_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.cmdFile_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.cmdFile_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdSet_Click()

        '<EhHeader>
        On Error GoTo cmdSet_Click_Err

        '</EhHeader>

100     getAction = "picSet"
    
110     If picSet.Visible = True Then
120         Timer1.Interval = 10
130         bhide = True
        Else
140         Timer1.Interval = 10
150         bhide = False
160         picSet.Visible = True
        End If

        '<EhFooter>
        Exit Sub

cmdSet_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.cmdSet_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.cmdSet_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DBF_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo DBF_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     bohu6.SHOW

        '<EhFooter>
        Exit Sub

DBF_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.DBF_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.DBF_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub diagrpar_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo diagrpar_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     par2frmOrder.SHOW

        '<EhFooter>
        Exit Sub

diagrpar_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.diagrpar_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.diagrpar_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub diortosis_Click()
        
    '<EhHeader>
    ' On Error GoTo metasx_LIAN_Click_Err
    '</EhHeader>
        
    Dim m As Integer

    Dim C As String
         
    PAR1.TDBGrid1.Col = 1
    C = Replace(PAR1.TDBGrid1.Text, ",", ".")
         
         Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)
       
        
         
         
    ' METASX_DIORTOSIS
    PAR1.METASX_SE_TIM 99, C, m_ID
         
    Exit Sub
        
    '------------------- OLD  ==========================================================================================
        
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
    '
    '        '<EhHeader>
    '        On Error GoTo diortosis_Click_Err
    '        '</EhHeader>
    '   PAR1.diortosis
    '
    '     If Len(PAR1.Grid1.TextMatrix(1, 1)) = 0 Then
    '
    '         MsgBox "OK"
    '         PAR1.diortosis
    '
    '        End If
    '
    '
    '
    '
    '        '<EhFooter>
    '        Exit Sub
    '
    'diortosis_Click_Err:
    '        'MsgBox Err.Description & vbCrLf & _
    '               "in ADOMERCNEW.MDIForm1.diortosis_Click " & _
    '               "at line " & Erl, _
    '               vbExclamation + vbOKOnly, "Application Error"
    '        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.diortosis_Click " & "at line " & Erl
    '        Resume Next
    '        '</EhFooter>
End Sub

Private Sub eisepit_Click()

        '<EhHeader>
        On Error GoTo eisepit_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     EPIT1.SHOW

        '<EhFooter>
        Exit Sub

eisepit_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.eisepit_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.eisepit_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub ekdoseis_Click()

        '<EhHeader>
        On Error GoTo ekdoseis_Click_Err

        '</EhHeader>

100     help1.WindowState = 2   'maximized

110     help1.SHOW

        '<EhFooter>
        Exit Sub

ekdoseis_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.ekdoseis_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.ekdoseis_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub ektyp_par_Click()

        '<EhHeader>
        On Error GoTo ektyp_par_Click_Err

        '</EhHeader>

        Dim m As Integer

        Dim C As String

        PAR1.TDBGrid1.Col = 2
        C = Replace(PAR1.TDBGrid1.Text, ",", ".")
        PAR1.METASX_SE_TIM 99, C, 0

        ' 100     PAR1.diortosis
110     DoEvents

120     PAR1.Command5_Click

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

Private Sub epit311_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo epit311_Click_Err

        '</EhHeader>

100     epit3.epit311_Click Index

        '<EhFooter>
        Exit Sub

epit311_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.epit311_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.epit311_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub exit_Click()

        '<EhHeader>
        On Error GoTo exit_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     Unload Me

        '<EhFooter>
        Exit Sub

exit_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.exit_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.exit_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Form_Load()
'
'End Sub

Private Sub genlog_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo genlog_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     bohu5.SHOW

        '<EhFooter>
        Exit Sub

genlog_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.genlog_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.genlog_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Image11_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

        '   Set Resizer.BoundControl = Image11
        '<EhHeader>
        On Error GoTo Image11_MouseDown_Err

        '</EhHeader>

100     Me.Caption = "image11"
110     Image11.Move Image11.Left, Image11.Top, X, Image11.Height
    
        '<EhFooter>
        Exit Sub

Image11_MouseDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.Image11_MouseDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.Image11_MouseDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Image1_Click()

'   gApoMenu = True
'  HMEROL2.SHOW
'End Sub

'Private Sub Image2_Click()

'  gApoMenu = True
'   hmerol.SHOW

'End Sub

Private Sub helpform_Click()
    ''  helpforma.SHOW
    '
    '
    '  Dim sFilename As String
    'Dim appWord As Word.Application
    'Dim WRDdoc As Word.Document
    '
    'sFilename = "c:\mercvb\help\help.doc"
    'Set appWord = New Word.Application
    ''Set WRDdoc = appWord.Documents.Open(sFilename)
    ''appWord.Visible = True

End Sub

Private Sub Listview2_Click()

        '    Listview2.ListItems.Add , "a6", "Πίνακες", 16, 16
        '    Listview2.ListItems.Add , "a7", "Back-Up", 17, 17
        '<EhHeader>
        On Error GoTo Listview2_Click_Err

        '</EhHeader>
100     Select Case ListView2.SelectedItem.Key

            Case "b1":
110             BOHU1.SHOW

            Case "b2":
120             bohu8.SHOW
        End Select

        '<EhFooter>
        Exit Sub

Listview2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.Listview2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.Listview2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub man_Click()

        '<EhHeader>
        On Error GoTo man_Click_Err

        '</EhHeader>

        Dim sFileName As String

        Dim appWord   As Word.Application

        Dim WRDdoc    As Word.Document

100     sFileName = "c:\mercvb\help\help.doc"
110     Set appWord = New Word.Application
120     Set WRDdoc = appWord.Documents.Open(sFileName)
130     appWord.Visible = True

        '<EhFooter>
        Exit Sub

man_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.man_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.man_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mapot1_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo mapot1_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     APOT1.SHOW

        '<EhFooter>
        Exit Sub

mapot1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mapot1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mapot1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mapot2_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo mapot2_Click_Err

        '</EhHeader>

100     gApoMenu = True

        '
        ' MDIForm1.'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", "NormalOffice2007.ini"
        ' MDIForm1.'SkinFramework1.ApplyWindow APOT2.hWnd
        ' MDIForm1.'SkinFramework1.ApplyOptions = MDIForm1.'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
        '
110     APOT2.WindowState = 2
120     APOT2.SHOW

        '  'SkinFramework1.LoadSkin "", ""

        '<EhFooter>
        Exit Sub

mapot2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mapot2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mapot2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mapot31_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo mapot31_Click_Err

        '</EhHeader>

100     apot3.SHOW

        '<EhFooter>
        Exit Sub

mapot31_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mapot31_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mapot31_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mapot32_Click()
'ΥΠΟΛΟΙΠΑ ΑΠΟΘΗΚΩΝ
Dim FF As New apot3param
 FF.Label_titlos = "Υπόλοιπα Ειδών"

FF.load_sthles_EIDON
'FF.LOAD_KINHSEIS_EIDON
FF.Rep_Fields.Text = "select KOD as [Κωδικός],LEFT(ONO,35) as [Περιγραφή],POS01 AS [Υπόλοιπο]"
FF.gemisma_Krithrion
fkrithria.Grid1.TextMatrix(4, 1) = 0.1
fkrithria.Grid1.TextMatrix(4, 2) = 999999999
fkrithria.ΕΧΙΤ_Click
FF.sumes.Text = "001"
FF.WHERE_QUERY.Text = "POS>=0.1 AND POS<=999999999" ' StatusBar1.Panels(6).Text"
'YPOLOGISMOS_Click
FF.SHOW
'FF.YPOLOGISMOS_Click

End Sub

Private Sub mapot34_Click()




End Sub

Private Sub mapot33_Click()
'isozygio ΑΠΟΘΗΚΩΝ
Dim FF As New apot3param
Dim sql As String

FF.load_sthles_EIDON
'FF.LOAD_KINHSEIS_EIDON
sql = "select KOD as [Κωδικός],LEFT(ONO,35) as [Περιγραφή],"
sql = sql + "(CASE APOG01 WHEN NULL THEN 0 ELSE APOG01 END) AS [Απογ.Απ1],(CASE KOSAPOG01 WHEN NULL THEN 0 ELSE KOSAPOG01 END) AS [Αξ.Περ.Απογ.Απ1],"
sql = sql + "(CASE S1X WHEN NULL THEN 0 ELSE S1X-APOG01 END) AS [Εισ1.Περιόδ],(CASE AS1X WHEN NULL THEN 0 ELSE AS1X-KOSAPOG01 END)  AS [Αξ1.Αγ.Περιόδ],"
sql = sql + "(CASE S1P WHEN NULL THEN 0 ELSE S1P END)  AS [Εξ1.Περιόδ],(CASE AS1P WHEN NULL THEN 0 ELSE AS1P END)  AS [Αξ1.Εξ.Περιόδ],"
sql = sql + "XPOS01 as [Υπ.Απ1],XTI*(XPOS01) AS [Αξ.Απογρ.απ.1]"
FF.sumes.Text = "0011111111"
FF.Rep_Fields.Text = sql
FF.gemisma_Krithrion
fkrithria.Grid1.TextMatrix(4, 1) = 0.1
fkrithria.Grid1.TextMatrix(4, 2) = 999999999
fkrithria.ΕΧΙΤ_Click
FF.WHERE_QUERY.Text = "S1X+S1P+APOG01>=0 "
FF.Label_titlos = "Ισοζύγιο Αποθηκών"
FF.SHOW

End Sub



Private Sub mapot34_Partides_Click()
  'frmPARTIDES.SHOW
'================================================================================
        Dim f As New DataEntry3
        Dim R As New ADODB.Recordset
        Load f

f.Caption = "Παρτίδες"

   ' string adodc1
'
'130     UpdSQL = ParamGrid.TextMatrix(0, 4)    ' string adodc1
'
'140     fAddNewScript = ParamGrid.TextMatrix(0, 5)    ' string meta to addnew button
'
'150     F_FIELD_EPIPLEON = ParamGrid.TextMatrix(0, 6)    'EPIPLEON FIELD
'160     F_TIMH = ParamGrid.TextMatrix(0, 7)  ' TIMH TOY EPIPLEON FIELD
'
'170     F_2FIELD_EPIPLEON = ParamGrid.TextMatrix(0, 8)    'EPIPLEON FIELD
'180     F_2TIMH = ParamGrid.TextMatrix(0, 9)  ' TIMH TOY EPIPLEON FIELD
'
'        'f.ParamGrid.TextMatrix(0, 8) = "TIMH" '2 EPIPLEON FIELD
'        'f.ParamGrid.TextMatrix(0, 9) = Left(Combo4.Text, 2)  ' TIMH TOY 2 EPIPLEON FIELD


'100     sql = ParamGrid.TextMatrix(0, 1)
        f.ParamGrid.TextMatrix(0, 1) = "SELECT  * FROM PARTIDES " '    sql

        '------------------------------------------------------------------------------------
        'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
        f.ParamGrid.TextMatrix(0, 2) = "SELECT ID,KOD,LOTNUMBER,HMEL,POSOAG " & " FROM PARTIDES ORDER BY ID DESC"    'ParamGrid   sql2
'110     Sql2 = ParamGrid.TextMatrix(0, 2)    ' string adodc1
        '------------------------------------------------------------------------------------
        'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
        f.ParamGrid.TextMatrix(0, 3) = "delete from PARTIDES where  "
        
         '   "  ""DELETE FROM PARTIDES WHERE LOTNUMBER ='"" + t02 +""'  AND KOD='""+t11+""' "    'DELETE SQL
        
        '"delete from PARTIDES where LOTNUMBER='" + T02 + " AND "    'ParamGrid              DelSQL
'120     delSql = ParamGrid.TextMatrix(0, 3)
        '------------------------------------------------------------------------------------
        'ποιό update query θα δημιουργείται (προαιρετικά )
        'αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
        '                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
        ' DD = " UPDATE PINAKES SET TYPOS=" + Str(MTYPOS) + " where TYPOS IS NULL AND AYJON="" + t1   "
        ' AddNewScript
        ' ' DD = "Sub Main()" & vbCrLf & _
        '         " DIM mCon  " & vbCrLf & _
        '         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
        '         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
        '         " mCon.execute """ + DD + " " & vbCrLf & _
        '        "Set r=Nothing" & vbCrLf & _
        '        "Set mCon=Nothing" & vbCrLf & _
        '        "End Sub"

'150     f.ParamGrid.TextMatrix(0, 6) = "TYPOS"    'EPIPLEON FIELD
'160     f.ParamGrid.TextMatrix(0, 7) = MTYPOS    ' TIMH TOY EPIPLEON FIELD

        '------------------------------------------------------------------------------------
170     f.ParamGrid.TextMatrix(0, 4) = ""  'DD

        f.ParamGrid.TextMatrix(1, 1) = "ID εγγραφής"
        f.ParamGrid.TextMatrix(1, 2) = 1  ' ARITHMOS
        
        
180     f.ParamGrid.TextMatrix(2, 1) = "Κωδ.Είδους"
        f.ParamGrid.TextMatrix(2, 2) = 2 ' STRING


190     f.ParamGrid.TextMatrix(3, 1) = "LOT NUMBER"
        f.ParamGrid.TextMatrix(3, 2) = 2 ' STRING
        
        
        f.ParamGrid.TextMatrix(4, 1) = "Ημερ.Λήξης"
        f.ParamGrid.TextMatrix(4, 2) = 3 ' DATE
        
        f.ParamGrid.TextMatrix(5, 1) = "Ποσ.Αγοράς"
        f.ParamGrid.TextMatrix(5, 2) = 1 ' STRING



        'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
200     f.Labels(0).Width = f.Labels(0).Width + 1000    '* TextWidth("A")
210     f.Textboxes(0).Left = f.Textboxes(0).Left + 1000    ' * TextWidth("A")
        '------------------------------------------------------------------------------------
Dim OO
        'αν παει η φόρμα να φύγει από την οθόνη τότε να κεντραρεται δεξια με την mdiform1
220    ' If OO.Left + f.Width > MDIForm1.Left + MDIForm1.Width Then
230     '    f.Left = MDIForm1.Left + MDIForm1.Width - f.Width
       ' Else
240    '     f.Left = OO.Left
      '  End If

250  '   f.Top = OO.Top + MDIForm1.Top

260     f.SHOW 1

Gdb.Execute "update PARTIDES SET YPOLOIPO=POSOAG WHERE YPOLOIPO IS NULL"
  

End Sub

'Private Sub mapot32_Click(Index As Integer)
''apot32.SHOW
'
'End Sub

Private Sub mapot6_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo mapot6_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     APOT6.SHOW

        '<EhFooter>
        Exit Sub

mapot6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mapot6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mapot6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mapot7_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo mapot7_Click_Err

        '</EhHeader>
100     If gXEIRISTHS <= 2 Then
110         apot7.Dir1.Visible = False
120         apot7.Text2.Visible = False
130         apot7.apo.Visible = False
140         apot7.eos.Visible = False
150         apot7.Label1.Visible = False
160         apot7.Label2.Visible = False
        Else
170         apot7.SHOW
        End If

        '<EhFooter>
        Exit Sub

mapot7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mapot7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mapot7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mbohu1_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo mbohu1_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     BOHU1.SHOW

        '<EhFooter>
        Exit Sub

mbohu1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mbohu1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mbohu1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mbohu11_Click()

        '<EhHeader>
        On Error GoTo mbohu11_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     bohu11.SHOW

        '<EhFooter>
        Exit Sub

mbohu11_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mbohu11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mbohu11_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mbohu12_Click()

        '<EhHeader>
        On Error GoTo mbohu12_Click_Err

        '</EhHeader>

100     bohu12.SHOW

        '<EhFooter>
        Exit Sub

mbohu12_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mbohu12_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mbohu12_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mBohu13_Click()

        '<EhHeader>
        On Error GoTo mBohu13_Click_Err

        '</EhHeader>

100     BOHU13.SHOW

        '<EhFooter>
        Exit Sub

mBohu13_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mBohu13_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mBohu13_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mbohu2_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo mbohu2_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     BOHU2.SHOW

        '<EhFooter>
        Exit Sub

mbohu2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mbohu2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mbohu2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub MDIForm_DblClick()

        '<EhHeader>
        On Error GoTo MDIForm_DblClick_Err

        '</EhHeader>
        Dim a As Long, X

100     If gXEIRISTHS = 9 Then
110         f_supervisor_code = FINDPARAMETROI("MDIFORM1", "f_supervisor_code", "0", "Κωδικός")
120         X = f_supervisor_code
130         PARAMETROI.PARAM.Caption = "MDIFORM1"
140         PARAMETROI.SHOW 1
150         f_supervisor_code = FINDPARAMETROI("MDIFORM1", "f_supervisor_code", "0", "Κωδικός")
            
160         If f_supervisor_code = X Then
            Else

170             If IsNull(f_supervisor_code) Then f_supervisor_code = ""
180             Gdb.Execute "UPDATE DIK SET CODE='" + f_supervisor_code + "' WHERE  LEFT(KOD,6)='070109' ", a

190             If a > 0 Then
200                 MsgBox "ενημερώθηκε ο νέος κωδικός"
                End If
            End If

            '   If F_BACKGROUND > 0 Then
            '      On Error GoTo ERRORS
            '      Kill "C:\MERCVB\BACKGROUND.BMP"
            '      FileCopy "C:\MERCVB\BACKGROUND" + LTrim(Str(F_BACKGROUND)) + ".BMP", "C:\MERCVB\BACKGROUND.BMP"
            '      MILSEC 2000
            '      gPicture = "C:\MERCVB\BACKGROUND.BMP"
            '      MDIForm1.Picture = LoadPicture(gPicture)
            '
            '   End If

        End If

        Exit Sub

ERRORS:

210     MsgBox "Δεν έγινε εγκατάσταση " + Chr(13) + Err.Description

220     Resume Next

        '<EhFooter>
        Exit Sub

MDIForm_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.MDIForm_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.MDIForm_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub MDIForm_Load()

'   ' Instantiate the Toolbox
'    Set g_InteropToolbox = New InteropToolbox
'    g_InteropToolbox.Initialize
'
'    ' Call Initialize method only when first creating the toolbox
'    ' This aids in the debugging experience
'    g_InteropToolbox.Initialize
'
'    ' Signal Application Startup
'    g_InteropToolbox.EventMessenger.RaiseApplicationStartedupEvent
'
'






        ' Global Gdb As New ADODB.Connection
        ' UGridEdit.SHOW 1
        ' End
        '<EhHeader>
        On Error GoTo MDIForm_Load_Err

        '</EhHeader>

        Dim R As New ADODB.Recordset

100     g_Stop = 0
110     gCapitals = 1

        'Dim ARR()  As String
        ' ARR = Split("1,2,3,4", ",")

        Dim DUM

        'Resizer.ZOrder 0
        'Set Resizer.BoundControl = Text1(Index)

        'Image11.width = f_FARDOSIMAGE
        'Listview1.width = f_FARDOSIMAGE
        'Listview2.width = f_FARDOSIMAGE

        ''πως ανοιγω ενα exel αρχείο
        ' Dim xl As New excel.Application
        ' Dim xx As Worksheet
        ' openXLS "c:\apog.xls", xx, xl
        ' Print xx.cells(1, 1)
        'Set xx = Nothing
        'xl.Quit
        'Set xl = Nothing

120     gDIG_3 = 6
130     gDIG_4 = 8

        'DUM = ShellExecute(0&, "Print", "C:\PDF\PDFTEMP\*.PDF", "", "", vbNormalFocus)

        Dim wrkDefault, mydb, minima

        '   '   'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        '    'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        '    'SkinFramework1.ApplyWindow Me.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
        'SkinFramework1.LoadSkin "", ""

140     gPicture = "C:\MERCVB\background.BMP"
150     MDIForm1.Picture = LoadPicture(gPicture)

        ' Dim dum: dum = printing_etik("354454534534", 2, "tem", 1.31, 1.34, 1.23, 9, "lagakis", 67)

160     MDIForm1.SHOW

        On Error GoTo SEELINE

175     minima = "ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΑΡΧΕΙΟ C:\MERCPATH.TXT"

        If Len(Dir("C:\MERCPATH.TXT")) > 0 Then

180         Open "C:\MERCPATH.TXT" For Input As #1
190         Line Input #1, gDir
200         Line Input #1, gConnect
210         Close #1
        Else

211         Open "C:\MERCVB\MERCPATH.TXT" For Input As #1
            Line Input #1, gDir
            Line Input #1, gConnect
212         Close #1

        End If

        ' ok !!!!!  gConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=38983;Initial Catalog=MERCURY"

        '50 Close #1

        '
        'Dim dsoServer As DSO.Server
        '    Dim enuClassType As DSO.ClassTypes
        '
        '    ' Create instance of server and connect.
        '    ' "LocalHost" will default to the
        '    ' local Windows NT Server 4.0 where the
        '    ' Analysis server is installed.
        '    Set dsoServer = New DSO.Server
        '    dsoServer.Connect "LocalHost"
        '
        'Set oDatabase = dsoServer.Databases("Northwind", "stevenb")

220     MDIForm1.StatusBar1.Panels(1).Text = gDir
230     MDIForm1.StatusBar1.Panels(3).Text = "Mercury ver." + Format(App.Major, "00") + "." + Format(App.Minor, "00") + "." + Format(App.Revision, "00")
235     MDIForm1.StatusBar1.Panels(6).Text = "Lagakis +30 2521022225"

        'Exit Sub

240     pelat.Enabled = False
250     apot.Enabled = False
260     epit.Enabled = False
270     Par.Enabled = False
280     bohu.Enabled = False

        ' gConnect = "Provider=SQLOLEDB.1;Password=38983;Persist Security Info=True;User ID=sa;Data Source=GEORGE;Initial Catalog=MERCURY;"
290     frmLogin.SHOW 1
 
        f_site = Val(FINDPARAMETROI("MDIFORM1", "F_SITE", "0", "ΚΩΔΙΚΟΣ SITE 1=VM 2=TOYS"))

        If f_site = 111 Or f_site = 222 Or f_site = 333 Or f_site = 444 Then
            apo2scroll(28).Visible = True
            apot_anaz(29).Visible = True
        End If

295     MDIForm1.StatusBar1.Panels(7).Text = "SN:" + FINDPARAMETROI("SN", "SN", "0", "ΣEIRIAKOΣ ΑΡΙΘΜΟΣ")

300     Timer1.Enabled = True

310     If gConnect = "end" Then
320         Unload Me

330         End

        End If

340     F_BACKGROUND = Val(FINDPARAMETROI("MDIFORM1", "F_BACKGROUND", "1", "BACKGROUND ΑΡ.ΕΙΚΟΝΑΣ=1,2,3,.. ΟΧΙ=0"))

350     f_FARDOSIMAGE = Val(FINDPARAMETROI("MDIFORM1", "f_FARDOSIMAGE", "3040", "ΠΛΑΤΟΣ ΓΡΗΓΟΡΟΥ ΜΕΝΟΥ(ΑΡΧΙΚΟ=3040)"))

360     Picture1.Width = f_FARDOSIMAGE
        'If f_FARDOSIMAGE > 0 Then
        '  Picture1.Visible = True
        'End If

370     Listview1.ListItems.ADD , "a1", "Καρτέλλα", 2, 2
380     Listview1.ListItems.ADD , "a2", "Διόρθ.Πελάτη", 3, 3
390     Listview1.ListItems.ADD , "a3", "Εισ.Παρ/κου", 6, 6
400     Listview1.ListItems.ADD , "a4", "Παραμ.Παραστατικων", 15, 15
410     Listview1.ListItems.ADD , "a5", "Πωλήσεις/Αγορές", 16, 16
420     Listview1.ListItems.ADD , "a6", "Προβολη Παρ/κού", 17, 17
430     Listview1.ListItems.ADD , "a7", "Υπόλοιπα Πελατών", 18, 18
        'Listview1.ListItems.Add , "b3", "Aναζήτηση", 5, 5
    
440     ListView2.ListItems.ADD , "b1", "Πίνακες", 1, 1
450     ListView2.ListItems.ADD , "b2", "Back-Up", 17, 17
        'Listview2.ListItems.Add , "a8", "Business Info", 25, 25
460     ListView2.ListItems.ADD , "b3", "Για το πρόγραμμα ", 18, 18

470     pelat.Enabled = True
480     apot.Enabled = True
490     epit.Enabled = True
500     Par.Enabled = True
510     bohu.Enabled = True

520     If F_BACKGROUND = 1 Then
            ' OK ΑΦΗΝΩ ΤΟ ΑΡΧΕΙΟ ΩΣ ΕΧΕΙ
        Else
530         gPicture = ""
540         MDIForm1.Picture = LoadPicture(gPicture)
        End If

        ' Gdb.Open gConnect

        ' f_loadpar1 = Val(FindParametroi("MDIFORM1", "F_LOADPAR1", "1", "ΦΟΡΤΩΝΩ ΚΑΘΕ ΦΟΡΑ ΤΗΝ ΦΟΡΜΑ ΠΑΡΑΣΤΑΤΙΚΩΝ=1"))

550     f_telephone = Val(FINDPARAMETROI("MDIFORM1", "F_TELEPHONE", "0", "ΕΜΦΑΝΙΣΗ ΚΑΛΟΥΝΤΩΝ ΤΗΛΕΦΩΝΩΝ=1"))

560     F_TOOLBAR = Val(FINDPARAMETROI("MDIFORM1", "F_TOOLBAR", "550", "ΥΨΟΣ TOOLBAR"))
570     f_Ypenth = Val(FINDPARAMETROI("MDIFORM1", "F_YPENTH", "1", "ΥΠΕΝΘΥΜΙΣΗ=1 ΟΧΙ=0"))
        'f_supervisor_code = Val(FindParametroi("MDIFORM1", "f_supervisor_code", "1", "Κωδικός"))

580     tlbToolBar.Height = F_TOOLBAR

        On Error GoTo stopthlefona

590     If f_telephone = 1 Then
600         TELEFON.ConnectionString = gConnect
610         TELEFON.RecordSource = "SELECT TOP 1 * FROM THLEFNOW"
620         TELEFON.Refresh
        End If

630     If F_TOOLBAR = 0 Then
640         tlbToolBar.Visible = False
650         StatusBar1.Visible = False
        End If

        On Error Resume Next

        '150
      '  Update_Structures

660     If gXEIRISTHS >= 8 Then
670         MakeBackup ""
        End If
        
        
Dim f_IS_PARTIDES As Integer
f_IS_PARTIDES = Val(FINDPARAMETROI("APOT2", "f_IS_PARTIDES", "0", "Αν Χρησιμοποιώ παρτίδες=1"))
If f_IS_PARTIDES = 1 Then
  'ΣΒΗΝΩ ΤΙΣ ΠΑΡΤΙΔΕΣ ΠΟΥ ΔΗΜΙΟΥΡΓΗΣΑ ΚΑΙ ΔΕΝ ΚΑΤΑΧΩΡΗΘΗΚΕ ΤΟ ΤΙΜΟΛΟΓΙΟ
  Gdb.Execute "DELETE FROM PARTIDES WHERE POSOAG>0 AND ID_NUM=0"

End If





        On Error GoTo update_structure

680     R.Open "select TOP 1 KATHGORIA from EID", Gdb, adOpenForwardOnly, adLockReadOnly

        On Error Resume Next

690     R.Close

        ' Gdb.Execute "ALTER TABLE EID ADD KATHGORIA INT"

        'On Error GoTo UPDATE2_STRUCTURE
        Dim dum2

        '  DUM2 = ADD_FIELD("PARASTAT", "ID", "INT IDENTITY(1,1)")
        'R.Open "select TOP 1 ARITMISI,TAYTPEL,TAYTXPI from XREOPIS", Gdb, adOpenForwardOnly, adLockReadOnly
        'On Error Resume Next
        'R.Close

700     MDIForm1.StatusBar1.Panels(2).Text = gConnect

710     If gConnect = "dBase III;" Then

            On Error GoTo MDIForm_Load_Err

720         update_tables
        End If

        'frmLogin.SHOW 1

730     If gXEIRISTHS < 8 Then
            'bohu.Visible = False
740         mapot1(21).Visible = False
750         pareid(23).Visible = False
760         apo4(24).Visible = False
770         apo5(25).Visible = False
780         mapot6(26).Visible = False
790         epit.Visible = False

800         mBohu1(51).Visible = False
810         mbohu2(52).Visible = False

820         par3(53).Visible = False
830         parkin(54).Visible = False
840         pros(55).Visible = False
850         genlog(56).Visible = False

860         If gXEIRISTHS <= 2 Then

870             mapot2(22).Visible = False
880             apoemporiko(1).Visible = False    'True
890             DBF(57).Visible = False
                ' pontoif(58).Visible = False
900             apoemporiko(59).Visible = False
910             calculator(510).Visible = False
920             mpelat1.Visible = False
930             mpelat2.Visible = False
940             mpelat3.Visible = False
950             mpelat4.Visible = False
960             mpelat5.Visible = False
970             apot.Visible = False

980             mapot7(27).Visible = False
990             mpar7.Visible = False
1000            mBohu13.Visible = False

1010            diagrpar(42).Visible = False
1020            parenos(43).Visible = False
1030            mpar5.Visible = False

1040            If gXEIRISTHS = 2 Then
                Else
1050                mPar6.Visible = False
                End If

1060            mbohu12.Visible = False
1070            timspar(44).Visible = False

1080            mapot2(22).Visible = False

1090            mapot7(27).Visible = False

            End If

        Else

1100        If f_Ypenth = 1 Then utPROEID.SHOW

        End If

        ' gDir = "\\DCCSERVER\data\lageuro" '  "\\R1r5q6\d\lageuro"
        'gDir = "d:\lageuro3" '  "\\R1r5q6\d\lageuro"
        '  gConnect = "dBase IV;"

        'gDir = ektyp_forma("ss")

        '   gDir = "c:\MERCVB\EMP.mdb"
        '  gConnect = "" 'Access
        ' gDek_pos = "###,###.00"
        ' gDek_ajia = "###,###.00"

        'ektyp_forma ("c:\mercvb\f90")

        ' Get default Workspace.
1110    Set wrkDefault = DBEngine.Workspaces(0)

        'MYDB = wrkDefault.OpenDatabase("c:\MERCVB\EMP.mdb")
        'If MYDB.CollatingOrder <> dbLangGreek Then
        '   wrkDefault.MYDB.Close
        '   wrkDefault.CompactDatabase "emp.mdb", "emp2.mdb", dbLangGreek
        'End If

1120    Me.Caption = "LoadParam"

        On Error GoTo MDIForm_Load_Err

        On Error Resume Next

1130    LOADPARAMETERS

1140    Me.Caption = "Menu"

        On Error GoTo CREATEPARAMETROI

1150    R.Open "select * from PARAMETROI", Gdb, adOpenForwardOnly, adLockReadOnly

1160    R.Close

        Dim Sql2, SQL3 As String

        On Error GoTo CREATEXRHSEIS

1170    SQL3 = "select * from XRHSEIS where '" + Format(gWorkDay, "MM/DD/YYYY") + "'>=ENARXH AND '" + Format(gWorkDay, "MM/DD/YYYY") + "'<=LHXH"

        On Error GoTo MDIForm_Load_Err

1180    R.Open SQL3, Gdb, adOpenForwardOnly, adLockReadOnly
    
1190    If R.EOF Then
            'BAZΩ 1/1/2000 ΓΙΑ ΝΑ ΜΗΝ ΒΓΑΖΕΙ ΛΑΘΟΣ ΥΠΟΛΟΙΠΑ
1200        Sql2 = "INSERT INTO XRHSEIS (ETOS,ENARXH,LHXH) VALUES (" + LTrim(str(Year(gWorkDay))) + ",'01/01/2000','12/31/" + LTrim(str(Year(gWorkDay))) + "')"
1210        Gdb.Execute Sql2
1220        R.Close
1230        R.Open SQL3, Gdb, adOpenForwardOnly, adLockReadOnly
        End If

1240    gEnarjh = R!enarxh
1250    gLhjh = R!Lhxh

1260    R.Close

        'ΕΛΕΓΧΩ ΑΝ Ο ΠΙΝΑΚΑΣ EGGTIM EXEI ID
        'select object_name(object_id), name
        'FROM sys.Columns
        'where is_identity = 1 AND object_name(object_id)='EGGTIM'

        R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'LOGGING'", Gdb, adOpenDynamic, adLockOptimistic

        On Error Resume Next

        If R(0) = 0 Then
            Gdb.Execute "CREATE TABLE LOGGING (PROCED  CHAR(70),ENTOLH TEXT,HME  DATETIME,ID INT IDENTITY(1,1) )"
        End If

        R.Close
  
        'Dim R As New ADODB.Recordset
        '================================ EIDH_DEFAULT ================================
        R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'EIDH_DEFAULT'", Gdb, adOpenDynamic, adLockOptimistic

        If R(0) = 0 Then
            Gdb.Execute "CREATE TABLE EIDH_DEFAULT ( AA INT NOT NULL,KOD VARCHAR(16) NOT NULL,TIMM FLOAT NULL,EKPT FLOAT NULL,ID INT IDENTITY(1,1) NOT NULL, AA_TIMOL INT NULL ) "
        End If

        R.Close
  
        Exit Sub

        'UPDATE2_STRUCTURE:
        'On Error Resume Next
        '
        'On Error Resume Next
        '   Gdb.Execute "ALTER TABLE XREOPIS ADD ARITMISI INT "
        '   Gdb.Execute "ALTER TABLE XREOPIS ADD TAYTPEL  CHAR(15) "
        '   Gdb.Execute "ALTER TABLE XREOPIS ADD TAYTXPI  CHAR(1) "
        '
        '    Resume Next

stopthlefona:
1270    f_telephone = 0
1280    Timer1.Enabled = False

1290    Resume Next

update_structure:
1300    Gdb.Execute "ALTER TABLE EID ADD KATHGORIA INT"
1310    Gdb.Execute "UPDATE EID SET KATHGORIA=1"

1320    Resume Next

CREATEPARAMETROI:

1330    Gdb.Execute "CREATE TABLE PARAMETROI (FORMA CHAR(20),VAR CHAR(25) ,TIMH CHAR(10),SXOLIA CHAR(50)  )"

        Exit Sub

CREATEXRHSEIS:

        'Resume Next
1340    On Error Resume Next

        ' Gdb.Execute "CREATE TABLE XRHSEIS (ETOS SMALLINT,ENARXH DATETIME ,LHXH DATETIME,N1 SMALLINT,N2 SMALLINT,C1 CHAR(20),C2 CHAR(20),ID INT IDENTITY(1,1) )  "

        Exit Sub

SEELINE:
1350    HandleError "MdiForm-load"
1360    MsgBox str(Erl), vbOKOnly, minima + " " + Err.Description

        '     Resume Next
        Exit Sub    'Resume Next

        '<EhFooter>
        Exit Sub

MDIForm_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.MDIForm_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.MDIForm_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub MDIForm_Resize()

    '<EhHeader>
    'On Error Resume Next
    '</EhHeader>

    ' On Error Resume Next

    If MDIForm1.Width > 15225 Then
        '  MDIForm1.width = 15225  ' MDIForm1.Picture.width
    End If

    'MDIForm1.Picture.Width = MDIForm1.Width

    'MDIForm1.Picture2.Width = MDIForm1.Width
    If MDIForm1.Width > 15000 Then

        'StretchSourcePictureFromPicture MDIForm1.Picture2.Picture, MDIForm1.Picture2
        '            MDIForm1.Picture2.Refresh

        MDIForm1.BackColor = &HFF0000
        MDIForm1.Picture = LoadPicture()
    Else
        MDIForm1.Picture = LoadPicture(gPicture)
    End If

    'mainpic.Width = MDIForm1.Width
    'mainpic.Height = MDIForm1.Height
    'mainpic.Picture = LoadPicture(gPicture)

    'If MDIForm1.WindowState = 2 Then MDIForm1.WindowState = 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

  ' Signal Application Shutdown
 '   g_InteropToolbox.EventMessenger.RaiseApplicationShutdownEvent

End Sub

Private Sub METASX_DA_Click()
 
        Dim m As Integer

        Dim C As String

        PAR1.TDBGrid1.Col = 2
        C = Replace(PAR1.TDBGrid1.Text, ",", ".")
        m = Val(FINDPARAMETROI("PAR1", "METASX_DA", "3", "Α/Α Δ.Α. ΣΤΑ ΠΑΡΑΣΤΑΤΙΚΑ ΔΑ(ΜΕΤΑΣΧ.ΣΕ ΔΑ)0-99"))
100     PAR1.metasx.Value = vbChecked
        
        
         Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
         m_ID = Val(PAR1.TDBGrid1.Text)
         PAR1.METASX_SE_TIM m, C, m_ID

'        PAR1.METASX_SE_TIM m, C, 0

        Exit Sub

End Sub

Private Sub metasx_lian_Click()

        '<EhHeader>
        On Error GoTo metasx_LIAN_Click_Err

        '</EhHeader>
        
        Dim m As Integer

        Dim C As String
         
        PAR1.TDBGrid1.Col = 2
        C = Replace(PAR1.TDBGrid1.Text, ",", ".")
        'PAR1.Text5.Text = "*" + C
        
        m = Val(FINDPARAMETROI("PAR1", "METASX_LIAN", "3", "Α/Α ΔΛ.ΔΑ ΣΤΑ ΠΑΡΑΣΤΑΤΙΚΑ (ΜΕΤΑΣΧ.ΣΕ ΛΙΑΝΙΚΗ)0-99"))
100     PAR1.metasx.Value = vbChecked


         Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)
        PAR1.METASX_SE_TIM m, C, m_ID

        Exit Sub

metasx_LIAN_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.metasx_par_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.metasx_par_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub par1_Click()
'    PAR1.SHOW
'End Sub

Private Sub metasx_par_Click()
        
        Dim m As Integer

        Dim C As String

        PAR1.TDBGrid1.Col = 2
        C = Replace(PAR1.TDBGrid1.Text, ",", ".")
        
        PAR1.metasx.Value = vbChecked
        PAR1.METASX_SE_TIM 99, C, 0

        Exit Sub

        '----------------------------------------------------------------------------------------
        
        '<EhHeader>
        On Error GoTo metasx_par_Click_Err

        '</EhHeader>

        'Dim m As Integer

        m = Val(FINDPARAMETROI("PAR1", "METASX_TIM", "2", "ΠΑΡΑΣΤΑΤΙΚO ΤΠΔΑ (ΜΕΤΑΣΧ.ΣΕ ΤΙΜΟΛ)0-99"))
100     PAR1.metasx.Value = vbChecked

        PAR1.metasx.Value = vbChecked

110     'PAR1.diortosis

        PAR1.MONTE_CRISTO.Visible = True
        PAR1.MONTE_CRISTO.Caption = Format(m, "00") + "-" + "ΜΕΤΑΣΧΗΜΑΤΙΣΜΟΣ ΣΕ ΤΙΜ"

        'PAR1.Text2(0).SetFocus
        ' SendKeys "{enter}{enter}{enter}"

        '<EhFooter>
        Exit Sub

metasx_par_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.metasx_par_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.metasx_par_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub METASX_TIM_Click()

        '<EhHeader>
        On Error GoTo metasx_TIM_Click_Err

        '</EhHeader>
   
        Dim m As Integer

        Dim C As String
         
        PAR1.TDBGrid1.Col = 2
        C = Replace(PAR1.TDBGrid1.Text, ",", ".")
        'PAR1.Text5.Text = "*" + C
        
        m = Val(FINDPARAMETROI("PAR1", "METASX_TIM", "3", "ΠΑΡΑΣΤΑΤΙΚO ΔΛ.ΔΑ(ΜΕΤΑΣΧ.ΣΕ ΛΙΑΝΙΚΗ)0-99"))
100     PAR1.metasx.Value = vbChecked

         Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)
'        PAR1.METASX_SE_TIM m, C, m_ID
        PAR1.METASX_SE_TIM m, C, m_ID
        Exit Sub

        '============================================================================

metasx_TIM_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.metasx_par_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.metasx_par_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mpar1_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo mpar1_Click_Err

        '</EhHeader>

100     gApoMenu = True

110     PAR1.SHOW

        '<EhFooter>
        Exit Sub

mpar1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mpar1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mpar1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mpar5_Click()

        '<EhHeader>
        On Error GoTo mpar5_Click_Err

        '</EhHeader>

100     If gXEIRISTHS = 9 Then
        Else
110         par5.SHOW
        End If

        '<EhFooter>
        Exit Sub

mpar5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mpar5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mpar5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub ListView1_Click()
'     Dim selItem As ListItem
'    On Error GoTo RAE
'  '  Set selItem = Listview1.SelectedItem
'    Select Case Listview1.SelectedItem.Key
'        Case "a1": Label3.Caption = "Products"
'        Case "a2": Label3.Caption = "Suppliers"
'        Case "a3": Label3.Caption = "Category"
'        Case "a4": Label3.Caption = "Stock Monitoring"
'        Case "a5": Label3.Caption = "Sales Entry"
'        Case "b1": Label3.Caption = "Daily Sales"
'        Case "b2": Label3.Caption = "Stock Recieve"
'    End Select
'RAE:
'    Set selItem = Nothing
'End Sub
Private Sub ListView1_DblClick()

        '<EhHeader>
        On Error GoTo ListView1_DblClick_Err

        '</EhHeader>
        Dim selItem As ListItem
    
        On Error GoTo RAE
    
        On Error GoTo ListView1_DblClick_Err

        'Set selItem = Listview1.SelectedItem
  
        '  Listview1.ListItems.Add , "a1", "Καρτέλλα", 2, 2
        '    Listview1.ListItems.Add , "a2", "Διόρθ.Πελάτη", 3, 3
        '    Listview1.ListItems.Add , "a3", "Εισ.Παρ/κου", 6, 6
        '    Listview1.ListItems.Add , "a4", "Παραμ.Παραστατικων", 15, 15
        '    Listview1.ListItems.Add , "a5", "Πωλήσεις/Αγορές", 16, 16
        '    Listview1.ListItems.Add , "b1", "Προβολη Παρ/κού", 17, 17
        '    Listview1.ListItems.Add , "b2", "Υπόλοιπα Πελατών", 18, 18
        '    'Listview1.ListItems.Add , "b3", "Aναζήτηση", 5, 5
        '
        '    Listview2.ListItems.Add , "a6", "Πίνακες", 16, 16
        '    Listview2.ListItems.Add , "a7", "Back-Up", 17, 17
100     Select Case Listview1.SelectedItem.Key

            Case "a1":
110             pelat4.SHOW

            Case "a2":
120             pelat2.SHOW
          
            Case "a3":
130             PAR1.SHOW

            Case "a4":
140             bohu3.SHOW

            Case "a5":
150             par4.SHOW

            Case "a6":
160             PAR2.SHOW
        
                '        Unload frmSuppliers
                '        Unload frmCategorys
                '        Unload frmMonitoring
                '        Unload frmStockRecieves
                '        Unload frmProducts
                '        loadForm frmSalesDaily
                '
            Case "a7":
170             pelat3.SHOW
        
                '        Unload frmSuppliers
                '        Unload frmCategorys
                '        Unload frmMonitoring
                '        Unload frmSalesDaily
                '        Unload frmProducts
                '        loadForm frmStockRecieves
                '
                '        Case "b3":
                '        loadForm anazitisi
                '

        End Select

RAE:
180     Set selItem = Nothing

        '<EhFooter>
        Exit Sub

ListView1_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.ListView1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.ListView1_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mPar6_Click()

        '<EhHeader>
        On Error GoTo mPar6_Click_Err

        '</EhHeader>

100     par6.SHOW

        '<EhFooter>
        Exit Sub

mPar6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mPar6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mPar6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mpar7_Click()

        '<EhHeader>
        On Error GoTo mpar7_Click_Err

        '</EhHeader>

100     par7.SHOW

        '<EhFooter>
        Exit Sub

mpar7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mpar7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mpar7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mpar8_Click()
    ' par8.SHOW

End Sub

Private Sub mpar9_Click()
    ' par9.SHOW
End Sub

'Private Sub mpar1_Click()

'End Sub

Private Sub mpelat1_Click()

        '<EhHeader>
        On Error GoTo mpelat1_Click_Err

        '</EhHeader>

100     pelat1.SHOW

        '<EhFooter>
        Exit Sub

mpelat1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mpelat1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mpelat1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mpelat2_Click()

        '<EhHeader>
        On Error GoTo mpelat2_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     pelat2.SHOW

        '<EhFooter>
        Exit Sub

mpelat2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mpelat2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mpelat2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mpelat3_Click()

        '<EhHeader>
        On Error GoTo mpelat3_Click_Err

        '</EhHeader>

100     pelat3.SHOW

        '<EhFooter>
        Exit Sub

mpelat3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mpelat3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mpelat3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub mpelat4_Click()

        '<EhHeader>
        On Error GoTo mpelat4_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     pelat4.resizing = 1
120     pelat4.SHOW

        '<EhFooter>
        Exit Sub

mpelat4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mpelat4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mpelat4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub par1_Click()

'End Sub

Private Sub mpelat5_Click()

        '<EhHeader>
        On Error GoTo mpelat5_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     pelat5.SHOW

        '<EhFooter>
        Exit Sub

mpelat5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.mpelat5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mpelat5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub par1menu_Click()

'
'End Sub

Private Sub par2show_Click()

        '<EhHeader>
        On Error GoTo par2show_Click_Err

        '</EhHeader>

100     PAR1.PAR2SHOWING

        '<EhFooter>
        Exit Sub

par2show_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.par2show_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.par2show_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub par3_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo par3_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     bohu3.SHOW

        '<EhFooter>
        Exit Sub

par3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.par3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.par3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub pareid_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo pareid_Click_Err

        '</EhHeader>

100     gApoMenu = True
        '  apot3.SHOW

        '<EhFooter>
        Exit Sub

pareid_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.pareid_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.pareid_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub parenos_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo parenos_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     PAR2.SHOW

        '<EhFooter>
        Exit Sub

parenos_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.parenos_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.parenos_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub parepit_Click()

        '<EhHeader>
        On Error GoTo parepit_Click_Err

        '</EhHeader>

100     gApoMenu = True
        '   epit3.Top = 1450
        '  epit3.Left = 90
        ' epit3.width = MDIForm1.width - 100
        ' epit3.height = MDIForm1.height - 1500

110     epit3.SHOW

        '<EhFooter>
        Exit Sub

parepit_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.parepit_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.parepit_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub parkin_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo parkin_Click_Err

        '</EhHeader>

100     gApoMenu = True
110     bohu4.SHOW

        '<EhFooter>
        Exit Sub

parkin_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.parkin_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.parkin_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub pontoi_Click(Index As Integer)
'    gApoMenu = True
'    bohu6.SHOW
'
'End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Picture1_KeyPress_Err

        '</EhHeader>
        Dim a

100     If KeyAscii = 27 Then
110         a = 1

120         a = MsgBox("Εξοδος ;", vbYesNo)

130         If a = vbYes Then End

        End If

        '<EhFooter>
        Exit Sub

Picture1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.Picture1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.Picture1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub pontoif_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo pontoif_Click_Err

        '</EhHeader>

100     bohu7.SHOW

        '<EhFooter>
        Exit Sub

pontoif_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.pontoif_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.pontoif_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub pros_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo pros_Click_Err

        '</EhHeader>

100     bohu5.SHOW

        'Dim db, DUM, aa As String, pol As String, polepis As String, ago As String, agoepis As String
        '
        ''Set db = OpenDatabase(gDir, False, False, gConnect)
        'Gdb.Execute UCase("update eid set m01=0,m02=0,m03=0,m04=0,m05=0,m06=0,m07=0,m08=0,m09=0,m10=0,m11=0,m12=0,g01=0,g02=0,g03=0,g04=0,g05=0,g06=0,g07=0,g08=0,g09=0,g10=0,g11=0,g12=0;")
        '
        ''db.Execute "CREATE INDEX eidko ON eid(kod);"
        'pol = "": polepis = "": ago = "": agoepis = ""
        'DUM = Get_AJ(pol, polepis, ago, agoepis)
        '
        '
        '' Kill gDir + "\eidko.ndx"
        ''Gdb.Execute "CREATE  INDEX  eidko  ON   eid(kod);"
        '
        '
        'Dim R As New ADODB.Recordset, EID As New ADODB.Recordset
        '
        '
        '
        '
        ''eid.Index = "eidko"
        '
        'R.Open UCase("select kode,sum(poso) as sposo ,left(atim,1) as satim,month(hme) as shme from eggtim group by kode,left(atim,1),month(hme);"), Gdb, adOpenDynamic, adLockOptimistic
        '
        '
        ''Set r = db.OpenRecordset("eggtim")
        'Do While Not R.EOF
        '
        '
        '
        '   EID.Open "SELECT *FROM EID WHERE KOD='" + R("KODE") + "'", Gdb, adOpenDynamic, adLockOptimistic
        '   'eid.Seek "=", r("kode")
        '
        '   If Not EID.EOF Then
        '      ' eid.Edit
        '      aa = Format(R("shme"), "00")
        '
        '     If InStr(pol, R("satim")) > 0 Then EID("m" + aa) = EID("m" + aa) + R("sposo")
        '     If InStr(polepis, R("satim")) > 0 Then EID("m" + aa) = EID("m" + aa) - R("sposo")
        '     If InStr(ago, R("satim")) > 0 Then EID("g" + aa) = EID("g" + aa) + R("sposo")
        '     If InStr(agoepis, R("satim")) > 0 Then EID("g" + aa) = EID("g" + aa) - R("sposo")
        '     EID.Update
        '   End If
        '   R.MoveNext
        '
        'Loop
        '
        '
        'db.Close
        '

        '<EhFooter>
        Exit Sub

pros_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.pros_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.pros_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub RUNT_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo RUNT_Click_Err

        '</EhHeader>
        Dim k  ' INDEX=0 TO PROTO MENOY   INDEX=1 TO DEYTERO

100     gDir = ETAIR(Index)

110     MDIForm1.StatusBar1.Panels(1).Text = gDir

120     If InStr(UCase(gDir), "DSN") > 0 Then
130         gConnect = gDir
        Else
140         gConnect = gConnect + ";DATABASE=" + gDir

        End If

150     If Left(gDir, 1) = "D" Then
160         gPicture = "C:\MERCVB\PALIA.BMP"
        Else
170         gPicture = "C:\MERCVB\BACKGROUND.BMP"
        End If

180     For k = 1 To mc
190         Unload runt(k)
        Next

        'For K = 1 To 1000: DoEvents: Next
        'MDIForm1.Picture = LoadPicture()
        'MILSEC 1000

200     MsgBox "Aλλαγή εταιρείας"
210     MDIForm1.Picture = LoadPicture(gPicture)

        'For K = 1 To 1000: DoEvents: Next

        '<EhFooter>
        Exit Sub

RUNT_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.RUNT_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.RUNT_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)

        '<EhHeader>
        On Error GoTo StatusBar1_PanelClick_Err

        '</EhHeader>
        Dim a

        Dim DB As Database, R As New ADODB.Recordset, k As Integer

        On Error Resume Next

100     R.Open "select * from ETAIREIES", Gdb, adOpenForwardOnly, adLockReadOnly

        'MDIForm1.Picture = LoadPicture()
110     mc = 0
120     runt(0).Caption = R(0)
130     ETAIR(0) = R(0)

140     Do While Not R.EOF
150         mc = mc + 1
160         Load runt(mc)
170         runt(mc).Caption = R!TITLOS
180         ETAIR(mc) = Trim(R(0))
190         R.MoveNext
        Loop

200     PopupMenu RUNTIME(0)

        '<EhFooter>
        Exit Sub

StatusBar1_PanelClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.StatusBar1_PanelClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.StatusBar1_PanelClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub TELEFON_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

'  TELEFON.Refresh
'
'  If IsNull(TELEFON.Recordset("EPO")) Then
'     MDIForm1.Caption = TELEFON.Recordset("THL") + " "
'  Else
'     MDIForm1.Caption = TELEFON.Recordset("THL") + " " + TELEFON.Recordset("EPO")
'  End If
'
'
'
'End Sub

Private Sub picMenu_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    cmdset.Top = picMenu.Top + picMenu.Height
    picSet.Top = picMenu.Top + picMenu.Height + cmdset.Height
End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    Select Case getAction

        Case "picSet"

            If bhide = False Then
                If picSet.Height >= 3495 Then
                    picSet.Height = 3495
                    Timer1.Interval = 0
                Else
                    picSet.Height = picSet.Height + 249
                End If
                
            Else
            
                If picSet.Height <= 600 Then
                    picSet.Height = 0
                    Timer1.Interval = 0
                    picSet.Visible = False
                Else
                    picSet.Height = picSet.Height - 249

                    DoEvents
                End If
                
            End If
            
        Case "picFile"
        
            If bhide = False Then
        
                If picMenu.Height >= 4575 Then
                Else
                    picMenu.Height = picMenu.Height + 249
                End If
        
            Else

                If picMenu.Height <= 600 Then
                    picMenu.Height = 0
                    picMenu.Visible = False
                Else
                    picMenu.Height = picMenu.Height - 249

                    DoEvents
                End If
            End If

    End Select

    If f_telephone = 1 Then

        Dim M_EPO

        If IsNull(TELEFON.Recordset("EPO")) Then M_EPO = Left(TELEFON.Recordset("EPO"), 30) Else M_EPO = Space(30)

        TELEFON.Refresh

        'If IsNull(TELEFON.Recordset("EPO")) Or IsNull(TELEFON.Recordset("DIE")) Then
        '  MDIForm1.Caption = TELEFON.Recordset("THL") + " "
        'MDIForm1.Caption = Left(TELEFON.Recordset("THL") + " " + Space(40), 40) + Left(MDIForm1.Caption, 40)
        'Else

        If Left(MDIForm1.Caption, 16) = Left(TELEFON.Recordset("THL") + " " + TELEFON.Recordset("EPO"), 16) Then
        Else
            MDIForm1.StatusBar1.Panels(5).Text = Left(MDIForm1.Caption, 42)
            MDIForm1.Caption = Left(TELEFON.Recordset("THL"), 16) + "-" + TELEFON.Recordset("EPO") + "  προηγούμενο : " + Left(MDIForm1.Caption, 40)  '+ TELEFON.Recordset("DIE")
        End If

        'MDIForm1.Caption = Left(TELEFON.Recordset("THL") + " " + TELEFON.Recordset("EPO") + Space(40), 40) + Left(MDIForm1.Caption, 40)

        'End If

        Exit Sub

    End If

    Dim R As New ADODB.Recordset

    R.Open "SELECT * FROM CalendarEvents where ABS(DATEDIFF(mi,StartDateTime,Getdate()))< ReminderMinutesBeforeStart AND IsReminder=1 ", Gdb, adOpenForwardOnly, adLockReadOnly

    '   Exit Sub
    If R.EOF Then
        R.Close

        Exit Sub

    End If

    Dim DELT As String

    DELT = "---- Yπενθύμιση --- " + Chr(13)

    ' Exit Sub
    Do While Not R.EOF
        ' + "Προθεσμία min:" + Format(r("DIFF"), "#####") + "/"
        DELT = DELT + R("Subject") + "/" + R("Location") + "/" + R("Body") + "/" + Format(R("StartDateTime"), "dd/mm/yyyy hh:mm") + Chr(13)
        R.MoveNext
    Loop

    MsgBox DELT
    R.Close

End Sub

'Private Sub timspar_Click()

'End Sub
Private Sub timspar_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo timspar_Click_Err

        '</EhHeader>

100     par4.SHOW

        '<EhFooter>
        Exit Sub

timspar_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.timspar_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.timspar_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub CREATEDATABASE(DB)

        '<EhHeader>
        On Error GoTo CREATEDATABASE_Err

        '</EhHeader>

        Dim DefaultWorkspace As Workspace

        Dim CurrentDatabase  As Database, MyDatabase As Database

100     Set DefaultWorkspace = DBEngine.Workspaces(0)
        'Set CurrentDatabase = DefaultWorkspace.Databases(0)
        ' Create new, encrypted database.
110     Set MyDatabase = DefaultWorkspace.CREATEDATABASE(DB, dbLangGreek)
        ' Fill in new database.
        ' Set property to that of current database.
        'MyDatabase.QueryTimeout = CurrentDatabase.QueryTimeout
        '...
        ' Complete structure of MyDatabase, possibly using structure of
        ' CurrentDatabase.
        '...
120     MyDatabase.Close

        '<EhFooter>
        Exit Sub

CREATEDATABASE_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.CREATEDATABASE " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.CREATEDATABASE " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub MakeBackup(THESI As String)

        '-------------------------------------------------------------
        '<EhHeader>
        On Error GoTo MakeBackup_Err

        '</EhHeader>

        Dim k    As Integer

        Dim R    As New ADODB.Recordset

        Dim mkod As String

        Dim DUM

        Dim DD

        Dim M_BACKUPTYPE

100     M_BACKUPTYPE = FINDPARAMETROI("BOHU8", "F_BACKUPTYPE", 0, "ΤΥΠΟΣ BACKUP 0=ACCESS 1=SQL DATABASE")    'posa psifia tha exei h kathe seira

        Dim M_PATH

110     M_PATH = Trim(FINDPARAMETROI("BOHU8", "F_PATH", "C:\MERCVB", "ΔΙΑΔΡΟΜΗ ΑΝΤΙΓΡΑΦΩΝ"))    'posa psifia tha exei h kathe seira)

120     If Len(THESI) > 0 Then
130         M_PATH = THESI
        End If

        Dim DDTOP

140     If M_BACKUPTYPE = 1 Then
150         DD = M_PATH + "\BACK" + Format(Day(Now), "00") + ".BAK"

160         If Len(Dir(DD, vbNormal)) > 0 Then
170             If Format(FileDateTime(DD), "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then

                    Exit Sub

                Else
180                 Kill DD
                    'yparxei_palio = True
                End If
            End If

190         Gdb.Execute "BACKUP DATABASE [" + Gdb.DefaultDatabase + "] TO  DISK ='" + DD + "' WITH NOFORMAT, NOINIT, SKIP, NOREWIND, NOUNLOAD,  STATS = 10 ", k

            Exit Sub

        End If

200     DD = M_PATH + "\BACK" + Format(Day(Now), "00") + ".MDB"
210     DDTOP = "C:\MERCVB\BACK" + Format(Day(Now), "00") + ".MDB"

        Dim OK As Boolean

220     OK = False

        'Dim yparxei_palio As Boolean
        'yparxei_palio = False
230     If Len(Dir(DD, vbNormal)) > 0 Then
240         If Format(FileDateTime(DD), "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then

                Exit Sub

            Else
250             Kill DD
                'yparxei_palio = True
            End If

        End If

        Dim LISTA(20)

260     LISTA(1) = "PEL"
270     LISTA(2) = "EID"
280     LISTA(3) = "EGG"
290     LISTA(4) = "EGGTIM"
300     LISTA(5) = "TIM"
310     LISTA(6) = "GRA"
320     LISTA(7) = "MEM"
330     LISTA(8) = "PINAKES"
340     LISTA(9) = "XREOPIS"
350     LISTA(10) = "ARITMISI"
360     LISTA(11) = "PARASTAT"
370     LISTA(12) = "BARCODES"
380     LISTA(13) = "HMEROL"
390     LISTA(14) = "HMEROL2"
400     LISTA(15) = "PARAMETROI"
410     LISTA(16) = "CalendarEvents"
420     LISTA(17) = "CalendarRecurrencePatterns"

        On Error Resume Next

        'Kill DD

        On Error GoTo LATOS2

430     CREATEDATABASE DD

        On Error Resume Next

        Dim DB As Database

440     Set DB = OpenDatabase("", False, False, gConnect)
450     Me.MousePointer = vbHourglass

        Dim LATOS

460     Me.Caption = "παρακαλώ περιμένετε - διαδικασία αντιγράφων"

470     For k = 1 To 17    ' List1.ListCount - 1
480         LATOS = 0

            On Error GoTo LATOS

490         DoEvents
500         Me.Caption = "παρακαλώ περιμένετε - διαδικασία αντιγράφων " + LISTA(k) + " " + LTrim(str(k)) + "/17)"
510         DB.Execute "SELECT * INTO " + LISTA(k) + "  IN '" + DD + "'  FROM  " + LISTA(k)
        Next

520     Me.MousePointer = vbNormal

530     If M_PATH = "C:\MERCVB" Then
            ' OK
        Else
            ' FileCopy DD, DDTOP

        End If

        Exit Sub

LATOS:
        'List2.AddItem Err.Description
540     LATOS = 1

550     Resume Next

560     Return

LATOS2:
570     MsgBox "ΛΑΘΟΣ ΟΝΟΜΑ ΒΑΣΗΣ ACCESS  ΔΩΣΤΕ Π.Χ.  C:\BACK.MDB"

        '<EhFooter>
        Exit Sub

MakeBackup_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.MakeBackup " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.MakeBackup " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub tlbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

        '<EhHeader>
        On Error GoTo tlbToolBar_ButtonClick_Err

        '</EhHeader>
        Dim a

100     a = 1

        Dim DUM

110     Select Case Button.Key

            Case "Dynaset"
120             HMEROL2.SHOW

130         Case "Table"

                '   hmerol.SHOW
140         Case "BIBLIO"
150             paralabh.SHOW

160         Case "Schedule"
170             DUM = Shell("C:\MERCVB\CALENDAR.EXE", vbMaximizedFocus)
        End Select

180     tlbToolBar.Buttons(1).style = tbrDefault
190     tlbToolBar.Buttons(2).style = tbrDefault

        '<EhFooter>
        Exit Sub

tlbToolBar_ButtonClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.tlbToolBar_ButtonClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.tlbToolBar_ButtonClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
