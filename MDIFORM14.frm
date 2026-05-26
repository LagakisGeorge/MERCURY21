VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E9C5BC0B-5CEA-42E0-805D-ACA6384FFB51}#1.0#0"; "jcMDITabs12.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FF0000&
   Caption         =   "oe"
   ClientHeight    =   9540
   ClientLeft      =   165
   ClientTop       =   2205
   ClientWidth     =   16890
   Icon            =   "MDIFORM14.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin jc_MDITabs.jcMDITabs jcMDITabs2 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      Style           =   4
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   8805
      Left            =   0
      ScaleHeight     =   8805
      ScaleWidth      =   2985
      TabIndex        =   2
      Top             =   420
      Width           =   2988
      Begin VB.CommandButton cmdset 
         BackColor       =   &H80000010&
         Caption         =   "Ρυθμίσεις Συστήματος"
         Height          =   192
         Left            =   -96
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7872
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.PictureBox picSet 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2340
         Left            =   34
         ScaleHeight     =   2340
         ScaleWidth      =   2940
         TabIndex        =   6
         Top             =   7164
         Width           =   2940
         Begin MSComctlLib.ListView ListView2 
            Height          =   84
            Left            =   0
            TabIndex        =   8
            Top             =   1920
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   159
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
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "MDIFORM14.frx":0442
            OLEDragMode     =   1
            NumItems        =   0
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
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
         Height          =   5340
         Left            =   34
         ScaleHeight     =   5340
         ScaleWidth      =   2940
         TabIndex        =   3
         Top             =   312
         Width           =   2940
         Begin MSComctlLib.ListView Listview1 
            Height          =   276
            Left            =   2208
            TabIndex        =   4
            Top             =   5184
            Visible         =   0   'False
            Width           =   492
            _ExtentX        =   900
            _ExtentY        =   450
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
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "MDIFORM14.frx":0D1C
            OLEDragMode     =   1
            NumItems        =   0
         End
         Begin ComctlLib.TreeView tvDir 
            Height          =   9000
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   2964
            _ExtentX        =   5239
            _ExtentY        =   15875
            _Version        =   327682
            Indentation     =   0
            LineStyle       =   1
            Style           =   7
            ImageList       =   "ImgLstFolder"
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
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
         Picture         =   "MDIFORM14.frx":15F6
         Stretch         =   -1  'True
         Top             =   -1560
         Width           =   570
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   0
         Picture         =   "MDIFORM14.frx":17EE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3330
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   0
      Top             =   9228
      Width           =   16884
      _ExtentX        =   29792
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   9
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2893
            MinWidth        =   2893
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "GDIR"
            TextSave        =   "GDIR"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3175
            MinWidth        =   3175
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3492
            MinWidth        =   3492
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   ".."
            TextSave        =   ".."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16890
      _ExtentX        =   29792
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
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
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":19E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":2052
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":2164
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":2276
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":2388
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":249A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":3174
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":3E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":4B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":5802
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":64DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":71B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":7E90
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":8B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":9844
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":A51E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":B1F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":BED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":CBAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":D886
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":E560
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":F23A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":FF14
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":10BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":118C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":125A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1327C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":13F56
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":14C30
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1590A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":165E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":172BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":17F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":18C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1994C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1A626
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1B300
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1BFDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1CCB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1D98E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1E668
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1E8F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1EFA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1FD00
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1FD5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":1FDBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":23296
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFORM14.frx":23B70
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImgLstFolder 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFORM14.frx":2444A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFORM14.frx":2499C
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
         Begin VB.Menu mpel3PelatesGenika 
            Caption         =   "Πελάτες/Προμηθευτές"
         End
         Begin VB.Menu mpel3PelatYPOLOIPA 
            Caption         =   "Υπόλοιπα Πελατών"
         End
         Begin VB.Menu mpel3promYPOLOIPA 
            Caption         =   "Υπόλοιπα Προμηθευτών"
         End
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
         Begin VB.Menu mapot32c 
            Caption         =   "Απογραφή Αποθήκης(από προηγ.έτος)"
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
   Begin VB.Menu par 
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
         Caption         =   "&7.MyData"
      End
      Begin VB.Menu mpar8 
         Caption         =   "&8.Εισαγωγή Ψηφιακου Πελατολογίου"
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
         Caption         =   "10.Γέφυρα XML Epsilon"
         Index           =   510
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
      Begin VB.Menu mBohu14 
         Caption         =   "14.Εύκολη διαμόρφωση myDta"
      End
      Begin VB.Menu THLYPOST 
         Caption         =   "15.Tηλέφωνα Υποστήριξης"
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
      Begin VB.Menu paralepisk 
         Caption         =   "Παραλαβές Επισκευών"
      End
      Begin VB.Menu calendar 
         Caption         =   "Ημερολόγιο"
      End
      Begin VB.Menu Εξωτερικά_Προγράμματα 
         Caption         =   "Εξωτερικά Προγράμματα"
      End
      Begin VB.Menu Updater 
         Caption         =   "Ελεγχος Νέας έκδοσης"
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
         Caption         =   "Παρουσίαση/Διαγραφή Παρ/κού"
      End
      Begin VB.Menu diortosis 
         Caption         =   "Διόρθωση"
      End
      Begin VB.Menu ektyp_par 
         Caption         =   "Εκτύπωση Παρ/κού/ Σε PDF "
      End
      Begin VB.Menu email_par 
         Caption         =   "Σε email το παρασ/κό"
         Enabled         =   0   'False
      End
      Begin VB.Menu metasx_lian 
         Caption         =   "Μετασχηματισμός σε ΑΠΟΔ."
      End
      Begin VB.Menu METASX_TIM 
         Caption         =   "Μετασχηματισμός σε ΤΙΜ.ΔΑ"
      End
      Begin VB.Menu METASX_DA 
         Caption         =   "Μετασχηματισμός σε ΔΕΛ.ΑΠ"
      End
      Begin VB.Menu akyrotiko 
         Caption         =   "Ακυρωτικό Στοιχείο"
      End
      Begin VB.Menu toMydata 
         Caption         =   "Αποστολή σε Mydata"
      End
      Begin VB.Menu epanAkyr 
         Caption         =   "Επαναποστολή Ακυρωτικού"
      End
      Begin VB.Menu menuToTransport 
         Caption         =   "Σε Μεταφορά"
      End
      Begin VB.Menu menuParadosis 
         Caption         =   "Παραδόθηκε"
      End
      Begin VB.Menu menuArnisi 
         Caption         =   "Αρνηση Παραλαβής"
      End
      Begin VB.Menu menu_diak_status 
         Caption         =   "Κατάσταση Διακίνησης"
      End
   End
   Begin VB.Menu EPITAG311 
      Caption         =   "EPITAG311"
      Visible         =   0   'False
      Begin VB.Menu epit311 
         Caption         =   "epit311"
         Index           =   0
      End
      Begin VB.Menu epit312 
         Caption         =   "epit312"
      End
   End
   Begin VB.Menu MENUXREOPIS 
      Caption         =   "MENUXREOPISFORCHECK"
      Index           =   101
      Visible         =   0   'False
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


Dim F_UPDATESTRUCTURES As Integer


Dim f_Ypenth  As Long

Dim f_supervisor_code
Attribute f_supervisor_code.VB_VarUserMemId = 1073938436

Dim f_telephone  As Integer

Dim f_is_benzinadiko As Integer
Dim fgAsterisk As Integer
Dim f_auto_load_times   As Integer

Dim F_COLOR_BACKGROUND As Integer

Dim F_BACKGROUND As Integer

Dim f_proeid_mydata As Integer


Dim f_FARDOSIMAGE
Dim F_TAKENEESTIMES As Integer

Dim bhide     As Boolean

Dim getAction As String

Dim f_site    As Long

Private Sub akyrotiko_Click()
     
     Dim F_PAROCHOS As Integer
     
     F_PAROCHOS = Val(FINDPARAMETROI(1, "PAR7MYDATA", "f_parochos", "0", "0=MyData 1=Παροχος"))

       If F_PAROCHOS = 1 Then
             PAR1.TDBGrid1.Col = 1
           ' c = Replace(PAR1.TDBGrid1.Text, ",", ".")

            Dim mpar As String: mpar = Left(PAR1.TDBGrid1.Text, 1) 'TDBGrid1.CurrentRow.cells(1).Value.toString.Substring(0, 1)
            
            PAR1.TDBGrid1.Col = 8
            Dim mmark As String: mmark = PAR1.TDBGrid1.Text
          
            Dim MYEID As String: MYEID = GGET_CVALUE("SELECT MYEID FROM PARASTAT  WHERE EIDOS='" + mpar + "'")
            If MYEID = "9.3" Then
                PAR1.TDBGrid1.Col = 7
                Dim mmID_NUM As String
                 mmID_NUM = PAR1.TDBGrid1.Text
                  Par7MyData.Visible = False
                  Par7MyData.Caption = "" ' ΙΣΑ ΙΣΑ ΝΑ ΑΝΟΙΓΕΙ ΤΟ FORM_LOAD ΓΙΑ ΝΑ ΔIΑΒΑΣΕΙ P_USER
                Par7MyData.CancelDelivery mmID_NUM, mmark
                Exit Sub
            Else
                MsgBox ("?e p????? µ??? p?st?t???")
                Exit Sub
            End If

        End If














  '<EhHeader>
        On Error GoTo akyrotiko_Click_err

        Dim m As Integer
        Dim m_akyr_neo  As Integer
m_akyr_neo = Val(FINDPARAMETROI(1, "PAR1", "f_akyr_neo", "0", "Γ.49.ΝΕΟ ΑΚΥΡΩΤΙΚΟ ΑΝΤΙΘΕΤΟ ΧΩΡΙΣ ΔΙΑΓΡΑΦΗ=1 ΟΧΙ=0"))

Dim mID_NUM As String
 PAR1.TDBGrid1.Col = 7
    mID_NUM = PAR1.TDBGrid1.Text
    
    Dim MAKY As Integer
    MAKY = PAR1.GET_NVALUE("SELECT AKYROMENO FROM TIM WHERE ID_NUM=" + mID_NUM)
    If MAKY > 0 Then
       If MAKY = 1 Then
         MsgBox "ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ ΕΙΝΑΙ ΑΚΥΡΩΜΕΝΟ. ΔEN AKYΡΩNETAI"
       End If
       If MAKY = 2 Then
         MsgBox "ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ ΕΙΝΑΙ ΑΚΥΡΩTIKO. ΔEN AKYΡΩNETAI"
       End If
       
       Exit Sub
       
       
       
    End If
    
    
    
    Dim rrF As New ADODB.Recordset, MARK2 As String
    rrF.Open "select ENTITYMARK,isnull(JWT,'') AS JWT FROM TIM WHERE ID_NUM=" + mID_NUM, Gdb, adOpenDynamic, adLockOptimistic
    Dim PAROXOS As Integer
    PAROXOS = 0
    If Not rrF.EOF Then
      If Left(rrF(0), 3) = "ΔΟΚ" Then
         MARK2 = CNull(mID(rrF(0), 4, 15))
      Else
         MARK2 = CNull(rrF(0))
      End If
      If Len(rrF("JWT")) > 20 Then
         MsgBox "ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ ΕΙΝΑΙ ME ΠΑΡΟΧΟ. ΔEN AKYΡΩNETAI"
         Exit Sub
      End If
      
    Else
      MARK2 = ""
    End If
    rrF.Close
    
    
    
    

 Dim c As String
  PAR1.TDBGrid1.Col = 1
    c = Replace(PAR1.TDBGrid1.Text, ",", ".")







Dim cHME As String
 PAR1.TDBGrid1.Col = 0
    cHME = Replace(PAR1.TDBGrid1.Text, ",", ".")
'den katalabainei oti einai benzinadiko

If f_is_benzinadiko = 1 Then
 
   If DateDiff("D", CDate(cHME), Now) > 0 Then
      MsgBox "Δεν θα ακυρωθεί το Παραστατικό γιατι δεν ειναι σημερινό"
      Exit Sub
   End If

End If




Dim ANS As Integer
ANS = MsgBox("Προσοχή θα ακυρωθεί το Παραστατικό " + c + " !!!" + Chr(13) + "Nα ακυρωθεί; ", vbYesNo)

If ANS = vbNo Then
    Exit Sub
End If

PAR1.PAR_GIA_AKYROSH.Caption = mID(c, 2, 7)

 Dim R As New ADODB.Recordset

     R.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + Left(c, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
     PAR1.PAR_GIA_AKYROSH.Caption = Trim(CNull(R!C1)) + "#" + mID(c, 2, 7)  ' 221#151   TYPOS ELINE+ARIUMOS
     R.Close
     


   
         
  '  PAR1.TDBGrid1.Col = 1
  '  C = Replace(PAR1.TDBGrid1.Text, ",", ".")
         
         Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)
       
        
  Dim ISOK As Integer
  ISOK = 1
  If m_akyr_neo = 0 Then  ' ME DIAGRAFH PALIOY
    ' METASX_DIORTOSIS
      PAR1.METASX_SE_TIM 99, c, m_ID, True, 0
  Else
     
     PAR1.neo_akyr m_ID, c, ISOK  'XORIS

  End If
      
      ' ΒΡΙΣΚΩ ΤΟ ΚΑΙΝΟΥΡΙΟ ID_NUM ΠΟΥ ΔΗΜΙΟΥΡΓΗΘΗΚΕ ΚΑΙ ΒΑΖΩ ΤΑ ΜΑΡΚ
      If Len(MARK2) > 3 And ISOK = 1 Then
      
              gisOnlineMydata = 1
            Dim ll As Long, CANCMARK As String
            
            rrF.Open "select max(ID_NUM) FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
                ll = rrF(0)
            rrF.Close
            Par7MyData.Visible = False
            Par7MyData.Caption = "" ' ΙΣΑ ΙΣΑ ΝΑ ΑΝΟΙΓΕΙ ΤΟ FORM_LOAD ΓΙΑ ΝΑ ΔIΑΒΑΣΕΙ P_USER
            
             CANCMARK = Par7MyData.CANCEL_INVOICE(MARK2, c)
              Gdb.Execute "UPDATE TIM SET ENTITYMARK='" + CANCMARK + "',SXETMARK='" + MARK2 + "' WHERE ID_NUM=" + str(ll)
             If Len(CNull(CANCMARK)) > 10 Then
               ' ΜΑΡΚ ΑΚΥΡΩΣΗΣ ΣΤΟ ΤΙΜΟΛΟΓΙΟ
               Gdb.Execute "UPDATE TIM SET INCMARK='" + CANCMARK + "'  WHERE ID_NUM=" + str(m_ID)
             End If
             
      End If

       
          Exit Sub

akyrotiko_Click_err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MDIForm1.metasx_par_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.akyrotiko_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>



End Sub

Private Sub apo2scroll_Click(index As Integer)

    On Error Resume Next

    If f_site = 333 Then
       ' apot2ROULISnew.SHOW
    ElseIf f_site = 334 Then
       ' apot2ROULIS.SHOW
    Else
      '  apot2vmn.SHOW
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

Private Sub apo4_Click(index As Integer)

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

Private Sub apo5_Click(index As Integer)

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

Private Sub apoemporiko_Click(index As Integer)

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

Private Sub apot_anaz_Click(index As Integer)
   ' apot2ROULISSEARCH.SHOW
  
    'apot_searching.SHOW

End Sub

Private Sub calculator_Click(index As Integer)

        '<EhHeader>
        On Error GoTo calculator_Click_Err

        '</EhHeader>
        Dim DUM

100     bohu10.SHOW


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

Private Sub calendar_Click()
Dim DUM
   On Error Resume Next
   DUM = Shell("C:\MERCVB\CALENDAR.EXE", vbMaximizedFocus)
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

Private Sub DBF_Click(index As Integer)

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

Private Sub diagrpar_Click(index As Integer)

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

    Dim c As String
         
    PAR1.TDBGrid1.Col = 1
    c = Replace(PAR1.TDBGrid1.Text, ",", ".")
    
    Dim mmark As String
    
    PAR1.TDBGrid1.Col = 8
    mmark = Replace(PAR1.TDBGrid1.Text, ",", ".")
    
    If Len(Trim(mmark)) > 2 Then
      If Left$(mmark, 2) = "ER" Then
      
      Else
        MsgBox "Aδύνατη η διόρθωση γιατί έχει πάρει MARK"
        Exit Sub
      End If
      
    End If
    
         
         Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)
       
        
         
         
    ' METASX_DIORTOSIS
    PAR1.METASX_SE_TIM 99, c, m_ID, False, 0
         
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

     PAR1.menu_ektyp_par 1
     Exit Sub

'        '<EhHeader>
'        On Error GoTo ektyp_par_Click_Err
'
'        '</EhHeader>
'
'        Dim m As Integer
'
'        Dim c As String
'
'        PAR1.tdbgrid1.Col = 2
'        c = Replace(PAR1.tdbgrid1.Text, ",", ".")
'
'
'           Dim m_ID As Long
'         PAR1.tdbgrid1.Col = 7
'       m_ID = Val(PAR1.tdbgrid1.Text)
'
'
'        PAR1.METASX_SE_TIM 99, c, m_ID, False, 0
'        ' PAR1.METASX_SE_TIM 99, C, 0, False
'
'        ' 100     PAR1.diortosis
'110     DoEvents
'
'120     PAR1.Command5_Click
'        PAR1.Text2(0).Text = "" ' για να μην διορθώση κατα λαθος
'        '<EhFooter>
'        Exit Sub
'
'ektyp_par_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.MDIForm1.ektyp_par_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.ektyp_par_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>

End Sub

Private Sub email_par_Click()
   PAR1.menu_ektyp_par 2
     Exit Sub
End Sub

Private Sub epanAkyr_Click()
   PAR1.epanAkyr
End Sub

Private Sub epit311_Click(index As Integer)

        '<EhHeader>
        On Error GoTo epit311_Click_Err

        '</EhHeader>

100     epit3.epit311_Click index

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

Private Sub epit312_Click()


'
'        '<EhHeader>
'        On Error GoTo epit311_Click_Err
'
'        '</EhHeader>
'
'100     epit3.epit312_Click Index
'
'        '<EhFooter>
'        Exit Sub
'
'epit312_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.MDIForm1.epit311_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.epit312_Click " & "at line " & Erl
'
'        Resume Next
'
'



End Sub

Private Sub exit_Click()
    End
Dim DUM

' DUM = Shell("C:\mercvb\bin\PROJECT1.EXE", vbMaximizedFocus)

Exit Sub




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

Private Sub genlog_Click(index As Integer)

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

Private Sub IMAGE2_Click()
  If Picture1.Width > 1999 Then
     Picture1.Width = 100
     ' Image1.Width = 100
  Else
        Picture1.Width = f_FARDOSIMAGE
     
  End If
      
End Sub

Private Sub Listview2_Click()

        '    Listview2.ListItems.Add , "a6", "Πίνακες", 16, 16
        '    Listview2.ListItems.Add , "a7", "Back-Up", 17, 17
        '<EhHeader>
        On Error GoTo Listview2_Click_Err

        '</EhHeader>
100     Select Case ListView2.SelectedItem.key

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

Private Sub mapot1_Click(index As Integer)

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

Private Sub mapot2_Click(index As Integer)

        '<EhHeader>
        On Error GoTo mapot2_Click_Err

        '</EhHeader>

100     gApoMenu = True

        '
        ' MDIForm1.'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", "NormalOffice2007.ini"
        ' MDIForm1.'SkinFramework1.ApplyWindow APOT2.hWnd
        ' MDIForm1.'SkinFramework1.ApplyOptions = MDIForm1.'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
        '
       
        
        
110    ' APOT2.WindowState = 2






      Dim fff As New APOT2
    fff.WindowState = 2



'fff.Hide



   '   Load fff
      
      
    '   fff.Hide
      
''      fff.Caption = ""
    '  fff.Hide
     'fff.Caption = ""
'      fff.Hide
'
'
      'MILSEC 2000
      
      
     
     ' fff.SHOW
      

       

120   '  APOT2.SHOW

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

Private Sub mapot31_Click(index As Integer)

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
FF.Rep_Fields.Text = "select KOD as [Κωδικός],LEFT(ONO,35) as [Περιγραφή],POS AS [Υπόλοιπο]"
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

Private Sub mapot32c_Click()
'PAROYSIASH apografhS απο περυσι
'
  ''ΥΠΟΛΟΙΠΑΟΘΗΚΩΝ
'Dim FF As New apot3param
' FF.Label_titlos = "Aπογραφή Ειδών"
'
'FF.load_sthles_EIDON
''FF.LOAD_KINHSEIS_EIDON
'FF.Rep_Fields.Text = "select KOD as [Κωδικός],LEFT(ONO,35) as [Περιγραφή],POS AS [Υπόλοιπο]"
'FF.gemisma_Krithrion
'fkrithria.Grid1.TextMatrix(4, 1) = 0.1
'fkrithria.Grid1.TextMatrix(4, 2) = 999999999
'fkrithria.ΕΧΙΤ_Click
'FF.sumes.Text = "001"
'FF.WHERE_QUERY.Text = "POS>=0.1 AND POS<=999999999" ' StatusBar1.Panels(6).Text"
''YPOLOGISMOS_Click
'FF.SHOW
''FF.YPOLOGISMOS_Click

'ΥΠΟΛΟΙΠΑΟΘΗΚΩΝ
Dim FF As New apot7
 FF.Caption = "Aπογραφή Ειδών"
 FF.Plath_Grid(0) = 1820
 FF.Plath_Grid(1) = 5020
 FF.Plath_Grid(2) = 580
 FF.Plath_Grid(3) = 580
 FF.Plath_Grid(4) = 580
 FF.Plath_Grid(5) = 580
 FF.sumstring.Text = "000101"
 
FF.Text2 = "select LEFT(KODE,20) AS [ΚΩΔΙΚΟΣ],EID.ONO AS [ΠΕΡΙΓΡΑΦΗ] ,EID.MON AS [ΜΟΝ.ΜΕΤ],sum(XRE) AS [ΠΟΣΟΤΗΤΑ],TIMM AS [TIMH],sum(round(XRE*TIMM,2)) AS [ΑΞΙΑ],'' AS [-] FROM EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE LEFT(ATIM,1)='λ' group by LEFT(KODE,20),EID.ONO,EID.MON,TIMM order by EID.ONO"
FF.CMDPROVOLI_Click
FF.SHOW


End Sub

Private Sub mapot33_Click()

'isozygio ΑΠΟΘΗΚΩΝ
Dim FF As New apot3param
Dim sql As String

FF.load_sthles_EIDON
'FF.LOAD_KINHSEIS_EIDON
Dim Apog1PosT As String
Dim Apog1AjiaT As String
Apog1PosT = "  (CASE APOG01 WHEN NULL THEN 0 ELSE APOG01 END) "
Apog1AjiaT = " (CASE KOSAPOG01 WHEN NULL THEN 0 ELSE KOSAPOG01 END) "

sql = "select KOD as [Κωδικός],LEFT(ONO,35) as [Περιγραφή],"
sql = sql + Apog1PosT + " AS [Απογ.Απ1]," + Apog1AjiaT + " AS [Αξ.Περ.Απογ.Απ1],"
sql = sql + " ISNULL(S1X-APOG01,0) AS [Εισ1.Περιόδ], AS1X   AS [Αξ1.Αγ.Περιόδ],"
sql = sql + "ISNULL(S1P,0)  AS [Εξ1.Περιόδ],AS1P   AS [Αξ1.Εξ.Περιόδ],"
sql = sql + "XPOS01 as [Υπ.Απ1],"
sql = sql + " ( CASE WHEN   ISNULL(APOG01,0) +S1X-ISNULL(APOG01,0) >0 THEN  ( AS1X + ISNULL(KOSAPOG01,0)  )/(  S1X+ ISNULL(APOG01,0)- ISNULL(APOG01,0) ) ELSE 0 END)  "
sql = sql + " AS [Μέση Τιμή],ROUND( ( CASE WHEN ISNULL(APOG01,0) +S1X-ISNULL(APOG01,0)>0 THEN (AS1X+ISNULL(KOSAPOG01,0) ) /(S1X-ISNULL(APOG01,0)+ISNULL(APOG01,0) )  ELSE 0 END) ,3)*POS AS [ΑΞΙΑ ΥΠΟΛ],ISNULL(APOG01,0) +S1X AS [POS AGOR],ISNULL(KOSAPOG01,0) +AS1X AS [AJIA AGOR] "
FF.sumes.Text = "0011111111"
FF.Rep_Fields.Text = sql
FF.gemisma_Krithrion
'fkrithria.Grid1.TextMatrix(4, 1) = 0.1
'fkrithria.Grid1.TextMatrix(4, 2) = 999999999
'fkrithria.ΕΧΙΤ_Click
 ' FF.WHERE_QUERY.Text = "Apog1PosT+S1X+S1P>=0 AND ISNULL(S1X-APOG01,0)+ISNULL(S1P,0)   >0 "
 FF.WHERE_QUERY.Text = " EID.KOD IN (SELECT KODE FROM EGGTIM)  "

'FF.WHERE_QUERY.Text = "[Εισ1.Περιόδ] + [Εξ1.Περιόδ]>0 "
FF.Check1.Value = vbChecked
FF.Text2 = "ΙΣΟΖΥΓΙΟ"

FF.Label_titlos = "Ισοζύγιο Αποθηκών"
FF.ypologizo_Ypoloipa (True)

FF.SHOW

End Sub



Private Sub mapot34_Partides_Click()
  'frmPARTIDES.SHOW
'================================================================================
        Dim F As New DataEntry3
        Dim R As New ADODB.Recordset
        Load F

F.Caption = "Παρτίδες"

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
        F.ParamGrid.TextMatrix(0, 1) = "SELECT  * FROM PARTIDES " '    sql

        '------------------------------------------------------------------------------------
        'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
        F.ParamGrid.TextMatrix(0, 2) = "SELECT ID,KOD,LOTNUMBER,HMEL,POSOAG " & " FROM PARTIDES ORDER BY ID DESC"    'ParamGrid   sql2
'110     Sql2 = ParamGrid.TextMatrix(0, 2)    ' string adodc1
        '------------------------------------------------------------------------------------
        'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
        F.ParamGrid.TextMatrix(0, 3) = "delete from PARTIDES where  "
        
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
170     F.ParamGrid.TextMatrix(0, 4) = ""  'DD

        F.ParamGrid.TextMatrix(1, 1) = "ID εγγραφής"
        F.ParamGrid.TextMatrix(1, 2) = 1  ' ARITHMOS
        
        
180     F.ParamGrid.TextMatrix(2, 1) = "Κωδ.Είδους"
        F.ParamGrid.TextMatrix(2, 2) = 2 ' STRING


190     F.ParamGrid.TextMatrix(3, 1) = "LOT NUMBER"
        F.ParamGrid.TextMatrix(3, 2) = 2 ' STRING
        
        
        F.ParamGrid.TextMatrix(4, 1) = "Ημερ.Λήξης"
        F.ParamGrid.TextMatrix(4, 2) = 3 ' DATE
        
        F.ParamGrid.TextMatrix(5, 1) = "Ποσ.Αγοράς"
        F.ParamGrid.TextMatrix(5, 2) = 1 ' STRING



        'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
200     F.Labels(0).Width = F.Labels(0).Width + 1000    '* TextWidth("A")
210     F.Textboxes(0).Left = F.Textboxes(0).Left + 1000    ' * TextWidth("A")
        '------------------------------------------------------------------------------------
Dim OO
        'αν παει η φόρμα να φύγει από την οθόνη τότε να κεντραρεται δεξια με την mdiform1
220    ' If OO.Left + f.Width > MDIForm1.Left + MDIForm1.Width Then
230     '    f.Left = MDIForm1.Left + MDIForm1.Width - f.Width
       ' Else
240    '     f.Left = OO.Left
      '  End If

250  '   f.Top = OO.Top + MDIForm1.Top

260     F.SHOW 1

Gdb.Execute "update PARTIDES SET YPOLOIPO=POSOAG WHERE YPOLOIPO IS NULL"
  

End Sub

'Private Sub mapot32_Click(Index As Integer)
''apot32.SHOW
'
'End Sub

Private Sub mapot6_Click(index As Integer)

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

Private Sub mapot7_Click(index As Integer)

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

Private Sub mbohu1_Click(index As Integer)

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

Private Sub mBohu14_Click()
Dim R As New ADODB.Recordset
R.Open "SELECT SUM(ISNULL(N1,0) ) FROM PINAKES WHERE TYPOS=1", Gdb, adOpenDynamic, adLockOptimistic
If R(0) > 1 Then
    MsgBox "Εχει τρέξει ήδη ο οδηγός"
    Exit Sub

End If
R.Close

'r.Open "SELECT ISNULL(AYJON,0) AS AYJ,ISNULL(PERIGRAFH,'') AS PER FROM PINAKES WHERE AYJON>0 AND TYPOS=12", Gdb, adOpenDynamic, adLockOptimistic
'Dim CC  As String
'Do While Not r.EOF
'
'  CC = Format(r!AYJ, "#0")
'
'    MsgBox "Εχει τρέξει ήδη ο οδηγός"
'    Exit Sub
'
'  End If
'  r.MoveNext
'Loop
'r.Close




'SELECT SUM(ISNULL(N1,0) ) FROM PINAKES WHERE TYPOS=1

Dim sql As String
sql = "UPDATE PINAKES SET N1=1  WHERE TYPOS=1 AND TIMH=24;UPDATE PINAKES SET N1=2  WHERE TYPOS=1 AND TIMH=13;"
sql = sql + "UPDATE PINAKES SET N1=3  WHERE TYPOS=1 AND TIMH=6;UPDATE PINAKES SET N1=4  WHERE TYPOS=1 AND TIMH=17 ;"

sql = sql + "UPDATE PINAKES SET N1=5  WHERE TYPOS=1 AND TIMH=9;UPDATE PINAKES SET N1=6 WHERE TYPOS=1 AND TIMH=4;"

sql = sql + "UPDATE PINAKES SET N1=7 WHERE TYPOS=1 AND TIMH=0;"
Gdb.Execute sql

MsgBox "φπα οκ"

Gdb.Execute "UPDATE PARASTAT SET ISMYDATA=1, MYEID='1.1',MYTYPESOD='E3_561_001',MYCATEG='category1_1' WHERE EIDOS IN ('T','t');"
Gdb.Execute "UPDATE PARASTAT SET ISMYDATA=1, MYEID='2.1',MYTYPESOD='E3_561_001',MYCATEG='category1_3' WHERE EIDOS IN ('Y');"
Gdb.Execute "UPDATE PARASTAT SET ISMYDATA=1, MYEID='5.1',MYTYPESOD='E3_561_001',MYCATEG='category1_1' WHERE EIDOS IN ('P');"
Gdb.Execute "UPDATE PARASTAT SET ISMYDATA=1, MYEID='11.1',MYTYPESOD='E3_561_003',MYCATEG='category1_1' WHERE EIDOS IN ('L','l');"
Gdb.Execute "UPDATE PARASTAT SET ISMYDATA=1, MYEID='11.2',MYTYPESOD='E3_561_003',MYCATEG='category1_3' WHERE EIDOS IN ('y');"
Gdb.Execute "UPDATE PARASTAT SET ISMYDATA=1, MYEID='11.4',MYTYPESOD='E3_561_003',MYCATEG='category1_1' WHERE EIDOS IN ('p');"

MsgBox "Παραστατικά οκ"

Gdb.Execute "UPDATE PINAKES SET N1=3  WHERE TYPOS=12 AND PERIGRAFH  LIKE '%ΜΕΤΡ%';"
Gdb.Execute "UPDATE PINAKES SET N1=5  WHERE TYPOS=12 AND PERIGRAFH  LIKE '%ΠΙΣΤΩΣ%';"
Gdb.Execute "UPDATE PINAKES SET N1=1  WHERE TYPOS=12 AND PERIGRAFH  LIKE '%ΚΑΡΤ%';"

MsgBox "Τρόποι πληρωμής οκ"

Gdb.Execute "DELETE FROM PINAKES WHERE TYPOS=15;"
Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH,LOG1,C1) VALUES (15,1,'ΕΜΠΟΡΕΥΜΑ',0,'category2_1','category1_1')"
Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH,LOG1,C1) VALUES (15,2,'ΠΡΟΙΟΝ',0,'category2_1','category1_2')"
Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH,LOG1,C1) VALUES (15,3,'ΥΠΗΡΕΣΙΑ',0,'','category1_3')"
Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH,LOG1,C1) VALUES (15,4,'ΠΑΓΙΟ',0,'','category1_4')"

MsgBox "Κατηγορίες οκ"





End Sub

Private Sub mbohu2_Click(index As Integer)

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
110         f_supervisor_code = FINDPARAMETROI(1, "MDIFORM1", "f_supervisor_code", "0", "Κωδικός")
120         X = f_supervisor_code
130         PARAMETROI.PARAM.Caption = "MDIFORM1"
140         PARAMETROI.SHOW 1
150         f_supervisor_code = FINDPARAMETROI(1, "MDIFORM1", "f_supervisor_code", "0", "Κωδικός")
            
160         If f_supervisor_code = X Then
            Else

170             If IsNull(f_supervisor_code) Then f_supervisor_code = ""
180             Gdb.Execute "UPDATE DIK SET CODE='" + f_supervisor_code + "' WHERE  LEFT(KOD,6)='070109' ", a

190             If a > 0 Then
200                 MsgBox "ενημερώθηκε ο νέος κωδικός"
                End If
            End If

               If F_BACKGROUND > 0 Then
                  On Error Resume Next 'GoTo ERRORS
            '      Kill "C:\MERCVB\BACKGROUND.BMP"
            '      FileCopy "C:\MERCVB\BACKGROUND" + LTrim(Str(F_BACKGROUND)) + ".BMP", "C:\MERCVB\BACKGROUND.BMP"
            '      MILSEC 2000
                  gPicture = "C:\MERCVB\BACKGROUND.BMP"
                  MDIForm1.Picture = LoadPicture(gPicture)
            '
               End If

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


'Private Sub mdiForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If Shift = 1 Then
'     Dim oleNewColor As OLE_COLOR
'    oleNewColor = ShowColorDialog(Me.hwnd, True, vbRed)
'   ' If oleNewColor <> -1 Then ' Grid1.BackColor = oleNewColor
'
'
'
'    Gdb.Execute "UPDATE PARAMETROI SET TIMH=" + str(oleNewColor) + " WHERE  FORMA='MDIFORM1' AND VAR='F_COLOR_BACKGROUND' "
'
'    End If
'
'End Sub


Private Sub MDIForm_Load()



'Dim bridge As ClassLibrary1.JsonTools
'Set bridge = New ClassLibrary1.JsonTools  ' Στη VB6 το Set είναι υποχρεωτικό
''Dim dict As New Scripting.Dictionary
''dict.Add "ID", 101, dict.Add "Status", "Active"
' Dim jsonResult As String
' jsonResult = bridge.Serialize(dict)
''MsgBox jsonResult







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


'      On Error Resume Next
'        F_COLOR_BACKGROUND = Val(FINDPARAMETROI(1, "MDIFORM1", "F_COLOR_BACKGROUND", "1", "ΧΡΩΜΑ BACKGROUND ΜΠΛΕ=1,ΠΡΑΣ=2,ΚΑΦΕ=3,ΜΩΒ=4,ΡΟΖ=5"))
'
'If F_COLOR_BACKGROUND = 1 Then
'    Me.BackColor = vbBlue '  F_COLOR_BACKGROUND
'ElseIf F_COLOR_BACKGROUND = 2 Then
'    Me.BackColor = vbGreen
'ElseIf F_COLOR_BACKGROUND = 3 Then  'KAFE
'    Me.BackColor = &H8080&
' ElseIf F_COLOR_BACKGROUND = 4 Then  'MVB
'    Me.BackColor = &HFFC0C0
' Else
'
'    Me.BackColor = &HC0C0FF    'POZ
' End If
'
 
                '&H00FFC0C0&
    






gisOnlineMydata = 0




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

140     gPicture = "" '"C:\MERCVB\background.BMP"
150     MDIForm1.Picture = LoadPicture(gPicture)

        ' Dim dum: dum = printing_etik("354454534534", 2, "tem", 1.31, 1.34, 1.23, 9, "lagakis", 67)

160     MDIForm1.SHOW

        On Error GoTo SEELINE

175     minima = "ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΑΡΧΕΙΟ C:\MERCPATH.TXT"

        If Len(Dir(App.Path + "\MERCPATH.TXT")) > 0 Then

180         Open App.Path + "\MERCPATH.TXT" For Input As #1
190         Line Input #1, gDir
200         Line Input #1, gConnect
210         Close #1
        Else

211         Open "C:\MERCPATH.TXT" For Input As #1
            Line Input #1, gDir
            Line Input #1, gConnect
212         Close #1

        End If
        
        If Len(Dir("C:\MERCVB\MERCCONFIG.INI")) > 0 Then
           Line Input #1, gConnect
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

220   '  MDIForm1.StatusBar1.Panels(1).Text = gDir  ' "Τηλέφωνα" '
230     MDIForm1.StatusBar1.Panels(3).Text = "MercuryΔΑ v." + Format(App.Major, "00") + "." + Format(App.Minor, "00") + "." + Format(App.Revision, "00")
235     MDIForm1.StatusBar1.Panels(6).Text = "Lagakis +30 2521022225"

        'Exit Sub

240     pelat.Enabled = False
250     apot.Enabled = False
260     epit.Enabled = False
270     par.Enabled = False
280     bohu.Enabled = False

        ' gConnect = "Provider=SQLOLEDB.1;Password=38983;Persist Security Info=True;User ID=sa;Data Source=GEORGE;Initial Catalog=MERCURY;"
290     frmLogin.SHOW 1
 
 
 
 
 
 
 
 
 
  
 
 
 
     
 
      F_UPDATESTRUCTURES = Val(FINDPARAMETROI(1, "MDIFORM1", "F_UPDATESTRUCTURES", "1", "ΝΑ ΚΑΝΕΙ UPDATE STRUCTURES=1 OXI=0"))
 
 
        f_site = Val(FINDPARAMETROI(1, "MDIFORM1", "F_SITE", "0", "ΚΩΔΙΚΟΣ SITE 1=VM 2=TOYS"))

        Dim F_fontheight As Integer
        
        F_fontheight = Val(FINDPARAMETROI(1, "MDIFORM1", "F_fontheight", "10", "ΜΕΓΙΣΤΟ ΥΨΟΣ FONT(ΓΡΑΜΜΑΤΟΣΕΙΡΑΣ) ΜΕΤΑ ΤΗΝ ΜΕΓΙΣΤΟΠΟΙΗΣΗ(10-14)"))
        
        
         f_proeid_mydata = Val(FINDPARAMETROI(1, "MDIFORM1", "f_proeid_mydata", "1", "Προειδοποίηση για μη απεσταλμένα MyData=1,Οχι=0"))
        
        
        If f_site = 111 Or f_site = 222 Or f_site = 333 Or f_site = 444 Then
            apo2scroll(28).Visible = True
            apot_anaz(29).Visible = True
        End If

295     MDIForm1.StatusBar1.Panels(7).Text = "SN:" + FINDPARAMETROI(1, "SN", "SN", "0", "ΣEIRIAKOΣ ΑΡΙΘΜΟΣ")


300     Timer1.Enabled = True

310     If gConnect = "end" Then
320         Unload Me

330         End

        End If



340     F_BACKGROUND = Val(FINDPARAMETROI(1, "MDIFORM1", "F_BACKGROUND", "1", "BACKGROUND ΑΡ.ΕΙΚΟΝΑΣ=1,2,3,.. ΟΧΙ=0"))





350     f_FARDOSIMAGE = Val(FINDPARAMETROI(1, "MDIFORM1", "f_FARDOSIMAGE", "3040", "ΠΛΑΤΟΣ ΓΡΗΓΟΡΟΥ ΜΕΝΟΥ(ΑΡΧΙΚΟ=3040)"))

360     Picture1.Width = f_FARDOSIMAGE


        'If f_FARDOSIMAGE > 0 Then
        '  Picture1.Visible = True
        'End If

370     Listview1.ListItems.Add , "a1", "Καρτέλλα", 2, 2
380     Listview1.ListItems.Add , "a2", "Διόρθ.Πελάτη", 3, 3
390     Listview1.ListItems.Add , "a3", "Εισ.Παρ/κου", 6, 6
400     Listview1.ListItems.Add , "a4", "Παραμ.Παραστατικων", 15, 15
410     Listview1.ListItems.Add , "a5", "Πωλήσεις/Αγορές", 16, 16
420     Listview1.ListItems.Add , "a6", "Προβολη Παρ/κού", 17, 17
430     Listview1.ListItems.Add , "a7", "Υπόλοιπα Πελατών", 18, 18
        'Listview1.ListItems.Add , "b3", "Aναζήτηση", 5, 5
    
440     ListView2.ListItems.Add , "b1", "Πίνακες", 1, 1
450     ListView2.ListItems.Add , "b2", "Back-Up", 17, 17
        'Listview2.ListItems.Add , "a8", "Business Info", 25, 25
460     ListView2.ListItems.Add , "b3", "Για το πρόγραμμα ", 18, 18

470     pelat.Enabled = True
480     apot.Enabled = True
490     epit.Enabled = True
500     par.Enabled = True
510     bohu.Enabled = True

520     If F_BACKGROUND = 1 Then
            ' OK ΑΦΗΝΩ ΤΟ ΑΡΧΕΙΟ ΩΣ ΕΧΕΙ
            gPicture = "C:\MERCVB\background.BMP"
        Else
530         gPicture = ""
540         MDIForm1.Picture = LoadPicture(gPicture)
        End If

        ' Gdb.Open gConnect

        ' f_loadpar1 = Val(FINDPARAMETROI(1,"MDIFORM1", "F_LOADPAR1", "1", "ΦΟΡΤΩΝΩ ΚΑΘΕ ΦΟΡΑ ΤΗΝ ΦΟΡΜΑ ΠΑΡΑΣΤΑΤΙΚΩΝ=1"))

550     f_telephone = Val(FINDPARAMETROI(1, "MDIFORM1", "F_TELEPHONE", "0", "ΕΜΦΑΝΙΣΗ ΚΑΛΟΥΝΤΩΝ ΤΗΛΕΦΩΝΩΝ=1"))


        f_auto_load_times = Val(FINDPARAMETROI(1, "MDIFORM1", "F_AUTO_LOAD_TIMES", "0", "ΑΥΤΟΜΑΤΗ ΦΟΡΤΩΣΗ ΤΙΜΩΝ ΥΠΟΚΑΤΑΣΤΗΜΑΤΟΣ"))
        
        fgAsterisk = Val(FINDPARAMETROI(1, "MDIFORM1", "fgAsterisk", "0", "Το * σημαινει οτιδηποτε στην αναζητηση=1 οχι=0 "))
        
        gAsterisk = fgAsterisk
        
        
        F_TAKENEESTIMES = Val(FINDPARAMETROI(1, "BOHU7", "F_TAKENEESTIMES", 0, "ΕΝΗΜΕΡΩΣΗ ΤΙΜΩΝ ΑΠΟ ΙΝΤΕΡΝΕΤ 0=OXI 1=NAI 2=ΟΛΑ ΕΛΕΥΘΕΡΑ"))
        ' αν το τερματικο δεν εχει δικαιωμα να παιρνει τιμες μπλοκαρει το αυτοματο
        If F_TAKENEESTIMES = 0 Then
           f_auto_load_times = 0
        End If
        




560     F_TOOLBAR = Val(FINDPARAMETROI(1, "MDIFORM1", "F_TOOLBAR", "550", "ΥΨΟΣ TOOLBAR"))
570     f_Ypenth = Val(FINDPARAMETROI(1, "MDIFORM1", "F_YPENTH", "1", "ΥΠΕΝΘΥΜΙΣΗ=1 ΟΧΙ=0"))
        'f_supervisor_code = Val(FINDPARAMETROI(1,"MDIFORM1", "f_supervisor_code", "1", "Κωδικός"))

580     tlbToolBar.Height = F_TOOLBAR

        On Error GoTo stopthlefona

590     If f_telephone = 1 Then
600         TELEFON.ConnectionString = gConnect
610         TELEFON.RecordSource = "SELECT TOP 1 * FROM THLEFNOW"
620         TELEFON.Refresh
        End If

630     If F_TOOLBAR = 0 Then
640         tlbToolBar.Visible = False

650        ' StatusBar1.Visible = False
        End If
        
        
        If F_TOOLBAR < 0 Then
            tlbToolBar.Visible = False
            StatusBar1.Visible = False
        End If
        
        
        

        On Error Resume Next

        'KANEI TO UPDATE STO FRM LOGIN OTAN ALLAZEI VERSION
        'Update_Structures

660     If gXEIRISTHS >= 8 Then
670         MakeBackup ""
        End If
        
        
Dim f_IS_PARTIDES As Integer
f_IS_PARTIDES = Val(FINDPARAMETROI(1, "APOT2", "f_IS_PARTIDES", "0", "Αν Χρησιμοποιώ παρτίδες=1"))
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



















730     If gXEIRISTHS <= 9 Then

  Dim N As Integer
 '================================ PRODUCT_TO_CATEGORIES ================================
     Dim RR As New ADODB.Recordset
     
     
        RR.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'MONADES'", Gdb, adOpenDynamic, adLockOptimistic
          'On Error Resume Next
          If RR(0) = 0 Then
             'On Error Resume Next
             'Gdb.Execute "INSERT INTO ACTIVE_MENUS (MENU_ID,NAME,ENABLED,USER_ID) VALUES (" + str(N) + ",'" + Left(MDIForm1.Controls(N).Caption, 35) + "',1,1)"
          
             apot7.CREATE_MONADES
          End If
        RR.Close
     
     
     
     
     
     
     
     RR.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'ACTIVE_MENUS'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next
         If RR.EOF Then
            'On Error Resume Next
            Gdb.Execute "INSERT INTO ACTIVE_MENUS (MENU_ID,NAME,ENABLED,USER_ID) VALUES (" + str(N) + ",'" + Left(MDIForm1.Controls(N).Caption, 35) + "',1,1)"
         End If
         RR.Close

On Error GoTo OUT

On Error Resume Next

Dim R99 As New ADODB.Recordset
  ' Dim n As Integer
 '================================ PRODUCT_TO_CATEGORIES ================================
     R99.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'ACTIVE_MENUS'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

     If R99(0) = 0 Then
        Gdb.Execute "CREATE TABLE ACTIVE_MENUS ( ID INT IDENTITY(1,1),ENABLED INT NULL,KOD VARCHAR(16) NULL,MENU_ID int  NULL,NAME VARCHAR(35) NULL ,USER_ID INT null) "
     End If

        R99.Close




' MDIForm1.Controls("MTAMEIAKES").Visible = False
         Dim r1 As New ADODB.Recordset, M_ENABLE As Boolean
           r1.Open "SELECT * FROM ACTIVE_MENUS WHERE USER_ID=" + str(gUserId), Gdb, adOpenDynamic, adLockOptimistic
          ' If R1.RecordCount > 0 Then
           Do While Not r1.EOF
                If r1("ENABLED") = 1 Then M_ENABLE = True Else M_ENABLE = False
                MDIForm1.Controls(r1("MENU_ID")).Enabled = M_ENABLE
                If M_ENABLE = False Then
                    MDIForm1.Controls(r1("MENU_ID")).Visible = M_ENABLE
                    
                    If r1("MENU_ID") = 5 Then ' καρτελλα πελατη
                    
                    
                    
                    End If
                    
                End If
                r1.MoveNext
           Loop
          ' End If
           
           r1.Close
             



'740         mapot1(21).Visible = False
'750         pareid(23).Visible = False
'760         apo4(24).Visible = False
'770         apo5(25).Visible = False
'780         mapot6(26).Visible = False
'790         EPIT.Visible = False
'
'800         mBohu1(51).Visible = False
'810         mbohu2(52).Visible = False
'
'820         par3(53).Visible = False
'830         parkin(54).Visible = False
'840         pros(55).Visible = False
'850         genlog(56).Visible = False
'
'860         If gXEIRISTHS <= 2 Then
'
'870             mapot2(22).Visible = False
'880             apoemporiko(1).Visible = False    'True
'890             DBF(57).Visible = False
'                ' pontoif(58).Visible = False
'900             apoemporiko(59).Visible = False
'910             calculator(510).Visible = False
'920             mpelat1.Visible = False
'930             mpelat2.Visible = False
'940             mpelat3.Visible = False
'950             mpelat4.Visible = False
'960             mpelat5.Visible = False
'970             APOT.Visible = False
'
'980             mapot7(27).Visible = False
'990             mpar7.Visible = False
'1000            mBohu13.Visible = False
'
'1010            diagrpar(42).Visible = False
'1020            parenos(43).Visible = False
'1030            mpar5.Visible = False
'
'1040            If gXEIRISTHS = 2 Then
'                Else
'1050                mPar6.Visible = False
'                End If
'
'1060            mbohu12.Visible = False
'1070            timspar(44).Visible = False
'
'1080            mapot2(22).Visible = False
'
'1090            mapot7(27).Visible = False
'
'            End If
'
        Else



        End If
1100        If f_Ypenth = 1 Then
                  g_Logot = 0
                  utPROEID.SHOW
            ElseIf f_Ypenth = 2 Then
                 g_Logot = 0 '1
                  utPROEID.SHOW
            Else
            
            End If
Dim sql As String
If f_proeid_mydata = 1 Then
  sql = "SELECT COUNT(*) FROM TIM WHERE (ENTITYMARK IS NULL OR LEFT(ENTITYMARK,3)='ERR' ) AND    "
  sql = sql + " LEFT(ATIM,1) IN     (   SELECT EIDOS FROM PARASTAT WHERE ISMYDATA=1  )    and "
  sql = sql + " DATEDIFF(day, HME, GETDATE() ) < 31  " '   AND ISNULL(AJ1,0)+ISNULL(AJ2,0)+ISNULL(AJ3,0)+ISNULL(AJ4,0)+ISNULL(AJ5,0)+ISNULL(AJ6,0)+ISNULL(AJ7,0) >0 "
  Dim NV As Long: NV = GGET_NVALUE(sql)
  If NV > 0 Then
     MsgBox "Υπάρχουν τις τελευταίες 30 ημέρες  " + str(NV) + " εγγραφες που δεν εστάλησαν στο Mydata" + Chr(13) + " Πάτε :" + Chr(13) + "Παραστατικά=> 7.Μydata και πατηστε" + Chr(13) + "<Aποστολη Παραστατικών στο Mydata>"
  End If
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

1130    LOADPARAMETERS 1

1140    Me.Caption = "Menu"

        On Error GoTo CREATEPARAMETROI

1150    R.Open "select * from PARAMETROI", Gdb, adOpenForwardOnly, adLockReadOnly

1160    R.Close

        Dim SQL2, SQL3 As String

        On Error GoTo CREATEXRHSEIS

1170    SQL3 = "select * from XRHSEIS where '" + Format(gWorkDay, "MM/DD/YYYY") + "'>=ENARXH AND '" + Format(gWorkDay, "MM/DD/YYYY") + "'<=LHXH"

        On Error GoTo MDIForm_Load_Err

1180    R.Open SQL3, Gdb, adOpenForwardOnly, adLockReadOnly
    
1190    If R.EOF Then
            'BAZΩ 1/1/2000 ΓΙΑ ΝΑ ΜΗΝ ΒΓΑΖΕΙ ΛΑΘΟΣ ΥΠΟΛΟΙΠΑ
1200        SQL2 = "INSERT INTO XRHSEIS (ETOS,ENARXH,LHXH) VALUES (" + LTrim(str(Year(gWorkDay))) + ",'01/01/2000','12/31/" + LTrim(str(Year(gWorkDay))) + "')"
1210        Gdb.Execute SQL2
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
        
        
        Call GetDriveList
        
        MDIForm1.StatusBar1.Panels(1).Text = GGET_CVALUE("SELECT TOP 1 PELONO FROM MEM")
        
        
        
        
        On Error Resume Next
        F_COLOR_BACKGROUND = Val(FINDPARAMETROI(1, "MDIFORM1", "F_COLOR_BACKGROUND", "1", "ΧΡΩΜΑ BACKGROUND ΜΠΛΕ=1,ΠΡΑΣ=2,ΚΑΦΕ=3,ΜΩΒ=4,ΡΟΖ=5"))

If F_COLOR_BACKGROUND = 1 Then
    Me.BackColor = vbBlue '  F_COLOR_BACKGROUND
ElseIf F_COLOR_BACKGROUND = 2 Then
    Me.BackColor = vbGreen
ElseIf F_COLOR_BACKGROUND = 3 Then  'KAFE
    Me.BackColor = &H8080&
 ElseIf F_COLOR_BACKGROUND = 4 Then  'MVB
    Me.BackColor = &HFFC0C0
 Else

    Me.BackColor = &HC0C0FF    'POZ
 End If

'
'                '&H00FFC0C0&
'
    
    
    
        
        
        
        
        
        
  
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

OUT:
Exit Sub


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
    If MDIForm1.Width > 150000 Then

        'StretchSourcePictureFromPicture MDIForm1.Picture2.Picture, MDIForm1.Picture2
        '            MDIForm1.Picture2.Refresh

        MDIForm1.BackColor = &HFF0000
        MDIForm1.Picture = LoadPicture()
    Else
        MDIForm1.Picture = LoadPicture(gPicture)
    End If

    'mainpic.Width = MDIForm1.Width
    'mainpic.Height = MDIForm1.Height
    tvDir.Height = Picture1.Height
    
   ' mainpic.Picture = LoadPicture(gPicture)

    'If MDIForm1.WindowState = 2 Then MDIForm1.WindowState = 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

  ' Signal Application Shutdown
 '   g_InteropToolbox.EventMessenger.RaiseApplicationShutdownEvent
 End
 

End Sub

Private Sub menu_diak_status_Click()
    PAR1.menu_diak_status
End Sub

Private Sub menuParadosis_Click()
    PAR1.menuParadosis
    
End Sub

Private Sub menuToTransport_Click()
    PAR1.menuToTransport (0)
End Sub

Private Sub METASX_DA_Click()
Dim CC As Integer
 CC = Val(FINDPARAMETROI(3, "PAR1", "f_TIMES_METASX_DA", "0", "TO METAΣΧ.(DA) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))
' CC = Val(FINDPARAMETROI(3, "PAR1", "f_TIMES_METASX_TIM", "0", "TO METAΣΧ.(TIM) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))
'  f_TIMES_METASX_LIAN = Val(FINDPARAMETROI(3, "PAR1", "f_TIMES_METASX_LIAN", "0", "TO METAΣΧ.(LIAN) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))
 
 
 
        Dim m As Integer

        Dim c As String

        PAR1.TDBGrid1.Col = 2
        c = Replace(PAR1.TDBGrid1.Text, ",", ".")
        m = Val(FINDPARAMETROI(9, "PAR1", "METASX_DA", "3", "Α/Α Δ.Α. στα παραστατικά ΔΑ(ΜΕΤΑΣΧ.ΣΕ ΔΑ)0-99"))
100     PAR1.metasx.Value = vbChecked
        
        
         Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
         m_ID = Val(PAR1.TDBGrid1.Text)
         PAR1.METASX_SE_TIM m, c, m_ID, False, CC
         

'        PAR1.METASX_SE_TIM m, C, 0

        Exit Sub

End Sub

Private Sub metasx_lian_Click()
Dim CC As Integer
 'f_TIMES_METASX_DA = Val(FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_DA", "0", "TO METAΣΧ.(DA) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))
' CC = Val(FINDPARAMETROI(3, "PAR1", "f_TIMES_METASX_TIM", "0", "TO METAΣΧ.(TIM) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))
CC = Val(FINDPARAMETROI(3, "PAR1", "f_TIMES_METASX_LIAN", "0", "TO METAΣΧ.(LIAN) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))




        '<EhHeader>
        On Error GoTo metasx_LIAN_Click_Err

        '</EhHeader>
        
        Dim m As Integer

        Dim c As String
         
        PAR1.TDBGrid1.Col = 2
        c = Replace(PAR1.TDBGrid1.Text, ",", ".")
        'PAR1.Text5.Text = "*" + C
        
        m = Val(FINDPARAMETROI(9, "PAR1", "METASX_LIAN", "3", "Α/Α ΔΛ.ΔΑ στα παραστατικά (Μετασχ.σε Λιανική)0-99"))
100     PAR1.metasx.Value = vbChecked


         Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)
        PAR1.METASX_SE_TIM m, c, m_ID, False, CC
        

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

        Dim c As String

        PAR1.TDBGrid1.Col = 2
        c = Replace(PAR1.TDBGrid1.Text, ",", ".")
        
        PAR1.metasx.Value = vbChecked
        PAR1.METASX_SE_TIM 99, c, 0, False, 0
        

        Exit Sub

        '----------------------------------------------------------------------------------------
        
        '<EhHeader>
        On Error GoTo metasx_par_Click_Err

        '</EhHeader>

        'Dim m As Integer

        m = Val(FINDPARAMETROI(9, "PAR1", "METASX_TIM", "2", "Παραστατικό ΤΠΔΑ (Μετασχ.σε Τιμολ)0-99"))
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
Dim CC As Integer
 'f_TIMES_METASX_DA = Val(FINDPARAMETROI(3, "PAR1", "f_CAPTION_METASX_DA", "0", "TO METAΣΧ.(DA) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))
 CC = Val(FINDPARAMETROI(3, "PAR1", "f_TIMES_METASX_TIM", "0", "TO METAΣΧ.(TIM) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))
'  f_TIMES_METASX_LIAN = Val(FINDPARAMETROI(3, "PAR1", "f_TIMES_METASX_LIAN", "0", "TO METAΣΧ.(LIAN) ME TIMEΣ=0 ΧΩΡΙΣ ΤΙΜΕΣ=1"))


        '<EhHeader>
        On Error GoTo metasx_TIM_Click_Err

        '</EhHeader>
   
        Dim m As Integer

        Dim c As String
         
        PAR1.TDBGrid1.Col = 2
        c = Replace(PAR1.TDBGrid1.Text, ",", ".")
        'PAR1.Text5.Text = "*" + C
        
        m = Val(FINDPARAMETROI(1, "PAR1", "METASX_TIM", "3", "ΠΑΡΑΣΤΑΤΙΚO ΔΛ.ΔΑ(ΜΕΤΑΣΧ.ΣΕ ΛΙΑΝΙΚΗ)0-99"))
100     PAR1.metasx.Value = vbChecked

         Dim m_ID As Long
         PAR1.TDBGrid1.Col = 7
       m_ID = Val(PAR1.TDBGrid1.Text)
'        PAR1.METASX_SE_TIM m, C, m_ID
        PAR1.METASX_SE_TIM m, c, m_ID, False, CC
        
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

Private Sub mpar1_Click(index As Integer)

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

100   '  If gXEIRISTHS = 9 Then
      ' Else
110         par5.SHOW
       ' End If

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
100     Select Case Listview1.SelectedItem.key

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
    Dim temp As Integer
        temp = gisOnlineMydata
        ' to kano gia na deixnei ta grid
        gisOnlineMydata = 0
        Par7MyData.where.Caption = "" ' ΜΗΝΕΜΕΙΝΕ ΥΠΟΛΕΙΜΑ ΑΠΟ ΤΟ PAR1->ENA ENA POY TA STELNEI
        
40       Par7MyData.Adodc1.ConnectionString = gConnect  'ean einai anoixth den diabazei to connectionstring

        
        
        
        
100     Par7MyData.SHOW

        gisOnlineMydata = temp

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
     
     Dim ISCAR As Integer
     ISCAR = Val(FINDPARAMETROI(1, "PELAT1", "F_ISCAR", "0", "ΣYNERΓEIA=1 πλυντηρια=2 ΟΧΙ=0")) '
     
  If ISCAR > 0 Then
     Dim F As New psifiakopelat
     F.SHOW
   End If
   
    

End Sub

Private Sub mpar9_Click()
    ' par9.SHOW
End Sub

Private Sub mpel3PelatesGenika_Click()
 Dim F As New pelat3
 pelat3.SHOW
 



End Sub

Private Sub mpel3PelatYPOLOIPA_Click()





'Dim f As New pelat3
'pelat3.SHOW

Dim FF As New pelat3param
 'FF.Label_titlos = "Υπόλοιπα Πελατών"

FF.load_sthles_EIDON
'FF.LOAD_KINHSEIS_EIDON
FF.Rep_Fields.Text = "select ROW_NUMBER() OVER(ORDER BY EPO ASC) AS [AA],left(KOD,12) as [Κωδικός],LEFT(EPO,35) as [Περιγραφή],EPA AS [Επάγγελμα],THL AS [Τηλεφωνο],round(TYP,2) AS [Υπόλοιπο],' ' AS [_] fROM PEL " ' ORDER BY EPO ASC "
FF.gemisma_Krithrion
fkrithria.Grid1.TextMatrix(4, 1) = 0.1
fkrithria.Grid1.TextMatrix(4, 2) = 999999999
fkrithria.ΕΧΙΤ_Click
FF.sumes.Text = "0000010"
Dim Q As String: Q = ""
If Val(FINDPARAMETROI(1, "PAR1", "F_MONOENERGOS", "0", "ΔΕΙΧΝΕΙ ΜΟΝΟ ΤΟΥΣ ΕΝΕΡΓΟΥΣ=1 ΟΛΟΥΣ=0")) Then
    Q = " ENERGOS=1 AND "
End If




FF.WHERE_QUERY.Text = "WHERE " + Q + " EIDOS='e' "   'and TYP>=0.1 AND TYP<=999999999" ' StatusBar1.Panels(6).Text"
FF.YPOLOGISMOS_Click  'param



FF.TDBGrid.Splits(0).columns(0).Width = 500
FF.TDBGrid.Splits(0).columns(1).Width = 900
FF.TDBGrid.Splits(0).columns(2).Width = 3900
FF.TDBGrid.Splits(0).columns(3).Width = 3400
FF.TDBGrid.Splits(0).columns(4).Width = 1400
FF.TDBGrid.Splits(0).columns(5).Width = 1400
'FF.TDBGrid.Splits(0).columns(0).Width = 400




FF.SHOW








End Sub

Private Sub mpel3promYPOLOIPA_Click()

Dim FF As New pelat3param
 'FF.Label_titlos = "Υπόλοιπα Πελατών"

FF.load_sthles_EIDON
'FF.LOAD_KINHSEIS_EIDON
FF.Rep_Fields.Text = "select '@@@@@' as [AA],KOD as [Κωδικός],LEFT(EPO,35) as [Περιγραφή],EPA AS [Επάγγελμα],THL AS [Τηλεφωνο],round(TYP,2) AS [Υπόλοιπο],'' as [.] fROM PEL "
FF.gemisma_Krithrion
fkrithria.Grid1.TextMatrix(4, 1) = 0.1
fkrithria.Grid1.TextMatrix(4, 2) = 999999999
fkrithria.ΕΧΙΤ_Click
FF.sumes.Text = "0000010"


Dim Q As String: Q = ""
If Val(FINDPARAMETROI(1, "PAR1", "F_MONOENERGOS", "0", "ΔΕΙΧΝΕΙ ΜΟΝΟ ΤΟΥΣ ΕΝΕΡΓΟΥΣ=1 ΟΛΟΥΣ=0")) Then
    Q = " ENERGOS=1 AND "
End If




FF.WHERE_QUERY.Text = "WHERE " + Q + " EIDOS='r' "   'and TYP>=0.1 AND TYP<=999999999" ' StatusBar1.Panels(6).Text"








'FF.WHERE_QUERY.Text = "WHERE EIDOS='r' and TYP>=0.1 AND TYP<=999999999" ' StatusBar1.Panels(6).Text"
FF.YPOLOGISMOS_Click   'paramypologizo_Ypoloipa



FF.TDBGrid.Splits(0).columns(0).Width = 500
FF.TDBGrid.Splits(0).columns(1).Width = 900
FF.TDBGrid.Splits(0).columns(2).Width = 3900
FF.TDBGrid.Splits(0).columns(3).Width = 3400
FF.TDBGrid.Splits(0).columns(4).Width = 1400
FF.TDBGrid.Splits(0).columns(5).Width = 1400
'FF.TDBGrid.Splits(0).columns(0).Width = 400






FF.SHOW
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

'Private Sub mpelat3_Click()
'
'        '<EhHeader>
'        On Error GoTo mpelat3_Click_Err
'
'        '</EhHeader>
'
'100     pelat3.SHOW
'
'        '<EhFooter>
'        Exit Sub
'
'mpelat3_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.MDIForm1.mpelat3_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.MDIForm1.mpelat3_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub

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

Private Sub mtameiakes_Click()
On Error Resume Next
  TAMEIAKES.SHOW
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

Private Sub par3_Click(index As Integer)

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

Private Sub paralepisk_Click()
    ' paralabh.SHOW
    Main_Form.SHOW 1
    
     
End Sub

Private Sub pareid_Click(index As Integer)

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

Private Sub parenos_Click(index As Integer)

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

Private Sub parkin_Click(index As Integer)

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

Private Sub pontoif_Click(index As Integer)

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

Private Sub pros_Click(index As Integer)

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

Private Sub RUNT_Click(index As Integer)

        '<EhHeader>
        On Error GoTo RUNT_Click_Err

        '</EhHeader>
        Dim k  ' INDEX=0 TO PROTO MENOY   INDEX=1 TO DEYTERO

100     gDir = ETAIR(index)

110    ' MDIForm1.StatusBar1.Panels(1).Text = "Τηλέφωνα" 'gDir

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


If StatusBar1.Panels(1).Text = Panel Then
     R.Open " SELECT TOP 20  THLEFONA.THL,HME,PEL.EPO  FROM THLEFONA INNER JOIN PEL ON PEL.EIDOS=THLEFONA.EIDOS AND PEL.KOD=THLEFONA.KOD ORDER BY HME DESC", Gdb, adOpenForwardOnly, adLockReadOnly

     mc = 0
     runt(0).Caption = R(0)
     ETAIR(0) = R(0)

     Do While Not R.EOF
         mc = mc + 1
         Load runt(mc)
         runt(mc).Caption = R(0) + "*" + Format(R!hme, "HH:MM") + "*" + R!EPO
         ETAIR(mc) = Trim(R(0))
         R.MoveNext
     Loop

     PopupMenu RUNTIME(0)

        Exit Sub





End If








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

Private Sub test_Click()
    frmdummy.SHOW

End Sub



Private Sub THLYPOST_Click()
     YPOSTHRIXI.SHOW
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


If F_TAKENEESTIMES > 0 And f_auto_load_times > 0 Then
    bohu7.get_times_eid
    Kill "c:\LAGEURO\EID.DBF"
    Kill "c:\LAGEURO\EID.DBT"
    Dim jj As Integer
    jj = bohu7.UPD_NEA_BARCODE(0)
    Kill "c:\LAGEURO\BARCODES.DBF"
    
    
End If


   If f_telephone = 30 Then   'EISROON AYTOMATA IMPORT
     bohu5.EISROON
   
   End If




    If f_telephone = 1 Then

        Dim M_EPO

        If IsNull(TELEFON.Recordset("EPO")) Then M_EPO = Left(TELEFON.Recordset("EPO"), 30) Else M_EPO = Space(30)

        TELEFON.Refresh

        'If IsNull(TELEFON.Recordset("EPO")) Or IsNull(TELEFON.Recordset("DIE")) Then
        '  MDIForm1.Caption = TELEFON.Recordset("THL") + " "
        'MDIForm1.Caption = Left(TELEFON.Recordset("THL") + " " + Space(40), 40) + Left(MDIForm1.Caption, 40)
        'Else

        If InStr(Left(MDIForm1.Caption, 18), Left(TELEFON.Recordset("THL"), 10)) > 0 And InStr(Left(MDIForm1.Caption, 18), Format(TELEFON.Recordset("hme"), "hh:mm")) > 0 Then
        Else
            MDIForm1.StatusBar1.Panels(5).Text = Left(MDIForm1.Caption, 42)
            MDIForm1.Caption = Format(TELEFON.Recordset("hme"), "hh:mm") + " " + Left(TELEFON.Recordset("THL"), 16) + "-" + TELEFON.Recordset("EPO") + "  προηγούμενο : " + Left(MDIForm1.Caption, 40) '+ TELEFON.Recordset("DIE")
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
Private Sub timspar_Click(index As Integer)

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

        Dim MKOD As String

        Dim DUM

        Dim DD

        Dim M_BACKUPTYPE
    
        
        

100     M_BACKUPTYPE = FINDPARAMETROI(1, "BOHU8", "F_BACKUPTYPE", 0, "ΤΥΠΟΣ BACKUP 0=ACCESS 1=SQL DATABASE")   'posa psifia tha exei h kathe seira

        Dim M_PATH

110     M_PATH = Trim(FINDPARAMETROI(1, "BOHU8", "F_PATH", "C:\MERCVB", "ΔΙΑΔΡΟΜΗ ΑΝΤΙΓΡΑΦΩΝ"))   'posa psifia tha exei h kathe seira)

120     If Len(THESI) > 0 Then
130         M_PATH = THESI
        End If

         If Len(M_PATH) < 2 Then
             Exit Sub
         End If


        Dim DDTOP
        
             ' EAN EINAI TSENTIDIS
             Dim M_GFK As Integer: M_GFK = Val(FINDPARAMETROI(1, "BOHU8", "F_GFK", 0, "GFK 1000=ΕΒΔ/2000=ΜΗΝ ΑΠΟΣΤΟΛΗ ΑΡΧΕΙΩΝ 0=OXI ΑΛΛΟΙΩΣ ΚΩΔΙΚΟΣ")) ' = 1000
              Dim BACKDAY As Date
              Dim NWEEKUPD2 As Integer
              Dim NMONTH As Integer
              If M_GFK = 1000 Then '1000  posa psifia tha exei h kathe seira
                
                If Weekday(Now) = 2 Then
                   BACKDAY = DateAdd("d", -2, Now)
                Else
                   BACKDAY = DateAdd("d", -1, Now)
                End If
                
                
                NWEEKUPD2 = gVal(GGET_CVALUE("SELECT TIMH FROM PARAMETROI WHERE FORMA='APOT7' AND VAR='WEEKGFK'"))

                If DatePart("ww", Now) = NWEEKUPD2 Then
                   'ok ta esteile aythn thn ebdomada
                Else
                  If NWEEKUPD2 = 0 Then
                     If GGET_NVALUE("select count(*) from PARAMETROI WHERE  FORMA='APOT7' AND VAR='WEEKGFK'") = 0 Then
                       Gdb.Execute "INSERT INTO PARAMETROI (FORMA,VAR,TIMH,SXOLIA) VALUES ('APOT7','WEEKGFK','0','NIELSEN-GFK ΑΡ.ΕΒΔΟΜΑΔΑΣ')"
                     End If
                  End If
                  apot7.gfk_send BACKDAY, BACKDAY, 1
                  Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + LTrim(str(DatePart("ww", Now))) + "' WHERE FORMA='APOT7' AND VAR='WEEKGFK' "
                End If
                
              ElseIf M_GFK = 2000 Then
                  'POIO MHNA ESTEILA TELEYTAIA
                  NMONTH = gVal(GGET_CVALUE("SELECT TIMH FROM PARAMETROI WHERE FORMA='APOT7' AND VAR='WEEKGFK'"))
                  If NMONTH = 0 Then
                     If GGET_NVALUE("select count(*) from PARAMETROI WHERE  FORMA='APOT7' AND VAR='WEEKGFK'") = 0 Then
                       Gdb.Execute "INSERT INTO PARAMETROI (FORMA,VAR,TIMH,SXOLIA) VALUES ('APOT7','WEEKGFK','0','NIELSEN-GFK ΑΡ.ΕΒΔΟΜΑΔΑΣ')"
                     End If
                  End If
                 If Month(Now) = NMONTH Then ' ΟΚ ΕΣΤΕΙΛΑ ΓΙΑ ΤΟΝ ΜΗΝΑ
                   'BACKDAY = DateAdd("d", -2, Now)
                 Else
                    If Month(Now) - NMONTH = 1 Then '
                       Dim NDAY As Integer
                       NDAY = Day(Now)
                       NDAY = -(NDAY)   ' GIA NA PIASO THN TELEYTAIA MERA TOY PROHG MHNA
                       BACKDAY = DateAdd("d", NDAY, Now)
                       
                       
                        Dim FIRSTDAY As Date
                       BACKDAY = DateAdd("d", NDAY, Now)
                       FIRSTDAY = DateSerial(Year(BACKDAY), Month(BACKDAY), 1)
                        
                          apot7.gfk_send FIRSTDAY, BACKDAY, 0
                          Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + str(Month(Now)) + "' WHERE FORMA='APOT7' AND VAR='WEEKGFK' "
                        
                    End If
                  
        
                 End If
           
           
           
           
           
           
           
              End If
        
        
        
        
        
        
        

140     If M_BACKUPTYPE = 1 Then
            ' YPARXEI O FAKELOS BACKUP???
           If Len(Dir(M_PATH + "\", vbNormal)) < 1 Then
              MHNYMA2.Timer1.Interval = 2000
              MHNYMA2.Label2.Caption = "δεν υπάρχει ο φάκελος  " + M_PATH + " για backup"

              MHNYMA2.SHOW 1
              'Me.Caption = "δεν υπάρχει ο φάκελος  " + M_PATH
              Exit Sub
           End If



150         DD = M_PATH + "\BACK" + Format(Day(Now), "00") + ".BAK"

            DoEvents
155         Me.Caption = "παρακαλώ περιμένετε - διαδικασία αντιγράφων " + DD
160         If Len(Dir(DD, vbNormal)) > 0 Then
170             If Format(FileDateTime(DD), "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then

                    Exit Sub

                Else
180                 Kill DD
                    'yparxei_palio = True
                End If
            End If

190           Gdb.Execute "BACKUP DATABASE [" + Gdb.DefaultDatabase + "] TO  DISK ='" + DD + "' WITH NOFORMAT, NOINIT, SKIP, NOREWIND, NOUNLOAD,  STATS = 10 ", k
             CHECKVER
             ' CHECKVER
'             ' EAN EINAI TSENTIDIS
'              If FINDPARAMETROI(1, "BOHU8", "F_GFK", 0, "GFK ΑΠΟΣΤΟΛΗ ΑΡΧΕΙΩΝ 0=OXI ΑΛΛΟΙΩΣ ΚΩΔΙΚΟΣ") = 1000 Then  '1000  posa psifia tha exei h kathe seira
'                Dim BACKDAY As Date
'                If Weekday(Now) = 2 Then
'                   BACKDAY = DateAdd("d", -2, Now)
'                Else
'                   BACKDAY = DateAdd("d", -1, Now)
'                End If
'
'
'                apot7.gfk_send BACKDAY, BACKDAY,1
'              End If
              
             
              'apot7.gfk_send DateAdd("d", -1, Now), DateAdd("d", -1, Now),0
           
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
        
        CHECKVER

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



Sub CHECKVER()

On Error Resume Next

Dim CONN As ADODB.Connection
Dim rs As ADODB.Recordset
Dim fld As ADODB.FIELD
Dim sql As String
'connect to MySQL server using Connector/ODBC
Set CONN = New ADODB.Connection
CONN.ConnectionString = "Provider=SQLNCLI11.1;Persist Security Info=False;pwd=p@ssw0rd;User ID=sa;Initial Catalog=MERCUSERS;Data Source=62.103.69.140\SQLEXPRESS,49737;"
'OK MYSQL LAGAKIS ALLA DEN EXEI ELLINIKA   "DRIVER={MySQL ODBC 3.51 Driver};SERVER=95.216.12.237;UID=a2009_lagakis_op;PWD=Zqpv549#; database=lagakis_op;PTION=3"
'SERVER=localhost; DATABASE=test;UID=venu;PWD=venu; OPTION=3"
CONN.Open
'create table
'conn.Execute "DROP TABLE IF EXISTS my_ado"

Dim MAFM As String
MAFM = GGET_CVALUE("SELECT TOP 1 AFMPEL FROM MEM")

CONN.Execute "insert into LOGINS(LOGIN,HME,C1) VALUES ('" + MAFM + "',GETDATE(),'" + Format(App.Major, "00") + "." + Format(App.Minor, "00") + "." + Format(App.Revision, "00") + "')"


'Dim r As New ADODB.Recordset
'r.Open "SELECT TOP 10 * FROM PEL ", conn, adOpenDynamic, adLockOptimistic
'Do While Not r.EOF
'   List1.AddItem r!EPO + "*" + r!AFM
'   r.MoveNext
'Loop
'
'r.Close
'"Mercury ver." + Format(App.Major, "00") + "." + Format(App.Minor, "00") + "." + Format(App.Revision, "00")
CONN.Close



End Sub

















Private Sub tlbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

        '<EhHeader>
        On Error GoTo tlbToolBar_ButtonClick_Err

        '</EhHeader>
        Dim a

100     a = 1

        Dim DUM

110     Select Case Button.key

            Case "Dynaset"
120            ' HMEROL2.SHOW

130         Case "Table"

                '   hmerol.SHOW
140         Case "BIBLIO"
150           '  paralabh.SHOW

160         Case "Schedule"
170             'DUM = Shell("C:\MERCVB\CALENDAR.EXE", vbMaximizedFocus)
                DUM = Shell("C:\mercvb\CalendarSample\CALENDAR.EXE", vbMaximizedFocus)
        End Select

180     tlbToolBar.Buttons(1).Style = tbrDefault
190     tlbToolBar.Buttons(2).Style = tbrDefault

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

Private Sub GetDriveList()
    Dim idx As Integer, dPath As String
    tvDir.Nodes.Clear
'    roytina(1) = ""
'    roytina(2) = "eis"
'    roytina(3) = "dior"
    'For idx = 0 To 3 ' DriveBox.ListCount - 1
        'dPath = Left(DriveBox.List(idx), 2) & "\"
        tvDir.Nodes.Add , , "Πελάτες", "Πελάτες", 1
        tvDir.Nodes.Add "Πελάτες", tvwChild, , "Εισαγωγή"
        tvDir.Nodes.Add "Πελάτες", tvwChild, , "Διόρθωση"
        tvDir.Nodes.Add "Πελάτες", tvwChild, , "Παρουσίαση Πελατών"
        tvDir.Nodes.Add "Πελάτες", tvwChild, , "Καρτέλα Πελάτου"
        tvDir.Nodes.Add "Πελάτες", tvwChild, , "Κινήσεις Πελατών"
        
        
        
        tvDir.Nodes.Add , , "Προμηθευτές", "Προμηθευτές", 1
        tvDir.Nodes.Add "Προμηθευτές", tvwChild, , "Εισαγωγή."
        tvDir.Nodes.Add "Προμηθευτές", tvwChild, , "Διορθωση"
        tvDir.Nodes.Add "Προμηθευτές", tvwChild, , "Παρουσίαση Προμηθευτών"
        tvDir.Nodes.Add "Προμηθευτές", tvwChild, , "Καρτέλα Προμηθευτή"
        tvDir.Nodes.Add "Προμηθευτές", tvwChild, , "Κινήσεις Προμηθευτών"
        
        tvDir.Nodes.Add , , "Αποθήκη", "Αποθήκη", 1
        tvDir.Nodes.Add "Αποθήκη", tvwChild, , "Εισαγωγή"
        tvDir.Nodes.Add "Αποθήκη", tvwChild, , "Διόρθωση"
        tvDir.Nodes.Add "Αποθήκη", tvwChild, , "Παρουσίαση Ειδών"
        tvDir.Nodes.Add "Αποθήκη", tvwChild, , "Τιμοκατάλογοι"
        tvDir.Nodes.Add "Αποθήκη", tvwChild, , "Κινήσεις Ειδών"
        tvDir.Nodes.Add "Αποθήκη", tvwChild, , "Ρυθμιζόμενες Εκτυπώσεις"
        
        tvDir.Nodes.Add , , "Επιταγές", "Επιταγές", 1
        tvDir.Nodes.Add "Επιταγές", tvwChild, , "Εισαγωγή"
        tvDir.Nodes.Add "Επιταγές", tvwChild, , "Παρουσίαση-Διόρθωση"
        
        
         tvDir.Nodes.Add , , "Παραστατικά", "Παραστατικά", 1
        tvDir.Nodes.Add "Παραστατικά", tvwChild, , "Παραστατικά"
        tvDir.Nodes.Add "Παραστατικά", tvwChild, , "Προβολή κατά ημέρα"
        tvDir.Nodes.Add "Παραστατικά", tvwChild, , "Παρουσίαση ενός Παραστατικού"
        tvDir.Nodes.Add "Παραστατικά", tvwChild, , "Στατιστικά-Παρουσίαση Παραστατικών"
        tvDir.Nodes.Add "Παραστατικά", tvwChild, , "Ταμείο Ημέρας"
        
        
         tvDir.Nodes.Add , , "Βοηθητικά", "Βοηθητικά", 1
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "1.Πίνακες"
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "2.Παράμετροι Προγράμματος"
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "3.Παράμετροι Παραστατικών"
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "4.Παράμετροι Κινήσεων Πελατών"
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "5.Import-Export"
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "6.Αλλαγή Κωδικών Χρήστη"
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "7.Αποστολή σε φορητά τερματικά"
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "8.Μεταφορές Αρχείων-Πόντοι"
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "9.Αντίγραφα Επαναφορά"
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "10.Συγκεντρωτική Τιμολογίων"
        tvDir.Nodes.Add "Βοηθητικά", tvwChild, , "11.Ελεγχοι Πληρότητας Αρχείων"
        
   ' Next idx
End Sub

Private Sub toMydata_Click()

  PAR1.toMydata
  
  


End Sub

Private Sub tvDir_Expand(ByVal Node As ComctlLib.Node)
    On Error GoTo ErrHdler
    Dim idx As Integer, fPos As Integer
    Dim rPath As String, Fname As String, fNewPath As String
    MousePointer = 11
    If Node.child.Text = "" Then
'        tvDir.Nodes.Remove Node.Child.Index
'        rPath = Node.Key
'        DirBox.Path = rPath
'        fPos = Len(rPath) + 1
'        'Add Folders
'        For idx = 0 To DirBox.ListCount - 1
'            fName = Mid(DirBox.List(idx), fPos)
'            fNewPath = rPath & fName & "\"
'            tvDir.Nodes.Add rPath, tvwChild, fNewPath, fName, 1
'            DirBox.Path = fNewPath
'            If (FileBox.ListCount > 0) Or (DirBox.ListCount > 0) Then
'                tvDir.Nodes.Add fNewPath, tvwChild, , ""
'                tvDir.Nodes(fNewPath).ExpandedImage = 2
'            End If
'            DirBox.Path = rPath
'        Next idx
    End If
     tvDir.Nodes(Node.index + 1).Image = 1
      tvDir.Nodes(Node.index + 2).Image = 1
       tvDir.Nodes(Node.index + 3).Image = 1
        tvDir.Nodes(Node.index + 4).Image = 1
        
        tvDir.Nodes(Node.index + 5).Image = 1
      tvDir.Nodes(Node.index + 6).Image = 1
       tvDir.Nodes(Node.index + 7).Image = 1
        tvDir.Nodes(Node.index + 8).Image = 1
        
        
        
    GoTo ExitSub
ErrHdler:
    'If Drive not ready handle error and re-instate removed item
    tvDir.Nodes.Add Node.key, tvwChild, , ""
    Resume ExitSub
ExitSub:
    MousePointer = 0
End Sub
Private Sub tvDir_NodeClick(ByVal Node As Node)
    Static tempNK As Integer
    Me.Caption = Node.key
 '   Label2.Caption = Node.Key 'πελατες,προμηθευτες κ.λ.π.
   ' Label1.Caption = Node.Index  '1,2,3,4,5,6 απο πάνω προς τα κάτω
    tvDir.Nodes(Node.index).Image = 2
    tvDir.Nodes(Node.index).ExpandedImage = 2
    If tempNK <> Empty And tempNK <> Node.index Then
        tvDir.Nodes(tempNK).ExpandedImage = 1
        tvDir.Nodes(tempNK).Image = 1
    End If
    tempNK = Node.index
    'Call AddFiles  ' edv kalv thn roytina
    ' CallByName Me, roytina(tempNK), VbMethod
   '============================================================================================================
    If tvDir.Nodes(tempNK).Parent = "Πελάτες" And Left(tvDir.Nodes(tempNK).Text, 8) = "Εισαγωγή" Then 'tempNK = 2 Then
       pelat1.SHOW
    End If
     If tvDir.Nodes(tempNK).Parent = "Πελάτες" And tvDir.Nodes(tempNK).Text = "Διόρθωση" Then  ' tempNK = 3 Then
        gApoMenu = True
        pelat2.SHOW
    End If
    If tvDir.Nodes(tempNK).Parent = "Πελάτες" And tvDir.Nodes(tempNK).Text = "Παρουσίαση Πελατών" Then ' tempNK = 4 Then
        'gApoMenu = True
        pelat3.SHOW
    End If
    
    If tvDir.Nodes(tempNK).Parent = "Πελάτες" And Left(tvDir.Nodes(tempNK).Text, 7) = "Καρτέλα" Then ' tempNK = 5 Then
        'gApoMenu = True
        pelat4.SHOW
    End If
    
     If tvDir.Nodes(tempNK).Parent = "Πελάτες" And tvDir.Nodes(tempNK).Text = "Κινήσεις Πελατών" Then ' tempNK = 5 Then
        'gApoMenu = True
        pelat5.SHOW
    End If
    
    
    
    '=======================================================================================================================
    If tvDir.Nodes(tempNK).Parent = "Προμηθευτές" And Left(tvDir.Nodes(tempNK).Text, 8) = "Εισαγωγή" Then 'tempNK = 2 Then
       pelat1.SHOW
    End If
     If tvDir.Nodes(tempNK).Parent = "Προμηθευτές" And tvDir.Nodes(tempNK).Text = "Διόρθωση" Then  ' tempNK = 3 Then
        gApoMenu = True
        pelat2.SHOW
    End If
    If tvDir.Nodes(tempNK).Parent = "Προμηθευτές" And tvDir.Nodes(tempNK).Text = "Παρουσίαση Προμηθευτών" Then ' tempNK = 4 Then
        'gApoMenu = True
        pelat3.SHOW
    End If
    
    If tvDir.Nodes(tempNK).Parent = "Προμηθευτές" And Left(tvDir.Nodes(tempNK).Text, 7) = "Καρτέλα" Then ' tempNK = 5 Then
        'gApoMenu = True
        pelat4.SHOW
    End If
    
     If tvDir.Nodes(tempNK).Parent = "Προμηθευτές" And tvDir.Nodes(tempNK).Text = "Κινήσεις Προμηθευτών" Then ' tempNK = 5 Then
        'gApoMenu = True
        pelat5.SHOW
    End If
    
    
 '=======================================================================================================================
    If tvDir.Nodes(tempNK).Parent = "Αποθήκη" And Left(tvDir.Nodes(tempNK).Text, 8) = "Εισαγωγή" Then 'tempNK = 2 Then
       APOT1.SHOW
    End If
     If tvDir.Nodes(tempNK).Parent = "Αποθήκη" And tvDir.Nodes(tempNK).Text = "Διόρθωση" Then  ' tempNK = 3 Then
        gApoMenu = True
        APOT2.SHOW
    End If
    If tvDir.Nodes(tempNK).Parent = "Αποθήκη" And Left(tvDir.Nodes(tempNK).Text, 8) = "Παρουσία" Then ' tempNK = 4 Then
        'gApoMenu = True
        apot3.SHOW
    End If
    
    If tvDir.Nodes(tempNK).Parent = "Αποθήκη" And Left(tvDir.Nodes(tempNK).Text, 7) = "Τιμοκατ" Then ' tempNK = 5 Then
        'gApoMenu = True
        apot4.SHOW
    End If
    
     If tvDir.Nodes(tempNK).Parent = "Αποθήκη" And Left(tvDir.Nodes(tempNK).Text, 3) = "Κιν" Then ' tempNK = 5 Then
        'gApoMenu = True
        APOT5.SHOW
    End If
    
    
    
    
    
    
    
    
    
    
    
        
    
    
    
    
End Sub


Private Sub Updater_Click()
If Len(Dir("C:\mercvb\updatemerc.bat", vbNormal)) < 10 Then

 Shell "C:\mercvb\updatemerc.bat", vbHide
 End If
 
 '        End

'   If Len(Dir("C:\mercvb\updater\WindowsApp1.exe", vbNormal)) > 10 Then
'      Dim ANS As Integer
'      ANS = MsgBox("Να κλείσει το πρόγραμμα για να προχωρήσει η ανβάθμιση;", vbYesNo)
'      If ANS = vbYes Then
'         Shell "C:\mercvb\updater\WindowsApp1.exe", vbNormalFocus
'         End
'      End If
'   End If


End Sub

