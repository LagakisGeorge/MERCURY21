VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form par2frmOrder 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   Caption         =   "Παρ/κά ανά ημέρα"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   969
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11520
      TabIndex        =   33
      Top             =   6720
      Width           =   2085
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8385
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "20/02/2022"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3375
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   14
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrder.frx":0000
            Key             =   "SerieShut"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrder.frx":02F4
            Key             =   "SerieOpen"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrder.frx":05B8
            Key             =   "Teller"
            Object.Tag             =   "3"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabBestell 
      Height          =   7350
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   14025
      _ExtentX        =   24739
      _ExtentY        =   12965
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   794
      BackColor       =   16711680
      MouseIcon       =   "frmOrder.frx":08AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Παραστατικά"
      TabPicture(0)   =   "frmOrder.frx":08C8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TvwCustomer"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LvOrders"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "picSerie(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DTPicker1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DTPicker2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text4"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   11040
         TabIndex        =   32
         Top             =   6000
         Width           =   2085
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Υπολογισμός"
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   6000
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H000080FF&
         Height          =   372
         Left            =   8640
         TabIndex        =   28
         Top             =   6480
         Width           =   1332
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H000080FF&
         Height          =   372
         Left            =   7200
         TabIndex        =   27
         Top             =   6480
         Width           =   852
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Εξοδος"
         Height          =   330
         Left            =   11040
         TabIndex        =   26
         Top             =   6960
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   4710
         TabIndex        =   23
         Top             =   6435
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   153944065
         CurrentDate     =   39110
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1665
         TabIndex        =   22
         Top             =   6435
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   153944065
         CurrentDate     =   39110
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   11040
         TabIndex        =   19
         Top             =   5640
         Width           =   2085
      End
      Begin VB.PictureBox picSerie 
         AutoRedraw      =   -1  'True
         Height          =   1800
         Index           =   0
         Left            =   3720
         ScaleHeight     =   116
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   621
         TabIndex        =   2
         Top             =   720
         Width           =   9372
         Begin VB.TextBox txtDispatch 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1395
            TabIndex        =   21
            Tag             =   "1"
            Top             =   855
            Width           =   1005
         End
         Begin VB.TextBox txtFreight 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1395
            TabIndex        =   18
            Tag             =   "1"
            Top             =   1215
            Width           =   1005
         End
         Begin VB.TextBox txtShipCity 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3780
            TabIndex        =   17
            Tag             =   "1"
            Top             =   1260
            Width           =   1140
         End
         Begin VB.TextBox txtPostelCode 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3780
            TabIndex        =   16
            Tag             =   "1"
            Top             =   900
            Width           =   1140
         End
         Begin VB.TextBox txtAddress 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3780
            TabIndex        =   15
            Tag             =   "1"
            Top             =   495
            Width           =   3390
         End
         Begin VB.TextBox txtCustomer 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3780
            TabIndex        =   5
            Tag             =   "1"
            Top             =   135
            Width           =   3390
         End
         Begin VB.TextBox txtOrderDate 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1395
            TabIndex        =   4
            Tag             =   "1"
            Top             =   495
            Width           =   1005
         End
         Begin VB.TextBox txtOrderID 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1395
            TabIndex        =   3
            Tag             =   "1"
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label Label4 
            Caption         =   "Πόλη"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2565
            TabIndex        =   20
            Top             =   1260
            Width           =   1140
         End
         Begin VB.Label Label3 
            Caption         =   "Τ.Κ."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2520
            TabIndex        =   12
            Top             =   900
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Διεύθυνση"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2520
            TabIndex        =   11
            Top             =   540
            Width           =   1350
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label lblVersandDatum 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   60
            TabIndex        =   9
            Top             =   900
            Width           =   1275
         End
         Begin VB.Label lblBestellDatum 
            Caption         =   "Ημερ/νία"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   540
            Width           =   975
         End
         Begin VB.Label lblEmpfδnger 
            Caption         =   "Πελάτης"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2520
            TabIndex        =   7
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblBestellNr 
            BackStyle       =   0  'Transparent
            Caption         =   "Αρ.Παραστ #"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   6
            Top             =   180
            Width           =   1080
         End
      End
      Begin MSComctlLib.ListView LvOrders 
         Height          =   2436
         Left            =   3720
         TabIndex        =   13
         Top             =   2928
         Width           =   9372
         _ExtentX        =   16536
         _ExtentY        =   4286
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.TreeView TvwCustomer 
         Height          =   5292
         Left            =   132
         TabIndex        =   14
         Top             =   720
         Width           =   3552
         _ExtentX        =   6271
         _ExtentY        =   9340
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Κωδικός"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   8640
         TabIndex        =   30
         Top             =   6120
         Width           =   852
      End
      Begin VB.Label Label7 
         Caption         =   "Ποσότητα"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   7200
         TabIndex        =   29
         Top             =   6120
         Width           =   852
      End
      Begin VB.Label Label6 
         Caption         =   "Εως"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3975
         TabIndex        =   25
         Top             =   6450
         Width           =   600
      End
      Begin VB.Label Label5 
         Caption         =   "Από "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   780
         TabIndex        =   24
         Top             =   6480
         Width           =   750
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   3840
         X2              =   13080
         Y1              =   5496
         Y2              =   5496
      End
   End
End
Attribute VB_Name = "par2frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim REGGTIM        As New ADODB.Recordset

Dim Rtim           As Recordset

Dim REGGTIMDetails As Recordset

Dim m_par(255)     As String

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Dim lOrderKey As Long         ' we will need this when the User hit's the Treeview

Dim sOrderNr  As String    'pass this to the Listview

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     updateTree

        '   Call initializeForm
        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub oldForm_Activate()

        'par2frmOrder.height = 7356
        'par2frmOrder.width = 11520
        '<EhHeader>
        On Error GoTo Form_Activate_Err

        '</EhHeader>

100     dtpicker2.Value = DateAdd("D", 1, Now)
110     dtpicker1.Value = DateAdd("D", -Day(Now) + 1, Now)

        Dim dB, R As New ADODB.Recordset

120     R.Open "select TITLOS,EIDOS from PARASTAT;", Gdb, adOpenDynamic, adLockOptimistic

130     R.MoveFirst

140     Do While Not R.EOF

150         If Not IsNull((R("EIDOS"))) Then
160             If IsNull((R("titlos"))) Then
170                 m_par(Asc(R("EIDOS"))) = ""
                Else
180                 m_par(Asc(R("EIDOS"))) = R("TITLOS")
                End If
            End If

190         R.MoveNext
        Loop

200     R.Close
210     Call initializeForm

        '<EhFooter>
        Exit Sub

Form_Activate_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.Form_Activate " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.Form_Activate " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl

        'remember the Function to open the Database !!
        'if the Database isn't there you'll get a Message
        '<EhFooter>
        
        'par2frmOrder.height = 7356
        'par2frmOrder.width = 11520
102     dtpicker2.Value = DateAdd("D", 1, Now)
110     dtpicker1.Value = DateAdd("D", -Day(Now) + 1, Now)

        Dim dB, R As New ADODB.Recordset

120     R.Open "select TITLOS,EIDOS from PARASTAT;", Gdb, adOpenDynamic, adLockOptimistic

130     R.MoveFirst

140     Do While Not R.EOF

150         If Not IsNull((R("EIDOS"))) Then
160             If IsNull((R("titlos"))) Then
170                 m_par(Asc(R("EIDOS"))) = ""
                Else
180                 m_par(Asc(R("EIDOS"))) = R("TITLOS")
                End If
            End If

190         R.MoveNext
        Loop

200     R.Close
210     Call initializeForm

        '<EhFooter>
        
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub updateTree()

        'Dim RTIM  As Recordset 'Customers Table
        '<EhHeader>
        On Error GoTo updateTree_Err

        '</EhHeader>
        Dim sqlCustomer As String    'make a Sqlstatement for the Treeview

        'Dim REGGTIM     As Recordset ' Order Table
        Dim sqlOrder    As String    ' place the Orders to the Customer

        Dim Rtim        As New ADODB.Recordset

        '1.) Clear the Treeview first
100     TvwCustomer.Nodes.Clear

110     Rtim.Open "select distinct convert(CHAR(10),HME,3) FROM TIM WHERE HME>='" + Format(dtpicker1, "MM/DD/YYYY") + "' AND HME <='" + Format(dtpicker2, "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

        Dim k

        '-------------------------------------------------------------------------------
120     Do While Not Rtim.EOF
130         TvwCustomer.Nodes.ADD , , "L" + Format(Rtim(0), "DD/MM/YYYY"), Format(Rtim(0), "DD/MM/YYYY"), 1, 1
140         Rtim.MoveNext
        Loop

150     Rtim.Close

160     Rtim.Open "select *,PEL.EPO FROM TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD WHERE HME>='" + Format(dtpicker1, "MM/DD/YYYY") + "' AND HME <='" + Format(dtpicker2, "MM/DD/YYYY") + "' ORDER BY ATIM", Gdb, adOpenForwardOnly, adLockReadOnly

        '---------------------------------------------------------------------------------
170     Do While Not Rtim.EOF
            '   TvwCustomer.Nodes.Add "L" + _
            '   Format(rtim("hme"), "DD/MM/YYYY"), tvwChild, rtim("ATIM") + _
            '   Format(rtim("ID_NUM"), "000000"), rtim("ATIM") + " " + m_par(Asc(rtim("atim"))) _
            '   , 1, 2

            On Error Resume Next

180         TvwCustomer.Nodes.ADD "L" + Format(Rtim("hme"), "DD/MM/YYYY"), tvwChild, Rtim("ATIM") + Format(Rtim("ID_NUM"), "000000"), Rtim("ATIM") + " " + Rtim("EPO") + " " + m_par(Asc(Rtim("atim"))) + ";" + str(Rtim("ID_NUM")), 1, 2

190         Rtim.MoveNext
        Loop

        Exit Sub

        '
200     REGGTIM.Open "select * FROM EGGTIM WHERE DAY(HME)=4 ", Gdb, adOpenForwardOnly, adLockReadOnly

210     Do While Not REGGTIM.EOF
220         TvwCustomer.Nodes.ADD REGGTIM("ATIM") + Format(hme, "DD/MM/YYYY") + REGGTIM("PELKOD"), tvwChild, "ID" & CStr(REGGTIM("hme")), "Invoice #" & REGGTIM("atim"), 1, 2
230         REGGTIM.MoveNext

240         If (REGGTIM.EOF) Then

                Exit Do

            End If

            'REGGTIM.Close

        Loop

250     DoEvents
        'sbStatus.Panels.Item(1).Text = "There are currently " & _
        'rtim.RecordCount & " Customers in the Database"

260     Rtim.Close
270     REGGTIM.Close

        '<EhFooter>
        Exit Sub

updateTree_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.updateTree " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.updateTree " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub setUpListView()

        '<EhHeader>
        On Error GoTo setUpListView_Err

        '</EhHeader>
        Dim clmHdr As ColumnHeader

100     LvOrders.ColumnHeaders.ADD 1, , "Περιγραφή", 4000, lvwColumnLeft

110     LvOrders.ColumnHeaders.ADD 2, , "Ποσότητα", 1000, lvwColumnRight
120     LvOrders.ColumnHeaders.ADD 3, , "Τιμή", 1000, lvwColumnRight
130     LvOrders.ColumnHeaders.ADD 4, , "Εκπ%", 1000, lvwColumnRight
140     LvOrders.ColumnHeaders.ADD 5, , "Αξία", 1000, lvwColumnRight
150     LvOrders.ColumnHeaders.ADD 6, , "Φορτ", 1000, lvwColumnRight
160     LvOrders.ColumnHeaders.ADD 7, , "Υπολ", 1000, lvwColumnRight
170     LvOrders.ColumnHeaders.ADD 8, , "Κωδικός", 1000, lvwColumnRight
        'Set clmHdr = LvOrders.ColumnHeaders. _
        '             Add(, , "Quantity", 860, lvwColumnCenter)
        'Set clmHdr = LvOrders.ColumnHeaders. _
        '             Add(, , "Discount", 800, lvwColumnRight)
        'Set clmHdr = LvOrders.ColumnHeaders. _
        '             Add(, , "Amount", 1000, lvwColumnRight)
        'Set clmHdr = LvOrders.ColumnHeaders. _
        '             Add(, , "", 0)
180     LvOrders.View = lvwReport

        '<EhFooter>
        Exit Sub

setUpListView_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.setUpListView " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.setUpListView " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub initializeForm()

        '<EhHeader>
        On Error GoTo initializeForm_Err

        '</EhHeader>

100     Screen.MousePointer = vbHourglass
110     sbStatus.Panels.Item(2).Text = "Loading..."
120     tabBestell.Tab = 0

130     DoEvents
140     Call clearFields
150     Call lockFields(True)
160     Call updateTree
170     Call setUpListView
180     Screen.MousePointer = vbDefault
190     sbStatus.Panels.Item(2).Text = "ΟΚ.."

        '<EhFooter>
        Exit Sub

initializeForm_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.initializeForm " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.initializeForm " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub clearFields()

        '<EhHeader>
        On Error GoTo clearFields_Err

        '</EhHeader>
        Dim indx     As Integer

        Dim tempMask As String

100     With Me.Controls

110         For indx = 0 To .Count - 1

120             If Me.Controls(indx).Tag = "1" Then
130                 If (TypeOf Me.Controls(indx) Is TextBox) Then
140                     Me.Controls(indx).Text = ""
150                 ElseIf (TypeOf Me.Controls(indx) Is ComboBox) Then
                    End If
                End If

            Next

        End With

160     DoEvents

        '<EhFooter>
        Exit Sub

clearFields_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.clearFields " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.clearFields " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub populateFields()

        '<EhHeader>
        On Error GoTo populateFields_Err

        '</EhHeader>

100     Call clearFields

        Dim R As New ADODB.Recordset

        Dim meidos

110     meidos = REGGTIM("eidos")

120     R.Open "select * from PEL WHERE EIDOS='" + meidos + "' AND KOD='" + REGGTIM("PELKOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

130     With REGGTIM

140         If (Not IsNull(!ATIM)) Then txtOrderID = !ATIM
150         If (Not IsNull(!hme)) Then txtOrderDate = !hme
160         If (Not IsNull(R!EPO)) Then txtCustomer = R!EPO

            'If (Not IsNull(r!EPO)) Then txtFreight = !Freight
            'If (Not IsNull(!ShippedDate)) Then txtDispatch = !ShippedDate
170         If (Not IsNull(R!DIE)) Then txtAddress = R!DIE
180         If (Not IsNull(R!XRVMA)) Then txtPostelCode = R!XRVMA
190         If (Not IsNull(R!pol)) Then txtShipCity = R!pol

200         DoEvents
        End With

        '<EhFooter>
        Exit Sub

populateFields_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.populateFields " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.populateFields " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub lockFields(bDoLock As Boolean)

        '<EhHeader>
        On Error GoTo lockFields_Err

        '</EhHeader>
        Dim indx As Integer

100     For indx = 0 To Me.Controls.Count - 1

110         If Me.Controls(indx).Tag = "1" Then
120             If (TypeOf Me.Controls(indx) Is TextBox) Then
130                 If (bDoLock = True) Then
140                     Me.Controls(indx).Locked = True
150                     Me.Controls(indx).BackColor = &H808000
160                     Me.Controls(indx).ForeColor = vbWhite
                    Else
170                     Me.Controls(indx).Locked = False
180                     Me.Controls(indx).BackColor = vbWhite
190                     Me.Controls(indx).BackColor = vbBlack
                    End If

200             ElseIf (TypeOf Me.Controls(indx) Is ComboBox) Then

210                 If (bDoLock = True) Then
220                     Me.Controls(indx).Enabled = False
230                     Me.Controls(indx).BackColor = vbWhite
240                     Me.Controls(indx).ForeColor = vbBlack
                    Else
250                     Me.Controls(indx).Enabled = True
260                     Me.Controls(indx).BackColor = vbWhite
270                     Me.Controls(indx).BackColor = vbBlack
                    End If
                End If
            End If

        Next

280     DoEvents

        '<EhFooter>
        Exit Sub

lockFields_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.lockFields " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.lockFields " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 14760, 10085, fh, fw, ft, fl  'mForm_Resize Me,8,8, 14760, 9285, fh, fw, ft, fl
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Text3_KeyUp_Err

        '</EhHeader>
100     If KeyCode = 13 Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Text3_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.Text3_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.Text3_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text3_LostFocus()

        '<EhHeader>
        On Error GoTo Text3_LostFocus_Err

        '</EhHeader>

        Dim itemToAdd  As ListItem

        Dim BestellSQL As String

        '    LvOrders.ListItems.Clear
        Dim N
    
        Dim R As New ADODB.Recordset

100     If Len(Trim(Text3.Text)) = 0 Then

            Exit Sub

        End If
    
        Dim MKOD As String

110     MKOD = LTrim(Trim(Text3.Text))

120     If Left(MKOD, 1) = "/" Then
130         MKOD = mID(MKOD, 2, Len(MKOD) - 1)
        Else
140         R.Open "SELECT KOD FROM BARCODES WHERE ERG='" + MKOD + "'", Gdb, adOpenDynamic, adLockOptimistic

150         If R.EOF Then
160             MKOD = "ΩΩΩΩΩΩΩΩΩΩΩΩΩΩ" 'DEN YPARX
            Else
170             MKOD = R(0)
            End If

180         R.Close
        End If

        ' If Not REGGTIM.EOF Then
        '    REGGTIM.MoveFirst
        ' While Not REGGTIM.EOF
        'Dim ddd
        ' N = N + 1
        Dim FOUND As Integer

190     FOUND = 0

200     For N = 1 To LvOrders.ListItems.Count

210         If LvOrders.ListItems.Item(N).SubItems(7) = MKOD Then
220             FOUND = 1

230             If Val(Text2.Text) > Val(LvOrders.ListItems.Item(N).SubItems(1)) - Val(LvOrders.ListItems.Item(N).SubItems(5)) Then
240                 MsgBox "ΥΠΕΡΒΑΣΗ ΠΑΡΑΓΓΕΛΙΑΣ"
250                 Text3.SetFocus
                Else
260                 LvOrders.ListItems.Item(N).SubItems(5) = Val(LvOrders.ListItems.Item(N).SubItems(5)) + Val(Text2.Text)
270                 LvOrders.ListItems.Item(N).SubItems(6) = Val(LvOrders.ListItems.Item(N).SubItems(1)) - Val(LvOrders.ListItems.Item(N).SubItems(5))
                End If
            End If

        Next
       
280     If FOUND = 0 Then
290         MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ Ο ΚΩΔΙΚΟΣ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ"
300         Text3.SetFocus
        End If
       
        'Else
        '   Set itemToAdd = LvOrders.ListItems.Add(, , "None ")
        '  sbStatus.Panels.Item(1).Text = "0 Products " _
           & sOrderNr
        ' End If
310     LvOrders.SelectedItem = LvOrders.ListItems(1)
320     Call LvOrders_ItemClick(LvOrders.SelectedItem)

330     DoEvents
340     Text3.Text = "  "
350     Text3.SetFocus

        '<EhFooter>
        Exit Sub

Text3_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.Text3_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.Text3_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TvwCustomer_NodeClick(ByVal Node As MSComctlLib.Node)

        '<EhHeader>
        On Error GoTo TvwCustomer_NodeClick_Err

        '</EhHeader>
100     If (Len(Node.Key) = 6) Then Exit Sub
        'lOrderKey = CLng(Mid$(Node.Key, 3, Len(Node.Key)))
        'With REGGTIM
        '   .Index = "PrimaryKey"
        '   .Seek "=", lOrderKey

        On Error Resume Next

110     REGGTIM.Close

        Dim RTIM2 As New ADODB.Recordset

        RTIM2.Open "SELECT * FROM TIM WHERE ID_NUM=" + Split(Node, ";")(1), Gdb, adOpenForwardOnly, adLockReadOnly

120     REGGTIM.Open "SELECT * FROM EGGTIM WHERE ID_NUM=" + Split(Node, ";")(1), Gdb, adOpenForwardOnly, adLockReadOnly

130     If Not REGGTIM.EOF Then
140         bFieldsPopulated = True
150         sOrderNr = TvwCustomer.SelectedItem
160         Call populateFields
170         Call populateListView
180         Call ShowTotal
            Text4.Text = "Αξία ΦΠΑ      " + Format(RTIM2("FPA6") + RTIM2("FPA7") + RTIM2("FPA1") + RTIM2("FPA2") + RTIM2("FPA3") + RTIM2("FPA4"), "###0.00")
            Text5.Text = "Συν.Αξία      " + Format(RTIM2("aji"), "###0.00")
            RTIM2.Close

        Else
            'MsgBox ("Ohhhh Nooo")
        End If

        'End With
        '<EhFooter>
        Exit Sub

TvwCustomer_NodeClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.TvwCustomer_NodeClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.TvwCustomer_NodeClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub populateListView()

        '<EhHeader>
        On Error GoTo populateListView_Err

        '</EhHeader>
        Dim itemToAdd  As ListItem

        Dim BestellSQL As String

100     LvOrders.ListItems.Clear
    
        Dim N

        'Set REGGTIMDetails = dbNorthwind.OpenRecordset(BestellSQL)
110     If Not REGGTIM.EOF Then
120         REGGTIM.MoveFirst

130         Do While Not REGGTIM.EOF

                'Dim ddd
140             N = N + 1
150             LvOrders.ListItems.ADD , , REGGTIM!ONOMA
160             LvOrders.ListItems.Item(N).SubItems(1) = Format$(REGGTIM!poso, "#,##0.00;(#,##0.00)")
170             LvOrders.ListItems.Item(N).SubItems(2) = Format$(REGGTIM!TIMM, "#,##0.00;(#,##0.00)")
            
180             LvOrders.ListItems.Item(N).SubItems(3) = Format$(REGGTIM!EKPT, "#,##0.00;(#,##0.00)")
190             LvOrders.ListItems.Item(N).SubItems(4) = Format$(REGGTIM!poso * REGGTIM!TIMM * (100 - REGGTIM!EKPT) / 100, "#,##0.00;(#,##0.00)")
200             LvOrders.ListItems.Item(N).SubItems(7) = REGGTIM!KODE
            
210             REGGTIM.MoveNext
                If N > 200 Then ' το εβαλα γιατι κολουσε οταν πηγαινες πανω κατω με βελακια
                   Exit Do
                End If
            Loop

            '   sbStatus.Panels.Item(1).Text = " " & _
                REGGTIMDetails.RecordCount & " Position(s) in " & _
                sOrderNr
        Else
220         Set itemToAdd = LvOrders.ListItems.ADD(, , "None ")
230         sbStatus.Panels.Item(1).Text = "0 Products " & sOrderNr
        End If

240     LvOrders.SelectedItem = LvOrders.ListItems(1)
250     Call LvOrders_ItemClick(LvOrders.SelectedItem)

260     DoEvents

        '<EhFooter>
        Exit Sub

populateListView_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.populateListView " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.populateListView " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub LvOrders_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'If (Not REGGTIM.EOF) Then
    'REGGTIM.MoveFirst
    '    REGGTIM.FindFirst "OrderID = " & _
    '            LvOrders.ListItems(Item.Index).SubItems(5)
    '        End If
End Sub

Private Sub ShowTotal()

        '<EhHeader>
        On Error GoTo ShowTotal_Err

        '</EhHeader>
        Dim i      As Integer

        Dim cTotal As Currency

100     With LvOrders

110         For i = 1 To .ListItems.Count
120             cTotal = cTotal + CCur(.ListItems(i).SubItems(4))
            Next

        End With

130     Text1.Text = " Kαθ.Αξίας: " & Format$(cTotal, "#,##0.00;(#,##0.00)")

        '<EhFooter>
        Exit Sub

ShowTotal_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2frmOrder.ShowTotal " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2frmOrder.ShowTotal " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

