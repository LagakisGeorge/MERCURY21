VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Par7MyData 
   BackColor       =   &H00C0E0FF&
   Caption         =   "MyData"
   ClientHeight    =   12525
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21405
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12525
   ScaleWidth      =   21405
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton PL8_4 
      Caption         =   "PL8_4"
      Height          =   360
      Left            =   15960
      TabIndex        =   72
      Top             =   10440
      Width           =   990
   End
   Begin VB.CommandButton cmdPROPLHROMH 
      Caption         =   "PROPLHROMH"
      Height          =   360
      Left            =   10920
      TabIndex        =   71
      Top             =   11400
      Width           =   2055
   End
   Begin VB.CommandButton cmdKOYKAKH 
      Caption         =   "koykakh"
      Height          =   492
      Left            =   16800
      TabIndex        =   70
      Top             =   4800
      Width           =   1332
   End
   Begin VB.CommandButton CMDENARXIAPOSTOLIS 
      Caption         =   "ΕΝΑΡΞΗ αποστολης"
      Height          =   360
      Left            =   240
      TabIndex        =   69
      Top             =   7800
      Width           =   1812
   End
   Begin VB.CommandButton CMDLIXIAPOSTOLIS 
      Caption         =   "ληψη δελτιου αποστολης"
      Height          =   360
      Left            =   2280
      TabIndex        =   68
      Top             =   7800
      Width           =   1812
   End
   Begin VB.CommandButton Command14 
      Caption         =   "ελεγχος ΔΑ με μαρκ GetDeliveryNoteStatus "
      Height          =   360
      Left            =   360
      TabIndex        =   67
      Top             =   8280
      Width           =   2412
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   19920
      Top             =   10560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   360
      Left            =   15360
      TabIndex        =   66
      Top             =   4680
      Width           =   990
   End
   Begin VB.CommandButton cmdTEST2 
      Caption         =   "TEST2"
      Height          =   360
      Left            =   17400
      TabIndex        =   65
      Top             =   10680
      Width           =   990
   End
   Begin VB.CommandButton CMDDEHTEST 
      Caption         =   "DEHTEST"
      Height          =   360
      Left            =   18240
      TabIndex        =   64
      Top             =   9240
      Width           =   990
   End
   Begin VB.CommandButton eisagTIM 
      Caption         =   "ΕΙΣΑΓΩΓΗ ΠΑΡ/ΚΟΥ ΣΕ ΒΑΣΗ"
      Height          =   372
      Left            =   4200
      TabIndex        =   63
      Top             =   4680
      Width           =   3012
   End
   Begin VB.ComboBox MYDPAR 
      Height          =   315
      ItemData        =   "par7MyData.frx":0000
      Left            =   240
      List            =   "par7MyData.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cancelClient 
      BackColor       =   &H000080FF&
      Caption         =   "cancelclient"
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   12120
      Width           =   2775
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "request clients"
      Height          =   375
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   10440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdRELATECLIENT 
      BackColor       =   &H0080FF80&
      Caption         =   "RELATECLIENT"
      Height          =   360
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   9960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   360
      Left            =   14160
      TabIndex        =   58
      Top             =   10440
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton updateclient 
      BackColor       =   &H0080FF80&
      Caption         =   "updateclient"
      Height          =   480
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   10320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdEXPORTTOPDF 
      Caption         =   "EXPORTTOPDF"
      Height          =   360
      Left            =   18480
      TabIndex        =   56
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton jsontimol 
      Caption         =   "JSONTIMOL"
      Height          =   375
      Left            =   6360
      TabIndex        =   55
      Top             =   11160
      Width           =   2415
   End
   Begin VB.CommandButton cmdParochos2 
      Caption         =   "parochos2"
      Height          =   360
      Left            =   9120
      TabIndex        =   54
      Top             =   11760
      Width           =   990
   End
   Begin VB.CommandButton cmdPAROCHOS 
      Caption         =   "PAROCHOS test"
      Height          =   360
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CheckBox chkexodon 
      Alignment       =   1  'Right Justify
      Caption         =   "Εξόδων"
      Height          =   255
      Left            =   2880
      TabIndex        =   52
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton requestafm 
      Caption         =   "requestafm"
      Height          =   360
      Left            =   4200
      TabIndex        =   51
      Top             =   9360
      Width           =   1695
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   21405
      _ExtentX        =   37756
      _ExtentY        =   1799
      ButtonWidth     =   3016
      ButtonHeight    =   1640
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Συσχάτιση Πιστ."
            Key             =   "SYSX"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Επαν/στολή Ακυρ"
            Key             =   "EPAN"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "για Επαναποστολή"
            Key             =   "GIAEPAN"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "diagdok"
                  Text            =   "Διαγραφή Δοκιμαστικών"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "epanfile"
                  Text            =   "Επαναποστολή Από αρχείο INV.XML"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "cmdDeleteMark"
                  Text            =   "Διαγραφή Μαρκ(μόνο για τεχνικό)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "XML τρέχ.εγγραφής"
            Key             =   "XML"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ενα-ένα Αποστολή"
            Key             =   "ONEBYONE"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Κινήσεις ΑΦΜ"
            Key             =   "KINHSEIS"
            ImageIndex      =   13
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kinpel"
                  Text            =   "Πελάτης"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kinprom"
                  Text            =   "Προμηθευτής"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Αγορές"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EmporAney"
                  Text            =   "Εμπορικό χωρίς αντιστοίχιση"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MydataAney"
                  Text            =   "Mydata χωρις αντιστοίχιση"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command12 
      Caption         =   "MAZIKES AKYR ΑΠΟ APESTALMENA2"
      Height          =   375
      Left            =   4200
      TabIndex        =   44
      Top             =   10320
      Width           =   1695
   End
   Begin VB.CommandButton SENDONEBYONE 
      Caption         =   "ΑΠΟΣΤΟΛΗ ΕΝΑ ΕΝΑ"
      Height          =   360
      Left            =   16800
      TabIndex        =   43
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CheckBox CHXORISERR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "ΜΟΝΟ ΧΩΡΙΣ ERR"
      Height          =   315
      Left            =   2280
      TabIndex        =   42
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "XML"
      Height          =   360
      Left            =   16800
      TabIndex        =   41
      Top             =   7200
      Width           =   2295
   End
   Begin VB.CommandButton cmdDeleteMark 
      Height          =   495
      Left            =   6120
      TabIndex        =   39
      ToolTipText     =   "Διαγραφή MARK"
      Top             =   9240
      Width           =   135
   End
   Begin VB.CommandButton TESTXARAL 
      Caption         =   "MAZIKES AKYROSEIS DOKTIMCANC"
      Height          =   375
      Left            =   13200
      TabIndex        =   38
      Top             =   9120
      Width           =   3735
   End
   Begin VB.CommandButton RESEND_CANCEL 
      Caption         =   "Επαναποστολή Ακυρωτικού"
      Height          =   360
      Left            =   16800
      TabIndex        =   37
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton SXETIKO 
      Caption         =   "Συσχέτιση Πιστωτικού"
      Height          =   495
      Left            =   16800
      TabIndex        =   35
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CheckBox ektos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ολα εκτός τρέχουσας εγγρ."
      Height          =   480
      Left            =   2280
      TabIndex        =   34
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Μόνο τρέχουσα"
      Height          =   255
      Left            =   2280
      TabIndex        =   33
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdSendTest 
      Caption         =   "send-test"
      Height          =   360
      Left            =   12120
      TabIndex        =   32
      Top             =   10560
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdMARK 
      BackColor       =   &H000080FF&
      Caption         =   "Μαρκάρισμα για επαναποστολή"
      Height          =   480
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmddiagdokimi 
      Caption         =   "ΔΙΑΓΡΑΦΗ ΔΟΚΙΜΑΣΤΙΚΩΝ"
      Height          =   360
      Left            =   7800
      TabIndex        =   28
      Top             =   4680
      Width           =   2172
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "ΑΠΟΣΤΟΛΗ ΠΑΡΑΣΤΑΤΙΚΩΝ ΣΤΟ MYDATA"
      Height          =   600
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1920
      Width           =   3735
   End
   Begin VB.CommandButton APODOS 
      BackColor       =   &H0080FF80&
      Caption         =   "ΛΗΨΗ ΑΠΟ ΕΜΠΟΡΙΚΟ DOS"
      Height          =   480
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   12120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM TIM"
      Top             =   10800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command9 
      Caption         =   "cancelInvoice"
      Height          =   360
      Left            =   240
      TabIndex        =   25
      Top             =   10320
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.CommandButton RESTORELATHOS 
      Caption         =   "Επαναφορά ""ΛΑΘΟΣ.."" ΓΙΑ ΑΠΟΣΤΟΛΗ"
      Height          =   360
      Left            =   240
      TabIndex        =   24
      Top             =   4200
      Width           =   3732
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ΑΝΤΙΣΤΟΙΧΗΣΗ"
      Height          =   360
      Left            =   2280
      TabIndex        =   22
      Top             =   5760
      Width           =   1692
   End
   Begin VB.CommandButton SEND_EXPENS 
      Caption         =   "SEND_EXPENS"
      Height          =   360
      Left            =   240
      TabIndex        =   21
      Top             =   9600
      Width           =   1815
   End
   Begin VB.CommandButton cmdepanfromfile 
      Caption         =   "Επαναποστολή από αρχείο"
      Height          =   360
      Left            =   12240
      TabIndex        =   20
      Top             =   9240
      Visible         =   0   'False
      Width           =   3732
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Αρχείο ΤΕΧΤ λήψης"
      Height          =   360
      Left            =   240
      TabIndex        =   19
      Top             =   4680
      Width           =   3735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Αρχείο TEXT Αποστολής"
      Height          =   360
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Αρχείο TEXT απάντησης"
      Height          =   360
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CheckBox monoApestalmena 
      BackColor       =   &H00C0E0FF&
      Height          =   252
      Left            =   1320
      TabIndex        =   16
      Top             =   1560
      Width           =   252
   End
   Begin VB.CheckBox hkoLD 
      BackColor       =   &H0080FFFF&
      Caption         =   "Παλιότερες εγγραφές"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   5760
      Width           =   1932
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   14160
      Top             =   11280
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
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
      Bindings        =   "par7MyData.frx":001F
      Height          =   3495
      Left            =   4200
      TabIndex        =   12
      Top             =   1185
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   6165
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
      Splits(0).RecordSelectorWidth=   1005
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2672"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2752"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2672"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "ΛΙΣΤΑ ΠΑΡΑΣΤΑΤΙΚΩΝ ΜΑΣ"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80C0FF&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=161"
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
   Begin VB.CommandButton cmdΤιμολογηθενταΣε 
      BackColor       =   &H0000FFFF&
      Caption         =   "Λήψη τιμολογηθεντων σε εμας"
      Height          =   600
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CommandButton cmdTestantinvUpdateTim 
      Caption         =   "testantinvUpdateTim"
      Height          =   360
      Left            =   3480
      TabIndex        =   10
      Top             =   9840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "δικα μας παραστατικά που εχουν ανεβει στο MYDATA"
      Height          =   495
      Left            =   6600
      TabIndex        =   9
      Top             =   9000
      Width           =   5292
   End
   Begin VB.CommandButton UPLOAD 
      Caption         =   "UPLOAD INV"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   9840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "REQUEST ΤΙΜΟΛΟΓΙΩΝ ΠΟΥ ΕΣΤΕΙΛΑ ΣΤΟ SQLSERVER TABLE APESTALMENA"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   11520
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   3852
   End
   Begin VB.CommandButton cmdCommand3 
      Caption         =   "CREATE XML WITHOUT TEXTWRITER"
      Height          =   480
      Left            =   14040
      TabIndex        =   5
      Top             =   9720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdCommand4 
      Caption         =   "test DOMDocument CREATE XML"
      Height          =   480
      Left            =   12000
      TabIndex        =   4
      Top             =   9720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdCommand5 
      BackColor       =   &H0000FF00&
      Caption         =   "ΤΕΣΤ DOMDocument UTF8"
      Height          =   480
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdDOMDocumentUTF8 
      BackColor       =   &H0000FF00&
      Caption         =   "synergeio sendclient"
      Height          =   480
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker EOS 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1185
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   162594817
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker APO 
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   1185
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   162594817
      CurrentDate     =   36494
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
      Bindings        =   "par7MyData.frx":0034
      Height          =   3135
      Left            =   4200
      TabIndex        =   13
      Top             =   5040
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   5530
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
      Splits(0).RecordSelectorWidth=   1005
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).Caption=   "ΛΙΣΤΑ ΤΙΜΟΛΟΓΗΘΕΝΤΩΝ ΣΕ ΕΜΑΣ"
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2672"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2752"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2672"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=161"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   9600
      Top             =   10440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc AdoSYN 
      Height          =   495
      Left            =   15840
      Top             =   11400
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
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
      Caption         =   "AdoSYN"
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
   Begin TrueOleDBGrid80.TDBGrid tdbgridsyn 
      Bindings        =   "par7MyData.frx":0049
      Height          =   3375
      Left            =   16800
      TabIndex        =   30
      Top             =   1185
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5980
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
      Splits(0).RecordSelectorWidth=   1005
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2672"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=196612"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2672"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=196612"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
      Caption         =   "ΛΙΣΤΑ ΠΑΡΑΣΤΑΤΙΚΩΝ ΜΑΣ"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80C0FF&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=161"
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
      _StyleDefs(50)  =   ":id=35,.parent=33,.alignment=1,.valignment=2,.bgcolor=&H80000002&"
      _StyleDefs(51)  =   ":id=35,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   "Named:id=36:Selected"
      _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(54)  =   "Named:id=37:Caption"
      _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(56)  =   "Named:id=38:HighlightRow"
      _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(58)  =   "Named:id=39:EvenRow"
      _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(60)  =   "Named:id=40:OddRow"
      _StyleDefs(61)  =   ":id=40,.parent=33"
      _StyleDefs(62)  =   "Named:id=41:RecordSelector"
      _StyleDefs(63)  =   ":id=41,.parent=34"
      _StyleDefs(64)  =   "Named:id=42:FilterBar"
      _StyleDefs(65)  =   ":id=42,.parent=33"
   End
   Begin RichTextLib.RichTextBox XML 
      Height          =   975
      Left            =   4200
      TabIndex        =   40
      Top             =   8160
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   1720
      _Version        =   393217
      TextRTF         =   $"par7MyData.frx":005E
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   16800
      Top             =   9600
      _ExtentX        =   979
      _ExtentY        =   979
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":00E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":09BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":110C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":155E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":19B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":1E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":2254
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":26A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":2AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":2F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":35C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":3C3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":42B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":4932
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":4FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":5626
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par7MyData.frx":5CA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image PICTURE1 
      Height          =   372
      Left            =   17520
      Top             =   8520
      Width           =   372
   End
   Begin VB.Image Image1 
      Height          =   372
      Left            =   16920
      Top             =   8520
      Width           =   372
   End
   Begin VB.Label frmOptionsEos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   18240
      TabIndex        =   50
      Top             =   4680
      Width           =   45
   End
   Begin VB.Label frmOptionsapo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   18000
      TabIndex        =   49
      Top             =   4800
      Width           =   45
   End
   Begin MSForms.Label FRMOptionsafm 
      Height          =   375
      Left            =   16920
      TabIndex        =   48
      Top             =   4680
      Width           =   735
      BackColor       =   12640511
      Size            =   "1296;661"
      FontHeight      =   165
      FontCharSet     =   161
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Εως"
      Height          =   195
      Left            =   2000
      TabIndex        =   47
      Top             =   1200
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Από"
      Height          =   195
      Left            =   0
      TabIndex        =   46
      Top             =   1200
      Width           =   285
   End
   Begin ComctlLib.ImageList ImageList11 
      Left            =   18600
      Top             =   9960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   21
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":60F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":640C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":6726
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":6A40
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":6D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":7074
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":738E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":76A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":79C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":7CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":7FF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":8310
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":862A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":8944
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":8C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":8F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":9292
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":95AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":98C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":9BE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "par7MyData.frx":9EFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label SXETMART 
      Height          =   375
      Left            =   17520
      TabIndex        =   36
      Top             =   11280
      Width           =   1695
   End
   Begin VB.Label lblLabel2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".."
      Height          =   195
      Left            =   16920
      TabIndex        =   31
      Top             =   4560
      Width           =   90
   End
   Begin VB.Label where 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Height          =   195
      Left            =   480
      TabIndex        =   23
      Top             =   11280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aπεσταλμενα"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   1035
   End
End
Attribute VB_Name = "Par7MyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' να βαλουμε την χωρα GR ΣΤΟΥΣ ΕΛΛΗΝΕς   dokimastiko periballon parh afm=150144455  plagakis   ed99b37f3e99cc46ba284186185b5cdc
'EAN DEN EXEI AFM KAI EINAI TIMOLOGIO NA MPLOKAREI


Dim F_ANAPODO    As String ' = FINDPARAMETROI(3, "PAR1", "F_ANAPODO", "ΩΨ", "Κωδικοί παρ/κων με ανάποδο ΦΟΡΤΩΣΗ/ΠΡΟΟΡΙΣΜΟ ;")
        
Dim f_name_xml   As String ' για να βαζω ονομα στο xml
Dim f_SDA As String '     = FINDPARAMETROI(1, "PAR1", "F_SDA", "Ω;V", "ΧΑΡΑΚΤΗΡ.ΓΡΑΜΜΑ ΣΔΑ ΕΝΑΡΞΗΣ ΚΑΙ ΕΠΙΣΤΡΟΦΗΣ Π.Χ. Φ;Ψ") '
'Dim gf_Bearer As String
Dim f_YPOK       As String

Dim Col          As TrueOleDBGrid80.Column

Dim ColS         As TrueOleDBGrid80.columns

Dim f_Stelno_daiakin As Integer

Dim F_brat       As Integer
Dim F_demo       As Integer

Dim PAROX_URLADD As String

Dim PAROX_USER   As String

Dim PAROX_PASS   As String

Dim PAROX_KEY    As String
Dim F_posSignature As String

Dim f_pelono, f_peldie, f_pelarit, f_pelpol, f_peltk As String

Dim F_PAROCHOS As Integer

'F_PAROCHOS = Val(FINDPARAMETROI(1, "PAR7MYDATA", "f_parochos", "0", "0=MyData 1=Παροχος"))

Dim F_PAROX_SIGNATURE As String
Dim antFPA(10)

Dim timhFPA(10)        As Single

Dim F_TXTFILES         As String

Dim F_ID_NUMS(5000)    As Long ' ΑΠΟΘΗΚΕΥΕΙ ΤΑ ID_NUM ΓΙΑ ΝΑ ΤΑ ΘΥΜΑΤΑΙ ΟΤΑΝ ΕΡΘΕΙ Η ΑΠΑΝΤΗΣΗ ΑΠΟ ΑΑΔΕ

Dim F_dokimi           As Integer ' αριθμός δοκιμαστικου user

' ΠΩΣ ΣΤΕΛΝΩ ΧΑΡΑΚΤΗΡΙΣΜΟ ΕΣΟΔΩΝ
'imports System
'imports System.Net.Http.headers
'imports System.Text
'imports System.Net.Http
'imports System.Web
'
'Namespace CSHttpClientSample
'    Module Program
'        Private Sub Main()
'            MakeRequest()
'            Console.WriteLine ("Hit ENTER to exit...")
'            Console.ReadLine()
'        End Sub
'
'        Private Async Sub MakeRequest()
'            Dim client = New HttpClient()
'            Dim queryString = HttpUtility.ParseQueryString(String.Empty)
'            client.DefaultRequestHeaders.Add("{aade-user-id}", "{aade-user-id}")
'            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "{subscription key}")
'            Dim uri = "https://mydata-dev.azure-api.net/SendExpensesClassification?" & queryString
'            Dim response As HttpResponseMessage
'            Dim byteData As Byte() = Encoding.UTF8.GetBytes("{body}")
'
'            Using content = New ByteArrayContent(byteData)
'                content.Headers.ContentType = New MediaTypeHeaderValue("< your content type, i.e. application/json >")
'                response = Await client.PostAsync(uri, content)
'            End Using
'        End Sub
'    End Module
'End Namespace

'        <taxesTotals>
'            <taxes>
'                <taxType>1</taxType>
'                <taxCategory>2</taxCategory>
'                <underlyingValue>1000.00</underlyingValue>
'                <taxAmount>200.00</taxAmount>
'            </taxes>
'        </taxesTotals>

'Πεδίο           Τύπος        Υποχρεωτικό    Περιγραφή                  Αποδεκτές τιμές
'taxType         xs:byte        Ναι         Είδος Φόρου                  Λίστα τιμών:
'                                                                    1 = Παρακρατούμενος Φόρος
'                                                                    2 = Τέλη
'                                                                    3 = Λοιποί Φόροι
'                                                                    4 = Χαρτόσημο
'                                                                    5 = Κρατήσεις
'taxCategory     xs:byte       Όχι        Κατηγορία Φόρου                 Ελάχιστη τιμή = 1
'underlyingValue xs:decimal    Όχι        Υποκείμενη Αξία            Ελάχιστη τιμή = 0
'                                                                    Δεκαδικά ψηφία = 2
'taxAmount       xs:decimal    Ναι          Ποσό Φόρου               Ελάχιστη τιμή = 0
'                                                                     Δεκαδικά ψηφία = 2
'id              xs:byte       Όχι      Αύξων αριθμός γραμμής
'Παρατηρήσεις:
'1) Το πεδίο taxCategory μπορεί να πάρει κάθε φορά οποιαδήποτε τιμή από τον αντίστοιχο πίνακα του Παραρτήματος του φόρου που αναφέρεται στο πεδίο taxType
'2) Το πεδίο underlyingValue υποδηλώνει την αξία στην οποία εφαρμόζεται ο συγκεκριμένος φόρος

Dim f_par1ypok         As String ' παραστατικά που ανήκουν στο υποκ/μα 1

Dim f_par2ypok         As String ' παραστατικά που ανήκουν στο υποκ/μα 1

Dim fSynt              As String

Dim fKatEsod(10)       As String

Dim fKatEXod(10)       As String

Dim FkatE3Esod(10)     As String

Dim F_E3_APOKATHG_EID  As Integer

Dim sql                As String

'Dim gdb As New ADODB.Connection
Dim MARK               As String

Dim fMydataFromEID1    As Integer

Dim P_USER             As String

Dim P_KEY              As String

Dim afmCompany         As String

Dim UrlAADE            As String

Dim url2               As String

'Dim sqldt As New ADODB.Recordset
Dim AFM                As String ' ΠΕΛΑΤΗΣ COUNTERPART

Dim AFM0               As String ' ΕΚΔΟΤΗς ISSUER

Dim issueDate          As String

Dim AA                 As String

Dim invoiceType        As String

Dim f_apoDOS           As Integer

Dim paytype            As String

Dim payaji             As String

Dim totalNetValue      As String

Dim totalVatAmount     As String

Dim totalGrossValue    As String

Dim docStock           As MSXML2.DOMDocument

Dim classificationType As String

Const MinDate          As Date = #1/1/1900# '-- set MinDate as you need

'-- 2^31 is used as an offset for 32-bit-sign-integer

Dim fLong              As Long  ' το χρησιμοποιω σαν δεικτη να ξερω το update tim τι θα ενημερωνει

' βγαζω ενα ακεραιο μοναδικο απο την ωρα για να βαζω σε καθε αποστολη για να ειναι μοναδική
Function Dt2Lng(aDate As Date) As Long
        '-- acept datetime range from MinDate to
        '-- (MinDate + 49710 days + 6:28:15) ~ 136 years
100     Dt2Lng = CLng((aDate - MinDate) * 86400 - 2 ^ 31)  ' 2^31 ειναι ο μεγαλυτερος ακεραιοσ του sqlserver
End Function

' ton akeraio ayton ton kanei date
Function Lng2Dt(N As Long) As Date
100     Lng2Dt = MinDate + (N + 2 ^ 31) / 86400
End Function

Private Sub APO_Change()

        '<EhHeader>
        On Error GoTo APO_Change_Err

        '</EhHeader>
        Dim sql As String

100     sql = FIND_QUERY()

102     Adodc1.ConnectionString = gConnect
104     Adodc1.RecordSource = sql
106     Adodc1.Refresh

        '<EhFooter>
        Exit Sub

APO_Change_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.APO_Change " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub APODOS_Click()

100     update_sql_from_dbf "PEL"

        'update2_sql_from_dbf "SELECT * FROM EGGTIM ORDER BY HME DESC", "SELECT top 1 * FROM EGGTIM"
        'update2_sql_from_dbf "SELECT * FROM TIM ORDER BY HME DESC", "SELECT top 1 * FROM TIM"

        Dim synt As String

102     synt = " where  HME>=#" + Format(APO.Value, "MM/dd/yyyy") + "#  AND HME<#" + Format(DateAdd("d", 1, EOS.Value), "MM/dd/yyyy") + "# "
    
104     update2_sql_from_dbf "SELECT * FROM EGGTIM " + synt + " ORDER BY HME DESC", "SELECT top 1 * FROM EGGTIM"
106     update2_sql_from_dbf "SELECT * FROM TIM " + synt + " ORDER BY HME DESC", "SELECT top 1 * FROM TIM"

108     Gdb.Execute "UPDATE EGGTIM SET APOT=1, ID_NUM = (SELECT TOP 1 ID_NUM FROM TIM WHERE ATIM=EGGTIM.ATIM AND HME=EGGTIM.HME) WHERE ID_NUM IS NULL"

        'ΜΗΔΕΝΙΖΩ ΤΑ ΑΡΝΗΤΙΚΑ ΚΕΝΑ
110     Gdb.Execute " UPDATE TIM SET AJ5=0,AJI=AJI-AJ5 WHERE AJ5<0 "

112     Gdb.Execute "UPDATE TIM SET B_N1=1 WHERE B_N1 IS NULL"
114     Gdb.Execute "UPDATE TIM SET TRP='1.ΜΕΤΡ' WHERE TRP LIKE 'ΜΕ%' "
116     Gdb.Execute "UPDATE TIM SET TRP='2.ΠΙΣΤ' WHERE TRP LIKE 'ΠΙ%' "

118     APO.SetFocus
120     EOS.SetFocus

        Dim R As New ADODB.Recordset

122     R.Open "SELECT COUNT(*) FROM EID WHERE KOD='99~'", Gdb, adOpenDynamic, adLockOptimistic

124     If R(0) = 0 Then
126         Gdb.Execute "insert into EID (KOD,ONO,FPA,APALLFPA) VALUES ('99~','KENA',5,6)"
        End If

128     R.Close

        ' UPDATE TIM SET AJI=AJI+AJ5 WHERE AJ5<0;
        ' UPDATE TIM SET AJ5=0 WHERE AJ5<0

130     MsgBox "OK"

End Sub

Private Sub update2_sql_from_dbf(arxeio As String, SQLQ As String)

        '<EhHeader>
        On Error GoTo update_sql_from_dbf_Err

        '</EhHeader>
        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim rSQL2  As New ADODB.Recordset

        Dim conDBF As String

        Dim conSQL As String

        Dim k      As Long

        Dim Fname  As String

        ' G2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
        '2004 G2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
100     conSQL = gConnect    ' "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP;Data Source=LAGAKIS"
102     conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO"
        '2004 conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO\04"

104     sql.Open conSQL
        'SET DBF=DATA.Open conDBF

106     Data1.DatabaseName = "C:\LAGEURO"   'Text1.Text

108     Data1.RecordSource = arxeio
110     Data1.Refresh

112     Set rDBF = Data1.Recordset

        'Label1.Caption = arxeio + " " + Format(rDBF.RecordCount, "######")
        'sql.Execute "DELETE FROM " + arxeio

114     rSQL.Open SQLQ, sql, adOpenDynamic, adLockOptimistic
        ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

116     rDBF.MoveFirst

        Dim Z

118     Z = 0

        Dim j As Integer

120     If InStr(arxeio, " TIM") > 0 Then
122         j = 1
        End If

        On Error GoTo printERROR    ' Resume Next

124     Do While Not rDBF.EOF
            
            '        Case ?
            
            '   End Select

126         rSQL2.Open "SELECT COUNT(*) FROM TIM WHERE ATIM='" + rDBF("ATIM") + "' AND HME='" + Format(rDBF("HME"), "mm/dd/yyyy") + "'", sql, adOpenDynamic, adLockOptimistic

128         If rSQL2(0) > 0 Then
130             rSQL2.Close
                 
132             GoTo lab580
                'rSQL2.Close
                'Me.Caption = "ΔΕΝ ΥΠΑΡΧΟΥΝ ΝΕΑ ΤΙΜΟΛΟΓΙΑ"
                'Exit Do
            End If

134         rSQL2.Close
            
136         rSQL.AddNew

138         For k = 0 To rDBF.FIELDS.Count - 1
140             Fname = rDBF.FIELDS(k).Name    ' p.x. FNAME=epo    R(0).NAME

142             If IsNull(rDBF(k)) Then
                Else
144                 rSQL(Fname) = rDBF(k)    ' rsqk("epo")=r(0)
                End If

                ' ΑΥΤΑ ΠΟΥ ΕΙΝΑΙ ΝULL ΤΑ ΓΕΜΙΖΕΙ
146             If IsNull(rDBF(k)) Then
148                 If rDBF(k).Type = dbText Then    '10=STRING
150                     rSQL(Fname) = " "
152                 ElseIf rDBF(k).Type = dbDate Then
154                     rSQL(Fname) = CDate("01/01/1990")
156                 ElseIf rDBF(k).Type = dbBoolean Then
158                     rSQL(Fname) = False
160                 ElseIf rDBF(k).Type = dbDouble Then
162                     rSQL(Fname) = 0
                    Else
164                     conSQL = ""
                    End If
                End If

            Next

            ' ΜΕΤΑΤΡΕΠΕΙ ΤΗΝ "ΧRΕOSI" ΣΕ "XRE" KAI "PISTOSI"
166         If arxeio = "EGG" Then
168             rSQL("PISTOSI") = 0: rSQL("XREOSI") = 0

170             If Not IsNull(rSQL("XRE")) Then
172                 If rSQL("EIDOS") = "e" Then
174                     If rSQL("XRE") > 0 Then rSQL("XREOSI") = rSQL("XRE")
176                     If rSQL("XRE") < 0 Then rSQL("PISTOSI") = -rSQL("XRE")
                    Else

178                     If rSQL("XRE") > 0 Then rSQL("PISTOSI") = rSQL("XRE")
180                     If rSQL("XRE") < 0 Then rSQL("XREOSI") = -rSQL("XRE")
                    End If
                End If
            End If

            ' METΑΤΡΕΠΕΙ ΤΟ 'POSO' ΣΕ 'XRE' KAI 'PIS'
182         If arxeio = "EGGTIM" Then
184             rSQL("PIS") = 0: rSQL("XRE") = 0

186             If Not IsNull(rSQL("POSO")) Then
188                 If InStr(F_xre_APOU + F_axre_APOU, Left(rSQL("ATIM"), 1)) > 0 Then
190                     rSQL("XRE") = rSQL("POSO")
                    End If

192                 If InStr(F_pis_APOU + F_apis_APOU, Left(rSQL("ATIM"), 1)) > 0 Then
194                     rSQL("PIS") = rSQL("POSO")
                    End If
                End If
            End If

196         rSQL.Update
198         Z = Z + 1

200         If j = 1 Then 'Z Mod 10 = 0 Then
202             Me.Caption = "ΝΕΑ ΤΙΜΟΛΟΓΙΑ " + str(Z)
            End If
lab580:
204         rDBF.MoveNext

206         DoEvents

        Loop

208     rSQL.Close
210     rDBF.Close

        Exit Sub

printERROR:
        '  List2.Visible = True

        ' List2.AddItem Err.Description

212     Resume Next

        '<EhFooter>
        Exit Sub

update_sql_from_dbf_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.update_sql_from_dbf " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.update_sql_from_dbf " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub suma_hmeron()

        Dim synola(1000) As Single

        ' BRISKO TIS SUMES
        'For K = 0 To 1
100     synola(0) = 0
102     synola(1) = 0


104     Do While Not AdoSYN.Recordset.EOF
106         synola(0) = synola(0) + AdoSYN.Recordset(0)
108         synola(1) = synola(1) + AdoSYN.Recordset(1)

110         AdoSYN.Recordset.MoveNext
            'TDBGRIDSYN.TextMatrix(L, K) = Format(TDBGRIDSYN.TextMatrix(L, K), "####,##0.00")
        Loop

        ' GR2.TextMatrix(GR2.rows - 1, K) = Format(SYNOLA(K), "####,##0.00")
        ' tdbgridsyn.Splits(0).columns(K).FooterText = Format(synola(K), "######0.00")
        ' tdbgridsyn.Splits(0).columns(K).FooterFont.Size = 12 ' = Format(SYNOLA(k), "######0.00")

        'Next K

        ' TDBGRIDSYN.FooterFont.Size = 12

112     lblLabel2.Caption = "Σύνολα καθ.αξία :" + Format(synola(0), "######0.00") + " και ΦΠΑ :" + Format(synola(1), "######0.00")

End Sub



 Public Sub CancelDelivery(extID As String, MARK As String)
        Dim Reqpos As WinHttp.WinHttpRequest
        Set Reqpos = New WinHttp.WinHttpRequest
'        Dim cancel As New CancelDelivery
'        cancel.mark = mark
'        cancel.externalSystemId = extID
'        Dim ResultJSONpos As String: ResultJSONpos = JsonConvert.SerializeObject(cancel).toString
        
        Dim Q As String
        Q = " { ""mark"":" + MARK + ","
        Q = Q + " ""externalSystemId"":""" + extID + """ }"
        
        
            Dim uripos As String: uripos = gfURL1 + "api/Cancel"
680:        With Reqpos
                gf_Bearer = REQPAROCHOS
                .Open "POST", gfURL1 + "api/Cancel", async:=False
                .setRequestHeader "Content-Type", "application/hal+json"
                .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
                .setRequestHeader "Authorization", "Bearer " + gf_Bearer
                .setRequestHeader "X-Version", "3"
                
                Q = Replace(Q, "ΔΟΚ", "")
                .send (Q)
                MsgBox ("CancelPosPayment" + Chr(13) + Reqpos.responseText)
            End With
            
172        Open "C:\TXTFILES\CANCDELIV.TXT" For Output As #5
174         Print #5, Reqpos.responseText
176         Close #5
178         Set ser = JSON.parse(Reqpos.responseText)
180         'cMARK = ser.Item("cancelationMark")
182         cMARK = ser.Item("cancellationMark")
            
            
            
            
            
'            Dim posjson As String: posjson = Reqpos.responseText
'            MsgBox ("Cancel" + Chr(13) + posjson)
'           ' Dim posser As String: 'JObject = JObject.parse(posjson)
           MILSEC 500
            
            Gdb.Execute "update TIM SET INCMARK='" + LTrim(str(cMARK)) + "' WHERE ID_NUM=" + str(extID)
            
            
            
    End Sub
















Private Sub cancelClient_Click()

        Dim ANS    As String, TYPOS As String

        Dim ARKYKL As String

100     ANS = InputBox("2=PARKING 3=GARAGE", "")

102     ARKYKL = InputBox("αρ.κυκλοφοριας", "")

104     If Val(ANS) = 2 Then
106         TYPOS = "dcrnew:parkingcarwash"
        Else
108         TYPOS = "dcrnew:garage"
        End If
   
        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
110     Set docStock = New MSXML2.DOMDocument

112     With docStock
114         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
116         Set elemRoot = .createElement("dcrudtcor:ClientCorrelationDoc")
118         elemRoot.setAttribute "xmlns:dcrudtcor", "http://www.aade.gr/myDATA/dcrudtcor/v1.0"
120         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
122         elemRoot.setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/dcrudtcor/v1.0 SimpleTypes-v1.0.xsd clientCorrelationType-v1.0.xsd"

124         Set elemStock = docStock.createElement("dcrudtcor:clientCorrelation")

126         Set elem2Field = docStock.createElement("dcrudtcor:entityVatNumber"): elem2Field.Text = afmCompany: elemStock.appendChild elem2Field
                     
128         Set elem2Field = docStock.createElement("dcrudtcor:mark"): elem2Field.Text = "400000000000000": elemStock.appendChild elem2Field
            ' Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = "202502221033: elemStock.appendChild elem2Field"
130         Set elem2Field = docStock.createElement("dcrudtcor:correlatedDCLids"): elem2Field.Text = 100000000007161#: elemStock.appendChild elem2Field
                       

                    
132         elemRoot.appendChild elemStock
  
            '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
134         Set .documentElement = elemRoot

            On Error Resume Next
            
136         FormatXmlDocument docStock ' βαζει κενα να ειναι ευκολο στο διαβασμα

            Dim ANS2 As Integer

138         ANS2 = MsgBox(docStock.XML, vbYesNo)

140         If ANS2 = vbNo Then

                Exit Sub

            End If
            
142         Kill "C:\txtfiles\" + ARKYKL + "r.xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

144         .save "C:\txtfiles\" + ARKYKL + "r.xml"
        End With

        ' TO XML EINAI docStock.xml

        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
     
        Dim MARK As String

        '    MARK = InputBox("δωσε τo mark", "μαρκ")
     
146     url2 = url2 + MARKTIM
     
148     url2 = "https://mydataapidev.aade.gr/DCL/CancelClient?dclID=" + "100000000007161"

        'initialize
150     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
152     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
154     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
156     XMLServer.Option(9) = 2048
158     XMLServer.Option(6) = True
    
160     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
162     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
164     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

166     XMLServer.send 'docStock.XML 'TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
168     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
170     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
 
172     MsgBox v

174     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

176     mfile = "c:\" + F_TXTFILES + "\sendinv\aytok" + "-" + ARKYKL + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
178     Debug.Print v
180     Open mfile For Output As #8
182     Print #8, v
184     Close #8
    
186     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
        '  READ_CANCEL v, MARK2

        '  400001829337674

        '    CANCEL_INVOICE = MARK2
    
        '<EhFooter>
        Exit Sub

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.RELATE_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub
Public Function OLDkillClient(ByVal park2garage3 As Integer, ByVal clientID As String) As Integer
     Dim ANS    As String, TYPOS As String



104     If park2garage3 = 2 Then
106         TYPOS = "dcrnew:parkingcarwash"
        Else
108         TYPOS = "dcrnew:garage"
        End If
   
     

        Dim i         As Integer
    


        Dim MARK As String

        '    MARK = InputBox("δωσε τo mark", "μαρκ")
     
146     url2 = url2 + MARKTIM
     
148     url2 = UrlAADE + "DCL/CancelClient?dclID=" + clientID ' "100000000007161"

        'initialize
150     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
152     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
154     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
156     XMLServer.Option(9) = 2048
158     XMLServer.Option(6) = True
    
160     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
162     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
164     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

166     XMLServer.send 'docStock.XML 'TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
168     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
170     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
 
172     MsgBox v

174     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

176     mfile = "c:\" + F_TXTFILES + "\sendinv\aytok" + "-" + ARKYKL + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
178     Debug.Print v
180     Open mfile For Output As #8
182     Print #8, v
184     Close #8
    
186     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
    
    OLDkillClient = 1
    
    
        
        Exit Function

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.RELATE_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function





Private Sub Check1_Click()

100     If Check1.Value = vbChecked Then
102         ektos.Value = vbUnchecked
        End If

End Sub

Private Sub CHXORISERR_Click()

        Dim sql As String

100     sql = FIND_QUERY()

        'If monoApestalmena.Value = vbUnchecked Then

102     Adodc1.ConnectionString = gConnect
104     Adodc1.RecordSource = sql
106     Adodc1.Refresh

        '<EhFooter>
        
        'Else
        '    monoApestalmena_Click

        'End If

End Sub

Function getJson(sInputJson As String) As String
 
        Dim p As Object
   
        ' Dim sInputJson As String
        '  sInputJson = "{ width: '200', frame: false, height: 130, bodyStyle:'background-color: #ffffcc;',buttonAlign:'right', items: [{ xtype: 'form',  url: '/content.asp'},{ xtype: 'form2',  url: '/content2.asp'}] }"
   
        '  MsgBox "Input JSON string: " & sInputJson
   
        ' sets p
100     Set p = JSON.parse(sInputJson)
   
        'MsgBox "Parsed object output: " & JSON.toString(p)
   
102     getJson = JSON.toString(p)
   
        'MsgBox "Get Bodystyle data: " & p.Item("bodyStyle")
   
        'MsgBox "Get Form Url data: " & p.Item("items").Item(1).Item("url")
   
        'p.Item("items").Item(1).ADD "ExtraItem", "Extra Data Value"
   
        'MsgBox "Parsed object output with added item: " & JSON.toString(p)

End Function

Function REQPAROCHOS() As String

        Dim URI As String

100     URI = "https://" + PAROX_URLADD + "account.parochos.gr/api/account/loginToSubscription"

        Dim mEMAIL As String

102     mEMAIL = PAROX_USER ' "glagakis@gmail.com"

        '
        On Error GoTo 0

        Dim REQ As WinHttp.WinHttpRequest

104    Set REQ = New WinHttp.WinHttpRequest

        Dim Q As String

106     With REQ
108         .Open "POST", URI, async:=False
110         .setRequestHeader "Content-Type", "application/hal+json"
112         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
            'Note: Normally you don't include all of this whitespace, but
            'we'll use it in this example:
            ' q = "{  ""subscriptionKey"": ""32657FD7929B483AABBD1C633401E945"" , ""username"": """ & mEMAIL & """ ," & """password"":""y39213921!!Y""," & """email"": """ & mEMAIL & """  " & "}"
114         Q = "{  ""subscriptionKey"": """ & PAROX_KEY & """ , ""username"": """ & mEMAIL & """ ," & """password"":""" & PAROX_PASS & """," & """email"": """ & mEMAIL & """  " & "}"
116         .send (Q)
118        ' MsgBox (REQ.responseText)
        End With
        
        Dim CC   As String

        'CC = json.toString(q)


On Error Resume Next

120     Dim JSON2 As String: JSON2 = REQ.responseText
        
122    ' REQPAROCHOS = Split(Split(JSON, ":")(1), ",")(0)
        'Dim JSON2 As String
       ' JSON2 = JSON
        
124     'gfURL1 = Split(Split(JSON, ":")(3), ",")(0) Set ser = JSON.parse(ANSW)
        
        '
         Dim ser As Object: Set ser = JSON.parse(REQ.responseText)
        '        Dim c As String: c = ""
        '
        '        '  For K As Integer = 0 To ser("jwt").Count - 1
        '
        '
               ' gf_Bearer = ser.Item(1)("jwt")
               ' gfURL1 = ser("url1").toString
                
                gf_Bearer = ser("jwt")
                gfURL1 = ser("url1")
                
                
        '
        '
                REQPAROCHOS = gf_Bearer
        
        '
        '
        '

End Function

Private Sub CMDDEHTEST_Click()
 
 Dim ARPAROX As String
 ARPAROX = InputBox("ΔΩΣΕ ΑΡΙΘΜΟ ΠΑΡΟΧΗΣ", "ΑΡ.ΑΠΑΡΟΧ", "824048676")
Dim URI As String
 URI = "https://www1.aade.gr/aadeapps5/oilbft-deddhe/api/v1/checkDeddhe/" + ARPAROX '"https://" + PAROX_URLADD + "account.parochos.gr/api/account/loginToSubscription"

        'initialize  100000000007664
114     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
116     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
118     XMLServer.setTimeouts 5000, 60000, 10000, 10000

        '130     XMLServer.Open "GET", url2 & "?mark=" + MARK2, False
    
        'force TLS 1.2
120     XMLServer.Option(9) = 2048
122     XMLServer.Option(6) = True
    
124     XMLServer.Open "GET", URI, False
       
126    ' XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
128     'XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

130   On Error Resume Next
        XMLServer.send 'docStock.XML 'TXML 'strData  'S????O ?? STRING TXML STO MYDATA  'XMLServer.send TXML '  ?? T??O ?? ?? S????O ??? ?????? ???? ???O strData
132     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
134     v = XMLServer.responseText
        '--------------- ???S???? INVOICES -----------------------------------------------------
      If UBound(Split(v, ":")) = 1 Then
          MsgBox "FALSE " + v
      Else
          MsgBox Split(v, ":")(3) + "    " + v
      End If
      
      
136    ' MsgBox v
End Sub

Private Sub cmdDOMDocumentUTF8_Click()

        '<EhHeader>
        '  On Error GoTo cmdDOMDocumentUTF8_Click_Err
        '</EhHeader>

        'aade
        '<InvoicesDoc xmlns="http://www.aade.gr/myDATA/invoice/v1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"

        'xsi:            schemaLocation = "http://www.aade.gr/myDATA/invoice/v1.0/InvoicesDoc-v0.6.xsd"
        'xmlns:            icls = "https://www.aade.gr/myDATA/incomeClassificaton/v1.0"
        '            xmlns:ecls="https://www.aade.gr/myDATA/expensesClassificaton/v1.0">

        'xsd
        '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:dcr="http://www.aade.gr/myDATA/dcr/v1.0" targetNamespace="http://www.aade.gr/myDATA/dcr/v1.0" elementFormDefault="qualified" version="1.0">
        '   <xs:include schemaLocation="DigitalClientsDoc-v1.0.xsd"/>

        '<?xml version="1.0" encoding="utf-8"?>      ----------------------------------- paris
        '<dcrnew:NewDigitalClientDoc  --------
        'xmlns:    dcrnew = "http://www.aade.gr/myDATA/dcrnew/v1.0"---------------------
        'xmlns:    xsi = "http://www.w3.org/2001/XMLSchema-instance" -------------------
        '    xsi:schemaLocation="http://www.aade.gr/myDATA/dcrnew/v1.0 SendClient-v1.0.xsd">----------------------
        '    <dcrnew:newDigitalClient>
        '        <dcrnew:clientServiceType>3</dcrnew:clientServiceType>
        '        <dcrnew:useCase>
        '            <dcrnew:vehicleRegistrationNumber>RMT4231</dcrnew:vehicleRegistrationNumber>
        '        </dcrnew:useCase>
        '    </dcrnew:newDigitalClient>
        '</dcrnew:NewDigitalClientDoc>                  --------------------------------------

        Dim ANS    As String, TYPOS As String

        Dim ARKYKL As String

100     ANS = InputBox("2=PARKING 3=GARAGE", "")

102     ARKYKL = InputBox("αρ.κυκλοφοριας", "")

        Dim mIDcar As String

104  '   mIDcar = newclient(ANS, ARKYKL, "", "")

End Sub

Public Function OLD_newclient(ByVal garaz3epark2 As Integer, _
                          ark As String, _
                          vFACTORY As String, _
                          vCategory As String) As String

100     If Val(garaz3epark2) = 2 Then
102         TYPOS = "dcrnew:parkingcarwash"
        Else
104         TYPOS = "dcrnew:garage"
        End If

        Dim varStock  As Variant

        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
106     varStock = Array(Array("ΕΥΡΩ12345", 10.32), Array("¥45632", 5.43)) 'Yen sign used here to show Unicode.  '
    
108     Set docStock = New MSXML2.DOMDocument

110     With docStock
112         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
114         Set elemRoot = .createElement("dcrnew:NewDigitalClientDoc")

116         elemRoot.setAttribute "xmlns:dcrnew", "http://www.aade.gr/myDATA/dcrnew/v1.0"
118         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
120         elemRoot.setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/dcrnew/v1.0 SendClient-v1.0.xsd"

122         Set elemStock = docStock.createElement("dcrnew:newDigitalClient")
            ' With elemStock----------------------------------2002-05-30T09:00:00
                   
124         Set elem2Field = docStock.createElement("dcrnew:clientServiceType"): elem2Field.Text = garaz3epark2: elemStock.appendChild elem2Field
                      
            'HTAN OK ALLA EBGALE MHNYMA (NEXT LINE)     Set elem2Field = docStock.createElement("dcrnew:creationDateTime"): elem2Field.Text = Format(Now, "yyyy-MM-ddTHH:mm:ss"): elemStock.appendChild elem2Field
            '<message>creationDateTime is allowed only when transmissionFailure has value</message>
                  
126         Set elem2Field = docStock.createElement("dcrnew:entityVatNumber"): elem2Field.Text = afmCompany: elemStock.appendChild elem2Field
128         Set elem2Field = docStock.createElement("dcrnew:branch"): elem2Field.Text = "0": elemStock.appendChild elem2Field
               
130         Set elemField = docStock.createElement("dcrnew:useCase")
132         Set elem1Field = docStock.createElement(TYPOS)
                              
134         Set elem2Field = docStock.createElement("dcrnew:vehicleRegistrationNumber"): elem2Field.Text = ark: elem1Field.appendChild elem2Field
136         Set elem2Field = docStock.createElement("dcrnew:vehicleCategory"): elem2Field.Text = vCategory: elem1Field.appendChild elem2Field
138         Set elem2Field = docStock.createElement("dcrnew:vehicleFactory"): elem2Field.Text = vFACTORY: elem1Field.appendChild elem2Field
                              
140         elemField.appendChild elem1Field
                       
142         elemStock.appendChild elemField
                    
144         elemRoot.appendChild elemStock
  
            '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
146         Set .documentElement = elemRoot

            On Error Resume Next
            
148         FormatXmlDocument docStock ' βαζει κενα να ειναι ευκολο στο διαβασμα
150         MsgBox docStock.XML
            
152         Kill "C:\txtfiles\" + ARKYKL + ".xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

154         .save "C:\txtfiles\" + ARKYKL + ".xml"
        End With
  
156     url2 = UrlAADE + "CancelInvoice?mark=" '/RequestDocs"
     
        Dim MARK As String

        '    MARK = InputBox("δωσε τo mark", "μαρκ")
     
158     url2 = url2 + MARKTIM
     
      If F_demo = 1 Then
160      url2 = "https://mydataapidev.aade.gr/DCL/SendClient"
      Else
      
      url2 = "https://mydatapi.aade.gr/DCL/SendClient"

        ' url2 = UrlAADE + "DCL/SendClient"
        'initialize
      End If
        
        
        
        
        
        
        
162     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
164     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
166     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
168     XMLServer.Option(9) = 2048
170     XMLServer.Option(6) = True
    
172     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
174     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
176     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
178     XMLServer.send docStock.XML 'TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
180     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
  
182     v = XMLServer.responseText

        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
        Dim CC As String

184     CC = takefromxml(v)
186     newclient = CC
188     MsgBox v

190     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

192     mfile = "c:\" + F_TXTFILES + "\sendinv\aytok" + "-" + ARKYKL + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
194     Debug.Print v
196     Open mfile For Output As #8
198     Print #8, v
200     Close #8
    
202     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
        '  READ_CANCEL v, MARK2

        '  400001829337674

        '    CANCEL_INVOICE = MARK2
    
        '<EhFooter>
        Exit Function

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdDOMDocumentUTF8_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

        'Πεδίο 'Τύπος'Υποχρεωτικό'Περιγραφή'Αποδεκτές τιμές
        'idDcl'xs:long'Όχι'Μοναδικός Αριθμός Ψηφιακού Πελατολογίου'Συμπληρώνεται από την Υπηρεσία
        'ClientServiceType'ClientServiceType'Ναι'Τύπος Πελατολογίου'1 = Ενοικίαση'2 = Πάρκινγκ/'Digital Client REST API 13'Πλυντήρια'3 = Συνεργεία
        'creationDateTime'xs: DateTime'Όχι'Ημερομηνία και ώρα δημιουργίας εγγραφής στο Ψηφιακό Πελατολόγιο
        'entityVatNumber'xs:string'Όχι'ΑΦΜ Οντότητας Αναφοράς'entityVatNumber
        'branch'xs:int'Όχι'Αρ.Εγκατάστασης
        'recurringService'xs:boolean'Όχι'Επαναλαμβανόμενη Υπηρεσία
        'continuousServic'xs:boolean'Όχι'Διαρκής Υπηρεσία
        'fromAgreedPeriodDate'xs: Date'Όχι'Από Συμφωνημένη Περίοδο
        'toAgreedPeriodDate'xs: Date'Όχι'Έως Συμφωνημένη Περίοδο
        'mixedService'xs:boolean'Όχι'Μεικτή Υπηρεσία
        'customerVatNumber xs:string Όχι ΑΦΜ Πελάτη
        'customerCountry'CountryType'Όχι'Χώρα Πελάτη
        'transmissionFailure'xs:byte'Όχι'Αδυναμία Επικοινωνίας ή Αδυναμία διαβίβασης'Οι τιμές {1,2} είναι επιτρεπτές
        'correlatedDclId'xs:long'Όχι'Συσχετιζόμενα DC ID
        'Comments'Comments'Όχι'Σχόλια/ Παρατηρήσεις
        'useCase'ServiceType'Ναι'Περιπτώσεις χρήσης

End Function

Private Sub cmdEXPORTTOPDF_Click()

        Dim FILENAME As String

100     FILENAME = "C:\PDF\TEST.PDF"

        Dim FcrystalFilename As String

102     FcrystalFilename = "C:\MERCVB\REPORTS\TIMOL1.RPT"

104     ExportReportToPDF FcrystalFilename, FILENAME, "TEST"
106     MsgBox FileLen(FILENAME)

End Sub

Private Sub cmdPAROCHOS_Click()

        Dim c As String
      If F_dokimi = 3 Then
         ' makeJsonClass
           CreateSimSign
           
           
          ' c = getBratToken()
      Else
100     c = getToken()

      End If
      
      
      
102     MsgBox c


End Sub

Private Sub makeJsonClass()

' Dim invoice As Object, counterpart As Object, address As Object
'Set invoice = CreateObject("Scripting.Dictionary")
'Set counterpart = CreateObject("Scripting.Dictionary")
'Set address = CreateObject("Scripting.Dictionary")
'
'address("city") = "T000122"
'address("postalCode") = "54352"
'counterpart("vatNumber") = "026883248"
'counterpart("address") = address
'invoice("tidNsp") = "54888913"
'invoice("counterpart") = counterpart
'
'Dim invoices As Collection
'Set invoices = New Collection
'invoices.ADD invoice
'
'Dim root As Object
'Set root = CreateObject("Scripting.Dictionary")
'root("invoice") = invoices
'
'Dim jsonString As String
'
'jsonString = ConvertToJson(root)
'


End Sub


Private Sub CreateSimSign()

Dim JS As String
JS = "{""invoice"": [{" & _
"""B2G"": null," & _
"""counterpart"": {""address"": {""city"": ""ΘΕΣΣΑΛΟΝΙΚΗ"",""number"": null,""postalCode"": ""54352"",""street"": null},""branch"": 0,""country"": ""GR"",""countryDocumentId"": null,""documentIdNo"": null,""name"": null,""supplyAccountNo"": null,""vatNumber"": ""026883248""}," & _
"""downloadingInvoiceUrl"": null," & _
"""extra"": {""altCustAddress"": null,""altCustName"": null,""customerActivity"": ""ΧΟΝΔΡΙΚΟ ΕΜΠΟΡΙΟ ΗΛΕΚΤΡΟΝΙΚΩΝ ΥΠΟΛΟΓΙΣΤΩΝ, ΠΕΡΙΦΕΡΕΙΑΚΟΥ ΕΞΟΠΛΙΣΜΟΥ ΥΠΟΛΟΓΙΣΤΩΝ ΚΑΙ ΛΟΓΙΣΜΙΚΟΥ"",""customerCity"": ""ΘΕΣΣΑΛΟΝΙΚΗ"",""customerCode"": null,""customerDoyCode"": 4216,""customerDoyName"": ""ΣΤ ΘΕΣΣΑΛΟΝΙΚΗΣ"",""customerEmail"": ""info@bratnet.gr"",""customerName"": ""ΤΕΣΤ ΠΕΛΑΤΗΣ"",""customerPhone"": ""2310221222"",""customerSendEmail"": false,""customerStreetName"": ""ΛΑΜΠΡΑΚΗ 150"",""customerTk"": ""54352"",""customerVat"": ""026883248"",""destinationAddress"": null,""hotelEndDate"": null,""hotelRoom"": null,""hotelStartDate"": null,""invoiceRemarks"": null,""invoiceTypeName"": ""Τιμολόγιο Παροχής Υπηρεσιών"",""loadingAddress"": null,""movePurpose"": null,""nationalID"": null,""nspCode"": ""01"",""paymentMethodName"": ""Πίστωση 30 ημερών"",""salerActivity"": ""ΕΜΠΟΡΙΑ ΛΟΓΙΣΜΙΚΟΥ"",""salerAdditionalStreetName"": ""ΑΝΩ ΤΟΥΜΠΑ"",""salerCity"": ""ΘΕΣΣΑΛΟΝΙΚΗ"",""salerDoyCode"": 4216,""salerDoyName"":"
JS = JS + " ""ΣΤ ΘΕΣΣΑΛΟΝΙΚΗΣ"",""salerEmail"": ""info@demooe.gr"",""salerGemh"": ""000000000000"",""salerName"": ""ΜΠΡΑΤΣΑΣ ΑΘΑΝΑΣΙΟΣ"",""salerPhone"": ""2310989465"",""salerRepresentative"": null,""salerRepresentativeVat"": null,""salerStreetName"": ""ΣΕΜΕΛΗΣ 1"",""salerTitle"": ""BRATNET"",""salerTk"": ""54352"",""salerVat"": ""801670671"",""salerWebsite"": ""www.demooe.gr"",""signature"": ""304402201AFB5D9751F851D9EDAFF4E8F477BCD3B243AE78D8F4BC5BFD877692564C2CE202204B736425B163C155BCC1C14FA317A1AC2DE8E0F2AD4F580A76BE2BA6B2BE23DA"",""tipAmount"": 10,""transactionId"": ""x123456"",""vatExemptionCategoryName"": null,""vehicleNumber"": null}," & _
"""invoiceDetails"": [{" & _
"""TaricNo"": null,""code"": ""000005"",""deductionsAmount"": null,""dienergia"": null,""discountOption"": null,""expensesClassification"": null,""feesAmount"": null,""feesPercentCategory"": null,""fuelCode"": null,""incomeClassification"": [{" & _
"""amount"": 1450,""classificationCategory"": ""category1_3"",""classificationType"": ""E3_561_001"",""id"": 1}]," & _
"""invoiceDetailType"": null,""lineComments"": null,""lineNumber"": 1,""measurementUnit"": null,""measurementUnitName"": ""ΤΕΜ"",""movePurposeLine"": null,""name"": ""ΥΠΗΡΕΣΙΑ"",""netValue"": 1450,""netValueBeforeDiscount"": 1450,""notVat195"": null,""otherDeliveryNoteHeader"": null,""otherMeasurementUnitQuantity"": null,""otherMeasurementUnitTitle"": null,""otherMovePurposeLineTitle"": null,""otherTaxesAmount"": null,""otherTaxesPercentCategory"": null,""price"": 1450,""priceIncludeVAT"": 0,""quantity"": null,""quantity15"": null,""recType"": null,""stampDutyAmount"": null,""stampDutyPercentCategory"": null,""ublCpvCode"": null,""ublMeasurementUnit"": null,""ublVatCategory"": null,""vatAmount"": 348,""vatCategory"": 1,""vatExemptionCategory"": null,""vatPercent"": 24,""withheldAmount"": null,""withheldPercentCategory"": null}]," & _
"""invoiceHeader"": {""aa"": ""00120852"",""correlatedInvoices"": null,""currency"": ""EUR"",""dispatchDate"": null,""dispatchTime"": null,""exchangeRate"": null,""fuelInvoice"": null,""invoiceType"": ""2.1"",""invoiceVariationType"": null,""isDeliveryNote"": null,""issueDate"": ""{{date}}"",""issueTime"": ""{{time}}"",""movePurpose"": null,""multipleConnectedMarks"": null,""otherCorrelatedEntities"": null,""otherDeliveryNoteHeader"": null,""otherMovePurposeTitle"": null,""pInvoiceNote"": null,""reverseDeliveryNote"": null,""reverseDeliveryNotePurpose"": null,""selfPricing"": null,""series"": ""ΤΠΥ"",""specialInvoiceCategory"": null,""tableAA"": null,""thirdPartyCollection"": null,""toWeigh"": null,""totalCancelDeliveryOrders"": null,""vatPaymentSuspension"": null,""vehicleNumber"": null}," & _
"""invoiceSummary"": {""expensesClassification"": null,""incomeClassification"": [{" & _
"""amount"": 1450,""classificationCategory"": ""category1_3"",""classificationType"": ""E3_561_001"",""id"": 1}]," & _
"""totalDeductionsAmount"": 0,""totalFeesAmount"": 0,""totalGrossValue"": 1508,""totalNetValue"": 1450,""totalOtherTaxesAmount"": 0,""totalPrintGrossValue"": 1508,""totalStampDutyAmount"": 0,""totalVatAmount"": 348,""totalWithheldAmount"": 290}," & _
"""invoiceTaxVatAnalysis"": null," & _
"""invoiceVatAnalysis"": [{" & _
"""netValuePerVat"": 1450,""vatAmount"": 348,""vatCategory"": 1,""vatPercent"": 24}]," & _
"""isUnsigned"": true," & _
"""issuer"": {""address"": null,""branch"": 0,""country"": ""GR"",""countryDocumentId"": null,""documentIdNo"": null,""name"": null,""supplyAccountNo"": null,""vatNumber"": ""979708945""}," & _
"""packingsDeclarations"": null," & _
"""paymentMethods"": {""paymentMethodDetails"": [{" & _
"""amount"": 1508,""ecrToken"": null,""paymentMethodInfo"": null,""providersSignature"": null,""tid"": null,""tipAmount"": null,""transactionId"": null,""type"": 8}]}," & _
"""taxesDescriptions"": [{" & _
"""TaxCategoryId"": 3,""lineid"": 1,""taxAmount"": -290,""taxDescription"": ""Παρακράτηση Φόρου 20%"",""taxType"": 1,""taxVatAmount"": 0,""taxVatPercent"": 0,""vatCategory"": 7}]," & _
"""taxesTotals"": {""taxes"": [{" & _
"""id"": 1,""taxAmount"": 290,""taxCategory"": 3,""taxType"": 1,""ublMultiplierFactor"": null,""ublTaxCategory"": null,""underlyingValue"": 1450}]}," & _
"""tidNsp"": ""54888913""," & _
"""transmissionFailure"": null," & _
"""ublFields"": null" & _
"}]}"























    Dim http As Object
    Dim jsonBody As String
    Dim response As String
    Dim URL As String
    Dim username As String
    Dim password As String
    Dim credentials As String
    Dim base64Credentials As String

    ' Base URL - a?t??at?st?se µe t? p?a?µat??? s?? base URL
    Dim BASE_URL As String
    BASE_URL = "https://einvoicing-dev-api.etimologiera.gr/v4"

    URL = BASE_URL & "/sendSimInvoice"

    ' Basic Auth credentials
    username = "028783755"
    password = "Lagakis123!@#"

    ' S??d?asµ?? username:password ?a? ??d???p???s? se Base64
    credentials = username & ":" & password
    base64Credentials = EncodeBase64(credentials)

    ' JSON Body
    jsonBody = "{" & _
        """externalSystemId"": ""4053""," & _
        """issuerVatNumber"": ""028783755""," & _
        """invoiceIssueDate"": ""2026-03-23""," & _
        """invoiceIssueTime"": ""15:50:25""," & _
        """companyBranch"": 0," & _
        """invoiceType"": ""1.4""," & _
        """invoiceSeries"": ""\u03A4\u0394\u0391\u03A4\u03A7""," & _
        """netValue"": 20.0," & _
        """vatAmount"": 0.0," & _
        """totalValue"": 20.00," & _
        """paymentAmount"": 20.00," & _
        """terminalId"": ""54888913""," & _
        """nspCode"": ""01""" & _
    "}"



    ' ??µ??????a XMLHTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' ?????µa s??des?? - False = synchronous
    http.Open "POST", URL, False

    ' ???sµ?? headers
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Authorization", "Basic " & base64Credentials

    ' ?p?st??? request
    http.send JS 'onBody

    ' ??e???? response
    If http.Status = 200 Then
        response = http.responseText
        MsgBox "Success!" & vbCrLf & response, vbInformation, "Response"
    Else
        MsgBox "Error: " & http.Status & " - " & http.statusText, vbCritical, "HTTP Error"
    End If

    ' ?a?a??sµ??
    Set http = Nothing

End Sub


' -----------------------------------------------
' S????t?s? Base64 encoding ??a VB6
' -----------------------------------------------
Private Function EncodeBase64(Text As String) As String

    Dim arrData() As Byte
    Dim objxml As Object
    Dim objNode As Object

    arrData = StrConv(Text, vbFromUnicode)

    Set objxml = CreateObject("MSXML2.DOMDocument")
    Set objNode = objxml.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData

    EncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objxml = Nothing

End Function






















Function getBratToken() As String

        'παιρνω το τοκεν απο τον παροχο

        Dim URI As String
          'URL: https://einvoicing-dev-api.etimologiera.gr/v4/sendSimInvoice

100     URI = "https://einvoicing-dev-api.etimologiera.gr/v4/sendSimInvoice"   ' "https://" + PAROX_URLADD + "account.parochos.gr/api/account/loginToSubscription"

        Dim mEMAIL As String

102     mEMAIL = PAROX_USER  ' "glagakis@gmail.com" "glagakis@gmail.com"

        Dim REQ As WinHttp.WinHttpRequest

104     Set REQ = New WinHttp.WinHttpRequest

106     With REQ
108         .Open "POST", URI, async:=False
110         .setRequestHeader "Content-Type", "application/hal+json"
112         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"



              .setRequestHeader "Username", "028783755"
              .setRequestHeader "Password", "Lagakis123!@#"   '1ZZFXIz(bXmN^w_ruz!4H#EC9nHHZqz!"

            '   Note: Normally you don't include all of this whitespace, but
            '  we 'll use it in this example:
            Dim Q As String

            ' q = "{  ""subscriptionKey"": ""32657FD7929B483AABBD1C633401E945"" , ""username"": """ + mEMAIL + """ ," & """password"":""y39213921!!Y""," & """email"": """ + mEMAIL + """  " & "}"
114         Q = "{    }"
                              '""Username"": ""028783755"" ,""Password"":""1ZZFXIz(bXmN^w_ruz!4H#EC9nHHZqz!""
116          .send Q
118         MILSEC 1000
             MsgBox REQ.responseText
        End With

120     Dim sInputJson As String: sInputJson = REQ.responseText

122     Set p = JSON.parse(sInputJson)

        'MsgBox "Parsed object output: " & json.toString(p)

        'MsgBox "Get Bodystyle data: " & p.Item("jwt")

124     gfURL1 = p.Item("url1")

126     getBratToken = p.Item("jwt")  ' Split(Split(sInputJson, ":")(1), ",")(0)

End Function



' "https://www1.aade.gr/aadeapps5/oilbft-deddhe/api/v1/checkDeddhe/824048676"
'End Sub

Function getToken() As String

        'παιρνω το τοκεν απο τον παροχο

        Dim URI As String

100     URI = "https://" + PAROX_URLADD + "account.parochos.gr/api/account/loginToSubscription"

        Dim mEMAIL As String

102     mEMAIL = PAROX_USER  ' "glagakis@gmail.com" "glagakis@gmail.com"

        Dim REQ As WinHttp.WinHttpRequest

104     Set REQ = New WinHttp.WinHttpRequest

106     With REQ
108         .Open "POST", URI, async:=False
110         .setRequestHeader "Content-Type", "application/hal+json"
112         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"

            '   Note: Normally you don't include all of this whitespace, but
            '  we 'll use it in this example:
            Dim Q As String

            ' q = "{  ""subscriptionKey"": ""32657FD7929B483AABBD1C633401E945"" , ""username"": """ + mEMAIL + """ ," & """password"":""y39213921!!Y""," & """email"": """ + mEMAIL + """  " & "}"
114         Q = "{  ""subscriptionKey"": """ & PAROX_KEY & """ , ""username"": """ & mEMAIL & """ ," & """password"":""" & PAROX_PASS & """," & """email"": """ & mEMAIL & """  " & "}"
        
116         .send Q
118         MILSEC 1000
            ' MsgBox Req.responseText
        End With

120     Dim sInputJson As String: sInputJson = REQ.responseText

122     Set p = JSON.parse(sInputJson)

        'MsgBox "Parsed object output: " & json.toString(p)

        'MsgBox "Get Bodystyle data: " & p.Item("jwt")

124     gfURL1 = p.Item("url1")

126     getToken = p.Item("jwt")  ' Split(Split(sInputJson, ":")(1), ",")(0)

End Function

'Dim xhr As Object
'Dim uri  As String
'      uri = "https://beta-account.parochos.gr/api/account/loginToSubscription"
' Dim mEMAIL As String
'mEMAIL = "glagakis@gmail.com"
'
'
'
'Dim HttpReq As Object
'Set HttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
'HttpReq.setTimeouts 1000, 1000, 1000, 1000
'MousePointer = vbHourglass
'HttpReq.Open "POST", uri + "?subscriptionKey=32657FD7929B483AABBD1C633401E945,username=" + mEMAIL + ",password=y39213921!!Y,email=" + mEMAIL, False
'HttpReq.send
'MsgBox HttpReq.responseText 'getResponseHeader("BTC")
'MousePointer = vbDefault
'
'
'
'
'
''Dim REQ As WinHttpRequest
''Set REQ = Nothing
''Set REQ = New WinHttpRequest
''
''
''
''Dim sBody As String
''
'''With REQ
''   ' Dim sBody As String
''  sBody = " username=" + mEMAIL + "& emai=" + mEMAIL + "&password=y39213921!!Y "
''   ' sBody = "username=xxxxxxx&message=Hello%20Kind%20regards%2C%20xxx%20xxxx%20xxx%2E%20Tel%2E%200721000111&to=%2B254721000111&from=xxxx"
''
''    REQ.Open "POST", uri, False
''    REQ.setRequestHeader "Content-Type", "text/xml"
''     REQ.setRequestHeader "Cache-Control", "no-cache, no-store"
''      REQ.setRequestHeader "Connection", "close"
''
''
''
''    REQ.setRequestHeader "subscriptionKey", "32657FD7929B483AABBD1C633401E945"
''
''    REQ.send sBody
''    MsgBox REQ.responseText
''
''
''
''
'''End With
''
''      Exit Sub
'
''Set xhr = CreateObject("MSXML2.ServerXMLHTTP")
''xhr.Open "POST", uri, False
''
''xhr.setRequestHeader "Content-Type", "application/json"
''xhr.setRequestHeader "subscriptionKey", "32657FD7929B483AABBD1C633401E945"     '"Authorization", "Bearer " + accesstoken
''
'
''
''xhr.send "[ {""username"":""" + mEMAIL + """, ""email"":""glagakis@gmail.com"",""password"": ""y39213921!!Y"" }]"
''
'''  "[{""datetime"": ""2021-02-21 14:07:37"",""isRefund"": false," & _
''         """receiptNumber"": ""880090"",""amount"": 164.22}]"
''
''If xhr.Status = 200 Then
''    MsgBox xhr.responseText
''End If
''Exit Sub
'
'
'
'    '    Dim client = New HttpClient()
'    '    Dim queryString = HttpUtility.ParseQueryString(String.Empty)
'
'
'    '    Dim memail As String
'    '    Dim jsonContent As String = " {""username"":""" + memail + """, ""email"":""glagakis@gmail.com"",""password"": ""y39213921!!Y"",""subscriptionKey"":""32657FD7929B483AABBD1C633401E945"" }"
'
'    '    Dim content As New StringContent(jsonContent, Encoding.UTF8, "application/json")
'
'
'
'     '   Dim response = Await client.PostAsync(uri, content)
'    '   Dim result = Await response.Content.ReadAsStringAsync()
'   '     TextBox2.Text = Result.toString
'
'    '    Dim MF = "c:\txtfiles\apantPAROCHOS.xml"  'Inv" + Format(Now, "yyyyddMMHHmm") + ".xml"
'    '    FileOpen(1, MF, OpenMode.Output)
'    '    PrintLine(1, result.ToString)
'    '    FileClose (1)
'
'
'    '    Dim json As String = result.ToString
'    '    Dim ser As JObject = JObject.Parse(json)
'    '    Dim c As String = ""
'
'
'
''        fbearer = ser("jwt").toString
'
'
'
''--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
'134  '    url2 = UrlAADE + "SendInvoices" '/RequestDocs"
'
'       'initialize
'136     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
'138     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
'140     XMLServer.setTimeouts 5000, 60000, 10000, 10000
'
'        'force TLS 1.2
'142     XMLServer.Option(9) = 2048
'144     XMLServer.Option(6) = True
'
'146
'
' XMLServer.Open "POST", uri, False
'
'    XMLServer.Open "GET", url2 & "?mark=400000019698028", False
'     XMLServer.setRequestHeader "username", mEMAIL
'     XMLServer.setRequestHeader "email", mEMAIL '"glagakis2"
'
'XMLServer.setRequestHeader "password", "y39213921!!"
'
'XMLServer.setRequestHeader "subscriptionKey", "32657FD7929B483AABBD1C633401E945"
'
'
'
'   '    Dim jsonContent As String = " {""username"":""" + memail + """, ""email"":""glagakis@gmail.com"",""password"": ""y39213921!!Y"",""subscriptionKey"":""32657FD7929B483AABBD1C633401E945"" }"
'
'
'
'152     'XMLServer.send TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
'154    ' Debug.Print XMLServer.Status
'        '  debugPrint XMLServer.responseText
'156  v = XMLServer.responseText
'     '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
'

'End Function

Private Sub cmdParochos2_Click()

        Dim p          As Object
   
        Dim sInputJson As String

100     sInputJson = "{ width: '200', frame: false, height: 130, bodyStyle:'background-color: #ffffcc;',buttonAlign:'right', items: [{ xtype: 'form',  url: '/content.asp'},{ xtype: 'form2',  url: '/content2.asp'}] }"
   
102     MsgBox "Input JSON string: " & sInputJson
   
        ' sets p
104     Set p = JSON.parse(sInputJson)
   
106     MsgBox "Parsed object output: " & JSON.toString(p)
   
108     MsgBox "Get Bodystyle data: " & p.Item("bodyStyle")
   
110     MsgBox "Get Form Url data: " & p.Item("items").Item(1).Item("url")
   
112     p.Item("items").Item(1).ADD "ExtraItem", "Extra Data Value"
   
114     MsgBox "Parsed object output with added item: " & JSON.toString(p)
   
        Exit Sub

        '************************************************************************************************
        Dim oHttpReq As MSXML2.XMLHTTP40     '.XMLHTTP60

        Dim xdoc     As New DOMDocument

        Dim URL      As String, user As String, pass As String

        Dim xmldata  As String

116     URL = UrlAADE + "SendInvoices" ' "https://xxxxxxxx.xxx.xx:4445/xxxxx/xxxxxxxxxxxxx"
118     user = P_USER ' "my_user_name"
120     pass = P_KEY ' "my_password"

        '   On Error GoTo errlbl
    
122     Set oHttpReq = New MSXML2.XMLHTTP40 '    DOMDocument40 '.XMLHTTP60
    
124     oHttpReq.Open "POST", URL, True, user, pass
      
126     oHttpReq.setRequestHeader "Content-Type", "text/xml" '; charset = UTF-8"
    
128     If (xdoc.parseError.errorCode <> 0) Then

            Dim myErr

130         Set myErr = xdoc.parseError
132         MsgBox ("You have error " & myErr.reason)
        End If
       
134     xdoc.Load ("C:\" + F_TXTFILES + "\inv.xml")    'my xml file that needs to be sent
   
136     oHttpReq.send xdoc.XML

138     DoEvents

140     Do While True
142         DoEvents

144         If oHttpReq.readyState = 4 Then Exit Do
        Loop

146     If oHttpReq.Status = 200 Then
            ' ProcessResponse  'sub that gets the information from the xml response
        Else
148         MsgBox CStr(oHttpReq.Status)
            'this is the number 12019
            'I read somewhere that this is an "INVALID HANDLE STATE" error
        End If

End Sub

Private Sub cmdPROPLHROMH_Click()

    Dim C_SEIRA As String
    Dim SQLDTBRANCH As String
    Dim Mctypos
    Dim mID_NUM As String
    mID_NUM = Format(CLng(Now) * 24 * 60 + Hour(Now()) * 60 + Minute(Now) + Second(Now()), "##00000000")
    Dim MKAU, MFPA
     Dim TID     As String

                            TID = GGET_CVALUE("select ISNULL(C1,'') AS C1 from PINAKES where TYPOS=12 AND  PERIGRAFH LIKE '%POS%'") ' Mid(sqlDt.Rows(klm)("TRP"), 1, 1))
                            TID = Trim(TID) '.TrimEnd()
    
Dim c As String
      c = getToken()
    
    Mctypos = "8.4"
    
    

MKAU = 0: MFPA = 0
Dim POSO As String
gf_Bearer = c

POSO = InputBox("ΠΟΣΟ ; ", "ΠΟΣΟ ΠΡΟΕΙΣΠΡΑΞΗΣ")


    C_SEIRA = GET_CVALUE("SELECT SYNT_TITL FROM PARASTAT WHERE EIDOS='" + Left$(SQLDTATIM, 1) + "'")
   ' Mctypos = FINDTYPOS(mID(SQLDTATIM, 1, 1))

    On Error GoTo 0 'to_posERROR

    Dim uripos As String: uripos = gfURL1 + "api/requestPayment"

    Dim ANSW   As String
 
    Dim QQ     As String

    QQ = "{""externalSystemId"":""" + mID_NUM + """ " + ","
    QQ = QQ + " ""issuerVatNumber"":""" + afmCompany + """" + ","
    QQ = QQ + " ""invoiceIssueDate"":""" + Format(Now, "yyyy-MM-dd") + ""","

    QQ = QQ + " ""companyBranch"":""" + SQLDTBRANCH + ""","             'OK
    QQ = QQ + " ""invoiceType"":""" + "8.4" + ""","  'Split(Mctypos, ";")(0)
    QQ = QQ + " ""invoiceSeries"":""" + C_SEIRA + ""","
    QQ = QQ + " ""invoiceAA"":""" + mID(SQLDTATIM, 2, 6) + ""","
    'QQ = QQ + " ""netValue"":" + Replace(Format(MKAU, "#####0.00"), ",", ".") + ","
    
    
     QQ = QQ + " ""netValue"":" + Replace(POSO, ",", ".") + ","
    
    
    QQ = QQ + " ""VatAmount"":" + Replace(Format(MFPA, "#####0.00"), ",", ".") + ","
    QQ = QQ + " ""totalValue"":" + Replace(POSO, ",", ".") + ","

    'QQ = QQ + " ""paymentAmount"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
    
    QQ = QQ + " ""paymentAmount"":" + Replace(POSO, ",", ".") + ","
    
    QQ = QQ + " ""NspCode"":""2"" , "
    QQ = QQ + " ""terminalId"":""" + TID + """ }"     '"99999069"   ' ΟΚ

    Open "C:\TXTFILES\REQUES111TAPOPAROXO" + mID_NUM + ".TXT" For Output As #1
    Write #1, QQ
         
    Close #1

    ANSW = REQWINHTTP(QQ, uripos, LTrim(mID_NUM))

    Dim ser As Object
 
    Set ser = JSON.parse(ANSW)
        
  On Error Resume Next
  
  F_PAROX_SIGNATURE = ser.Item("paymentToken")("signature")
    posTimestamp = ser.Item("paymentToken")("timestamp")
    posuid = ser.Item("uid")
 
    Open "C:\TXTFILES\answREQUEST2222APOPAROXO" + extID + ".TXT" For Output As #1
    Write #1, ANSW
    Write #1, F_PAROX_SIGNATURE
    Close #1
 
    If Len(F_PAROX_SIGNATURE) < 5 Then
       ' to_pos = "error"

       ' Exit Function

    End If
 
    Dim timestampStart As String: timestampStart = """timestamp"":""" '-------------------------

    Dim timestampEnd   As String: timestampEnd = ""","

    Dim startIndex     As Integer: startIndex = InStr(posTimestamp, timestampStart)

    Dim endIndex       As Integer: endIndex = InStr(posTimestamp, timestampEnd)

    posTimestamp = mID(posTimestamp, 1, 33) ' startIndex + 1, endIndex - startIndex)























End Sub

Private Sub cmdRELATECLIENT_Click()

        Dim ANS    As String, TYPOS As String

        Dim ARKYKL As String

100     ANS = InputBox("2=PARKING 3=GARAGE", "")

102     ARKYKL = InputBox("αρ.κυκλοφοριας", "")

104     If Val(ANS) = 2 Then
106         TYPOS = "dcrnew:parkingcarwash"
        Else
108         TYPOS = "dcrnew:garage"
        End If

        Dim FF As String

110     'FF = RELATECLIEND("400000000000", "100000000013314", 1, "", "", "")

End Sub

Public Function OLD_RELATECLIEND(ByVal MARK As String, ByVal CORID As String, mark1fim2 As Integer, mfimaa As String, mfimdate As String, ByVal artam As String) As String
        'CORID =100000000007161#
        'MARK=4000000000
   
        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
100     Set docStock = New MSXML2.DOMDocument

102     With docStock
104         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
106         Set elemRoot = .createElement("dcrudtcor:ClientCorrelationDoc")
108         elemRoot.setAttribute "xmlns:dcrudtcor", "http://www.aade.gr/myDATA/dcrudtcor/v1.0"
110         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
112         elemRoot.setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/dcrudtcor/v1.0 SimpleTypes-v1.0.xsd clientCorrelationType-v1.0.xsd"

114         Set elemStock = docStock.createElement("dcrudtcor:clientCorrelation")

116         Set elem2Field = docStock.createElement("dcrudtcor:entityVatNumber"): elem2Field.Text = afmCompany: elemStock.appendChild elem2Field
                                         
118        ' Set ELEM2FIELD = docStock.createElement("dcrudtcor:mark"): ELEM2FIELD.Text = MARK: elemStock.appendChild ELEM2FIELD
            '' Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = "202502221033: elemStock.appendChild elem2Field"

          If mark1fim2 = 1 Then  ' me mark
              Set elem2Field = docStock.createElement("dcrudtcor:mark"): elem2Field.Text = MARK: elemStock.appendChild elem2Field
          Else
            Set elem2Field = docStock.createElement("dcrudtcor:FIM")
            

                    Set elem22field = docStock.createElement("dcrudtcor:FIMNumber"): elem22field.Text = artam: elem2Field.appendChild elem22field
                    Set elem22field = docStock.createElement("dcrudtcor:FIMAA"): elem22field.Text = mfimaa: elem2Field.appendChild elem22field
                    Set elem22field = docStock.createElement("dcrudtcor:FIMIssueDate"): elem22field.Text = mfimdate: elem2Field.appendChild elem22field
                    Set elem22field = docStock.createElement("dcrudtcor:FIMIssueTime"): elem22field.Text = "00:00:00": elem2Field.appendChild elem22field

              elemStock.appendChild elem2Field
           End If




















120         Set elem2Field = docStock.createElement("dcrudtcor:correlatedDCLids"): elem2Field.Text = CORID: elemStock.appendChild elem2Field
                       

                    
122         elemRoot.appendChild elemStock
  
            '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
124         Set .documentElement = elemRoot

            On Error Resume Next
            
126         FormatXmlDocument docStock ' βαζει κενα να ειναι ευκολο στο διαβασμα

            Dim ANS2 As Integer

128         ANS2 = MsgBox(docStock.XML, vbYesNo)

130         If ANS2 = vbNo Then

                Exit Function

            End If
            
132         Kill "C:\txtfiles\" + ARKYKL + "r.xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

134         .save "C:\txtfiles\" + ARKYKL + "r.xml"
        End With

        ' TO XML EINAI docStock.xml

        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
     
        ' Dim MARK As String
        '    MARK = InputBox("δωσε τo mark", "μαρκ")
     
136     url2 = url2 + MARKTIM
     
138     url2 = UrlAADE + "DCL/ClientCorrelations"
     url2 = UrlAADE + "DCL/ClientCorrelations"
        'initialize
140     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
142     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
144     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
146     XMLServer.Option(9) = 2048
148     XMLServer.Option(6) = True
    
150     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
152     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
154     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

156     XMLServer.send docStock.XML 'TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
158     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
160     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
 
        Dim CC As String

162     CC = takefromxml(v)
 
164     RELATECLIEND = CC '  <clientCorrelationID>100000000008565</clientCorrelationID>
166     CC = takefromxml(v)
        'CC = takefromxml(v)
 
168     MsgBox v

170     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

172     mfile = "c:\" + F_TXTFILES + "\sendinv\aytok" + "-" + ARKYKL + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
174     Debug.Print v
176     Open mfile For Output As #8
178     Print #8, v
180     Close #8
    
182     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
        '  READ_CANCEL v, MARK2

        '  400001829337674

        '    CANCEL_INVOICE = MARK2
    
        '<EhFooter>
        Exit Function

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.RELATE_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Private Sub cmdSendTest_Click()

        Dim oHttpReq As MSXML2.XMLHTTP40     '.XMLHTTP60

        Dim xdoc     As New DOMDocument

        Dim URL      As String, user As String, pass As String

        Dim xmldata  As String

        ''https://mydataapidev.aade.gr/DCL/SendClient
   
100     user = P_USER ' "my_user_name"
102     pass = P_KEY ' "my_password"
    
        '  user = "glagakis@gmail.com"
        '28783755
        'pass = "dfeaf73f6ba598a2ab9f25a5be216829"

        On Error GoTo 0 'errlbl

104     URL = UrlAADE + "DCL/SendClient" ' "https://xxxxxxxx.xxx.xx:4445/xxxxx/xxxxxxxxxxxxx"
106     Set oHttpReq = New MSXML2.XMLHTTP40 '    DOMDocument40 '.XMLHTTP60
    
108     oHttpReq.Open "POST", URL, True, user, pass
      
110     oHttpReq.setRequestHeader "Content-Type", "text/xml" '; charset = UTF-8"
    
112     If (xdoc.parseError.errorCode <> 0) Then

            Dim myErr

114         Set myErr = xdoc.parseError
116         MsgBox ("You have error " & myErr.reason)
        End If
       
118     xdoc.Load ("C:\" + F_TXTFILES + "\synerg.xml")    'my xml file that needs to be sent
   
120     oHttpReq.send xdoc.XML

122     DoEvents

124     Do While True
126         DoEvents

128         If oHttpReq.readyState = 4 Then Exit Do
        Loop

130     If oHttpReq.Status = 200 Then
            ' ProcessResponse  'sub that gets the information from the xml response
        Else
132         MsgBox CStr(oHttpReq.Status)
            'this is the number 12019
            'I read somewhere that this is an "INVALID HANDLE STATE" error
        End If

        Exit Sub

errlbl:
134     Debug.Print Err.Number
136     Debug.Print Err.Description

138     MsgBox "error---------"

        'Dim oHttpReq As MSXML2.XMLHTTP40     '.XMLHTTP60
        'Dim xdoc As New DOMDocument
        '    Dim url As String, user As String, pass As String
        '    Dim xmldata As String
        '
        '    url = UrlAADE + "SendInvoices" ' "https://xxxxxxxx.xxx.xx:4445/xxxxx/xxxxxxxxxxxxx"
        '    user = P_USER ' "my_user_name"
        '    pass = P_KEY ' "my_password"
        '
        '    On Error GoTo errlbl
        '
        '    Set oHttpReq = New MSXML2.XMLHTTP40 '    DOMDocument40 '.XMLHTTP60
        '
        '    oHttpReq.Open "POST", url, True, user, pass
        '
        '    oHttpReq.setRequestHeader "Content-Type", "text/xml" '; charset = UTF-8"
        '
        '
        '    If (xdoc.parseError.errorCode <> 0) Then
        '        Dim myErr
        '        Set myErr = xdoc.parseError
        '        MsgBox ("You have error " & myErr.reason)
        '    End If
        '
        '    xdoc.Load ("C:\" + F_TXTFILES + "\inv.xml")    'my xml file that needs to be sent
        '
        '    oHttpReq.send xdoc.XML
        '    DoEvents
        '
        '
        '    Do While True
        '        DoEvents
        '        If oHttpReq.readyState = 4 Then Exit Do
        '    Loop
        '
        '    If oHttpReq.Status = 200 Then
        '       ' ProcessResponse  'sub that gets the information from the xml response
        '    Else
        '        MsgBox CStr(oHttpReq.Status)
        '              'this is the number 12019
        '              'I read somewhere that this is an "INVALID HANDLE STATE" error
        '    End If
        '
        '    Exit Sub
        'errlbl:
        '    Debug.Print Err.Number
        '    Debug.Print Err.Description
        '
        'MsgBox "error---------"

End Sub

Private Sub cmddiagdokimi_Click()
100     Gdb.Execute "UPDATE TIM SET ENTITYMARK=NULL WHERE ENTITYMARK LIKE 'ΔΟΚ%'"
End Sub

Private Sub cmdMARK_Click()

        On Error Resume Next

        Dim ANS As Integer

        Dim k   As Integer

100     ANS = MsgBox(Adodc1.Recordset("atim") + " ** Προσοχη θα αφαιρεθεί το mark. Πρέπει να γίνει επαναποστολή. Στο Mydata εξακολουθεί να υπάρχει το υπάρχον mark." + Chr(13) + " Ειστε Σίγουρος/η; ", vbYesNo)

102     If ANS = vbYes Then
104         Gdb.Execute "update TIM set ENTITYMARK=NULL WHERE ID_NUM=" + str(Adodc1.Recordset("id_num")), k

106         If k = 1 Then
108             MsgBox "OK"
            End If

        End If

110     Adodc1.Refresh

End Sub

Private Sub cmdDeleteMark_Click()

100     deletemarK

End Sub

Sub deletemarK()

        Dim a, B

100     B = InputBox("Προσοχή θα ακυρωθεί μαρκ.Δωσε κωδικό αξουσιοδότησης")

102     If B = "3921" Then
104         a = InputBox("δωσε μαρκ που θα ακυρωθεί")
106         c = CANCEL_INVOICE(Trim(a), Trim(a))
108         MsgBox "ok"
        Else
110         MsgBox "λαθος"
        End If

End Sub

Private Sub Command10old_Click()

        '=============== request client ======================================
        Dim ANS    As String, TYPOS As String

        Dim ARKYKL As String

100     ANS = InputBox("2=PARKING 3=GARAGE", "")

102     ARKYKL = InputBox("αρ.κυκλοφοριας", "")

104     If Val(ANS) = 2 Then
106         TYPOS = "dcrnew:parkingcarwash"
        Else
108         TYPOS = "dcrnew:garage"
        End If
   
        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
110     Set docStock = New MSXML2.DOMDocument

112     With docStock
114         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
116         Set elemRoot = .createElement("dcrudtcor:ClientCorrelationDoc")
118         elemRoot.setAttribute "xmlns:dcrudtcor", "http://www.aade.gr/myDATA/dcrudtcor/v1.0"
120         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
122         elemRoot.setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/dcrudtcor/v1.0 SimpleTypes-v1.0.xsd clientCorrelationType-v1.0.xsd"

124         Set elemStock = docStock.createElement("dcrudtcor:clientCorrelation")

126         Set elem2Field = docStock.createElement("dcrudtcor:entityVatNumber"): elem2Field.Text = afmCompany: elemStock.appendChild elem2Field
                     
128         Set elem2Field = docStock.createElement("dcrudtcor:mark"): elem2Field.Text = "400000000000000": elemStock.appendChild elem2Field
            ' Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = "202502221033: elemStock.appendChild elem2Field"
130         Set elem2Field = docStock.createElement("dcrudtcor:correlatedDCLids"): elem2Field.Text = 100000000007161#: elemStock.appendChild elem2Field
                       

                    
132         elemRoot.appendChild elemStock
  
            '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
134         Set .documentElement = elemRoot

            On Error Resume Next
            
136         FormatXmlDocument docStock ' βαζει κενα να ειναι ευκολο στο διαβασμα

            Dim ANS2 As Integer

138         ANS2 = MsgBox(docStock.XML, vbYesNo)

140         If ANS2 = vbNo Then

                Exit Sub

            End If
            
142         Kill "C:\txtfiles\" + ARKYKL + "r.xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

144         .save "C:\txtfiles\" + ARKYKL + "r.xml"
        End With

        ' TO XML EINAI docStock.xml

        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
     
        Dim MARK As String

        '    MARK = InputBox("δωσε τo mark", "μαρκ")
     
146     url2 = url2 + MARKTIM
     
148     url2 = "https://mydataapidev.aade.gr/DCL/RequestClients?" ' 100000000007664"

        'initialize
150     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
152     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
154     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
156     XMLServer.Option(9) = 2048
158     XMLServer.Option(6) = True
    
160     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
162     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
164     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

166     XMLServer.send 'docStock.XML 'TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
168     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
170     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
 
172     MsgBox v

174     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

176     mfile = "c:\" + F_TXTFILES + "\sendinv\aytok" + "-" + ARKYKL + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
178     Debug.Print v
180     Open mfile For Output As #8
182     Print #8, v
184     Close #8
    
186     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
        '  READ_CANCEL v, MARK2

        '  400001829337674

        '    CANCEL_INVOICE = MARK2
    
        '<EhFooter>
        Exit Sub

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.RELATE_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

 Function FERE_TRANSACTIONID(mID_NUM As String) As String 'result@transcactionid
'---------------------------------------------------------------------
 Dim URI As String
Dim Q As String
Dim ANSW As String
Dim ctoken As String


 Dim userpos As String, PWDPOS As String
        userpos = GGET_CVALUE("SELECT ISNULL(USERPOS,'') AS U FROM MEM")
        PWDPOS = GGET_CVALUE("SELECT ISNULL(PWDPOS,'') AS U FROM MEM")
 Q = " { ""identity"": """ + userpos + """,    ""password"": """ + PWDPOS + """ }"
        
        
    ctoken = request_pliromis(userpos, PWDPOS)
        
        
        
        
        
        CC = "0"
'
'      If Val(ccc) > 0 Then
'            uripos = "https://lambdapi.gr/backend/webecr/transaction_intents?timol=" + ccc
'        Else
'            uripos = "https://lambdapi.gr/backend/webecr/transaction_intents"   '?timol=" + ccc
'        End If
'
'If Val(mID_NUM) > 0 Then
'            uripos = "https://lambdapi.gr/backend/webecr/transaction_intents?timol=" + ccc
'        Else
'            uripos = "https://lambdapi.gr/backend/webecr/transaction_intents"   '?timol=" + ccc
'        End If

If Val(mID_NUM) > 0 Then
            uripos = "https://lambdapi.gr/backend/webecr/transaction_intents?timol=" + mID_NUM
        Else
            uripos = "https://lambdapi.gr/backend/webecr/transaction_intents"   '?timol=" + ccc
        End If






Dim Reqpos As WinHttp.WinHttpRequest
       Set Reqpos = New WinHttp.WinHttpRequest
        
        With Reqpos
            .Open "GET", uripos, async:=False
            .setRequestHeader "Content-Type", "application/json"
            .setRequestHeader "Accept", "*/*"
            .setRequestHeader "Authorization", "Bearer " + ctoken
           
            .send
            MsgBox ("ΑΠΑΝΤΗΣΗ ΑΠΟ ΠΛΗΡΩΜΗ" + Chr(13) + Reqpos.responseText)
        End With
        MILSEC (1000)
    ANSW = Reqpos.responseText
        
        Dim ser As Object
 Set ser = JSON.parse(ANSW)
        'Set ser = JSON.parse(ANSW)
        
F_posSignature = ser.Item(1)("TransactionId")   'ser.Item(0)("TransactionId") '("signature")
'posTimestamp = ser.Item("paymentToken")("timestamp")
        
        
        
        
        
        
        
        
        
        
        
FERE_TRANSACTIONID = F_posSignature 'ser.Item("TransactionId") '("signature")
        
        
        
        
        
        
        
        
        
 End Function

Private Sub cmdTEST2_Click()

Dim URI As String
Dim Q As String
Dim ANSW As String
Dim ctoken As String


 Dim userpos As String, PWDPOS As String
        userpos = GGET_CVALUE("SELECT ISNULL(USERPOS,'') AS U FROM MEM")
        PWDPOS = GGET_CVALUE("SELECT ISNULL(PWDPOS,'') AS U FROM MEM")
 Q = " { ""identity"": """ + userpos + """,    ""password"": """ + PWDPOS + """ }"
        
        
    ctoken = request_pliromis(userpos, PWDPOS)
        
        
        
        
        
        CC = "0"
        
      If Val(ccc) > 0 Then
            uripos = "https://lambdapi.gr/backend/webecr/transaction_intents?timol=" + ccc
        Else
            uripos = "https://lambdapi.gr/backend/webecr/transaction_intents"   '?timol=" + ccc
        End If




Dim Reqpos As WinHttp.WinHttpRequest
       Set Reqpos = New WinHttp.WinHttpRequest
        
        With Reqpos
            .Open "GET", uripos, async:=False
            .setRequestHeader "Content-Type", "application/json"
            .setRequestHeader "Accept", "*/*"
            .setRequestHeader "Authorization", "Bearer " + ctoken
           
            .send
            MsgBox ("ΑΠΑΝΤΗΣΗ ΑΠΟ ΠΛΗΡΩΜΗ" + Chr(13) + Reqpos.responseText)
        End With
        MILSEC (1000)
    
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
     Exit Sub
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
  URI = "https://lambdapi.gr/api/collections/users/auth-with-password"
  
 ' Dim Reqpos As WinHttp.WinHttpRequest
       Set Reqpos = New WinHttp.WinHttpRequest
  
  Reqpos.Open "POST", URI, async:=False
  Reqpos.setRequestHeader "Content-Type", "application/json"
  Reqpos.setRequestHeader "Accept", "*/*"
  Reqpos.send Q
  MILSEC 500
  
  MsgBox Reqpos.responseText
  ANSW = Reqpos.responseText
  Dim ser As Object
 Set ser = JSON.parse(ANSW)
        
F_posSignature = ser.Item("token") '("signature")
  
  
  
  
  
  
  
  


Exit Sub



  URI = "https://" + PAROX_URLADD + "account.parochos.gr/api/account/loginToSubscription"
  
  mEMAIL = "glagakis@gmail.com"
  Q = "{  ""subscriptionKey"": """ & PAROX_KEY & """ , ""username"": """ & mEMAIL & """ ," & """password"":""" & PAROX_PASS & """," & """email"": """ & mEMAIL & """  " & "}"
 
' Dim uripos As String: uripos = gfURL1 + "api/requestPayment"
ANSW = REQWINHTTP(Q, URI, id_num)
MsgBox ANSW









End Sub


Function REQWINHTTP(Q As String, URI As String, id_num) As String

'ta parakatv einai to post poy kanv gia to timologio

'136     With REQ
'138         .Open "POST", URI, async:=False
'140         .setRequestHeader "Content-Type", "application/hal+json"
'142         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
'144         .setRequestHeader "Authorization", "Bearer " + gf_Bearer
'            'Note: Normally you don't include all of this whitespace, but
'            'we'll use it in this example:
'146         Q = " { ""Bearer"":""" + gf_Bearer + """, ""source"":" & ResultJSON & ", ""identifier"":""Parochos"",    ""externalSystemId"": """ + str(id_num) + """ }"
'            ' q = " {   ""externalSystemId"": ""188"", ""source"":" & ResultJSON & "    }"
'148         .send (Q)
'150         MILSEC 1000
'        End With


        'Dim URI As String

100     'URI = "https://" + PAROX_URLADD + "account.parochos.gr/api/account/loginToSubscription"

        Dim mEMAIL As String

102     mEMAIL = PAROX_USER ' "glagakis@gmail.com"

        '
        On Error GoTo 0

        Dim REQ As WinHttp.WinHttpRequest

104    Set REQ = New WinHttp.WinHttpRequest

        'Dim Q As String
        ' Q = " { ""Bearer"":""" + gf_Bearer + """, ""source"":" & Q & ", ""identifier"":""Parochos"",    ""externalSystemId"": """ + str(id_num) + """ }"
106     With REQ
108         .Open "POST", URI, async:=False
110         .setRequestHeader "Content-Type", "application/hal+json"
112         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
             .setRequestHeader "Authorization", "Bearer " + gf_Bearer
             .setRequestHeader "X-Version", "3"
            'Note: Normally you don't include all of this whitespace, but
            'we'll use it in this example:
            ' q = "{  ""subscriptionKey"": ""32657FD7929B483AABBD1C633401E945"" , ""username"": """ & mEMAIL & """ ," & """password"":""y39213921!!Y""," & """email"": """ & mEMAIL & """  " & "}"
114     '    Q = "{  ""subscriptionKey"": """ & PAROX_KEY & """ , ""username"": """ & mEMAIL & """ ," & """password"":""" & PAROX_PASS & """," & """email"": """ & mEMAIL & """  " & "}"
116         .send (Q)
118       '  MsgBox (REQ.responseText)
        
        End With
        
        Dim CC   As String

        'CC = json.toString(q)

120     Dim JSON As String: JSON = REQ.responseText
        
'122     REQPAROCHOS = Split(Split(JSON, ":")(1), ",")(0)
'
'124     gfURL1 = Split(Split(JSON, ":")(3), ",")(0)
        
        '
        '       ' Dim ser As JObje


         REQWINHTTP = JSON



End Function





Private Sub Command10_Click()

        '=============== request client ======================================
        Dim ANS    As String, TYPOS As String

        Dim ARXID As String

100     ANS = InputBox("2=PARKING 3=GARAGE", "")

102     ARXID = InputBox("ID", , "100000000000000")

104     If Val(ANS) = 2 Then
106         TYPOS = "dcrnew:parkingcarwash"
        Else
108         TYPOS = "dcrnew:garage"
        End If
   
        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
        ''102     Set docStock = New MSXML2.DOMDocument
        ''104     With docStock
        ''106         .appendChild .createProcessingInstruction("xml", _
        ''                                                      "version=""1.0"" encoding=""utf-8""")
        ''108         Set elemRoot = .createElement("dcrudtcor:ClientCorrelationDoc")
        ''                elemRoot.setAttribute "xmlns:dcrudtcor", "http://www.aade.gr/myDATA/dcrudtcor/v1.0"
        ''                elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
        ''                elemRoot.setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/dcrudtcor/v1.0 SimpleTypes-v1.0.xsd clientCorrelationType-v1.0.xsd"
        ''
        ''116                 Set elemStock = docStock.createElement("dcrudtcor:clientCorrelation")
        ''
        ''                       Set elem2Field = docStock.createElement("dcrudtcor:entityVatNumber"): elem2Field.Text = "028783755": elemStock.appendChild elem2Field
        ''
        ''                      Set elem2Field = docStock.createElement("dcrudtcor:mark"): elem2Field.Text = "400000000000000": elemStock.appendChild elem2Field
        ''                      ' Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = "202502221033: elemStock.appendChild elem2Field"
        ''                        Set elem2Field = docStock.createElement("dcrudtcor:correlatedDCLids"): elem2Field.Text = 100000000007161#: elemStock.appendChild elem2Field
        ''
        ''126
        ''
        ''130                 elemRoot.appendChild elemStock
        ''
        ''          '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
        ''132         Set .documentElement = elemRoot
        ''            On Error Resume Next
        ''
        ''             FormatXmlDocument docStock ' ίa?e? ?e?a ?a e??a? e????? st? d?aίasµa
        ''            Dim ANS2 As Integer
        ''            ANS2 = MsgBox(docStock.XML, vbYesNo)
        ''            If ANS2 = vbNo Then
        ''                Exit Sub
        ''            End If
        ''
        ''
        ''134         Kill "C:\txtfiles\" + ARKYKL + "r.xml"
        ''            On Error GoTo cmdDOMDocumentUTF8_Click_Err
        ''136         .save "C:\txtfiles\" + ARKYKL + "r.xml"
        ''        End With
        ''
        ''

        ' TO XML EINAI docStock.xml

        '--------------- ???S???? INVOICES -----------------------------------------------------
     
        Dim MARK      As String

        '    MARK = InputBox("d?se to mark", "µa??")
     
110     url2 = url2 + MARKTIM
     
112     url2 = "https://mydataapidev.aade.gr/DCL/RequestClients?DCLID=" + ARXID '100000000007663" ''''''100000000007739"


      If F_demo = 1 Then
         ' url2 = "https://mydataapidev.aade.gr/DCL/SendClient"
            url2 = "https://mydataapidev.aade.gr/DCL/RequestClients?DCLID=" + ARXID '100000000007663" ''''''100000000007739"

      Else
      
             ' url2 = "https://mydatapi.aade.gr/DCL/SendClient"
             url2 = "https://mydatapi.aade.gr/DCL/RequestClients?DCLID=" + ARXID '100000000007663" ''''''100000000007739"
        
      End If





        'initialize  100000000007664
114     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
116     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
118     XMLServer.setTimeouts 5000, 60000, 10000, 10000

        '130     XMLServer.Open "GET", url2 & "?mark=" + MARK2, False
    
        'force TLS 1.2
120     XMLServer.Option(9) = 2048
122     XMLServer.Option(6) = True
    
124     XMLServer.Open "GET", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
126     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
128     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

130   On Error Resume Next
        XMLServer.send 'docStock.XML 'TXML 'strData  'S????O ?? STRING TXML STO MYDATA  'XMLServer.send TXML '  ?? T??O ?? ?? S????O ??? ?????? ???? ???O strData
132     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
134     v = XMLServer.responseText
        '--------------- ???S???? INVOICES -----------------------------------------------------
 
136     MsgBox v

138     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

140     mfile = "c:\" + F_TXTFILES + "\sendinv\OLAaytok" + "-" + ARXID + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
142     Debug.Print v
144     Open mfile For Output As #8
146     Print #8, v
148     Close #8
    
150     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
        ' -------------- ΛΗΦΘΕΙΣΑ ΑΠΑΝΤΗΣΗ ------------------https://mydataapidev.aade.gr/DCL/RequestClients?DCLID=100000000007664"
        '<?xml version="1.0" encoding="utf-8"?>
        '<RequestedDoc xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="http://www.aade.gr/myDATA/dcr/v1.0">
        '  <entityVatNumber>028783755</entityVatNumber>
        '  <clientsDoc>
        '    <DigitalClient>
        '      <InitialClientData>
        '        <idDcl>100000000007739</idDcl>
        '        <clientServiceType>3</clientServiceType>
        '        <creationDateTime xsi:nil="true" />
        '        <entityVatNumber>028783755</entityVatNumber>
        '        <branch>0</branch>
        '        <useCase>
        '          <garage>
        '            <vehicleRegistrationNumber>RRR8787</vehicleRegistrationNumber>
        '            <vehicleCategory>AYTOK</vehicleCategory>
        '            <vehicleFactory>CITROEN</vehicleFactory>
        '          </garage>
        '        </useCase>
        '      </InitialClientData>
        '    </DigitalClient>
        '  </clientsDoc>
        '  <updateclientRequestsDoc />
        '  <clientcorrelationsRequestsDoc />
        '  <cancelClientRequestsDoc />
        '</RequestedDoc>

        '===================================== https://mydataapidev.aade.gr/DCL/RequestClients?DCLID=100000000007663"
        '<?xml version="1.0" encoding="utf-8"?>
        '<RequestedDoc xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="http://www.aade.gr/myDATA/dcr/v1.0">
        '  <entityVatNumber>028783755</entityVatNumber>
        '  <clientsDoc>
        '    <DigitalClient>
        '      <InitialClientData>
        '        <idDcl>100000000007664</idDcl>
        '        <clientServiceType>3</clientServiceType>
        '        <creationDateTime xsi:nil="true" />
        '        <entityVatNumber>028783755</entityVatNumber>
        '        <branch>0</branch>
        '        <useCase>
        '          <garage>
        '            <vehicleRegistrationNumber>RMK5555</vehicleRegistrationNumber>
        '            <vehicleCategory>FORTHGAKI</vehicleCategory>
        '            <vehicleFactory>FIAT</vehicleFactory>
        '          </garage>
        '        </useCase>
        '      </InitialClientData>
        '    </DigitalClient>
        '    <DigitalClient>
        '      <InitialClientData>
        '        <idDcl>100000000007739</idDcl>
        '        <clientServiceType>3</clientServiceType>
        '        <creationDateTime xsi:nil="true" />
        '        <entityVatNumber>028783755</entityVatNumber>
        '        <branch>0</branch>
        '        <useCase>
        '          <garage>
        '            <vehicleRegistrationNumber>RRR8787</vehicleRegistrationNumber>
        '            <vehicleCategory>AYTOK</vehicleCategory>
        '            <vehicleFactory>CITROEN</vehicleFactory>
        '          </garage>
        '        </useCase>
        '      </InitialClientData>
        '    </DigitalClient>
        '  </clientsDoc>
        '  <updateclientRequestsDoc />
        '  <clientcorrelationsRequestsDoc />
        '  <cancelClientRequestsDoc />
        '</RequestedDoc>
    
        '<EhFooter>
        Exit Sub

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.RELATE_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command11_Click()

        Dim N As Long

100     N = Adodc1.Recordset("ID_NUM")
102     XML.TextRTF = GET_CVALUE("SELECT XMLTEXT FROM TIM WHERE ID_NUM=" + str(N))

End Sub

Private Sub Command12_Click()
        'ΣΒΗΝΕΙ ΤΙΣ ΔΙΠΛΟΕΓΓΡΑΦΕΣ ΠΟΥ ΔΗΛΩΘΗΚΑΝ ΣΤΟ ΚΕΝΤΡΙΚΟ ΕΝΩ ΗΤΑΝ ΥΠΟΚΑΤΑΣΤΗΜΑ

        Dim a, B

100     B = InputBox("Προσοχή θα ακυρωθOYN ΑΠΟ TABLE APESTALMENA2 .Δωσε κωδικό αξουσιοδότησης")

102     If B = "3921" Then
            ' a = InputBox("δωσε μαρκ που θα ακυρωθεί")
            ' C = CANCEL_INVOICE(Trim(a), Trim(a))
            ' MsgBox "ok"
        Else
104         MsgBox "λαθος"
        End If

        On Error Resume Next

106     Gdb.Execute "ALTER TABLE  APESTALMENA2 ADD CANCELATIONMARK NVARCHAR(20)"

        On Error GoTo 0
 
        Dim c, D, SX As String

        Dim R As New ADODB.Recordset
 
108     R.Open "SELECT * FROM  APESTALMENA2  WHERE ISSUEDATE>='" + Format(APO.Value, "yyyy-MM-dd") + "'  AND ISSUEDATE<'" + Format(DateAdd("d", 1, EOS.Value), "yyyy-MM-dd") + "' ORDER BY MARK", Gdb, adOpenDynamic, adLockOptimistic
 
110     Do While Not R.EOF
 
            'd = InputBox("δωσε τον αριθμό του id", 0)
   
            ' SX = GET_CVALUE("select ISNULL(SXETMARK,'00') AS SXX from TIM WHERE ID_NUM=" + d)
            'If SX = "00" Then
            '  MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΜΑΡΚ ΠΟΥ ΘΑ ΑΚΥΡΩΘΕΙ"
            ' Exit Sub
            ' End If
            Dim TIMOL As String

            ' TIMOL = R!ATIM    ' AA   '    GET_CVALUE("select ATIM from TIM WHERE ID_NUM=" + d)
            '  SX = R!entitymark   '   R!MARK
            'List1.AddItem
   
112         TIMOL = R!AA   '    GET_CVALUE("select ATIM from TIM WHERE ID_NUM=" + d)
114         SX = R!MARK
   
116         c = CANCEL_INVOICE(Trim(SX), TIMOL)

118         If Len(Trim(c)) > 5 Then
120             Gdb.Execute "update APESTALMENA2 set CANCELATIONMARK='" + c + "' WHERE ID=" + str(R!ID)
                'Gdb.Execute "update DOKTIMCANC set SXETMARK='" + C + "' WHERE ID_NUM=" + str(R!ID_NUM)
122             List1.AddItem SX + " AK " + c
            End If   '
   
124         R.MoveNext
   
        Loop
   
126     MsgBox "OK. ΤΕΛΑΕΥΤΑΊΟ ΑΚΥΡΩΜΕΝΟ ΜΑΡΚ  " + SX
   
End Sub

Private Sub Command13_Click()
Dim CC As String
CC = InputBox("id")

'cc = "40576"

Dim cc3 As String
uripos = "api/getPayments/" + CC  ' {externalSystemId}"
'ANSW = REQWINHTTP(QQ, uripos, LTrim(mID_NUM))
cc3 = REQPAROCHOS()
 'XMLServer.Open "GET", url2 & "" + mmark,

     Dim REQ As WinHttp.WinHttpRequest

    Set REQ = New WinHttp.WinHttpRequest
106     With REQ
108

' .Open("GET", uripos, Async:=False)
'            .SetRequestHeader("Content-Type", "application/json")
'            .SetRequestHeader("Accept", "*/*")
'            .SetRequestHeader("Authorization", "Bearer " + ctoken)
'            ' q = offline
'            .Send()







            .Open "GET", gfURL1 + uripos, async:=False
110         .setRequestHeader "Content-Type", "application/json"
112         .setRequestHeader "Accept", "*/*"
             .setRequestHeader "Authorization", "Bearer " + gf_Bearer
             .setRequestHeader "X-Version", "3"
             .send ""
118         MsgBox (REQ.responseText)
        
        End With
        
Open "C:\TXTFILES\PAYMENTS" + CC + ".TXT" For Output As #1
Write #1, REQ.responseText
Close #1



















'ANSW = REQWINHTTP(QQ, uripos, LTrim(mID_NUM))




End Sub

Private Sub Command14_Click()
'https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=NziZ2b5egY41ooVijo%2fjtSMRj7CvuLB%2fxG1YnEyrsSdlsH7ef9ZcgJZCLIFhuHMUa6Cp8pQjpI4ErLYp4fpsmjaTE7asa34Phn2SIB5cM4c%3d
''https://mydataapidev.aade.gr/GetDeliveryNoteStatus?mark=%7bmark%7d
Dim MARK As String
        MARK = InputBox("δωσε τo mark", "μαρκ")

Dim DUM As String
DUM = GetDeliveryStatus(MARK, "aaa")


End Sub
Public Function GetDeliveryStatus(ByVal MARKTIM As String, TIMOL As String) As String
loadpar7
        '<EhHeader>
        On Error GoTo CANCEL_INVOICE_Err

        '</EhHeader>
100     If Len(Trim(MARKTIM)) < 1 Then
102         GetDeliveryStatus = ""

            Exit Function

        End If

        '            Dim uri As String: uri = "https://mydata-dev.azure-api.net/CancelInvoice?mark=400000020235194"

        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
104     url2 = UrlAADE + "GetDeliveryNoteStatus?mark=" '/RequestDocs"
     
        Dim MARK As String

        '    MARK = InputBox("δωσε τo mark", "μαρκ")
     
106     url2 = url2 + MARKTIM

        'initialize
108     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
110     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
112     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
114     XMLServer.Option(9) = 2048
116     XMLServer.Option(6) = True
    
118     'XMLServer.Open "POST", url2, False
        XMLServer.Open "GET", url2, False
120     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
122     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
124     XMLServer.send 'TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
126     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
128     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
MsgBox v
Dim c55 As String
Dim c6 As String
c6 = ""
c55 = READ_xml(v, "GetDeliveryNoteStatusResponse", "status", c6)


130     TIMOL = Replace(TIMOL, "/", "-")

132     Dim mfile As String: mfile = "c:\" + F_TXTFILES + "\sendinv\apantdeliv-" + TIMOL + "-" + MARKTIM + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"

134     Debug.Print v
136     Open mfile For Output As #8
138     Print #8, v
140     Close #8
    
142     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
144    ' READ_CANCEL v, MARK2

        '  400001829337674

146     GetDeliveryStatus = c55

        '<EhFooter>
        Exit Function

CANCEL_INVOICE_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.CANCEL_INVOICE " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>
End Function

Private Sub CMDLIXIAPOSTOLIS_Click()
Dim qrl As String: qrl = "https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=vZGdPC4qtv61lfj3tMcBWCVzle6HNmunsN6JjmX4e9gLqhLIIIUNGzIx6YzWdXmeg0fsTOtxriHc%2bCRo3shXRfuFVdEayG8L9Xt9%2fSH8HaE%3d"
'"https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=NziZ2b5egY41ooVijo%2fjtSMRj7CvuLB%2fxG1YnEyrsSdlsH7ef9ZcgJZCLIFhuHMUa6Cp8pQjpI4ErLYp4fpsmjaTE7asa34Phn2SIB5cM4c%3d"
'https://mydataapidev.aade.gr/RegisterTransfer

'https://mydatapi.aade.gr/myDATA/ConfirmDeliveryOutcome
Dim TXML As String
TXML = "<?xml version=""1.0"" encoding=""utf-8""?>"   'qrUrl,outcome"
TXML = TXML + "<ConfirmDeliveryOutcomeRequest>"
TXML = TXML + "<qrUrl>" + qrl + "</qrUrl>"
TXML = TXML + "<outcome>FULL</outcome>"
TXML = TXML + "</ConfirmDeliveryOutcomeRequest>"

Dim DUM As String: DUM = SendConfirm(TXML, " ")

End Sub
Public Function SendConfirm(ByVal TXML As String, TIMOL As String) As String
        loadpar7
        '<EhHeader>
        On Error GoTo CANCEL_INVOICE_Err

        '</EhHeader>
100     If Len(Trim(TXML)) < 1 Then
102         SendConfirm = ""

            Exit Function

        End If

        '           ' Dim uri As String: uri = "https://mydata-dev.azure-api.net/CancelInvoice?mark=400000020235194"

        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
104     url2 = UrlAADE + "ConfirmDeliveryOutcome"
     
        Dim MARK As String

        '    MARK = InputBox("δωσε τo mark", "μαρκ")
     
106    ' url2 = url2 + MARKTIM

        'initialize
108     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
110     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
112     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
114     XMLServer.Option(9) = 2048
116     XMLServer.Option(6) = True
    
118     XMLServer.Open "POST", url2, False
       ' XMLServer.Open "GET", url2, False
120     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
122     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
124     XMLServer.send TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
126     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
128     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
MsgBox v
'Cannot confirm delivery outcome for Invoice with MARK: 400001961787310. It has not been dispatched yet. Current status: 0</message>
 '       <code>813</code>

130     TIMOL = Replace(TIMOL, "/", "-")

132     Dim mfile As String: mfile = "c:\" + F_TXTFILES + "\sendinv\apantdeliv-" + TIMOL + "-" + MARKTIM + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"

134     Debug.Print v
136     Open mfile For Output As #8
138     Print #8, v
140     Close #8
    
142     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
144     READ_CANCEL v, MARK2

        '  400001829337674

146     SendConfirm = MARK2

        '<EhFooter>
        Exit Function

CANCEL_INVOICE_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.CANCEL_INVOICE " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>
End Function



Private Sub ENARXI2APOSTOLIS(ByVal carrierAFM As String, ByVal qrl As String, ByVal ark As String)

'Dim qrl As String: qrl = "https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=vZGdPC4qtv61lfj3tMcBWCVzle6HNmunsN6JjmX4e9gLqhLIIIUNGzIx6YzWdXmeg0fsTOtxriHc%2bCRo3shXRfuFVdEayG8L9Xt9%2fSH8HaE%3d"

'https://mydatapi.aade.gr/myDATA/ConfirmDeliveryOutcome
Dim TXML As String
TXML = "<?xml version=""1.0"" encoding=""utf-8""?>"   'qrUrl,outcome"
TXML = TXML + "<Transport>"
TXML = TXML + "<qrUrl>" + qrl + "</qrUrl>"

TXML = TXML + "<transportDetail>"
'                                         Dim ark As String: ark = "PMT4231"

TXML = TXML + "<vehicleNumber>" + ark + "</vehicleNumber>"

Dim TYPOS As String: TYPOS = "7"
TXML = TXML + "<transportType>" + TYPOS + "</transportType>"

TXML = TXML + "<carrierVatNumber>" + carrierAFM + "</carrierVatNumber>"

Dim LoN, LAT As String
LoN = "24.12468"
LAT = "41.136"

'txml = txml + "<longitude>" + LoN + "</longitude>"
'txml = txml + "<latitude>" + LAT + "</latitude>"

'Κωδικός Περιγραφή 1 Φορτηγό Δημόσιας Χρήσης 2 Φορτηγό Ιδιωτικής Χρήσης
'3 Πλοίο 4 Τρένο 5 Αεροπλάνο 6 Λοιπά Μεταφορικά Μέσα (π.χ Δίκυκλα, ..) 7 Άνευ


TXML = TXML + "</transportDetail>"

TXML = TXML + "</Transport>"

Dim DUM As String: DUM = ENARXIAPOS(TXML, "00")

End Sub



Private Sub CMDENARXIAPOSTOLIS_Click()

Dim qrl As String: qrl = "https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=vZGdPC4qtv61lfj3tMcBWCVzle6HNmunsN6JjmX4e9gLqhLIIIUNGzIx6YzWdXmeg0fsTOtxriHc%2bCRo3shXRfuFVdEayG8L9Xt9%2fSH8HaE%3d"
'"https://mydataapidev.aade.gr/TimologioQR/QRInfo?q=NziZ2b5egY41ooVijo%2fjtSMRj7CvuLB%2fxG1YnEyrsSdlsH7ef9ZcgJZCLIFhuHMUa6Cp8pQjpI4ErLYp4fpsmjaTE7asa34Phn2SIB5cM4c%3d"
'https://mydataapidev.aade.gr/RegisterTransfer

'https://mydatapi.aade.gr/myDATA/ConfirmDeliveryOutcome
Dim TXML As String
TXML = "<?xml version=""1.0"" encoding=""utf-8""?>"   'qrUrl,outcome"
TXML = TXML + "<Transport>"
TXML = TXML + "<qrUrl>" + qrl + "</qrUrl>"

TXML = TXML + "<transportDetail>"
Dim ark As String: ark = "PMT4231"

TXML = TXML + "<vehicleNumber>" + ark + "</vehicleNumber>"

Dim TYPOS As String: TYPOS = "7"
TXML = TXML + "<transportType>" + TYPOS + "</transportType>"

TXML = TXML + "<carrierVatNumber>150144455</carrierVatNumber>"

Dim LoN, LAT As String
LoN = "24.12468"
LAT = "41.136"

'txml = txml + "<longitude>" + LoN + "</longitude>"
'txml = txml + "<latitude>" + LAT + "</latitude>"

'Κωδικός Περιγραφή 1 Φορτηγό Δημόσιας Χρήσης 2 Φορτηγό Ιδιωτικής Χρήσης
'3 Πλοίο 4 Τρένο 5 Αεροπλάνο 6 Λοιπά Μεταφορικά Μέσα (π.χ Δίκυκλα, ..) 7 Άνευ


TXML = TXML + "</transportDetail>"

TXML = TXML + "</Transport>"

Dim DUM As String: DUM = ENARXIAPOS(TXML, " ")

End Sub
Public Function ENARXIAPOS(ByVal TXML As String, TIMOL As String) As String
      
      If TIMOL = "00" Then
      Else
         Par7MyData.loadpar7
      End If
        '<EhHeader>
        On Error GoTo CANCEL_INVOICE_Err

        '</EhHeader>
100     If Len(Trim(TXML)) < 1 Then
102         ENARXIAPOS = ""

            Exit Function

        End If

        '           ' Dim uri As String: uri = "https://mydata-dev.azure-api.net/CancelInvoice?mark=400000020235194"

        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
104     url2 = UrlAADE + "RegisterTransfer"
     
        Dim MARK As String

        '    MARK = InputBox("δωσε τo mark", "μαρκ")
     
106    ' url2 = url2 + MARKTIM

        'initialize
108     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
110     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
112     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
114     XMLServer.Option(9) = 2048
116     XMLServer.Option(6) = True
    
118     XMLServer.Open "POST", url2, False
       ' XMLServer.Open "GET", url2, False
120     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
122     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
124     XMLServer.send TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
126     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
128     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
MsgBox v
'Cannot confirm delivery outcome for Invoice with MARK: 400001961787310. It has not been dispatched yet. Current status: 0</message>
 '       <code>813</code>

130     TIMOL = Replace(TIMOL, "/", "-")

132     Dim mfile As String: mfile = "c:\" + F_TXTFILES + "\sendinv\apantdeliv-" + TIMOL + "-" + MARKTIM + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"

134     Debug.Print v
136     Open mfile For Output As #8
138     Print #8, v
140     Close #8
    
142     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
144     READ_CANCEL v, MARK2

        '  400001829337674

146     ENARXIAPOS = MARK2

        '<EhFooter>
        Exit Function

CANCEL_INVOICE_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.CANCEL_INVOICE " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>
End Function

Private Sub cmdKOYKAKH_Click()

    Dim sql As String

    sql = "SELECT        G.ID_NUM AS DocID, R.TITLOS AS DocDescr, R.SYNT_TITL AS DocCode, G.ATIM AS DocNum, SUBSTRING(G.ATIM, 2, 6) AS DocAA, FORMAT(G.HME, 'yyyy-MM-ddTHH:mm:ss.fffZ') AS DocDate, G.PROOD_AJ AS LineNum, P.KOD AS CustCode,"
    sql = sql + " P.EPO AS CustName, P.AFM AS CustAFM, P.EPA AS CustProfession, P.DOY AS CustDOY, P.XRVMA AS CustZipCode, P.DIE AS CustAddress, P.POL AS CustCity, '1' AS VatStatus, LEFT(T.TRP, 1) AS Payment, "
    sql = sql + "  T.TRP AS PaymentDesc, G.KODE AS EidosCode, E.ONO AS EidosDescription, G.MONA AS UNIT, G.POSO AS Qty, G.TIMM AS Price, G.POSO * G.TIMM AS NetValue, (CASE WHEN G.FPA = 1 THEN 13 WHEN G.FPA = 2 THEN 24 END) "
    sql = sql + "   AS FPAPRC, G.MIK_AJIA - G.KAU_AJIA AS VatValue, G.POSO * G.TIMM - G.KAU_AJIA AS DiscVal, 0 AS DiscVal2, CONVERT(CHAR(10), G.HME, 103) AS InsertDate, CONVERT(varCHAR(10), G.HME, 103) "
    sql = sql + "  + ' ' + T.ORA AS LastDate into TEMP1010"
    sql = sql + " FROM            dbo.EGGTIM AS G INNER JOIN "
    sql = sql + "  dbo.PEL AS P ON G.EIDOS = P.EIDOS AND G.PELKOD = P.KOD INNER JOIN "
    sql = sql + "  dbo.TIM AS T ON G.ID_NUM = T.ID_NUM INNER JOIN "
    sql = sql + "  dbo.EID AS E ON G.KODE = E.KOD INNER JOIN "
    sql = sql + "  dbo.PARASTAT AS R ON LEFT(G.ATIM, 1) = R.EIDOS "
    sql = sql + "where R.MYEID='1.1' AND E.PROM LIKE '%ΚΟΥΚΑ%' AND G.HME>='" + Format(APO.Value, "MM/dd/yyyy") + "' AND G.HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'"

    Dim Reqpos As WinHttp.WinHttpRequest

    Set Reqpos = New WinHttp.WinHttpRequest
    '

    Dim Q As String

    Dim R As New ADODB.Recordset
  Gdb.Execute "DROP TABLE IF EXISTS TEMP1010;"
 
 Gdb.Execute "IF OBJECT_ID('dbo.TEMP1010', 'U') IS NOT NULL  DROP TABLE dbo.TEMP1010;"
 Gdb.Execute sql
  
    R.Open "SELECT * FROM TEMP1010", Gdb, adOpenDynamic, adLockOptimistic

    Q = "["

    Do While Not R.EOF
        If Len(Q) > 30 Then
            Q = Q + ",{"
         Else
            Q = Q + "{"
        End If
        Q = Q + " ""DocID"": """ + str(R!DocID) + """,  "
        Q = Q + " ""DocDescr"": """ + R!DocDescr + """,  "
        '=====
        Q = Q + " ""DocCode"": """ + R!DocCode + """,  "
        Q = Q + " ""DocNum"": """ + R!Docnum + """,  "

        Q = Q + " ""DocAA"": """ + R!Docaa + """,  "
        Q = Q + " ""DocDate"": """ + Left(R!Docdate, 19) + """,  "

        Q = Q + " ""AlterDoc"": """ + R!DocCode + """,  " 'ebala to idio me doccode
        Q = Q + " ""LineNum"": " + str(R!LINENUM) + ",  "

        Q = Q + " ""CustCode"": """ + R!custCode + """,  "
        Q = Q + " ""CustName"": """ + R!CustName + """,  "
        Q = Q + " ""CustAFM"": """ + R!custafm + """,  "
        Q = Q + " ""CustProfession"": """ + R!custprofession + """,  "

        Q = Q + " ""CustDOY"": """ + R!custdoy + """,  "
        Q = Q + " ""CustZipCode"": """ + R!custzipCode + """,  "
        Q = Q + " ""CustAddress"": """ + R!CustAddress + """,  "
        Q = Q + " ""CustCity"": """ + R!custCity + """,  "
        '-----------------------------------------------------------

        Q = Q + " ""VatStatus"": 1,  "
        Q = Q + " ""Payment"": """ + R!Payment + """,  "
    
        Q = Q + " ""PaymentDescr"": """ + R!PaymentDesc + """,  "
        Q = Q + " ""EidosCode"": """ + R!EidosCode + """,  "

        Q = Q + " ""EidosDescription"": """ + R!eidosdescription + """,  "
        Q = Q + " ""Unit"": """ + R!Unit + """,  "
        Q = Q + " ""Qty"":" + str(R!QTY) + ",  "
        Q = Q + " ""Price"": " + Replace(Format(R!Price, "###0.00"), ",", ".") + ",  "

        Q = Q + " ""NetValue"":" + Replace(Format(R!NetValue, "###0.00"), ",", ".") + ",  "
        Q = Q + " ""FPAPrc"": """ + str(nNull(R!FPAPRC)) + """,  "
        Q = Q + " ""VATValue"":" + Replace(Format(R!VATValue, "###0.00"), ",", ".") + ",  "
        Q = Q + " ""Discval"": " + Replace(Format(R!DISCVal, "###0.00"), ",", ".") + ",  "

        Q = Q + " ""DiscVal2"":" + Replace(Format(R!DISCVal2, "###0.00"), ",", ".") + ",  "
        Q = Q + " ""InsertDate"": """ + Left(R!Docdate, 19) + """,  "
        Q = Q + " ""LastUpdate"": """ + Left(R!Docdate, 19) + """,  "
        Q = Q + " ""IsActive"": 1,  "

        Q = Q + " ""OrderID"": null,  "
        Q = Q + " ""debugging"": ""d""  "

        Q = Q + " }  "
    
        DoEvents
        Me.Caption = R!Docnum
        R.MoveNext
    
    Loop

    Q = Q + " ] "
    Gdb.Execute "DROP TABLE TEMP1010"
    Open "C:\MERCVB\KOYKAKIS.TXT" For Output As #3
    Print #3, Q
    Close #3

    If Len(Q) < 20 Then
        MsgBox "δεν υπαρχουν εγγραφες"

        Exit Sub

    End If

    Dim uripos As String: uripos = "https://apps.koukakisfarm.gr/dc/dc.php"

    '
    With Reqpos
        .Open "POST", uripos, async:=False
      
        .setRequestHeader "accept", "application/json"
        .setRequestHeader "X-Client-Code", "ANTHOPOULOS"     ', "Bearer " + gf_Bearer
        .setRequestHeader "X-Client-Password", "Y7uaGuNqjWkGcQH5R2f5"
        .setRequestHeader "Content-Type", "application/json"
        ' qpos = " {  ""Bearer"":""" + gf_Bearer + """, " & ResultJSONpos & ",     ""externalSystemId"": """ + LTrim(str(SQLDT("ID_NUM"))) + """ }"
        .send (Q)
                        
        MsgBox ("απαντηση : " + Chr(13) + Reqpos.responseText)
                         
    End With
                        
    Dim posjson As String: posjson = Reqpos.responseText
                        
    Dim ser     As Object
                        
    Set ser = JSON.parse(posjson)

    'CURL -x 'POST' \
    'https://virtserver.swaggerhub.com/koukakisfarmsa/DataCollection_FarmaKoukaki_Infosupport/1.0.0/dc/dc.php' \
    '  -H 'accept: application/json' \
    '  -H 'X-Client-Code: ANTHOPOULOS' \
    '  -H 'X-Client-Password: Y7uaGuNqjWkGcQH5R2f5' \
    '  -H 'Content-Type: application/json' \
    '  -d '[
    '  {
    '    "DocID": "example-doc-000001",
    '    "DocDescr": "?e????af? pa?astat????",

    '    "DocCode": "se???-001",
    '    "DocNum": "???000001",

    '    "DocAA": "000001",
    '    "DocDate": "2026-01-16T00:00:00",

    '    "AlterDoc": "???000001",
    '    "LineNum": 1,

    '    "CustCode": "pe??t??-001",
    '    "CustName": "???µa ?e??t?",
    '    "CustAFM": "123456789",
    '    "CustProfession": "??ast????t?ta ?e??t?",

    '    "CustDOY": "??? ?e??t?",
    '    "CustZipCode": "12345",
    '    "CustAddress": "??e????s? ?e??t? 1",
    '    "CustCity": "???? ?e??t?",

    '    "VatStatus": 1,
    '    "Payment": "1",
    '    "PaymentDescr": "?e????af? ?????µ??",
    '    "EidosCode": "e?d??-001",

    '    "EidosDescription": "?e????af? ??d??? 1",
    ''    "Unit": "teµ",
    '    "Qty": 1,
    '    "Price": 2.5,

    '    "NetValue": 2.5,
    '    "FPAPrc": "13",
    '    "VATValue": 0.33,
    '    "Discval": 0,

    '    "DiscVal2": 0,
    '    "InsertDate": "2026-01-16T11:36:20",
    '    "LastUpdate": "2026-01-16T11:36:20",
    '    "IsActive": 1,
    '    "OrderID": null
    '  },
    '  {
    '    "DocID": "example-doc-000001",
    '    "DocDescr": "?e????af? pa?astat????",
    '    "DocCode": "se???-001",
    '    "DocNum": "pa?-000001",
    '    "DocAA": "000001",
    '    "DocDate": "2026-01-16T00:00:00",
    '    "AlterDoc": "pa?-000001",
    '    "LineNum": 2,
    ''    "CustCode": "pe??t??-001",
    '    "CustName": "???µa ?e??t?",
    '    "CustAFM": "123456789",
    '    "CustProfession": "??ast????t?ta ?e??t?",
    '    "CustDOY": "??? ?e??t?",
    '    "CustZipCode": "12345",
    '    "CustAddress": "??e????s? ?e??t? 1",
    '    "CustCity": "???? ?e??t?",
    '    "VatStatus": 1,
    '    "Payment": "1",
    '    "PaymentDescr": "?e????af? ?????µ??",
    '    "EidosCode": "e?d??-002",
    '    "EidosDescription": "?e????af? ??d??? 2",
    '    "Unit": "teµ",
    '    "Qty": 3,
    '    "Price": 4.25,
    '    "NetValue": 12.75,
    '    "FPAPrc": "13",
    '    "VATValue": 1.66,
    '    "Discval": 0,
    '    "DiscVal2": 0,
    '    "InsertDate": "2026-01-16T11:36:20",
    '    "LastUpdate": "2026-01-16T11:36:20",
    '    "IsActive": 1,
    '    "OrderID": null,
    '    "debugging": "d"
    '  }
    ']'
End Sub

Private Sub PL8_4_Click()
Dim SJ As String

  
         SJ = " {""invoice"":{"   ' arxh invoice  ----
   SJ = " {""issuer"":{"   ' arxh invoice  ----

 SJ = SJ + Chr(13) + " ""vatNumber"": """ + "AFM_COMPANY" + """ ,  "
 SJ = SJ + Chr(13) + " ""country"": """ + "GR" + """ },  "
 SJ = SJ + Chr(13) + " ""counterpart"": { "

 SJ = SJ + Chr(13) + " ""vatNumber"": """ + "000000000" + """ },  "

 SJ = SJ + Chr(13) + " ""name"": """ + "000000000" + """ },  "
  SJ = SJ + Chr(13) + " ""country"": """ + "GR" + """ },  "
 
 SJ = SJ + Chr(13) + " ""invoiceHeader"": { "
 
  SJ = SJ + Chr(13) + " ""series"": """ + "ΕΙΣ" + """ },  "
   SJ = SJ + Chr(13) + " ""aa"": " + "00001" + "  ,  "
    SJ = SJ + Chr(13) + " ""issueDate"": """ + "2026-06-10" + """ },  "
     SJ = SJ + Chr(13) + " ""invoiceType"": """ + "8.4" + """ },  "
      SJ = SJ + Chr(13) + " ""currency"": """ + "EUR" + """ },  "
  SJ = SJ + Chr(13) + " ""invoiceNote"": """ + "SXOLIA" + """ }, "
 
 
  SJ = SJ + Chr(13) + " ""invoiceDetails"": [{ "
   SJ = SJ + Chr(13) + " ""netValue"": " + "150.00" + "  ,  "
  SJ = SJ + Chr(13) + " ""vatCategory"": " + "8" + "  ,  "
   SJ = SJ + Chr(13) + " ""classificationCategory"": """ + "category1_95" + """ } ], "
   
    SJ = SJ + Chr(13) + " ""paymentMethods"": [{ "
  SJ = SJ + Chr(13) + " ""type"": " + "7" + "  ,  "
   SJ = SJ + Chr(13) + " ""amount"": " + "150.00" + " }] ,  "
   
  SJ = SJ + Chr(13) + " ""invoiceSummary"": { "
   SJ = SJ + Chr(13) + " ""totalValue"": " + "150.00" + "  },  "
   
   SJ = SJ + Chr(13) + " ""Messages"": [{ "
   SJ = SJ + Chr(13) + " ""type"": " + "0" + "  ,  "
    SJ = SJ + Chr(13) + " ""recipients"": " + """""" + "  }]}}  "
  

'   { invoice: { "
'            "issuer": {
'                "vatNumber": "000000000",
'                "country": "GR"
'            },
'            "counterpart": {
'                "vatNumber": "802023570",
'                "name": "ΠΑΡΟΧΟΣ ΛΥΣΕΩΝ ΠΛΗΡΟΦΟΡΙΚΗΣ ΑΝΩΝΥΜΗ ΕΤΑΙΡΕΙΑ",
'                "country": "GR"
'            },

'            "invoiceHeader": {
'                "series": "ΕΙΣ-001",
'                "aa": 1194474,
'                "issueDate": "2026-06-10",
'                "invoiceType": "8.4",
'                "currency": "EUR",


'                "invoiceNote": "ΣΧΟΛΙΟ ΠΑΡΑΣΤΑΤΙΚΟΥ"
'            },
'            "invoiceDetails": [
'                {
'                    "netValue": 150.00,
'                    "vatCategory": 8,
'                    "classificationCategory": "category1_95"
'                }
'            ],
'            "paymentMethods": [
'                {
'                    "type": 7
'                    "amount": 150.00
'                }
'            ],
'            "invoiceSummary": {
'                "totalValue": 150.00
'            },
'            "Messages": [
'                {
'                    "type": 0,
'                    "recipients": ""
'                }
'            ]
'        }
'    }


































End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err
If F_PAROCHOS = 1 Then
  If MYDPAR.Text = "MYDATA" Then
  
      If chkexodon.Value = vbChecked Then
           Shell "notepad.exe c:\" + F_TXTFILES + "\apantSendEXP.XML", vbMaximizedFocus
        
        Else
           Shell "notepad.exe c:\" + F_TXTFILES + "\apantSendInv.XML", vbMaximizedFocus

        End If

  
  
  
  
  Else
     Shell "notepad.exe c:\" + F_TXTFILES + "\apantinv.jsn", vbMaximizedFocus  '400010164786452
  End If


Else
        '</EhHeader>
100     If chkexodon.Value = vbChecked Then
102         Shell "notepad.exe c:\" + F_TXTFILES + "\apantSendEXP.XML", vbMaximizedFocus
        
        Else
104         Shell "notepad.exe c:\" + F_TXTFILES + "\apantSendInv.XML", vbMaximizedFocus

        End If

End If


        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Command3_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        ' getReqANSI.XML
        '<EhHeader>
        On Error GoTo Command5_Click_Err
     If F_PAROCHOS = 1 Then
          Shell "notepad.exe c:\" + F_TXTFILES + "\inv.jsn", vbMaximizedFocus
 
     Else
        '</EhHeader>
100     If chkexodon.Value = vbChecked Then
102         Shell "notepad.exe c:\" + F_TXTFILES + "\EXP.xml", vbMaximizedFocus
        Else
104         Shell "notepad.exe c:\" + F_TXTFILES + "\Inv.XML", vbMaximizedFocus
        End If
        
       
        
        
        
        
        End If
        
        
        Exit Sub

Command5_Click_Err:
106     MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Command5_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

108     Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        'getReqANSI.xml
        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>

100     Shell "notepad.exe c:\" + F_TXTFILES + "\getReqANSI.XML", vbMaximizedFocus

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Command6_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdepanfromfile_Click()

        '<EhHeader>
        On Error GoTo cmdepanfromfile_Click_Err

        '</EhHeader>
        Dim m_idnum As String
        
        Dim ANS     As Integer

100     ANS = MsgBox("Προσοχή!!! Να σταλεί το inv.xml στο MyData", vbYesNo)

102     If ANS = vbYes Then
104         m_idnum = InputBox("ΔΩΣΕ ΤΟ ID_NUM ΤΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ ΠΟΥ ΘΑ ΠΕΣΕΙ ΤΟ ΜΑΡΚ", "", "0")

            Dim MARK As String

106         MARK = GGET_CVALUE("select ENTITYMARK FROM TIM WHERE ID_NUM=" + m_idnum)

108         If gVal(MARK) = 0 Then
                Dim markReal As String: markReal = ""
110             MAKE_request "", 1, 0, m_idnum, markReal
112             MsgBox "OK " + markReal
            Else
114             MsgBox " ΈΧΕΙ ΗΔΗ ΜΑΡΚ ΤΟ ID_NUM ΠΟΥ ΖΗΤΗΣΑΤΕ"
            End If
        End If
       
        Exit Sub

cmdepanfromfile_Click_Err:
116     MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Command7_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

118     Resume Next

        '</EhFooter>

End Sub

Private Sub Command7_Click()

    '  Dim stringa As String
    '  stringa = "{""data"":{""GET:oauth.openapi.it\/counters"":{""counter"":22,""paid"":0,""limit"":false},""GET:oauth.openapi.it\/scopes"":{""counter"":6,""paid"":0,""limit"":false},""POST:oauth.openapi.it\/token"":{""counter"":1,""paid"":0,""limit"":false},""GET:imprese.openapi.it\/advance"":{""counter"":14,""paid"":0,""limit"":false},""GET:imprese.openapi.it\/base"":{""counter"":2,""paid"":0,""limit"":false}},""success"":true,""message"":"""",""error"":null}"
    
    '  Dim oJson As Object
    ' Set oJson = JsonParseObject(stringa)
    ' Debug.Print JsonDump(oJson, maxWidth:=0)

End Sub

Private Sub Command8_Click()

        ' TREXV TA APESTALMENA KAI TA ANTISOIXIZV STO TIM
        '<EhHeader>
        On Error GoTo Command8_Click_Err

        '</EhHeader>
        Dim r0  As New ADODB.Recordset

        Dim r1  As New ADODB.Recordset

        Dim sql As String

100     'sql = "SELECT * FROM APESTALMENA WHERE ISSUEDATE>='" + Format(apo.Value, "yyyy-MM-dd") + "'  AND ISSUEDATE<'" + Format(DateAdd("d", 1, eos.Value), "yyyy-MM-dd") + "' AND ISANTIST IS NULL order by MARK DESC"
 sql = "SELECT * FROM APESTALMENA WHERE ISANTIST IS NULL order by MARK DESC"
102     r0.Open sql, Gdb, adOpenDynamic, adLockOptimistic

104     Do While Not r0.EOF

106         If IsNull(r0!EPO) Then
108            ' Gdb.Execute "UPDATE APESTALMENA SET EPO=(SELECT TOP 1 '**'+ISNULL(LEFT(EPO,23),'') AS EPO FROM PEL WHERE AFM='" + r0!AFM + "') WHERE ID=" + str(r0!ID)
            End If

110         r1.Open "SELECT PEL.EPO,PEL.AFM,TIM.* FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD WHERE CONVERT(CHAR(10),HME,103)='" + Format(r0!issueDate, "dd/MM/yyyy") + "' and ATIM LIKE '%" + Right(Trim(LTrim((r0!AA))), 6) + "'", Gdb, adOpenDynamic, adLockOptimistic

112         If Not r1.EOF Then
                 Gdb.Execute "UPDATE APESTALMENA SET EPO=(SELECT TOP 1 '**'+ISNULL(LEFT(EPO,23),'') AS EPO FROM PEL WHERE AFM='" + r0!AFM + "') WHERE ID=" + str(r0!ID)
                'If r1!AFM = r0!AFM Then
                '     Gdb.Execute "UPDATE TIM SET SXETMARK='" + r0!mark + "' WHERE ID_NUM=" + str(r1!ID_NUM)
                'End If
         
                'If Trim(r1!AFM) = Trim(r0!AFM) Then
                ' Gdb.Execute "UPDATE TIM SET AADEKAU=" + Replace(str(r0!totalNetValue), ",", ".") + ",AADEFPA=" + Replace(str(r0!totalVatAmount), ",", ".") + ", SXETMARK='" + r0!mark + "' WHERE ID_NUM=" + str(r1!ID_NUM)
                ' End If
         
114             If Trim(r1!AFM) = Trim(r0!AFM) Then
116                ' Gdb.Execute "UPDATE TIM SET AADEKAU=" + Replace(str(r0!totalNetValue), ",", ".") + ",AADEFPA=" + Replace(str(r0!totalVatAmount), ",", ".") + ", SXETMARK='" + r0!MARK + "' WHERE ID_NUM=" + str(r1!id_num)
                End If
         
                '      Else
                '
                'Exit Do
      
118             Gdb.Execute "update APESTALMENA SET ATIM='" + r1!ATIM + "' , ISANTIST=1 WHERE ID= " + str(r0!ID)

        
            End If
       
            ' Gdb.Execute "update APESTALMENA SET ATIM='" + r1!ATIM + "' , ISANTIST=1 WHERE ID= " + str(r0!ID)
            ' Gdb.Execute "UPDATE TIM SET SXETMARK='" + Trim(r0!mark) + "' WHERE ID_NUM=" + str(r1!ID_NUM)
       
120         r1.Close
 
122         r0.MoveNext
        Loop
 
124     r0.Close


     sql = "select  P.EPO,A.* from APESTALMENA A inner JOIN PEL P ON P.EIDOS='r' and P.AFM=A.AFM   WHERE A.AA IN ( SELECT ATIM FROM APESTALMENADET WHERE LEN(KODE)>1 )"
216     Adodc2.ConnectionString = gConnect
218     Adodc2.RecordSource = sql
220     Adodc2.Refresh




        '<EhFooter>
        Exit Sub

Command8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Command8_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

        '    ' TREXV TA APESTALMENA KAI TA ANTISOIXIZV STO TIM
        '        '<EhHeader>
        '        On Error GoTo Command8_Click_Err
        '        '</EhHeader>
        '    Dim r0 As New ADODB.Recordset
        '    Dim r1 As New ADODB.Recordset
        '
        '100  r0.Open "SELECT * FROM APESTALMENA WHERE ISANTIST IS NULL", Gdb, adOpenDynamic, adLockOptimistic
        '102  Do While Not r0.EOF
        '104    r1.Open "SELECT * FROM TIM WHERE CONVERT(CHAR(10),HME,103)='" + Format(r0!issueDate, "dd/MM/yyyy") + "' and ATIM LIKE '%" + LTrim(str(r0!AA)) + "'", Gdb, adOpenDynamic, adLockOptimistic
        '106    If Not r1.EOF Then
        '
        '110       Gdb.Execute "UPDATE TIM SET ENTITYMARK='" + r0!MARK + "' WHERE ID_NUM=" + str(r1!ID_NUM)
        ''
        '
        '
        '
        '114        Gdb.Execute "update APESTALMENA SET ATIM='" + r1!ATIM + "' , ISANTIST=1 WHERE ID= " + str(r0!ID)
        '116         Gdb.Execute "UPDATE TIM SET ENTITYMARK='" + Trim(r0!MARK) + "' WHERE ID_NUM=" + str(r1!ID_NUM)
        '
        '
        '       End If
        '120    r1.Close
        '
        '
        '
        '122    r0.MoveNext
        '     Loop
        '
        '124  r0.Close
        '
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
        'Command8_Click_Err:
        '        MsgBox Err.Description & vbCrLf & _
        '               "in ADOMERCNEW.Par7MyData.Command8_Click " & _
        '               "at line " & Erl, _
        '               vbExclamation + vbOKOnly, "Application Error"
        '        Resume Next
        '        '</EhFooter>
End Sub

Private Sub Command9_Click()

        '<EhHeader>
        On Error GoTo Command9_Click_Err

        '</EhHeader>
        Dim MARK As String, CANCMARK As String

100     MARK = InputBox("δωσε τo mark", "μαρκ")
 
102     CANCMARK = CANCEL_INVOICE(MARK, " ")
 
        '<EhFooter>
        Exit Sub

Command9_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Command9_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub
 
Public Function CANCEL_INVOICE(ByVal MARKTIM As String, TIMOL As String) As String

        '<EhHeader>
        On Error GoTo CANCEL_INVOICE_Err

        '</EhHeader>
100     If Len(Trim(MARKTIM)) < 1 Then
102         CANCEL_INVOICE = ""

            Exit Function

        End If

        '            Dim uri As String: uri = "https://mydata-dev.azure-api.net/CancelInvoice?mark=400000020235194"

        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
104     url2 = UrlAADE + "CancelInvoice?mark=" '/RequestDocs"
     
        Dim MARK As String

        '    MARK = InputBox("δωσε τo mark", "μαρκ")
     
106     url2 = url2 + MARKTIM

        'initialize
108     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
110     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
112     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
114     XMLServer.Option(9) = 2048
116     XMLServer.Option(6) = True
    
118     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
120     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
122     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
124     XMLServer.send TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
126     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
128     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------

130     TIMOL = Replace(TIMOL, "/", "-")

132     Dim mfile As String: mfile = "c:\" + F_TXTFILES + "\sendinv\apantCANC-" + TIMOL + "-" + MARKTIM + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"

134     Debug.Print v
136     Open mfile For Output As #8
138     Print #8, v
140     Close #8
    
142     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
144     READ_CANCEL v, MARK2

        '  400001829337674

146     CANCEL_INVOICE = MARK2

        '<EhFooter>
        Exit Function

CANCEL_INVOICE_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.CANCEL_INVOICE " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Sub READ_CANCEL(ByVal v As String, ByRef MARK As String)

        '=========================================
        '<EhHeader>
        On Error GoTo READ_CANCEL_Err

        '</EhHeader>
        Dim C1, c2, c3, c4, SXOLIA As String

        Dim objxml As New MSXML2.DOMDocument

100     objxml.async = True
102     objxml.loadXML v

        ' TO IDIO DIABAZEI APO ARXEIO  =>     objxml.Load ("C:\"+F_TXTFILES+"\apantSendCANC.xml")
        Dim nLathosApost As Long

104     nLathosApost = 0

        Dim nApost As Long

106     nApost = 0
    
        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode

108     Set nodeList = objxml.SelectNodes("ResponseDoc/response")

110     For Each Node In nodeList

            Dim xmlNode As IXMLDOMNode

            ' Dim xml2Node As IXMLDOMNode
            ' Dim xml3Node As IXMLDOMNode
112         If Node.hasChildNodes Then
114             C1 = "": c2 = "": c3 = ""

116             For Each xmlNode In Node.childNodes

118                 If xmlNode.nodeName = "cancellationMark" Then
120                     C1 = xmlNode.Text

                        Exit For
                 
                    End If
                    
                    '<statusCode>Success</statusCode>
                    
                       If xmlNode.nodeName = "errors" Then 'statusCode" Then
                           If xmlNode.hasChildNodes Then
                               Dim Node2     As IXMLDOMNode
                               For Each Node2 In xmlNode.childNodes
                                   If Node2.nodeName = "error" Then MsgBox Node2.Text
                               
                               Next
                               
                               
                               
                           End If
                 
                       End If
                    
                    
                    

122             Next xmlNode

            End If

        Next

124     MARK = C1

        '<EhFooter>
        Exit Sub

READ_CANCEL_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.READ_CANCEL " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Function READ_xml(ByVal v As String, mNode As String, mName As String, ByRef lathos As String) As String

        '=========================================
        '<EhHeader>
        On Error GoTo READ_CANCEL_Err

        '</EhHeader>
        Dim C1, c2, c3, c4, SXOLIA As String

        Dim objxml As New MSXML2.DOMDocument

100     objxml.async = True
102     objxml.loadXML v

        ' TO IDIO DIABAZEI APO ARXEIO  =>     objxml.Load ("C:\"+F_TXTFILES+"\apantSendCANC.xml")
        Dim nLathosApost As Long

104     nLathosApost = 0

        Dim nApost As Long

106     nApost = 0
    
        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode

108     Set nodeList = objxml.SelectNodes(mNode) '"ResponseDoc/response")

110     For Each Node In nodeList

            Dim xmlNode As IXMLDOMNode

            ' Dim xml2Node As IXMLDOMNode
            ' Dim xml3Node As IXMLDOMNode
112         If Node.hasChildNodes Then
114             C1 = "": c2 = "": c3 = ""

116             For Each xmlNode In Node.childNodes

118                 If xmlNode.nodeName = mName Then  '"cancellationMark"
120                     C1 = xmlNode.Text

                        Exit For
                 
                    End If
                    
                    '<statusCode>Success</statusCode>
                    
                       If xmlNode.nodeName = "errors" Then 'statusCode" Then
                           If xmlNode.hasChildNodes Then
                               Dim Node2     As IXMLDOMNode
                               For Each Node2 In xmlNode.childNodes
                                   If Node2.nodeName = "error" Then
                                       'MsgBox
                                       lathos = Node2.Text
                                   End If
                               
                               Next
                               
                               
                               
                           End If
                 
                       End If
                    
                    
                    

122             Next xmlNode

            End If

        Next

124     READ_xml = C1

        '<EhFooter>
        Exit Function

READ_CANCEL_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.READ_CANCEL " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function






Private Sub eisagTIM_Click()

On Error GoTo OUTOUT
    Dim m As String
    m = Adodc2.Recordset("AA")
    m = Trim(m)
    MAFM = Trim(Adodc2.Recordset("AFM"))
    MKOD = GGET_CVALUE("SELECT KOD FROM PEL WHERE EIDOS='r' AND AFM='" + MAFM + "'")
    
On Error GoTo OUTDOOR
Dim n2 As Integer
n2 = GGET_NVALUE("select count(*) from TIM WHERE ATIM='" + m + "'")
If n2 > 0 Then
  MsgBox ("υπαρχει ηδη")
  Exit Sub
End If

Gdb.BeginTrans
Gdb.Execute ("INSERT INTO TIM (B_N1,TRP,KPE,EIDOS,ATIM,KLEIDI,AJI,HME) SELECT 1 AS B1,'2.ΠΙ','" + MKOD + "','r',AA,AA,TOTALNETVALUE+TOTALVATAMOUNT,ISSUEDATE FROM APESTALMENA WHERE AA='" + m + "'")
Dim mID_NUM As String: mID_NUM = GGET_CVALUE("SELECT STR(MAX(ID_NUM)) FROM TIM ")
'MKOD = GGET_CVALUE("SELECT KOD FROM PEL WHERE EIDOS='e' AND AFM='" + MAFM + "'")

Gdb.Execute ("INSERT INTO EGGTIM (APOT,FPA,KODE,MONA,TIMM,KAU_AJIA,ONOMA,ID_NUM,PELKOD,EIDOS,ATIM,POSO,XRE,HME) SELECT 1,FPA,ISNULL(KODE,'@@@') AS KODE,MONA,TIMM,KAU_AJIA,ONOMA," + mID_NUM + ",'" + MKOD + "','r',ATIM,POSO,XRE,HME  FROM APESTALMENADET WHERE ATIM='" + m + "'")
Gdb.Execute ("UPDATE TIM SET AJ1=(SELECT SUM(POSO*TIMM)FROM EGGTIM WHERE ID_NUM=" + mID_NUM + " AND FPA=1) WHERE ID_NUM=" + mID_NUM)
Gdb.Execute ("UPDATE TIM SET AJ2=(SELECT SUM(POSO*TIMM)FROM EGGTIM WHERE ID_NUM=" + mID_NUM + " AND FPA=2) WHERE ID_NUM=" + mID_NUM)
Gdb.Execute ("UPDATE TIM SET AJ3=(SELECT SUM(POSO*TIMM)FROM EGGTIM WHERE ID_NUM=" + mID_NUM + " AND FPA=3) WHERE ID_NUM=" + mID_NUM)
Gdb.Execute ("UPDATE TIM SET AJ4=(SELECT SUM(POSO*TIMM)FROM EGGTIM WHERE ID_NUM=" + mID_NUM + " AND FPA=4) WHERE ID_NUM=" + mID_NUM)
Gdb.Execute ("UPDATE TIM SET AJ5=(SELECT SUM(POSO*TIMM)FROM EGGTIM WHERE ID_NUM=" + mID_NUM + " AND FPA=5) WHERE ID_NUM=" + mID_NUM)
Gdb.Execute ("UPDATE TIM SET AJ6=(SELECT SUM(POSO*TIMM)FROM EGGTIM WHERE ID_NUM=" + mID_NUM + " AND FPA=6) WHERE ID_NUM=" + mID_NUM)

Gdb.Execute ("UPDATE TIM SET FPA1=AJ1/100*(SELECT TIMH FROM PINAKES WHERE TYPOS=1 AND AYJON=1) WHERE ID_NUM=" + mID_NUM)
Gdb.Execute ("UPDATE TIM SET FPA2=AJ1/100*(SELECT TIMH FROM PINAKES WHERE TYPOS=1 AND AYJON=2) WHERE ID_NUM=" + mID_NUM)
Gdb.Execute ("UPDATE TIM SET FPA3=AJ1/100*(SELECT TIMH FROM PINAKES WHERE TYPOS=1 AND AYJON=3) WHERE ID_NUM=" + mID_NUM)
Gdb.Execute ("UPDATE TIM SET FPA4=AJ1/100*(SELECT TIMH FROM PINAKES WHERE TYPOS=1 AND AYJON=4) WHERE ID_NUM=" + mID_NUM)
Gdb.Execute ("UPDATE TIM SET FPA6=AJ1/100*(SELECT TIMH FROM PINAKES WHERE TYPOS=1 AND AYJON=6) WHERE ID_NUM=" + mID_NUM)
Gdb.Execute ("UPDATE EGGTIM SET KODE='@@@'  WHERE LEN(KODE)<1 AND ID_NUM=" + mID_NUM)
Dim R As New ADODB.Recordset
R.Open "SELECT ISNULL(KODE,'@@@') AS KODE ,ONOMA,ISNULL(FPA,2) AS FPA,ISNULL(TIMM,0) AS TIMM FROM EGGTIM WHERE ID_NUM=" + mID_NUM, Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
    If Len(GGET_CVALUE("SELECT KOD FROM EID WHERE KOD='" + R!KODE + "'")) > 2 Then
       ' OK YPARXEI TO EIDOS
    Else
        m = R!KODE
        If Len(m) < 1 Then m = "@@@"
       Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,XTI) VALUES ('" + m + "','" + R!ONOMA + "'," + str(R!FPA) + "," + Replace(str(R!TIMM), ",", ".") + ")"
    End If
    R.MoveNext
    
Loop
Gdb.CommitTrans
MsgBox "OK"
Exit Sub



OUTDOOR:
On Error Resume Next
Gdb.RollbackTrans

MsgBox "ΛΑΘΟΣ" + Err.Description

OUTOUT:
MsgBox "ΛΑΘΟΣ" + Err.Description

End Sub

Private Sub ektos_Click()

100     If ektos.Value = vbChecked Then
102         Check1.Value = vbUnchecked
        End If
    
End Sub

Private Sub Form_DblClick()
       
100     PARAMETROI.PARAM.Caption = "PAR7MYDATA"
102     PARAMETROI.SHOW 1

End Sub

Private Sub Form_Paint()

    '  Unload Me
    Dim a As Long
    If F_PAROCHOS = 1 Then
        cmdPAROCHOS.BackColor = vbGreen
    End If
        

    'a = Dt2Lng(Now)

End Sub

Private Sub hkoLD_Click()

        '<EhHeader>
        On Error GoTo hkoLD_Click_Err

        '</EhHeader>
 
        ' ΣΥΝΔΕΣΗ ΜΕ ΤΙΜΛΟΓΙΑ ΑΓΟΡΑΣ
        ' SELECT T.ATIM,* FROM APESTALMENA  A  LEFT JOIN TIM T ON TOTALNETVALUE=AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7
    
        'Dim sql As String
'100     sql = "select  (select TOP 1 TITLOS FROM PARASTAT WHERE MYEID= RTRIM(SUBSTRING(A.TYPOS,1,4))+'%') AS [ΠΑΡΑΣΤ],P.EPO,ATIM,MARK,TOTALNETVALUE AS [ΚΑΘ.ΑΞΙΑ], A.* from APESTALMENA A LEFT JOIN PEL P ON P.AFM=A.AFM AND P.EIDOS='r'  ORDER BY ISSUEDATE DESC" ' ISSUEDATE >='" + Format(APO.Value, "yyyy-MM-dd") + "'  AND  ISSUEDATE<'" + Format(DateAdd("d", 1, EOS.Value), "yyyy-MM-dd") + "'  ORDER BY ISSUEDATE DESC"
'        '"SELECT * FROM APESTALMENA WHERE issueDate>='" + Format(apo.Value, "MM/dd/yyyy") + "'  AND issueDate<'" + Format(DateAdd("d", 1, eos.Value), "MM/dd/yyyy") + "' AND ISANTIST IS NULL order by MARK DESC"
'
'102     Adodc2.ConnectionString = gConnect
'104     Adodc2.RecordSource = sql
'106     Adodc2.Refresh
'


    sql = "select ISANTIST, P.EPO,A.* from APESTALMENA A inner JOIN PEL P ON P.EIDOS='r' and P.AFM=A.AFM   WHERE A.AA IN ( SELECT ATIM FROM APESTALMENADET WHERE LEN(KODE)>1 )"
216     Adodc2.ConnectionString = gConnect
218     Adodc2.RecordSource = sql
220     Adodc2.Refresh





        '        '<EhHeader>
        '        On Error GoTo hkoLD_Click_Err
        '        '</EhHeader>
        '
        '
        '
        '        ' ΣΥΝΔΕΣΗ ΜΕ ΤΙΜΛΟΓΙΑ ΑΓΟΡΑΣ
        '        ' SELECT T.ATIM,* FROM APESTALMENA  A  LEFT JOIN TIM T ON TOTALNETVALUE=AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7
        '
        '        'Dim sql As String
        '100    sql = "select  (select TOP 1 TITLOS FROM PARASTAT WHERE MYEID= RTRIM(SUBSTRING(A.TYPOS,1,4))+'%') AS [ΠΑΡΑΣΤ],P.EPO,ATIM,MARK, A.* from APESTALMENA A LEFT JOIN PEL P ON P.AFM=A.AFM AND P.EIDOS='r' ORDER BY ID DESC"
        '
        '
        '102 Adodc2.ConnectionString = gConnect
        '104 Adodc2.RecordSource = sql
        '106 Adodc2.Refresh

        '<EhFooter>
        Exit Sub

hkoLD_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.hkoLD_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdCommand3_Click()

        'Set a reference to "Microsoft XML, v3.0"
        '<EhHeader>
        On Error GoTo cmdCommand3_Click_Err

        '</EhHeader>

        '<?xml version="1.0" encoding="UTF-8"?>
        '<log>
        '<error>
        ' <DateError>12/10/2020</DateError>
        '  <ErrorCode>1562</ErrorCode>
        '</error>
        '
        '</log>

        Dim i As Integer

100     Randomize
    
102     i = Int(Rnd() * 5000) + 1
104     LogError Now, i

        '<EhFooter>
        Exit Sub

cmdCommand3_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdCommand3_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub LogError(ByVal ErrDate As Date, ByVal ErrCode As Long)

        '<EhHeader>
        On Error GoTo LogError_Err

        '</EhHeader>

        Dim objDoc                   As MSXML2.DOMDocument

        Dim xmlProcessingInstruction As MSXML2.IXMLDOMProcessingInstruction

        Dim objRoot                  As IXMLDOMNode

        Dim objNode                  As IXMLDOMNode

        Dim objNodeErr               As IXMLDOMNode

        Dim objNodeDetails           As IXMLDOMNode

        Dim strLogPath               As String

100     strLogPath = "c:\ErrorLog.xml"

        'Create an XML Document object
102     Set objDoc = New MSXML2.DOMDocument
    
        'Check if the log exists. If it does open it otherwise create it
104     If Dir(strLogPath) = "" Then
            'Creating standard headers
106         Set xmlProcessingInstruction = objDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
108         objDoc.appendChild xmlProcessingInstruction
110         Set xmlProcessingInstruction = Nothing
        
            'Create the Root Node
112         Set objRoot = objDoc.createElement("log")
114         objDoc.appendChild objRoot
116         Set objRoot = Nothing
        Else
118         objDoc.Load strLogPath
        End If
    
        'Get the Root Node
120     Set objRoot = objDoc.SelectSingleNode("log")
    
        'Add an error node to the root node
122     Set objNodeErr = objDoc.createElement("error")
124     objRoot.appendChild objNodeErr
126     Set objRoot = Nothing
    
        'Add the details to the error node
128     Set objNodeDetails = objDoc.createElement("DateError")
130     objNodeDetails.Text = FormatDateTime(ErrDate, 2)
132     objNodeErr.appendChild objNodeDetails
    
134     Set objNodeDetails = objDoc.createElement("ErrorCode")
136     objNodeDetails.Text = ErrCode
138     objNodeErr.appendChild objNodeDetails
    
        'clean up
140     Set objNodeDetails = Nothing
142     Set objNodeErr = Nothing
144     Set objNode = Nothing
    
        '<<<<<<< HEAD
        'Save the log
146     objDoc.save strLogPath

148     Set objDoc = Nothing

        '<EhFooter>
        Exit Sub

LogError_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.LogError " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdCommand4_Click()

        ' This procedure creates XML document
        ' and saves it to disk.
        ' Requires msxml.dll (Go to Project --> References and
        ' and choose Microsoft XML version 2.0, or whatever the
        ' current version you have installed)
        ' The example given below will write the following XML
        ' documents.
        '
        ' <Family>
        '    <Member Relationship="Father">
        '       <Name>Some Guy</name>
        '    </member>
        ' </family>
        '
        'but it should be clear how to modify the code
        'to create your own documents
        '<EhHeader>
        On Error GoTo cmdCommand4_Click_Err

        '</EhHeader>
   
        Dim objDom        As DOMDocument

        Dim objRootElem   As IXMLDOMElement

        Dim objMemberElem As IXMLDOMElement

        Dim objMemberRel  As IXMLDOMAttribute

        Dim objMemberName As IXMLDOMElement
   
100     Set objDom = New DOMDocument
   
        ' Creates root element
102     Set objRootElem = objDom.createElement("Family")
104     objDom.appendChild objRootElem
   
        ' Creates Member element
106     Set objMemberElem = objDom.createElement("Member")
108     objRootElem.appendChild objMemberElem
   
        ' Creates Attribute to the Member Element
110     Set objMemberRel = objDom.createAttribute("Relationship")
112     objMemberRel.nodeValue = "Father"
114     objMemberElem.setAttributeNode objMemberRel
   
        ' Create element under Member element, and
        ' gives value "some guy"
116     Set objMemberName = objDom.createElement("Name")
118     objMemberElem.appendChild objMemberName
120     objMemberName.Text = "Some Guy"

        '<<<<<<< HEAD
        ' Saves XML data to disk.
122     objDom.save ("c:\andrew.xml")

        '<EhFooter>
        Exit Sub

        '=======

cmdCommand4_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdCommand4_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdCommand5_Click()

        '<EhHeader>
        On Error GoTo cmdCommand5_Click_Err

        '</EhHeader>

        '<?xml version="1.0" encoding="utf-8"?>
        '<ArrayOfStock xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">

        '<Stock>
        '<ProductCode>12345</ProductCode>
        '<ProductPrice>10.32</ProductPrice>
        '</Stock>

        '<Stock>
        '<ProductCode>¥45632</ProductCode>
        '<ProductPrice>5.43</ProductPrice>
        '</Stock>

        '</ArrayOfStock>

        Dim varStock  As Variant

        'global       Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
100     varStock = Array(Array("ΕΥΡΩ12345", 10.32), Array("¥45632", 5.43)) 'Yen sign used here to show Unicode.
    
102     Set docStock = New MSXML2.DOMDocument

104     With docStock
106         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
108         Set elemRoot = .createElement("ArrayOfStock")

110         With elemRoot
112             .setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
114             .setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
            
116             For i = 0 To UBound(varStock)
118                 Set elemStock = docStock.createElement("Stock")

120                 With elemStock
122                     Set elemField = docStock.createElement("ProductCode")
124                     elemField.Text = CStr(varStock(i)(0))
126                     .appendChild elemField
128                     Set elemField = docStock.createElement("ProductPrice")
130                     elemField.Text = CStr(varStock(i)(1))
132                     .appendChild elemField
                    End With

134                 .appendChild elemStock
                Next

            End With

136         Set .documentElement = elemRoot

            On Error Resume Next

138         Kill "C:\created.xml"

            On Error GoTo cmdCommand5_Click_Err

140         .save "C:\created.xml"
        End With
    
        '<EhFooter>
        Exit Sub

cmdCommand5_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdCommand5_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

' Add formatting to the document.
Public Sub FormatXmlDocument(ByVal xml_doc As DOMDocument)

        '<EhHeader>
        On Error GoTo FormatXmlDocument_Err

        '</EhHeader>
100     FormatXmlNode xml_doc.documentElement, 0

        '<EhFooter>
        Exit Sub

FormatXmlDocument_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.FormatXmlDocument " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

' Add formatting to this element. Indent it and add a
' carriage return before its children. Then recursively
' format the children with increased indentation.
Public Sub FormatXmlNode(ByVal Node As IXMLDOMNode, ByVal indent As Integer)

        '<EhHeader>
        On Error GoTo FormatXmlNode_Err

        '</EhHeader>
        Dim child     As IXMLDOMNode

        Dim text_only As Boolean

        ' Do nothing if this is a text node.
100     If TypeOf Node Is IXMLDOMText Then Exit Sub

        ' See if this node contains only text.
102     text_only = True

104     If Node.hasChildNodes Then

106         For Each child In Node.childNodes

108             If Not (TypeOf child Is IXMLDOMText) Then
110                 text_only = False

                    Exit For

                End If

112         Next child

        End If

        ' Process child nodes.
114     If Node.hasChildNodes Then

            ' Add a carriage return before the children.
116         If Not text_only Then
118             Node.insertBefore Node.ownerDocument.createTextNode(vbCrLf), Node.firstChild
            End If

            ' Format the children.
120         For Each child In Node.childNodes

122             FormatXmlNode child, indent + 2
124         Next child

        End If

        ' Format this element.
126     If indent > 0 Then
            ' Indent before this element.
128         Node.parentNode.insertBefore Node.ownerDocument.createTextNode(Space$(indent)), Node

            ' Indent after the last child node.
130         If Not text_only Then Node.appendChild Node.ownerDocument.createTextNode(Space$(indent))

            ' Add a carriage return after this node.
132         If Node.nextSibling Is Nothing Then
134             Node.parentNode.appendChild Node.ownerDocument.createTextNode(vbCrLf)
            Else
136             Node.parentNode.insertBefore Node.ownerDocument.createTextNode(vbCrLf), Node.nextSibling
            End If
        End If

        '<EhFooter>
        Exit Sub

FormatXmlNode_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.FormatXmlNode " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdCommand6_Click()

End Sub

Private Sub cmdDOMDocumentUTF8ΟΛΔ_Click()

        '<EhHeader>
        On Error GoTo cmdDOMDocumentUTF8_Click_Err

        '</EhHeader>

        Dim varStock  As Variant

        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
100     varStock = Array(Array("ΕΥΡΩ12345", 10.32), Array("¥45632", 5.43)) 'Yen sign used here to show Unicode.
    
102     Set docStock = New MSXML2.DOMDocument

104     With docStock
106         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
108         Set elemRoot = .createElement("ArrayOfStock")
            '  With elemRoot/////////////////////////////////////////////////////////////////////////////////////////
110         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
112         elemRoot.setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
            
114         For i = 0 To UBound(varStock)
116             Set elemStock = docStock.createElement("Stock")
                ' With elemStock----------------------------------
               
                '                    add_c Stock, "ProductCode", "11111"
                '                     add_c Stock, "ProductPrice", "1.35"
               
118             Set elemField = docStock.createElement("ProductCode")
                       
                ' δημιουργω εσοχη
120             Set elem2Field = docStock.createElement("Product2Price")
122             elem2Field.Text = "22222"
124             elemField.appendChild elem2Field
                       
126             elemStock.appendChild elemField
                    
128             Set elemField = docStock.createElement("ProductPrice"): elemField.Text = "22222": elemStock.appendChild elemField
                    
                '<<<<<<< HEAD
                ' End With--------------------------------------
130             elemRoot.appendChild elemStock
            Next

            '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
132         Set .documentElement = elemRoot

            On Error Resume Next

134         Kill "C:\created.xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

136         .save "C:\created.xml"
        End With
    
        '<EhFooter>
        Exit Sub

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdDOMDocumentUTF8_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'Sub add_1(elem As MSXML2.IXMLDOMElement, klados As String)
'
'    Set elem = docStock.createElement(klados)
'
'End Sub

Sub add_c(ByRef mParent As MSXML2.IXMLDOMElement, mChild As String, MTEXT As String)

        '<EhHeader>
        On Error GoTo add_c_Err

        '</EhHeader>
        Dim elemField As MSXML2.IXMLDOMElement

100     Set elemField = docStock.createElement(mChild)
102     elemField.Text = MTEXT
104     elemField.appendChild elemField

        '<EhFooter>
        Exit Sub

add_c_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.add_c " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdTestantinvUpdateTim_Click()
    'UpdateTIMwithMARK
End Sub

Private Sub Command1_Click()

        'Dim doc as XmlDocument = new XmlDocument()
        'doc.Load ("config.xml")
        'Dim root As XmlElement = doc.DocumentElement
        'PackageName = root.GetAttribute("id")
        'ProjectVersion = root.GetAttribute("version")
        'ProjectName = root.GetElementsByTagName("name").Item(0).InnerText
        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     Set objxml = CreateObject("Msxml.DOMDocument")
102     objxml.async = True
104     objxml.Load "C:\" + F_TXTFILES + "\apantreq.xml"

        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode

106     Gdb.Open "DSN=MERCSQL;"
108     Set nodeList = objxml.SelectNodes("RequestedDoc/invoicesDoc/invoice")

110     For Each Node In nodeList

112         Print Node.nodeName  ' this works'
114         List1.AddItem Node.nodeName + "*******" + Node.Text  ' INVOICE
       
            '       Print node.n
116         Call printNode(Node, 0)     'here is the problem explained below'
    
118         If Len(Trim(payaji)) = 0 Then payaji = "0"
120         If Len(Trim(totalGrossValue)) = 0 Then totalGrossValue = "0"

122         sql = "INSERT INTO APESTALMENA  ([MARK],[AFM],[ISSUEDATE],[AA],[TYPOS],[PAYTYPE],[PAYAJI],[TOTALNETVALUE],[TOTALVATAMOUNT],[TOTALGROSSVALUE],[CLASSIFICATIONTYPE])"
124         sql = sql + " Values ('" + MARK + "','" + AFM + "','" + issueDate + "','" + AA + "','" + invoiceType + "','" + paytype + "'," + Replace(payaji, ",", ".") + "," + Replace(totalNetValue, ",", ".") + "," + Replace(totalVatAmount, ",", ".") + "," + Replace(totalGrossValue, ",", ".") + ",'" + classificationType + "')"

126         Gdb.Execute sql
128     Next Node
    
        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Command1_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>6976082345

End Sub

Public Sub printNode(Node As IXMLDOMNode, ADDRECORD As Integer)

        '<EhHeader>
        On Error GoTo printNode_Err

        '</EhHeader>
        Dim xmlNode  As IXMLDOMNode

        Dim xml2Node As IXMLDOMNode

        Dim xml3Node As IXMLDOMNode
     
100     If Node.hasChildNodes Then
       
102         For Each xmlNode In Node.childNodes

104             List1.AddItem xmlNode.nodeName + "*" + xmlNode.Text  ' MARK - ISSUER -COUNTERPART-HEADER
                ' Print xmlNode.nodeName
           
106             If xmlNode.nodeName = "mark" Then MARK = xmlNode.Text
                '<issuer>
           
108             If UCase(xmlNode.nodeName) = "ISSUER" Then

110                 For Each xml2Node In xmlNode.childNodes

112                     If UCase(xml2Node.nodeName) = "VATNUMBER" Then AFM0 = xml2Node.Text
                        'Call printNode(xmlNode)
                    Next

                End If

114             If UCase(xmlNode.nodeName) = "COUNTERPART" Then

116                 For Each xml2Node In xmlNode.childNodes

118                     If UCase(xml2Node.nodeName) = "VATNUMBER" Then AFM = xml2Node.Text
                        'Call printNode(xmlNode)
                    Next

                End If
                
                
                
                
                
                
                
120             If (xmlNode.nodeName) = "invoiceHeader" Then

122                 For Each xml2Node In xmlNode.childNodes

124                     If (xml2Node.nodeName) = "issueDate" Then
                            issueDate = xml2Node.Text
                        End If
126
                        If (xml2Node.nodeName) = "aa" Then
                                   AA = xml2Node.Text
                        End If
128                     If (xml2Node.nodeName) = "invoiceType" Then
                                invoiceType = xml2Node.Text
                                Dim gg As String
                                Select Case invoiceType
                                     Case "1.1"
                                         gg = "G"
                                      Case "5.1"
                                         gg = "D"
                                      Case "5.2"
                                         gg = "D"
                                      Case "9.3"
                                         gg = "E"
                                      Case "2.1"
                                         gg = "Ξ"
                                         
                                End Select
                                AA = gg + Right(Format(Val(AA), "000000"), 6)
                                         
                         End If
                                
                         
                        'Call printNode(xmlNode) invoiceType
                    Next

                End If

130             If (xmlNode.nodeName) = "paymentMethods" Then

132                 For Each xml2Node In xmlNode.childNodes

134                     If (xml2Node.nodeName) = "type" Then paytype = xml2Node.Text
136                     If (xml2Node.nodeName) = "amount" Then payaji = xml2Node.Text
                        'Call printNode(xmlNode)
                    Next

                End If

138             If (xmlNode.nodeName) = "invoiceDetails" Then
                    Dim KODE, ONOMA, MONA As String
                    Dim POSO, TIMM, kau_ajia As Single
                    Dim FPA As String
140                 For Each xml2Node In xmlNode.childNodes

142                     If (xml2Node.nodeName) = "lineNumber" Then
                            'add eggtim line
                        End If
                        
                        'itemCode
                        If xml2Node.nodeName = "itemCode" Then
                            KODE = Replace(xml2Node.Text, "'", "`")
                        End If
                        
                                                
                        'itemDescr,quantity
                         If (xml2Node.nodeName) = "itemDescr" Then
                            'add eggtim line
                            Debug.Print xml2Node.Text
                            ONOMA = Replace(xml2Node.Text, "'", "`")
                        End If
                       If (xml2Node.nodeName) = "quantity" Then
                            'add eggtim line
                            Debug.Print xml2Node.Text
                            POSO = Val(xml2Node.Text)
                        End If
                         If (xml2Node.nodeName) = "measurementUnit" Then
                            'add eggtim line
                            Debug.Print xml2Node.Text
                            MONA = (GGET_CVALUE("select PERIGRAFH FROM PINAKES WHERE TYPOS=2 AND N1=" + xml2Node.Text)) '(xml2Node.Text)
                        End If
                      '  <netValue>14.17</netValue>
       ' <vatCategory>1</vatCategory>
                       If (xml2Node.nodeName) = "netValue" Then
                            'add eggtim line
                            Debug.Print "net" + xml2Node.Text
                            kau_ajia = Val(xml2Node.Text)
                        End If
                         If (xml2Node.nodeName) = "vatCategory" Then
                            'add eggtim line
                            Debug.Print "fpa" + xml2Node.Text
                            FPA = str(GGET_NVALUE("select AYJON FROM PINAKES WHERE TYPOS=1 AND N1=" + xml2Node.Text))
                        End If
                                             
                                             
                                             
                                             
                                             
                       
                    Next
                    If ADDRECORD = 1 Then
                       TIMM = kau_ajia / IIf(Val(POSO) = 0, 1, Val(POSO))
                       Dim Q As String: Q = "insert into APESTALMENADET (FPA,MONA,EIDOS,KAU_AJIA,HME,ATIM,KODE,ONOMA,POSO,TIMM) VALUES(" + FPA + ",'" + MONA + "','r'," + Replace(Format(kau_ajia, "#######0.00"), ",", ".") + ",'" + issueDate + "','" + AA + "','" + KODE + "','" + ONOMA + "'," + Replace(Format(Val(POSO), "####0.00"), ",", ".") + "," + Replace(Format(Val(TIMM), "####0.00"), ",", ".") + ")"
                       Gdb.Execute Q
                   End If

                End If

                'invoiceSummary
144             If (xmlNode.nodeName) = "invoiceSummary" Then

146                 For Each xml2Node In xmlNode.childNodes

148                     If (xml2Node.nodeName) = "totalNetValue" Then totalNetValue = xml2Node.Text
150                     If (xml2Node.nodeName) = "totalVatAmount" Then totalVatAmount = xml2Node.Text
                       
152                     If (xml2Node.nodeName) = "incomeClassification" Then

154                         For Each xml3Node In xml2Node.childNodes
                       
156                             If (xml3Node.nodeName) = "icls:classificationType" Then
158                                 classificationType = xml3Node.Text
                                End If
                       
                            Next
                       
                        End If
                       
                        'Call printNode(xmlNode)
                    Next

                End If
            
                '             <totalNetValue>53.65</totalNetValue>
                '        <totalVatAmount>12.88</totalVatAmount>
                '        <totalWithheldAmount>0</totalWithheldAmount>
                '        <totalFeesAmount>0</totalFeesAmount>
                '        <totalStampDutyAmount>0</totalStampDutyAmount>
                '        <totalOtherTaxesAmount>0</totalOtherTaxesAmount>
                '        <totalDeductionsAmount>0</totalDeductionsAmount>
                '        <totalGrossValue>66.53</totalGrossValue>
                '        <incomeClassification>
                '          <icls:classificationType>E3_561_001</icls:classificationType>
                '          <icls:classificationCategory>category1_1</icls:classificationCategory>
                '          <icls:amount>53.65</icls:amount>
                '        </incomeClassification>
                '      </invoiceSummary>
            
                ' Call print2Node(xmlNode)
160         Next xmlNode

            ' Print node.nodeName
            'List1.AddItem node.nodeName + "*" + node.Text
        
        End If

        '<EhFooter>
        Exit Sub

printNode_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.printNode " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Sub printUSNode(Node As IXMLDOMNode, ByRef AFM As String)

        '<EhHeader>
        On Error GoTo printNode_Err

        '</EhHeader>
        Dim xmlNode  As IXMLDOMNode

        Dim xml2Node As IXMLDOMNode

        Dim xml3Node As IXMLDOMNode
     
100     If Node.hasChildNodes Then
       
102         For Each xmlNode In Node.childNodes

104             List1.AddItem xmlNode.nodeName + "*" + xmlNode.Text  ' MARK - ISSUER -COUNTERPART-HEADER
                ' Print xmlNode.nodeName
           
106             If xmlNode.nodeName = "mark" Then MARK = xmlNode.Text
                '<issuer>
           
108             If UCase(xmlNode.nodeName) = "ISSUER" Then

110                 For Each xml2Node In xmlNode.childNodes

112                     If UCase(xml2Node.nodeName) = "VATNUMBER" Then AFM0 = xml2Node.Text
114                     If UCase(xml2Node.nodeName) = "BRANCH" Then
116                         AFM0 = xml2Node.Text
                        End If

                        'Call printNode(xmlNode)
                    Next

                End If

118             If UCase(xmlNode.nodeName) = "COUNTERPART" Then

120                 For Each xml2Node In xmlNode.childNodes

122                     If UCase(xml2Node.nodeName) = "VATNUMBER" Then AFM = xml2Node.Text
                        'Call printNode(xmlNode)
                    Next

                End If

124             If (xmlNode.nodeName) = "invoiceHeader" Then

126                 For Each xml2Node In xmlNode.childNodes

128                     If (xml2Node.nodeName) = "issueDate" Then issueDate = xml2Node.Text
130                     If (xml2Node.nodeName) = "aa" Then AA = xml2Node.Text
132                     If (xml2Node.nodeName) = "invoiceType" Then invoiceType = xml2Node.Text
                        'Call printNode(xmlNode) invoiceType
                    Next

                End If

134             If (xmlNode.nodeName) = "paymentMethods" Then

136                 For Each xml2Node In xmlNode.childNodes

138                     If (xml2Node.nodeName) = "type" Then paytype = xml2Node.Text
140                     If (xml2Node.nodeName) = "amount" Then payaji = xml2Node.Text
                        'Call printNode(xmlNode)
                    Next

                End If

142             If (xmlNode.nodeName) = "invoiceDetails" Then

144                 For Each xml2Node In xmlNode.childNodes

146                     If (xml2Node.nodeName) = "lineNumber" Then
                            'add eggtim line
                        End If
                       
                    Next

                End If

                'invoiceSummary
148             If (xmlNode.nodeName) = "invoiceSummary" Then

150                 For Each xml2Node In xmlNode.childNodes

152                     If (xml2Node.nodeName) = "totalNetValue" Then totalNetValue = xml2Node.Text
154                     If (xml2Node.nodeName) = "totalVatAmount" Then totalVatAmount = xml2Node.Text
                       
156                     If (xml2Node.nodeName) = "incomeClassification" Then

158                         For Each xml3Node In xml2Node.childNodes
                       
160                             If (xml3Node.nodeName) = "icls:classificationType" Then
162                                 classificationType = xml3Node.Text
                                End If
                       
                            Next
                       
                        End If
                       
                        'Call printNode(xmlNode)
                    Next

                End If
            
                ' Call print2Node(xmlNode)
164         Next xmlNode

            ' Print node.nodeName
            'List1.AddItem node.nodeName + "*" + node.Text
        
        End If

        '<EhFooter>
        Exit Sub

printNode_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.printNode " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Sub print2Node(Node As IXMLDOMNode)

        '<EhHeader>
        On Error GoTo print2Node_Err

        '</EhHeader>
        Dim xmlNode As IXMLDOMNode

100     If Node.hasChildNodes Then

            'List1.AddItem node.childNodes.Item(0)
102         For Each xmlNode In Node.childNodes

104             Call printNode(xmlNode, 0)
106         Next xmlNode

            ' Print node.nodeName
108         List1.AddItem Node.nodeName + "*" + Node.Text
        
        End If

        '<EhFooter>
        Exit Sub

print2Node_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.print2Node " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Sub test()

        '<EhHeader>
        On Error GoTo test_Err

        '</EhHeader>

        Dim objxml, arrNodes, nodesXML, i

100     Set objxml = CreateObject("MSXML2.DOMDocument.6.0")

102     With objxml
104         .setProperty "SelectionLanguage", "XPath"
106         .setProperty "SelectionNamespaces", "xmlns:s='http://www.w3.org/ns/widgets'"
108         .validateOnParse = True
110         .async = False
112         .Load "C:\" + F_TXTFILES + "\apantreq.xml"
        End With

114     arrNodes = Array("/s:widget/s:preference[@name='android-minSdkVersion']/@value", "/s:widget/s:preference[@name='android-versionCode']/@value", "/s:widget/s:preference[@name='android-installLocation']/@value", "/s:widget/s:preference[@name='android-targetSdkVersion']/@value", "/s:widget/s:preference[@name='orientation']/@value", "/s:widget/s:preference[@name='fullscreen']/@value", "/s:widget/@id", "/s:widget/@version", "/s:widget/s:name")

116     For i = LBound(arrNodes) To UBound(arrNodes)
118         Set nodesXML = objxml.documentElement.SelectSingleNode(arrNodes(i))
120         MsgBox nodesXML.Text
        Next

122     Set nodesXML = Nothing: Set objxml = Nothing

        '<EhFooter>
        Exit Sub

test_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.test " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim N As Integer:

100     If F_PAROCHOS = 1 Then
          If MYDPAR.Text = "MYDATA" Then
              N = ToXMLsub(0, "")
          Else
              N = ToJason(0, 1)
          End If
          
102
        Else
104         N = ToXMLsub(0, "") ' 1 means no msgbox
        End If

        'N = ToXMLsub(0)
   
        Dim SQL2 As String

106     SQL2 = FIND_QUERY()
108     Adodc1.RecordSource = SQL2
110     Adodc1.Refresh
   
        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Command2_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Function FIND_QUERY()

        '=================================================================================================
        '<EhHeader>
        On Error GoTo FIND_QUERY_Err

        '</EhHeader>
        Dim pol As String, polepis As String, ago As String, AGOEPIS As String, PARASPAROXOY As String

100     pol = "": polepis = "": ago = "": AGOEPIS = "": PARASPAROXOY = ""

        'If checkServer(0) Then
        ' MsgBox("OK")
        'End If
102     If Year(APO) < 2021 Then
104         MsgBox "ΛΑΘΟΣ ΗΜΕΡΟΜΗΝΙΑ"
106         FIND_QUERY = ""
108         Unload Me
           
        End If

110     Dim DUM As Boolean: DUM = Get_AJ_ASCII(pol, polepis, ago, AGOEPIS, PARASPAROXOY)

        Dim par

112     If Len(pol) = 0 Then
114         MsgBox "Δεν έχουν επιλεχθεί παραστατικά για MyData" + Chr(13) + "Παράμετροι παραστατικών->MyData->Ενημερώνει MyData "
116         FIND_QUERY = "0"

            Exit Function

        End If

118     If Len(polepis) = 0 Then
120         par = pol
        Else
122         par = pol + "," + polepis 'pol + polepis
        End If
        
124     If Len(ago) > 0 Then
126         par = par + "," + ago
        End If

        Dim sql As String
        
128     If Len(where.Caption) < 1 Then

130         If CHXORISERR.Value = vbChecked Then
        
132             fSynt = " WHERE ( (isnull(ENTITYMARK,'')='') OR (INCMARK='-' AND AKYROMENO=1) ) AND    LEFT(ATIM,1) IN     (  " + par + "  )   "

            Else
134             fSynt = " WHERE ( isnull(ENTITYMARK,'')='' OR LEFT(ENTITYMARK,3)='ERR'  OR (INCMARK='-' AND AKYROMENO=1)     ) AND    LEFT(ATIM,1) IN     (  " + par + "  )   "
            End If
        
136         fSynt = fSynt + " and HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<'" + Format(DateAdd("d", 1, EOS.Value), "MM/dd/yyyy") + "'    AND ISNULL(AJ1,0)+ISNULL(AJ2,0)+ISNULL(AJ3,0)+ISNULL(AJ4,0)+ISNULL(AJ5,0)+ISNULL(AJ6,0)+ISNULL(AJ7,0) >=0  "
    
138         If Check1.Value = vbChecked Then
                'where.Caption =
140             fSynt = fSynt + " and ID_NUM=" + str(Adodc1.Recordset("id_num"))
142             Check1.Value = vbUnchecked
            End If
    
144         If ektos.Value = vbChecked Then
                'where.Caption =
146             fSynt = fSynt + " and ID_NUM<>" + str(Adodc1.Recordset("id_num"))
148             ektos.Value = vbUnchecked
            End If
    
150         If F_PAROCHOS = 1 Then
                 If Len(PARASPAROXOY) = 0 Then PARASPAROXOY = "''"
                'where.Caption =
152             If MYDPAR.Text = "MYDATA" Then
154                 fSynt = fSynt + " and  LEFT(ATIM,1) Not In (" + PARASPAROXOY + ") "
                Else
156                 fSynt = fSynt + " and  LEFT(ATIM,1)  In (" + PARASPAROXOY + ") "
                End If
                
            End If
    
        Else
158         fSynt = where.Caption
        End If
    
160     sql = "SELECT ENTITYMARK, HME,ATIM,PEL.EPO,isnull(PEL.AFM,'000000000') as AFM,KPE,isnull(PEL.DIE,'') AS DIE,ISNULL(PEL.STREETNUMBER,'0') AS PELSTREETNUMBER,ISNULL(PEL.BRANCH,'0') AS PELBRANCH, ISNULL(PEL.CONTRACTINGAUTHORITYID,'') AS CONTRACTINGAUTHORITYID,ISNULL(PEL.CONTRACTIDENTIFIER,'') AS CONTRACTIDENTIFIER,ENTITYUID,ID_NUM, "

162     sql = sql & "ISNULL(PEL.ORGANIZATIONALUNITNAME,'') AS ORGANIZATIONALUNITNAME,ISNULL(PEL.ORGANIZATIONALUNITCODE,'') AS ORGANIZATIONALUNITCODE,ISNULL(PEL.BUDTYPE,'3') AS BUDTYPE,ISNULL(PEL.BUDIDENTIFIER,'') AS BUDIDENTIFIER, "
164     sql = sql + "isnull(PEL.CH1,'') AS DEH,ISNULL(AJ1,0) AS AJ1,ISNULL(AJ2,0) AS AJ2 ,ISNULL( AJ3,0 ) AS AJ3,ISNULL(AJ4,0) AS AJ4,ISNULL(AJ5,0) AS AJ5,ISNULL(AJ6,0) AS AJ6,ISNULL(AJ7,0) AS AJ7,"
166     sql = sql + "ISNULL(FPA1,0) AS FPA1,ISNULL(FPA2,0) AS FPA2,ISNULL(FPA3,0) AS FPA3,ISNULL(FPA4,0) AS FPA4 ,ISNULL(FPA6,0) AS FPA6,ISNULL(FPA7,0) AS FPA7,"
168     sql = sql + "AJI,PEL.EPA,ISNULL(PEL.POL,'.') AS POL,PEL.COUNTRY,TRP,ISNULL(APALAGIFPA,0) AS APALAGIFPA ,"
170     sql = sql + "ISNULL(PEL.XRVMA,'') AS TK,SXETMARK,ISNULL(PARAKRATISI,0) AS PARAKRAT,ISNULL(B_C2,'') AS TYPOSPARAKRAT,isnull(EXCHANGERATE,1) AS EXCHANGERATE,isnull(CURRENCY,'EUR') AS CURRENCY ,ISNULL(ELGA,0) AS ELGA,isnull(KR1,0) AS KR1,ISNULL(KR2,0) AS KR2,PEL.EIDOS AS EIDOS,isnull(ANASTOLHFPA,0) as ANASTOLHFPA ,OTHERMOVEPURPOSETITLE, ORA,AYTOK,ISNULL(SXETMARKS,',') AS SXETMARKS,SKOPOS,AKYROMENO,isnull(JWT,'') AS JWT, "
        sql = sql + "STREET,TIM.STREETNUMBER,POSTALCODE,CITY,ISNULL(TIM.BRANCH,'0') AS BRANCH,ISNULL(HMEPARAD,HME) AS HMEPARAD,ISNULL(ORAPARAD,'00:00') AS ORAPARAD,FORTOSH "
172     sql = sql + "   FROM TIM LEFT JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
174     sql = sql + fSynt ' " WHERE (ENTITYMARK IS NULL OR LEFT(ENTITYMARK,3)='ERR' ) AND    LEFT(ATIM,1) IN     (  " + par + "  )    and HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(DateAdd("d", 1, EOS.Value), "MM/dd/yyyy") + "'  "
        'sql = sql + "  AND AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7>0  " + synt
176     sql = sql + " order by HME,ATIM"       '  OR INCMARK IS NULL OR INCMARK='ERROR'

178     FIND_QUERY = sql
    
        'soyma kata kathgoria kai kata fpa
        'SELECT SUM(KAU_AJIA) AS AJIA,G.FPA,KATHGORIA  FROM EGGTIM G  INNER JOIN EID D  ON G.KODE=D.KOD
        '  WHERE ATIM='T000592' GROUP BY G.FPA,KATHGORIA
    
        '    AJIA    FPA KATHGORIA
        '     100     1   1
        '     10      2   1
        '     20      1   2
        '    400      2   2
        '
        '
    
        '<EhFooter>
        Exit Function

FIND_QUERY_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.FIND_QUERY " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Public Function ToXMLsub(ByVal noask As Integer, ByRef markReal As String) As Integer ' 1=no messbog
     
   Dim m_is_benzinadiko As Integer
     m_is_benzinadiko = Val(FINDPARAMETROI(1, "PAR1", "f_is_benzinadiko", "0", "Γ.34.EINAI BENZINAΔIKO=1 ΟΧΙ=0"))  'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))

        '<EhHeader>
        On Error GoTo ToXMLsub_Err

        '</EhHeader>

        Dim R8     As New ADODB.Recordset, CR8 As String, CR8B As String
       
        Dim SUMNET As Double

        '===?G??O ?? XML G?? ?? ????S?????? =================================================================================
        'WHERE (ENTITYMARK IS NULL OR ENTITYMARK='ERROR' ) AND
        'Left(ATIM, 1) In     (  " + PAR + "  )    And
        'HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
        'sxetiko   <correlatedInvoices>400000017716190</correlatedInvoices>
        '  <invoiceHeader>                       σχετιζομενο
        '            <series>A</series>
        '            <aa>101</aa>
        '            <issueDate>2020-04-08</issueDate>
        '            <invoiceType>5.1</invoiceType>
        '            <currency>EUR</currency>
        '            <correlatedInvoices>400000017716190</correlatedInvoices>
        '    </invoiceHeader>
    
        '    <invoiceHeader>                       αυτοτιμολογηση
        '            <series>A</series>
        '            <aa>101</aa>
        '            <issueDate>2020-04-08</issueDate>
        '            <invoiceType>1.1</invoiceType>
        '            <currency>EUR</currency>
        '            <selfPricing>true</selfPricing>
        '        </invoiceHeader>
    
100     ToXMLsub = 1

        Dim ctypos As String, cTyposExod As String
   
        Dim sql    As String

102     sql = FIND_QUERY()
    
104     fLong = Dt2Lng(Now)

106     Gdb.Execute "UPDATE TIM SET ENTITY=0,ENTLINEN=" + str(fLong) + fSynt
    
        'ΕΑΝ ΕΙΝΑΙ ΠΡΟΒΛΗΜΑΤΙΚΟ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ ΑΠΟΘΗΚΕΥΕΙ ΤΟ ΠΡΟΒΛΗΜΑ ΣΤΟ ENTITYUID
        'KAI BAZEI ENTITYΜΑΡΚ="ERROR-ΑΝΕΥ ΑΠΟΣΤΟΛΗΣ"
        Dim sqlDtemp As New ADODB.Recordset

108     sqlDtemp.Open sql, Gdb, adOpenDynamic, adLockOptimistic

        Dim mProbl    As String, mTypPar As String

        Dim SYN_PROBL As String
        
110     Do While Not sqlDtemp.EOF
112         mProbl = ""

            ' AN EINAI AKYRVMENO KAI EXEI MARK NA STEILEI TO AKYRVTIKO
114         If Val(Left(CNull(sqlDtemp!entitymark) + " ", 1)) > 1 Or Left(CNull(sqlDtemp!entitymark) + " ", 1) = "Δ" Then
116             If sqlDtemp!AKYROMENO = 1 Then
118                 If Len(where.Caption) > 0 Then  ' ONLINE ΔΕΝ ΕΧΡΕΙΑΖΕΤΑΙ GRID
                        ' TIPOTA DEN KANV
                    Else

                        Dim CCCL As String

120                     CCCL = CANCEL_INVOICE(IIf(Left(CNull(sqlDtemp!entitymark) + " ", 1) = "Δ", mID(sqlDtemp!entitymark, 4, 15), sqlDtemp!entitymark), sqlDtemp!ATIM)    ' CANCEL_INVOICE(sqlDtemp!ENTITYMARK, sqlDtemp!ATIM)

122                     If Len(CNull(CCCL)) > 1 Then
124                         Gdb.Execute "UPDATE TIM SET INCMARK='" + CCCL + "' WHERE ID_NUM=" + str(sqlDtemp!id_num)
                        End If
                    End If
                End If

            Else

126             If isDiakin = 2 Then
                Else

128                 If checkIntegrity(i, sqlDtemp, mProbl) = False Then
130                     If mProbl = "" Then
132                         mProbl = "ΠΡΟΒΛΗΜΑ ΑΚΕΡΑΙΟΤΗΤΑΣ " + sqlDtemp("ATIM")
                        End If
                    End If
                End If

134             ctypos = FINDTYPOS(mID(sqlDtemp("ATIM"), 1, 1))
                  
136             mTypPar = mID(Split(ctypos, ";")(0), 1, 3)
                    
138             If Left(mTypPar, 2) = "11" Then
                    'lianikh den xreiazetai pelaths
                Else

140                 If sqlDtemp("COUNTRY") <> "GR" Then
                    Else

142                     If Len(Trim(sqlDtemp("AFM"))) < 9 Then
144                         mProbl = "124Β.ΔEN EXEI ΑΦΜ "
                        Else

                            'αν δεν ειναι ενδοκοινοτικό ή τρίτων χωρών?
146                         If InStr("1.2  1.3 2.2 2.3  2.4 ", mTypPar) = 0 Then
148                             If check_afm(Trim(sqlDtemp("AFM"))) = 0 Then
150                                 mProbl = "130Β. EXEI ΛΑΘΟΣ ΑΦΜ "
                                End If
                            End If
                            
                        End If
                    End If
                End If
                    
152             If mTypPar = "5.1" Then   ' -------- συσχετιζομενο
154                 If IsNull(sqlDtemp("SXETMARK")) Then
156                     mProbl = "ΔΕΝ EXEI ΣΧΕΤΙΚΟ ΠΑΡΑΣΤΑΤΙΚΟ "
                    Else

158                     If Len(sqlDtemp("SXETMARK")) < 10 Then
160                         mProbl = "ΔΕΝ EXEI ΣΧΕΤΙΚΟ ΠΑΡΑΣΤΑΤΙΚΟ "
                        End If
                    End If
                End If
                    
162             isDiakin = GGET_NVALUE("SELECT ISNULL(ISDIAKIN,0) FROM PARASTAT WHERE EIDOS='" + Left(sqlDtemp("ATIM"), 1) + "'")
                    
164             If isDiakin > 0 Then
166                 If Len(Trim(sqlDtemp("TK"))) < 3 Then
168                     mProbl = "ΔΕΝ EXEI TK   " + sqlDtemp("EPO")
                    End If
                
                End If
                    
170             If Len(Trim(mProbl)) > 2 Then
172                 Gdb.Execute "UPDATE TIM SET ENTITYMARK='ERROR-ΠΡΙΝ ΑΠΟΣΤ', ENTITYUID='" + mProbl + "' where ID_NUM=" + str(sqlDtemp("ID_NUM"))
174                 SYN_PROBL = SYN_PROBL + mProbl + Chr(13)
                End If
                          
            End If
                      
176         sqlDtemp.MoveNext
        Loop
    
        ' εδω καθυστερει στην online ενημερωση (ανα τιμολογιο ενημερωση )
178     If Len(where.Caption) > 1 Then  ' an erxetai apo online mhn kaneis refresh
   
        Else
180         Adodc1.Refresh   ' εδω καθυστερει στην online ενημερωση (ανα τιμολογιο ενημερωση )
    
        End If
    
182     If Len(SYN_PROBL) > 0 Then
184         If noask = 0 Then
186             MsgBox SYN_PROBL
188             MsgBox "ΑΔΥΝΑΤΗ Η ΣΥΝΕΧΙΣΗ"
            End If

            Exit Function

        End If
    
        Dim SQLDT As New ADODB.Recordset

        '  SQL = "SELECT  top 20  AJ1 ,AJ2  from TIM  order by HME"
    
        '==================================================================================
190     SQLDT.Open sql, Gdb, adOpenDynamic, adLockOptimistic

        '==================================================================================
192     If SQLDT.EOF Then
194         If noask = 0 Then
196             MsgBox ("ΔΕΝ ΒΡΕΘΗΚΑΝ ΕΓΓΡΑΦΕΣ")
            End If

198         ToXMLsub = 0

            Exit Function

        End If
     
        Dim IsAytotim As Integer

        Dim varStock  As Variant

        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim invoice   As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement
        
        
          ' ΑΝ ΤΟ ΒΕΝΖΙΝΑΔΙΚΟ ΚΟΒΕΙ ΤΙΜ ΓΙΑ ΛΑΔΙΑ ΝΑ ΓΥΡΙΖΕΙ ΒΕΝΖΙΝΑΔΙΚΟ=0

                                 If m_is_benzinadiko = 1 Then
                                    Dim mm_code As String
                                   
                                     mm_code = GGET_CVALUE("select KODE FROM EGGTIM WHERE ID_NUM=" + str(SQLDT("ID_NUM")))
                                     If InStr("10 11 12 13 14 15 20 21 30 31 32 33 34 35 36 37 38 40 41 42 43 44 50 60 61 70 71 72", Trim(mm_code)) > 0 Then      'μονο στου βενζιναδικου να βγαζει fuelcode
                                         m_is_benzinadiko = 1
                                    Else
                                         m_is_benzinadiko = 0
                                    End If
                                 End If
        
    
        '<InvoicesDoc   xmlns="http://www.aade.gr/myDATA/invoice/v1.0"
        ' xsi:schemaLocation="http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd"
        ' xmlns:N1="https://www.aade.gr/myDATA/incomeClassificaton/v1.0"
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
200     Set docStock = New MSXML2.DOMDocument

202     With docStock
204         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
206         Set elemRoot = .createElement("InvoicesDoc")

208         With elemRoot '/////////////////////////////////////////////////////////////////////////////////////////
210             .setAttribute "xmlns", "http://www.aade.gr/myDATA/invoice/v1.0"
212             .setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd"
214             .setAttribute "xmlns:n1", "https://www.aade.gr/myDATA/incomeClassificaton/v1.0"
216             .setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
218             .setAttribute "xmlns:ecls", "https://www.aade.gr/myDATA/expensesClassificaton/v1.0"
            
                '
                '    '        <?xml version="1.0" encoding="UTF-8"?>
                ''<InvoicesDoc xmlns="http://www.aade.gr/myDATA/invoice/v1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                'xsi:            schemaLocation = "http://www.aade.gr/myDATA/invoice/v1.0/InvoicesDoc-v0.6.xsd"
                'xmlns:            icls = "https://www.aade.gr/myDATA/incomeClassificaton/v1.0"
                'xmlns:            ecls = "https://www.aade.gr/myDATA/expensesClassificaton/v1.0"
            
                'Dim I As Long:
220             i = -1

222             Dim nbb As Long: For nbb = 1 To 5000: F_ID_NUMS(nbb) = 0: Next ' ΑΠΟΘΗΚΕΥΕΙ ΤΑ ID_NUM ΓΙΑ ΝΑ ΤΑ ΘΥΜΑΤΑΙ ΟΤΑΝ ΕΡΘΕΙ Η ΑΠΑΝΤΗΣΗ ΑΠΟ ΑΑΔΕ

                'Dim isDiakin As Integer
224             isDiakin = GGET_NVALUE("SELECT ISNULL(ISDIAKIN,0) FROM PARASTAT WHERE EIDOS='" + Left(SQLDT("ATIM"), 1) + "'")

226             Do While Not SQLDT.EOF
            
228                 i = i + 1
               
230                 f_name_xml = SQLDT("atim")
                    
                    '  If checkIntegrity(i, sqlDt) = False Then
                    '     MsgBox (" ΠΡΟΒΛΗΜΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + sqlDt(0)("ATIM").toString)
                    '  End If
               
                    '  If checkIntegrity(I) = False Then

                    '    MsgBox (" ΠΡΟΒΛΗΜΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + sqldt(0)("ATIM").ToString)
                    '  End If

                    Dim EGGTIM As New ADODB.Recordset

232                 Me.Caption = "ΠΑΡΑΣΤΑΤΙΚΑ " + str(i + 1)
                
234                 List1.AddItem "Δημιουργήθηκαν  " + str(i + 1) + " παραστατικά "
                    'EGGTIM.Open "SELECT KODE,POSO,TIMM,EKPT,FPA,ISNULL(KAU_AJIA,0) AS KAU_AJIA,ISNULL(MIK_AJIA,0) AS MIK_AJIA FROM EGGTIM WHERE TIMM<>0 AND POSO<>0 AND ID_NUM=" + str(sqlDt("ID_NUM")), Gdb, adOpenDynamic, adLockOptimistic
                
236                 Dim C56       As String: C56 = ""
                    
                    Dim exei_axia As String

                    'ctypos = FINDTYPOS(mID(sqlDt("ATIM"), 1, 1))

                    'DA Ή ΤΕΛΟΣ ΑΝΘΕΚΤΙΚΟΤΗΤΑΣ
238                 If isDiakin = 2 Or Split(FINDTYPOS(mID(SQLDT("ATIM"), 1, 1)), ";")(0) = "8.2" Then
240                     exei_axia = ""
                    Else
242                     exei_axia = " and TIMM<>0 AND ISNULL(KAU_AJIA,0)<>0 "
                    End If

                    ' Gdb.Execute "UPDATE EGGTIM SET FPAROW=ROUND(MIK_AJIA,2)-ROUND(KAU_AJIA,2) WHERE ID_NUM=" + str(sqlDt("ID_NUM"))

244                 If fMydataFromEID1 = 1 Then  'ΕΝΗΜΕΡΩΣΗ ΑΠΟ ΚΑΤΗΓΟΡΙΑ ΕΙΔΩΝ
                        '  SELECT SUM(KAU_AJIA) AS AJIA,G.FPA,KATHGORIA  FROM EGGTIM G  INNER JOIN EID D  ON G.KODE=D.KOD
                        '  WHERE ATIM='T000592' GROUP BY G.FPA,KATHGORIA
246                     C56 = "SELECT D.DISCOUNTOPTION,G.OTHERMEASUREMENTUNITQUANTITY , G.OTHERMEASUREMENTUNITTITLE,G.KODE,G.POSO,ISNULL(G.MIKTAKILA,0) AS MIKTAKILA,G.ONOMA,G.MONA,G.ID, CONVERT(INTEGER,KATHGORIA) AS KATHGORIA, CONVERT(INTEGER,G.FPA) AS FPA,ROUND(SUM(ISNULL(KAU_AJIA,0)),2) AS KAU_AJIA,ROUND(SUM(ISNULL(MIK_AJIA,0)),2) AS MIK_AJIA, ISNULL(D.APALLFPA,0) AS APAL "
248                     C56 = C56 + " FROM EGGTIM G  INNER JOIN EID D  ON G.KODE=D.KOD "
250                     C56 = C56 + " WHERE  POSO<>0  " + exei_axia + "  and ID_NUM=" + str(SQLDT("ID_NUM"))
252                     C56 = C56 + " GROUP BY G.ID,CONVERT(INTEGER,G.FPA),CONVERT(INTEGER,KATHGORIA), ISNULL(D.APALLFPA,0),G.KODE,G.POSO,MIKTAKILA,G.ONOMA,G.MONA,G.OTHERMEASUREMENTUNITQUANTITY , G.OTHERMEASUREMENTUNITTITLE,D.DISCOUNTOPTION"   ' G.FPA,KATHGORIA,APALLFPA "
254                     EGGTIM.Open C56, Gdb, adOpenDynamic, adLockOptimistic
                    Else   ' ΜΟΝΟ ΑΠΟ ΠΑΡΑΣΤΑΤΙΚΟ
256                     C56 = "SELECT D.DISCOUNTOPTION,G.OTHERMEASUREMENTUNITQUANTITY , G.OTHERMEASUREMENTUNITTITLE,G.KODE,G.POSO,ISNULL(G.MIKTAKILA,0) AS MIKTAKILA,G.ONOMA,G.MONA,G.ID,'1' AS KATHGORIA,CONVERT(INTEGER,G.FPA) AS FPA,ROUND(SUM(ISNULL(KAU_AJIA,0)),3) AS KAU_AJIA, "
258                     C56 = C56 + " ROUND(SUM(ISNULL(MIK_AJIA,0)),2) AS MIK_AJIA, ISNULL(D.APALLFPA,0) AS APAL  "
260                     C56 = C56 + " FROM EGGTIM  G  INNER JOIN EID D  ON G.KODE=D.KOD WHERE POSO<>0 " + exei_axia + " AND ID_NUM=" + str(SQLDT("ID_NUM")) + " GROUP BY G.ID,CONVERT(INTEGER,G.FPA), ISNULL(D.APALLFPA,0),G.KODE,G.POSO, MIKTAKILA,G.ONOMA,G.MONA ,G.OTHERMEASUREMENTUNITQUANTITY , G.OTHERMEASUREMENTUNITTITLE,D.DISCOUNTOPTION " ' G.FPA,APALLFPA "
262                     EGGTIM.Open C56, Gdb, adOpenDynamic, adLockOptimistic
                    
                    End If
                
                    ' ελεγχω αν ησ ουμα των κατηγοριών είναι ίδια με την σουμα του τιμολογιου
                    Dim suma_kathg As Double

264                 suma_kathg = 0

                    Dim SUM_MIK_AJIA As Single

                    Dim SUM_MIK2     As Double

266                 SUM_MIK2 = 0

                    Dim SUMESFPASEIRON As Single
                    Dim PROTO          As Long: PROTO = 0 ' ΤΟ ΠΡΟΤΟ ΕΓΓΤΙΜ ΓΙΑ ΝΑ PAREI THN DIAFORA STO KAU_AJIA
268                 Do While Not EGGTIM.EOF
                        If PROTO = 0 Then PROTO = EGGTIM!ID
270                     suma_kathg = suma_kathg + EGGTIM!kau_ajia
272                     SUM_MIK2 = SUM_MIK2 + EGGTIM!MIK_AJIA

                        SUMESFPASEIRON = SUMESFPASEIRON + (EGGTIM!MIK_AJIA - EGGTIM!kau_ajia)

274                     EGGTIM.MoveNext
                    Loop

276                 SUM_MIK2 = GGET_NVALUE("SELECT SUM(MIK_AJIA) FROM EGGTIM WHERE ID_NUM=" + str(SQLDT("ID_NUM")))

278                 EGGTIM.MoveFirst
                
                    'Dim DUM As New DataTable
280                 F_ID_NUMS(i + 1) = SQLDT("ID_NUM") ' ' ΑΠΟΘΗΚΕΥΕΙ ΤΑ ID_NUM ΓΙΑ ΝΑ ΤΑ ΘΥΜΑΤΑΙ ΟΤΑΝ ΕΡΘΕΙ Η ΑΠΑΝΤΗΣΗ ΑΠΟ ΑΑΔΕ
282                 Gdb.Execute "UPDATE TIM SET ENTITY=" + str(i + 1) + " WHERE ID_NUM=" + str(SQLDT("ID_NUM"))

284                 SUMNET = SQLDT("aj1") + SQLDT("aj2") + SQLDT("aj3") + SQLDT("aj4") + SQLDT("aj5") + SQLDT("aj6") + SQLDT("aj7")
286                 SumFpa = SQLDT("fpa1") + SQLDT("fpa2") + SQLDT("fpa3") + SQLDT("fpa4") + SQLDT("fpa6") + SQLDT("fpa7")

                    ' ΣΥΜΦΩΝΟΥΝ ΟΙ ΚΑΘΑΡΕΣ ΑΞΙΕΣ ?
288                 If Abs(SUMNET - suma_kathg) > 0.0099 Then
290                     If noask = 0 Then
292                         MsgBox ("δεν συμφωνουν οι ΚΑΘΑΡΕΣ αξιες στις κατηγορίες με το παρ/κό " + SQLDT("ATIM")) + " " + Format(SUMNET, "#######0.0000") + "<>" + Format(suma_kathg, "#######0.0000")
294                         MsgBox ("Ακυρώθηκαν όλες οι αποστολές ")
                        End If

                        Exit Function

                    End If
                     
                    Dim DIFF_FPA As Single

                    Dim OK_DIFF  As Integer

296                 OK_DIFF = 1

                    ' ΣΥΜΦΩΝΟΥΝ ΟΙ ΜΙΚΤΕΣ ΑΞΙΕΣ ?
298                 If Abs(SUMNET + SumFpa - SUM_MIK2) > 0.005 Then
                     
300                     If Abs(SUMNET + SumFpa - SUM_MIK2) > 0.03 Then
302                         If noask = 0 Then
304                             MsgBox ("δεν συμφωνουν οι ΜΙΚΤΕΣ αξιες στις κατηγορίες με το παρ/κό " + SQLDT("ATIM")) + " " + str(SUMNET + SumFpa) + "<>" + str(SUM_MIK2)
306                             MsgBox ("Ακυρώθηκαν όλες οι αποστολές ")
                            End If

                            Exit Function

                        Else
308                         DIFF_FPA = SUMNET + SumFpa - SUM_MIK2
310                         OK_DIFF = 0
                        End If
                       
                    End If
                    
                    Dim TESTSUMES As Double
                    TESTSUMES = GGET_NVALUE("SELECT SUM(ROUND(MIK_AJIA,2) -ROUND(KAU_AJIA,2)) FROM EGGTIM WHERE ID_NUM=" + str(SQLDT("ID_NUM")))
                    
                    Dim KL1 As Integer, SUMFPASEIRON As Single
                    
                    'ΕΧΩ ΔΙΑΦΟΡΑ ΣΤΟ TIM.FPA1+TIM.FPA2.. ME TO SUM(ROUND(MIK_AJIA,2) -ROUND(KAU_AJIA,2))
                    ' KAI PREPEI NA RIZV THN DIAFORA STO PROTO RECORD TOY EGGTIM
                    If Abs(SumFpa - SUMESFPASEIRON) > 0.005 And Abs(SumFpa - Round(SUMESFPASEIRON, 2)) < 0.1 Then
                        Dim DIAFORA_APO_TIM As Double
                        DIAFORA_APO_TIM = SumFpa - Round(TESTSUMES, 2)
                    
                        DIAFORA_APO_TIM = Round(DIAFORA_APO_TIM, 2)
                        If DIAFORA_APO_TIM > 0 Then
                            Gdb.Execute "UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + LTrim(str(Abs(DIAFORA_APO_TIM))) + "  WHERE ID=" + str(PROTO)
                        Else
                              
                            Gdb.Execute "UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA-" + LTrim(str(Abs(DIAFORA_APO_TIM))) + "  WHERE ID=" + str(PROTO)
                        
                        End If
                    
                    End If

                    '1.1;E3_561_001;category1_1;0
312                 ctypos = FINDTYPOS(mID(SQLDT("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
                
314                 cTyposExod = FINDEXODTYPOS(mID(SQLDT("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
                
316                 IsAytotim = Val(Trim(Split(ctypos, ";")(3)))
                
318                 If Len(Trim(Split(ctypos, ";")(2))) = 0 And Len(Trim(Split(cTyposExod, ";")(1))) = 0 Then  ' MPOREI TO E3 NA EINAI KENO P.X. 1_95

                        ' If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
                        'writer.Close()
320                     If noask = 0 Then
322                         MsgBox ("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " + SQLDT("ATIM"))
324                         MsgBox ("Ακυρώθηκαν όλες οι αποστολές ")
                        End If

                        Exit Function

                    End If

                    ' writer.WriteComment (sqldt("ATIM") + " " + Format(sqldt("HME"), "dd/MM/yyyy"))
            
                    ' For I = 0 To UBound(varStock)
326                 Set invoice = docStock.createElement("invoice")

328                 With invoice  '----------------------------------

                        ' Dim isDiakin As Integer:
                        'isDiakin = 2

330                     List1.AddItem SQLDT("ATIM") + " " + Format(SQLDT("hme"), "dd/MM/yyyy")
332                    ' Set elemField = docStock.createElement("uid"): elemField.Text = str(i + 1): INVOICE.appendChild elemField
                      
                        Dim IssuerAFM, IssuerCountry, iss
                      
                        Dim m_ypok As String
                      
                        If Month(SQLDT("HME")) < Month(Now) Or Day(SQLDT("HME")) < Day(Now) Or Year(SQLDT("HME")) < Year(Now) Then
                      
                            ' If Day(sqlDt("HME")) < Day(Now) Then
                            Set elemField = docStock.createElement("transmissionFailure"): elemField.Text = "3": .appendChild elemField
                            ' End If
                        End If
                      
334                     If SQLDT("EIDOS") = "r" And InStr("3.1  9.3 ", Split(ctypos, ";")(0)) = 0 And IsAytotim = 0 Then '  IsAytotim = 0  gia na petaei thn αυτοτιμολόγηση στο else
           
                            '------------ ISSUER ------------------------
336                         Set elemField = docStock.createElement("issuer") ' δημιουργω εσοχη
338                         Set elem2Field = docStock.createElement("vatNumber"): elem2Field.Text = Trim(SQLDT("AFM")): elemField.appendChild elem2Field
340                         Set elem2Field = docStock.createElement("country"): elem2Field.Text = SQLDT("COUNTRY"): elemField.appendChild elem2Field
                                     
342                         Set elem2Field = docStock.createElement("branch"): elem2Field.Text = CNull(SQLDT("PELBRANCH")): elemField.appendChild elem2Field

344                         If SQLDT("COUNTRY") <> "GR" Then
                        
346                             Set elem2Field = docStock.createElement("name"): elem2Field.Text = SQLDT("epo"): elemField.appendChild elem2Field
                            End If
                        
348                         Set elem2Field = docStock.createElement("address") ' δημιουργω εσοχ' δημιουργω εσοχ  sqlDT(i)("XRVMA")= TK
350                         Set elem3Field = docStock.createElement("postalCode"): elem3Field.Text = CNull(SQLDT("TK")): elem2Field.appendChild elem3Field
352                         Set elem3Field = docStock.createElement("city"): elem3Field.Text = CNull(SQLDT("POL")): elem2Field.appendChild elem3Field
354                         elemField.appendChild elem2Field
                          
356                         .appendChild elemField
                                
                            '------------ COUNTER PART ------------------------
358                         If mID(Split(ctypos, ";")(0), 1, 2) = "11" Then
                                'lianikh den xreiazetai pelaths
                            Else
                                
360                             Set elemField = docStock.createElement("counterpart") ' δημιουργω εσοχ
362                             Set elem2Field = docStock.createElement("vatNumber"): elem2Field.Text = afmCompany: elemField.appendChild elem2Field
364                             Set elem2Field = docStock.createElement("country"): elem2Field.Text = "GR": elemField.appendChild elem2Field
                         
366                             m_ypok = f_YPOK

368                             If InStr(f_par1ypok, Left(SQLDT("ATIM"), 1)) > 0 Then
370                                 m_ypok = "1"
                                End If

372                             If InStr(f_par2ypok, Left(SQLDT("ATIM"), 1)) > 0 Then
374                                 m_ypok = "2"
                                End If

                                'f_YPOK
                                ' Set elem2Field = docStock.createElement("branch"): elem2Field.Text = "0": elemField.appendChild elem2Field
                                Set elem2Field = docStock.createElement("branch"): elem2Field.Text = m_ypok:
376                             elemField.appendChild elem2Field
  
378                             .appendChild elemField
                 
                            End If
                      
                        Else
                            '------------ ISSUER ------------------------
380                         Set elemField = docStock.createElement("issuer") ' δημιουργω εσοχη
382                         Set elem2Field = docStock.createElement("vatNumber"): elem2Field.Text = afmCompany: elemField.appendChild elem2Field
384                         Set elem2Field = docStock.createElement("country"): elem2Field.Text = "GR": elemField.appendChild elem2Field
                        
386                         m_ypok = f_YPOK

388                         If InStr(f_par1ypok, Left(SQLDT("ATIM"), 1)) > 0 Then
390                             m_ypok = "1"
                            End If

392                         If InStr(f_par2ypok, Left(SQLDT("ATIM"), 1)) > 0 Then
394                             m_ypok = "2"
                            End If

                            'f_YPOK
                            ' Set elem2Field = docStock.createElement("branch"): elem2Field.Text = "0": elemField.appendChild elem2Field
396                         Set elem2Field = docStock.createElement("branch"): elem2Field.Text = m_ypok: elemField.appendChild elem2Field
                          
398                         Dim issuer_country As String: issuer_country = "GR"
                          
400                         If issuer_country <> "GR" Or isDiakin >= 1 Then
                            
402                             Set elem2Field = docStock.createElement("name"): elem2Field.Text = f_pelono: elemField.appendChild elem2Field
                               
                                        
                                        
404                             Set elem2Field = docStock.createElement("address") ' δημιουργω εσοχ'
406                             Set elem3Field = docStock.createElement("street"): elem3Field.Text = f_peldie: elem2Field.appendChild elem3Field
408                             Set elem3Field = docStock.createElement("number"): elem3Field.Text = f_pelarit: elem2Field.appendChild elem3Field

410                             Set elem3Field = docStock.createElement("postalCode"): elem3Field.Text = f_peltk: elem2Field.appendChild elem3Field
412                             Set elem3Field = docStock.createElement("city"): elem3Field.Text = f_pelpol: elem2Field.appendChild elem3Field
414                             elemField.appendChild elem2Field
                            End If
                            
416                         .appendChild elemField
                                
                            '------------ COUNTER PART ------------------------
418                         If mID(Split(ctypos, ";")(0), 1, 2) = "11" And isDiakin = 0 Then
                                'lianikh den xreiazetai pelaths
                            Else
                                
420                             Set elemField = docStock.createElement("counterpart") ' δημιουργω εσοχ
422                             Set elem2Field = docStock.createElement("vatNumber"): elem2Field.Text = Trim(SQLDT("AFM")): elemField.appendChild elem2Field
424                             Set elem2Field = docStock.createElement("country"): elem2Field.Text = SQLDT("COUNTRY"): elemField.appendChild elem2Field

                                Dim pelst2 As String: pelst2 = Trim(SQLDT("PELBRANCH"))
                                If Len(pelst2) = 0 Then
                                    pelst2 = "0"
                                End If

426                             Set elem2Field = docStock.createElement("branch"): elem2Field.Text = pelst2: elemField.appendChild elem2Field

428                             If SQLDT("COUNTRY") <> "GR" Or isDiakin >= 1 Then
                        
430                                 Set elem2Field = docStock.createElement("name"): elem2Field.Text = SQLDT("epo"): elemField.appendChild elem2Field
                                End If
                                
                               ' If m_is_benzinadiko = 1 Then  'PEL.CH1 AS DEH
                                     'supplyAccountNo  deh
                               '      Set elem2Field = docStock.createElement("supplyAccountNo"): elem2Field.Text = CNull(sqlDt("DEH")): elemField.appendChild elem2Field
                              '  End If
                                
                        
432                             Set elem2Field = docStock.createElement("address") ' δημιουργω εσοχ' δημιουργω εσοχ  sqlDT(i)("XRVMA")= TK
                                
434                             Set elem3Field = docStock.createElement("street"): elem3Field.Text = CNull(SQLDT("DIE")): elem2Field.appendChild elem3Field
                                Dim pelst As String: pelst = Trim(SQLDT("PELSTREETNUMBER"))
                                If Len(pelst) = 0 Then
                                    pelst = "0"
                                End If
436                             Set elem3Field = docStock.createElement("number"): elem3Field.Text = pelst: elem2Field.appendChild elem3Field
                                
438                             Set elem3Field = docStock.createElement("postalCode"): elem3Field.Text = CNull(SQLDT("TK")): elem2Field.appendChild elem3Field
                                
440                             Set elem3Field = docStock.createElement("city"): elem3Field.Text = CNull(SQLDT("POL")): elem2Field.appendChild elem3Field
442                             elemField.appendChild elem2Field

444                             .appendChild elemField
                            End If
                        End If

                        '---------------- HEADER---------------------------
446                     Set elemField = docStock.createElement("invoiceHeader") ' δημιουργω εσοχη

                        Dim C_SEIRA As String

448                     C_SEIRA = GET_CVALUE("SELECT SYNT_TITL FROM PARASTAT WHERE EIDOS='" + Left$(SQLDT("ATIM"), 1) + "'")

                        ' If Left(C_SEIRA, 1) = "Τ" Then C_SEIRA = "0" Else C_SEIRA = Trim$(C_SEIRA)
450                     If Len(C_SEIRA) = 0 Then C_SEIRA = "0"

452                     Set elem2Field = docStock.createElement("series"): elem2Field.Text = C_SEIRA: elemField.appendChild elem2Field
454                     Set elem2Field = docStock.createElement("aa"): elem2Field.Text = mID(SQLDT("ATIM"), 2, 6): elemField.appendChild elem2Field
456                     Set elem2Field = docStock.createElement("issueDate"): elem2Field.Text = Format(SQLDT("hme"), "yyyy-MM-dd"): elemField.appendChild elem2Field
458                     Set elem2Field = docStock.createElement("invoiceType"): elem2Field.Text = Split(ctypos, ";")(0): elemField.appendChild elem2Field

460                     If SQLDT("ANASTOLHFPA") = 1 Then
462                         Set elem2Field = docStock.createElement("vatPaymentSuspension"): elem2Field.Text = "true": elemField.appendChild elem2Field
                        End If

464                     If isDiakin = 2 Then
                        Else
466                         Set elem2Field = docStock.createElement("currency"): elem2Field.Text = SQLDT("CURRENCY"): elemField.appendChild elem2Field
                        End If
                        
                        Dim isot As String

468                     If SQLDT("CURRENCY") = "EUR" Then isot = "1" Else isot = Replace(Format(SQLDT("EXCHANGERATE"), "00.00000"), ",", ".")

470                     Set elem2Field = docStock.createElement("exchangeRate"): elem2Field.Text = isot: elemField.appendChild elem2Field

                 
524                     If Split(ctypos, ";")(0) = "5.1" Or Split(ctypos, ";")(0) = "9.1" Then  ' -------- συσχετιζομενο
526                         CR8 = "": CR8B = ""

528                         If IsNull(SQLDT("SXETMARK")) Then
                        
                            Else
                                ' If Len(sqlDt("SXETMARK")) > 10 Then
                                '    CR8 = Split(sqlDt("SXETMARK"), " ")(1)
                                '    CR8B = Split(sqlDt("SXETMARK"), " ")(2)
                                ' End If
                           
530                             If Len(SQLDT("SXETMARK")) > 10 Then
532                                 CR8 = Split(SQLDT("SXETMARK"), "*")(0)
                                    ' CR8B = Split(sqlDt("SXETMARK"), " ")(2)
                                    '  Else
                                 
                                    Dim NSX As Integer

534                                 For NSX = 0 To 10

536                                     If Len(Split(SQLDT("SXETMARK"), ";")(NSX)) > 0 Then
538                                         CR8 = Split(SQLDT("SXETMARK"), ";")(NSX)
540                                         Set elem2Field = docStock.createElement("correlatedInvoices"): elem2Field.Text = CR8: elemField.appendChild elem2Field
                                        Else

                                            Exit For

                                        End If
                                 
                                    Next
                                 
                                End If
                           
                            End If
                        
                                                 
                        End If

                       
                       
                       
                       
                       
472                     If isDiakin >= 1 Then   'tda=1  δελ.αποστ=2
                              If Split(ctypos, ";")(0) = "10.2" Then
                              
                              Else

474                             Set elem2Field = docStock.createElement("dispatchDate"): elem2Field.Text = Format(SQLDT("HMEPARAD"), "yyyy-MM-dd"): elemField.appendChild elem2Field
476                             Set elem2Field = docStock.createElement("dispatchTime"): elem2Field.Text = Trim(SQLDT("ORAPARAD")) + ":00.0000000Z": elemField.appendChild elem2Field
478                             Set elem2Field = docStock.createElement("vehicleNumber"): elem2Field.Text = SQLDT("AYTOK"): elemField.appendChild elem2Field
480                             Set elem2Field = docStock.createElement("movePurpose"): elem2Field.Text = Left(SQLDT("SKOPOS"), 2): elemField.appendChild elem2Field
                             End If


                      If m_is_benzinadiko = 1 Then
                        Dim CODE_MLADIA As String
                        CODE_MLADIA = GGET_CVALUE("SELECT KODE FROM EGGTIM WHERE ID_NUM=" + str(SQLDT("ID_NUM"))) 'ΠΑΙΡΝΩ ΤΟΝ ΚΩΔΙΚΟ ΠΟΥ ΠΟΥΛΗΘΗΚΕ
                        
                        'ΜΟΝΟ ΟΤΑΝ ΕΙΝΑΙ ΚΑΥΣΙΜΑ
                        If InStr("10 11 12 13 14 15 20 21 30 31 32 33 34 35 36 37 38 40 41 42 43 44 50 60 61 70 71 72", CODE_MLADIA) > 0 Then   'μονο στου βενζιναδικου να βαζει fuelInvoice
                            Set elem2Field = docStock.createElement("fuelInvoice"): elem2Field.Text = "true": elemField.appendChild elem2Field
                        End If
                     End If



                            'Δ.Α. για λογαριασμο τριτων
                            If isDiakin = 2 And Val(Left(SQLDT("SKOPOS"), 2)) = 2 Then
                                '                         <otherCorrelatedEntities>
                                '                            <type>6</type>
                                '                            <entityData>
                                '                                 <vatNumber>101875470</vatNumber>
                                '                                 <country>GR</country>
                                '                                 <branch>0</branch>
                                '                                 <name>PAPPOY</name>
                                '                                <address>
                                '                           <street>ΦΤΕΛΙΑ</street>
                                '        <number>0</number>
                                '        <postalCode>66100 </postalCode>
                                '        <city>ΔΡΑΜΑ</city>
                                '      </address>
                                ' </entityData>
                                '</otherCorrelatedEntities>
                                Dim otherCorrelatedEntities As MSXML2.IXMLDOMElement
                                Set otherCorrelatedEntities = docStock.createElement("otherCorrelatedEntities") ' δημιουργω εσοχ'
                                Set elem3Field = docStock.createElement("type"): elem3Field.Text = "6": otherCorrelatedEntities.appendChild elem3Field
                             
                                '------------------------------------------------
                                Set elem2Field = docStock.createElement("entityData") ' δημιουργω εσοχ'
                                Set elem3Field = docStock.createElement("vatNumber"): elem3Field.Text = "000000000": elem2Field.appendChild elem3Field
                                Set elem3Field = docStock.createElement("country"): elem3Field.Text = "GR": elem2Field.appendChild elem3Field
                                Set elem3Field = docStock.createElement("branch"): elem3Field.Text = "0": elem2Field.appendChild elem3Field
                                Set elem3Field = docStock.createElement("name"): elem3Field.Text = "TEST": elem2Field.appendChild elem3Field
                              
                                Set elem3Field = docStock.createElement("address") ' δημιουργω εσοχ' δημιουργω εσοχ  sqlDT(i)("XRVMA")= TK
                                
                                Set elem33Field = docStock.createElement("street"): elem33Field.Text = "street---": elem3Field.appendChild elem33Field
                                Set elem33Field = docStock.createElement("number"): elem33Field.Text = "00": elem3Field.appendChild elem33Field
                                Set elem33Field = docStock.createElement("postalCode"): elem33Field.Text = "66100": elem3Field.appendChild elem33Field
                                Set elem33Field = docStock.createElement("city"): elem33Field.Text = "drama": elem3Field.appendChild elem33Field
                                elem2Field.appendChild elem3Field  'address
                         
                                otherCorrelatedEntities.appendChild elem2Field ' entityData
                                '------------------------------------------------
                                elemField.appendChild otherCorrelatedEntities '  otherCorrelatedEntities
                            End If

                            Dim otherDeli As MSXML2.IXMLDOMElement
                         If Split(ctypos, ";")(0) = "10.2" Then
                              
                         Else
482                         Set otherDeli = docStock.createElement("otherDeliveryNoteHeader") ' δημιουργω εσοχ'
484                         Set elem2Field = docStock.createElement("loadingAddress") ' δημιουργω εσοχ' δημιουργω εσοχ  sqlDT(i)("XRVMA")= TK
                                     Dim ADDRESS1(10) As String, FORT As String
                                     FORT = SQLDT("FORTOSH")
                                     ADDRESS1(0) = f_peldie
                                     ADDRESS1(1) = IIf(Len(f_pelarit) = 0, "0", f_pelarit)
                                     ADDRESS1(2) = f_pelpol
                                     ADDRESS1(3) = f_peltk
                                   If InStr(FORT, ";") > 0 Then ' (Split(ctypos, ";")(0)
                                        ADDRESS1(0) = Split(FORT, ";")(0)
                                        ADDRESS1(1) = Split(FORT, ";")(1)
                                        ADDRESS1(2) = Split(FORT, ";")(2)
                                        ADDRESS1(3) = Split(FORT, ";")(3)
                                   End If
486                         Set elem3Field = docStock.createElement("street"): elem3Field.Text = ADDRESS1(0): elem2Field.appendChild elem3Field
488                         Set elem3Field = docStock.createElement("number"): elem3Field.Text = ADDRESS1(1): elem2Field.appendChild elem3Field
490                         Set elem3Field = docStock.createElement("postalCode"): elem3Field.Text = ADDRESS1(3): elem2Field.appendChild elem3Field
492                         Set elem3Field = docStock.createElement("city"): elem3Field.Text = ADDRESS1(2): elem2Field.appendChild elem3Field
494                         otherDeli.appendChild elem2Field

496                         Set elem2Field = docStock.createElement("deliveryAddress") ' δημιουργω εσοχ' δημιουργω εσοχ  sqlDT(i)("XRVMA")= TK
498                         Set elem3Field = docStock.createElement("street"): elem3Field.Text = CNull(SQLDT("STREET")): elem2Field.appendChild elem3Field
500                         Set elem3Field = docStock.createElement("number"): elem3Field.Text = CNull(SQLDT("STREETNUMBER")): elem2Field.appendChild elem3Field
502                         Set elem3Field = docStock.createElement("postalCode"): elem3Field.Text = CNull(SQLDT("POSTALCODE")): elem2Field.appendChild elem3Field
504                         Set elem3Field = docStock.createElement("city"): elem3Field.Text = CNull(SQLDT("CITY")): elem2Field.appendChild elem3Field
506                         otherDeli.appendChild elem2Field
508                         elemField.appendChild otherDeli
                         End If


                            'If IsAytotim = 1 Then   'tda
                            '     Set elem2Field = docStock.createElement("isDeliveryNote"): elem2Field.Text = "true": elemField.appendChild elem2Field
                            'End If
                            '<isDeliveryNote>*****true</isDeliveryNote>
510                         If isDiakin = 1 Then   'tda=1  δελ.αποστ=2
512                             Set elem2Field = docStock.createElement("isDeliveryNote"): elem2Field.Text = "true": elemField.appendChild elem2Field
                            End If
                             
                             
                              ' αν ειναι στα παραστατικα που ειναι αναποδα
                         If InStr(F_ANAPODO, Left(SQLDT("ATIM"), 1)) > 0 Then
                               Set elem2Field = docStock.createElement("reverseDeliveryNote"): elem2Field.Text = "true": elemField.appendChild elem2Field
                               Set elem2Field = docStock.createElement("reverseDeliveryNotePurpose"): elem2Field.Text = "2": elemField.appendChild elem2Field
                         
                         
                         
                         End If
                         
                                 
                             
                             
                             
                             
                             
                        Else
                            '2800                            Dim NSXETMARKS As Integer, RR2() As String
                            '                            RR2 = Split(sqlDt("SXETMARKS"), ",")
                            '2810                            If Len(sqlDt("SXETMARKS")) > 2 Then ' ΓΙΑΤΙ ΤΟ  NULL ="," ΟΠΟΤΕ ΘΑ ΕΧΕΙ ΚΟΜΑ ΠΑΝΤΑ
                            '                             For NSXETMARKS = 0 To UBound(RR2)
                            '
                            '                                If Len(Split(sqlDt("SXETMARKS"), ",")(NSXETMARKS)) > 0 Then
                            '2830                                      Set elem2Field = docStock.createElement("multipleConnectedMarks"): elem2Field.Text = Split(sqlDt("SXETMARKS"), ",")(NSXETMARKS): elemField.appendChild elem2Field
                            '
                            '
                            '                                Else
                            '                                    Exit For
                            '                                End If
                            '
                            '                             Next
                            '                           End If
                        
                        End If
                        
                        If Val(Left(SQLDT("SKOPOS"), 2)) = 19 Then
                            '<otherMovePurposeTitle>ΛΟΙΠΕΣ ΔΙΑΚΙΝΗΣΕΙΣ</otherMovePurposeTitle>
                            Set elem2Field = docStock.createElement("otherMovePurposeTitle"): elem2Field.Text = SQLDT("OTHERMOVEPURPOSETITLE"): elemField.appendChild elem2Field
                            
                        End If
                                                       
                        '------------- ΑΥΤΟΤΙΜΟΛΟΓΗΣΗ --------------------<selfPricing>true</selfPricing>
514                     If IsAytotim = 1 Then

                            ' ΑΥΤΟΤΙΜΟΛΟΓΗΣΗ ΜΗΛΙΟΥ
                               
516                         If SQLDT("EIDOS") = "r" Then
518                             If Left(mTypPar, 3) = "1.1" Then     ' Left(mTypPar, 3) = "3.1" Or Left(mTypPar, 3) = "13." Then ' TITLOS KTHSHS- αγορα εε
520                                 Set elem2Field = docStock.createElement("selfPricing"): elem2Field.Text = "true": elemField.appendChild elem2Field
                                End If

                            Else
522                             Set elem2Field = docStock.createElement("selfPricing"): elem2Field.Text = "true": elemField.appendChild elem2Field
                            End If
                        End If
                                                       
'524                     If Split(ctypos, ";")(0) = "5.1" Then  ' -------- συσχετιζομενο
'526                         CR8 = "": CR8B = ""
'
'528                         If IsNull(SQLDT("SXETMARK")) Then
'
'                            Else
'                                ' If Len(sqlDt("SXETMARK")) > 10 Then
'                                '    CR8 = Split(sqlDt("SXETMARK"), " ")(1)
'                                '    CR8B = Split(sqlDt("SXETMARK"), " ")(2)
'                                ' End If
'
'530                             If Len(SQLDT("SXETMARK")) > 10 Then
'532                                 CR8 = Split(SQLDT("SXETMARK"), "*")(0)
'                                    ' CR8B = Split(sqlDt("SXETMARK"), " ")(2)
'                                    '  Else
'
'                                    Dim NSX As Integer
'
'534                                 For NSX = 0 To 10
'
'536                                     If Len(Split(SQLDT("SXETMARK"), ";")(NSX)) > 0 Then
'538                                         CR8 = Split(SQLDT("SXETMARK"), ";")(NSX)
'540                                         Set elem2Field = docStock.createElement("correlatedInvoices"): elem2Field.Text = CR8: elemField.appendChild elem2Field
'                                        Else
'
'                                            Exit For
'
'                                        End If
'
'                                    Next
'
'                                End If
'
'                            End If
'
'                            ' CR8B = Split(sqlDt("SXETMARK"), " ")(3)
'
'                            '288                         R8.Open "SELECT ENTITYMARK FROM TIM WHERE ATIM='" + CR8 + "' AND CONVERT(CHAR(10),HME,103)='" + CR8B + "'", Gdb, adOpenDynamic, adLockOptimistic
'                            '290                         CR8 = ""
'                            '292                         If IsNull(R8(0)) Then
'                            '294                             CR8 = ""
'                            '                            Else
'                            '
'                            '296                             If R8.EOF Then
'                            '                                Else
'                            '298                               CR8 = correlatedInvoice
'                            '                                End If
'                            '                            End If
'                            '300                         R8.Close
'                            'CorrelatedInvoices
'
'                            ' Set elem2Field = docStock.createElement("correlatedInvoices"): elem2Field.Text = CR8: elemField.appendChild elem2Field
'
'                        End If

542                     .appendChild elemField
                        
                        '----------------------- PAYMENT ---------------------------------
                        
544                     If isDiakin <= 1 Then  ' if isDiakin =2 then ειναι δελτιο αποστολης οποτε οχι payment
                        
546                         Dim cTrp As String: cTrp = FindTRP(mID(SQLDT("TRP"), 1, 1))

548                         If Len(cTrp) = 0 Then
550                             If noask = 0 Then
552                                 MsgBox ("ΔΕΝ ΕΧΩ ΑΝΤΙΣΤΟΙΧΙΣΗ ΣΤΟΝ ΤΡΟΠΟ ΠΛΗΡΩΜΗΣ " + SQLDT("TRP"))
                                End If

                                Exit Function

                            End If

554                         If SQLDT("EIDOS") = "r" And Left(mTypPar, 3) <> "3.1" And IsAytotim = 0 Then   ' TITLOS KTHSHS Then
                            Else

556                             Set elemField = docStock.createElement("paymentMethods") ' δημιουργω εσοχ
                            
558                             Set elem2Field = docStock.createElement("paymentMethodDetails") ' δημιουργω εσοχ' δημιουργω εσοχ
560                             Set elem3Field = docStock.createElement("type"): elem3Field.Text = cTrp: elem2Field.appendChild elem3Field
562                             Set elem3Field = docStock.createElement("amount"): elem3Field.Text = chDec(Format(SUMNET + SumFpa, "######0.##")): elem2Field.appendChild elem3Field
564                             elemField.appendChild elem2Field
                            
566                             .appendChild elemField
                            End If

                        End If

                        Dim SYN_KAU As Double, SYN_FPA As Double

568                     SYN_KAU = 0
570                     SYN_FPA = 0

                        Dim fpaRow          As Double

                        Dim L               As Integer

572                     Dim SumEsodExod(10) As Double: For L = 1 To 10: SumEsodExod(L) = 0: Next
                     
574                     L = 0

                        '====================================================================================================================
576                     Do While Not EGGTIM.EOF
                            'For n = 1 To 3 ' SEIRES TIMOLOGIOY
578                         L = L + 1

                            Dim AJ As Double

580                         If IsNull(EGGTIM("KAU_AJIA")) Then
582                             AJ = 0
                            Else
584                             AJ = EGGTIM("KAU_AJIA")  ' Math.Round(EGGTIM(L)("POSO") * EGGTIM(L)("TIMM") * (1 - EGGTIM(L)("EKPT") / 100), 2)
                            End If

586                         SumEsodExod(EGGTIM("KATHGORIA")) = SumEsodExod(EGGTIM("KATHGORIA")) + AJ

                            Dim VAT As String

                            '1 ΦΠΑ συντελεστής 24% 24%
                            '2 ΦΠΑ συντελεστής 13% 13%
                            '3 ΦΠΑ συντελεστής 6% 6%
                            '4 ΦΠΑ συντελεστής 17% 17%
                            '5 ΦΠΑ συντελεστής 9% 9%
                            '6 ΦΠΑ συντελεστής 4% 4%
                            '7 Άνευ Φ.Π.Α. 0%
                            '8 Εγγραφές χωρίς ΦΠΑ  (πχ Μισθοδοσία, Αποσβέσεις)
                            
588                         VAT = Format(antFPA(EGGTIM("FPA")), "0")
590                         SYN_KAU = SYN_KAU + AJ

592                         If EGGTIM("FPA") = 5 Then
594                             fpaRow = 0
                            Else
596                             fpaRow = EGGTIM("MIK_AJIA") - EGGTIM("KAU_AJIA")

598                             If fpaRow > 0 And OK_DIFF = 0 Then
600                                 fpaRow = fpaRow + DIFF_FPA
602                                 OK_DIFF = 1
604                                 DIFF_FPA = 0
                                  
                                End If
                                  
                            End If

606                         SYN_FPA = SYN_FPA + fpaRow
                            
                            '
                            '-----------------------------------------------  invoiceDetails
                        
608                         Set elemField = docStock.createElement("invoiceDetails") ' δημιουργω εσοχ
610                         Set elem2Field = docStock.createElement("lineNumber"): elem2Field.Text = str(L): elemField.appendChild elem2Field

612                         If isDiakin >= 1 Then

                                '     <itemCode>***01</itemCode>
                                '        <itemDescr>*****01 - ????? ?4</itemDescr>
                                '        <quantity>****1</quantity>
                                '        <measurementUnit>***1</measurementUnit>
                            
                                Dim monada As Integer

614                             monada = GGET_NVALUE("SELECT ISNULL(TIMH,1) as mm FROM PINAKES WHERE TYPOS=2 AND PERIGRAFH LIKE '%" + EGGTIM("MONA") + "%' ")

616                             If monada = 0 Then monada = 1
618                             Set elem2Field = docStock.createElement("itemCode"): elem2Field.Text = EGGTIM("kode"): elemField.appendChild elem2Field
620                             Set elem2Field = docStock.createElement("itemDescr"): elem2Field.Text = EGGTIM("onoma"): elemField.appendChild elem2Field


                                 If m_is_benzinadiko = 1 Then
                                   If isDiakin = 1 Then
                                     If InStr("10 11 12 13 14 15 20 21 30 31 32 33 34 35 36 37 38 40 41 42 43 44 50 60 61 70 71 72", Trim(EGGTIM("kode"))) > 0 Then   'μονο στου βενζιναδικου να βγαζει fuelcode
                                        Set elem2Field = docStock.createElement("fuelCode"): elem2Field.Text = Trim(EGGTIM("kode")): elemField.appendChild elem2Field
                                     End If
                                   End If
                                 End If



                                'f_SDA = FINDPARAMETROI(1, "PAR1", "F_SDA", "Ω;V", "ΧΑΡΑΚΤΗΡ.ΓΡΑΜΜΑ ΣΔΑ ΕΝΑΡΞΗΣ ΚΑΙ ΕΠΙΣΤΡΟΦΗΣ Π.Χ. Φ;Ψ") '
                                ' EAN EINAI SYGK.DELTIO EPISTROFHS
                                If isDiakin = 2 Then ' DA ΣΥΓΚΕΝΤΡΩΤΙΚΟ ΚΛΠ    'And Split(f_SDA, ";")(1) = Left(sqlDt("ATIM"), 1) Then
                                    Set elem2Field = docStock.createElement("quantity"): elem2Field.Text = Replace(Format(EGGTIM("poso") - EGGTIM("MIKTAKILA"), "######0.000"), ",", "."): elemField.appendChild elem2Field
                                Else
622                                 Set elem2Field = docStock.createElement("quantity"): elem2Field.Text = Replace(Format(EGGTIM("poso"), "######0.000"), ",", "."): elemField.appendChild elem2Field
                                End If

624                             Set elem2Field = docStock.createElement("measurementUnit"): elem2Field.Text = str(monada): elemField.appendChild elem2Field
                                
                             End If   'isDiakin >= 1 Then

                           

                            Dim mVat As String

626                         If isDiakin = 2 Then
628                             mVat = "8"
                            Else
630                             mVat = VAT
                            End If

632                         Set elem2Field = docStock.createElement("netValue"): elem2Field.Text = chDec(Format(AJ, "######0.##")): elemField.appendChild elem2Field
634                         Set elem2Field = docStock.createElement("vatCategory"): elem2Field.Text = mVat: elemField.appendChild elem2Field
636                         Set elem2Field = docStock.createElement("vatAmount"): elem2Field.Text = chDec(Format(fpaRow, "######0.##")): elemField.appendChild elem2Field
                            If IsNull(EGGTIM("DISCOUNTOPTION")) Then
                            
                            Else
                               If nNull(EGGTIM("DISCOUNTOPTION")) = 0 Then
                                    Set elem2Field = docStock.createElement("discountOption"): elem2Field.Text = "false": elemField.appendChild elem2Field
                               Else
                                    Set elem2Field = docStock.createElement("discountOption"): elem2Field.Text = "true": elemField.appendChild elem2Field
                               End If
                            
                            End If
                            
638                         If fpaRow = 0 And isDiakin <> 2 Then  ' στο δα δεν θελει απαλλαγες
                                ' If InStr("1.2  1.3  2.3  2.4", mTypPar) > 0 Then 'ENDOK - TRITON XORON
                                '   Set elem2Field = docStock.createElement("vatExemptionCategory"): elem2Field.Text = "14": elemField.appendChild elem2Field
                                ' Else
                               
640                             Dim mAPAL As String: mAPAL = ""

642                             If SQLDT("APALAGIFPA") > 0 Then ' αν εχω ολικη απαλλαγη φπα να μην ασχολουμαι με τισ σειρες
644                                 mAPAL = str(SQLDT("APALAGIFPA"))
                                Else

646                                 If EGGTIM("apal") = 0 Then mAPAL = str(SQLDT("APALAGIFPA")) Else mAPAL = str(EGGTIM("apal"))
                                End If

648                             If Val(mAPAL) = 99 Then
                                Else
650                                 Set elem2Field = docStock.createElement("vatExemptionCategory"): elem2Field.Text = mAPAL: elemField.appendChild elem2Field
                                End If
                            End If
                           
                            '------------- ΑΥΤΟΤΙΜΟΛΟΓΗΣΗ --------------------<selfPricing>true</selfPricing>
652                         If IsAytotim = 1 Then
                                '------------------------ expenses CLASSIFICATION ------------------------
654                             Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχη
656                             Set elem3Field = docStock.createElement("ecls:classificationType"): elem3Field.Text = Split(cTyposExod, ";")(0): elem2Field.appendChild elem3Field

658                             If fMydataFromEID1 <> 1 Then ' κατηγορία από το παραστατικό

660                                 Set elem3Field = docStock.createElement("ecls:classificationCategory"): elem3Field.Text = Split(cTyposExod, ";")(1): elem2Field.appendChild elem3Field
                                Else
662                                 Set elem3Field = docStock.createElement("ecls:classificationCategory"): elem3Field.Text = fKatEXod(EGGTIM("kathgoria")): elem2Field.appendChild elem3Field
                                    ' fKatEsod(EGGTIM("kathgoria"))
                                End If

664                             Set elem3Field = docStock.createElement("ecls:amount"): elem3Field.Text = chDec(Format(AJ, "######0.##")): elem2Field.appendChild elem3Field
666                             elemField.appendChild elem2Field
                                ' .appendChild elem2Field
                  
                                'fKatEsod(EGGTIM("kathgoria"))
668                             If Left(mTypPar, 3) = "3.1" Then ' TITLOS KTHSHS
                                Else
670                                 Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχη
672                                 Set elem3Field = docStock.createElement("ecls:classificationType"): elem3Field.Text = "VAT_361": elem2Field.appendChild elem3Field
674                                 Set elem3Field = docStock.createElement("ecls:amount"): elem3Field.Text = chDec(Format(AJ, "######0.##")): elem2Field.appendChild elem3Field
676                                 elemField.appendChild elem2Field
                                End If

                            Else
                                '------------------------ INCOME CLASSIFICATION ------------------------
678                             Set elem2Field = docStock.createElement("incomeClassification") ' δημιουργω εσοχ' δημιουργω εσοχ
                                ' Set elem3Field = docStock.createElement("n1:classificationType"): elem3Field.Text = Split(ctypos, ";")(1): elem2Field.appendChild elem3Field

680                             If fMydataFromEID1 <> 1 Then ' κατηγορία από το παραστατικό

682                                 If Len(Trim(Split(ctypos, ";")(1))) = 0 Or isDiakin = 2 Then ' δεν εχει Ε3 π.χ. 1_95 ή ειναι δελτιο αποστολης
                                    Else '-------------------------------  ΠΑΙΡΝΕΙ ΑΠΟ ΠΑΡΑΣΤΑΤΤΙΚΟ
684                                     Set elem3Field = docStock.createElement("n1:classificationType"): elem3Field.Text = Split(ctypos, ";")(1): elem2Field.appendChild elem3Field
                                    End If

686                                 Set elem3Field = docStock.createElement("n1:classificationCategory"): elem3Field.Text = IIf(isDiakin = 2, "category3", Split(ctypos, ";")(2)): elem2Field.appendChild elem3Field
                                Else
                            
688                                 If EGGTIM("kathgoria") = 8 Or Len(Trim(Split(ctypos, ";")(1))) = 0 Or fKatEsod(EGGTIM("kathgoria")) = "category1_95" Then ' εγγυοδοσια Ή ΕΧΕΙ ΚΕΝΟ Ε3 Π.Χ. 1_95 CATEGORY
                                    Else ' 3/4/23 ---------------- ΠΑΙΡΝΕΙ ΤΟ Ε3  ΑΠΟ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ

                                        Dim E3type As String

690                                     If F_E3_APOKATHG_EID = 1 Then E3type = FkatE3Esod(EGGTIM("kathgoria")) Else E3type = Split(ctypos, ";")(1)
692                                     Set elem3Field = docStock.createElement("n1:classificationType"): elem3Field.Text = E3type: elem2Field.appendChild elem3Field
                                    End If

                                    ' κατηγορια απο το "Κατηγορία έιδους"
694                                 Set elem3Field = docStock.createElement("n1:classificationCategory"): elem3Field.Text = IIf(isDiakin = 2, "category3", fKatEsod(EGGTIM("kathgoria"))): elem2Field.appendChild elem3Field

                                End If

696                             Set elem3Field = docStock.createElement("n1:amount"): elem3Field.Text = chDec(Format(AJ, "######0.##")): elem2Field.appendChild elem3Field
698                             elemField.appendChild elem2Field
                            End If
                      
700                         .appendChild elemField
                        
702                         EGGTIM.MoveNext
                            'Next
                        Loop

                        '====================================================================================================================
      
704                     Gdb.Execute "UPDATE TIM SET AADEKAU=" + Replace(Format(SYN_KAU, "######0.#####"), ",", ".") + ",AADEFPA=" + Replace(Format(SYN_FPA, "######0.#####"), ",", ".") + " WHERE ID_NUM=" + str(SQLDT("ID_NUM"))
                       
                        ' AN DEN ΕΣΤΕΙΛΕ ΣΤΟ DETAILS ( SYN_KAU ) ΟΛΟ ΤΟ ΠΟΣΟ ΤΟΥ ΤΙΜΟΛΟΓΙΟΥ (SUMNET)
706                     If Abs(Math.Round(SUMNET, 2) - Math.Round(SYN_KAU, 2)) >= 0.01 Then
                           
708                         Gdb.Execute "UPDATE TIM SET ENTITYMARK='ERROR-ΠΡΙΝ ΑΠΟΣΤ', ENTITYUID='" + mProbl + "' where ID_NUM=" + str(SQLDT("ID_NUM"))
                             
                            ' den εστειλε ολο το τιμολογιο
710                         If noask = 0 Then
712                             MsgBox "456.ΑΣΥΜΦΩΝΙΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + SQLDT("ATIM")
                            End If

                            Exit Function

                        End If

                        '  Exit Function
  
                        Dim TYPOS_FOROY  As Integer  '1=PARAKRA WITHHELD  2=TELH FEES  3=LOIPOI OTHER 4=XARTOS STAMP (5=KRATHSEIS(DEN TO YPOSTHRIZO 27-10-2021)

                        Dim TYPOS_2FOROY As Integer  '1=PARAKRA WITHHELD  2=TELH FEES  3=LOIPOI OTHER 4=XARTOS STAMP (5=KRATHSEIS(DEN TO YPOSTHRIZO 27-10-2021)
  
714                     TYPOS_FOROY = 0
716                     TYPOS_2FOROY = 0

                        Dim CPOSO_F        As String

                        Dim SYN_KRATIS(10) As Double, K_KR As Integer

718                     For K_KR = 1 To 5: SYN_KRATIS(K_KR) = 0: Next

                        'TYPOSPARAKRAT exei 1 ή 2 φόρουυς / παρακρατησεις
                        'π.χ.  1;4;3;1   σημαινει οτι εχει παρακρατηση(1) και λοιπους φόρους (3)
                        '      1=parakrathsh   4=typos4 parakrathhshw   3=λοιποι φοροι   1=τυπος λοιπου φορου

                        '  '------------------------ taxes CLASSIFICATION ------------------------
                        'ISNULL(PARAKRATISI,0) AS PARAKRAT,ISNULL(B_C2,'') AS TYPOSPARAKRAT "    B_C2=>ΤΥPΟSPΑRΑΚRΑΤ   1;4;3;5(PARAKRATHSH+XARTOSHMO)  , PARAKRATISI=ΑΞΙΑ 1ΗΣ ΠΑΡΑΚΡΑΤΗΣΗΣ , KR1=>2Η ΑΞΙΑ Π.Χ. ΧΑΡΤΟΣΗΜΟΥ
720                     If SQLDT("PARAKRAT") > 0 Then

                            Dim RR() As String

722                         RR = Split(SQLDT("TYPOSPARAKRAT"), ";")
724                         TYPOS_FOROY = Val(RR(0))
726                         SYN_KRATIS(TYPOS_FOROY) = SQLDT("PARAKRAT")
728                         Set elemField = docStock.createElement("taxesTotals") ' δημιουργω εσοχη
730                         Set elem2Field = docStock.createElement("taxes") ' δημιουργω εσοχη
732                         Set elem3Field = docStock.createElement("taxType"): elem3Field.Text = RR(0): elem2Field.appendChild elem3Field

734                         If gVal(RR(1)) > 0 Then
736                             Set elem3Field = docStock.createElement("taxCategory"): elem3Field.Text = RR(1): elem2Field.appendChild elem3Field
                            End If
                            If SYN_KAU > 0 Then
738                             Set elem3Field = docStock.createElement("underlyingValue"): elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): elem2Field.appendChild elem3Field
                            End If
740                         Set elem3Field = docStock.createElement("taxAmount"): elem3Field.Text = chDec(Format(Round(SQLDT("PARAKRAT"), 2), "####0.00")): elem2Field.appendChild elem3Field
742                         elemField.appendChild elem2Field  ' taxes
                            ' .appendChild elemField  '.appendChild elemField

                            ' εχει και 2η παρακρατηση
744                         If UBound(RR) > 2 Then
746                             TYPOS_2FOROY = Val(RR(2))
748                             SYN_KRATIS(TYPOS_2FOROY) = SYN_KRATIS(TYPOS_2FOROY) + SQLDT("KR1")
                                ' Set elemField = docStock.createElement("taxesTotals") ' δημιουργω εσοχη
750                             Set elem2Field = docStock.createElement("taxes") ' δημιουργω εσοχη
752                             Set elem3Field = docStock.createElement("taxType"): elem3Field.Text = RR(2): elem2Field.appendChild elem3Field

754                             If gVal(RR(1)) > 0 Then
756                                 Set elem3Field = docStock.createElement("taxCategory"): elem3Field.Text = RR(3): elem2Field.appendChild elem3Field
                                End If

758                             Set elem3Field = docStock.createElement("underlyingValue"): elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): elem2Field.appendChild elem3Field
760                             Set elem3Field = docStock.createElement("taxAmount"): elem3Field.Text = chDec(Format(Round(SQLDT("KR1"), 2), "####0.00")): elem2Field.appendChild elem3Field
762                             elemField.appendChild elem2Field  ' taxes
                            End If
                           
764                         .appendChild elemField  '.appendChild elemField
 
                        End If
      
                        '------------------------------------------------ InvoiceSummary
                        
766                     Set elemField = docStock.createElement("invoiceSummary") ' δημιουργω εσοχ
768                     Set elem2Field = docStock.createElement("totalNetValue"): elem2Field.Text = chDec(Format(Round(SYN_KAU, 2), "######0.##")): elemField.appendChild elem2Field
770                     Set elem2Field = docStock.createElement("totalVatAmount"): elem2Field.Text = chDec(Format(Round(SYN_FPA, 2), "######0.##")): elemField.appendChild elem2Field

                        ' PARAKRATHSH
772                     If SYN_KRATIS(1) > 0 Then CPOSO_F = chDec(Format(Round(SYN_KRATIS(1), 2), "####0.00")) Else CPOSO_F = "0.00"
774                     Set elem2Field = docStock.createElement("totalWithheldAmount"): elem2Field.Text = CPOSO_F: elemField.appendChild elem2Field
                            
                        'totalFeesAmount
776                     If SYN_KRATIS(2) > 0 Then CPOSO_F = chDec(Format(Round(SYN_KRATIS(2), 2), "####0.00")) Else CPOSO_F = "0.00"
778                     Set elem2Field = docStock.createElement("totalFeesAmount"): elem2Field.Text = CPOSO_F:: elemField.appendChild elem2Field

                        'totalStampDutyAmount
780                     If SYN_KRATIS(4) > 0 Then CPOSO_F = chDec(Format(Round(SYN_KRATIS(4), 2), "####0.00")) Else CPOSO_F = "0.00"
782                     Set elem2Field = docStock.createElement("totalStampDutyAmount"): elem2Field.Text = CPOSO_F: elemField.appendChild elem2Field

                        'totalOtherTaxesAmount
784                     If SYN_KRATIS(3) > 0 Then CPOSO_F = chDec(Format(Round(SYN_KRATIS(3), 2), "####0.00")) Else CPOSO_F = "0.00"
786                     Set elem2Field = docStock.createElement("totalOtherTaxesAmount"): elem2Field.Text = CPOSO_F: elemField.appendChild elem2Field
                            
788                     Set elem2Field = docStock.createElement("totalDeductionsAmount"): elem2Field.Text = "0.00": elemField.appendChild elem2Field
                            
                        '  If SYN_KRATIS(1) > 0 Then
                        '      Set elem2Field = docStock.createElement("totalGrossValue"): elem2Field.Text = chDec(Format(SYN_KAU + SYN_FPA - SYN_KRATIS(1), "######0.##")): elemField.appendChild elem2Field
                        ' End If
790                     If SYN_KRATIS(1) + SYN_KRATIS(2) + SYN_KRATIS(3) + SYN_KRATIS(4) + SYN_KRATIS(5) > 0 Then
792                         Set elem2Field = docStock.createElement("totalGrossValue"): elem2Field.Text = chDec(Format(Round(SYN_KAU, 2) + Round(SYN_FPA, 2) - Round(SYN_KRATIS(1), 2) + Round(SYN_KRATIS(2), 2) + Round(SYN_KRATIS(3), 2) + Round(SYN_KRATIS(4), 2) + Round(SYN_KRATIS(5), 2), "######0.##")): elemField.appendChild elem2Field
                        Else
794                         Set elem2Field = docStock.createElement("totalGrossValue"): elem2Field.Text = chDec(Format(Round(SYN_KAU, 2) + Round(SYN_FPA, 2), "######0.##")): elemField.appendChild elem2Field
                        End If

796                     If isDiakin = 2 Then
798                         Set elem2Field = docStock.createElement("incomeClassification") ' δημιουργω εσοχ' δημιουργω εσοχ
800                         Set elem3Field = docStock.createElement("n1:classificationCategory"): elem3Field.Text = "category3": elem2Field.appendChild elem3Field
802                         Set elem3Field = docStock.createElement("n1:amount"): elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): elem2Field.appendChild elem3Field
804                         elemField.appendChild elem2Field

                        Else

806                         If fMydataFromEID1 <> 1 Then '  ΑΠΟ ΠΑΡΑΣΤΑΤΙΚΟ    fKatEsod (EGGTIM("kathgoria"))
                  
808                             If IsAytotim = 1 Then
                                    '------------------------ expenses CLASSIFICATION ------------------------
810                                 Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχη
812                                 Set elem3Field = docStock.createElement("ecls:classificationType"): elem3Field.Text = Split(cTyposExod, ";")(0): elem2Field.appendChild elem3Field
814                                 Set elem3Field = docStock.createElement("ecls:classificationCategory"): elem3Field.Text = Split(cTyposExod, ";")(1): elem2Field.appendChild elem3Field
816                                 Set elem3Field = docStock.createElement("ecls:amount"): elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): elem2Field.appendChild elem3Field
818                                 elemField.appendChild elem2Field

820                                 If Left(mTypPar, 3) = "3.1" Then ' TITLOS KTHSHS
                                    Else
822                                     Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχη
824                                     Set elem3Field = docStock.createElement("ecls:classificationType"): elem3Field.Text = "VAT_361": elem2Field.appendChild elem3Field
826                                     Set elem3Field = docStock.createElement("ecls:amount"): elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): elem2Field.appendChild elem3Field
828                                     elemField.appendChild elem2Field
                                    End If

                                Else '------------------------ income CLASSIFICATION ------------------------
                           
830                                 Set elem2Field = docStock.createElement("incomeClassification") ' δημιουργω εσοχ' δημιουργω εσοχ

832                                 If Len(Trim(Split(ctypos, ";")(1))) = 0 Then  ' P.X 1_95 ΚΕΝΟ Ε3
                              
                                    Else
834                                     Set elem3Field = docStock.createElement("n1:classificationType"): elem3Field.Text = Split(ctypos, ";")(1): elem2Field.appendChild elem3Field
                                    End If
                              
836                                 Set elem3Field = docStock.createElement("n1:classificationCategory"): elem3Field.Text = Split(ctypos, ";")(2): elem2Field.appendChild elem3Field
838                                 Set elem3Field = docStock.createElement("n1:amount"): elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): elem2Field.appendChild elem3Field
840                                 elemField.appendChild elem2Field
                             
                                End If

                            Else  '-------------------------------------------   ΑΠΟ ΚΑΤΗΓΟΡΙΑ ΕΙΔΩΝ

                                Dim K90 As Integer

842                             If IsAytotim <> 1 Then ' oxi aytotimologisi

844                                 For K90 = 1 To 10

846                                     If SumEsodExod(K90) > 0 Then
848                                         Set elem2Field = docStock.createElement("incomeClassification") ' δημιουργω εσοχ' δημιουργω εσοχ

850                                         If fKatEsod(K90) = "category1_95" Then  ' εγγυοδοσια-GIA EJODA TRITVN
                                            Else

                                                ' 3/4/23  ΠΑΙΡΝΕΙ ΤΟ Ε3 ΑΠΟ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ
                                                Dim E3TYPE2 As String

852                                             If F_E3_APOKATHG_EID = 1 Then E3TYPE2 = FkatE3Esod(K90) Else E3TYPE2 = Split(ctypos, ";")(1)
854                                             Set elem3Field = docStock.createElement("n1:classificationType"): elem3Field.Text = E3TYPE2: elem2Field.appendChild elem3Field
                                            End If
                                
                                            ' ΠΑΙΡΝΕΙ ΤΗΝ ΚΑΤΗΓΟΡΙΑ ΑΠΟ ΤΗΝ ΚΑΤΗΓΟΡΙΑ ΕΙΔΟΥΣ
856                                         Set elem3Field = docStock.createElement("n1:classificationCategory")
858                                         elem3Field.Text = fKatEsod(K90)
860                                         elem2Field.appendChild elem3Field
                             
862                                         Set elem3Field = docStock.createElement("n1:amount")
864                                         elem3Field.Text = chDec(Format(SumEsodExod(K90), "######0.##"))
866                                         elem2Field.appendChild elem3Field
                             
868                                         elemField.appendChild elem2Field
                                        End If

                                    Next

                                Else  ' -------------------------- AYTOTIMOLOGISI  -----
                      
870                                 For K90 = 1 To 10

872                                     If SumEsodExod(K90) > 0 Then
874                                         Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχ' δημιουργω εσοχ
876                                         Set elem3Field = docStock.createElement("ecls:classificationType"): elem3Field.Text = Split(cTyposExod, ";")(0): elem2Field.appendChild elem3Field
878                                         Set elem3Field = docStock.createElement("ecls:classificationCategory")
880                                         elem3Field.Text = fKatEXod(K90)
882                                         elem2Field.appendChild elem3Field
                             
884                                         Set elem3Field = docStock.createElement("ecls:amount")
886                                         elem3Field.Text = chDec(Format(SumEsodExod(K90), "######0.##"))
888                                         elem2Field.appendChild elem3Field
                             
890                                         elemField.appendChild elem2Field
                                        End If

                                    Next
                          
892                                 If Left(mTypPar, 3) = "3.1" Then ' TITLOS KTHSHS
                                    Else
894                                     Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχη
896                                     Set elem3Field = docStock.createElement("ecls:classificationType"): elem3Field.Text = "VAT_361": elem2Field.appendChild elem3Field
898                                     Set elem3Field = docStock.createElement("ecls:amount"): elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): elem2Field.appendChild elem3Field
900                                     elemField.appendChild elem2Field
                                    End If
                                End If
                    
                            End If
                      
                        End If 'isDiakin=2
                      
902                     .appendChild elemField
                    
                    End With  '--------------------------------------

904                 .appendChild invoice
                
906                 EGGTIM.Close
                
908                 SQLDT.MoveNext
                Loop

                ' Next
            End With ' /////////////////////////////////////////////////////////////////////////////////////////////

910         Set .documentElement = elemRoot

            On Error Resume Next

912         Kill "C:\" + F_TXTFILES + "\inv.xml"

            '<<<<<<< HEAD
            On Error GoTo ToXMLsub_Err

            ' MAKE_request docStock
            'DOYLEYEI OK ALLA ΥΠΟΨΙΑΖΟΜΑΙ ΟΤΙ ΣΕ ΜΕΓΑΛΑ ΑΡΧΕΙΑ ΤΑ ΜΠΕΡΔΕΥΕΙ
914         FormatXmlDocument docStock ' βαζει κενα να ειναι ευκολο στο διαβασμα

916         Dim ccc As String: ccc = "C:\" + F_TXTFILES + "\inv.xml"

918         .save ccc
        End With

        ' List1.AddItem "Δημιουργήθηκαν " + " απεσταλη"
       ' Dim markReal As String: markReal = ""
920     MAKE_request docStock.XML, 0, noask, "0", markReal
    
922     If Len(where.Caption) > 0 Then  ' ONLINE ΔΕΝ ΕΧΡΕΙΑΖΕΤΑΙ GRID
     
        Else
    
            ' ΠΑΩ ΝΑ ΠΙΑΣΩ ΤΑ ΤΙΜΟΛΟΓΙΑ ΠΟΥ ΤΥΠΩΘΗΚΑΝ ΚΑΙ ΑΚΥΡΩΘΗΚΑΝ ΧΩΡΙΣ ΝΑ ΠΑΡΟΥΝ MYDATA
            ' ΗΔΗ ΕΧΟΥΝ ΠΑΡΕΙ ΜΑΡΚ ΤΑ ΤΙΜΟΛΟΓΙΑ ΑΛΛΑ ΔΕΝ ΠΗΡΑΝ ΜΑΡΚ ΤΑ ΑΚΥΡΩΤΙΚΑ
            Dim sqlDtemp2 As New ADODB.Recordset

924         sqlDtemp2.Open sql, Gdb, adOpenDynamic, adLockOptimistic
    
926         Do While Not sqlDtemp2.EOF

                ' AN EINAI AKYRVMENO KAI EXEI MARK NA STEILEI TO AKYRVTIKO
928             If Val(Left(CNull(sqlDtemp2!entitymark) + " ", 1)) > 1 Or Left(CNull(sqlDtemp2!entitymark) + " ", 1) = "Δ" Then
930                 If sqlDtemp2!AKYROMENO = 1 Then
                              
932                     CCCL = CANCEL_INVOICE(IIf(Left(CNull(sqlDtemp2!entitymark) + " ", 1) = "Δ", mID(sqlDtemp2!entitymark, 4, 15), sqlDtemp2!entitymark), sqlDtemp2!ATIM)

934                     If Len(CNull(CCCL)) > 1 Then
936                         Gdb.Execute "UPDATE TIM SET INCMARK='" + CCCL + "' WHERE ID_NUM=" + str(sqlDtemp2!id_num)
                        End If
                    End If
                End If

938             sqlDtemp2.MoveNext
            Loop

        End If
   
        '<EhFooter>
        Exit Function

ToXMLsub_Err:

        If noask = 0 Then
            MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.ToXMLsub " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        End If

        Resume Next

        '</EhFooter>

End Function

Sub UpdateTIMwithMARK(ByVal v As String, _
                      txmlapost As String, _
                      noask As Integer, _
                      ByVal file_idnum As String, ByRef markReal)

        '<EhHeader>
        On Error GoTo UpdateTIMwithMARK_Err

        '</EhHeader>

100     List1.Clear

        Dim RECS As Integer

        Dim CCV  As String

        Dim C1, c2, c3, c4, CURL, SXOLIA As String

        '
        '    Set objxml = CreateObject("Msxml.DOMDocument")
        '    objxml.async = True
        '    objxml.Load "C:\F_"+F_TXTFILES+"\apantSendinv.xml"
    
        Dim objxml As New MSXML2.DOMDocument

102     objxml.async = True
104     objxml.loadXML v
    
        Dim nLathosApost As Long

106     nLathosApost = 0

        Dim nApost As Long, nAll As Long
        
108     nApost = 0
110     nAll = 0

        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode

        'Gdb.Open "DSN=MERCSQL;"
112     Set nodeList = objxml.SelectNodes("ResponseDoc/response")

114     For Each Node In nodeList

            'List1.AddItem Node.nodeName   ' this works'
            'List1.AddItem Node.nodeName + "*******" + Node.Text  ' INVOICE
        
            Dim r2       As New ADODB.Recordset

            Dim xmlNode  As IXMLDOMNode

            Dim xml2Node As IXMLDOMNode

            Dim xml3Node As IXMLDOMNode
        
116         If Node.hasChildNodes Then
                'List1.AddItem node.childNodes.Item(0)
118             C1 = "": c2 = "": c3 = ""

120             For Each xmlNode In Node.childNodes

                    ' List1.AddItem xmlNode.nodeName
            
122                 If xmlNode.nodeName = "index" Then
124                     C1 = xmlNode.Text
                        ' List1.AddItem c1
                    End If
            
126                 If xmlNode.nodeName = "invoiceUid" Then
128                     c4 = xmlNode.Text
                        'List1.AddItem c2
                    End If
            
130                 If xmlNode.nodeName = "invoiceMark" Then
132                     c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If
                
                    'cURL
134                 If xmlNode.nodeName = "qrUrl" Then
136                     CURL = xmlNode.Text
                        'List1.AddItem c2
                    End If
                
                    'CCV = GET_CVALUE("SELECT CONVERT(CHAR(10),ID_NUM) FROM TIM WHERE ENTLINEN=" + str(fLong) + " and ENTITY=" + C1)
                
                    'NEA EKDOSH ΑΠΟ ΤΟΝ ΠΙΝΑΚΑ F_ID_NUMS POY EXEI TA IDS
138                 CCV = str(F_ID_NUMS(Val(C1)))
                
                    ' αν ερχεται από την επαναποστολή από αρχείο τότε να παίρνει το id_num που πληκτρολόγησα
140                 If Val(file_idnum) > 0 Then
142                     CCV = file_idnum
                    End If

                    'statusCode
144                 If xmlNode.nodeName = "statusCode" Then
146                     If xmlNode.Text = "Success" Then
148                         If F_demo = 1 Then c2 = "'ΔΟΚ" + c2 + "'"
                                      markReal = c2
150                         RECS = GDBEXE("UPDATETIMWITHMARK", "UPDATE TIM SET ENTITY=0,  ENTITYUID='" + c4 + "',ENTITYMARK=" + c2 + " WHERE  ID_NUM=" + CCV)  ' ENTLINEN=" + str(fLong) + " and ENTITY=" + C1, RECS
                          
                            ' On Error GoTo 0
152                         RECS = GDBEXE("UPDATETIMWITHMARK", "UPDATE TIM SET QRURL='" + CURL + "' WHERE  ID_NUM=" + CCV)
                            '                                  Set Image1.Picture = QRCodegenBarcode(CURL)
                            '                                  Picture1.Picture = QRCodegenBarcode(CURL, 0, QRCodegenEcc_LOW, 1, 40, , , 500) ' Edited by Saturnian
                            '                                  Clipboard.Clear ' Edited by Saturnian
                            '                                  Clipboard.SetData QRCodegenBarcode(CURL, 0, QRCodegenEcc_LOW, 1, 40, , , 500) ' Edited by Saturnian
                            '                                  SavePicture Picture1.Picture, "C:\MERCVB\FOTOp.BMP"
                            '
                            '                                  send_to_Access
                           
154                         If Val(file_idnum) > 0 Then
156                             GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET   ENTITYUID='" + c4 + "',ENTITYMARK=" + c2 + " WHERE ID_NUM=" + file_idnum
158                             GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET   QRURL='" + CURL + "'  WHERE ID_NUM=" + file_idnum
                          
                            End If
                            
                            If f_Stelno_daiakin = 1 Then
                                  ENARXI2APOSTOLIS afmCompany, CURL, PAR1.AYTOKINHTO ' SQLDT("AYTOK") 'ark
                            End If
                        
160                         SXOLIA = " OK "
                           
162                         nApost = nApost + 1
                            ' Gdb.Execute "UPDATE TIM SET XMLTEXT='" + DIAXOR(txmlapost, nApost) + "' WHERE ID_NUM=" + CCV
                        Else
164                         RECS = GDBEXE("UPDATETIMWITHMARK", "UPDATE TIM SET  ENTITYMARK='ERROR'   WHERE   ID_NUM=" + CCV)   ' ENTLINEN=" + str(fLong) + " and ENTITY=" + C1                          'sxolia    ,ENTITYUID='" + Left(xmlNode.Text, 40) + "'
166                         SXOLIA = xmlNode.Text + " ************* ΛΑΘΟΣ ****************"
                            markReal = xmlNode.Text
168                         nLathosApost = nLathosApost + 1
                        End If

170                     r2.Open "select * from TIM WHERE   ID_NUM=" + CCV, Gdb, adOpenDynamic, adLockOptimistic

                        'If fromFile = 1 Then
                        'Else
172                     If r2.EOF Then
                        Else

174                         If IsNull(r2!ATIM) Then
                            Else
176                             List1.AddItem r2("ATIM") + "  " + SXOLIA
                            End If
                        End If

178                     r2.Close
                    End If
            
180                 If xmlNode.nodeName = "errors" Then
            
182                     If xmlNode.hasChildNodes Then  'CHILD=ERROR
                            'List1.AddItem XMLnode.childNodes.Item(0)
                  
184                         For Each xml2Node In xmlNode.childNodes  'ERROR
                        
186                             If xml2Node.hasChildNodes Then 'CHILD=MESSAGE

188                                 For Each xml3Node In xml2Node.childNodes

190                                     If xml3Node.nodeName = "message" Then
192                                         SXOLIA = xml3Node.Text
                                         
194                                         SXOLIA = Left(xml3Node.Text, 40)
196                                         SXOLIA = Replace(SXOLIA, "'", "`")


                                            markReal = markReal + Chr(13) + xml3Node.Text

                                            On Error GoTo UpdateTIMwithMARK_Err

198                                         GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET ENTITY=0,ENTITYUID='" + SXOLIA + "'  WHERE   ID_NUM=" + CCV   ' ENTLINEN=" + str(fLong) + " and ENTITY=" + C1
200                                         List1.AddItem "  " + SXOLIA
                                        End If

202                                     If xml3Node.nodeName = "code" Then
204                                         SXOLIA = xml3Node.Text + " ************* ΛΑΘΟΣ ****************"
206                                         List1.AddItem "  " + SXOLIA
                                        End If

                                    Next
                        
                                End If
                        
                                'c2 = xmlNode.Text
                                'List1.AddItem c2
                            Next

                        End If
               
                    End If
            
208                 nAll = nAll + 1

                    On Error Resume Next

210                 GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET XMLTEXT='" + DIAXOR(txmlapost, nAll) + "' WHERE ID_NUM=" + CCV
            
                    On Error GoTo UpdateTIMwithMARK_Err
            
212             Next xmlNode

                ' Print node.nodeName
                ' List1.AddItem Node.nodeName + "*" + Node.Text
        
            End If
        
            ' <response>
            '    <index>1</index>
            '    <invoiceUid>89A3C206EE098D4B2960EBE8161956614DA9B100</invoiceUid>
            '    <invoiceMark>400001827772457</invoiceMark>
            '    <statusCode>Success</statusCode>
            '  </response>
            '
        
            '        <response>
            '    <index>2</index>
            '    <statusCode>ValidationError</statusCode>
            '    <errors>
            '      <error>
            '        <message>Counterpart's country for this invoice type must not be in EU</message>
            '        <code>244</code>
            '      </error>
            '    </errors>
            '  </response>
        
            '
            '        '       Print node.n
            ''        Call printNode(Node)     'here is the problem explained below'
            '
            '    If Len(Trim(payaji)) = 0 Then payaji = "0"
            '    If Len(Trim(totalGrossValue)) = 0 Then totalGrossValue = "0"
            '
            '
            'sql = "INSERT INTO APESTALMENA  ([MARK],[AFM],[ISSUEDATE],[AA],[TYPOS],[PAYTYPE],[PAYAJI],[TOTALNETVALUE],[TOTALVATAMOUNT],[TOTALGROSSVALUE],[CLASSIFICATIONTYPE])"
            'sql = sql + " Values ('" + MARK + "','" + AFM + "','" + issueDate + "','" + AA + "','" + invoiceType + "','" + paytype + "'," + Replace(payaji, ",", ".") + "," + Replace(totalNetValue, ",", ".") + "," + Replace(totalVatAmount, ",", ".") + "," + Replace(totalGrossValue, ",", ".") + ",'" + classificationType + "')"
            '
            ''Gdb.Execute sql
        Next

        '1111
        'DIAXOR txmlapost

214     If nApost = 1 Then
216         GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET XMLTEXT='" + txmlapost + "' WHERE ID_NUM=" + CCV
            ' Gdb.Execute "UPDATE TIM SET XMLMYDATA='" + txmlapost + "' WHERE ID_NUM=" + CCV

        End If

        'GoTo 1111

218     List1.AddItem "Πήραν ΜΑΡΚ " + str(nApost)

220     List1.AddItem "Λάθος , χωρίς  ΜΑΡΚ " + str(nLathosApost)

        '<EhFooter>
        Exit Sub

UpdateTIMwithMARK_Err:

        If noask = 0 Then
            MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.UpdateTIMwithMARK " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        End If

        Resume Next

        '</EhFooter>

End Sub
        
Function takefromxml(ByVal v As String) As String

        On Error GoTo UpdateTIMwithMARK_Err

        '<?xml version="1.0" encoding="utf-8"?>
        '<ResponseDoc xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
        '  <response>
        '    <index>1</index>
        '    <newClientDclID>100000000007417</newClientDclID>
        '    <statusCode>Success</statusCode>
        '  </response>
        '</ResponseDoc>

100     List1.Clear

        Dim RECS As Integer

        Dim CCV  As String

        Dim C1, c2, c3, c4, CURL, SXOLIA As String
   
        Dim objxml As New MSXML2.DOMDocument

102     objxml.async = True
104     objxml.loadXML v
    
        Dim nLathosApost As Long

106     nLathosApost = 0

        Dim nApost As Long, nAll As Long
        
108     nApost = 0
110     nAll = 0

        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode
    
112     Set nodeList = objxml.SelectNodes("ResponseDoc/response")

114     For Each Node In nodeList
        
            Dim r2       As New ADODB.Recordset

            Dim xmlNode  As IXMLDOMNode

            Dim xml2Node As IXMLDOMNode

            Dim xml3Node As IXMLDOMNode
        
116         If Node.hasChildNodes Then
                'List1.AddItem node.childNodes.Item(0)
118             C1 = "": c2 = "": c3 = ""

120             For Each xmlNode In Node.childNodes

                    ' List1.AddItem xmlNode.nodeName
            
122                 If xmlNode.nodeName = "index" Then
124                     C1 = xmlNode.Text
                        ' List1.AddItem c1
                    End If
            
126                 If xmlNode.nodeName = "invoiceUid" Then
128                     c4 = xmlNode.Text
                        'List1.AddItem c2
                    End If
            
130                 If xmlNode.nodeName = "newClientDclID" Then
132                     c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If

                    'clientCorrelationID
134                 If xmlNode.nodeName = "clientCorrelationID" Then
136                     c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If

                    'updatedClientDclID
138                 If xmlNode.nodeName = "updatedClientDclID" Then
140                     c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If
                
                    'cURL
142                 If xmlNode.nodeName = "qrUrl" Then
144                     CURL = xmlNode.Text
                        'List1.AddItem c2
                    End If
                
146                 CCV = str(F_ID_NUMS(Val(C1)))
                
                    ' αν ερχεται από την επαναποστολή από αρχείο τότε να παίρνει το id_num που πληκτρολόγησα
148                 If Val(file_idnum) > 0 Then
150                     CCV = file_idnum
                    End If
            
152                 If xmlNode.nodeName = "errors" Then
            
154                     If xmlNode.hasChildNodes Then  'CHILD=ERROR
                            'List1.AddItem XMLnode.childNodes.Item(0)
                  
156                         For Each xml2Node In xmlNode.childNodes  'ERROR
                        
158                             If xml2Node.hasChildNodes Then 'CHILD=MESSAGE

160                                 For Each xml3Node In xml2Node.childNodes

162                                     If xml3Node.nodeName = "message" Then
164                                         SXOLIA = xml3Node.Text
                                         
166                                         SXOLIA = Left(xml3Node.Text, 40)
168                                         SXOLIA = Replace(SXOLIA, "'", "`")

                                            On Error GoTo UpdateTIMwithMARK_Err

                                            ' GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET ENTITY=0,ENTITYUID='" + SXOLIA + "'  WHERE   ID_NUM=" + CCV   ' ENTLINEN=" + str(fLong) + " and ENTITY=" + C1
170                                         List1.AddItem "  " + SXOLIA
                                        End If

172                                     If xml3Node.nodeName = "code" Then
174                                         SXOLIA = xml3Node.Text + " ************* ΛΑΘΟΣ ****************"
176                                         List1.AddItem "  " + SXOLIA
                                        End If

                                    Next
                        
                                End If

                            Next

                        End If
               
                    End If
            
178                 nAll = nAll + 1

                    On Error Resume Next

                    On Error GoTo UpdateTIMwithMARK_Err

180             Next xmlNode
      
            End If
        
        Next

        '1111
        'DIAXOR txmlapost

182     If nApost = 1 Then
            '  GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET XMLTEXT='" + txmlapost + "' WHERE ID_NUM=" + CCV
            ' Gdb.Execute "UPDATE TIM SET XMLMYDATA='" + txmlapost + "' WHERE ID_NUM=" + CCV

        End If

        'GoTo 1111

184     List1.AddItem "Πήραν ΜΑΡΚ " + str(nApost)

186     List1.AddItem "Λάθος , χωρίς  ΜΑΡΚ " + str(nLathosApost)

188     takefromxml = c2

        '<EhFooter>
        Exit Function

UpdateTIMwithMARK_Err:

        If noask = 0 Then
            MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.UpdateTIMwithMARK " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        End If

        Resume Next

        '</EhFooter>

End Function
        
'        Sub TEST22()
'
'
'        Dim I As Integer
'        Dim sumNet As Single
'        Dim sumFpa As Single
'        Dim ctypos As String
'
'        '======================================= ΠΕΡΠΑΤΑΩ ΤΟ ΤΙΜ ====================================
'        For I = 0 To sqldt.Rows.Count - 1
'
'
'
'            If checkIntegrity(I) = False Then
'
'                MsgBox (" ΠΡΟΒΛΗΜΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + sqldt(0)("ATIM").ToString)
'            End If
'
'
'
'            Dim EGGTIM As New DataTable
'            Me.Text = "ΠΑΡΑΣΤΑΤΙΚΑ " + Str(I)
'            ExecuteSQLQuery("SELECT KODE,POSO,TIMM,EKPT,FPA,ISNULL(KAU_AJIA,0) AS KAU_AJIA,ISNULL(MIK_AJIA,0) AS MIK_AJIA FROM EGGTIM WHERE TIMM<>0 AND POSO<>0 AND ID_NUM=" +SQLDT("ID_NUM").ToString, EGGTIM)
'            Dim DUM As New DataTable
'            ExecuteSQLQuery("UPDATE TIM SET ENTITY=" + Str(i + 1) + " WHERE ID_NUM=" +SQLDT("ID_NUM").ToString, DUM)
'
'
'
'
'            sumNet = sqldt("aj1") + sqldt("aj2") + sqldt("aj3") + sqldt("aj4") + sqldt("aj5") + sqldt("aj6") + sqldt("aj7")
'            sumFpa = sqldt("fpa1") + sqldt("fpa2") + sqldt("fpa3") + sqldt("fpa4") + sqldt("fpa6") + sqldt("fpa7")
'
'            '1.1;E3_561_001;category1_1
'            ctypos = FINDTYPOS(Mid(sqldt(I)("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
'            If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
'                writer.Close()
'                MsgBox ("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " + sqldt("ATIM"))
'                Exit Sub
'
'            End If
'
'
'            writer.WriteComment (sqldt(I)("ATIM") + " " + Format(sqldt(I)("HME"), "dd/MM/yyyy"))
'            writer.WriteStartElement ("invoice")
'            crNode("uid", Str(i + 1), writer)
'            'crNode("mark", "", writer)
'
'
'            '---------------------------------------------εκδοτης
'            writer.WriteStartElement ("issuer")
'            crNode("vatNumber", "028783755", writer)
'            crNode("country", "GR", writer)
'            crNode("branch", "0", writer)
'            ' writer.WriteStartElement("address")
'            ' crNode("postalCode", """66100""", writer)
'            ' crNode("city", """ΔΡΑΜΑ""", writer)
'            ' writer.WriteEndElement() '/address
'            writer.WriteEndElement() '/issuer
'
'            '--------------------------------------------- πελατης
'            If Mid(Split(ctypos, ";")(0), 1, 2) = "11" Then
'                'lianikh den xreiazetai pelaths
'            Else
'                ' End If
'                writer.WriteStartElement ("counterpart")
'                crNode("vatNumber", Trim(sqlDT(i)("AFM")), writer)  ' crNode("vatNumber", "026677115", writer)
'                crNode("country", "GR", writer)
'                crNode("branch", "0", writer)
'                writer.WriteStartElement ("address")
'                crNode("postalCode", """66100""", writer)  'crNode("postalCode", """66100""", writer)
'                crNode("city",SQLDT("POL"), writer)  ' crNode("city", """ΔΡΑΜΑ""", writer)
'                writer.WriteEndElement() ' /address
'                writer.WriteEndElement() ' /counterpart
'            End If
'
'
'            '----------------------------------------------- header
'            writer.WriteStartElement ("invoiceHeader")
'            crNode("series", "0", writer)
'            crNode("aa", Mid(sqlDT(i)("ATIM"), 2, 6), writer)   '  crNode("aa", "15", writer)
'            crNode("issueDate", Format(sqlDT(i)("hme"), "yyyy-MM-dd"), writer) ' crNode("issueDate", "2019-12-15", writer)
'
'
'            'ctypos = FINDTYPOS(Mid(sqlDT(i)("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
'            'If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
'            '    writer.Close()
'            '    MsgBox("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " +SQLDT("ATIM"))
'            '    Exit Function
'
'            'End If
'
'
'            crNode("invoiceType", Split(ctypos, ";")(0), writer)   ' ειδος παραστατικού
'            crNode("currency", "EUR", writer)
'            crNode("exchangeRate", "1.0", writer)
'            writer.WriteEndElement() ' /invoiceHeader
'
'            '          <paymentMethods>
'            '   <paymentMethodDetails>
'            '       <type>3</type>
'            '       <amount>66.53</amount>
'            '   </paymentMethodDetails>
'            '</paymentMethods>
'            '----------------------------------------------- paymentMethods
'            writer.WriteStartElement ("paymentMethods")
'            writer.WriteStartElement ("paymentMethodDetails")
'            Dim cTrp As String = FindTRP(Mid(sqlDT(i)("TRP"), 1, 1))
'            If Len(cTrp) = 0 Then
'                MsgBox ("ΔΕΝ ΕΧΩ ΑΝΤΙΣΤΟΙΧΙΣΗ ΣΤΟΝ ΤΡΟΠΟ ΠΛΗΡΩΜΗΣ " + sqldt("TRP"))
'                writer.Close()
'                Exit Sub
'            End If
'            crNode("type", cTrp, writer)
'            crNode("amount", Format(sumNet + sumFpa, "######0.##"), writer)   '  crNode("aa", "15", writer)
'
'            writer.WriteEndElement() ' /paymentMethodDetails
'            writer.WriteEndElement() ' /paymentMethods
'
'
'            Dim SYN_KAU, SYN_FPA As Double
'            SYN_KAU = 0
'            SYN_FPA = 0
'            Dim fpaRow As Double
'
'            '======================================= ΠΕΡΠΑΤΑΩ ΤΟ EGGΤΙΜ ====================================
'            For L As Integer = 0 To EGGTIM.Rows.Count - 1
'
'                Dim AJ As Single
'                If IsDBNull(EGGTIM(L)("TIMM")) Then
'                    AJ = 0
'                Else
'                    AJ = EGGTIM(L)("KAU_AJIA")  ' Math.Round(EGGTIM(L)("POSO") * EGGTIM(L)("TIMM") * (1 - EGGTIM(L)("EKPT") / 100), 2)
'                End If
'
'                Dim VAT As String
'                '1 ΦΠΑ συντελεστής 24% 24%
'                '2 ΦΠΑ συντελεστής 13% 13%
'                '3 ΦΠΑ συντελεστής 6% 6%
'                '4 ΦΠΑ συντελεστής 17% 17%
'                '5 ΦΠΑ συντελεστής 9% 9%
'                '6 ΦΠΑ συντελεστής 4% 4%
'                '7 Άνευ Φ.Π.Α. 0%
'                '8 Εγγραφές χωρίς ΦΠΑ  (πχ Μισθοδοσία, Αποσβέσεις)
'
'
'
'                If EGGTIM(L)("FPA") = 1 Then '13%
'                    VAT = "2"
'                    SYN_KAU = SYN_KAU + AJ
'                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA") ' AJ * 0.13
'                    SYN_FPA = SYN_FPA + fpaRow
'
'                    'ElseIf EGGTIM(L)("FPA") = 2 Then
'                    '   VAT = "1"
'                ElseIf EGGTIM(L)("FPA") = 2 Then
'                    VAT = "1"
'                    SYN_KAU = SYN_KAU + AJ
'                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA") 'AJ * 0.24
'                    SYN_FPA = SYN_FPA + fpaRow
'
'                    ' SYN_FPA = SYN_FPA + AJ * 0.24
'
'                ElseIf EGGTIM(L)("FPA") = 5 Then
'                    VAT = "7"
'                    SYN_KAU = SYN_KAU + AJ
'                    fpaRow = 0
'                    SYN_FPA = SYN_FPA + fpaRow
'
'                ElseIf EGGTIM(L)("FPA") = 6 Then
'                    VAT = "1"
'                    SYN_KAU = SYN_KAU + AJ
'                    ' SYN_FPA = SYN_FPA + AJ * 0.24
'
'                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA")  ' AJ * 0.24
'                    SYN_FPA = SYN_FPA + fpaRow
'
'
'
'
'                ElseIf EGGTIM(L)("FPA") = 4 Then
'                    VAT = "4"
'                    SYN_KAU = SYN_KAU + AJ
'                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA")  ' AJ * 0.06
'                    SYN_FPA = SYN_FPA + fpaRow
'                Else ' If EGGTIM(L)("FPA") = 2 Then
'                    VAT = "1"
'                End If
'                '-----------------------------------------------  invoiceDetails
'                writer.WriteStartElement ("invoiceDetails")
'                crNode("lineNumber", Str(L + 1), writer) '  crNode("lineNumber", "1", writer)
'                '    crNode("quantity", EGGTIM(L)("POSO").ToString, writer)
'                '    crNode("measurementUnit", "1", writer)
'                crNode("netValue", Format(AJ, "######0.##"), writer)  ' crNode("netValue", "100", writer)
'
'                crNode("vatCategory", VAT, writer) '1=24%   2=13%
'
'                crNode("vatAmount", Format(fpaRow, "######0.##"), writer)  ' c
'
'                If fpaRow = 0 Then
'
'                    crNode("vatExemptionCategory", "1", writer) 'APALAGIFPA
'
'                End If
'
'                writer.WriteStartElement ("incomeClassification")
'                crNode("N1:classificationType", Split(ctypos, ";")(1), writer)
'                crNode("N1:classificationCategory", Split(ctypos, ";")(2), writer)
'                crNode("N1:amount", Format(AJ, "######0.##"), writer)
'
'                writer.WriteEndElement() '/incomeClassification
'
'
'
'
'
'                '                <invoiceDetails>
'                '  <lineNumber> 1</lineNumber>
'                '  <netValue>1185</netValue>
'                '  <vatCategory>7</vatCategory>
'                '  <vatAmount>0</vatAmount>
'                '<vatExemptionCategory>1</vatExemptionCategory>
'                '  <incomeClassification>
'                '    <N1:classificationType> E3_561_001</N1:classificationType>
'                '    <N1:classificationCategory>category1_1</N1:classificationCategory>
'                '    <N1:amount> 1185</N1:amount>
'                '  </incomeClassification>
'
'                '</invoiceDetails>
'
'
'
'                '               <incomeClassification>
'                '<N1:classificationType> E3_561_001</N1:classificationType>
'                '            <N1:classificationCategory>category1_1</N1:classificationCategory>
'                '<N1:amount> 100.0</N1:amount>
'                '    </incomeClassification>
'
'
'                writer.WriteEndElement()   ' /invoiceDetails
'            Next
'
'            ExecuteSQLQuery "UPDATE TIM SET AADEKAU=" + Replace(Format(SYN_KAU, "######0.#####"), ",", ".") + ",AADEFPA=" + Replace(Format(SYN_FPA, "######0.#####"), ",", ".") + " WHERE ID_NUM=" + sqldt("ID_NUM").ToString, DUM
'            '------------------------------------------------ InvoiceSummary
'            writer.WriteStartElement ("invoiceSummary")
'            crNode("totalNetValue", Format(SYN_KAU, "######0.##"), writer)  ' crNode("totalNetValue", "100", writer)
'            crNode("totalVatAmount", Format(SYN_FPA, "######0.##"), writer)  '  crNode("totalVatAmount", "24", writer)
'            crNode("totalWithheldAmount", "0", writer)
'            crNode("totalFeesAmount", "0", writer)
'            crNode("totalStampDutyAmount", "0", writer)
'            crNode("totalOtherTaxesAmount", "0", writer)
'            crNode("totalDeductionsAmount", "0", writer)
'            crNode("totalGrossValue", Format(SYN_KAU + SYN_FPA, "######0.##"), writer)
'
'
'            writer.WriteStartElement ("incomeClassification")
'            crNode("N1:classificationType", Split(ctypos, ";")(1), writer)
'            crNode("N1:classificationCategory", Split(ctypos, ";")(2), writer)
'            crNode("N1:amount", Format(SYN_KAU, "######0.##"), writer)
'            writer.WriteEndElement() '  /invoicesummary
'
'
'
'
'            writer.WriteEndElement() '  /invoicesummary
'            '=========================================================
'            writer.WriteEndElement() ' / Invoice
'
'        Next
'
'
'
'
'
'        writer.WriteEndElement() 'InvoicesDoc
'
'
'
'
'
'
'
'
'        writer.WriteEndDocument()
'        writer.Close()
'        '  MsgBox("ok")
'
'
'        ListBox2.Items.Clear()
'
'
'
'        '------ τοπικος ελεγχος xml που τον καταργησα γιατι μπηκε και το "https://www.aade.gr/myDATA/incomeClassificaton/v1.0
'        'FileOpen(1, "C:\F_"+F_TXTFILES+"\CHECKXSD.TXT", OpenMode.Output)
'        ''        Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'        'Dim myDocument As New XmlDocument
'        'myDocument.Load(ff) ' m_filename)  ' "C:\somefile.xml"
'        'myDocument.Schemas.Axmlns:N1dd("http://www.aade.gr/myDATA/invoice/v1.0", "c:\"+F_"+F_TXTFILES+"+"\invoicesDoc-v0.6.xsd") 'namespace here or empty string
'        'Dim eventHandler As ValidationEventHandler = New ValidationEventHandler(AddressOf ValidationEventHandler)
'        'myDocument.Validate(eventHandler)
'        ''       MsgBox("ok ελεγχος")
'        'For n As Integer = 0 To ListBox2.Items.Count - 1
'        '    PrintLine(1, ListBox2.Items(n).ToString)
'        'Next
'        'FileClose(1)
'
'
'
'
'
'
'        paint_ergasies(DataGridView1, "SELECT   ATIM,HME,ENTITY,AADEKAU,AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7 AS KAUTIM,AADEFPA,FPA1+FPA2+FPA3+FPA4+FPA6+FPA7 AS FPATIM,ENTITYUID,ENTITYMARK FROM TIM WHERE ENTITY>0")
'
'    End Sub

Sub ExecuteSQLQuery(sql As String, ByRef RR As ADODB.Recordset)

        '<EhHeader>
        On Error GoTo ExecuteSQLQuery_Err

        '</EhHeader>
100     If InStr(UCase(sql), "SELECT") > 0 Then
102         RR.Open sql, Gdb, adOpenDynamic, adLockOptimistic
    
        Else  ' EXECUTE
    
104         Gdb.Execute sql
    
        End If

        '<EhFooter>
        Exit Sub

ExecuteSQLQuery_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.ExecuteSQLQuery " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Function Get_AJ_ASCII(ByRef pol As String, _
                      ByRef polepis As String, _
                      ByRef ago As String, _
                      ByRef AGOEPIS As String, _
                      ByRef PARASPAROXOY) As Boolean

        '<EhHeader>
        On Error GoTo Get_AJ_ASCII_Err

        '</EhHeader>

        ' Dim R As New ADODB.Recordset
        Dim X      As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If
        Dim sqlDT2 As New ADODB.Recordset

100     sqlDT2.Open "select POL,EIDOS,AJIA_APOU,ISNULL(UBL,'000') AS UBL,POS_APOU from PARASTAT WHERE ISMYDATA=1", Gdb, adOpenDynamic, adLockOptimistic

102     pol = " "

        Dim row As Integer

104     Do While Not sqlDT2.EOF
            ' For row = 0 To sqlDT2.Rows.Count - 1

106         If IsNull(sqlDT2("eidos")) Or IsNull(sqlDT2("pol")) Or IsNull(sqlDT2("ajia_apou")) Then

            Else

108             If sqlDT2("pol") = "1" And sqlDT2("ajia_apou") = "3" Then
110                 pol = pol + "'" + sqlDT2("eidos") + "',"
                End If


                If sqlDT2("pol") = "1" And sqlDT2("POS_apou") = "3" Then
                    pol = pol + "'" + sqlDT2("eidos") + "',"
                End If





112             If sqlDT2("pol") = "1" And sqlDT2("ajia_apou") = "4" Then
114                 polepis = polepis + "'" + sqlDT2("eidos") + "',"
                End If

116             If sqlDT2("pol") = "2" And (sqlDT2("ajia_apou") = "1" Or Val(sqlDT2("POS_apou")) > 0) Then
118                 ago = ago + "'" + sqlDT2("eidos") + "',"
                End If

120

                 If sqlDT2("pol") = "2" And sqlDT2("POS_apou") = "3" Then
122                 AGOEPIS = AGOEPIS + "'" + sqlDT2("eidos") + "',"
                End If




124             DoEvents

            End If
           
126         If Val(sqlDT2("UBL")) = 0 Then ' = "000" Then
           
            Else
128             PARASPAROXOY = PARASPAROXOY + "'" + sqlDT2("eidos") + "',"
            End If
           
130         sqlDT2.MoveNext
        Loop  'Next

132     pol = mID(pol, 1, Len(pol) - 1)

134     If Len(polepis) > 0 Then
136         polepis = mID(polepis, 1, Len(polepis) - 1)
        Else
138         polepis = ""  'ME KENO DHMIOYRGEI PROBLHMA
        End If

140     If Len(ago) > 0 Then
142         ago = mID(ago, 1, Len(ago) - 1)
        Else
144         ago = ""
        End If

146     If Len(PARASPAROXOY) > 0 Then
148         PARASPAROXOY = mID(PARASPAROXOY, 1, Len(PARASPAROXOY) - 1)
        Else
150         PARASPAROXOY = ""
        End If

152     Get_AJ_ASCII = True

154     If Len(AGOEPIS) > 0 Then
156         AGOEPIS = mID(AGOEPIS, 1, Len(AGOEPIS) - 1)
        Else
158         AGOEPIS = "" 'ME KENO DHMIOYRGEI PROBLHMA
        End If

        '<EhFooter>
        Exit Function

Get_AJ_ASCII_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Get_AJ_ASCII " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

'Sub UPLOAD2()
''BALTE EDO TO USERNAME KAI TO SUBSCRIPTION KEY SAS
'
'
''KALESTE THN YPOROUTINA GIA NA FTIAXETE TO XML POY 8A STEILETE STO myData
''Call create_xml(TXML)
'
'    Const URL1 = UrlAADE + "SendInvoices"
'
'    'initialize
'    Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
'    Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
'    XMLServer.setTimeouts 5000, 60000, 10000, 10000
'
'    'force TLS 1.2
'    XMLServer.Option(9) = 2048
'    XMLServer.Option(6) = True
'
'    XMLServer.Open "POST", URL1, False
'    XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
'    XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY ' "555bc57c80634243958f62b629316aaa"
'    XMLServer.send TXML
'    Dim v As String
'    Debug.Print XMLServer.Status
'
'    v = XMLServer.responseText
' Debug.Print v

'End Sub

'Private Sub MakeRequest()
'        ' ΣΤΕΛΝΩ ΣΤΟ SENDINVOICES
'        ' "c:\"+F_TXTFILES+"\inv.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
'        ' H APANTHSH EINAI STO  "c:\"+F_TXTFILES+"\apantSendInv.XML")
'
'
'
'
'        Dim client = New HttpClient()
'        'Dim queryString = HttpUtility.ParseQueryString(String.Empty)
'        Try
'            client.DefaultRequestHeaders.Add("aade-user-id", gUserId) '"glagakis2")
'            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", gSubKey) ' "555bc57c80634243958f62b629316aaa")
'
'            Dim uri = "https://mydata-dev.azure-api.net/SendInvoices" ' + queryString.ToString
'
'            Dim response As HttpResponseMessage
'            Dim xl = XDocument.Load("c:\"+F_TXTFILES+"\inv.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
'            Dim byteData As Byte() = Encoding.UTF8.GetBytes(xl)
'
'            Using content = New ByteArrayContent(byteData)
'                content.Headers.ContentType = New MediaTypeHeaderValue("application/xml")
'                response = Await client.PostAsync(uri, content)
'                Dim result = Await response.Content.ReadAsStringAsync()
'                Dim MF = "c:\"+F_TXTFILES+"\sendinv\apantSendInv" + Format(Now, "yyyyddMMHHmm") + ".xml"
'                FileOpen(1, MF, OpenMode.Output)
'                PrintLine(1, result.ToString)
'                FileClose (1)
'                TextBox2.Text = Result.toString
'                ' "είναι το textbox πανω στη φόρμα που σου επιστρέφει το response xml"
'                'Dim byteData2 As Byte() = File.ReadAllBytes("c:\"+F_TXTFILES+"\inv.xml")
'                ' sept 2020 debug   Rename("c:\"+F_TXTFILES+"\inv.xml", "c:\"+F_TXTFILES+"\inv" + Format(Now, "yyyyddMMHHmm") + ".xml")
'                FileCopy(MF, "c:\"+F_TXTFILES+"\apantSendInv.XML")
'
'            End Using
'
'
'        Catch ex As Exception
'            MsgBox (ex.toString)
'        End Try
'
'
'    End Sub

Sub MAKE_request(TXML As String, _
                 fromFile As Integer, _
                 noask As Integer, _
                 ByVal file_id_num As String, ByRef markReal As String) '(TXML As MSXML2.DOMDocument)

        'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
        ' ΤΗΝ ΑΠΑΝΤΗΣΗ ΤΗΝ ΑΝΤΙΓΡΑΦΩ ΣΤΟ ΦΑΚΕΛΟ  FileCopy mfile, "C:\"+F_TXTFILES+"\apantSendInv.XML"
        'ΕΝΗΜΕΡΨΝΩ ΤΟ ΤΙΜ ΜΕ ΤΑ ΜΑΡΚΣ    UpdateTIMwithMARK
        '<EhHeader>
        On Error GoTo MAKE_request_Err

        '</EhHeader>

        'RequestDocs (GET) =========================
        '        ' ΣΤΕΛΝΩ ΣΤΟ SENDINVOICES
        '        ' "c:\"+F_TXTFILES+"\inv.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
        '        ' H APANTHSH EINAI STO  "c:\"+F_TXTFILES+"\apantSendInv.XML")
        'Dim XL As MSXML2.DOMDocument
        '  .Load("c:\"+F_TXTFILES+"\inv.xml") ' "--> e?e? ??? ap????e?se? t? xml p?? eft?a?a"

        Dim strXML As String, intFile As Integer

        'Open file
100     intFile = FreeFile

102     strXML = ""

        Dim c As String

        '---------------------------------------------------- DIABAZEI ALLA ANSI ----------------------------
        'Open "c:\"+F_TXTFILES+"\inv.xml" For Input As intFile
        '
        ''Load XML into string strXML
        'While Not EOF(intFile)
        '    Line Input #intFile, c
        '    strXML = strXML + c
        'Wend
        'Close intFile

        'If Dir("c:\"+F_TXTFILES+"", vbDirectory) Then

        '--------------------------- αποθηκεύω το ΑΡΧΕΙΟ  inv ΣΕ fakelo "c:\"+F_TXTFILES+"\sendinv\yyyy-mm-dd  ------------------------------
104     If Len(Dir("c:\" + F_TXTFILES + "", vbDirectory)) = 0 Then
106         MkDir "c:\" + F_TXTFILES + ""
        End If

108     If Len(Dir("c:\" + F_TXTFILES + "\sendinv", vbDirectory)) = 0 Then
110         MkDir "c:\" + F_TXTFILES + "\sendinv"
        End If

112     If Len(Dir("c:\" + F_TXTFILES + "\inv.xml", vbNormal)) = 0 Then
114         MsgBox "δεν υπάρχει το αρχείο των τιμολογιων inv.xml"

            Exit Sub

        End If

        Dim todaydir As String

116     todaydir = "c:\" + F_TXTFILES + "\sendinv\" + Format(Now, "yyyyMMdd")

118     If Len(Dir(todaydir, vbDirectory)) = 0 Then
120         MkDir todaydir
        End If
   
122     FileCopy "c:\" + F_TXTFILES + "\inv.xml", todaydir + "\" + f_name_xml + "-" + Format(Now, "yyyy-MM-dd-HH-mm")

        '--------------------------- ΔΙΑΒΑΖΕΙ ΑΡΧΕΙΟ ΣΕ ΜΟΡΦΗ UTF8 ------------------------------
        Dim objStream, strData

124     Set objStream = CreateObject("ADODB.Stream")

126     objStream.Charset = "utf-8"
128     objStream.Open
130     objStream.LoadFromFile ("c:\" + F_TXTFILES + "\inv.xml")

132     strData = objStream.ReadText()

134     objStream.Close

136     If fromFile = 1 Then

138         TXML = strData

        End If

140     Set objStream = Nothing
        '-------------------------------------------------------------------------------------

        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
142     url2 = UrlAADE + "SendInvoices" '/RequestDocs"

        'initialize
144     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
146     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
148     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
150     XMLServer.Option(9) = 2048
152     XMLServer.Option(6) = True
    
154     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
156     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
158     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
160     XMLServer.send TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
162     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
164     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
 
166     Dim mfile As String: mfile = "c:\" + F_TXTFILES + "\sendinv\" + Format(Now, "yyyyMMdd") + "\" + f_name_xml + "-apantSendInv-" + Format(Now, "yyyyddMMHHmm") + ".xml"

        ' Debug.Print v
168     Open mfile For Output As #8
170     Print #8, v
172     Close #8
    
174     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendInv.XML"
    
176     UpdateTIMwithMARK v, TXML, noask, str(file_id_num), markReal
    
        '<EhFooter>
        Exit Sub

MAKE_request_Err:

        Me.Caption = Err.Description & "in Par7MyData.MAKE_request " & "at line " & Erl
               
        Resume Next

        '</EhFooter>

End Sub

Sub read_utf8_file()

        '<EhHeader>
        On Error GoTo read_utf8_file_Err

        '</EhHeader>
        Dim objStream, strData

100     Set objStream = CreateObject("ADODB.Stream")

102     objStream.Charset = "utf-8"
104     objStream.Open
106     objStream.LoadFromFile ("C:\" + F_TXTFILES + "\inv.xlm")

108     strData = objStream.ReadText()

110     objStream.Close
112     Set objStream = Nothing

        '<EhFooter>
        Exit Sub

read_utf8_file_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.read_utf8_file " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Sub KINPROM_CLICK()

End Sub

Public Sub print2USNode(Node As IXMLDOMNode)

        '<EhHeader>
        On Error GoTo printNode_Err

        '</EhHeader>
        Dim xmlNode  As IXMLDOMNode

        Dim xml2Node As IXMLDOMNode

        Dim xml3Node As IXMLDOMNode
     
100     If Node.hasChildNodes Then
       
102         For Each xmlNode In Node.childNodes

104             List1.AddItem xmlNode.nodeName + "*" + xmlNode.Text  ' MARK - ISSUER -COUNTERPART-HEADER
                ' Print xmlNode.nodeName
           
106             If xmlNode.nodeName = "minMark" Then MARK = xmlNode.Text
                '<issuer>
           
                ' If UCase(xmlNode.nodeName) = "counterVatNumber" Then
                '           AFM = xmlNode.Text
                'End If

108             If UCase(xmlNode.nodeName) = UCase("counterVatNumber") Then
110                 AFM = xmlNode.Text
                End If

112             If (xmlNode.nodeName) = "issueDate" Then
114                 issueDate = xmlNode.Text
                End If

116             If (xmlNode.nodeName) = "netValue" Then
118                 totalNetValue = xmlNode.Text
                End If

                'totalVatAmount
120             If (xmlNode.nodeName) = "vatAmount" Then
122                 totalVatAmount = xmlNode.Text
                End If

124             If (xmlNode.nodeName) = "invType" Then
126                 invoiceType = xmlNode.Text
                End If
            
                ' Call print2Node(xmlNode)   netValue
128         Next xmlNode

            ' Print node.nodeName
            'List1.AddItem node.nodeName + "*" + node.Text
        
        End If
        
        '        <bookInfo>
        '    <counterVatNumber>139737839</counterVatNumber>
        '    <issueDate>2022-08-03</issueDate>
        '    <invType>1.1</invType>
        '    <netValue>56.45</netValue>
        '    <vatAmount>13.55</vatAmount>
        '    <grossValue>70</grossValue>
        '    <count>1</count>
        '    <minMark>400001273579873</minMark>
        '    <maxMark>400001273579873</maxMark>
        '  </bookInfo>
        
        '<EhFooter>
        Exit Sub

printNode_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.printNode " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub KINPEL_CLICK(ByVal PEL1 As Integer)   'Sub KINPEL_CLICK()
        'RequestDocs (GET) =========================
    
        Gdb.Execute "DELETE FROM APESTALMENA2"
    
        '<EhHeader>
        On Error GoTo request_Err

        '</EhHeader>

        ' url2 = UrlAADE + "RequestTransmittedDocs" '/RequestDocs"
        ' url2 = "https://mydata-dev.azure-api.net/RequestMyIncome"  ' ?dateFrom={dateFrom}&dateTo={dateTo}[&entityVatNumber][&counterVatNumber][&invType]"
        'initialize
100     url2 = UrlAADE + IIf(PEL1 = 1, "RequestMyIncome", "RequestMyExpenses")

102     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
104     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
106     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
108     XMLServer.Option(9) = 2048
110     XMLServer.Option(6) = True

        Dim mmark As String

        'mmark = InputBox("απο ποια ημερομηνια και μετά  ηη/μμ/εεεε", , Format(Now, "dd/MM/yyyy"))

        Dim mmday As String

        'mmday = InputBox("ΜΕΧΡΙ ποια ημερομηνια         ηη/μμ/εεεε", , Format(Now, "dd/MM/yyyy"))

        Dim mmafm As String

        'mmafm = InputBox("ΑΦΜ", , "000000000")

112     frmOptions.SHOW 1

114     mmark = frmOptionsapo.Caption
116     mmday = frmOptionsEos.Caption
118     mmafm = FRMOptionsafm.Caption

120     If Len(mmark) < 2 Or Len(mmday) < 2 Then
122         MsgBox "ακυρώθηκε"

            Exit Sub

        End If

        '  queryString["dateFrom"] = "{string}";
        '            queryString["dateTo"] = "{string}";
        '            queryString["entityVatNumber"] = "{string}";
        '            queryString["counterVatNumber"] = "{string}";
        '            queryString["invType"] = "{string}";
        '            queryString["maxMark"] = "{string}";

        '?mark={mark}
124     mmark = "?dateFrom=" + mmark + "&dateTo=" + mmday + "&" + afmCompany + "&counterVatNumber=" + mmafm   ' 28/04/2022"  ' deixnei to 29/4/22
    
126     XMLServer.Open "GET", url2 & "" + mmark, False
128     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
130     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY ' "555bc57c80634243958f62b629316aaa"
132     XMLServer.send ""

134     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
136     v = XMLServer.responseText
138     Debug.Print v
    
140     Dim mfile As String: mfile = "c:\" + F_TXTFILES + "\RequestMyIncome" + Format(Now, "yyyyddMMHHmm") + ".xml"

142     Debug.Print v
144     Open mfile For Output As #8
146     Print #8, v
148     Close #8
    
150     XML.Text = v
    
        Dim objxml As MSXML2.DOMDocument

152     Set objxml = New MSXML2.DOMDocument
    
154     objxml.async = True
156     objxml.loadXML v
    
        Dim r0       As New ADODB.Recordset

        Dim mFound   As Integer

        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode

        Dim m_ID     As String

        'Gdb.Open "DSN=MERCSQL;"
158     Set nodeList = objxml.SelectNodes("RequestedBookInfo/bookInfo")

160     For Each Node In nodeList

162         Print Node.nodeName  ' this works'
164         List1.AddItem Node.nodeName + "*******" + Node.Text  ' INVOICE
       
            '       Print node.n
166         Call print2USNode(Node)     'here is the problem explained below'
    
168         If Len(Trim(payaji)) = 0 Then payaji = "0"
170         If Len(Trim(totalGrossValue)) = 0 Then totalGrossValue = "0"
172         If gVal(totalVatAmount) = 0 Then totalVatAmount = "0"
            '    r0.Open "select * from APESTALMENA WHERE AFM='" + AFM0 + "' AND ISSUEDATE='" + issueDate + "' AND TYPOS='" + invoiceType + "' AND AA='" + AA + "'", Gdb, adOpenDynamic, adLockOptimistic
174         FOUND = 0
            '   If r0.EOF Then
            '
            '   Else
            '     If IsNull(r0("ID")) Then
            '
            '     Else
            '        FOUND = 1
            '        m_ID = str(r0("ID"))
            '     End If
            '   End If
        
176         If FOUND = 0 Then
178             AA = "1"
180             sql = "INSERT INTO APESTALMENA2  ([MARK],[AFM],[ISSUEDATE],[AA],[TYPOS],[TOTALNETVALUE],[TOTALVATAMOUNT])"
182             sql = sql + " Values ('" + MARK + "','" + AFM + "','" + issueDate + "','" + AA + "','" + invoiceType + "'," + Replace(totalNetValue, ",", ".") + "," + Replace(totalVatAmount, ",", ".") + ")"

                Dim nn As Long

184             Gdb.Execute sql, nn
            Else
186             Gdb.Execute "update APESTALMENA2 SET MARK='" + MARK + "',PAYTYPE='" + paytype + "',PAYAJI=" + Replace(payaji, ",", ".") + ",TOTALNETVALUE=" + Replace(totalNetValue, ",", ".") + ",TOTALVATAMOUNT=" + Replace(totalVatAmount, ",", ".") + ",TOTALGROSSVALUE=" + Replace(totalGrossValue, ",", ".") + ",CLASSIFICATIONTYPE='" + classificationType + "' WHERE ID=" + m_ID
            End If
        
            ' r0.Close
        
            'Exit For
        
188     Next Node
    
        'Dim sql As String
190     sql = "select  (select TOP 1 TITLOS FROM PARASTAT WHERE ETIK LIKE RTRIM(SUBSTRING(TYPOS,1,4))+'%') AS [ΠΑΡΑΣΤ],(SELECT TOP 1 ISNULL(EPO,'')  FROM PEL WHERE AFM=APESTALMENA2.AFM) AS [ΕΠΩΝΥΜΙΑ],* from APESTALMENA2 ORDER BY ID DESC"

192     Adodc2.ConnectionString = gConnect
194     Adodc2.RecordSource = sql
196     Adodc2.Refresh
    
        '<EhFooter>
        Exit Sub

request_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.request " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Sub requestafm2(Userafm As String, _
                       passafm As String, _
                       askafm As String, _
                       ByRef EPO As String, _
                       ByRef DIE As String, _
                       ByRef EPA As String, _
                       ByRef pol As String, _
                       ByRef tk As String, _
                       ByRef doy As String, _
                       ByRef ENERGOS As Integer)

        'RequestDocs (GET) =========================
        '<EhHeader>
        On Error GoTo request_Err

        '</EhHeader>

100     url2 = "https://www1.gsis.gr:443/wsaade/RgWsPublic2/RgWsPublic2"  '  UrlAADE + "RequestTransmittedDocs" '/RequestDocs"

        'initialize
102     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
104     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
106     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
108     XMLServer.Option(9) = 2048
110     XMLServer.Option(6) = True

        Dim mmark As String

        'mmark = InputBox("απο ποιο μαρκ και μετά", , "400000000000000")

        Dim mmday As String

        'mmday = InputBox("απο ποια ημερομηνια και μετά", , "ηη/μμ/εεεε")

        '  queryString["dateFrom"] = "{string}";
        '            queryString["dateTo"] = "{string}";
        '            queryString["entityVatNumber"] = "{string}";
        '            queryString["counterVatNumber"] = "{string}";
        '            queryString["invType"] = "{string}";
        '            queryString["maxMark"] = "{string}";
        '?mark={mark}
112     mmark = "<env:Envelope xmlns:env=""http://www.w3.org/2003/05/soap-envelope"" xmlns:ns1=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"" xmlns:ns2=""http://rgwspublic2/RgWsPublic2Service"" xmlns:ns3=""http://rgwspublic2/RgWsPublic2"">"
114     mmark = mmark + "   <env:Header>     <ns1:Security>         <ns1:UsernameToken> <ns1:Username>" + Userafm + "</ns1:Username>            <ns1:Password>" + passafm + "</ns1:Password>    </ns1:UsernameToken>      </ns1:Security> </env:Header> "
116     mmark = mmark + "    <env:Body><ns2:rgWsPublic2AfmMethod><ns2:INPUT_REC><ns3:afm_called_by/><ns3:afm_called_for>" + askafm + "</ns3:afm_called_for> </ns2:INPUT_REC></ns2:rgWsPublic2AfmMethod>  </env:Body></env:Envelope>"
        ' mmark "?dateFrom=01/01/2022&dateTo=" + mmday   ' 28/04/2022"  ' deixnei to 29/4/22
    
118     XMLServer.Open "POST", url2, False
120     XMLServer.setRequestHeader "Content-Type", "application/soap+xml"                       '"aade-user-id", P_USER '"glagakis2"
        ' XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY ' "555bc57c80634243958f62b629316aaa"
122     XMLServer.send mmark

124     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
126     v = XMLServer.responseText
        '124  Debug.Print v
        '
        '
        '126  Dim mfile As String: mfile = "c:\" + F_TXTFILES + "\getReqInv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        '128  Debug.Print v
        '130     Open mfile For Output As #8
        '132        Print #8, v
        '134     Close #8
    
        Dim objxml As MSXML2.DOMDocument

128     Set objxml = New MSXML2.DOMDocument
    
130     objxml.async = True
132     objxml.loadXML v

        Dim error_code As String

134     error_code = objxml.getElementsByTagName("error_code").Item(0).Text
        tk = ""
136     If error_code = "RG_WS_PUBLIC_WRONG_AFM" Then
138         ENERGOS = -1
        Else

            Dim cenergos As String

140         cenergos = objxml.getElementsByTagName("deactivation_flag_descr").Item(0).Text

142         If Left$(cenergos, 6) = "ΕΝΕΡΓΟ" Then
144             ENERGOS = 1
146             EPO = objxml.getElementsByTagName("onomasia").Item(0).Text
148             DIE = objxml.getElementsByTagName("postal_address").Item(0).Text + " " + objxml.getElementsByTagName("postal_address_no").Item(0).Text
150             pol = objxml.getElementsByTagName("postal_area_description").Item(0).Text
152             EPA = objxml.getElementsByTagName("firm_act_descr").Item(0).Text

154             doy = objxml.getElementsByTagName("doy_descr").Item(0).Text
156             tk = objxml.getElementsByTagName("postal_zip_code").Item(0).Text
            Else
158             ENERGOS = 0
            End If
 
        End If
 
        Exit Sub

request_Err:
        ' MsgBox Err.Description & vbCrLf & _
          "in ADOMERCNEW.Par7MyData.request " & _
          "at line " & Erl, _
          vbExclamation + vbOKOnly, "Application Error"

160     Resume Next

        '</EhFooter>

End Sub

Sub REQUEST_DIKAMAS()

        'RequestDocs (GET) =========================
        '<EhHeader>
        On Error GoTo REQUEST_DIKAMAS_Err

        '</EhHeader>

100     url2 = UrlAADE + "RequestTransmittedDocs" '/RequestDocs"

        'initialize
102     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
104     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
106     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
108     XMLServer.Option(9) = 2048
110     XMLServer.Option(6) = True

        Dim mmark As String

112     mmark = "" 'InputBox("απο ποιο μαρκ και μετά(-1) ", , "400000000000000")

        Dim mmday As String

114     mmday = InputBox("ΕΩΣ ποια ημερομηνια DD/MM/YYYY Π.Χ. 25/08/2025(ΓΙΑ 25/8/2025)", , Format(Now, "dd/MM/yyyy"))


Gdb.Execute "DELETE FROM APESTALMENA2"
        
mmark = GGET_CVALUE("select TOP 1 ENTITYMARK AS MARK, *from TIM WHERE ENTITYMARK LIKE '4%' AND HME<'" + Format(mmday, "MM/DD/YYYY") + "' ORDER BY ENTITYMARK DESC ")
  'mmark = "400010725396138"
  '        400010726168544"   400010725396139
        
        '  queryString["dateFrom"] = "{string}";
        '            queryString["dateTo"] = "{string}";
        '            queryString["entityVatNumber"] = "{string}";
        '            queryString["counterVatNumber"] = "{string}";
        '            queryString["invType"] = "{string}";
        '            queryString["maxMark"] = "{string}";
        '?mark={mark}
116    ' mmark = mmark + "?dateFrom=14/07/2025&dateTo=" + mmday  ' + "&maxMark=10"    ' 28/04/2022"  ' deixnei to 29/4/22
    
  '  mmark = "?dateFrom=14/07/2025&dateTo=" + mmday   ' + "&maxMark=10"    ' 28/04/2022"  ' deixnei to 29/4/22
   
  ' XMLServer.Open "GET", url2 & "?mark=" + mmark + "&dateTo=19/03/2025"   ΒΑΖΩ ΤΟ ΜΑΡΚ ΚΑΙ ΗΜΕΡΟΜΗΝΙΑ ΚΑΙ ΜΟΥ ΒΓΑΖΕΙ ΤΗΣ ΗΜΕΡΑΣ ΑΠΟ ΤΟ ΜΑΡΚ -1  ΚΑΙ ΜΕΤΑ
    
    'ΤΟ ΠΑΡΚΑΤΩ ΔΟΥΛΕΥΕΙ ΟΚ ΜΕ ΗΜΕΡΟΜΗΝΙΑ ΚΑΙ ΠΙΑΝΕΙ ΟΛΗ ΤΗΝ ΗΜΕΡΑ ΑΝ ΔΩΣΩ ΧΑΜΗΛΟ ΜΑΡΚ, ΑΛΛΟΙΩΣ ΑΝ ΒΑΛΩ ΜΑΡΚ ΓΙΑ ΤΗΝ ΙΔΙΑ ΜΕΡΑ ΑΠΟ ΤΟ ΜΑΡΚ(-1) ΚΑΙ ΠΕΡΑ
118        XMLServer.Open "GET", url2 & "?mark=" + mmark + "&dateTo=" + mmday ' 19/03/2025"

' ΤΕΣΤ ΓΙΑ ΜΕΓΙΣΤΟ
 'XMLServer.Open "GET", url2 & "?mark=" + mmark + "&dateTo=" + mmday + "&maxMark=2" '' 19/03/2025"

' XMLServer.Open "GET", url2 & "?dateFrom=14/07/2025&dateTo=15/07/2025"
 
 
 '+ "&maxMark=400010164786455", False
'    If F_PAROCHOS = 1 Then
'       If F_demo = 1 Then
'             P_USER = PAROX_USER
'             P_KEY = PAROX_KEY
'       End If
'
'    End If

120     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
122     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY ' "555bc57c80634243958f62b629316aaa"



124     XMLServer.send ""

126     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
128     v = XMLServer.responseText
130     Debug.Print v
    
132     Dim mfile As String: mfile = "c:\" + F_TXTFILES + "\getReqInv" + Format(Now, "yyyyddMMHHmm") + ".xml"

134     MsgBox v
136     Open mfile For Output As #8
138     Print #8, v
140     Close #8

'400010164786452
          Shell "notepad.exe " + mfile, vbMaximizedFocus  '400010164786452

    
        Dim objxml As MSXML2.DOMDocument

142     Set objxml = New MSXML2.DOMDocument
    
144     objxml.async = True
146     objxml.loadXML v
    
        Dim r0       As New ADODB.Recordset

        Dim mFound   As Integer

        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode

        Dim m_ID     As String

        'Gdb.Open "DSN=MERCSQL;"
148     Set nodeList = objxml.SelectNodes("RequestedDoc/invoicesDoc/invoice")
        Dim AFM As String
150     For Each Node In nodeList

152         Print Node.nodeName  ' this works'
154         List1.AddItem Node.nodeName + "*******" + Node.Text  ' INVOICE
       
            '       Print node.n
156         Call printUSNode(Node, AFM)    'here is the problem explained below'
    
158         If Len(Trim(payaji)) = 0 Then payaji = "0"
160         If Len(Trim(totalGrossValue)) = 0 Then totalGrossValue = "0"
    
162         r0.Open "select * from APESTALMENA2 WHERE AFM='" + AFM + "' AND ISSUEDATE='" + issueDate + "' AND TYPOS='" + invoiceType + "' AND AA='" + AA + "'", Gdb, adOpenDynamic, adLockOptimistic
164         FOUND = 0

166         If r0.EOF Then
        
            Else

168             If IsNull(r0("ID")) Then
          
                Else
170                 FOUND = 1
172                 m_ID = str(r0("ID"))
                End If
            End If
        
174         If FOUND = 0 Then
176             sql = "INSERT INTO APESTALMENA2  ([MARK],[AFM],[ISSUEDATE],[AA],[TYPOS],[PAYTYPE],[PAYAJI],[TOTALNETVALUE],[TOTALVATAMOUNT],[TOTALGROSSVALUE],[CLASSIFICATIONTYPE])"
178             sql = sql + " Values ('" + MARK + "','" + Left(AFM, 9) + "','" + issueDate + "','" + AA + "','" + invoiceType + "','" + paytype + "'," + Replace(payaji, ",", ".") + "," + Replace(totalNetValue, ",", ".") + "," + Replace(totalVatAmount, ",", ".") + "," + Replace(totalGrossValue, ",", ".") + ",'" + classificationType + "')"

                Dim nn As Long

180             Gdb.Execute sql, nn
            Else
182             Gdb.Execute "update APESTALMENA2 SET MARK='" + MARK + "',PAYTYPE='" + paytype + "',PAYAJI=" + Replace(payaji, ",", ".") + ",TOTALNETVALUE=" + Replace(totalNetValue, ",", ".") + ",TOTALVATAMOUNT=" + Replace(totalVatAmount, ",", ".") + ",TOTALGROSSVALUE=" + Replace(totalGrossValue, ",", ".") + ",CLASSIFICATIONTYPE='" + classificationType + "' WHERE ID=" + m_ID
               
            End If
        
184         r0.Close
        
            'Exit For
        
186     Next Node
    
        'Dim sql As String
188     sql = "select  (select TOP 1 TITLOS FROM PARASTAT WHERE ETIK LIKE RTRIM(SUBSTRING(TYPOS,1,4))+'%') AS [ΠΑΡΑΣΤ],(SELECT TOP 1 EPO FROM PEL WHERE PEL.AFM=APESTALMENA2.AFM) AS [ΠΕΛ],* from APESTALMENA2 ORDER BY ID DESC"

190     Adodc2.ConnectionString = gConnect
192     Adodc2.RecordSource = sql
194     Adodc2.Refresh
Dim m_ID_NUM As Long
Dim M_AFM As String
Dim M_KOD As String
Do While Not Adodc2.Recordset.EOF
   If GGET_NVALUE("select COUNT(*) from TIM WHERE ENTITYMARK='" + Adodc2.Recordset("MARK") + "'") >= 1 Then
      'OK
   Else
      m_ID_NUM = GGET_NVALUE("select ID_NUM from TIM WHERE LEFT(ISNULL(ENTITYMARK,' '),1)<>'4' AND HME='" + Format(Adodc2.Recordset("ISSUEDATE"), "MM/dd/yyyy") + "' and AJ1+AJ2+AJ3+AJ4+AJ5+AJ6=" + Replace(Format(Adodc2.Recordset("totalNETValue"), "#####0.00"), ",", "."))
      M_AFM = GGET_CVALUE("select KPE from TIM WHERE ID_NUM=" + str(m_ID_NUM))
      M_AFM = GGET_CVALUE("SELECT AFM FROM PEL WHERE EIDOS='e' and KOD='" + M_AFM + "'")
      
      If M_AFM = Trim(Adodc2.Recordset("AFM")) Then
           Gdb.Execute ("UPDATE TIM SET ENTITYMARK='" + Trim(Adodc2.Recordset("MARK")) + "' WHERE ID_NUM=" + str(m_ID_NUM))
           MsgBox "OK ENHMEΡΩΘΗΚΕ TO MARK ΠΟΥ ΕΛΕΙΠΕ"
      End If
      
   End If

   Adodc2.Recordset.MoveNext
Loop





    
        '<EhFooter>
        Exit Sub

REQUEST_DIKAMAS_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.REQUEST_DIKAMAS " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdΤιμολογηθενταΣε_Click()

        '<EhHeader>
        On Error GoTo cmdΤιμολογηθενταΣε_Click_Err

        '</EhHeader>
 create_TABLES_APESTALMENA2 ' Command4_Click
100     REQUEST_PROMITH
  
        '<EhFooter>
        Exit Sub

cmdΤιμολογηθενταΣε_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdΤιμολογηθενταΣε_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Function DIAXOR(ByVal v As String, ByVal uid As Integer) As String

        Dim Q As String

100     Q = ""

        Dim objxml As MSXML2.DOMDocument

102     Set objxml = New MSXML2.DOMDocument
    
104     objxml.async = True
106     objxml.loadXML v
    
        Dim r0       As New ADODB.Recordset

        Dim mFound   As Integer

        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode

        Dim m_ID     As String

        'Gdb.Open "DSN=MERCSQL;"
108     Set nodeList = objxml.SelectNodes("InvoicesDoc/invoice")

110     For Each Node In nodeList

112         If Node.firstChild.Text = uid Then
       
114             For Each xmlNode In Node.childNodes

116                 Q = Q + xmlNode.nodeName + "=" + xmlNode.Text + Chr(13)  ' MARK - ISSUER -COUNTERPART-HEADER
                  
118                 If xmlNode.hasChildNodes Then
       
120                     For Each xml2Node In xmlNode.childNodes
                          
122                         If xml2Node.hasChildNodes Then

124                             For Each xml3Node In xml2Node.childNodes

126                                 Q = Q + "          " + xml2Node.nodeName + "=" + xml3Node.Text + Chr(13)
                               
                                Next

                            Else
128                             Q = Q + "  " + xml2Node.nodeName + "=" + xml2Node.Text + Chr(13)
                            End If

                        Next
                          
                    Else
                        
130                     Q = Q + Node.nodeName + "=" + Node.Text + Chr(13)   ' INVOICE
                    
                    End If

                Next
        
            End If

        Next
        
132     DIAXOR = Q
End Function

Function toDatemmdd(ByVal CCDATE As String) As String
toDatemmdd = ""
On Error GoTo OUT
Dim nd As String
nd = mID$(CCDATE, 4, 2) + "/" + mID$(CCDATE, 1, 2) + "/" + mID$(CCDATE, 7, 4)
toDatemmdd = nd
Exit Function
OUT:

End Function




Sub REQUEST_PROMITH()

'Gdb.Execute "DELETE FROM APESTALMENA"
'Gdb.Execute "DELETE FROM APESTALMENADET"



        '<EhHeader>
        On Error GoTo REQUEST_PROMITH_Err

'GoTo 153        '</EhHeader>

        On Error Resume Next



Dim ANS As Integer
ANS = MsgBox("ΝΑ ΣΒΗΣΤΟΥΝ ΤΑ ΠΑΛΙΑ ΠΟΥ ΗΡΘΑΝ ΑΠΟ MYDATA;", vbYesNo)

If ANS = vbYes Then

    Gdb.Execute "DELETE FROM APESTALMENA"

End If






        'RequestDocs (GET) τα παραστατικα που εκοψαν οι προμηθευτες για εμάς
        '=======================================================================================================================
        Dim r3 As New ADODB.Recordset

100     r3.Open "select max(MARK) from APESTALMENA", Gdb, adOpenDynamic, adLockOptimistic

        Dim MARK2 As String

102     If r3.EOF Then
104         MARK2 = "4000100000000"
        Else

            'MARK = r3(0)
106         If IsNull(r3(0)) Then
108             MARK2 = "0000000000000"
            Else
110             MARK2 = Trim(r3(0))
            End If
        
        End If

112     r3.Close
    
    Dim CC As String, CCDATE As String
   CCDATE = InputBox("ΑΠΟ ΠΟΙΑ ΗΜΕΡΟΜΗΝΙΑ ΚΑΙ ΜΕΤΑ; HH/MM/EEEE", "", Format(Now, "dd/MM/yyyy"))
    CCDATE = toDatemmdd(CCDATE)
    
    CC = GGET_CVALUE("SELECT TOP 1 ENTITYMARK FROM TIM WHERE LEN(ENTITYMARK)>10 AND HME<='" + CCDATE + "' order by HME DESC ")
        
      MARK2 = CC
      
        
        
        MARK2 = InputBox("ΑΠΟ ΠΟΙΟ ΜΑΡΚ ΚΑΙ ΜΕΤΑ;", "", MARK2)
    
    
    
114     If Len(Dir("c:\" + F_TXTFILES + "\GETREQ", vbDirectory)) = 0 Then
116         MkDir "c:\" + F_TXTFILES + "\GETREQ"
        End If
    
118     url2 = UrlAADE + "RequestDocs" '/RequestDocs"

        'initialize
120     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
122     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
124     XMLServer.setTimeouts 5000, 60000, 10000, 10000

        'force TLS 1.2
126     XMLServer.Option(9) = 2048
128     XMLServer.Option(6) = True

130     XMLServer.Open "GET", url2 & "?mark=" + MARK2, False
132     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
134     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY ' "555bc57c80634243958f62b629316aaa"
136     XMLServer.send ""

138     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
140     v = XMLServer.responseText
142    ' Debug.Print v

        '

144     Dim mfile As String: mfile = "c:\" + F_TXTFILES + "\GETREQ\getReqANSI.xml" 'getReqISSUER" + Format(Now, "yyyyddMMHHmm") + ".xml"

146    ' Debug.Print v
148     Open mfile For Output As #8
150     Print #8, v
152     Close #8


153
     Dim mfile2 As String: mfile2 = "c:\" + F_TXTFILES + "\GETREQ\getReqANSI.xml" 'getReqISSUER" + Format(Now, "yyyyddMMHHmm") + ".xml"

    ' Debug.Print v
     Open mfile2 For Input As #8
      Dim v2 As String
     Line Input #8, v2
     v2 = v2 + Chr(13)
    Do While Not EOF(8)
      Line Input #8, v
      v2 = v2 + v + Chr(13)
    Loop
    Close #8

v = v2
'select    *    from APESTALMENA   WHERE AA IN ( SELECT ATIM FROM APESTALMENADET WHERE LEN(KODE)>1 )

        'ME ADODB STREAM ΓΡΑΦΕΙ ΣΕ UTF-8 TEXNIKO KAI DEN MAS KANEI
        'Dim St2  As ADODB.Stream
        'Set St2 = New ADODB.Stream
        'St2.Mode = adModeReadWrite
        'St2.Type = adTypeText
        'St2.Charset = "UTF-8"
        '
        '
        'St2.Open
        'St2.LoadFromFile (mfile)
        'St2.Position = St2.Size
        ''St2.WriteText v
        'St2.SaveToFile ("c:\"+F_TXTFILES+"\getReqIS2.XML")
        'St2.Close
        '
        '
        '
        '
        '
        '
        '

        'ME ADODB STREAM ΓΡΑΦΕΙ ΣΕ UTF-8 TEXNIKO KAI DEN MAS KANEI
        'Dim St2  As ADODB.Stream
        'Dim Streamout  As ADODB.Stream
        '
        'Set St2 = New ADODB.Stream
        'St2.Mode = adModeReadWrite
        'St2.Type = adTypeText
        'St2.Charset = "UTF-8"
        '
        '
        'St2.Open
        'St2.LoadFromFile (mfile)
        '
        '
        'St2.Position = 0
        '
        '
        'Set Streamout = New ADODB.Stream
        'Streamout.Charset = "utf-8"
        'Streamout.Open
        'Streamout.WriteText St2.ReadText
        '
        '
        '
        ''St2.WriteText v
        'Streamout.SaveToFile "c:\"+F_TXTFILES+"\getRequtf.XML", adSaveCreateOverWrite
        '
        'St2.Close
        'Streamout.Close
        '
    
        '    Set objXML = CreateObject("Msxml.DOMDocument")
        '   objXML.async = True
        '  objXML.Load "C:\"+F_TXTFILES+"\GETREQ2.xml"     ' mfile  '
    
        Dim objxml As MSXML2.DOMDocument

154     Set objxml = New MSXML2.DOMDocument
    
156     objxml.async = True
158     objxml.loadXML v
    
        Dim r0       As New ADODB.Recordset

        Dim mFound   As Integer

        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode

        Dim m_ID     As String

        'Gdb.Open "DSN=MERCSQL;"
160     Set nodeList = objxml.SelectNodes("RequestedDoc/invoicesDoc/invoice")

162     For Each Node In nodeList

            ' Print Node.nodeName  ' this works'
164         List1.AddItem Node.nodeName + "*******" + Node.Text  ' INVOICE
       
            '       Print node.n
166         Call printNode(Node, 1)   'here is the problem explained below'
    
168         If Len(Trim(payaji)) = 0 Then payaji = "0"
170         If Len(Trim(totalGrossValue)) = 0 Then totalGrossValue = "0"
    
172         r0.Open "select * from APESTALMENA WHERE AFM='" + AFM0 + "' AND ISSUEDATE='" + issueDate + "' AND TYPOS='" + invoiceType + "' AND AA='" + AA + "'", Gdb, adOpenDynamic, adLockOptimistic
174         FOUND = 0

176         If r0.EOF Then
        
            Else

178             If IsNull(r0("ID")) Then
          
                Else
180                 FOUND = 1
182                 m_ID = str(r0("ID"))
                End If
            End If
        
184         If FOUND = 0 Then
                'AA = "G" + Right(Format(Val(AA), "000000"), 6)
186             sql = "INSERT INTO APESTALMENA  ([MARK],[AFM],[ISSUEDATE],[AA],[TYPOS],[PAYTYPE],[PAYAJI],[TOTALNETVALUE],[TOTALVATAMOUNT],[TOTALGROSSVALUE],[CLASSIFICATIONTYPE])"
188             sql = sql + " Values ('" + MARK + "','" + AFM0 + "','" + issueDate + "','" + AA + "','" + invoiceType + "','" + paytype + "'," + Replace(payaji, ",", ".") + "," + Replace(totalNetValue, ",", ".") + "," + Replace(totalVatAmount, ",", ".") + "," + Replace(totalGrossValue, ",", ".") + ",'" + classificationType + "')"
190             Gdb.Execute sql
            Else
192             Gdb.Execute "update APESTALMENA SET MARK='" + MARK + "',PAYTYPE='" + paytype + "',PAYAJI=" + Replace(payaji, ",", ".") + ",TOTALNETVALUE=" + Replace(totalNetValue, ",", ".") + ",TOTALVATAMOUNT=" + Replace(totalVatAmount, ",", ".") + ",TOTALGROSSVALUE=" + Replace(totalGrossValue, ",", ".") + ",CLASSIFICATIONTYPE='" + classificationType + "' WHERE ID=" + m_ID
        
            End If
        
194         r0.Close
        
196     Next Node
    
        ' -----------------------------------ακυρωμενα--------------------------------------------------------------------------------------------------
        ' <cancelledInvoicesDoc>
        '    <cancelledInvoice>
        '      <invoiceMark>400001398277034</invoiceMark>
        '      <cancellationMark>400001400657944</cancellationMark>
        '      <cancellationDate>2022-09-15</cancellationDate>
        '    </cancelledInvoice>
        '  </cancelledInvoicesDoc>

        '    If Node.hasChildNodes Then
        '
        '102         For Each xmlNode In Node.childNodes
        '104            List1.AddItem xmlNode.nodeName + "*" + xmlNode.Text  ' MARK - ISSUER -COUNTERPART-HEADER
        '               ' Print xmlNode.nodeName
        '
        '106            If xmlNode.nodeName = "mark" Then mark = xmlNode.Text
        '               '<issuer>
        'On Error GoTo 0
198    Set nodeList = objxml.SelectNodes("RequestedDoc/cancelledInvoicesDoc/cancelledInvoice")

200    For Each Node In nodeList

           'Print Node.nodeName  ' this works'
202        List1.AddItem Node.nodeName + "*******" + Node.Text  ' INVOICE
       
            '       Print node.n
204        For Each xmlNode In Node.childNodes

206             If xmlNode.nodeName = "invoiceMark" Then
208                 MARK = xmlNode.Text
 
210                Gdb.Execute "update APESTALMENA SET PAYAJI=-1,TOTALNETVALUE=0,TOTALVATAMOUNT=0,TOTALGROSSVALUE=0 WHERE MARK='" + MARK + "'"
212                 Gdb.Execute "update TIM SET SXETMARK='' WHERE SXETMARK='" + MARK + "'"
        
                End If

            Next

           'r0.Close
        
        Next
    
        'Dim sql As String
214   '  sql = "select  (select TOP 1 TITLOS FROM PARASTAT WHERE ETIK LIKE RTRIM(SUBSTRING(TYPOS,1,4))+'%') AS [ΠΑΡΑΣΤ],ATIM,* from APESTALMENA WHERE YEAR(ISSUEDATE)=" + str(Year(gLhjh)) + "    ORDER BY ID DESC"
     sql = "select ISANTIST,P.EPO,A.* from APESTALMENA A inner JOIN PEL P ON P.EIDOS='r' and P.AFM=A.AFM   WHERE A.AA IN ( SELECT ATIM FROM APESTALMENADET WHERE LEN(KODE)>1 )"
216     Adodc2.ConnectionString = gConnect
218     Adodc2.RecordSource = sql
220     Adodc2.Refresh
    
        '<EhFooter>
        Exit Sub

REQUEST_PROMITH_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.REQUEST_PROMITH " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Sub create_TABLES_APESTALMENA2()

        Dim sql As String
         
100     sql = " CREATE TABLE [dbo].[APESTALMENA2](" & "[AFM] [nchar](10) NULL," & "[ISSUEDATE] [nchar](10) NULL," & "[AA] [nchar](10) NULL," & "[TYPOS] [nchar](10) NULL," & "[PAYTYPE] [nchar](10) NULL," & "[PAYAJI] [float] NULL," & "[TOTALNETVALUE] [float] NULL," & "[TOTALVATAMOUNT] [float] NULL," & "[TOTALGROSSVALUE] [float] NULL, " & "[CLASSIFICATIONTYPE] [nchar](10) NULL," & "[MARK] [nchar](20) NULL," & "[YPOK] [int],[ID] [int] IDENTITY(1,1) NOT NULL" & ") ON [PRIMARY]"

        Dim R As New ADODB.Recordset
         
102     R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'APESTALMENA2'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

104     If R(0) = 0 Then
106         Gdb.Execute sql
        End If
        R.Close
        
        
         R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'APESTALMENA'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next
         sql = " CREATE TABLE [dbo].[APESTALMENA](" & "[AFM] [nchar](10) NULL," & "[ISSUEDATE] [nchar](10) NULL," & "[AA] [nchar](10) NULL," & "[TYPOS] [nchar](10) NULL," & "[PAYTYPE] [nchar](10) NULL," & "[PAYAJI] [float] NULL," & "[TOTALNETVALUE] [float] NULL," & "[TOTALVATAMOUNT] [float] NULL," & "[TOTALGROSSVALUE] [float] NULL, " & "[CLASSIFICATIONTYPE] [nchar](10) NULL," & "[MARK] [nchar](20) NULL," & "[YPOK] [int],[ID] [int] IDENTITY(1,1) NOT NULL" & ") ON [PRIMARY]"

        If R(0) = 0 Then
            Gdb.Execute sql
        End If
        R.Close
        
        
        
        
        R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'APESTALMENA2DET'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

     If R(0) = 0 Then
         Gdb.Execute "SELECT TOP 0 * INTO APESTALMENA2DET FROM EGGTIM"
        End If
        R.Close
            
            
        R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'APESTALMENADET'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

     If R(0) = 0 Then
         Gdb.Execute "SELECT TOP 0 * INTO APESTALMENADET FROM EGGTIM"
        End If
        R.Close
        
End Sub
Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>
        create_TABLES_APESTALMENA2
        
108    REQUEST_DIKAMAS

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Command4_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub EOS_Change()

        '<EhHeader>
        On Error GoTo EOS_Change_Err

        '</EhHeader>
        Dim sql As String

100     sql = FIND_QUERY()

102     If monoApestalmena.Value = vbUnchecked Then

104         Adodc1.ConnectionString = gConnect
106         Adodc1.RecordSource = sql
108         Adodc1.Refresh

            '<EhFooter>
        
        Else
            monoApestalmena_Click

        End If
        
        Exit Sub

EOS_Change_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.EOS_Change " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'
'
'
Private Sub Form_Load()
100     loadpar7
End Sub
    
Public Sub loadpar7()

        '  gdb.Open "DSN=MERCSQL;DATABASE=EMPMYDATA"
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        f_SDA = FINDPARAMETROI(1, "PAR1", "F_SDA", "Ω;V", "ΧΑΡΑΚΤΗΡ.ΓΡΑΜΜΑ ΣΔΑ ΕΝΑΡΞΗΣ ΚΑΙ ΕΠΙΣΤΡΟΦΗΣ Π.Χ. Φ;Ψ") '
        If Len(f_SDA) = 0 Or InStr(f_SDA, ";") = 0 Then
            f_SDA = " ; "
        End If
        
        F_brat = FINDPARAMETROI(1, "PAR7MYDATA", "f_brat", "0", "0=parochos.gr 1=bratnet ")
        
        
        F_ANAPODO = FINDPARAMETROI(3, "PAR1", "F_ANAPODO", "ΩΨ", "Κωδικοί παρ/κων με ανάποδο ΦΟΡΤΩΣΗ/ΠΡΟΟΡΙΣΜΟ ;")
        
        
        
100     F_TXTFILES = Trim(FINDPARAMETROI(1, "PAR7MYDATA", "F_TXTFILES", "TXTFILES", "ΦΑΚΕΛΟΣ ΑΠΟΣΤΩΛΩΝ "))

102     F_E3_APOKATHG_EID = Val(FINDPARAMETROI(1, "PAR7MYDATA", "F_E3_APOKATHG_EID", "0", "E3 από κατηγορία ειδών=1 Απο Παραστατικά=0 "))

104     F_dokimi = Val(FINDPARAMETROI(1, "PAR7MYDATA", "F_dokimi", "1", "1=δοκιμη 1ος user 2=δοκμη 2ος user 3=bratnet"))

106     f_YPOK = FINDPARAMETROI(1, "PAR7MYDATA", "f_YPOK", "0", "ΑΡΙΘΜΟΣ ΥΠΟΚΑΤΑΣΤΗΜΑΤΟΣ 0=ΚΕΝΤΡΙΚΟ ")
108     f_YPOK = Left$(f_YPOK, 1)

f_Stelno_daiakin = FINDPARAMETROI(1, "PAR7MYDATA", "f_Stelno_daiakin", "0", "1=Να φαίνεται αυτόματα οτι ειναι σε διακινηση ")



110     If InStr("0123", f_YPOK) = 0 Then
112         f_YPOK = "0"
        End If

        'Dim f_par1ypok As String ' παραστατικά που ανήκουν στο υποκ/μα 1
114     f_par1ypok = FINDPARAMETROI(1, "PAR7MYDATA", "f_par1ypok", "", "παραστατικά που ανήκουν στο υποκ/μα 1 π.χ.  αβ  ")

116     f_par2ypok = FINDPARAMETROI(1, "PAR7MYDATA", "f_par2ypok", "", "παραστατικά που ανήκουν στο υποκ/μα 2 π.χ.  gW  ")

118     F_PAROCHOS = Val(FINDPARAMETROI(1, "PAR7MYDATA", "f_parochos", "0", "0=MyData 1=Παροχος"))
        MDIForm1.StatusBar1.Panels(6).Text = IIf(F_PAROCHOS = 1, "ΠΑΡΟΧΟΣ", "myData")

120     If F_PAROCHOS = 1 Then
122         MYDPAR.Visible = True
124         MYDPAR.Text = MYDPAR.List(1)
        End If

126     f_apoDOS = Val(FINDPARAMETROI(1, "PAR7MYDATA", "f_apoDOS", "0", "ΜΠΟΡΩ ΝΑ ΤΡΑΒΗΞΩ ΑΠΟ DOS C:\LAGEURO=1 OXI=0 "))

128     If f_apoDOS = 1 Then
130         APODOS.Visible = True
   
        End If

        Dim R As New ADODB.Recordset

132     R.Open "select * from MEM", Gdb, adOpenDynamic, adLockOptimistic

134     f_pelono = R("pelono")
136     f_peldie = R("peldie")
138     f_pelarit = R("PELARIT")
140     f_pelpol = R("PELPOL")
142     f_peltk = R("PELTK")

144     F_demo = 0

146     If R!mydataIsDemo = 1 Then
148         PAROX_URLADD = "beta-"

150         If F_dokimi = 1 Then
152             afmCompany = "028783755"
154             P_USER = "glagakis22"   ' "555bc57c80634243958f62b629316aaa"
156             P_KEY = "534044b048f4023151f37c2a44282200"  ' "555bc57c80634243958f62b629316aaa" ' CNull(R!MYDATAPASW)
                 
158             PAROX_USER = "glagakis@gmail.com"
160             PAROX_PASS = "Lagakis123!@#" ' y39213921!!Y"
162             PAROX_KEY = "32657FD7929B483AABBD1C633401E945"



           ElseIf F_dokimi = 3 Then 'bratnet===============================
1522             afmCompany = "028783755"
1542             P_USER = "glagakis22"   'mydata    "555bc57c80634243958f62b629316aaa"
1562             P_KEY = "534044b048f4023151f37c2a44282200"  'mydata key    "555bc57c80634243958f62b629316aaa" ' CNull(R!MYDATAPASW)
                 
1582             PAROX_USER = "028783755"
1602             PAROX_PASS = "Lagakis123!@#" ' y39213921!!Y"
1622             PAROX_KEY = "1ZZFXIz(bXmN^w_ruz!4H#EC9nHHZqz!"
            Else
164             afmCompany = "150144455"
166             P_USER = "plagakis"
168             P_KEY = "ed99b37f3e99cc46ba284186185b5cdc"
       
170             PAROX_USER = "glagakis@gmail.com"
172             PAROX_PASS = "y39213921!!Y"
174             PAROX_KEY = "32657FD7929B483AABBD1C633401E945"
       
            End If
       
176         F_demo = 1
178         UrlAADE = "https://mydataapidev.aade.gr/" '"https://mydata-dev.azure-api.net/"
        
            '  XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
            '  XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
       
180         MsgBox "προσοχή δοκιμαστικο περιβαλλον " + IIf(F_PAROCHOS = 1, "παροχου", "myData")
     
        Else
182         UrlAADE = "https://mydatapi.aade.gr/myDATA/"  'SendInvoices
184         PAROX_URLADD = ""
            ' PAROX_URLADD = "beta-" ' προσοχη debug only να σβηστει στην παραγωγή

186         P_USER = CNull(R!MYDATAUSER)
188         P_KEY = CNull(R!MYDATAPASW)

190         afmCompany = CNull(R!AFMPEL)

192         If IsNull(R!AFMPEL) Then
194             MsgBox "δεν εχει δηλωθει  το ΑΦΜ (Βοηθητικα παράμετροι προγραμματος)"
196             Unload Me

                Exit Sub

            End If
   
198         If IsNull(R!MYDATAUSER) Then
200             MsgBox "δεν εχει δηλωθει  το USER NAME ΓΙΑ ΤΟ MYDATA.(Βοηθητικα παράμετροι προγραμματος)"
202             Unload Me

                Exit Sub

            End If

204         If IsNull(R!MYDATAPASW) Then
206             MsgBox "δεν εχει δηλωθει ο ΚΩΔΙΚΟΣ ΓΙΑ ΤΟ MYDATA.(Βοηθητικα παράμετροι προγραμματος)"
208             Unload Me

                Exit Sub

            End If


210         PAROX_USER = CNull(R.FIELDS("PAROXUSER").Value) '"glagakis@gmail.com"
212         PAROX_PASS = CNull(R.FIELDS("PAROXPASS").Value) '"y39213921!!Y"
214         PAROX_KEY = CNull(R.FIELDS("PAROXKEY").Value) '"32657FD7929B483AABBD1C633401E945"

        End If
 
216     R.Close
 
218     R.Open "select ISNULL(AYJON,0) AS AYJON,ISNULL(TIMH,0) AS TIMH,ISNULL(N1,0) AS N1  FROM PINAKES WHERE TYPOS=1", Gdb, adOpenDynamic, adLockOptimistic

        Dim lp As Integer

220     Do While Not R.EOF
222         antFPA(R!AYJON) = R!n1
224         timhFPA(R!AYJON) = R!TIMH
    
226         If R!n1 = 0 Or R!n1 > 8 Then
228             MsgBox "ΛΑΘΟΣ ΣΤΙΣ ΑΝΤΙΣΤΟΙΧΗΣΕΙΣ ΦΠΑ"
230             APO.Enabled = False
232             EOS.Enabled = False
234             Command2.Enabled = False
236             monoApestalmena.Enabled = False
            End If

238         R.MoveNext
        Loop
 
240     R.Close
  
242     APO.Value = Now
244     EOS.Value = Now

        Dim sql As String

246     sql = FIND_QUERY()

248     If sql = "0" Then

250         APO.Enabled = False
252         EOS.Enabled = False
254         Command2.Enabled = False
256         monoApestalmena.Enabled = False

258         GoTo lab1000
        End If

260     If Len(where.Caption) > 0 Then  ' ONLINE ΔΕΝ ΕΧΡΕΙΑΖΕΤΑΙ GRID
    
        Else

262         If gisOnlineMydata = 0 Then
264             Adodc1.ConnectionString = gConnect  '1
266             Adodc1.RecordSource = sql
268             Adodc1.Refresh
            End If

        End If

        Dim R56 As New ADODB.Recordset

270     R56.Open "select MYDATAFROMEID1 FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
272     fMydataFromEID1 = nNull(R56(0))
274     R56.Close

        Dim n50 As Integer

276     If fMydataFromEID1 = 1 Then '  'ΕΝΗΜΕΡΩΣΗ ΑΠΟ ΚΑΤΗΓΟΡΙΑ ΕΙΔΩΝ
278         R.Open "SELECT AYJON,C1,LOG1,ISNULL(C3,'') AS C3 FROM PINAKES WHERE TYPOS=15", Gdb, adOpenDynamic, adLockOptimistic
280         n50 = 0

282         Do While Not R.EOF
284             n50 = n50 + 1
286             fKatEsod(n50) = Trim(CNull(R!C1))
288             FkatE3Esod(n50) = Trim(CNull(R!c3)) 'E3-561-1 ΑΝ ΤΡΑΒΑΕΙ ΤΟ Ε3 ΑΠΟ ΤΗΝ ΚΑΤΗΓΟΡΙΑ ΕΙΔΟΥΣ ΚΑΙ ΟΧΙ ΑΠΟ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ
290             fKatEXod(n50) = Trim(CNull(R!log1))
292             R.MoveNext
            Loop

294         R.Close
    
296         If n50 = 1 Then
298             MsgBox "δεν εχουν δηλωθει οι κατηγορίες ειδών"

                Exit Sub

            End If
   
        End If
 
        '174  sql = " CREATE TABLE [dbo].[APESTALMENA](" _
        '       & "[AFM] [nchar](10) NULL," _
        '       & "[ISSUEDATE] [nchar](10) NULL," _
        '       & "[AA] [nchar](10) NULL," _
        '       & "[TYPOS] [nchar](10) NULL," _
        '       & "[PAYTYPE] [nchar](10) NULL," _
        '       & "[PAYAJI] [float] NULL," _
        '       & "[TOTALNETVALUE] [float] NULL," _
        '       & "[TOTALVATAMOUNT] [float] NULL," _
        '       & "[TOTALGROSSVALUE] [float] NULL, " _
        '       & "[CLASSIFICATIONTYPE] [nchar](10) NULL," _
        '       & "[MARK] [nchar](20) NULL," _
        '       & "[ID] [int] IDENTITY(1,1) NOT NULL" _
        '    & ") ON [PRIMARY]"
        '
        '
        '
        '176       R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'APESTALMENA'", Gdb, adOpenDynamic, adLockOptimistic
        '        'On Error Resume Next
        '
        '178     If R(0) = 0 Then
        '180         Gdb.Execute sql
        '        End If
        '
        '
        '182   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE11", "[nchar](10) NULL")
        '184   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE12", "[nchar](10) NULL")
        '186   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT11", "[nchar](10) NULL")
        '188   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT12", "[nchar](10) NULL")
        '
        '190   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE21", "[nchar](10) NULL")
        '192   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE22", "[nchar](10) NULL")
        '194   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT21", "[nchar](10) NULL")
        '196   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT22", "[nchar](10) NULL")
        '
        '198   dum2 = ADD_FIELD("APESTALMENA", "MARK2", "[nchar](20) NULL")
        '
        '200    dum2 = ADD_FIELD("APESTALMENA", "AMOUNT1", "float NULL")
        '202    dum2 = ADD_FIELD("APESTALMENA", "AMOUNT2", "float NULL")
        '
        '204    dum2 = ADD_FIELD("APESTALMENA", "ISANTIST", "INT NULL")
        '206    dum2 = ADD_FIELD("APESTALMENA", "ATIM", "NVARCHAR(12) NULL ")
        '
        '207    dum2 = ADD_FIELD("APESTALMENA", "DEMO", "INT NULL ")
        '

        'paris user           lagakis2
        'paris key            c10c55456c3a411b866cd0f5e69155dc

        'XMLServer.setRequestHeader "aade-user-id", "glagakis2"
        'XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa"

        '<EhFooter>
lab1000:
       Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Form_Load " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Function FINDTYPOS(c As String) As String

        '<EhHeader>
        On Error GoTo FINDTYPOS_Err

        '</EhHeader>
        Dim s4 As New ADODB.Recordset

100     s4.Open "select * from PARASTAT where EIDOS='" + c + "'", Gdb, adOpenDynamic, adLockOptimistic

102     If IsNull(s4!MYEID + s4!myTypEsod + s4!mycateg) Then
104         FINDTYPOS = ";;;;;;"
        Else
106         FINDTYPOS = s4!MYEID + ";" + s4!myTypEsod + ";" + s4!mycateg + ";" + CNull(s4!MYISAYTOTIM) + ";"
        End If

        '       dum2 = ADD_FIELD("PARASTAT", "MYEID", "nvarCHAR(10)")
        'dum2 = ADD_FIELD("PARASTAT", "MYTYPESOD", "nvarCHAR(20)")
        'dum2 = ADD_FIELD("PARASTAT", "MYCATEG", "nvarCHAR(20)")
        'dum2 = ADD_FIELD("PARASTAT", "MYISAYTOTIM", "INT")

108     FINDTYPOS = Trim(FINDTYPOS)

        ' If InStr(FINDTYPOS, ";") = 0 Then
110     FINDTYPOS = FINDTYPOS + ";;" ' gia na mhn skaei to split()
        ' Else

        ' End If

        '<EhFooter>
        Exit Function

FINDTYPOS_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.FINDTYPOS " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function
    
Function FINDEXODTYPOS(c As String) As String

        '<EhHeader>
        On Error GoTo FINDEXODTYPOS_Err

        '</EhHeader>
        
        Dim s4 As New ADODB.Recordset

100     s4.Open "select * from PARASTAT where EIDOS='" + c + "'", Gdb, adOpenDynamic, adLockOptimistic

102     If IsNull(s4!myTypEXod + s4!myCatExod) Then
104         FINDEXODTYPOS = ""
        Else
106         FINDEXODTYPOS = s4!myTypEXod + ";" + s4!myCatExod + ";;"
        End If

108     FINDEXODTYPOS = Trim(FINDEXODTYPOS)

        ' If InStr(FINDTYPOS, ";") = 0 Then
110     FINDEXODTYPOS = FINDEXODTYPOS + ";;" ' gia na mhn skaei to split()
        ' Else

        ' End If

        '<EhFooter>
        Exit Function

FINDEXODTYPOS_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.FINDEXODTYPOS " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function
    
Function FIND2_EXODTYPOS(c As String) As String

        '<EhHeader>
        On Error GoTo FINDEXODTYPOS_Err

        '</EhHeader>
        
        Dim s4 As New ADODB.Recordset

100     s4.Open "select * from PARASTAT where EIDOS='" + c + "'", Gdb, adOpenDynamic, adLockOptimistic

102     If IsNull(s4!myTypEXod + s4!myCatExod) Then
104         FIND2_EXODTYPOS = ""
        Else
106         FIND2_EXODTYPOS = s4!myTypEXod + ";" + s4!myCatExod + ";;"
        End If

108     FIND2_EXODTYPOS = Trim(FIND2_EXODTYPOS)

        ' If InStr(FINDTYPOS, ";") = 0 Then
110     FIND2_EXODTYPOS = FIND2_EXODTYPOS + ";;" ' gia na mhn skaei to split()
        ' Else

        ' End If

        '<EhFooter>
        Exit Function

FINDEXODTYPOS_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.FINDEXODTYPOS " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function
    
Function FindTRP(c As String) As String

        '<EhHeader>
        On Error GoTo FindTRP_Err

        '</EhHeader>
        Dim sqlDT4 As New ADODB.Recordset

100     sqlDT4.Open "select N1 from PINAKES where TYPOS=12 AND  AYJON=" + c + "", Gdb, adOpenDynamic, adLockOptimistic

102     If IsNull(sqlDT4(0)) Then
104         FindTRP = ""
        Else
106         FindTRP = sqlDT4(0)
        End If

        '<EhFooter>
        Exit Function

FindTRP_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.FindTRP " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Sub TEST_sendtimol()

End Sub

Private Sub jsontimol_Click()
 
        ' Dim iss As New Issuer
        '        iss.vatNumber = "028783755"
        '        iss.branch = 0
        '        iss.street = "?.????????G??"
        '        iss.streetNumber = "21"
        '        iss.postalCode = "66100" '       "postalCode": "postalCode",
        '        iss.city = "?????" '       "city": "city",
        '        iss.COUNTRY = "GR" '       "country": "GR"
        '
        '
        '
        '
        '
        '
        '
        '        Dim cou As New counterpart
        '        cou.vatNumber = "150144455"
        '        cou.Name = "??G???S ????S????S"
        '        cou.branch = 0
        '        cou.street = "S?????????"
        '        cou.streetNumber = "13"
        '        cou.postalCode = "66100" '       "postalCode": "postalCode",
        '        cou.city = "?????" '       "city": "city",
        '        cou.COUNTRY = "GR" '
        '        cou.email = "test@email.com" '
        '
        '
        '        Dim deliv As New deliveryAddress
        '        deliv.street = "S????????? "
        '        deliv.streetNumber = "13"
        '        deliv.postalCode = "66100" '       "postalCode": "postalCode",
        '        deliv.city = "?????" '       "city": "city",
        '        deliv.COUNTRY = "GR" '
        '
        '        '  Dim det(10) As InvoiceDetail
        '        Dim newitem(10) As InvoiceDetail
        '
        '        newitem(0) = New InvoiceDetail
        '        newitem(0).lineNumber = 1
        '        newitem(0).entityName = "KARIDIA"
        '        'newitem(0).recType = Null
        '        newitem(0).invoiceDetailType = 0
        '        newitem(0).quantity = 1
        '        newitem(0).netValue = 10
        '        newitem(0).vatCategory = 1
        '        newitem(0).vatCategoryUbl = "S"
        '        newitem(0).vatAmount = 2.4
        '        newitem(0).vatPercent = 24
        '        newitem(0).MeasurementUnit = 1
        '        newitem(0).measurementUnitUbl = "H87"
        '        newitem(0).lineComments = "S?????"
        '
        '        newitem(0).totalValue = 12.4
        '        newitem(0).classificationCategory = "category1_1"
        '        newitem(0).classificationType = "E3_561_001"
        '        newitem(0).cpvCode = "45233222-1"
        '
        '
        '
        '        newitem(1) = New InvoiceDetail
        '        newitem(1).lineNumber = 2
        '        newitem(1).entityName = "KARIDIA" 'ONOMA
        '        'newitem(1).recType = Null
        '        newitem(1).invoiceDetailType = 0
        '        newitem(1).quantity = 1  'POSO
        '        newitem(1).netValue = 10#   '
        '        newitem(1).vatCategory = 1
        '        newitem(1).vatCategoryUbl = "S"
        '        newitem(1).vatAmount = 2.4
        '        newitem(1).vatPercent = 24 ' G_FPA(FPA)
        '        newitem(1).MeasurementUnit = 1
        '        newitem(1).measurementUnitUbl = "H87"
        '        newitem(1).lineComments = ""  ' "S?????"
        '
        '        newitem(1).totalValue = 12.4
        '        newitem(1).classificationCategory = "category1_1"
        '        newitem(1).classificationType = "E3_561_001"
        '        newitem(1).cpvCode = "45233222-1"
        '
        '        Dim pay(5) As PaymentMethod
        '
        '        pay(0) = New PaymentMethod
        '        pay(0).Type = 5
        '        pay(0).amount = 24.8
        '
        '
        '
        '
        '
        '        Dim he As New invoiceHeader
        '        he.Series = "b"
        '        he.issueDate = "2024-04-21"
        '        he.dispatchDate = "2024-04-21"
        '        he.AA = 218
        '        he.invoiceCode = "INV000218"
        '        he.invoiceTypeUbl = "380"
        '        he.invoiceType = "1.1"
        '        he.Currency = "EUR"
        '        he.movePurpose = 1
        '        he.selfPricing = False
        '
        '        he.fuelInvoice = False
        '
        '
        '
        '
        '        Dim invs As New invoiceSummary
        '
        '        invs.totalVatAmount = 4.8
        '        invs.totalValue = 24.8
        '        invs.totalNetValue = 20#
        '
        '
        '
        '
        '
        '        Dim inv As New INVOICE
        '        inv.Issuer = iss
        '        inv.counterpart = cou
        '        inv.invoiceHeader = he
        '        inv.invoiceDetails.ADD (newitem(0))
        '        inv.invoiceDetails.ADD (newitem(1))
        '        inv.paymentMethods.ADD (pay(0))
        '        inv.invoiceSummary = invs
        '        inv.deliveryAddress = deliv
        '        inv.invoiceSummary = invs
        '
        '
        '
        '
        '        Dim SOU As New Source2
        '        SOU.INVOICE = inv
        '
        '        Dim ex2 As New Example
        '        ex2.externalSystemId = "197"
        '        ex2.Source = SOU
        '
        '

        '  Dim ResultJSON As String = JsonConvert.SerializeObject(SOU).ToString
        ' Debug.Print (ResultJSON) '
        
        ' Dim jsonContent As String = ResultJSON ' " {""externa

        Dim REQ As WinHttp.WinHttpRequest

100     REQ = New WinHttp.WinHttpRequest

102     Dim URI As String: URI = gfURL1 + "api/send" '?Bearer=" + gf_Bearer

        Dim Q   As String

104     With REQ
            '            .Open("POST", uri, Async:=False)
            '            .SetRequestHeader("Content-Type", "application/hal+json")
            '            .SetRequestHeader("Accept", "text/*, application/hal+json, application/json")
            '            .SetRequestHeader("Authorization", "Bearer " + gf_Bearer)
            '            q = " { ""Bearer"":""" + gf_Bearer + """, ""source"":" & ResultJSON & ",     ""externalSystemId"": ""198"" }"

106         .send (Q)
108         MsgBox (REQ.responseText)
110         MsgBox (REQ.Status)
        End With

        '        Dim ResultJSON12 As String = " { ""externalSystemId"" : ""185"",  ""source"": {
        '        ""invoice"": {
        '            ""issuer"": {
        '                ""vatNumber"": ""028783755"",
        '                ""branch"": 0,
        '                ""street"": ""street"",
        '                ""streetNumber"": ""streetNumber"",
        '                ""postalCode"": ""postalCode"",
        '                ""city"": ""city"",
        '                ""country"": ""GR""
        '            },
        '            ""counterpart"": {
        '                ""vatNumber"": ""028783755"",
        '                ""branch"": 0,
        '                ""name"": ""name"",
        '                ""street"": ""street"",
        '                ""streetNumber"": ""streetNumber"",
        '                ""postalCode"": ""postalCode"",
        '                ""city"": ""city"",
        '                ""country"": ""GR"",
        '                ""email"": ""test@email.com""
        '            },
        '            ""representative"": {
        '                ""vatNumber"": ""000000000"",
        '                ""name"": ""name""
        '            },
        '            ""deliveryAddress"": {
        '                ""street"": ""street"",
        '                ""streetNumber"": ""streetNumber"",
        '                ""postalCode"": ""postalCode"",
        '                ""city"": ""city"",
        '                ""country"": ""GR""
        '            },

        '            ""invoiceHeader"": {
        '                ""series"": ""INV"",
        '                ""aa"": 13,
        '                ""issueDate"": ""2024-04-14"",
        '                ""dispatchDate"": ""2024-04-14"",
        '                ""invoiceCode"": ""INV-00013"",
        '                ""invoiceType"": ""1.1"",
        '                ""invoiceTypeUbl"": ""380"",
        '                ""currency"": ""EUR"",
        '                ""selfPricing"": false,
        '                ""movePurpose"": 1,
        '                ""fuelInvoice"": false
        '            },
        '            ""invoiceDetails"": [
        '                {
        '                    ""lineNumber"": 1,
        '                    ""recType"": 1,
        '                    ""quantity"": 1,
        '                    ""entityName"": ""string"",
        '                    ""invoiceDetailType"": 0,
        '                    ""netValue"": 100,
        '                    ""totalValue"": 124,
        '                    ""vatCategory"":  ""24"",

        '                    ""vatAmount"": 24,
        '                    ""vatPercent"": 24,
        '                    ""measurementUnit"": ""H87"",
        '                    ""measurementUnitUbl"": ""H87"",
        '                    ""lineComments"": ""S?????"",
        '                    ""classificationCategory"": ""category1_1"",
        '                    ""classificationType"": ""E3_561_1"",

        '                    ""cpvCode"": ""45233222-1""

        '                }
        '            ],

        '            ""paymentMethods"": [
        '                {
        '                    ""type"": 5,
        '                    ""amount"": 124
        '                }
        '            ],

        '            ""invoiceSummary"": {
        '                ""totalNetValue"": 100,
        '                ""totalVatAmount"": 24,
        '                ""totalValue"": 124
        '            },
        '            ""Messages"": [
        '                {
        '                    ""type"": 0,
        '                    ""recipients"": ""test1@email.com;test2@email.com"",
        '                    ""cc"": ""test3@email.com;test4@email.com"",
        '                    ""templateIdentifier"": ""0000001""
        '                },
        '                {
        '                    ""type"": 1,
        '                    ""recipients"": ""6983022568;6983022868""
        '                }
        '            ]
        '        }
        '    }
        '}"

        '""taxInfo"": {
        '               ""taxCategory"": 1,
        '               ""taxCategoryUbl"": ""50"",
        '               ""underlyingValue"": 100,
        '               ""taxPercent"": 20
        '           }

        ' AYTO XREIAZETAI GIA TA TICKETS
        'client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "YOUR_ACCESS_TOKEN")

        '=========================================================================================

        'var client = new JsonServiceClient("http://192.168.0.5:12345/hello");
        '   Try
        '{
        '   HelloResponse response = await client.GetAsync(new Hello{ Name = Username ?? "" });
        '}

        '??   Dim jsonContent As String = " {""syncAll"":0, ""syncTables"": [1,2,3,8,9,10,11,12],""devicesUuid"":""ERP"",""lastSyncDate"":null,""fetchAllUsers"" : true }"
        'Dim jsonContent As String = "{""+ ResultJSON2+"" }" ' " {""externalSystemId"":""185"", ""source"" : """ + ResultJSON + """ }"
        ' Dim jsonContent As String = ResultJSON ' " {""externa

        ' Dim jsonContent As String = "{""syncAll"":""0"", ""syncTables"": ""[1,2,3,8,9,10,11,12"",""devicesUuid"":""ERP"",""lastSyncDate"":""null"",""fetchAllUsers"" : ""true"" }"

        '  Dim content As New StringContent(jsonContent, Encoding.UTF8, "application/json")

        'client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", fbearer)

        'uri = "https://" + PAROX_URLADD + "srv.parochos.gr/api/send"

        ' Dim response = Await client.PostAsync(uri, content)
        '  Dim result = Await response.Content.ReadAsStringAsync()
        '  TextBox2.Text = result.ToString

        ' Exit Sub

        'Dim MF = "c:\txtfiles\apanttimPAROCHOS.xml"  'Inv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        'FileOpen(1, MF, OpenMode.Output)
        '    PrintLine(1, result.ToString)
        'FileClose(1)

        'Dim json As String = "" ' result.ToString
        'Dim ser As JObject = JObject.Parse(json)
        ' Dim c As String = ""

        ' For K As Integer = 0 To ser("jwt").Count - 1

        '  fbearer = ser("jwt")(K)("userId").ToString

        ' Next

End Sub

Private Sub monoApestalmena_Click()

        '<EhHeader>
        On Error GoTo monoApestalmena_Click_Err

        '</EhHeader>
        Dim sql As String

        Dim pol As String, polepis As String, ago As String, AGOEPIS As String

100     pol = "": polepis = "": ago = "": AGOEPIS = ""

        Dim PARASPAROX As String

102     PARASPAROX = ""
        
        Dim DUM As Boolean
  
104     If monoApestalmena.Value = vbUnchecked Then
    
106         sql = FIND_QUERY()

108         Adodc1.ConnectionString = gConnect
110         Adodc1.RecordSource = sql
112         Adodc1.Refresh
  
            Exit Sub

        End If
  
114     DUM = Get_AJ_ASCII(pol, polepis, ago, AGOEPIS, PARASPAROX)

        Dim par, synt As String
        
116     If Len(polepis) = 0 Then polepis = "'~'"

118     par = pol + "," + polepis

120     If Len(ago) > 0 Then
        
122         par = par + "," + ago
        End If

124     synt = ""
126     sql = "SELECT ENTITYMARK, HME,ATIM,PEL.EPO,PEL.AFM,KPE,ISNULL(PEL.DIE,'') AS DIE,ENTITYUID,ID_NUM, AJ1  ,AJ2 , AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,"
128     sql = sql + ""    '"CONVERT(CHAR(10),HME,3) AS HMEP
130     sql = sql + "PEL.EPA,PEL.POL,AJ6,FPA6,AJ7,FPA7,TRP,ISNULL(APALAGIFPA,0) AS APALAGIFPA ,ISNULL(PEL.XRVMA,'') AS TK "

132     sql = sql + "   FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
134     sql = sql + " WHERE     LEFT(ATIM,1) IN     (  " + par + "  )    and HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
136     sql = sql + "  AND AJ1+AJ2+AJ3+AJ4+isnull(AJ5,0)+isnull(AJ6,0)+isnull(AJ7,0)>0  " + synt
138     sql = sql + " order by HME"       '  OR INCMARK IS NULL OR INCMARK='ERROR'

140     Adodc1.RecordSource = sql
142     Adodc1.Refresh

144     synt = " and len(ENTITYMARK)>5 "
146    sql = "SELECT SUM( AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7) AS [ΣΥΝ.ΚΑΘ],SUM( FPA1+FPA2+FPA3+FPA4+FPA6+FPA7) AS [ΣΥΝ.ΦΠΑ],CONVERT(CHAR(10),HME,103) AS HMEP "

148    sql = sql + "   FROM TIM  "
150    sql = sql + " WHERE     LEFT(ATIM,1) IN     (  " + par + "  )    and HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
152    sql = sql + "  AND AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7>0  " + synt
154    sql = sql + " GROUP BY HME order by HME"       '  OR INCMARK IS NULL OR INCMARK='ERROR'

156     AdoSYN.ConnectionString = gConnect
158    AdoSYN.RecordSource = sql
160    AdoSYN.Refresh

162     suma_hmeron

        '<EhFooter>
        Exit Sub

monoApestalmena_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.monoApestalmena_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub MYDPAR_Change()
          Dim SQL2 As String

106     SQL2 = FIND_QUERY()
108     Adodc1.RecordSource = SQL2
110     Adodc1.Refresh
     
End Sub

Private Sub requestafm_Click()
    ' requestafm2
   
End Sub

Private Sub RESEND_CANCEL_Click()

        Dim c, D, SX As String

100     D = InputBox("δωσε τον αριθμό του id", 0)
   
102     SX = GET_CVALUE("select ISNULL(SXETMARK,'00') AS SXX from TIM WHERE ID_NUM=" + D)

104     If SX = "00" Then
106         MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΜΑΡΚ ΠΟΥ ΘΑ ΑΚΥΡΩΘΕΙ"

            Exit Sub

        End If

        Dim TIMOL As String

108     TIMOL = GET_CVALUE("select ATIM from TIM WHERE ID_NUM=" + D)
   
110     c = CANCEL_INVOICE(SX, TIMOL)

112     If Len(c) > 5 Then
114         Gdb.Execute "update TIM set ENTITYMARK='" + c + "' WHERE ID_NUM=" + D
        End If

End Sub

Private Sub RESTORELATHOS_Click()

        '<EhHeader>
        On Error GoTo RESTORELATHOS_Click_Err

        '</EhHeader>

100     Gdb.Execute "UPDATE TIM SET ENTITYMARK=NULL WHERE ENTITYMARK LIKE '%ΛΑΘ%'"

        '<EhFooter>
        Exit Sub

RESTORELATHOS_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.RESTORELATHOS_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub SENDONEBYONE_Click()

        Dim sql              As String

        Dim mgisOnlineMydata As Integer

100     mgisOnlineMydata = gisOnlineMydata
102     sql = FIND_QUERY()
  
        Dim r85 As New ADODB.Recordset
  
104     r85.Open sql, Gdb, adOpenDynamic, adLockOptimistic
  
106     gisOnlineMydata = 1

        Dim nm As Long

108     nm = 0

110     Do While Not r85.EOF
112         where.Caption = " where ID_NUM=" + str(r85!id_num) + "  " ' M_idnum
114         List1.AddItem r85!ATIM

116         DoEvents
118         ToXMLsub 1, "" ' 1 means no msgbox
120         nm = nm + 1
122         List1.AddItem Format(r85!hme, "dd/MM/yyyy") + "--" + Format(nm, "####0")
124         r85.MoveNext
        Loop
 
126     r85.Close
 
128     gisOnlineMydata = mgisOnlineMydata

End Sub

Private Sub SXETIKO_Click()

        Dim CC As String

100     CC = (Adodc1.Recordset("atim"))

102     CC = FINDTYPOS(Left(CC, 1))

104     If Left(CC, 3) = "5.1" Then  ' pistvtiko διαλεγω το σχετιζομενο παραστατικο

            'ΓΙΑ ΝΑ ΚΑΤΑΛΑΒΕΙ ΟΤΙ ΤΟ ΚΑΛΕΣΕ ΤΟ PAR7MYDATA
106         par1b.Label1.Caption = "1"
108         par1b.PEL.Text = Trim(Adodc1.Recordset("KPE")) ' Text2(0)
110         par1b.SHOW 1
   
112         Gdb.Execute "UPDATE TIM SET SXETMARK='" + SXETMART.Caption + "'  WHERE ID_NUM= " + str(Adodc1.Recordset("ID_NUM"))
   
        End If

End Sub

Private Sub TDBGrid1_Click()
    '
    '   Dim a As String
    '   a = 1
    '140        TDBGrid1.Col = 3
    '150    a = TDBGrid1.Text

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid1_HeadClick_Err

        '</EhHeader>

        '--------------------------------------------------------------------------------------=
        Dim sumes(100) As String

100     For k = 0 To Adodc1.Recordset.FIELDS.Count - 1
102         sumes(k) = TDBGrid1.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
        Next

104     If Adodc1.Recordset.sort = "[" & TDBGrid1.columns(ColIndex).DataField & "] asc" Then   ' strSort
106         Adodc1.Recordset.sort = "[" & TDBGrid1.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
108         Adodc1.Recordset.sort = "[" & TDBGrid1.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

110     For k = 0 To Adodc1.Recordset.FIELDS.Count - 1
112         TDBGrid1.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
        Next

        '<EhFooter>
        Exit Sub

TDBGrid1_HeadClick_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.TDBGrid1_HeadClick " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub TESTXARAL_Click()
On Error GoTo OUT

        'ΣΒΗΝΕΙ ΤΙΣ ΔΙΠΛΟΕΓΓΡΑΦΕΣ ΠΟΥ ΔΗΛΩΘΗΚΑΝ ΣΤΟ ΚΕΝΤΡΙΚΟ ΕΝΩ ΗΤΑΝ ΥΠΟΚΑΤΑΣΤΗΜΑ
Dim KPE As String, HM1 As String, HM2 As String
KPE = InputBox("ΔΩΣΕ ΚΩΔΙΚΟ ΠΕΛΑΤΗ")
HM1 = InputBox("ΑΠΟ ηη/MM/εεεε", "", "01/01/2025")
HM2 = InputBox("ΕΩΣ ηη/MM/εεεε", "", "31/01/2025")

HM1 = mID(HM1, 4, 2) + "/" + Left$(HM1, 2) + "/" + mID$(HM1, 7, 4)
HM2 = mID(HM2, 4, 2) + "/" + Left$(HM2, 2) + "/" + mID$(HM2, 7, 4)

Dim sql As String
On Error Resume Next
Gdb.Execute "DROP TABLE DOKTIMCANCEL"
On Error GoTo OUT
sql = "select ENTITYMARK,ATIM,HME,KPE,AJI,ID_NUM,'                ' AS CANCELATIONMARK,'                ' AS NEOMARK  INTO DOKTIMCANCEL from TIM where KPE='" + KPE + "' and HME>='" + HM1 + "' AND HME<='" + HM2 + "'"
Gdb.Execute sql
Dim a, B

100

Dim R As New ADODB.Recordset
R.Open "SELECT COUNT(*) FROM DOKTIMCANCEL", Gdb, adOpenDynamic, adLockOptimistic
Dim num As Integer
num = R(0)
R.Close
If num = 0 Then
  MsgBox "ΔΕΝ ΒΡΕΘΗΚΑΝ ΕΓΓΡΑΦΕΣ"
  Exit Sub
End If


Gdb.Execute "UPDATE DOKTIMCANCEL SET ENTITYMARK=SUBSTRING(ENTITYMARK,4,15) WHERE ENTITYMARK LIKE 'ΔΟΚ%' "
B = InputBox("Προσοχή θα ακυρωθOYN " + str(num) + "EGGRFAES " + Chr(13) + " ΑΠΟ TABLE DOKTIMCANCEL" + HM1 + " " + HM2 + " " + KPE + " .Δωσε κωδικό αξουσιοδότησης")

102     If B = "3921" Then
            ' a = InputBox("δωσε μαρκ που θα ακυρωθεί")
            ' C = CANCEL_INVOICE(Trim(a), Trim(a))
            ' MsgBox "ok"
        Else
104         MsgBox "λαθος"
            Exit Sub

        End If

        On Error GoTo 0
 
        Dim c, D, SX As String

        'Dim R2 As New ADODB.Recordset
 
106     R.Open "SELECT ATIM,ENTITYMARK,CANCELATIONMARK,KPE,ID_NUM FROM  DOKTIMCANCEL", Gdb, adOpenDynamic, adLockOptimistic  '  WHERE ISSUEDATE>='" + Format(APO.Value, "yyyy-MM-dd") + "'  AND ISSUEDATE<'" + Format(DateAdd("d", 1, EOS.Value), "yyyy-MM-dd") + "' ORDER BY MARK"
            Dim TIMOL As String
            Dim id_num As Integer
            Dim NEOMARK As String
          
108     Do While Not R.EOF
  
            'd = InputBox("δωσε τον αριθμό του id", 0)
   
            ' SX = GET_CVALUE("select ISNULL(SXETMARK,'00') AS SXX from TIM WHERE ID_NUM=" + d)
            'If SX = "00" Then
            '  MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΜΑΡΚ ΠΟΥ ΘΑ ΑΚΥΡΩΘΕΙ"
            ' Exit Sub
            ' End If
           
            id_num = R!id_num

110         TIMOL = R!ATIM    ' AA   '    GET_CVALUE("select ATIM from TIM WHERE ID_NUM=" + d)
112         SX = R!entitymark   '   R!MARK
            'List1.AddItem
   
114         c = CANCEL_INVOICE(Trim(SX), TIMOL)

116         If Len(c) > 5 Then


                 Gdb.Execute "update TIM set C12='" + c + "',C13='" + SX + "' WHERE ID_NUM=" + str(id_num)
      
                 Gdb.Execute "update DOKTIMCANCEL set CANCELATIONMARK='" + c + "' WHERE ID_NUM=" + str(id_num)
118             List1.AddItem SX + " AK " + c


                           ' loadpar7
                             
                             gisOnlineMydata = 1
                          where.Caption = " where ID_NUM=" + str(id_num) + "  " ' M_idnum
                           NEOMARK = "000000000000000"
                          ToXMLsub 0, NEOMARK

            End If   '
            
            'loadpar7
            ' NEOMARK = "000000000000000"
                          'ToXMLsub 0, NEOMARK

            

120        ' Write #3, SX + "/" + c
   
122         R.MoveNext
   
        Loop
   
124     Close #3

Exit Sub
OUT:
   MsgBox "ΛΑΘΟΣ" + Err.Description
   
End Sub

Private Sub Toolbar1_ButtonDropDown(ByVal Button As MSComctlLib.Button)

100     If Button.index = 3 Then
            ' remove existing buttons
            '  Button.ButtonMenus.Clear

            'add two ButtonMenu to Button 3:
            '  Toolbar1.Buttons(3).ButtonMenus.ADD , "btn1", "ButtonMenu1"
            '  Toolbar1.Buttons(3).ButtonMenus.ADD , "btn2", "ButtonMenu2"
  
        End If
  
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

        '  Private Sub xxToolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
100     Select Case Button.key

            Case "SYSX"
102             SXETIKO_Click

104         Case "EPAN"
106             RESEND_CANCEL_Click

108         Case "GIAEPAN"
110             cmdMARK_Click

112         Case "ONEBYONE"
114             SENDONEBYONE_Click
    
116         Case "XML"
118             Command11_Click

120         Case "KINHSEIS"
                'kinhseis_click
122             MsgBox "διαλεξτε πελάτη ή προμηθευτή"
    
        End Select
    
        '     If Button.Index = 3 Then
        '  ' remove existing buttons
        '        Button.ButtonMenus.Clear
        '
        '
        '        'add two ButtonMenu to Button 3:
        '        Toolbar1.Buttons(3).ButtonMenus.ADD , "btn1", "ButtonMenu1"
        '        Toolbar1.Buttons(3).ButtonMenus.ADD , "btn2", "ButtonMenu2"
        '
        '  End If
    
End Sub
    
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  
        Dim N As Integer

100     N = ButtonMenu.index

102     Select Case ButtonMenu.key

            Case "diagdok"
104             cmddiagdokimi_Click

106         Case "epanfile"
                ' Else
108             cmdepanfromfile_Click

110         Case "cmdDeleteMark"
112             deletemarK

114         Case "kinpel"
        
116             KINPEL_CLICK 1

118         Case "kinprom"
120             KINPEL_CLICK 2
        
122         Case "EmporAney"
        
124             EmporAney

126         Case "MydataAney"
128             MydataAney
        
        End Select
  
End Sub

Private Sub MydataAney()
        'Dim sql As String
100     sql = "select  (select TOP 1 TITLOS FROM PARASTAT WHERE MYEID= RTRIM(SUBSTRING(A.TYPOS,1,4))+'%') AS [ΠΑΡΑΣΤ],P.EPO,ATIM,MARK,TOTALNETVALUE AS [ΚΑΘ.ΑΞΙΑ], A.* from APESTALMENA A LEFT JOIN PEL P ON P.AFM=A.AFM AND P.EIDOS='r' "
102     sql = sql + " WHERE (ISANTIST IS NULL ) AND ISSUEDATE>='" + Format(APO.Value, "yyyy-MM-dd") + "'  AND  ISSUEDATE<'" + Format(DateAdd("d", 1, EOS.Value), "yyyy-MM-dd") + "'  ORDER BY ISSUEDATE DESC"
        '"SELECT * FROM APESTALMENA WHERE issueDate>='" + Format(apo.Value, "MM/dd/yyyy") + "'  AND issueDate<'" + Format(DateAdd("d", 1, eos.Value), "MM/dd/yyyy") + "' AND ISANTIST IS NULL order by MARK DESC"

104     Adodc2.ConnectionString = gConnect
106     Adodc2.RecordSource = sql
108     Adodc2.Refresh

End Sub

Private Sub EmporAney()

100     sql = "select  ATIM,CONVERT(CHAR(10),HME,103) AS HMEP,P.EPO from TIM A LEFT JOIN PEL P ON P.KOD=A.KPE AND P.EIDOS='r' "
102     sql = sql + " WHERE LEFT(ATIM,1) IN ('G','g','Ξ') and (AADEKAU IS NULL) AND HME>='" + Format(APO.Value, "yyyy-MM-dd") + "'  AND HME<'" + Format(DateAdd("d", 1, EOS.Value), "yyyy-MM-dd") + "'  ORDER BY HME DESC"
        '"SELECT * FROM APESTALMENA WHERE issueDate>='" + Format(apo.Value, "MM/dd/yyyy") + "'  AND issueDate<'" + Format(DateAdd("d", 1, eos.Value), "MM/dd/yyyy") + "' AND ISANTIST IS NULL order by MARK DESC"

104     Adodc2.ConnectionString = gConnect
106     Adodc2.RecordSource = sql
108     Adodc2.Refresh

End Sub

Private Sub updateclient_Click()
        '<?xml version="1.0" encoding="utf-8"?>
        '<dcrudt:UpdateClientDoc xmlns:dcrudt="http://www.aade.gr/myDATA/dcrudt/v1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.aade.gr/myDATA/dcrudt/updateClientType-v1.0.xsd">
        '  <dcrudt:updateClient>
        '    <dcrudt:initialDclId>2</dcrudt:initialDclId>
        '
        '  </dcrudt:updateClient>
        '</dcrudt:UpdateClientDoc>

        Dim ANS    As String, TYPOS As String

        Dim ARKYKL As String

100     ANS = InputBox("2=PARKING 3=GARAGE", "")

102     ARKYKL = InputBox("αρ.κυκλοφοριας", "")

104     If Val(ANS) = 2 Then
106         TYPOS = "dcrnew:parkingcarwash"
        Else
108         TYPOS = "dcrnew:garage"
        End If

        Dim DUM As String

110    ' DUM = fupdateclient("100000000006453", "2", "202502221033")

End Sub

Public Function OLD_fupdateclient(ByVal clientID As String, _
                              ByVal ServiceType As String, _
                              ByVal completionDateTime As String) As String '"202502221033
   
        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
100     Set docStock = New MSXML2.DOMDocument

102     With docStock
104         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
106         Set elemRoot = .createElement("dcrudt:UpdateClientDoc")
108         elemRoot.setAttribute "xmlns:dcrudt", "https://www.aade.gr/myDATA/dcrudt/v1.0"
110         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
112         elemRoot.setAttribute "xsi:schemaLocation", "https://www.aade.gr/myDATA/dcrudt/v1.0 updateClientType-v1.0.xsd updateClient-v1.0.xsd"

114         Set elemStock = docStock.createElement("dcrudt:updateClient")
            ' With elemStock----------------------------------  100000000006453
116         Set elem2Field = docStock.createElement("dcrudt:initialDclId"): elem2Field.Text = clientID: elemStock.appendChild elem2Field
118         Set elem2Field = docStock.createElement("dcrudt:clientServiceType"): elem2Field.Text = ServiceType: elemStock.appendChild elem2Field
120         Set elem2Field = docStock.createElement("dcrudt:entryCompletion"): elem2Field.Text = "true": elemStock.appendChild elem2Field
            'ProvidedCategoryServiceType
            ' Set elem2Field = docStock.createElement("dcrudt:ProvidedCategoryServiceType"): elem2Field.Text = "3": elemStock.appendChild elem2Field
                      
            ' Set elem2Field = docStock.createElement("dcrudt:amount"): elem2Field.Text = 10#: elemStock.appendChild elem2Field
                      
            ' Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = completionDateTime: elemStock.appendChild elem2Field
                       

            'elemStock.appendChild elemField
                    
122         elemRoot.appendChild elemStock
  
            '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
124         Set .documentElement = elemRoot

            On Error Resume Next
            
126         FormatXmlDocument docStock ' βαζει κενα να ειναι ευκολο στο διαβασμα

            Dim ANS2 As Integer

128         ANS2 = MsgBox(docStock.XML, vbYesNo)

130         If ANS2 = vbNo Then

                Exit Function

            End If
            
132         Kill "C:\txtfiles\" + clientID + ".xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

134         .save "C:\txtfiles\" + clientID + ".xml"
        End With

136     FileCopy "C:\txtfiles\" + clientID + ".xml", "c:\" + F_TXTFILES + "\sendinv\update-aytok" + "-" + clientID + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"

        ' TO XML EINAI docStock.xml

        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
     
        Dim MARK As String

        '    MARK = InputBox("δωσε τo mark", "μαρκ")
     
138     url2 = url2 + MARKTIM
     
140     url2 = UrlAADE + "DCL/UpdateClient"

        'initialize
142     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
144     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
146     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
148     XMLServer.Option(9) = 2048
150     XMLServer.Option(6) = True
    
152     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
154     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
156     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

158     XMLServer.send docStock.XML 'TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
160     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
162     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
 
164     MsgBox v

        Dim CC As String

166     CC = takefromxml(v)
168     fupdateclient = CC

170     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

172     mfile = "c:\" + F_TXTFILES + "\sendinv\apant-update-aytok" + "-" + clientID + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
  
        ' Dim mfileAsk As String
        ' mfileAsk = "C:\txtfiles\" + clientID + ".xml"
174     Debug.Print v
176     Open mfile For Output As #8
178     Print #8, v
180     Close #8
    
        '  FileCopy mfile, "C:\" + F_TXTFILES + "\apantupdate.XML"
    
        Dim MARK2 As String
    
        '  READ_CANCEL v, MARK2

        '  400001829337674

        '    CANCEL_INVOICE = MARK2
    
        '<EhFooter>
        Exit Function

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdDOMDocumentUTF8_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Private Sub UPLOAD_Click()
    '  UPLOAD2
       
End Sub

Function chDec(c As String) As String

        '<EhHeader>
        On Error GoTo chDec_Err

        '</EhHeader>

100     chDec = Replace(c, ",", ".")

        '<EhFooter>
        Exit Function

chDec_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.chDec " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Public Function checkIntegrity(ByVal i As Long, _
                                ByVal SQLDT As ADODB.Recordset, _
                                ByRef PROBL As String) As Boolean
  
        'ELEGXOS AN SUM(KAU_AJIA) TOY EGGTIM ΣΥΜΦΩΝΕΙ ΜΕ TA AJ1,AJ2,.. TOY TIM ΑΝ ΟΧΙ ΒΑΖΕΙ ΤΗΝ ΔΙΑΦΟΡΑ ΣΤΟ EGGTIM GIA NA SYMFVNHSEI
  
        '<EhHeader>
        On Error GoTo checkIntegrity_Err

        '</EhHeader>

        'If checkIntegrity(I) = False Then

        '    MsgBox(" ΠΡΟΒΛΗΜΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + sqlDT("ATIM").ToString)
        'End If

        Dim SQLDOS As String

        Dim nERR   As Long

100     nERR = 0 ' ποσες φορες εκανε error=> resume next

        Dim DIFF1 As Double

102     Dim S_AJ  As Double: S_AJ = SQLDT("aj1") + SQLDT("aj2") + SQLDT("aj3") + SQLDT("aj4") + SQLDT("aj5") + SQLDT("aj6") + SQLDT("aj7")

104     Dim S_FPA As Double: S_FPA = SQLDT("fpa1") + SQLDT("fpa2") + SQLDT("FPA3") + SQLDT("FPA4") + SQLDT("FPA6") + SQLDT("FPA7")

106     diff = SQLDT("aji") - (S_AJ + S_FPA)

108     If diff = 0 Then
            'OK
110     ElseIf Math.Abs(diff) < 0.08 Then ' ΕΧΟΥΝ ΜΙΚΡΗ ΔΙΑΦΟΡΑ ΚΑΙ ΤΗΝ ΚΑΛΥΠΤΩ ΑΛΛΑΖΟΝΤΑ ΤΟ AJI
            ' ετσι βγαζω διαφορά στο mydata σε σχαση με το χαρτι του τιμολογιου
                


            Dim n21   As Integer

            Dim max21 As Single

            Dim maxaj As String

112         max21 = 0: maxaj = ""

            ' βρισκω το μεγαλυτερο ποσο για να βαλω την διαφορα
114         For n21 = 1 To 7

116             If SQLDT("aj" + Format(n21, "0")) > max21 Then
118                 maxaj = "aj" + Format(n21, "0")
                   
                End If
                
            Next

120         If Len(maxaj) > 1 Then
                'Gdb.Execute "UPDATE TIM SET " + maxaj + "=" + maxaj + "+" + Replace(str(Diff), ",", ".") + "   WHERE ID_NUM=" + str(sqlDt("ID_NUM"))
                
122             Gdb.Execute "UPDATE TIM SET " + UCase(maxaj) + "=" + UCase(maxaj) + "+" + Replace(str(diff), ",", ".") + "   WHERE ID_NUM=" + str(SQLDT("ID_NUM"))
                
            End If

        Else
            ' Dim DUM As New DataTable
            ' DUM11.OPEN "SELECT SUM(POSO*TIMM*(100-EKPT)/100)   WHERE ID_NUM=" + STR(SQLDT("ID_NUM")), DUM)
124         checkIntegrity = False

            Exit Function

        End If

        '    Dim DUM150 As New DataTable
        ' GDB.EXECUTE "UPDATE EGGTIM
        '        SET MIK_AJIA=ROUND(POSO*TIMM*(100-EKPT)/100*1.24,2)
        '     ,KAU_AJIA=ROUND(POSO*TIMM*(100-EKPT)/100,2)  ", DUM150)

        Dim DUM11 As New ADODB.Recordset

        Dim DUM10 As New ADODB.Recordset
            
        Dim SQL10 As String

126     SQL10 = "SELECT SUM(ISNULL(KAU_AJIA,0)) AS KAU,SUM(ISNULL(MIK_AJIA,0)) AS MIK," _
           & " SUM(CASE WHEN FPA=1 THEN ISNULL(KAU_AJIA,0) ELSE 0 END ) AS KAU1, " _
           & " SUM(CASE WHEN FPA=2 THEN ISNULL(KAU_AJIA,0) ELSE 0 END ) AS KAU2, " _
           & " SUM(CASE WHEN FPA=3 THEN ISNULL(KAU_AJIA,0) ELSE 0 END ) AS KAU3, " _
           & " SUM(CASE WHEN FPA=4 THEN ISNULL(KAU_AJIA,0) ELSE 0 END ) AS KAU4, " _
           & " SUM(CASE WHEN FPA=5 THEN ISNULL(KAU_AJIA,0) ELSE 0 END ) AS KAU5, " _
           & " SUM(CASE WHEN FPA=6 THEN ISNULL(KAU_AJIA,0) ELSE 0 END ) AS KAU6, " _
           & " SUM(CASE WHEN FPA=7 THEN ISNULL(KAU_AJIA,0) ELSE 0 END ) AS KAU7, " _
           & " SUM(CASE WHEN FPA=1 THEN ISNULL(MIK_AJIA,0) ELSE 0 END ) AS MIK1, " _
           & " SUM(CASE WHEN FPA=2 THEN ISNULL(MIK_AJIA,0) ELSE 0 END ) AS MIK2, " _
           & " SUM(CASE WHEN FPA=3 THEN ISNULL(MIK_AJIA,0) ELSE 0 END ) AS MIK3, " _
           & " SUM(CASE WHEN FPA=4 THEN ISNULL(MIK_AJIA,0) ELSE 0 END ) AS MIK4, " _
           & " SUM(CASE WHEN FPA=5 THEN ISNULL(MIK_AJIA,0) ELSE 0 END ) AS MIK5, " _
           & " SUM(CASE WHEN FPA=6 THEN ISNULL(MIK_AJIA,0) ELSE 0 END ) AS MIK6, " _
           & " SUM(CASE WHEN FPA=7 THEN ISNULL(MIK_AJIA,0) ELSE 0 END ) AS MIK7 " _
           & " FROM EGGTIM    WHERE ID_NUM=" + str(SQLDT("ID_NUM"))
            
128     DUM10.Open SQL10, Gdb, adOpenDynamic, adLockOptimistic

130     If IsNull(DUM10("KAU")) Then
132         checkIntegrity = False

            Exit Function

        End If
             
134     If (DUM10("KAU")) = 0 Then ' ΕΠΑΝΥΠΟΛΟΓΙΣΜΟΣ KAU_AJIA
            ' Gdb.Execute " UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA=POSO*TIMM*(100-EKPT)/100*(100+(SELECT TOP 1 TIMH FROM PINAKES WHERE TYPOS=1 AND AYJON=FPA))/100    WHERE ID_NUM=" + str(sqlDt("ID_NUM"))
              
136         Gdb.Execute " UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*((100-EKPT2)/100)*(100-EKPT)/100,MIK_AJIA=POSO*TIMM*((100-EKPT2)/100)*(100-EKPT)/100*(100+(SELECT TOP 1 TIMH FROM PINAKES WHERE TYPOS=1 AND AYJON=FPA))/100    WHERE ID_NUM=" + str(SQLDT("ID_NUM"))
                     
138         DUM10.Close
                   
140         DUM10.Open SQL10, Gdb, adOpenDynamic, adLockOptimistic
              
        End If
             
142     If fMydataFromEID1 = 1 Then  'ΕΝΗΜΕΡΩΣΗ ΑΠΟ ΚΑΤΗΓΟΡΙΑ ΕΙΔΩΝ

            '  SELECT SUM(KAU_AJIA) AS AJIA,G.FPA,KATHGORIA  FROM EGGTIM G  INNER JOIN EID D  ON G.KODE=D.KOD
            '  WHERE ATIM='T000592' GROUP BY G.FPA,KATHGORIA
            Dim C56 As String, R56 As New ADODB.Recordset

144         C56 = "SELECT KATHGORIA,KOD,ATIM "
146         C56 = C56 + " FROM EGGTIM G  INNER JOIN EID D  ON G.KODE=D.KOD "
148         C56 = C56 + " WHERE ( (KATHGORIA IS NULL) OR KATHGORIA>5 ) AND  TIMM<>0 AND POSO<>0 AND ID_NUM=" + str(SQLDT("ID_NUM"))
150         C56 = C56 + "  "
152         R56.Open C56, Gdb, adOpenDynamic, adLockOptimistic

            Dim mERR As Integer

154         mERR = nERR

156         Do While Not R56.EOF
                        
158             List1.AddItem R56!kod + " ΔΕΝ ΕΧΕΙ ΚΑΤΗΓΟΡΙΑ ΠΑΡΑΣΤ:" + SQLDT("atim")
                         
160             R56.MoveNext

162             If mERR - nERR > 1 Then  ' για να μην κολλησει το προγραμμα
164                 PROBL = "ΠΡΟΒΛΗΜΑ ΜΕ ΚΑΤΗΓΟΡΙΕΣ ΠΑΡΑΣΤ:" + SQLDT("atim")
166                 checkIntegrity = False

                    Exit Function

                End If

            Loop
                    
168         R56.Close
            'PROBL = "ΕΙΔΗ ΧΩΡΙΣ ΚΑΤΗΓ στο " + sqlDt("atim")
        End If

        'ΚΑΘΑΡΕΣ ΑΞΙΕΣ ΕΛΕΓΧΟΣ
170     If Abs(Round(DUM10("KAU"), 2) - Round(S_AJ, 2)) < 0.002 Then 'OK   SYMFVNEI Η ΚΑΘΑΡΗ ΑΞΙΑ
            'OK

        Else ' ΚΑΛΥΠΤΩ ΤΗΝ ΔΙΑΦΟΡΑ EGGTIM - TIM ΑΝΑ ΦΠΑ

172         If Math.Abs(Round(SQLDT("aj1"), 2) - Round(DUM10("KAU1"), 2)) > 0.005 Then
                
174             DIFF1 = SQLDT("aj1") - DUM10("KAU1")
176             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=1 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
178             Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If

180         If Math.Abs(Round(SQLDT("aj2"), 2) - Round(DUM10("KAU2"), 2)) > 0.005 Then
                
182             DIFF1 = SQLDT("aj2") - DUM10("KAU2")
184             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=2 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
186             Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If

188         If Math.Abs(Round(SQLDT("aj3"), 2) - Round(DUM10("KAU3"), 2)) > 0.005 Then
                
190             DIFF1 = SQLDT("aj3") - DUM10("KAU3")
192             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=3 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
194             Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If

196         If Math.Abs(Round(SQLDT("aj4"), 2) - Round(DUM10("KAU4"), 2)) > 0.005 Then
                
198             DIFF1 = SQLDT("aj4") - DUM10("KAU4")
200             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=4 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
202             Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If

204         If Math.Abs(Round(SQLDT("aj5"), 2) - Round(DUM10("KAU5"), 2)) > 0.005 Then
                
206             DIFF1 = SQLDT("aj5") - DUM10("KAU5")

208             If f_apoDOS = 1 Then ' ερχεται απο DOS ΟΠΟΤΕ ΤΟΥ ΛΕΙΠΕΙ Η ΕΓΓΡΑΦΗ ΤΩΝ ΚΕΝΩΝ ΚΑΙ ΤΗΝ ΔΗΜΙΟΥΡΓΩ
210                 SQLDOS = "INSERT INTO EGGTIM(KODE,EIDOS,ATIM,HME,FPA,POSO,XRE,TIMM,KAU_AJIA,MIK_AJIA,ID_NUM) VALUES('99~','e',"
212                 SQLDOS = SQLDOS + "'" + SQLDT("ATIM") + "',"
214                 SQLDOS = SQLDOS + "'" + Format(SQLDT("HME"), "mm/dd/yyyy") + "',5,1,1,"
216                 SQLDOS = SQLDOS + Replace(str(DIFF1), ",", ".") + ","  'TIMM
218                 SQLDOS = SQLDOS + Replace(str(DIFF1), ",", ".") + ","  ' KAU_AJIA
220                 SQLDOS = SQLDOS + Replace(str(DIFF1), ",", ".") + "," + str(SQLDT("ID_NUM")) + ")"  ' MIK_AJIA
222                 Gdb.Execute SQLDOS
                Else

224                 DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=5 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
226                 Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                    DUM11.Close:
                End If
                    
            End If

228         If Math.Abs(Round(SQLDT("aj6"), 2) - Round(DUM10("KAU6"), 2)) > 0.005 Then
                
230             DIFF1 = SQLDT("aj6") - DUM10("KAU6")
232             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=6 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
234             Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If

236         If Math.Abs(Round(SQLDT("aj7"), 2) - Round(DUM10("KAU7"), 2)) > 0.005 Then
                
238             DIFF1 = SQLDT("aj7") - DUM10("KAU7")
240             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=7 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
242             Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If
        End If

        'ΚΑΘΑΡΕΣ MIKTES ΑΞΙΕΣ ΕΛΕΓΧΟΣ
244     If DUM10("MIK") = S_AJ + S_FPA Then 'OK   SYMFVNEI Η ΑΞΙΑ ME FPA
            'OK

        Else ' ΚΑΛΥΠΤΩ ΤΗΝ ΔΙΑΦΟΡΑ EGGTIM - TIM ΑΝΑ ΦΠΑ

246         If Math.Abs(SQLDT("aj1") + SQLDT("FPA1") - DUM10("MIK1")) > 0.005 Then
                
248             DIFF1 = SQLDT("aj1") + SQLDT("FPA1") - DUM10("MIK1")
250             DUM11.Open "SELECT TOP 1 * FROM EGGTIM    WHERE FPA=1 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
252             Gdb.Execute "UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
                    
                '                    ' ΑΝ Η ΔΙΑΦΟΡΑ ΕΙΝΑΙ ΚΑΤΩ ΤΟΥ 0.05 ΤΟΤΕ ΝΑ ΠΡΟΣΑΡΜΟΣΩ ΤΟ ΜΙΚ_ΑJIA
                '                    Dim DIFF_MIK As Single
                '                    DIFF_MIK = (sqlDt("aj1") + sqlDt("FPA1") - GGET_NVALUE("select MIK_AJIA FROM EGGTIM WHERE ID=" + str(DUM11("ID")))) '0.10-0.97 =0.03
                '                     If Math.Abs(DIFF_MIK) < 0.05 Then
                '                         Gdb.Execute "UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+DIFF_MIK " + " WHERE ID=" + str(DUM11("ID"))
                '
                '                     End If
                    
            End If

254         If Math.Abs(SQLDT("aj2") + SQLDT("FPA2") - DUM10("MIK2")) > 0.005 Then
                
256             DIFF1 = SQLDT("aj2") + SQLDT("FPA2") - DUM10("MIK2")
258             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=2 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
260             Gdb.Execute "UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If

262         If Math.Abs(SQLDT("aj3") + SQLDT("FPA3") - DUM10("MIK3")) > 0.005 Then
                
264             DIFF1 = SQLDT("aj3") + SQLDT("FPA3") - DUM10("MIK3")
266             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=3 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
268             Gdb.Execute "UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If

270         If Math.Abs(SQLDT("aj4") + SQLDT("FPA4") - DUM10("MIK4")) > 0.005 Then
                
272             DIFF1 = SQLDT("aj4") + SQLDT("FPA4") - DUM10("MIK4")
274             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=4 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
276             Gdb.Execute "UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If

278         If Math.Abs(SQLDT("aj6") + SQLDT("FPA6") - DUM10("MIK6")) > 0.005 Then
                
280             DIFF1 = SQLDT("aj6") + SQLDT("FPA6") - DUM10("MIK6")
282             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=6 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
284             Gdb.Execute "UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If

286         If Math.Abs(SQLDT("aj7") + SQLDT("FPA7") - DUM10("MIK7")) > 0.005 Then
                
288             DIFF1 = SQLDT("aj7") + SQLDT("FPA7") - DUM10("MIK7")
290             DUM11.Open "SELECT * FROM EGGTIM    WHERE FPA=7 AND ID_NUM=" + str(SQLDT("ID_NUM")) + "ORDER BY KAU_AJIA DESC", Gdb, adOpenDynamic, adLockOptimistic
292             Gdb.Execute "UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(str(DIFF1), ",", ".") + " WHERE ID=" + str(DUM11("ID"))
                DUM11.Close:
            End If

        End If
           
        Dim rkathg As New ADODB.Recordset

294     If fMydataFromEID1 = 1 Then  'ΕΝΗΜΕΡΩΣΗ ΑΠΟ ΚΑΤΗΓΟΡΙΑ ΕΙΔΩΝ
            
296         rkathg.Open "select KODE,EID.KATHGORIA from EGGTIM INNER JOIN EID ON  EID.KOD=EGGTIM.KODE WHERE (EID.KATHGORIA=0 OR EID.KATHGORIA IS NULL ) AND ID_NUM=" + str(SQLDT("ID_NUM")), Gdb, adOpenDynamic, adLockOptimistic

298         If rkathg.EOF Then
300             rkathg.Close
            Else
302             MsgBox " Ο κωδικός είδους " + rkathg(0) + " δεν εχει κατηγορία "
304             rkathg.Close
306             checkIntegrity = False

                Exit Function
                  
            End If
        End If
 
308     If SQLDT("APALAGIFPA") > 0 Then   ' ean exei olikh apallagh fpa tote ok den me endiaferoyn ta eidh
         
        Else
310         rkathg.Open "select KODE,EID.APALLFPA from EGGTIM INNER JOIN EID ON  EID.KOD=EGGTIM.KODE WHERE (EID.APALLFPA<1) AND EID.FPA=5  AND ID_NUM=" + str(SQLDT("ID_NUM")), Gdb, adOpenDynamic, adLockOptimistic

312         If rkathg.EOF Then
314             rkathg.Close
            Else
316             MsgBox " Ο κωδικός είδους " + rkathg(0) + " δεν εχει ΑΠΑΛΛΑΓΗ ΦΠΑ "
318             rkathg.Close
320             checkIntegrity = False

                Exit Function
                  
            End If
        End If

322     checkIntegrity = True

        '<EhFooter>
        Exit Function

checkIntegrity_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.checkIntegrity " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        nERR = nERR + 1

        Resume Next

        '</EhFooter>

End Function

Private Sub SEND_EXPENS_Click()

        '<EhHeader>
        On Error GoTo SEND_EXPENS_Click_Err

        '</EhHeader>

        'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
        ' ΤΗΝ ΑΠΑΝΤΗΣΗ ΤΗΝ ΑΝΤΙΓΡΑΦΩ ΣΤΟ ΦΑΚΕΛΟ  FileCopy mfile, "C:\"+F_TXTFILES+"\apantSendInv.XML"
        'ΕΝΗΜΕΡΨΝΩ ΤΟ ΤΙΜ ΜΕ ΤΑ ΜΑΡΚΣ    UpdateTIMwithMARK

        'RequestDocs (GET) =========================
        '        ' ΣΤΕΛΝΩ ΣΤΟ SENDINVOICES
        '        ' "c:\"+F_TXTFILES+"\inv.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
        '        ' H APANTHSH EINAI STO  "c:\"+F_TXTFILES+"\apantSendInv.XML")
        'Dim XL As MSXML2.DOMDocument
        '  .Load("c:\"+F_TXTFILES+"\inv.xml") ' "--> e?e? ??? ap????e?se? t? xml p?? eft?a?a"

        'Dim TXML As New MSXML2.DOMDocument
   
        Dim TXML As String

100     TXML = "" ' New MSXML2.DOMDocument
        '= ""
102     CREATE_EXP (TXML)

        'Exit Sub

        Dim strXML As String, intFile As Integer

        'Open file
104     intFile = FreeFile

        'Exit Sub

106     strXML = ""

        Dim c As String

        '---------------------------------------------------- DIABAZEI ALLA ANSI ----------------------------
        'Open "c:\"+F_TXTFILES+"\inv.xml" For Input As intFile
        '
        ''Load XML into string strXML
        'While Not EOF(intFile)
        '    Line Input #intFile, c
        '    strXML = strXML + c
        'Wend
        'Close intFile

        'If Dir("c:\"+F_TXTFILES+"", vbDirectory) Then

        '
        '106 If Len(Dir("c:\" + F_TXTFILES + "", vbDirectory)) = 0 Then
        '108    MkDir "c:\" + F_TXTFILES + ""
        '    End If
        '110 If Len(Dir("c:\" + F_TXTFILES + "\sendinv", vbDirectory)) = 0 Then
        '112    MkDir "c:\" + F_TXTFILES + "\sendinv"
        '    End If
        '
        '114 If Len(Dir("c:\" + F_TXTFILES + "\inv.xml", vbNormal)) = 0 Then
        '116    MsgBox "δεν υπάρχει το αρχείο των τιμολογιων inv.xml"
        '       Exit Sub
        '    End If

        '--------------------------- ΔΙΑΒΑΖΕΙ ΑΡΧΕΙΟ ΣΕ ΜΟΡΦΗ UTF8 ------------------------------
        Dim objStream, strData

108     Set objStream = CreateObject("ADODB.Stream")

110     objStream.Charset = "utf-8"
112     objStream.Open
        '124 objStream.LoadFromFile ("c:\" + F_TXTFILES + "\EXP.xml") 'εχει δειγμα μετά το τέλος της SUB

114     objStream.LoadFromFile ("c:\" + F_TXTFILES + "\EXP.xml") 'εχει δειγμα μετά το τέλος της SUB

116     strData = objStream.ReadText()

118     objStream.Close

        'If fromFile = 1 Then

        'to exei parei hdh apo to create_exp
        ' ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΤΟ ΑΡΧΕΙΟ  ΤΟΤΕ :
        ' TXML = strData

        'End If

120     Set objStream = Nothing
        '-------------------------------------------------------------------------------------
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
122     url2 = UrlAADE + "SendExpensesClassification" '/RequestDocs"

        'initialize
124     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
126     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
128     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
130     XMLServer.Option(9) = 2048
132     XMLServer.Option(6) = True
    
134     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
136     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
138     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
140     XMLServer.send strData ' TXML 'strData  'ΣΤΕΛΝΩ ΤΟ STRING TXML STO MYDATA  'XMLServer.send TXML '  ΑΝ ΘΕΛΩ ΝΑ ΤΟ ΣΤΕΙΛΩ ΑΠΟ ΑΡΧΕΙΟ ΤΟΤΕ ΒΑΖΩ strData
142     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
144     v = XMLServer.responseText
        '--------------- ΑΠΟΣΤΟΛΗ INVOICES -----------------------------------------------------
 
146     Dim mfile As String: mfile = "c:\" + F_TXTFILES + "\sendinv\apantSendEXP" + Format(Now, "yyyyddMMHHmm") + ".xml"

148     Debug.Print v
150     Open mfile For Output As #8
152     Print #8, v
154     Close #8
    
156     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendEXP.XML"
    
        'UpdateTIMwithMARK v

        '<EhFooter>
        Exit Sub

SEND_EXPENS_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.SEND_EXPENS_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        
        Resume Next

        '</EhFooter>

End Sub

'----------------- exp.xml ---------------------------------------------------------------
'elemroot      <ExpensesClassificationsDoc xmlns="https://www.aade.gr/myDATA/expensesClassificaton/v1.0">
'  elemField  <expensesInvoiceClassification>
'         <invoiceMark>400001827580768</invoiceMark>

'elem0Field    <invoicesExpensesClassificationDetails>
' <lineNumber>1</lineNumber>

' elem10Field  <expensesClassificationDetailData>
'   <classificationType>E3_585_016</classificationType>
'   <classificationCategory>category2_3</classificationCategory>
'   <amount>10.00000</amount>
'   <id>1</id>
' </expensesClassificationDetailData>

' <expensesClassificationDetailData>
'   <classificationType>VAT_361</classificationType>
'   <classificationCategory>category2_3</classificationCategory>
'   <amount>10.00000</amount>
'   <id>2</id>
' </expensesClassificationDetailData>

'</invoicesExpensesClassificationDetails>
'</expensesInvoiceClassification>
'</ExpensesClassificationsDoc>

Sub CREATE_EXP(ByRef TXML As String)  ' MSXML2.DOMDocument)

        '-------------------- ΔΗΜΙΟΥΡΓΕΙ ΤΟ ΑΡΧΕΙΟ EXP.XML --------------------------------------
        '<EhHeader>
        On Error GoTo CREATE_EXP_Err

        Dim elemRoot   As MSXML2.IXMLDOMElement

        Dim elem0Root  As MSXML2.IXMLDOMElement

        Dim elem10Root As MSXML2.IXMLDOMElement

        Dim elem2Root  As MSXML2.IXMLDOMElement

        'αυτη ειναι η σωστη διατυπωση
        '<ExpensesClassificationsDoc
        ' xmlns="https://www.aade.gr/myDATA/expensesClassificaton/v1.0"
        ' xmlns:inv="http://www.aade.gr/myDATA/invoice/v1.0"
        ' xsi:schemaLocation="https://www.aade.gr/myDATA/expensesClassificaton/v1.0 schema.xsd"
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">

        'Set elemRoot = .createElement("InvoicesDoc")

        '166         With elemRoot

100     Dim ctyp       As String: ctyp = FIND2_EXODTYPOS(Left$(Adodc2.Recordset("ATIM"), 1))

102     Set docStock = New MSXML2.DOMDocument

104     With docStock
106         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
            '                .setAttribute "xmlns", "https://www.aade.gr/myDATA/expensesClassificaton/v1.0"
            '                .setAttribute "xmlns:inv", "http://www.aade.gr/myDATA/invoice/v1.0"
            '                .setAttribute "xsi:schemaLocation", "https://www.aade.gr/myDATA/expensesClassificaton/v1.0 schema.xsd"
            '                .setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
            '
            '  Set elemRoot = .createElement("InvoicesDoc")

            ' With elemRoot '/////////////////////////////////////////////////////////////////////////////////////////

            ' End With
      
108         Set elemRoot = .createElement("ExpensesClassificationsDoc")

110         With elemRoot
                            
                ' elemRoot.setAttribute "xmlns", UrlAADE + "expensesClassificaton/v1.0"

112             .setAttribute "xmlns", "https://www.aade.gr/myDATA/expensesClassificaton/v1.0"
114             .setAttribute "xmlns:inv", "http://www.aade.gr/myDATA/invoice/v1.0"
116             .setAttribute "xsi:schemaLocation", "https://www.aade.gr/myDATA/expensesClassificaton/v1.0 schema.xsd"
118             .setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
        
120             Set elemField = docStock.createElement("expensesInvoiceClassification") ' δημιουργω εσοχη

122             With elemField
        
124                 Set elem2Field = docStock.createElement("invoiceMark"): elem2Field.Text = Adodc2.Recordset("MARK"): elemField.appendChild elem2Field
        
126                 Set elem0Field = docStock.createElement("invoicesExpensesClassificationDetails") ' δημιουργω εσοχη

128                 With elem0Field
             
130                     Set elem2Field = docStock.createElement("lineNumber"): elem2Field.Text = "1": elem0Field.appendChild elem2Field
        
                        '------------ expensesClassificationDetailData ------------------------
132                     Set elem10Field = docStock.createElement("expensesClassificationDetailData") ' δημιουργω εσοχη
134                     Set elem2Field = docStock.createElement("classificationType"): elem2Field.Text = Split(ctyp, ";")(0): elem10Field.appendChild elem2Field
136                     Set elem2Field = docStock.createElement("classificationCategory"): elem2Field.Text = Split(ctyp, ";")(1): elem10Field.appendChild elem2Field
138                     Set elem2Field = docStock.createElement("amount"): elem2Field.Text = Replace(str(Adodc2.Recordset("TOTALNETVALUE")), ",", "."): elem10Field.appendChild elem2Field
140                     Set elem2Field = docStock.createElement("id"): elem2Field.Text = "1": elem10Field.appendChild elem2Field
142                     elem0Field.appendChild elem10Field
    
                        '------------ expensesClassificationDetailData ------------------------
144                     Set elem10Field = docStock.createElement("expensesClassificationDetailData") ' δημιουργω εσοχη
146                     Set elem2Field = docStock.createElement("classificationType"): elem2Field.Text = "VAT_361": elem10Field.appendChild elem2Field
148                     Set elem2Field = docStock.createElement("classificationCategory"): elem2Field.Text = Split(ctyp, ";")(1): elem10Field.appendChild elem2Field
150                     Set elem2Field = docStock.createElement("amount"): elem2Field.Text = Replace(str(Adodc2.Recordset("TOTALVATAMOUNT")), ",", "."): elem10Field.appendChild elem2Field
152                     Set elem2Field = docStock.createElement("id"): elem2Field.Text = "2": elem10Field.appendChild elem2Field
154                     elem0Field.appendChild elem10Field
    
156                     .appendChild elem10Field
                    End With
            
158                 elemField.appendChild elem0Field
                End With
       
160             elemRoot.appendChild elemField
            End With
     
162         .appendChild elemRoot
        End With
  
164     FormatXmlDocument docStock ' βαζει κενα να ειναι ευκολο στο διαβασμα
  
        '<<<<<<< HEAD
166     docStock.save "C:\" + F_TXTFILES + "\EXP.xml"

168     TXML = docStock.XML
    
        Exit Sub
    
        Dim R As New ADODB.Recordset

170     R.Open "SELECT * FROM APESTALMENA WHERE CLASSTYPE11 IS NULL", Gdb, adOpenDynamic, adLockOptimistic

172     Do While Not R.EOF

        Loop

        '<EhFooter>
        Exit Sub

CREATE_EXP_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.CREATE_EXP " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub update_sql_from_dbf(arxeio As String)

        '<EhHeader>
        On Error GoTo update_sql_from_dbf_Err

        '</EhHeader>
        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim conDBF As String

        Dim conSQL As String

        Dim k      As Long

        Dim Fname  As String

100     g2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
        '2004 G2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
102     conSQL = gConnect    ' "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP;Data Source=LAGAKIS"
104     conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO"
        '2004 conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO\04"

106     sql.Open conSQL
        'SET DBF=DATA.Open conDBF
        
        '  On Error GoTo 0

108     Data1.DatabaseName = "C:\LAGEURO" ' Text1.Text

110     Data1.RecordSource = "SELECT * FROM PEL WHERE EIDOS='e' ORDER BY KOD DESC"
112     Data1.Refresh

114     Set rDBF = Data1.Recordset

116     Label1.Caption = arxeio + " " + Format(rDBF.RecordCount, "######")
        ' sql.Execute "DELETE FROM " + arxeio

118     rSQL.Open "SELECT *FROM " + arxeio, sql, adOpenDynamic, adLockOptimistic
        ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

120     rDBF.MoveFirst

        Dim Z

122     Z = 0

        Dim r2SQL As New ADODB.Recordset

        On Error GoTo WRITEERROR    ' Resume Next

124     Do While Not rDBF.EOF

126         r2SQL.Open "select count(*) from PEL WHERE KOD='" + rDBF("KOD") + "'AND EIDOS='e'", Gdb, adOpenDynamic, adLockOptimistic

128         If r2SQL(0) = 0 Then

130             rSQL.AddNew

132             For k = 0 To rDBF.FIELDS.Count - 1
134                 Fname = rDBF.FIELDS(k).Name    ' p.x. FNAME=epo    R(0).NAME

136                 If IsNull(rDBF(k)) Then
                    Else
138                     rSQL(Fname) = rDBF(k)    ' rsqk("epo")=r(0)
                    End If

                    ' ΑΥΤΑ ΠΟΥ ΕΙΝΑΙ ΝULL ΤΑ ΓΕΜΙΖΕΙ
140                 If IsNull(rDBF(k)) Then
142                     If rDBF(k).Type = dbText Then    '10=STRING
144                         rSQL(Fname) = " "
146                     ElseIf rDBF(k).Type = dbDate Then
148                         rSQL(Fname) = CDate("01/01/1990")
150                     ElseIf rDBF(k).Type = dbBoolean Then
152                         rSQL(Fname) = False
154                     ElseIf rDBF(k).Type = dbDouble Then
156                         rSQL(Fname) = 0
                        Else
158                         conSQL = ""
                        End If
                    End If

                Next

160             rSQL.Update
162             Z = Z + 1
          
            End If

164         r2SQL.Close
          
166         If Z Mod 10 = 0 Then
168             Me.Caption = Z
            End If

170         rDBF.MoveNext

172         DoEvents

        Loop

174     rSQL.Close
176     rDBF.Close

        Exit Sub

WRITEERROR:
        'List2.Visible = True

        ' List2.AddItem Err.Description

178     Resume Next

        '<EhFooter>
        Exit Sub

update_sql_from_dbf_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu8.update_sql_from_dbf " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu8.update_sql_from_dbf " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Πώς διαβιβάζεται η εγγυοδοσία (κενές φιάλες) στη ψηφιακή πλατφόρμα myData;

'Επίλυση
'Το συνολικό ποσό που αφορά εγγυοδοσία, έρχεται ως διακριτή γραμμή στο παραστατικό που διαβιβάζεται και
'  χαρακτηρίζεται ως 1.95 «Λοιπά Πληροφοριακά Στοιχεία Εσόδων» από τον Εκδότη και 2.95 «Λοιπά Πληροφοριακά Στοιχεία Εξόδων» από το Λήπτη.
'   Και στις δύο περιπτώσεις (Εκδότη και Λήπτη) οι αναγραφόμενες αξίες δεν σχετίζονται με κωδικό Ε3 και δεν επηρεάζουν το λογιστικό και φορολογικό αποτέλεσμα.
'<invoiceDetails>
'      <lineNumber> 1</lineNumber>
'      <netValue>10.</netValue>
'      <vatCategory>1</vatCategory>
'      <vatAmount>2.4</vatAmount>
'      <incomeClassification>
'        <n1:classificationType>E3_561_001</n1:classificationType>
'        <n1:classificationCategory>category1_1</n1:classificationCategory>
'        <n1:amount>10.</n1:amount>
'      </incomeClassification>
'    </invoiceDetails>
'
'   <invoiceDetails>
'      <lineNumber> 2</lineNumber>
'      <netValue>4.</netValue>
'      <vatCategory>7</vatCategory>
'      <vatAmount>0</vatAmount>
'     <vatExemptionCategory>6</vatExemptionCategory>
'      <incomeClassification>
'
'        <n1:classificationCategory>category1_95</n1:classificationCategory>
'        <n1:amount>4.</n1:amount>
'      </incomeClassification>
'    </invoiceDetails>
'
'
'    <invoiceSummary>
'      <totalNetValue>14.</totalNetValue>
'      <totalVatAmount>2.4</totalVatAmount>
'      <totalWithheldAmount>0.00</totalWithheldAmount>
'      <totalFeesAmount>0.00</totalFeesAmount>
'      <totalStampDutyAmount>0.00</totalStampDutyAmount>
'      <totalOtherTaxesAmount>0.00</totalOtherTaxesAmount>
'      <totalDeductionsAmount>0.00</totalDeductionsAmount>
'      <totalGrossValue>16.4</totalGrossValue>
'      <incomeClassification>
'        <n1:classificationType>E3_561_001</n1:classificationType>
'        <n1:classificationCategory>category1_1</n1:classificationCategory>
'        <n1:amount>10.</n1:amount>
'      </incomeClassification>
'      <incomeClassification>
'
'        <n1:classificationCategory>category1_95</n1:classificationCategory>
'        <n1:amount>4.</n1:amount>
'      </incomeClassification>
'
'
''
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

Private Sub TDBGrid2_HeadClick(ByVal ColIndex As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid1_HeadClick_Err

        '</EhHeader>

        '--------------------------------------------------------------------------------------=
        Dim sumes(100) As String

100     For k = 0 To Adodc2.Recordset.FIELDS.Count - 1
102         sumes(k) = TDBGrid2.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
        Next

104     If Adodc2.Recordset.sort = "[" & TDBGrid2.columns(ColIndex).DataField & "] asc" Then   ' strSort
106         Adodc2.Recordset.sort = "[" & TDBGrid2.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
108         Adodc2.Recordset.sort = "[" & TDBGrid2.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

110     For k = 0 To Adodc2.Recordset.FIELDS.Count - 1
112         TDBGrid2.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
        Next

        '<EhFooter>
        Exit Sub

TDBGrid1_HeadClick_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.TDBGrid1_HeadClick " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Sub send_to_Access()

        Dim stm     As ADODB.Stream

        Dim strFile As String

100     strFile = "c:\mercvb\fotoP.bmp"
  
        Dim DS As New ADODB.Connection

        'ds.Open "provider=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=c:\mercvb\reports\reports.mdb;PERSIST SECURITY=FALSE"
102     DS.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mercvb\reports\REPORTS.MDB;Persist Security Info=False"
  
        Dim RR As New ADODB.Recordset

104     RR.Open "SELECT * FROM MEM", DS, adOpenDynamic, adLockOptimistic
  
        ' Dim imgByte()   As Byte
    
        ' imgByte = GetBytes(Picture1.Picture, "C:\mercvb\FOTOP.bmp")
 
        ' Dim DB As DAO.Database
        ' Dim rmem As DAO.Recordset
        '
        '
        '  Set DB = OpenDatabase("c:\mercvb\reports\reports.mdb")
        '  Set rmem = DB.OpenRecordset("TIM")
        '
106     Set stm = New ADODB.Stream
108     stm.Type = adTypeBinary
110     stm.Open
112     stm.LoadFromFile strFile
  
        ' RR("MEMO").Value = "1345SD"
114     RR("foto").Value = stm.Read  'imgByte  'stm.Read
116     RR.Update
  
118     RR.Close

        ' R.Recordset("memo").Value = "asdasdas"
        'R.Recordset("foto") = stm.Read
        'R.Recordset.Update

        'rmem.edit
        'rmem("foto") = stm.Read
        'rmem("MEMO") = "AAAAA"   ' stm.Read
        'rmem.Update

120     stm.Close
122     Set stm = Nothing

End Sub

Public Function ToJason(ByVal noask As Integer, _
                        ByVal UPLOAD As Integer) As Integer ' 1=no messbog
  
        Dim m_is_benzinadiko As Integer

        m_is_benzinadiko = Val(FINDPARAMETROI(1, "PAR1", "f_is_benzinadiko", "0", "Γ.34.EINAI BENZINAΔIKO=1 ΟΧΙ=0"))  'αγορ=1 πωλ=2 παντου=3 ΟΧΙ=0"))

        Dim POSTransactionId As String

        'αν θελω διαχωρισμο
        '    Sj = Sj + Chr(13) +Chr(13) +

        '<EhHeader>
        On Error GoTo Tojason_Err

        '</EhHeader>

        Dim R8     As New ADODB.Recordset, CR8 As String, CR8B As String
       
        Dim SUMNET As Double

        Dim SJ     As String
    
100     ToJason = 1

        Dim ctypos As String, cTyposExod As String
   
        Dim sql    As String

102     sql = FIND_QUERY()
    
104     fLong = Dt2Lng(Now)

106     Gdb.Execute "UPDATE TIM SET ENTITY=0,ENTLINEN=" + str(fLong) + fSynt
    
        'ΕΑΝ ΕΙΝΑΙ ΠΡΟΒΛΗΜΑΤΙΚΟ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ ΑΠΟΘΗΚΕΥΕΙ ΤΟ ΠΡΟΒΛΗΜΑ ΣΤΟ ENTITYUID
        'KAI BAZEI ENTITYΜΑΡΚ="ERROR-ΑΝΕΥ ΑΠΟΣΤΟΛΗΣ"
        Dim sqlDtemp As New ADODB.Recordset

108     sqlDtemp.Open sql, Gdb, adOpenDynamic, adLockOptimistic

        Dim mProbl    As String, mTypPar As String

        Dim SYN_PROBL As String
        
110     Do While Not sqlDtemp.EOF
112         mProbl = ""

            ' AN EINAI AKYRVMENO KAI EXEI MARK NA STEILEI TO AKYRVTIKO
114         If Val(Left(CNull(sqlDtemp!entitymark) + " ", 1)) > 1 Or Left(CNull(sqlDtemp!entitymark) + " ", 1) = "Δ" Then
116             If sqlDtemp!AKYROMENO = 1 Then
118                 If Len(where.Caption) > 0 Then  ' ONLINE ΔΕΝ ΕΧΡΕΙΑΖΕΤΑΙ GRID
                        ' TIPOTA DEN KANV
                    Else

                        Dim CCCL As String

120                     CCCL = CANCEL_INVOICE(IIf(Left(CNull(sqlDtemp!entitymark) + " ", 1) = "Δ", mID(sqlDtemp!entitymark, 4, 15), sqlDtemp!entitymark), sqlDtemp!ATIM)    ' CANCEL_INVOICE(sqlDtemp!ENTITYMARK, sqlDtemp!ATIM)

122                     If Len(CNull(CCCL)) > 1 Then
124                         Gdb.Execute "UPDATE TIM SET INCMARK='" + CCCL + "' WHERE ID_NUM=" + str(sqlDtemp!id_num)
                        End If
                    End If
                End If

            Else

126             If checkIntegrity(i, sqlDtemp, mProbl) = False Then
128                 If mProbl = "" Then
130                     mProbl = "ΠΡΟΒΛΗΜΑ ΑΚΕΡΑΙΟΤΗΤΑΣ " + sqlDtemp("ATIM")
                    End If
                End If

132             ctypos = FINDTYPOS(mID(sqlDtemp("ATIM"), 1, 1))
                  
134             mTypPar = mID(Split(ctypos, ";")(0), 1, 3)
                    
136             If Left(mTypPar, 2) = "11" Then
                    'lianikh den xreiazetai pelaths
                Else

138                 If sqlDtemp("COUNTRY") <> "GR" Then
                    Else

140                     If Len(Trim(sqlDtemp("AFM"))) < 9 Then
142                         mProbl = "124J.ΔEN EXEI ΑΦΜ "
                        Else

                            'αν δεν ειναι ενδοκοινοτικό ή τρίτων χωρών?
144                         If InStr("1.2  1.3 2.2 2.3  2.4", mTypPar) = 0 Then
146                             If check_afm(Trim(sqlDtemp("AFM"))) = 0 Then
148                                 mProbl = "130J. EXEI ΛΑΘΟΣ ΑΦΜ "
                                End If
                            End If
                            
                        End If
                    End If
                End If
                    
150             If mTypPar = "5.1" Then   ' -------- συσχετιζομενο
152                 If IsNull(sqlDtemp("SXETMARK")) Then
154                     mProbl = "ΔΕΝ EXEI ΣΧΕΤΙΚΟ ΠΑΡΑΣΤΑΤΙΚΟ "
                    Else

156                     If Len(sqlDtemp("SXETMARK")) < 10 Then
158                         mProbl = "ΔΕΝ EXEI ΣΧΕΤΙΚΟ ΠΑΡΑΣΤΑΤΙΚΟ "
                        End If
                    End If
                End If
                    
160             If Len(Trim(mProbl)) > 2 Then
162                 Gdb.Execute "UPDATE TIM SET ENTITYMARK='ERROR-ΠΡΙΝ ΑΠΟΣΤ', ENTITYUID='" + mProbl + "' where ID_NUM=" + str(sqlDtemp("ID_NUM"))
164                 SYN_PROBL = SYN_PROBL + mProbl + Chr(13)
                End If
                
166             Dim ubl3 As String: ubl3 = GGET_CVALUE("SELECT left(ISNULL(UBL,'000'),3) AS UBL2 FROM PARASTAT WHERE EIDOS='" & Left(sqlDtemp("ATIM"), 1) & "'")

                'e?? paroxo mono sthn xondriki kai to parastattiko anti na exei 380 exei "mydata"
168             If ubl3 = "000" Then 'mydata
170                 SYN_PROBL = "ΕΙΝΑΙ ΓΙΑ MYDATA.ΠΑΤΕ ΜΕ ΔΙΟΡΘΩΣΗ ΠΑΡΑΣΤΑΤΙΚΟΥ"
                Else
                    'OK
                End If
                          
            End If
                      
172         sqlDtemp.MoveNext
        Loop
    
        ' εδω καθυστερει στην online ενημερωση (ανα τιμολογιο ενημερωση )
174     If Len(where.Caption) > 1 Then  ' an erxetai apo online mhn kaneis refresh
   
        Else
176         Adodc1.Refresh   ' εδω καθυστερει στην online ενημερωση (ανα τιμολογιο ενημερωση )
    
        End If
    
178     If Len(SYN_PROBL) > 0 Then
180         If noask = 0 Then
182             MsgBox SYN_PROBL
184             MsgBox "ΑΔΥΝΑΤΗ Η ΣΥΝΕΧΙΣΗ"
            End If

            Exit Function

        End If
    
        Dim SQLDT As New ADODB.Recordset

        '  SQL = "SELECT  top 20  AJ1 ,AJ2  from TIM  order by HME"
    
        '==================================================================================
186     SQLDT.Open sql, Gdb, adOpenDynamic, adLockOptimistic

        '==================================================================================
188     If SQLDT.EOF Then
190         If noask = 0 Then
192             MsgBox ("ΔΕΝ ΒΡΕΘΗΚΑΝ ΕΓΓΡΑΦΕΣ")
            End If

194         ToJason = 0

            Exit Function

        End If

        Dim isDiakin  As Integer

        Dim IsAytotim As Integer

        Dim varStock  As Variant

        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim invoice   As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement
    
        '<InvoicesDoc   xmlns="http://www.aade.gr/myDATA/invoice/v1.0"
        ' xsi:schemaLocation="http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd"
        ' xmlns:N1="https://www.aade.gr/myDATA/incomeClassificaton/v1.0"
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
196     Set docStock = New MSXML2.DOMDocument

198     With docStock
200        ' .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
202         Set elemRoot = .createElement("InvoicesDoc")

204         With elemRoot '/////////////////////////////////////////////////////////////////////////////////////////
206             .setAttribute "xmlns", "http://www.aade.gr/myDATA/invoice/v1.0"
208             .setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd"
210             .setAttribute "xmlns:n1", "https://www.aade.gr/myDATA/incomeClassificaton/v1.0"
212             .setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
214             .setAttribute "xmlns:ecls", "https://www.aade.gr/myDATA/expensesClassificaton/v1.0"
            
                '
                '    '        <?xml version="1.0" encoding="UTF-8"?>
                ''<InvoicesDoc xmlns="http://www.aade.gr/myDATA/invoice/v1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                'xsi:            schemaLocation = "http://www.aade.gr/myDATA/invoice/v1.0/InvoicesDoc-v0.6.xsd"
                'xmlns:            icls = "https://www.aade.gr/myDATA/incomeClassificaton/v1.0"
                'xmlns:            ecls = "https://www.aade.gr/myDATA/expensesClassificaton/v1.0"
            
                'Dim I As Long:
216             i = -1

218             Dim nbb As Long: For nbb = 1 To 5000: F_ID_NUMS(nbb) = 0: Next ' ΑΠΟΘΗΚΕΥΕΙ ΤΑ ID_NUM ΓΙΑ ΝΑ ΤΑ ΘΥΜΑΤΑΙ ΟΤΑΝ ΕΡΘΕΙ Η ΑΠΑΝΤΗΣΗ ΑΠΟ ΑΑΔΕ

220             Do While Not SQLDT.EOF
            
                    'Dim isDiakin As Integer
222                 isDiakin = GGET_NVALUE("SELECT ISNULL(ISDIAKIN,0) FROM PARASTAT WHERE EIDOS='" + Left(SQLDT("ATIM"), 1) + "'")
            
224                 i = i + 1
               
226                 f_name_xml = SQLDT("atim")

                    ' ΑΝ ΤΟ ΒΕΝΖΙΝΑΔΙΚΟ ΚΟΒΕΙ ΤΙΜ ΓΙΑ ΛΑΔΙΑ ΝΑ ΓΥΡΙΖΕΙ ΒΕΝΖΙΝΑΔΙΚΟ=0

                    If m_is_benzinadiko = 1 Then

                        Dim mm_code As String
                                   
                        mm_code = GGET_CVALUE("select KODE FROM EGGTIM WHERE ID_NUM=" + str(SQLDT("ID_NUM")))

                        If InStr("10 11 12 13 14 15 20 21 30 31 32 33 34 35 36 37 38 40 41 42 43 44 50 60 61 70 71 72", Trim(mm_code)) > 0 Then      'μονο στου βενζιναδικου να βγαζει fuelcode
                            m_is_benzinadiko = 1
                        Else
                            m_is_benzinadiko = 0
                        End If
                    End If
                    
                    '  If checkIntegrity(i, sqlDt) = False Then
                    '     MsgBox (" ΠΡΟΒΛΗΜΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + sqlDt(0)("ATIM").toString)
                    '  End If
               
                    '  If checkIntegrity(I) = False Then

                    '    MsgBox (" ΠΡΟΒΛΗΜΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + sqldt(0)("ATIM").ToString)
                    '  End If

                    Dim EGGTIM As New ADODB.Recordset

228                 Me.Caption = "ΠΑΡΑΣΤΑΤΙΚΑ " + str(i + 1)
                
230                 List1.AddItem "Δημιουργήθηκαν  " + str(i + 1) + " παραστατικά "
                    'EGGTIM.Open "SELECT KODE,POSO,TIMM,EKPT,FPA,ISNULL(KAU_AJIA,0) AS KAU_AJIA,ISNULL(MIK_AJIA,0) AS MIK_AJIA FROM EGGTIM WHERE TIMM<>0 AND POSO<>0 AND ID_NUM=" + str(sqlDt("ID_NUM")), Gdb, adOpenDynamic, adLockOptimistic
                
232                 Dim C56       As String: C56 = ""

                    Dim exei_axia As String

                    'Δ.Α.  ή ΤΕΛΟς ΑΝΘΕΚΤΙΚΟΤΗΤΑΣ
233                 If isDiakin = 2 Or Split(FINDTYPOS(mID(SQLDT("ATIM"), 1, 1)), ";")(0) = "8.2" Then
                        exei_axia = ""
                    Else
                        exei_axia = " and TIMM<>0 AND ISNULL(KAU_AJIA,0)<>0 "
                    End If
                  
                    'ΜΟΝΟ ΓΙΑ ΠΑΡΟΧΟ  (ΔΕΝ ΚΑΝΕΙ ΔΙΑΧΩΡΙΣΜΟ ΑΠΟ ΕΙΔΗ/ΠΑΡΑΣΤΑΤΙΚΟ !!!!!!!!!!!!!!!
234                 C56 = "SELECT  KODE,ONOMA,POSO,CONVERT(INTEGER,ISNULL(KATHGORIA,'1') ) AS KATHGORIA, CONVERT(INTEGER,G.FPA) AS FPA,ROUND((ISNULL(KAU_AJIA,0)),2) AS KAU_AJIA,ROUND((ISNULL(MIK_AJIA,0)),2) AS MIK_AJIA, ISNULL(D.APALLFPA,0) AS APAL,ISNULL(D.CPV,'') AS CPV,ISNULL(G.ONOMA,'') AS ONOMA,ISNULL(G.MONA,'???') AS MONA "
236                 C56 = C56 & " FROM EGGTIM G  INNER JOIN EID D  ON G.KODE=D.KOD "
238                 C56 = C56 & " WHERE POSO<>0  " + exei_axia + "  and ID_NUM=" & str(SQLDT("ID_NUM"))
240                 EGGTIM.Open C56, Gdb, adOpenDynamic, adLockOptimistic

                    ' ΠΡΟΣΟΧΗ ΟΧΙ ΓΙΑ ΠΑΡΟΧΟ
242                 If fMydataFromEID1 = 1 Then  'ΕΝΗΜΕΡΩΣΗ ΑΠΟ ΚΑΤΗΓΟΡΙΑ ΕΙΔΩΝ
                     
                        'C56 = "SELECT  CONVERT(INTEGER,KATHGORIA) AS KATHGORIA, CONVERT(INTEGER,G.FPA) AS FPA,ROUND(SUM(ISNULL(KAU_AJIA,0)),2) AS KAU_AJIA,ROUND(SUM(ISNULL(MIK_AJIA,0)),2) AS MIK_AJIA, ISNULL(D.APALLFPA,0) AS APAL "
                        'C56 = C56 + " FROM EGGTIM G  INNER JOIN EID D  ON G.KODE=D.KOD "
                        'C56 = C56 + " WHERE TIMM<>0 AND POSO<>0 AND ISNULL(KAU_AJIA,0)<>0 and ID_NUM=" + str(sqlDt("ID_NUM"))
                        'C56 = C56 + " GROUP BY CONVERT(INTEGER,G.FPA),CONVERT(INTEGER,KATHGORIA), ISNULL(D.APALLFPA,0) "   ' G.FPA,KATHGORIA,APALLFPA "
                        ' EGGTIM.Open C56, Gdb, adOpenDynamic, adLockOptimistic
                    Else   ' ΜΟΝΟ ΑΠΟ ΠΑΡΑΣΤΑΤΙΚΟ
                        'C56 = "SELECT '1' AS KATHGORIA,CONVERT(INTEGER,G.FPA) AS FPA,ROUND(SUM(ISNULL(KAU_AJIA,0)),2) AS KAU_AJIA, "
                        'C56 = C56 + " ROUND(SUM(ISNULL(MIK_AJIA,0)),2) AS MIK_AJIA, ISNULL(D.APALLFPA,0) AS APAL "
                        'C56 = C56 + " FROM EGGTIM  G  INNER JOIN EID D  ON G.KODE=D.KOD WHERE TIMM<>0 AND POSO<>0  AND ISNULL(KAU_AJIA,0)<>0 AND ID_NUM=" + str(sqlDt("ID_NUM")) + " GROUP BY CONVERT(INTEGER,G.FPA), ISNULL(D.APALLFPA,0)" ' G.FPA,APALLFPA "
                        ' EGGTIM.Open C56, Gdb, adOpenDynamic, adLockOptimistic
                    
                    End If
                
                    ' ελεγχω αν ησ ουμα των κατηγοριών είναι ίδια με την σουμα του τιμολογιου
                    Dim suma_kathg As Double

244                 suma_kathg = 0

                    Dim SUM_MIK_AJIA As Single

                    Dim SUM_MIK2     As Double

246                 SUM_MIK2 = 0

248                 Do While Not EGGTIM.EOF
250                     suma_kathg = suma_kathg + EGGTIM!kau_ajia
252                     SUM_MIK2 = SUM_MIK2 + EGGTIM!MIK_AJIA
254                     EGGTIM.MoveNext
                    Loop

256                 SUM_MIK2 = GGET_NVALUE("SELECT SUM(MIK_AJIA) FROM EGGTIM WHERE ID_NUM=" + str(SQLDT("ID_NUM")))

258                 EGGTIM.MoveFirst
                
                    'Dim DUM As New DataTable
260                 F_ID_NUMS(i + 1) = SQLDT("ID_NUM") ' ' ΑΠΟΘΗΚΕΥΕΙ ΤΑ ID_NUM ΓΙΑ ΝΑ ΤΑ ΘΥΜΑΤΑΙ ΟΤΑΝ ΕΡΘΕΙ Η ΑΠΑΝΤΗΣΗ ΑΠΟ ΑΑΔΕ
262                 Gdb.Execute "UPDATE TIM SET ENTITY=" + str(i + 1) + " WHERE ID_NUM=" + str(SQLDT("ID_NUM"))

264                 SUMNET = SQLDT("aj1") + SQLDT("aj2") + SQLDT("aj3") + SQLDT("aj4") + SQLDT("aj5") + SQLDT("aj6") + SQLDT("aj7")
266                 SumFpa = SQLDT("fpa1") + SQLDT("fpa2") + SQLDT("fpa3") + SQLDT("fpa4") + SQLDT("fpa6") + SQLDT("fpa7")

                    ' ΣΥΜΦΩΝΟΥΝ ΟΙ ΚΑΘΑΡΕΣ ΑΞΙΕΣ ?
268                 If Abs(SUMNET - suma_kathg) > 0.005 Then
270                     If noask = 0 Then
272                         MsgBox ("δεν συμφωνουν οι ΚΑΘΑΡΕΣ αξιες στις κατηγορίες με το παρ/κό " + SQLDT("ATIM")) + " " + str(SUMNET) + "<>" + str(suma_kathg)
274                         MsgBox ("Ακυρώθηκαν όλες οι αποστολές ")
                        End If

                        Exit Function

                    End If
                     
                    Dim DIFF_FPA As Single

                    Dim OK_DIFF  As Integer

276                 OK_DIFF = 1

                    ' ΣΥΜΦΩΝΟΥΝ ΟΙ ΜΙΚΤΕΣ ΑΞΙΕΣ ?
278                 If Abs(SUMNET + SumFpa - SUM_MIK2) > 0.005 Then
                     
280                     If Abs(SUMNET + SumFpa - SUM_MIK2) > 0.05 Then
282                         If noask = 0 Then
284                             MsgBox ("δεν συμφωνουν οι ΜΙΚΤΈΣ αξιες στις κατηγορίες με το παρ/κό " + SQLDT("ATIM")) + " " + str(SUMNET + SumFpa) + "<>" + str(SUM_MIK2)
286                             MsgBox ("Ακυρώθηκαν όλες οι αποστολές ")
                            End If

                            Exit Function

                        Else
288                         DIFF_FPA = SUMNET + SumFpa - SUM_MIK2
290                         OK_DIFF = 0
                        End If
                       
                    End If

                    '1.1;E3_561_001;category1_1;0
292                 ctypos = FINDTYPOS(mID(SQLDT("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
                
294                 cTyposExod = FINDEXODTYPOS(mID(SQLDT("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
                
296                 IsAytotim = Val(Trim(Split(ctypos, ";")(3)))
                
298                 If Len(Trim(Split(ctypos, ";")(2))) = 0 And Len(Trim(Split(cTyposExod, ";")(1))) = 0 Then  ' MPOREI TO E3 NA EINAI KENO P.X. 1_95

                        ' If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
                        'writer.Close()
300                     If noask = 0 Then
302                         MsgBox ("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " + SQLDT("ATIM"))
304                         MsgBox ("Ακυρώθηκαν όλες οι αποστολές ")
                        End If

                        Exit Function

                    End If
                
306                 gf_Bearer = getToken() ' REQPAROCHOS()

308                 If Len(gf_Bearer) < 2 Then
310                     MsgBox ("ΑΔΥΝΑΤΗ Η ΣΥΝΔΕΣΗ ΜΕ ΤΟΝ ΠΑΡΟΧΟ")
                        'Exit Function
                    End If
                
                    Dim is_B2G As Integer, ubl As String
                 
312                 is_B2G = GGET_NVALUE("SELECT ISNULL(B2G,0) AS DIMOSIO FROM PARASTAT WHERE EIDOS='" & Left(SQLDT("ATIM"), 1) & "'")
314                 ubl = GGET_CVALUE("SELECT ISNULL(UBL,'') AS UBL2 FROM PARASTAT WHERE EIDOS='" & Left(SQLDT("ATIM"), 1) & "'")
            
316                 SJ = " {""invoice"":{"   ' arxh invoice  ----------------------------------------------------------------
            
                    ' For I = 0 To UBound(varStock)
318                 Set invoice = docStock.createElement("invoice")

320                 With invoice  '----------------------------------
322                     List1.AddItem SQLDT("ATIM") + " " + Format(SQLDT("hme"), "dd/MM/yyyy")
324                     ' ' Set elemField = docStock.createElement("uid"):'elemField.Text = str(i + 1): INVOICE.appendChild elemField
                      
                        Dim IssuerAFM, IssuerCountry, iss
                      
                        Dim m_ypok As String
                      
326                     If SQLDT("EIDOS") = "r" And Split(ctypos, ";")(0) <> "3.1" And IsAytotim = 0 Then '  IsAytotim = 0  gia na petaei thn αυτοτιμολόγηση στο else
           
                            '------------ ISSUER ------------------------
328                         ' Set elemField = docStock.createElement("issuer") ' δημιουργω εσοχη
330                         ' Set elem2Field = docStock.createElement("vatNumber"):'elem2Field.Text = Trim(SQLDT("AFM")): ' ' elemField.appendChild elem2Field
332                         ' Set elem2Field = docStock.createElement("country"):'elem2Field.Text = SQLDT("COUNTRY"): ' ' elemField.appendChild elem2Field
                                     
334                         ' Set elem2Field = docStock.createElement("branch"):'elem2Field.Text = "0": ' ' elemField.appendChild elem2Field

336                         If SQLDT("COUNTRY") <> "GR" Then
                        
338                             ' Set elem2Field = docStock.createElement("name"):'elem2Field.Text = SQLDT("epo"): ' ' elemField.appendChild elem2Field
                            End If
                        
340                         ' Set elem2Field = docStock.createElement("address") ' δημιουργω εσοχ' δημιουργω εσοχ  sqlDT(i)("XRVMA")= TK
342                         ' Set elem3Field = docStock.createElement("postalCode"):'elem3Field.Text = CNull(SQLDT("TK")): ' elem2Field.appendChild elem3Field
344                         ' Set elem3Field = docStock.createElement("city"):'elem3Field.Text = CNull(SQLDT("POL")): ' elem2Field.appendChild elem3Field
346                         ' ' elemField.appendChild elem2Field
                          
348                        ' .appendChild elemField
                                
                            '------------ COUNTER PART ------------------------
350                         If mID(Split(ctypos, ";")(0), 1, 2) = "11" Then
                                'lianikh den xreiazetai pelaths
                            Else
                                
352                             ' Set elemField = docStock.createElement("counterpart") ' δημιουργω εσοχ
354                             ' Set elem2Field = docStock.createElement("vatNumber"):'elem2Field.Text = afmCompany: ' ' elemField.appendChild elem2Field
356                             ' Set elem2Field = docStock.createElement("country"):'elem2Field.Text = "GR": ' ' elemField.appendChild elem2Field
                         
358                             m_ypok = f_YPOK

360                             If InStr(f_par1ypok, Left(SQLDT("ATIM"), 1)) > 0 Then
362                                 m_ypok = "1"
                                End If

364                             If InStr(f_par2ypok, Left(SQLDT("ATIM"), 1)) > 0 Then
366                                 m_ypok = "2"
                                End If

                                'f_YPOK
                                ' ' Set elem2Field = docStock.createElement("branch"):'elem2Field.Text = "0": ' ' elemField.appendChild elem2Field
                                ' Set elem2Field = docStock.createElement("branch"):'elem2Field.Text = m_ypok:
368                             ' ' elemField.appendChild elem2Field
  
370                            ' .appendChild elemField
                 
                            End If
                      
                        Else
                            '------------ ISSUER ------------------------
372                         ' Set elemField = docStock.createElement("issuer") ' δημιουργω εσοχη
374                         ' Set elem2Field = docStock.createElement("vatNumber"):'elem2Field.Text = afmCompany: ' ' elemField.appendChild elem2Field
376                         ' Set elem2Field = docStock.createElement("country"):'elem2Field.Text = "GR": ' ' elemField.appendChild elem2Field
                        
378                         m_ypok = f_YPOK

380                         If InStr(f_par1ypok, Left(SQLDT("ATIM"), 1)) > 0 Then
382                             m_ypok = "1"
                            End If

384                         If InStr(f_par2ypok, Left(SQLDT("ATIM"), 1)) > 0 Then
386                             m_ypok = "2"
                            End If

                            'f_YPOK
                            ' ' Set elem2Field = docStock.createElement("branch"):'elem2Field.Text = "0": ' ' elemField.appendChild elem2Field
388                         ' Set elem2Field = docStock.createElement("branch"):'elem2Field.Text = m_ypok: ' ' elemField.appendChild elem2Field
390                        ' .appendChild elemField
                            
392                         SJ = SJ + Chr(13) + " ""issuer"":{ "

394                         SJ = SJ + Chr(13) + " ""vatNumber"": """ + afmCompany + """ "
396                         SJ = SJ + Chr(13) + ",""branch"": " + m_ypok
398                         SJ = SJ + Chr(13) + ",""street"": """ + GGET_CVALUE("SELECT TOP 1 ISNULL(PELDIE,'') AS PELDIE FROM MEM ") + """"
400                         SJ = SJ + Chr(13) + ",""streetNumber"": " + GGET_CVALUE("SELECT TOP 1 ISNULL(PELARIT,'0') AS PELARIT FROM MEM ")
402                         SJ = SJ + Chr(13) + ",""postalCode"": """ + GGET_CVALUE("SELECT TOP 1 ISNULL(PELTK,'0') AS PELTK FROM MEM ") + """"
404                         SJ = SJ + Chr(13) + ",""city"": """ + GGET_CVALUE("SELECT TOP 1 ISNULL(PELPOL,'') AS PELPOL FROM MEM ") + """"
406                         SJ = SJ + Chr(13) + ",""COUNTRY"": " + " ""GR"" "
408                         SJ = SJ + Chr(13) + "}, " 'telos issuer
                                
                            '------------ COUNTER PART ------------------------
410                         If mID(Split(ctypos, ";")(0), 1, 2) = "ΩΩ" Then  'PANTA XREIAZETAI
                                'lianikh den xreiazetai pelaths
                            Else
                                
412                             ' Set elemField = docStock.createElement("counterpart") ' δημιουργω εσοχ
414                             ' Set elem2Field = docStock.createElement("vatNumber"):'elem2Field.Text = Trim(SQLDT("AFM")): ' ' elemField.appendChild elem2Field
416                             ' Set elem2Field = docStock.createElement("country"):'elem2Field.Text = SQLDT("COUNTRY"): ' ' elemField.appendChild elem2Field

                                Dim pelst3 As String: pelst3 = Trim(SQLDT("PELBRANCH"))

                                If Len(pelst3) = 0 Then
                                    pelst3 = "0"
                                End If
                                    
                                Dim pelst2 As String: pelst2 = Trim(SQLDT("PELSTREETNUMBER"))

                                If Len(pelst2) = 0 Then
                                    pelst2 = "0"
                                End If

418                             ' Set elem2Field = docStock.createElement("branch"):'elem2Field.Text = pelst3: ' ' elemField.appendChild elem2Field

420                             If SQLDT("COUNTRY") <> "GR" Then
                        
422                                 ' Set elem2Field = docStock.createElement("name"):'elem2Field.Text = SQLDT("epo"): ' ' elemField.appendChild elem2Field
                                End If
                        
424                             ' Set elem2Field = docStock.createElement("address") ' δημιουργω εσοχ' δημιουργω εσοχ  sqlDT(i)("XRVMA")= TK
426                             ' Set elem3Field = docStock.createElement("postalCode"):'elem3Field.Text = CNull(SQLDT("TK")): ' elem2Field.appendChild elem3Field
428                             ' Set elem3Field = docStock.createElement("city"):'elem3Field.Text = CNull(SQLDT("POL")): ' elem2Field.appendChild elem3Field
430                             ' ' elemField.appendChild elem2Field
432                            ' .appendChild elemField
                                
434                             SJ = SJ + Chr(13) + " ""counterpart"":{ "
436                             SJ = SJ + Chr(13) + " ""vatNumber"":""" + Trim(SQLDT("AFM")) + """"
438                             SJ = SJ + Chr(13) + ", ""name"":""" + Trim(SQLDT("EPO")) + """"
440                             SJ = SJ + Chr(13) + ",""branch"":" + pelst3 ' + """"
442                             SJ = SJ + Chr(13) + ",""street"":""" + Trim(SQLDT("DIE")) + """"
444                             SJ = SJ + Chr(13) + ",""streetNumber"":""" + pelst2 + """"
446                             SJ = SJ + Chr(13) + ",""postalCode"":""" + Trim(SQLDT("TK")) + """"
448                             SJ = SJ + Chr(13) + ",""city"":""" + Trim(SQLDT("pol")) + """"
450                             SJ = SJ + Chr(13) + ",""country"":""" + Trim(SQLDT("COUNTRY")) + """"
                                If m_is_benzinadiko = 1 Then
                                     SJ = SJ + Chr(13) + ",""supplyAccountNo"":""" + Trim(SQLDT("DEH")) + """"
                                End If

452                             If is_B2G = 1 Then
454                                 SJ = SJ + Chr(13) + ",""municipality"":""" + Trim(SQLDT("POL")) + """"
                                End If

456                             SJ = SJ + Chr(13) + "}, " 'telos counterpart
                                
                            End If
                        End If

                        '---------------- HEADER---------------------------
458                     ' Set elemField = docStock.createElement("invoiceHeader") ' δημιουργω εσοχη

                        Dim C_SEIRA As String

460                     C_SEIRA = GET_CVALUE("SELECT SYNT_TITL FROM PARASTAT WHERE EIDOS='" + Left$(SQLDT("ATIM"), 1) + "'")

                        ' If Left(C_SEIRA, 1) = "Τ" Then C_SEIRA = "0" Else C_SEIRA = Trim$(C_SEIRA)
462                     If Len(C_SEIRA) = 0 Then C_SEIRA = "0"

464                     ' Set elem2Field = docStock.createElement("series"):'elem2Field.Text = C_SEIRA: ' ' elemField.appendChild elem2Field
466                     ' Set elem2Field = docStock.createElement("aa"):'elem2Field.Text = mID(SQLDT("ATIM"), 2, 6): ' ' elemField.appendChild elem2Field
468                     ' Set elem2Field = docStock.createElement("issueDate"):'elem2Field.Text = Format(SQLDT("hme"), "yyyy-MM-dd"): ' ' elemField.appendChild elem2Field
470                     ' Set elem2Field = docStock.createElement("invoiceType"):'elem2Field.Text = Split(ctypos, ";")(0): ' ' elemField.appendChild elem2Field

472                     If SQLDT("ANASTOLHFPA") = 1 Then
474                         ' Set elem2Field = docStock.createElement("vatPaymentSuspension"):'elem2Field.Text = "true": ' ' elemField.appendChild elem2Field
                        End If

476                     ' Set elem2Field = docStock.createElement("currency"):'elem2Field.Text = SQLDT("CURRENCY"): ' ' elemField.appendChild elem2Field

                        Dim isot As String

478                     If SQLDT("CURRENCY") = "EUR" Then isot = "1" Else isot = Replace(Format(SQLDT("EXCHANGERATE"), "00.00000"), ",", ".")

480                     ' Set elem2Field = docStock.createElement("exchangeRate"):'elem2Field.Text = isot: ' ' elemField.appendChild elem2Field
                                                       
482                     SJ = SJ + Chr(13) + " ""invoiceHeader"":{ "
484                     SJ = SJ + Chr(13) + " ""series"":""" + C_SEIRA + """"
486                     SJ = SJ + Chr(13) + ",""aa"":""" + mID(SQLDT("ATIM"), 2, 6) + """"
488                     SJ = SJ + Chr(13) + ",""issueDate"":""" + Format(SQLDT("hme"), "yyyy-MM-dd") + """"
490                     SJ = SJ + Chr(13) + ",""dispatchDate"":""" + Format(SQLDT("HMEPARAD"), "yyyy-MM-dd") + """"
                           
492                     SJ = SJ + Chr(13) + ",""invoiceCode"":""" + mID(SQLDT("ATIM"), 1, 7) + """"
494                     SJ = SJ + Chr(13) + ",""invoiceTypeUbl"":""" + ubl + """"
496                     SJ = SJ + Chr(13) + ",""invoiceType"":""" + Split(ctypos, ";")(0) + """"

                        If isDiakin = 2 Then
                        Else
498                         SJ = SJ + Chr(13) + ",""currency"":""" + "EUR" + """"
                        End If
                        
                        '                           he.movePurpose = Val(SQLDT.rows(klm)("SKOPOS").toString.Substring(0, 2))
                        '                        If he.movePurpose = 19 Then
                        '                            he.OTHERMOVEPURPOSETITLE = SQLDT.rows(klm)("OTHERMOVEPURPOSETITLE").toString
                        '                        End If

500                     SJ = SJ + Chr(13) + ",""movePurpose"":" + LTrim(str(Val(SQLDT("SKOPOS"))))

                        If Val(SQLDT("SKOPOS")) = 19 Then
                            SJ = SJ + Chr(13) + ",""otherMovePurposeTitle"":""" + SQLDT("OTHERMOVEPURPOSETITLE") + """"
                        
                        End If

           

502

                        If m_is_benzinadiko = 0 Then

504                         SJ = SJ + Chr(13) + ",""fuelInvoice"":false"
                        Else
                            SJ = SJ + Chr(13) + ",""fuelInvoice"":true"
                        End If

506                     If is_B2G = 1 Then
508                         SJ = SJ + Chr(13) + ",""paymentTerms"":""" + mID(SQLDT("TRP"), 3, 7) + """"
                        End If
                        
                        
                        
                                     '------------- ΑΥΤΟΤΙΜΟΛΟΓΗΣΗ --------------------<selfPricing>true</selfPricing>
                        If IsAytotim = 1 Then

                            ' ΑΥΤΟΤΙΜΟΛΟΓΗΣΗ ΜΗΛΙΟΥ
                            If SQLDT("EIDOS") = "r" Then
                                If Left(mTypPar, 3) = "1.1" Then     ' Left(mTypPar, 3) = "3.1" Or Left(mTypPar, 3) = "13." Then ' TITLOS KTHSHS- αγορα εε
                                    '' Set elem2Field = docStock.createElement("selfPricing"):'elem2Field.Text = "true": ' ' elemField.appendChild elem2Field
                                    SJ = SJ + Chr(13) + ",""selfPricing"":true"
                                End If

                            Else
                                '' Set elem2Field = docStock.createElement("selfPricing"):'elem2Field.Text = "true": ' ' elemField.appendChild elem2Field
                                SJ = SJ + Chr(13) + ",""selfPricing"":false"
                            End If
                            
                        End If
                        
                        

'510                     If IsAytotim = 1 Then
'512                         Sj = Sj + Chr(13) + ",""selfPricing"":true"
'                        End If
                        
                        
                        
                        
                        
                        
602                     SJ = SJ + Chr(13) + ",""vatPaymentSuspension"":false "
604                     SJ = SJ + Chr(13) + ",""vehicleNumber"":""" + SQLDT("AYTOK") + """  "
                         
                        ' "specialInvoiceCategory": 11,
                        '"otherCorrelatedEntities": [
                        '{
                        '"type": ??, => T? S???????????? ??? ?????? ??? MyData
                        '"entityData": {
                        '"vatNumber": "999999999",
                        '"branch": 0,
                        '"name": "Name",
                        '"documentIdNo": null,
                        '"countryDocumentId": null,
                        '"supplyAccountNo": "123456789",

                        '"address": {
                        '"country": "GR",  **
                        '"street": "str",***
                        '"number": "123",
                        '"postalCode": "12345",
                        '"city": "Athens"
                        '}
                        '}
                        '}
                        ']
                        '},
                        '"invoiceDetails":
                        If m_is_benzinadiko = 1 Then

                            Dim mTHERM As Integer

                            mTHERM = Val(FINDPARAMETROI(1, "PAR1", "F_THERM", "0", "προκειται για θερμανση=1 οχι=0")) '
                            'otherent(1) = New otherCorrelatedEntitity
                            
                            Dim CODE_kays As String

                            CODE_kays = GGET_CVALUE("SELECT KODE FROM EGGTIM WHERE ID_NUM=" + str(SQLDT("ID_NUM"))) 'ΠΑΙΡΝΩ ΤΟΝ ΚΩΔΙΚΟ ΠΟΥ ΠΟΥΛΗΘΗΚΕ
                            
                            If mTHERM = 1 And CODE_kays = "30" Then '=========== T?????S? =========
                                SJ = SJ + Chr(13) + ",""specialInvoiceCategory"":11  "
                                'he.specialInvoiceCategory = 11
                            End If
                            
                            SJ = SJ + Chr(13) + ",""otherCorrelatedEntitities"":[{  "                                   'otherCorrelatedEntitity
                            SJ = SJ + Chr(13) + " ""type"":1 "
                            SJ = SJ + Chr(13) + ",""entityData"":{ "
                            SJ = SJ + Chr(13) + " ""vatNumber"":""" + Trim(SQLDT("AFM")) + """"
                            SJ = SJ + Chr(13) + " , ""branch"":0 "
                            SJ = SJ + Chr(13) + ",""Name"":""" + Trim(SQLDT("EPO")) + """"
                                                           
                      '      Sj = Sj + Chr(13) + ",""supplyAccountNo"":""" + Trim(SQLDT("DEH")) + """"
                            SJ = SJ + Chr(13) + ",""address"":{ "
                            SJ = SJ + Chr(13) + " ""country"":""GR"" "
                            SJ = SJ + Chr(13) + ",""street"":""" + Trim(SQLDT("DIE")) + """"
                            SJ = SJ + Chr(13) + ",""number"":""" + Trim(SQLDT("STREETNUMBER")) + """"
                            SJ = SJ + Chr(13) + ",""postalCode"":""" + Trim(SQLDT("TK")) + """"
                            SJ = SJ + Chr(13) + " ,""city"":""" + Trim(SQLDT("POL")) + """"
                            SJ = SJ + Chr(13) + " }}}]"
                        End If  '  If m_is_Benzinadiko = 1 Then
                        
606                     If isDiakin >= 1 Then   'tda=1  δελ.αποστ=2
                       
608                         If isDiakin = 1 Then   'tda=1
610                             SJ = SJ + Chr(13) + ",""isDeliveryNote"":1 "
                            End If
                       
612                         SJ = SJ + Chr(13) + ",""otherDeliveryNoteHeader"": { " '----------------
                                 
614                         SJ = SJ + Chr(13) + " ""deliveryAddress"": { "
616                         SJ = SJ + Chr(13) + " ""street"": """ + CNull(SQLDT("STREET")) + """  "
618                         SJ = SJ + Chr(13) + ",""number"": """ + CNull(SQLDT("STREETNUMBER")) + """  "
620                         SJ = SJ + Chr(13) + ",""postalCode"": """ + CNull(SQLDT("POSTALCODE")) + """  "
622                         SJ = SJ + Chr(13) + ",""city"": """ + CNull(SQLDT("CITY")) + """  "
624                         SJ = SJ + Chr(13) + " } "
                                    
626                         SJ = SJ + Chr(13) + ",""loadingAddress"": { "
                                   
628                         SJ = SJ + Chr(13) + " ""street"": """ + GGET_CVALUE("SELECT TOP 1 ISNULL(PELDIE,'') AS PELDIE FROM MEM ") + """"
630                         SJ = SJ + Chr(13) + ",""number"": " + GGET_CVALUE("SELECT TOP 1 ISNULL(PELARIT,'0') AS PELARIT FROM MEM ")
632                         SJ = SJ + Chr(13) + ",""postalCode"": """ + GGET_CVALUE("SELECT TOP 1 ISNULL(PELTK,'0') AS PELTK FROM MEM ") + """"
634                         SJ = SJ + Chr(13) + ",""city"": """ + GGET_CVALUE("SELECT TOP 1 ISNULL(PELPOL,'') AS PELPOL FROM MEM ") + """"
                        
636                         SJ = SJ + Chr(13) + " } "
                                     
638                         SJ = SJ + Chr(13) + ",""startShippingBranch"": 0  "
640                         SJ = SJ + Chr(13) + ",""completeShippingBranch"": 0  "
642                         SJ = SJ + Chr(13) + " } " '---------------otherDeliveryNoteHeader"

                        Else
                            '                            Dim NSXETMARKS As Integer
                            '                            If InStr(sqlDt.rows(klm)("SXETMARKS"), ",") > 0 Then
                            '                                For NSXETMARKS = 0 To 99
                            '                                    If Len(Split(sqlDt.rows(klm)("SXETMARKS"), ",")(NSXETMARKS)) > 0 Then
                            '                                        elem2Field = docStock.createElement("multipleConnectedMarks"):'elem2Field.Text = Split(sqlDt.rows(klm)("SXETMARKS"), ",")(NSXETMARKS): ' ' elemField.appendChild (elem2Field)
                            '                                    Else
                            '                                        Exit For
                            '                                    End If
                            '                                Next
                            '                            End If
                        End If
                              
514                     SJ = SJ + Chr(13) + "} " 'telos invoiceHeader
                          
516                     If is_B2G = 1 Then
                          
518                         SJ = SJ + Chr(13) + " , ""publishType"":2"

                            'organizationalUnitName 2 NO String BT-10 Name of internal organizational unit of public authority that will receive the invoice, if any.
                            'organizationalUnitCode 2 NO String BT-10 Code of internal organizational unit of public authority that will receive the invoice, if any.

                            'name 2 YES String BT-44 Counterpart’s Company Official Title
                           
                            '                           "contractingAuthorityID": "997687965",
                            '                           "budget": {
                            '                                  "type": 3,
                            '                                  "identifier": "??F?7??-?3S"
                            '                                      },
                           
                            ' publ.contractingAuthorityID = sqlDt.Rows(klm)("CONTRACTINGAUTHORITYID")  ' "1017.0000000000.0001"
                            ' bud.type = sqlDt.Rows(klm)("BUDTYPE") ' 1         '"??p?? ????p?????sµ??"   '?p????? ??µ?? ap? ta pa?a??t?:
                           
520                         SJ = SJ + Chr(13) + ",""publishDetails"":{ "
522                         SJ = SJ + Chr(13) + " ""contractingAuthorityID"":""" + SQLDT("contractingAuthorityID") + """"
524                         SJ = SJ + Chr(13) + ",""budget"":{ "
526                         SJ = SJ + Chr(13) + " ""type"": " + SQLDT("BUDTYPE") + ""
528                         SJ = SJ + Chr(13) + ",""identifier"":""" + SQLDT("BUDIDENTIFIER") + """"
530                         SJ = SJ + Chr(13) + "} " 'telos budget
532                         SJ = SJ + Chr(13) + ",""contractIdentifier"":""" + SQLDT("contractIdentifier") + """"
534                         SJ = SJ + Chr(13) + ",""organizationalUnitName"":""" + SQLDT("ORGANIZATIONALUNITNAME") + """"
536                         SJ = SJ + Chr(13) + ",""organizationalUnitCode"":""" + SQLDT("organizationalUnitCode") + """"
                           
538                         SJ = SJ + Chr(13) + "} " 'telos publishDetails
                          
                        End If
                                                       
                        '                        '------------- ΑΥΤΟΤΙΜΟΛΟΓΗΣΗ --------------------<selfPricing>true</selfPricing>
                        '540                     If IsAytotim = 1 Then
                        '
                        '                            ' ΑΥΤΟΤΙΜΟΛΟΓΗΣΗ ΜΗΛΙΟΥ
                        '
                        '542                         If SQLDT("EIDOS") = "r" Then
                        '544                             If Left(mTypPar, 3) = "1.1" Then     ' Left(mTypPar, 3) = "3.1" Or Left(mTypPar, 3) = "13." Then ' TITLOS KTHSHS- αγορα εε
                        '546                                 ' Set elem2Field = docStock.createElement("selfPricing"):'elem2Field.Text = "true": ' ' elemField.appendChild elem2Field
                        '                                End If
                        '
                        '                            Else
                        '548                             ' Set elem2Field = docStock.createElement("selfPricing"):'elem2Field.Text = "true": ' ' elemField.appendChild elem2Field
                        '                            End If
                        '
                        '
                        '
                        '
                        '
                        '                        End If

                        '                                 "paymentMethods": [
                        '                                 {          "type": 5,
                        '"amount": 0       }     ],
                        '  "correlatedInvoices": [
                        ' {         "extSystemId": "string",          "mark": 0,        }     ],
                        '
                                                       
550                     If Split(ctypos, ";")(0) = "5.1" Then  ' -------- συσχετιζομενο
552                         CR8 = "": CR8B = ""

554                         If IsNull(SQLDT("SXETMARK")) Then
                        
                            Else
                                ' If Len(sqlDt("SXETMARK")) > 10 Then
                                '    CR8 = Split(sqlDt("SXETMARK"), " ")(1)
                                '    CR8B = Split(sqlDt("SXETMARK"), " ")(2)
                                ' End If
                           
556                             If Len(SQLDT("SXETMARK")) > 10 Then
558                                 CR8 = Split(SQLDT("SXETMARK"), "*")(0)
                                    ' CR8B = Split(sqlDt("SXETMARK"), " ")(2)
                                    '  Else
                                 
                                    Dim NSX As Integer

560                                 For NSX = 0 To 10

562                                     If Len(Split(SQLDT("SXETMARK"), ";")(NSX)) > 0 Then
564                                         CR8 = Split(SQLDT("SXETMARK"), ";")(NSX)
566                                         ' Set elem2Field = docStock.createElement("correlatedInvoices"):'elem2Field.Text = CR8: ' ' elemField.appendChild elem2Field
                                        Else

                                            Exit For

                                        End If
                                 
                                    Next
                                 
                                End If
                                
568                             SJ = SJ + Chr(13) + ",""correlatedInvoices"":[{ ""extSystemId"":null, ""mark"":" + CR8 + " }] "
                                
                                '       "correlatedInvoices": [
                                '        {
                                '             "extSystemId": null,
                                '              "mark": "400001946173070"
                                '           }
                                '        ],
                           
                            End If
                        
                            ' CR8B = Split(sqlDt("SXETMARK"), " ")(3)
                        
                            '288                         R8.Open "SELECT ENTITYMARK FROM TIM WHERE ATIM='" + CR8 + "' AND CONVERT(CHAR(10),HME,103)='" + CR8B + "'", Gdb, adOpenDynamic, adLockOptimistic
                            '290                         CR8 = ""
                            '292                         If IsNull(R8(0)) Then
                            '294                             CR8 = ""
                            '                            Else
                            '
                            '296                             If R8.EOF Then
                            '                                Else
                            '298                               CR8 = correlatedInvoice
                            '                                End If
                            '                            End If
                            '300                         R8.Close
                            'CorrelatedInvoices
                             
                            ' ' Set elem2Field = docStock.createElement("correlatedInvoices"):'elem2Field.Text = CR8: ' ' elemField.appendChild elem2Field
                                                       
                        End If

570                    ' .appendChild elemField
                        
                        Dim IS_POS As Integer: IS_POS = 0

                        If InStr(mID(SQLDT("TRP"), 3, 7), "POS") > 0 Then
                            IS_POS = 1
                        End If

                        If IS_POS = 1 And ubl = "380" Then

                            Dim mTRP    As String: mTRP = SQLDT("TRP")

                            Dim MATIM   As String: MATIM = SQLDT("atim")

                            Dim mID_NUM As String: mID_NUM = LTrim(str(SQLDT("ID_NUM"))) 'Aritmitiras(56)  F_posSignature,POSTransactionId

                            Dim MKAU    As Double, MFPA As Double, M_SYN As Double

                            Dim TID     As String

                            TID = GGET_CVALUE("select ISNULL(C1,'') AS C1 from PINAKES where TYPOS=12 AND  AYJON=" + mID(mTRP, 1, 1)) ' Mid(sqlDt.Rows(klm)("TRP"), 1, 1))
                            TID = Trim(TID) '.TrimEnd()
                           
                            
                            
                            MFPA = SQLDT("FPA1") + SQLDT("FPA2") + SQLDT("FPA3") + SQLDT("FPA4") + SQLDT("FPA6") + SQLDT("FPA7")
                            MKAU = SQLDT("AJ1") + SQLDT("AJ2") + SQLDT("AJ3") + SQLDT("AJ4") + SQLDT("AJ5") + SQLDT("AJ6") + SQLDT("AJ7")
                            
                            If Len(TID) > 2 Then
                                POSTransactionId = to_pos(TID, mID_NUM, MKAU, MFPA, SQLDT("HME"), SQLDT("ATIM"), SQLDT("BRANCH")) ', POSTransactionId, F_posSignature, tid, issu, he, invs)
                            Else
                                MsgBox "δεν εχει δηλωθει TID POS  (C1 PINAKES TYPOS 12)"
                                ToJason = 0
                                Exit Function
                            End If
                            
                            If POSTransactionId = "error" Then
                                ToJason = 0

                                Exit Function

                            End If

                            MILSEC 2000

                            'POSTransactionId = FERE_TRANSACTIONID(mID_NUM)
                            If Len(POSTransactionId) < 5 Then
                                ToJason = 0

                                ' POSTransactionId = FERE_TRANSACTIONID(mID_NUM - 1)
                                Exit Function

                            End If
                            
                        End If

                        'Dim TID As String
                        '                        If IS_POS = 1222 And ubl = "380" Then
                        '
                        '                            TID = GGET_CVALUE("select ISNULL(C1,'') AS C1 from PINAKES where TYPOS=12 AND  AYJON=" + mID(SQLDT("TRP"), 1, 1))
                        '                            TID = Trim(TID)
                        '
                        '                            Dim QQ As String
                        '
                        '                            QQ = "{""externalSystemId"":""" + SQLDT("id_num") + """ " + ","
                        '                            QQ = QQ + " ""issuerVatNumber"":""" + afmCompany + """" + ","
                        '                            QQ = QQ + " ""invoiceIssueDate"":""" + Format(SQLDThme, "yyyy-MM-dd") + ""","
                        '
                        '                            QQ = QQ + " ""companyBranch"":""" + SQLDTBRANCH + ""","
                        '
                        '                            QQ = QQ + " ""invoiceSeries"":""" + C_SEIRA + ""","
                        '                            QQ = QQ + " ""invoiceAA"":""" + mID(SQLDTATIM, 2, 6) + ""","
                        '                             QQ = QQ + " ""invoiceType"":""" + Split(ctypos, ";")(0) + ""","
                        '
                        '                            QQ = QQ + " ""netValue"":" + Replace(Format(MKAU, "#####0.00"), ",", ".") + ","
                        '                            QQ = QQ + " ""VatAmount"":" + Replace(Format(MFPA, "#####0.00"), ",", ".") + ","
                        '                            QQ = QQ + " ""totalValue"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
                        '
                        '                            QQ = QQ + " ""paymentAmount"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
                        '                            QQ = QQ + " ""NspCode"":""2"" , "
                        '                            QQ = QQ + " ""terminalId"":""" + TID + """ }"     '"99999069"
                        '
                        '
                        '
                        '
                        '
                        '
                        '
                        '                            QQ = "{""externalSystemId"":""" + LTrim(str(SQLDT("id_num"))) + """ " + ","
                        '                            QQ = QQ + " ""issuerVatNumber"":""" + afmCompany + """" + ","
                        '                            QQ = QQ + " ""invoiceIssueDate"":""" + Format(SQLDT("hme"), "yyyy-MM-dd") + ""","
                        '
                        '                            QQ = QQ + " ""companyBranch"":""" + SQLDT("BRANCH") + ""","
                        '
                        '                            QQ = QQ + " ""invoiceSeries"":""" + C_SEIRA + ""","
                        '                            QQ = QQ + " ""invoiceAA"":""" + mID(SQLDT("ATIM"), 2, 6) + ""","
                        '                             QQ = QQ + " ""invoiceType"":""" + Split(ctypos, ";")(0) + ""","
                        '
                        '                             MFPA = SQLDT("FPA1") + SQLDT("FPA2") + SQLDT("FPA3") + SQLDT("FPA4") + SQLDT("FPA6") + SQLDT("FPA7")
                        '                                                MKAU = SQLDT("AJ1") + SQLDT("AJ2") + SQLDT("AJ3") + SQLDT("AJ4") + SQLDT("AJ5") + SQLDT("AJ6") + SQLDT("AJ7")
                        '
                        '                        QQ = QQ + " ""netValue"":" + Replace(Format(MKAU, "#####0.00"), ",", ".") + ","
                        '                            QQ = QQ + " ""VatAmount"":" + Replace(Format(MFPA, "#####0.00"), ",", ".") + ","
                        '                            QQ = QQ + " ""totalValue"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
                        '
                        '                            QQ = QQ + " ""paymentAmount"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
                        '                            QQ = QQ + " ""NspCode"":""2"" , "
                        '                            QQ = QQ + " ""terminalId"":""" + TID + """ }"     '"99999069"
                        '
                        '
                        '
                        '
                        '
                        '
                        '
                        '                            Dim ResultJSONpos As String
                        '
                        '                            ResultJSONpos = QQ
                        '
                        '                            Dim Reqpos As WinHttp.WinHttpRequest
                        '
                        '                            Set Reqpos = New WinHttp.WinHttpRequest
                        '
                        '                            ResultJSONpos = Replace(ResultJSONpos, "{", "")
                        '                            ResultJSONpos = Replace(ResultJSONpos, "}", "")
                        '
                        '                            Dim posuid       As String: posuid = ""
                        '
                        '                            Dim posTimestamp As String: posTimestamp = ""
                        '
                        '                            Dim KN           As Integer
                        '
                        '                            For KN = 1 To 1
                        '
                        '                                '
                        '                                Dim qpos   As String
                        '
                        '                                Dim uripos As String: uripos = gfURL1 + "api/requestPayment"
                        '
                        '680:                            With Reqpos
                        '                                    .Open "POST", uripos, async:=False
                        '                                    .setRequestHeader "Content-Type", "application/hal+json"
                        '                                    .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
                        '                                    .setRequestHeader "Authorization", "Bearer " + gf_Bearer
                        '                                    .setRequestHeader "X-Version", "3"
                        '
                        '                                    qpos = " {  ""Bearer"":""" + gf_Bearer + """, " & ResultJSONpos & ",     ""externalSystemId"": """ + LTrim(str(SQLDT("ID_NUM"))) + """ }"
                        '                                    .send (qpos)
                        '
                        '                                    MsgBox ("πατηστε την Γεφυρα." + Chr(13) + Reqpos.responseText)
                        '
                        '                                End With
                        '
                        '                                Dim posjson As String: posjson = Reqpos.responseText
                        '
                        '                                Dim ser     As Object
                        '
                        '                                Set ser = JSON.parse(posjson)
                        '
                        '                                F_PAROX_SIGNATURE = ser.Item("paymentToken")("signature")
                        '                                posTimestamp = ser.Item("paymentToken")("timestamp")
                        '                                posuid = ser.Item("uid")
                        '                            Next
                        '
                        '                            Dim userpos As String, PWDPOS As String
                        '
                        '                            userpos = GGET_CVALUE("SELECT ISNULL(USERPOS,'') AS U FROM MEM")
                        '                            PWDPOS = GGET_CVALUE("SELECT ISNULL(PWDPOS,'') AS U FROM MEM")
                        '
                        '                            'zhtao token (paris1) apo lamdapi
                        '                            Dim paris1 As String: paris1 = request_pliromis(userpos, PWDPOS)
                        '
                        '                            QQ = "{""terminal_id"":""" + TID + """ " + ","
                        '                            QQ = QQ + " ""paymentAmount"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
                        '                            QQ = QQ + " ""netValue"":" + Replace(Format(MKAU, "#####0.00"), ",", ".") + ","
                        '                            QQ = QQ + " ""vatAmount"":" + Replace(Format(MFPA, "#####0.00"), ",", ".") + ","
                        '                            QQ = QQ + " ""totalValue"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
                        '                            QQ = QQ + " ""externalSystemId"":""" + LTrim(str(SQLDT("id_num"))) + """ " + ","
                        '                            QQ = QQ + " ""uid"":""" + posuid + """ " + ","
                        '                            QQ = QQ + " ""TimeStamp"":""" + posTimestamp + """ " + ","
                        '                            QQ = QQ + " ""signature"":""" + F_PAROX_SIGNATURE + """ " + "}"
                        '                            offlineJSONpos = QQ
                        '
                        '                            Dim req2 As String: req2 = request2_pliromis(paris1, offlineJSONpos)
                        '
                        '                            Dim mRes As String: mRes = Split(req2, "@")(0)
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
                        '8                            If mRes <> "1" Then
                        '
                        '                                Dim HH As Integer
                        '
                        '                                'Exit Function
                        '                                For HH = 1 To 15
                        '                                    MILSEC 5000 ' Threading.Thread.Sleep (5000) '5sec
                        '                                    req2 = FERE_TRANSACTIONID(extID)
                        '
                        '                                    If Len(req2) > 5 Then
                        '                                        POSTransactionId = Split(req2, "@")(1)
                        '
                        '                                        Exit For
                        '
                        '                                    End If
                        '
                        '                                Next
                        '
                        '                                If Len(req2) < 5 Then
                        '                                    ANSPOS = MsgBox("Εγινε η συναλλαγή;", vbYesNo)
                        '
                        '                                    If ANSPOS = vbYes Then
                        '
                        '                                        If Len(req2) > 5 Then
                        '                                            POSTransactionId = Split(req2, "@")(1)
                        '                                        Else
                        '                                            ToJasonSub = 0
                        '
                        '                                            Exit Function
                        '
                        '                                        End If 'Len(req2)
                        '
                        '                                    Else
                        '
                        '                                        If gVal(mRes) = 2 Or gVal(mRes) = 3 Then ' ap????f???e, a???????e
                        '
                        '                                            CancelPosPayment extID, posSignature
                        '                                            ToJason = 0
                        '
                        '                                            Exit Function
                        '
                        '                                        End If 'gVal(mRes) = 2
                        '
                        '                                        ToJason = 0
                        '
                        '                                        Exit Function
                        '
                        '                                    End If 'ANSPOS = vbYes Then
                        '
                        '                                End If ' Len(req2) < 5
                        '
                        '                            End If ' mRes <> "1"
                        '
                        '                        End If  ' is_pos=1

572                     Dim cTrp As String: cTrp = FindTRP(mID(SQLDT("TRP"), 1, 1))

574                     If Len(cTrp) = 0 Then
576                         If noask = 0 Then
578                             MsgBox ("ΔΕΝ ΕΧΩ ΑΝΤΙΣΤΟΙΧΙΣΗ ΣΤΟΝ ΤΡΟΠΟ ΠΛΗΡΩΜΗΣ " + SQLDT("TRP"))
                            End If

                            Exit Function

                        End If
                        
                        If isDiakin <= 1 Then  ' if isDiakin =2 then ειναι δελτιο αποστολης οποτε οχι payment
                        
                            '----------------------- PAYMENT ---------------------------------

580                         If SQLDT("EIDOS") = "r" And Left(mTypPar, 3) <> "3.1" And IsAytotim = 0 Then   ' TITLOS KTHSHS Then
                            Else

582                             ' Set elemField = docStock.createElement("paymentMethods") ' δημιουργω εσοχ
                            
584                             ' Set elem2Field = docStock.createElement("paymentMethodDetails") ' δημιουργω εσοχ' δημιουργω εσοχ
586                             ' Set elem3Field = docStock.createElement("type"):'elem3Field.Text = cTrp: ' elem2Field.appendChild elem3Field
588                             ' Set elem3Field = docStock.createElement("amount"):'elem3Field.Text = chDec(Format(SUMNET + SumFpa, "######0.##")): ' elem2Field.appendChild elem3Field
590                             ' ' elemField.appendChild elem2Field
592                            ' .appendChild elemField

                                '594                             Sj = Sj + Chr(13) + ",""paymentMethods"":[{ ""type"":" + cTrp ' """"
                                '596                             Sj = Sj + Chr(13) + ",""amount"":" + chDec(Format(Round(SUMNET + SumFpa, 2), "####0.00"))
                                '598
                                '
                                '                                If IS_POS = 1 Then
                                '
                                ''                                    Sj = Sj + Chr(13) + ",""terminalId"":""" + TID + """"
                                ''                                    Sj = Sj + Chr(13) + ",""nspCode"":2"
                                ''                                    Sj = Sj + Chr(13) + ",""tipAmount"":0"
                                ''                                    Sj = Sj + Chr(13) + ",""signature"":""" + F_PAROX_SIGNATURE + """"
                                ''                                    Sj = Sj + Chr(13) + ",""transactionId"":""" + POSTransactionId + """"
                                '
                                '                                End If
                                '
                                '600                             Sj = Sj + Chr(13) + "}] " 'telos paymentMethods

                            End If
                        
                        End If ' isDiakin <= 1 Then  ' if isDiakin =2 then ειναι δελτιο αποστολης οποτε οχι payment

                        'gia na doyleyei h diakinish
                     
                        '602                       Sj = Sj + Chr(13) + ",""vatPaymentSuspension"":false "
                        '604                       Sj = Sj + Chr(13) + ",""vehicleNumber"":""xxx0000""  "
 
                        '606                     If isDiakin >= 1 Then   'tda=1  δελ.αποστ=2
                        '
                        '608                         If isDiakin = 1 Then   'tda=1
                        '610                               Sj = Sj + Chr(13) + ",""isDeliveryNote"":1 "
                        '                            End If
                        '
                        '612                           Sj = Sj + Chr(13) + ",""OtherDeliveryNoteHeader"": { " '----------------
                        '
                        '614                           Sj = Sj + Chr(13) + " ""deliveryAddress"": { "
                        '616                           Sj = Sj + Chr(13) + " ""street"": """ + CNull(sqlDt("STREET")) + """  "
                        '618                           Sj = Sj + Chr(13) + ",""number"": """ + CNull(sqlDt("STREETNUMBER")) + """  "
                        '620                           Sj = Sj + Chr(13) + ",""postalCode"": """ + CNull(sqlDt("POSTALCODE")) + """  "
                        '622                           Sj = Sj + Chr(13) + ",""city"": """ + CNull(sqlDt("CITY")) + """  "
                        '624                           Sj = Sj + Chr(13) + " } "
                        '
                        '626                           Sj = Sj + Chr(13) + ",""loadingAddress"": { "
                        '
                        '628                           Sj = Sj + Chr(13) + " ""street"": """ + GGET_CVALUE("SELECT TOP 1 ISNULL(PELDIE,'') AS PELDIE FROM MEM ") + """"
                        '630                           Sj = Sj + Chr(13) + ",""number"": " + GGET_CVALUE("SELECT TOP 1 ISNULL(PELARIT,'0') AS PELARIT FROM MEM ")
                        '632                           Sj = Sj + Chr(13) + ",""postalCode"": """ + GGET_CVALUE("SELECT TOP 1 ISNULL(PELTK,'0') AS PELTK FROM MEM ") + """"
                        '634                           Sj = Sj + Chr(13) + ",""city"": """ + GGET_CVALUE("SELECT TOP 1 ISNULL(PELPOL,'') AS PELPOL FROM MEM ") + """"
                        '
                        '636                           Sj = Sj + Chr(13) + " } "
                        '
                        '638                           Sj = Sj + Chr(13) + ",""startShippingBranch"": 0  "
                        '640                           Sj = Sj + Chr(13) + ",""completeShippingBranch"": 0  "
                        '642                           Sj = Sj + Chr(13) + " } " '---------------otherDeliveryNoteHeader"
                        '
                        '                        Else
                        '                            '                            Dim NSXETMARKS As Integer
                        '                            '                            If InStr(sqlDt.rows(klm)("SXETMARKS"), ",") > 0 Then
                        '                            '                                For NSXETMARKS = 0 To 99
                        '                            '                                    If Len(Split(sqlDt.rows(klm)("SXETMARKS"), ",")(NSXETMARKS)) > 0 Then
                        '                            '                                        elem2Field = docStock.createElement("multipleConnectedMarks"):'elem2Field.Text = Split(sqlDt.rows(klm)("SXETMARKS"), ",")(NSXETMARKS): ' ' elemField.appendChild (elem2Field)
                        '                            '                                    Else
                        '                            '                                        Exit For
                        '                            '                                    End If
                        '                            '                                Next
                        '                            '                            End If
                        '                        End If

                        Dim SYN_KAU As Double, SYN_FPA As Double

644                     SYN_KAU = 0
646                     SYN_FPA = 0

                        Dim fpaRow          As Double

                        Dim L               As Integer
                    
648                     Dim SumEsodExod(10) As Double: For L = 1 To 10: SumEsodExod(L) = 0: Next
                     
650                     L = 0
652                     SJ = SJ + Chr(13) + ",""invoiceDetails"":[" 'arxizoyn oi seires=======================================================================

                        '====================================================================================================================
654                     Do While Not EGGTIM.EOF
                            'For n = 1 To 3 ' SEIRES TIMOLOGIOY
656                         L = L + 1

                            Dim AJ As Double

658                         If IsNull(EGGTIM("KAU_AJIA")) Then
660                             AJ = 0
                            Else
662                             AJ = EGGTIM("KAU_AJIA")  ' Math.Round(EGGTIM(L)("POSO") * EGGTIM(L)("TIMM") * (1 - EGGTIM(L)("EKPT") / 100), 2)
                            End If

664                         SumEsodExod(EGGTIM("KATHGORIA")) = SumEsodExod(EGGTIM("KATHGORIA")) + AJ

                            Dim VAT As String

                            '1 ΦΠΑ συντελεστής 24% 24%
                            '2 ΦΠΑ συντελεστής 13% 13%
                            '3 ΦΠΑ συντελεστής 6% 6%
                            '4 ΦΠΑ συντελεστής 17% 17%
                            '5 ΦΠΑ συντελεστής 9% 9%
                            '6 ΦΠΑ συντελεστής 4% 4%
                            '7 Άνευ Φ.Π.Α. 0%
                            '8 Εγγραφές χωρίς ΦΠΑ  (πχ Μισθοδοσία, Αποσβέσεις)
                            
666                         VAT = Format(antFPA(EGGTIM("FPA")), "0")
668                         SYN_KAU = SYN_KAU + AJ

670                         If EGGTIM("FPA") = 5 Then
672                             fpaRow = 0
                            Else
674                             fpaRow = EGGTIM("MIK_AJIA") - EGGTIM("KAU_AJIA")

676                             If fpaRow > 0 And OK_DIFF = 0 Then
678                                 fpaRow = fpaRow + DIFF_FPA
                                    OK_DIFF = 1
682                                 DIFF_FPA = 0
                                  
                                End If
                                  
                            End If

684                         SYN_FPA = SYN_FPA + fpaRow
                            
                            '
                            '-----------------------------------------------  invoiceDetails
                        
686                         ' Set elemField = docStock.createElement("invoiceDetails") ' δημιουργω εσοχ
688                         ' Set elem2Field = docStock.createElement("lineNumber"):'elem2Field.Text = str(L): ' ' elemField.appendChild elem2Field
690                         ' Set elem2Field = docStock.createElement("netValue"):'elem2Field.Text = chDec(Format(AJ, "######0.##")): ' ' elemField.appendChild elem2Field
692                         ' Set elem2Field = docStock.createElement("vatCategory"):'elem2Field.Text = VAT: ' ' elemField.appendChild elem2Field

694                         ' Set elem2Field = docStock.createElement("vatAmount"):'elem2Field.Text = chDec(Format(fpaRow, "######0.##")): ' ' elemField.appendChild elem2Field

696                         If L = 1 Then
698                             SJ = SJ + Chr(13) + "{ ""lineNumber"":" + str(L)
                            Else
700                             SJ = SJ + Chr(13) + ", { ""lineNumber"":" + str(L)
                            End If
                           
702                         SJ = SJ + Chr(13) + ",""recType"":0"
704                         SJ = SJ + Chr(13) + ",""quantity"":" + Replace(Format(Round(EGGTIM("POSO"), 2), "#####0.00"), ",", ".") ' Replace(str(EGGTIM("POSO")), ",", ".")
706                         SJ = SJ + Chr(13) + ",""entityName"":""" + Replace(EGGTIM("onoma"), """", "`") + """"
                           
708                         SJ = SJ + Chr(13) + ",""netValue"":" + Replace(Format(Round(AJ, 2), "#####0.00"), ",", ".")
710                         SJ = SJ + Chr(13) + ",""vatCategory"":" + Replace(str(VAT), ",", ".")
                            ' 1ο μερος detail
                            
                            '  Sj = Sj + Chr(13) + ",""itemCode"":" + EGGTIM("KODE")
                            '   Sj = Sj + Chr(13) + ",""itemDescr"":" + EGGTIM("ONOMA")

                            If m_is_benzinadiko = 1 Then
                                ' newitem(NEGG).fuelCode = EGGTIM.rows(NEGG)("KODE") ' 30 ' "ITEMCODE"
                                SJ = SJ + Chr(13) + ",""fuelCode"":" + EGGTIM("KODE")
                            End If

712                         Dim FPAUBL As String: FPAUBL = GGET_CVALUE("Select ISNULL(C2,'') AS UBL FROM PINAKES  WHERE TYPOS=1 AND N1=" + VAT)

714                         If Len(FPAUBL) = 0 Then
716                             MsgBox ("δεν εχω ΦΠΑ UBL ΣΕΙΡΑ " + str(L))
718                             ToJason = 0

                                Exit Function

                            End If
                            
720                         SJ = SJ + Chr(13) + ", ""vatCategoryUbl"":""" + Trim(FPAUBL) + """"
722                         SJ = SJ + Chr(13) + ",""vatAmount"":" + Replace(Format(Round(fpaRow, 2), "####0.00"), ",", ".")
724                         SJ = SJ + Chr(13) + ",""vatPercent"":" + Replace(str(IIf(VAT = 1, 24, IIf(VAT = 2, 13, IIf(VAT = 3, 6, IIf(VAT = 4, 17, IIf(VAT = 5, 9, IIf(VAT = 6, 4, 0))))))), ",", ".")

                            'newitem(NEGG).vatAmount = System.Math.Round(fpaRow, 2) 'fpaRow '24
                            'newitem(NEGG).vatPercent = IIf(VAT = 1, 24, IIf(VAT = 2, 13, IIf(VAT = 3, 6, IIf(VAT = 4, 17, IIf(VAT = 5, 9, IIf(VAT = 6, 4, 0))))))
                            ' newitem(NEGG).MeasurementUnit = 1
                          
                            Dim monada As String: monada = EGGTIM("MONA")

                            Dim Unitmy As String

                            Unitmy = GGET_NVALUE("select str(TIMH) FROM PINAKES  WHERE TYPOS=2 AND PERIGRAFH LIKE '%" + monada + "%'")
                          
                            SJ = SJ + Chr(13) + ", ""measurementUnit"":""" + Trim(Unitmy) + """"

726                         SJ = SJ + Chr(13) + ",""lineComments"":""---""   "       '       "lineComments": "string",

728

730                         If Len(monada) < 2 Then
732                             MsgBox ("9382.ΔΕΝ ΕΧΩ MON.MET  ΣΕΙΡΑ " + str(NEGG) + "--" + EGGTIM("ONOMA"))
734                             ToJasonSub = 0

                                Exit Function

                            End If

736                         Dim NCC     As Integer: NCC = GGET_NVALUE("select COUNT(*) FROM PINAKES  WHERE TYPOS=2 AND PERIGRAFH LIKE '%" + monada + "%'")

                            Dim UNITUBL As String

738                         If NCC = 0 Then
740                             UNITUBL = "H87"
                            Else
742                             UNITUBL = GGET_CVALUE("select ISNULL(C2,'') AS UBL FROM PINAKES  WHERE TYPOS=2 AND PERIGRAFH LIKE '%" + monada + "%'")

                            End If

744                         If Len(UNITUBL) = 0 Then
746                             MsgBox ("9389.ΔΕΝ ΕΧΩ MON.MET  UBL ΣΕΙΡΑ ? " + str(NEGG) + "--" + EGGTIM("ONOMA"))
748                             ToJasonSub = 0

                                Exit Function

                            End If

750                         SJ = SJ + Chr(13) + ", ""measurementUnitUbl"":""" + Trim(UNITUBL) + """"

                            '  newitem(NEGG).measurementUnitUbl = Trim(UNITUBL) ' "H87"

                            'newitem(NEGG).lineComments = ""  ' "S?????"
752                         SJ = SJ + Chr(13) + ", ""lineComments"":""" + "" + """"

754                         SJ = SJ + Chr(13) + ",""totalValue"":" + Replace(Format(Round(AJ + fpaRow, 2), "#####0.00"), ",", ".")

756                         SJ = SJ + Chr(13) + ", ""cpvCode"":""" + EGGTIM("CPV") + """"

                            '-------------------------------------------------------------------------------
758                         If Len(EGGTIM("CPV")) = 0 Then
760                             If is_B2G = 1 Then
762                                 MsgBox ("9405.ΔΕΝ ΕΧΩ CPV ΣΤΟ  " + EGGTIM("ONOMA"))
764                                 ToJasonSub = 0

                                    Exit Function

                                End If

                            End If

766                         If fpaRow = 0 And isDiakin <> 2 Then  ' στο δα δεν θελει απαλλαγες
                                ' If InStr("1.2  1.3  2.3  2.4", mTypPar) > 0 Then 'ENDOK - TRITON XORON
                                '   ' Set elem2Field = docStock.createElement("vatExemptionCategory"):'elem2Field.Text = "14": ' ' elemField.appendChild elem2Field
                                ' Else
                               
768                             Dim mAPAL As String: mAPAL = ""

770                             If SQLDT("APALAGIFPA") > 0 Then ' αν εχω ολικη απαλλαγη φπα να μην ασχολουμαι με τισ σειρες
772                                 mAPAL = str(SQLDT("APALAGIFPA"))
                                Else

774                                 If EGGTIM("apal") = 0 Then mAPAL = str(SQLDT("APALAGIFPA")) Else mAPAL = str(EGGTIM("apal"))
                                End If

776                             If Val(mAPAL) = 99 Then
                                Else
778                                 ' Set elem2Field = docStock.createElement("vatExemptionCategory"):'elem2Field.Text = mAPAL: ' ' elemField.appendChild elem2Field
                                    
                                    'newitem(NEGG).vatExemption = mAPAL ' EGGTIM.Rows(NEGG)("CPV")  ' "45233222-1"
                                     
780                                 SJ = SJ + Chr(13) + ",""vatExemption"":""" + Trim(mAPAL) + """"
                                     
782                                 Dim APALUBL As String: APALUBL = GGET_CVALUE("Select ISNULL(C2,'') AS UBL FROM PINAKES  WHERE TYPOS=44 AND AYJON=" + mAPAL)

784                                 SJ = SJ + Chr(13) + ",""vatExemptionUbl"":""" + Trim(APALUBL) + """"
                                   
786                                 If Len(APALUBL) = 0 Then
788                                     MsgBox ("δεν έχω κωδικό απαλλαγής ΦΠΑ UBL σειρά " + str(NEGG))
790                                     ToJasonSub = 0

                                        Exit Function

                                    End If
                                  
                                End If
                            End If
                           
                            '------------- ΑΥΤΟΤΙΜΟΛΟΓΗΣΗ --------------------<selfPricing>true</selfPricing>
792                         If IsAytotim = 1 Then
                                '------------------------ expenses CLASSIFICATION ------------------------
794                             ' Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχη
796                             ' Set elem3Field = docStock.createElement("ecls:classificationType"):'elem3Field.Text = Split(cTyposExod, ";")(0): ' elem2Field.appendChild elem3Field

798                             If fMydataFromEID1 <> 1 Then ' κατηγορία από το παραστατικό

800                                 ' Set elem3Field = docStock.createElement("ecls:classificationCategory"):'elem3Field.Text = Split(cTyposExod, ";")(1): ' elem2Field.appendChild elem3Field
                                Else
802                                 ' Set elem3Field = docStock.createElement("ecls:classificationCategory"):'elem3Field.Text = fKatEXod(EGGTIM("kathgoria")): ' elem2Field.appendChild elem3Field
                                    ' fKatEsod(EGGTIM("kathgoria"))
                                End If

804                             ' Set elem3Field = docStock.createElement("ecls:amount"):'elem3Field.Text = chDec(Format(AJ, "######0.##")): ' elem2Field.appendChild elem3Field
806                             ' ' elemField.appendChild elem2Field
                                '' .appendChild elem2Field
                  
                                'fKatEsod(EGGTIM("kathgoria"))
808                             If Left(mTypPar, 3) = "3.1" Then ' TITLOS KTHSHS
                                Else
810                                 ' Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχη
812                                 ' Set elem3Field = docStock.createElement("ecls:classificationType"):'elem3Field.Text = "VAT_361": ' elem2Field.appendChild elem3Field
814                                 ' Set elem3Field = docStock.createElement("ecls:amount"):'elem3Field.Text = chDec(Format(AJ, "######0.##")): ' elem2Field.appendChild elem3Field
816                                 ' ' elemField.appendChild elem2Field
                                End If

                            Else
                                '------------------------ INCOME CLASSIFICATION ------------------------
818                             ' Set elem2Field = docStock.createElement("incomeClassification") ' δημιουργω εσοχ' δημιουργω εσοχ
                                ' ' Set elem3Field = docStock.createElement("n1:classificationType"):'elem3Field.Text = Split(ctypos, ";")(1): ' elem2Field.appendChild elem3Field

820                             If fMydataFromEID1 <> 1 Then ' κατηγορία από το παραστατικό
822                                 If Len(Trim(Split(ctypos, ";")(1))) = 0 Or isDiakin = 2 Then   ' δεν εχει Ε3 π.χ. 1_95  ή ειναι δελτιο αποστολης
                                    Else '-------------------------------  ΠΑΙΡΝΕΙ ΑΠΟ ΠΑΡΑΣΤΑΤΤΙΚΟ
824                                     '' Set elem3Field = docStock.createElement("n1:classificationType"):'elem3Field.Text = Split(ctypos, ";")(1): ' elem2Field.appendChild elem3Field
                                        SJ = SJ + Chr(13) + ",""classificationType"":""" + Split(ctypos, ";")(1) + """"

                                    End If

826                                 ' Set elem3Field = docStock.createElement("n1:classificationCategory"):'elem3Field.Text = Split(ctypos, ";")(2): ' elem2Field.appendChild elem3Field

                                    '  newitem(NEGG).classificationCategory = Split(ctypos, ";")(2) ' "category1_1"
                                    '  newitem(NEGG).classificationType = Split(ctypos, ";")(1) '"E3_561_001"
828                                 SJ = SJ + Chr(13) + ",""classificationCategory"":""" + IIf(isDiakin = 2, "category3", Split(ctypos, ";")(2)) + """"
830

                                Else
                            
832                                 If EGGTIM("kathgoria") = 8 Or Len(Trim(Split(ctypos, ";")(1))) = 0 Or fKatEsod(EGGTIM("kathgoria")) = "category1_95" Then ' εγγυοδοσια Ή ΕΧΕΙ ΚΕΝΟ Ε3 Π.Χ. 1_95 CATEGORY
                                    Else ' 3/4/23 ---------------- ΠΑΙΡΝΕΙ ΤΟ Ε3  ΑΠΟ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ

                                        Dim E3type As String

834                                     If F_E3_APOKATHG_EID = 1 Then
836                                         E3type = FkatE3Esod(EGGTIM("kathgoria"))
                                        Else
838                                         E3type = Split(ctypos, ";")(1)
                                        End If
                                             
840                                     '' Set elem3Field = docStock.createElement("n1:classificationType"):'elem3Field.Text = E3type: ' elem2Field.appendChild elem3Field
                                        ' newitem(NEGG).classificationType = E3type ' Split(ctypos, ";")(1) '"E3_561_001"
842                                     SJ = SJ + Chr(13) + ",""classificationType"":""" + E3type + """"
                                    End If

                                    ' κατηγορια απο το "Κατηγορία έιδους"
844                                 '' Set elem3Field = docStock.createElement("n1:classificationCategory"):'elem3Field.Text = fKatEsod(EGGTIM("kathgoria")): ' elem2Field.appendChild elem3Field

                                    'newitem(NEGG).classificationCategory = fKatEsod(EGGTIM.Rows(NEGG)("kathgoria"))
846                                 SJ = SJ + Chr(13) + ",""classificationCategory"":""" + IIf(isDiakin = 2, "category3", fKatEsod(EGGTIM("kathgoria"))) + """"

                                End If

848                             ' Set elem3Field = docStock.createElement("n1:amount"):'elem3Field.Text = chDec(Format(AJ, "######0.##")): ' elem2Field.appendChild elem3Field
850                             ' ' elemField.appendChild elem2Field
                            End If
                      
852                        ' .appendChild elemField
854                         SJ = SJ + Chr(13) + "}" ' εδω κλεινει η σειρα=============================================
856                         EGGTIM.MoveNext
                            'Next
                        Loop

858                     SJ = SJ + Chr(13) + "]" 'ληγουν oi seires=======================================================================
                        '====================================================================================================================
      
860                     Gdb.Execute "UPDATE TIM SET AADEKAU=" + Replace(Format(SYN_KAU, "######0.#####"), ",", ".") + ",AADEFPA=" + Replace(Format(SYN_FPA, "######0.#####"), ",", ".") + " WHERE ID_NUM=" + str(SQLDT("ID_NUM"))
                       
                        ' AN DEN ΕΣΤΕΙΛΕ ΣΤΟ DETAILS ( SYN_KAU ) ΟΛΟ ΤΟ ΠΟΣΟ ΤΟΥ ΤΙΜΟΛΟΓΙΟΥ (SUMNET)
862                     If Abs(Math.Round(SUMNET, 2) - Math.Round(SYN_KAU, 2)) >= 0.01 Then
                           
864                         Gdb.Execute "UPDATE TIM SET ENTITYMARK='ERROR-ΠΡΙΝ ΑΠΟΣΤ', ENTITYUID='" + mProbl + "' where ID_NUM=" + str(SQLDT("ID_NUM"))
                             
                            ' den εστειλε ολο το τιμολογιο
866                         If noask = 0 Then
868                             MsgBox "456.ΑΣΥΜΦΩΝΙΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + SQLDT("ATIM") + " " + str(SUMNET) + "<>" + str(SYN_KAU)
                            End If

870                         ToJason = 0

                            Exit Function

                        End If

                        '  Exit Function
  
                        Dim TYPOS_FOROY  As Integer  '1=PARAKRA WITHHELD  2=TELH FEES  3=LOIPOI OTHER 4=XARTOS STAMP (5=KRATHSEIS(DEN TO YPOSTHRIZO 27-10-2021)

                        Dim TYPOS_2FOROY As Integer  '1=PARAKRA WITHHELD  2=TELH FEES  3=LOIPOI OTHER 4=XARTOS STAMP (5=KRATHSEIS(DEN TO YPOSTHRIZO 27-10-2021)
  
872                     TYPOS_FOROY = 0
874                     TYPOS_2FOROY = 0

                        Dim CPOSO_F        As String

                        Dim SYN_KRATIS(10) As Double, K_KR As Integer

876                     For K_KR = 1 To 5: SYN_KRATIS(K_KR) = 0: Next

                        'taxCategoryUbl  ΠΑΡΑΚΡΑΤΟΥΜΕΝΟΙ ΦΟΡΟΙ   104
                        '1 - Περιπτ. β’- Τόκοι - 15% 104 - Standard
                        '2 - Περιπτ. γ’ - Δικαιώματα - 20%   104 - Standard
                        '3 - Περιπτ. δ’ - Αμοιβές Συμβουλών Διοίκησης - 20%  104 - Standard
                        '4 - Περιπτ. δ’ - Τεχνικά Έργα - 3%  104 - Standard
                        '5 - Υγρά καύσιμα και προϊόντα καπνοβιομηχανίας 1%   104 - Standard
                        '6 - Λοιπά Αγαθά 4%  104 - Standard
                        '7 - Παροχή Υπηρεσιών 8% 104 - Standard
                        '8 - Προκ/τέος Φόρος Αρχ/νων & Μηχ/κών & Συμβατικών Αμοιβών, για Εκπόνηση Μελετών & Σχεδίων 4%   104 - Standard
                        '9 - Προκ/τέος Φόρος Αρχιτ/νων & Μηχ/κών & Συμβατικών Αμοιβών, αφορούν άλλης φύσης έργα 10%  104 - Standard
                        '10 - Προκ/τέος Φόρος στις Αμοιβές Δικηγόρων 15% 104 - Standard
                        '11 - Παρακράτηση Φ.Μ.Υ. παρ. 1 αρ. 15 ν. 4172/2013  104 - Standard
                        '12 - Παρακράτηση Φ.Μ.Υ. παρ. 2 αρ. 15 ν. 4172/2013 - Αξιωματικών Εμπορικού Ναυτικού 104 - Standard
                        '13 - Παρακράτηση Φ.Μ.Υ. παρ. 2 αρ. 15 ν. 4172/2013 -Κατώτερο Πλήρωμα Εμπορικού Ναυτικού 104 - Standard
                        '14 - Παρακράτηση Ειδικής Εισφοράς Αλληλεγγύης   104 - Standard
                        '15 - Παρακράτηση Φόρου Αποζημίωσης λόγω Διακοπής Σχέσης Εργασίας παρ. 3 αρ. 15 ν. 4172/2013 104 - Standard
                        '16 - Παρακρατήσεις συναλλαγών αλλοδαπής βάσει συμβάσεων αποφυγής διπλής φορολογίας (Σ.Α.Δ.Φ.)   104 - Standard
                        '17 - Λοιπές Παρακρατήσεις Φόρου 104 - Standard
                        '18 - Παρακράτηση Φόρου Μερίσματα περ.α παρ. 1 αρ. 64 ν. 4172/2013   104 - Standard

                        'taxCategoryUbl  ΤΕΛΗ    AAA,AEV,ABK
                        '1 - Για μηνιαίο λογαριασμό μέχρι και 50 ευρώ 12%    AAA - Telecommunication
                        '2 - Για μηνιαίο λογαριασμό από 50,01 μέχρι και 100 ευρώ 15% AAA - Telecommunication
                        '3 - Για μηνιαίο λογαριασμό από 100,01 μέχρι και 150 ευρώ 18%    AAA - Telecommunication
                        '4 - Για μηνιαίο λογαριασμό από 150,01 ευρώ και άνω 20%  AAA - Telecommunication
                        '5 - Τέλος καρτοκινητής επί της αξίας του χρόνου ομιλίας (12%)   AAA - Telecommunication
                        '6 - Τέλος στη συνδρομητική τηλεόραση 10%    AAA - Telecommunication
                        '7 - Τέλος συνδρομητών σταθερής τηλεφωνίας 5%    AAA - Telecommunication
                        '8 - Περιβαλλοντικό Τέλος & πλαστικής σακούλας ν. 2339/2001 αρ. 6α 0,07 ευρώ ανά τεμάχιο AEV - Environmental protection service
                        '9 - Εισφορά δακοκτονίας 2%  ACF - Miscellaneous treatment
                        '10 - Λοιπά τέλη ABK - Miscellaneous
                        '11 - Τέλη Λοιπών Φόρων  ABK - Miscellaneous
                        '12 - Εισφορά δακοκτονίας    ACF - Miscellaneous treatment
                        '13 - Για μηνιαίο λογαριασμό κάθε σύνδεσης (10%) AAA - Telecommunication
                        '14 - Τέλος καρτοκινητής επί της αξίας του χρόνου ομιλίας (10%)  AAA - Telecommunication
                        '15 - Τέλος κινητής και καρτοκινητής για φυσικά πρόσωπα ηλικίας 15 έως και 29 ετών (0%)  AAA - Telecommunication
                        '16 - Εισφορά προστασίας περιβάλλοντος πλαστικών προϊόντων 0,04 λεπτά ανά τεμάχιο [άρθρο 4 ν. 4736/2020] AEV - Environmental protection service
                        '17 - Τέλος ανακύκλωσης 0,08 λεπτά ανά τεμάχιο [άρθρο 80 ν. 4819/2021]   AEV - Environmental protection service
                        '18 - Τέλος διαμονής παρεπιδημούντων ABK - Miscellaneous
                        '19 - Tέλος επί των ακαθάριστων εσόδων των εστιατορίων και συναφών καταστημάτων  ABK - Miscellaneous
                        '20 - Τέλος επί των ακαθάριστων εσόδων των κέντρων διασκέδασης   ABK - Miscellaneous
                        '21 - Τέλος επί των ακαθάριστων εσόδων των καζίνο    ABK - Miscellaneous
                        '22 - Λοιπά τέλη επί των ακαθάριστων εσόδων  ABK - Miscellaneous

                        'taxCategoryUbl  (ΛΟΙΠΟΙ ΦΟΡΟΙ)  CAP,AEF
                        '1 - α1) ασφάλιστρα κλάδου πυρός 20% CAP - Insurance brokerage service
                        '2 - α2) ασφάλιστρα κλάδου πυρός 20% CAP - Insurance brokerage service
                        '3 - β) ασφάλιστρα κλάδου ζωής 4%    CAP - Insurance brokerage service
                        '4 - γ) ασφάλιστρα λοιπών κλάδων 15% CAP - Insurance brokerage service
                        '5 - δ) απαλλασσόμενα φόρου ασφαλίστρων 0%   CAP - Insurance brokerage service
                        '6 - Ξενοδοχεία 1-2 αστέρων 0,50 €   AEF - Rents and leases
                        '7 - Ξενοδοχεία 3 αστέρων 1,50 € AEF - Rents and leases
                        '8 - Ξενοδοχεία 4 αστέρων 3,00 € AEF - Rents and leases
                        '9 - Ξενοδοχεία 5 αστέρων 4,00 € AEF - Rents and leases
                        '10 - Ενοικιαζόμενα -επιπλωμένα δωμάτια -διαμερίσματα 0,50 € AEF - Rents and leases
                        '11 - 2.Ειδικός Φόρος στις διαφημίσεις που προβάλλονται από την τηλεόραση (ΕΦΤΔ) 5%  AA - Advertising
                        '12 - 3.1 Φόρος πολυτελείας επί της φορολ. αξίας για ενδοκ/κώς αποκτούμενα & εισαγόμενα από τρίτες χώρες ABK - Miscellaneous
                        '13 - 3.2 Φόρος πολυτελείας 10% επί της τιμής πώλησης προ Φ.Π.Α. για τα εγχωρίως παραγόμενα είδη 10% ABK - Miscellaneous
                        '14 - Δικαίωμα του Δημοσίου στα εισητήρια των καζίνο (80% επί του εισητηρίου)    ABK - Miscellaneous
                        '15 - ασφάλιστρα κλάδου πυρός 20%    CAP - Insurance brokerage service
                        '16 - Λοιποί Τελωνειακοί Δασμοί-Φόροι    ABK - Miscellaneous
                        '17 - Λοιποί Φόροι   ABK - Miscellaneous
                        '18 - Επιβαρύνσεις Λοιπών Φόρων  ABK - Miscellaneous
                        '19 - ΕΦΚ    ABK - Miscellaneous
                        '20 - Ξενοδοχεία 1-2 αστέρων 1,50€ (ανά Δωμ./Διαμ.)  AEF - Rents and leases
                        '21 - Ξενοδοχεία 3 αστέρων 3,00€ (ανά Δωμ./Διαμ.)    AEF - Rents and leases
                        '22 - Ξενοδοχεία 4 αστέρων 7,00€ (ανά Δωμ./Διαμ.)    AEF - Rents and leases
                        '23 - Ξενοδοχεία 5 αστέρων 10,00€ (ανά Δωμ./Διαμ.)   AEF - Rents and leases
                        '24 - Ενοικιαζόμενα επιπλωμένα δωμάτια – διαμερίσματα 1,50€ (ανά Δωμ./Διαμ.) AEF - Rents and leases
                        '25 - Ακίνητα βραχυχρόνιας μίσθωσης 1,50€    AEF - Rents and leases
                        '26 - Ακίνητα βραχυχρόνιας μίσθωσης μονοκατοικίες άνω των 80 τ.μ. 10,00€ AEF - Rents and leases
                        '27 - Αυτοεξυπηρετούμενα καταλύματα – τουριστικές επιπλωμένες επαύλεις (βίλες) 10,00€    AEF - Rents and leases
                        '28 - Ακίνητα βραχυχρόνιας μίσθωσης 0,50€    AEF - Rents and leases
                        '29 - Ακίνητα βραχυχρόνιας μίσθωσης μονοκατοικίες άνω των 80 τ.μ. 4,00€  AEF - Rents and leases
                        '30 - Αυτοεξυπηρετούμενα καταλύματα – τουριστικές επιπλωμένες επαύλεις (βίλες) 4,00€ AEF - Rents and leases

                        'TYPOSPARAKRAT exei 1 ή 2 φόρουυς / παρακρατησεις
                        'π.χ.  1;4;3;1   σημαινει οτι εχει παρακρατηση(1) και λοιπους φόρους (3)
                        '      1=parakrathsh   4=typos4 parakrathhshw   3=λοιποι φοροι   1=τυπος λοιπου φορου

                        '  '------------------------ taxes CLASSIFICATION ------------------------
                        'ISNULL(PARAKRATISI,0) AS PARAKRAT,ISNULL(B_C2,'') AS TYPOSPARAKRAT "    B_C2=>ΤΥPΟSPΑRΑΚRΑΤ   1;4;3;5(PARAKRATHSH+XARTOSHMO)  , PARAKRATISI=ΑΞΙΑ 1ΗΣ ΠΑΡΑΚΡΑΤΗΣΗΣ , KR1=>2Η ΑΞΙΑ Π.Χ. ΧΑΡΤΟΣΗΜΟΥ
                           
                        'φοροσ διαμονης
                        If Split(ctypos, ";")(0) = "8.2" Then

                        End If

878                     If SQLDT("PARAKRAT") > 0 Then

                            Dim RR() As String

880                         RR = Split(SQLDT("TYPOSPARAKRAT"), ";")
882                         TYPOS_FOROY = Val(RR(0))

884                         SYN_KRATIS(TYPOS_FOROY) = SQLDT("PARAKRAT")

886                         '' Set elemField = docStock.createElement("taxesTotals") ' δημιουργω εσοχη
888                         '' Set elem2Field = docStock.createElement("taxes") ' δημιουργω εσοχη
890                         '' Set elem3Field = docStock.createElement("taxType"):'elem3Field.Text = RR(0): ' elem2Field.appendChild elem3Field

                            'taxesTotals
892                         SJ = SJ + Chr(13) + ",""taxesTotals"":[ {"

                            Dim CATUBL As String: CATUBL = " "

'                            If RR(0) = 1 Then ' ΠΑΡΑΚΡΑΤΗΣΗ ΦΟΡΟΥ
'                                CATUBL = "104"
'                            ElseIf RR(0) = 4 Then
'                                CATUBL = "SAE"
'                            End If
                            
                            If RR(0) = 1 Then ' ΠΑΡΑΚΡΑΤΗΣΗ ΦΟΡΟΥ
                            
894                             If gVal(RR(1)) > 0 Then  ' KATHGORIA PARAKRATHSHS KAI POSOSTO PARAKRATHSHS
896                                 SJ = SJ + Chr(13) + "  ""taxCategory"":" + RR(1)
898                                 SJ = SJ + Chr(13) + ",""taxCategoryUbl"":""104"" "

                                    Dim taxPercent As Single

902                                 If RR(1) = 4 Then
904                                     taxPercent = 3  '??????????  ?a?f?t?
906                                 ElseIf RR(1) = 2 Or RR(1) = 3 Then
908                                     taxPercent = 20
910                                 ElseIf RR(1) = 6 Then
912                                     taxPercent = 4
914                                 ElseIf RR(1) = 7 Then
916                                     taxPercent = 8
918                                 ElseIf RR(1) = 8 Then
920                                     taxPercent = 4
922                                 ElseIf RR(1) = 9 Then
924                                     taxPercent = 10
926                                 ElseIf RR(1) = 10 Then
928                                     taxPercent = 15
930                                 ElseIf RR(1) = 5 Then 'ygra kays
932                                     taxPercent = 1
934                                 ElseIf RR(1) = 1 Then
936                                     taxPercent = 15
                                    End If

938                                 SJ = SJ + Chr(13) + ",""taxPercent"":" + str(taxPercent)
                                End If
                           
                            ElseIf RR(0) = 3 Then 'LOPOI FOROI
                            
                                If gVal(RR(1)) > 0 Then  ' KATHGORIA LOIPOI FOROI
                                    SJ = SJ + Chr(13) + "  ""taxCategory"":" + RR(1)
                                    SJ = SJ + Chr(13) + ",""taxType"":3 "
                                    SJ = SJ + Chr(13) + ",""taxCategoryUbl"":""AEF"" "
                                End If
                             ElseIf RR(0) = 4 Then 'psifiako telos
                            
                               If gVal(RR(1)) > 0 Then  '
                                    SJ = SJ + Chr(13) + "  ""taxCategory"":" + RR(1)
                                    SJ = SJ + Chr(13) + ",""taxType"":4 "
                                    SJ = SJ + Chr(13) + ",""taxCategoryUbl"":""SAE"" "
                                End If
                            
                            
                            End If

940                         ' ' Set elem3Field = docStock.createElement("underlyingValue"):'elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): ' elem2Field.appendChild elem3Field
942                         ' ' Set elem3Field = docStock.createElement("taxAmount"):'elem3Field.Text = chDec(Format(Round(sqlDt("PARAKRAT"), 2), "####0.00")): ' elem2Field.appendChild elem3Field
944                         ' ' ' elemField.appendChild ELEM2FIELD  ' taxes
                            '' .appendChild elemField  '.appendChild elemField

                            ' newTax(1).UnderlyingValue = SYN_KAU
                            'newTax(1).taxAmount = sqlDt.rows(klm)("PARAKRAT")
946                         SJ = SJ + Chr(13) + " ,""UnderlyingValue"":" + Replace(Format(SYN_KAU, "#####0.00"), ",", ".")
948                         SJ = SJ + Chr(13) + ",""taxAmount"":" + Replace(Format(SQLDT("PARAKRAT"), "######0.00"), ",", ".")
950                         SJ = SJ + Chr(13) + "}" 'κλεινει το taxtotal  Replace(Format(Round(fpaRow, 2), "####0.00"), ",", ".")

                            ' εχει και 2η παρακρατηση
952                         If UBound(RR) > 2 Then
954                             TYPOS_2FOROY = Val(RR(2))
956                             SYN_KRATIS(TYPOS_2FOROY) = SYN_KRATIS(TYPOS_2FOROY) + SQLDT("KR1")
                                ' ' Set elemField = docStock.createElement("taxesTotals") ' δημιουργω εσοχη
958                             '   ' Set elem2Field = docStock.createElement("taxes") ' δημιουργω εσοχη
960                             '  ' Set elem3Field = docStock.createElement("taxType"):'elem3Field.Text = RR(2): ' elem2Field.appendChild elem3Field

962                             If gVal(RR(1)) > 0 Then
964                                 ' ' Set elem3Field = docStock.createElement("taxCategory"):'elem3Field.Text = RR(3): ' elem2Field.appendChild elem3Field
                                End If

966                             '  ' Set elem3Field = docStock.createElement("underlyingValue"):'elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): ' elem2Field.appendChild elem3Field
968                             '  ' Set elem3Field = docStock.createElement("taxAmount"):'elem3Field.Text = chDec(Format(Round(sqlDt("KR1"), 2), "####0.00")): ' elem2Field.appendChild elem3Field
970                             ' ' ' elemField.appendChild ELEM2FIELD  ' taxes
                            End If
                           
972                         '.appendChild elemField  '.appendChild elemField
                            
974                         SJ = SJ + Chr(13) + " ] " ' kleinei to taxestotals
 
                        End If
      
                        If isDiakin <> 2 Then  ' στο δα δεν θελει PAYMENT
      
594                         SJ = SJ + Chr(13) + ",""paymentMethods"":[{ ""type"":" + cTrp ' """"
596                         SJ = SJ + Chr(13) + ",""amount"":" + chDec(Format(Round(SUMNET + SumFpa, 2), "####0.00"))
598

                            If IS_POS = 1 Then
                          
                                '                                    Sj = Sj + Chr(13) + ",""terminalId"":""" + TID + """"
                                '                                    Sj = Sj + Chr(13) + ",""nspCode"":2"
                                '                                    Sj = Sj + Chr(13) + ",""tipAmount"":0"
                                '                                    Sj = Sj + Chr(13) + ",""signature"":""" + F_PAROX_SIGNATURE + """"
                                '                                    Sj = Sj + Chr(13) + ",""transactionId"":""" + POSTransactionId + """"
                          
                            End If
                           
600                         SJ = SJ + Chr(13) + "}] " 'telos paymentMethods
      
                        End If
      
                        '------------------------------------------------ InvoiceSummary
                        
976                     '  ' Set elemField = docStock.createElement("invoiceSummary") ' δημιουργω εσοχ
978                     ' ' Set elem2Field = docStock.createElement("totalNetValue"):'elem2Field.Text = chDec(Format(Round(SYN_KAU, 2), "######0.##")): ' ' elemField.appendChild ELEM2FIELD
980                     ' ' Set elem2Field = docStock.createElement("totalVatAmount"):'elem2Field.Text = chDec(Format(Round(SYN_FPA, 2), "######0.##")): ' ' elemField.appendChild ELEM2FIELD

982                     SJ = SJ + Chr(13) + " , ""invoiceSummary"":{ "

                        Dim synoloxorisFPA As Double

                        If SQLDT("PARAKRAT") > 0 And SYN_KAU = 0 Then
                            synoloxorisFPA = SQLDT("PARAKRAT")
                        Else
                            synoloxorisFPA = SYN_KAU
                        End If

984                     SJ = SJ + Chr(13) + " ""totalNetValue"":" + Replace(Format(synoloxorisFPA, "#####0.00"), ",", ".")

986                     SJ = SJ + Chr(13) + ",""totalVatAmount"":" + Replace(Format(Round(SYN_FPA, 2), "#####0.00"), ",", ".") ' Replace(str(SYN_FPA), ",", ".")
                          
                        Dim synn As Double

                        synn = SYN_FPA + SYN_KAU

                        If SQLDT("PARAKRAT") > 0 Then

                            Dim RR2() As String

                            RR2 = Split(SQLDT("TYPOSPARAKRAT"), ";")
                              
                            If RR2(0) = "1" Then ' parakratisi
                                synn = synn - SQLDT("PARAKRAT")
                            Else
                                synn = synn + SQLDT("PARAKRAT")
                            End If
                        End If

988                     SJ = SJ + Chr(13) + ",""totalValue"":" + Replace(Format(Round(synn, 2), "######0.00"), ",", ".")
990                     SJ = SJ + Chr(13) + " } " ' kleinei to Invoice sUMMARY
                              
992                     If Len(SQLDT("jwt")) > 10 Then
994                         SJ = SJ + Chr(13) + ", ""delayedProcessCode"":""" + SQLDT("jwt") + """ "
                        End If
                              
996                     SJ = SJ + Chr(13) + " }} " ' kleinei to Invoice

                        ' PARAKRATHSH
998                     If SYN_KRATIS(1) > 0 Then CPOSO_F = chDec(Format(Round(SYN_KRATIS(1), 2), "####0.00")) Else CPOSO_F = "0.00"
1000                    ' Set elem2Field = docStock.createElement("totalWithheldAmount"):'elem2Field.Text = CPOSO_F: ' ' elemField.appendChild elem2Field
                            
                        'totalFeesAmount
1002                    If SYN_KRATIS(2) > 0 Then CPOSO_F = chDec(Format(Round(SYN_KRATIS(2), 2), "####0.00")) Else CPOSO_F = "0.00"
1004                    ' Set elem2Field = docStock.createElement("totalFeesAmount"):'elem2Field.Text = CPOSO_F:: ' ' elemField.appendChild elem2Field

                        'totalStampDutyAmount
1006                    If SYN_KRATIS(4) > 0 Then CPOSO_F = chDec(Format(Round(SYN_KRATIS(4), 2), "####0.00")) Else CPOSO_F = "0.00"
1008                    ' Set elem2Field = docStock.createElement("totalStampDutyAmount"):'elem2Field.Text = CPOSO_F: ' ' elemField.appendChild elem2Field

                        'totalOtherTaxesAmount
1010                    If SYN_KRATIS(3) > 0 Then CPOSO_F = chDec(Format(Round(SYN_KRATIS(3), 2), "####0.00")) Else CPOSO_F = "0.00"
1012                    ' Set elem2Field = docStock.createElement("totalOtherTaxesAmount"):'elem2Field.Text = CPOSO_F: ' ' elemField.appendChild elem2Field
                            
1014                    ' Set elem2Field = docStock.createElement("totalDeductionsAmount"):'elem2Field.Text = "0.00": ' ' elemField.appendChild elem2Field
                            
                        '  If SYN_KRATIS(1) > 0 Then
                        '      ' Set elem2Field = docStock.createElement("totalGrossValue"):'elem2Field.Text = chDec(Format(SYN_KAU + SYN_FPA - SYN_KRATIS(1), "######0.##")): ' ' elemField.appendChild elem2Field
                        ' End If
1016                    If SYN_KRATIS(1) + SYN_KRATIS(2) + SYN_KRATIS(3) + SYN_KRATIS(4) + SYN_KRATIS(5) > 0 Then
1018                        ' Set elem2Field = docStock.createElement("totalGrossValue"):'elem2Field.Text = chDec(Format(Round(SYN_KAU, 2) + Round(SYN_FPA, 2) - Round(SYN_KRATIS(1), 2) + Round(SYN_KRATIS(2), 2) + Round(SYN_KRATIS(3), 2) + Round(SYN_KRATIS(4), 2) + Round(SYN_KRATIS(5), 2), "######0.##")): ' ' elemField.appendChild elem2Field
                        Else
1020                        ' Set elem2Field = docStock.createElement("totalGrossValue"):'elem2Field.Text = chDec(Format(Round(SYN_KAU, 2) + Round(SYN_FPA, 2), "######0.##")): ' ' elemField.appendChild elem2Field
                        End If
                            
                        If isDiakin = 2 Then
                      
                            '  ' Set elem2Field = docStock.createElement("incomeClassification") ' δημιουργω εσοχ' δημιουργω εσοχ
                            ' ' Set elem3Field = docStock.createElement("n1:classificationCategory"):'elem3Field.Text = "category3": ' elem2Field.appendChild elem3Field
                            '  ' Set elem3Field = docStock.createElement("n1:amount"):'elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): ' elem2Field.appendChild elem3Field
                            '  ' ' elemField.appendChild elem2Field

                            'Sj = Sj + Chr(13) + ",""classificationCategory"":""" + IIf(isDiakin = 2, "category3", fKatEsod(EGGTIM("kathgoria"))) + """"""
                            
                        Else

1022                        If fMydataFromEID1 <> 1 Then '  ΑΠΟ ΠΑΡΑΣΤΑΤΙΚΟ    fKatEsod (EGGTIM("kathgoria"))
                  
1024                            If IsAytotim = 1 Then
                                    '------------------------ expenses CLASSIFICATION ------------------------
1026                                ' Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχη
1028                                ' Set elem3Field = docStock.createElement("ecls:classificationType"):'elem3Field.Text = Split(cTyposExod, ";")(0): ' elem2Field.appendChild elem3Field
1030                                ' Set elem3Field = docStock.createElement("ecls:classificationCategory"):'elem3Field.Text = Split(cTyposExod, ";")(1): ' elem2Field.appendChild elem3Field
1032                                ' Set elem3Field = docStock.createElement("ecls:amount"):'elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): ' elem2Field.appendChild elem3Field
1034                                ' ' elemField.appendChild elem2Field

1036                                If Left(mTypPar, 3) = "3.1" Then ' TITLOS KTHSHS
                                    Else
1038                                    ' Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχη
1040                                    ' Set elem3Field = docStock.createElement("ecls:classificationType"):'elem3Field.Text = "VAT_361": ' elem2Field.appendChild elem3Field
1042                                    ' Set elem3Field = docStock.createElement("ecls:amount"):'elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): ' elem2Field.appendChild elem3Field
1044                                    ' ' elemField.appendChild elem2Field
                                    End If

                                Else '------------------------ income CLASSIFICATION ------------------------
                           
1046                                ' Set elem2Field = docStock.createElement("incomeClassification") ' δημιουργω εσοχ' δημιουργω εσοχ

1048                                If Len(Trim(Split(ctypos, ";")(1))) = 0 Then  ' P.X 1_95 ΚΕΝΟ Ε3
                              
                                    Else
1050                                    ' Set elem3Field = docStock.createElement("n1:classificationType"):'elem3Field.Text = Split(ctypos, ";")(1): ' elem2Field.appendChild elem3Field
                                    End If
                              
1052                                ' Set elem3Field = docStock.createElement("n1:classificationCategory"):'elem3Field.Text = Split(ctypos, ";")(2): ' elem2Field.appendChild elem3Field
1054                                ' Set elem3Field = docStock.createElement("n1:amount"):'elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): ' elem2Field.appendChild elem3Field
1056                                ' ' elemField.appendChild elem2Field
                             
                                End If

                            Else  '-------------------------------------------   ΑΠΟ ΚΑΤΗΓΟΡΙΑ ΕΙΔΩΝ

                                Dim K90 As Integer

1058                            If IsAytotim <> 1 Then ' oxi aytotimologisi

1060                                For K90 = 1 To 10

1062                                    If SumEsodExod(K90) > 0 Then
1064                                        ' Set elem2Field = docStock.createElement("incomeClassification") ' δημιουργω εσοχ' δημιουργω εσοχ

1066                                        If fKatEsod(K90) = "category1_95" Then  ' εγγυοδοσια-GIA EJODA TRITVN
                                            Else

                                                ' 3/4/23  ΠΑΙΡΝΕΙ ΤΟ Ε3 ΑΠΟ ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ
                                                Dim E3TYPE2 As String

1068                                            If F_E3_APOKATHG_EID = 1 Then E3TYPE2 = FkatE3Esod(K90) Else E3TYPE2 = Split(ctypos, ";")(1)
1070                                            ' Set elem3Field = docStock.createElement("n1:classificationType"):'elem3Field.Text = E3TYPE2: ' elem2Field.appendChild elem3Field
                                            End If
                                
                                            ' ΠΑΙΡΝΕΙ ΤΗΝ ΚΑΤΗΓΟΡΙΑ ΑΠΟ ΤΗΝ ΚΑΤΗΓΟΡΙΑ ΕΙΔΟΥΣ
1072                                        ' Set elem3Field = docStock.createElement("n1:classificationCategory")
1074                                       'elem3Field.Text = fKatEsod(K90)
1076                                        ' elem2Field.appendChild elem3Field
                             
1078                                        ' Set elem3Field = docStock.createElement("n1:amount")
1080                                       'elem3Field.Text = chDec(Format(SumEsodExod(K90), "######0.##"))
1082                                        ' elem2Field.appendChild elem3Field
                             
1084                                        ' ' elemField.appendChild elem2Field
                                        End If

                                    Next

                                Else  ' -------------------------- AYTOTIMOLOGISI  -----
                      
1086                                For K90 = 1 To 10

1088                                    If SumEsodExod(K90) > 0 Then
1090                                        ' Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχ' δημιουργω εσοχ
1092                                        ' Set elem3Field = docStock.createElement("ecls:classificationType"):'elem3Field.Text = Split(cTyposExod, ";")(0): ' elem2Field.appendChild elem3Field
1094                                        ' Set elem3Field = docStock.createElement("ecls:classificationCategory")
1096                                       'elem3Field.Text = fKatEXod(K90)
1098                                        ' elem2Field.appendChild elem3Field
                             
1100                                        ' Set elem3Field = docStock.createElement("ecls:amount")
1102                                       'elem3Field.Text = chDec(Format(SumEsodExod(K90), "######0.##"))
1104                                        ' elem2Field.appendChild elem3Field
                             
1106                                        ' ' elemField.appendChild elem2Field
                                        End If

                                    Next
                          
1108                                If Left(mTypPar, 3) = "3.1" Then ' TITLOS KTHSHS
                                    Else
1110                                    ' Set elem2Field = docStock.createElement("expensesClassification") ' δημιουργω εσοχη
1112                                    ' Set elem3Field = docStock.createElement("ecls:classificationType"):'elem3Field.Text = "VAT_361": ' elem2Field.appendChild elem3Field
1114                                    ' Set elem3Field = docStock.createElement("ecls:amount"):'elem3Field.Text = chDec(Format(SYN_KAU, "######0.##")): ' elem2Field.appendChild elem3Field
1116                                    ' ' elemField.appendChild elem2Field
                                    End If
                                End If  'ayt
                    
                            End If ' mydatafromEID
                        End If 'isdiakin=2
                            
1118                   ' .appendChild elemField
                    
                    End With  '--------------------------------------

1120               ' .appendChild invoice
                
1122                EGGTIM.Close

                    ' End If  'pos=1
                    
                    '            QQ = "{""terminal_id"":""" + tid + """ " + ","
                    '   QQ = QQ + " ""paymentAmount"":" + Format(MFPA + MKAU, "#####0.00") + ","
                    '   QQ = QQ + " ""netValue"":" + Format(MKAU, "#####0.00") + ","
                    'QQ = QQ + " ""VatAmount"":" + Format(MFPA, "#####0.00") + ","
                    '  QQ = QQ + " ""totalValue"":" + Format(MFPA + MKAU, "#####0.00") + ","
                    ' QQ = QQ + " ""externalSystemId"":""" + MID_NUM + """ " + ","
                    '
                    ' QQ = QQ + " ""uid"":""" + posuid + """ " + ","
                    ' QQ = QQ + " ""TimeStamp"":""" + posTimestamp + """ " + ","
                    ' QQ = QQ + " ""signature"":""" + F_posSignature + """ " + "}"
                    '
                    ' offlineJSONpos = QQ
                    ' req2 = request2_pliromis(paris1, offlineJSONpos)

                    '========================================================================================
                
1124                'MATIM = FERE_TRANSACTIONID(mID_NUM) Dim F_PAROX_SIGNATURE As String

                    SendToParochos SJ, SQLDT("id_num"), F_PAROX_SIGNATURE, TID, MFPA + MKAU, POSTransactionId, IS_POS, ubl, UPLOAD
                  
1126                SQLDT.MoveNext
                Loop

                ' Next
            End With ' /////////////////////////////////////////////////////////////////////////////////////////////

            'Exit Function

            '570         Set .documentElement = elemRoot
            '
            '            On Error Resume Next
            '
            '572         Kill "C:\" + F_TXTFILES + "\inv.xml"
            '
            '            '<<<<<<< HEAD
            '            On Error GoTo Tojason_Err
            '
            '            ' MAKE_request docStock
            '            'DOYLEYEI OK ALLA ΥΠΟΨΙΑΖΟΜΑΙ ΟΤΙ ΣΕ ΜΕΓΑΛΑ ΑΡΧΕΙΑ ΤΑ ΜΠΕΡΔΕΥΕΙ
            '574         FormatXmlDocument docStock ' βαζει κενα να ειναι ευκολο στο διαβασμα
            '            Dim ccc As String: ccc = "C:\" + F_TXTFILES + "\inv.xml"
            '576        ' .save ccc
        End With

        ' List1.AddItem "Δημιουργήθηκαν " + " απεσταλη"
     
        ' MAKE_request docStock.XML, 0, noask, "0"
        ' se paroxo
       
        'stelno to pa?astaT??? st?? pa????  // ResultJason
                    
        'On Error GoTo 0
       
        '        If Len(where.Caption) > 0 Then  ' ONLINE ΔΕΝ ΕΧΡΕΙΑΖΕΤΑΙ GRID
        '
        '        Else
        '
        '            ' ΠΑΩ ΝΑ ΠΙΑΣΩ ΤΑ ΤΙΜΟΛΟΓΙΑ ΠΟΥ ΤΥΠΩΘΗΚΑΝ ΚΑΙ ΑΚΥΡΩΘΗΚΑΝ ΧΩΡΙΣ ΝΑ ΠΑΡΟΥΝ MYDATA
        '            ' ΗΔΗ ΕΧΟΥΝ ΠΑΡΕΙ ΜΑΡΚ ΤΑ ΤΙΜΟΛΟΓΙΑ ΑΛΛΑ ΔΕΝ ΠΗΡΑΝ ΜΑΡΚ ΤΑ ΑΚΥΡΩΤΙΚΑ
        '            Dim sqlDtemp2 As New ADODB.Recordset
        '            sqlDtemp2.Open sql, Gdb, adOpenDynamic, adLockOptimistic
        '
        '            Do While Not sqlDtemp2.EOF
        '
        '                ' AN EINAI AKYRVMENO KAI EXEI MARK NA STEILEI TO AKYRVTIKO
        '                If Val(Left(CNull(sqlDtemp2!entitymark) + " ", 1)) > 1 Or Left(CNull(sqlDtemp2!entitymark) + " ", 1) = "Δ" Then
        '                    If sqlDtemp2!AKYROMENO = 1 Then
        '
        '                        CCCL = CANCEL_INVOICE(IIf(Left(CNull(sqlDtemp2!entitymark) + " ", 1) = "Δ", mID(sqlDtemp2!entitymark, 4, 15), sqlDtemp2!entitymark), sqlDtemp2!ATIM)
        '
        '                        If Len(CNull(CCCL)) > 1 Then
        '                            Gdb.Execute "UPDATE TIM SET INCMARK='" + CCCL + "' WHERE ID_NUM=" + str(sqlDtemp2!id_num)
        '                        End If
        '                    End If
        '                End If
        '
        '                sqlDtemp2.MoveNext
        '            Loop
        '
        '        End If
        '
        '<EhFooter>
        Exit Function

Tojason_Err:

        If noask = 0 Then
            MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Tojason " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        End If

        Resume Next

        '</EhFooter>

End Function





Sub SendToParochos(SJ As String, id_num, posSignature As String, TID As String, POS_AMOUNT As Single, POSTransactionId As String, IS_POS As Integer, ubl As String, ByVal UPLOAD As Integer)

        '==========================================================================================
        On Error Resume Next

        Dim mqq As String, fnamejson As String

100     fnamejson = "vb6jsonInstantReply" + str(id_num) + str(KN) + ".txt"
102     mqq = "c:\txtfiles\" + fnamejson
104     Open mqq For Output As #2
106     Print #2, SJ
108     Close #2
 
        '--------------------------- αποθηκεύω το ΑΡΧΕΙΟ  inv ΣΕ fakelo "c:\"+F_TXTFILES+"\sendinv\yyyy-mm-dd  ------------------------------
110     If Len(Dir("c:\" + F_TXTFILES + "", vbDirectory)) = 0 Then
112         MkDir "c:\" + F_TXTFILES + ""
        End If

114     If Len(Dir("c:\" + F_TXTFILES + "\sendinv", vbDirectory)) = 0 Then
116         MkDir "c:\" + F_TXTFILES + "\sendinv"
        End If

118     If Len(Dir("c:\" + F_TXTFILES + "\inv.xml", vbNormal)) = 0 Then
120         MsgBox "δεν υπάρχει το αρχείο των τιμολογιων inv.xml"

            Exit Sub

        End If

        Dim todaydir As String

122     todaydir = "c:\" + F_TXTFILES + "\sendinv\" + Format(Now, "yyyyMMdd")

124     If Len(Dir(todaydir, vbDirectory)) = 0 Then
126         MkDir todaydir
        End If
  
128     FileCopy mqq, todaydir + "\" + fnamejson + "-" + Format(Now, "yyyy-MM-dd-HH-mm")
        
        
       ' TO ANTIGRAFEI GIA NA BLEPV TO TELEYTAIO  JSON
        FileCopy mqq, "C:\" + F_TXTFILES + "\INV.JSN"
        '--------------ΑΠΟΣΤΟΛΗ ΣΕ ΠΑΡΟΧΟ ----------------------------------
130     Dim ResultJSON As String: ResultJSON = SJ 'ResultJSON ' " {""externa

        Dim REQ        As WinHttp.WinHttpRequest

132     Set REQ = New WinHttp.WinHttpRequest

        Dim URI As String:
                    
134     URI = gfURL1 + "api/send" '?Bearer=" + gf_Bearer==============================
                    
        Dim Q As String

        On Error GoTo lab70

        ' Dim JSON As String '= Req.ResponseText
        'Dim ser As JObject '= JObject.Parse(json)
        Dim cMARK              As String '= ser.SelectToken("signing")("mark").ToString

        Dim qr                 As String '= ser.SelectToken("signing")("qrCode").ToString

        Dim uid                As String '= ser.SelectToken("signing")("uid").ToString

        Dim authenticationCode As String '= ser.SelectToken("signing")("authenticationCode").ToString

136     With REQ
138         .Open "POST", URI, async:=False
140         .setRequestHeader "Content-Type", "application/hal+json"
142         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
144         .setRequestHeader "Authorization", "Bearer " + gf_Bearer
            'Note: Normally you don't include all of this whitespace, but
            'we'll use it in this example:
146        ' Q = " { ""Bearer"":""" + gf_Bearer + """, ""source"":" & ResultJSON & ", ""identifier"":""Parochos"",    ""externalSystemId"": """ + str(id_num) + """ }"
            ' q = " {   ""externalSystemId"": ""188"", ""source"":" & ResultJSON & "    }"
148       Q = " { ""Bearer"":""" + gf_Bearer + """, ""source"":" & ResultJSON & ", ""identifier"":""Parochos"",    ""externalSystemId"": """ + LTrim(str(id_num)) + """ }"
              
              .send (Q)
150         MILSEC 1000
        End With
        'Dim authenticationCode As String  ' : "AF40818DB06A43955C7C16DCA9649AB99FC16B14"
        Dim JSONC As String

        Dim TempFile As String
        TempFile = ""
        'Dim KN As Integer
152     For KN = 1 To 20

            'PAIRNO APANTISI APO TON PAROXO
154         URI = gfURL1 + "api/get"    ' \\\\\\\\\\\\\\\\\ get \\\\\\\\\\\\\\\\\\\\\\\\\

156         With REQ
158             .Open "POST", URI, async:=False
160             .setRequestHeader "Content-Type", "application/hal+json"
162             .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
164             .setRequestHeader "Authorization", "Bearer " + gf_Bearer
                'Note: Normally you don't include all of this whitespace, but
                'we'll use it in this example:
166             Q = " { ""Bearer"":""" + gf_Bearer + """, ""source"":" & ResultJSON & ",     ""externalSystemId"": """ + LTrim(str(id_num)) + """ }"
                ' q = " {   ""externalSystemId"": ""188"", ""source"":" & ResultJSON & "    }"
168             .send (Q)
                '  MsgBox("10231." + Req.ResponseText)

                'var beautified = parsedJson.ToString(Formatting.Indented);
                ' var minified = parsedJson.ToString(Formatting.None);
                ' MsgBox(Req.Status)
            End With

            '??af? t? json sto c:\txtfiles gia debugging
            '                                Dim oWrite4 As System.IO.StreamWriter
            '                                oWrite4 = file.CreateText("C:\txtfiles\jsonInstantReply" + ID_NUM + str(KN) + ".txt")
            '                                oWrite4.WriteLine (Req.responseText)
            '                                oWrite4.Close
            On Error Resume Next

            Dim ser As Object

170         JSONC = REQ.responseText
            ' FileCopy mqq, todaydir + "\" + fnamejson + "-" + Format(Now, "yyyy-MM-dd-HH-mm")
            TempFile = todaydir + "\" + "vb6jsonReply" + str(id_num) + str(KN) + ".txt"
172         Open TempFile For Output As #2
174         Print #2, JSONC
176         Close #2
                                
178         Set ser = JSON.parse(JSONC)
180         cMARK = ser.Item("signing")("mark")
182         qr = ser.Item("signing")("qrCode")
184         uid = ser.Item("signing")("uid")
            authenticationCode = ser.Item("signing")("authenticationCode") '": "AF40818DB06A43955C7C16DCA9649AB99FC16B14"
                                
186         GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET   JWT='" + Replace(gf_Bearer, "'", "`") + "'  WHERE ID_NUM=" + str(id_num)
                                
188         If Len(cMARK) > 10 Then


189             If F_demo = 1 Then cMARK = "'ΔΟΚ" + cMARK + "'"
'150                         RECS = GDBEXE("UPDATETIMWITHMARK", "UPDATE TIM SET ENTITY=0,  ENTITYUID='" + c4 + "',ENTITYMARK=" + c2 + " WHERE  ID_NUM=" + CCV)  ' ENTLINEN=" + str(fLong) + " and ENTITY=" + C1, RECS
                          
    
                GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET   AUTHENTICATIONCODE='" + authenticationCode + "'  WHERE ID_NUM=" + str(id_num)
190             GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET   ENTITYUID='" + uid + "',ENTITYMARK=" + cMARK + " WHERE ID_NUM=" + str(id_num)
192             GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET   QRURL='" + qr + "'  WHERE ID_NUM=" + str(id_num)
                                
                If Len(cMARK) > 10 Then  'ste??? t? submit payment
                         If IS_POS = 1 And ubl = "380" Then
                           Gdb.Execute "update TIM SET TRANSACTIONID='" + POSTransactionId + "',POSPAYSIGN='" + posSignature + "' WHERE ID_NUM=" + str(id_num)
                                    MsgBox ("OK")
                            Dim SSW As String:
                            SSW = POS_submitpayment(cMARK, LTrim(str(id_num)), posSignature, TID, Format(POS_AMOUNT, "###0.00"), POSTransactionId)
                            'SSW = POS_submitpayment(cMARK, str(id_num), posSignature, TID, Format(POS_AMOUNT, "###0.00"), POSTransactionId)
                         
                         End If
                         
                         'SSW = POS_submitpayment(cMARK, str(id_num), posSignature, TID, Format(POS_AMOUNT, "###0.00"), POSTransactionId)
                         
                         
                         
                End If '
                                
                                
                                
                                
                                
                                
                                
                                
                Exit For

            End If
                                
194         authenticationCode = ser.Item("signing")("authenticationCode")

196         Dim cMARK1 As String: cMARK1 = ser.Item("errorMessage")

198         If Len(cMARK1) > 2 Then
200          MsgBox ("Πρόβλημa (10103):" + cMARK1)

                'ToJasonSub = 0
                Exit Sub

            End If

202         If Len(cMARK) > 10 And Len(qr) > 10 Then ' ok

                Exit For

            Else

204             MILSEC 1000

            End If

        Next

        FileCopy TempFile, "C:\" + F_TXTFILES + "\apantINV.JSN"

'SSW = POS_submitpayment(cMARK, str(id_num), posSignature, TID, Format(POS_AMOUNT, "###0.00"), POSTransactionId)
                
'SSW = POS_submitpayment(cMARK, str(id_num), posSignature, tid, Format(POS_AMOUNT, "###0.00"), POSTransactionId)




        If Len(where.Caption) = 0 Or UPLOAD = 1 Then ' αν δεν ερχεται απο το par1 ή πρεπει να κανω upload τοτε στειλτα απο εδω
        
                        
660                Dim f_matim As String, F_FORMA1 As String
                      f_matim = GGET_CVALUE("SELECT ATIM from TIM where ID_NUM=" + str(id_num))
                      F_FORMA1 = GGET_CVALUE("SELECT FORMA1 FROM PARASTAT WHERE EIDOS='" & Left(f_matim, 1) & "'")
                     ' CrystalReport1.ReportFileName = F_FORMA1
                      Dim F_EKTYPOTHS2: F_EKTYPOTHS2 = FINDPARAMETROI(1, "PAR1", "F_EKTYPOTHS", "..", "Εκτυπωτής παραστατικών(με .. προεπιλεγμένος)") '
662
                         Dim mf_dat As Date
                         mf_dat = GGET_DVALUE("select HME FROM TIM WHERE ID_NUM=" + str(id_num))
                         
                         
                
                
664                 If F_PAROCHOS = 1 Then

                        'πρεπει να τυπωσειτο παραστατικο για να παρει το barcode,hmer,ariumo k.l.p
                         PAR1.TDBGrid1.Col = 2 ' ισα ισα να ανοιξει το παρ1
                         Dim cc3 As Double
                         cc3 = GGET_NVALUE("select AJI FROM TIM WHERE ID_NUM=" + str(id_num))
                         cc3c = Replace(Format(cc3, "######0.00"), ",", ".")
                         
                         
                          'DEN EXREIAZETAI TO KANONIZEI TO LASTTIMOL PAR1.Text1.Text = mID(f_matim, 2, 6)
                           PAR1.LASTTIMOL.RecordSource = "SELECT  HME,ATIM,AJI AS [AΞΙΑ],PEL.EPO AS [ΕΠΩΝΥΜΙΑ],PEL.KOD AS [ΚΩΔΙΚΟΣ],PARAT AS [ΣΧΟΛΙΑ],SXETIKO AS [ΣΧΕΤ.ΠΑΡ],ID_NUM,ENTITYMARK,SXETMARK FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where ID_NUM=" + str(id_num)
                           PAR1.LASTTIMOL.Refresh
                            PAR1.METASX_SE_TIM 99, cc3c, id_num, False, 0
                         
                         'PAR1.Text1.Text = mID(f_matim, 2, 6)
                        
                        Dim CPOSO As String
                        
                        
                          DUM = PAR1.printCrystal(f_matim, CDate(mf_dat))
                         
                         
                         Ucr9print.printingCR9 F_EKTYPOTHS, F_FORMA1 ' CrystalReport1.ReportFileName








666                     Dim fFILE As String: fFILE = Trim(f_matim) + ".PDF"  '"c:\mercvb\" +
                        Dim Size As Long
                        CrystalReport1.ReportFileName = F_FORMA1
668                     Size = gExport2ReportToPDF(F_FORMA1, Trim(f_matim), fFILE, F_EKTYPOTHS, CrystalReport1)
670                     Dim ubl2 As String: ubl2 = GGET_CVALUE("SELECT left(ISNULL(UBL,'000'),3) AS UBL2 FROM PARASTAT WHERE EIDOS='" & Left(f_matim, 1) & "'")
                       
                        
672                     If ubl2 <> "000" Then
674                         If GGET_NVALUE("select MYDATAISDEMO from MEM") = 1 Then
676                             PAROX_URLADD = "beta-"
                            Else
678                             PAROX_URLADD = ""
                            End If
                                                      
680                         PAR1.uploadAzure id_num, fFILE, Size
                        End If
                    End If

      End If
      Exit Sub

lab70:
206     MsgBox Err.Description

208     Resume Next

End Sub


'Sub Typose_Parast(ByVal cPoso As String, ByVal m_id As Long)
'     Dim m As Integer
'     Dim c As String
'        PAR1.METASX_SE_TIM 99, c, m_id, False, 0
'110     DoEvents
'120     PAR1.Command5_Click
'        PAR1.Text2(0).Text = "" ' για να μην διορθώση κατα λαθος
'End Sub




Function to_pos(ByVal TID As String, _
                ByVal mID_NUM As String, _
                ByVal MKAU As Double, _
                ByVal MFPA As Double, _
                ByVal SQLDThme As Date, _
                ByVal SQLDTATIM As String, _
                ByVal SQLDTBRANCH) As String ', POSTransactionId As String, posSignature As String, tid As String, issu As Issuer, he As InvoiceHeader, invs As InvoiceSummary) As String

    '=============================================================
'1. kalo ton paroxo gia na moy dosei ypografh posSignature
'2.pairno token apo to paris
'3.για να καλεσω για την πληρωμη θα χρειαστω τα
    'posSignature,postimestamp, posuid
    ' και ενα token που θα παρω απο το lambdapi
 




'----------- 11111111111111111111111111111111111

    Dim C_SEIRA As String

    Dim Mctypos

    C_SEIRA = GET_CVALUE("SELECT SYNT_TITL FROM PARASTAT WHERE EIDOS='" + Left$(SQLDTATIM, 1) + "'")
    Mctypos = FINDTYPOS(mID(SQLDTATIM, 1, 1))

    On Error GoTo to_posERROR

    Dim uripos As String: uripos = gfURL1 + "api/requestPayment"

    Dim ANSW   As String
 
    Dim QQ     As String

    QQ = "{""externalSystemId"":""" + mID_NUM + """ " + ","
    QQ = QQ + " ""issuerVatNumber"":""" + afmCompany + """" + ","
    QQ = QQ + " ""invoiceIssueDate"":""" + Format(SQLDThme, "yyyy-MM-dd") + ""","

    QQ = QQ + " ""companyBranch"":""" + SQLDTBRANCH + ""","
    QQ = QQ + " ""invoiceType"":""" + Split(Mctypos, ";")(0) + ""","
    QQ = QQ + " ""invoiceSeries"":""" + C_SEIRA + ""","
    QQ = QQ + " ""invoiceAA"":""" + mID(SQLDTATIM, 2, 6) + ""","
    QQ = QQ + " ""netValue"":" + Replace(Format(MKAU, "#####0.00"), ",", ".") + ","
    QQ = QQ + " ""VatAmount"":" + Replace(Format(MFPA, "#####0.00"), ",", ".") + ","
    QQ = QQ + " ""totalValue"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","

    QQ = QQ + " ""paymentAmount"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
    QQ = QQ + " ""NspCode"":""2"" , "
    QQ = QQ + " ""terminalId"":""" + TID + """ }"     '"99999069"

    Open "C:\TXTFILES\REQUES111TAPOPAROXO" + extID + ".TXT" For Output As #1
    Write #1, QQ
         
    Close #1

    ANSW = REQWINHTTP(QQ, uripos, LTrim(mID_NUM))

    Dim ser As Object
 
    Set ser = JSON.parse(ANSW)
        
    F_PAROX_SIGNATURE = ser.Item("paymentToken")("signature")
    posTimestamp = ser.Item("paymentToken")("timestamp")
    posuid = ser.Item("uid")
 
    Open "C:\TXTFILES\answREQUEST2222APOPAROXO" + extID + ".TXT" For Output As #1
    Write #1, ANSW
    Write #1, F_PAROX_SIGNATURE
    Close #1
 
    If Len(F_PAROX_SIGNATURE) < 5 Then
        to_pos = "error"

        Exit Function

    End If
 
    Dim timestampStart As String: timestampStart = """timestamp"":""" '-------------------------

    Dim timestampEnd   As String: timestampEnd = ""","

    Dim startIndex     As Integer: startIndex = InStr(posTimestamp, timestampStart)

    Dim endIndex       As Integer: endIndex = InStr(posTimestamp, timestampEnd)

    posTimestamp = mID(posTimestamp, 1, 33) ' startIndex + 1, endIndex - startIndex)
    '=================================22222222222222222222222222222222222===============================================================================================
    'για να καλεσω για την πληρωμη θα χρειαστω τα
    'posSignature,postimestamp, posuid
    ' και ενα token που θα παρω απο το lambdapi
 
    ' και ενα token που θα παρω απο το lambdapi
    Dim userpos As String, PWDPOS As String

    userpos = GGET_CVALUE("SELECT ISNULL(USERPOS,'') AS U FROM MEM")
    PWDPOS = GGET_CVALUE("SELECT ISNULL(PWDPOS,'') AS U FROM MEM")

    Dim paris1 As String: paris1 = request_pliromis(userpos, PWDPOS) ' username password prepei na mpoyn san parametroi

    '==================================333333333333333333333333===============================================
  
    Dim req2   As String
       
  
    QQ = "{""terminal_id"":""" + TID + """ " + ","
    QQ = QQ + " ""paymentAmount"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
    QQ = QQ + " ""netValue"":" + Replace(Format(MKAU, "#####0.00"), ",", ".") + ","
    QQ = QQ + " ""vatAmount"":" + Replace(Format(MFPA, "#####0.00"), ",", ".") + ","
    QQ = QQ + " ""totalValue"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
    QQ = QQ + " ""externalSystemId"":""" + mID_NUM + """ " + ","
  
    QQ = QQ + " ""uid"":""" + posuid + """ " + ","
    QQ = QQ + " ""TimeStamp"":""" + posTimestamp + """ " + ","
    QQ = QQ + " ""signature"":""" + F_PAROX_SIGNATURE + """ " + "}"
 
    offlineJSONpos = QQ
 
    Open "C:\TXTFILES\paris1qq333APOPAROXO" + extID + ".TXT" For Output As #1
    Write #1, QQ
         
    Close #1
    MsgBox "ΠΑΤΗΣΤΕ ΤΟ ΚΟΥΜΠΙ ΓΙΑ ΓΕΦΥΡΑ"
    req2 = request2_pliromis(paris1, offlineJSONpos)
 
    Open "C:\TXTFILES\requ2apopliromi5555POPAROXO" + extID + ".TXT" For Output As #1
    Write #1, req2
         
    Close #1
  
    Dim mRes As String: mRes = Split(req2, "@")(0)

    POSTransactionId = Split(req2, "@")(1)
    ' pay(0).transactionId = POSTransactionId  '

    If mRes <> "1" Then

        Dim HH As Integer

        For HH = 1 To 5
            MILSEC 2000 '5sec
            req2 = FERE_TRANSACTIONID(mID_NUM)
                 If req2 = "error" Then
                    
                    ToJasonSub = 0
                    ' CancelPosPayment mID_NUM, F_PAROX_SIGNATURE
                     CancelPosPayment mID_NUM, F_PAROX_SIGNATURE
                    ' CancelPosPayment mID_NUM, F_PAROX_SIGNATURE
                     Exit Function
                 End If
                 
                 
            If Len(req2) > 5 Then
                POSTransactionId = req2 '.toString.Split("@")(1)
                to_pos = POSTransactionId
                Exit For

            End If

        Next
            
        If Len(req2) < 5 Then
            
            Dim ANSPOS As Integer: ANSPOS = vbYes

            ANSPOS = MsgBox("Εγινε η συναλλαγή;", vbYesNo)
                           
            If ANSPOS = vbYes Then
                req2 = FERE_TRANSACTIONID(mID_NUM) ' 1;123......"

                If Len(req2) > 5 Then
                    POSTransactionId = req2 ' Split(req2, "@")(1)
                    to_pos = POSTransactionId
                Else
                    ToJasonSub = 0

                    Exit Function

                End If
                                 
            Else

                If gVal(mRes) = 2 Or gVal(mRes) = 3 Then ' ap????f???e, a???????e
                                    
                          ToJasonSub = 0
 CancelPosPayment mID_NUM, F_posSignature
                    Exit Function

                End If

                ToJasonSub = 0

                Exit Function

            End If
                            
        End If

    End If
 to_pos = POSTransactionId
    Exit Function

to_posERROR:
    to_pos = ""

    Resume Next

End Function
  Function request2_pliromis(token, offline) As String
 ' apo dotnet     {""terminal_id"":""01843721"",""paymentAmount"":2.0,""netValue"":1.61,
 '""vatAmount"":0.3899999999999999,""totalValue"":2.0,""externalSystemId"":""40479"",
 '""uid"":""5852D6190FA34B5AF56F3662E0F3FBB111A56322"",""timestamp"":""2026-02-16T08:03:11.2406733+02:00"",
 '""signature"":""304402200B53EE2FD874442FCC47F63D3E50B6FE5D80B4457E328ECF583C1C3B042959A3022059B4DADD27009AEFAA914CE81D2A4055B1E682457AA073CF6C0E3BC602FCF3FC""}"

  '  apo vb6 {"terminal_id":"01843721" , "paymentAmount":2.00, "netValue":1.61, "vatAmount":0.39, "totalValue":2.00, "externalSystemId":"40480" ,
  '"uid":"5DC5E9B978430BEDEC12A57143A13EB1501801BE" , "TimeStamp":"2026-02-16T08:13:41.8491877+02:00" ,
  '"signature":"3045022100B4C79C0BED0B71BF2628BA8A6C08B454F709172C787F7C4F4BEE4DCC1DB23C4D02207AA766F6D69E0B13767108F8F02E5B5DE9024BA9C0947C1C8E6002C0508E7C56" }
   
        request2_pliromis = "0"
        
        Dim uripos As String: uripos = "https://lambdapi.gr/backend/webecr/offline_payment"
        Dim Reqpos As WinHttp.WinHttpRequest
       Set Reqpos = New WinHttp.WinHttpRequest
       
       offline = Replace(offline, "TimeStamp", "timestamp")
       
        For KN = 1 To 3
            
            
            With Reqpos
                .Open "POST", uripos, async:=False
                .setRequestHeader "Content-Type", "application/json"
                .setRequestHeader "Accept", "*/*"
                .setRequestHeader "Authorization", "Bearer " + token
                Dim Q  As String: Q = offline
               On Error GoTo ADYNAMIA
                    .send (Q)
                

                MsgBox ("ΑΠΑΝΤΗΣΗ ΑΠΟ offline_payment" + Chr(13) + Reqpos.responseText)
            
            End With
            MILSEC (1000)
          



            If InStr("REQ1" + Chr(13) + Reqpos.responseText, "Result") > 0 Then
                ' MsgBox(Reqpos.ResponseText)
                Dim JSON2 As String: JSON2 = Reqpos.responseText
                 Dim ser As Object
                ser = JSON.parse(JSON2)

               


                Dim ctoken As String: ctoken = ser.Item("transaction")("Result")
                'Dim transid As String = ser.SelectToken("transaction")("Id").ToString
                Dim transid As String: transid = ser.SelectToken("transaction")("TransactionId")
                request2_pliromis = ctoken + "@" + transid '   Id '72012914','Reqpos.ResponseText
                Exit For
            End If
        Next
        Exit Function
        
ADYNAMIA:
                    MsgBox ("ΑΔΥΝΑΜΙΑ ΠΛΗΡΩΜΗΣ")
                    request2_pliromis = "error"
                    Exit Function
               
        
End Function

 Function request_pliromis(userpos As String, PWDPOS As String) As String


        Dim Reqpos As WinHttp.WinHttpRequest
       Set Reqpos = New WinHttp.WinHttpRequest
        For KN = 1 To 1
            'pa???? t?? apa?t?s? ap? t?? pa????
            Dim uripos As String: uripos = "https://lambdapi.gr/api/collections/users/auth-with-password"
            With Reqpos
                .Open "POST", uripos, async:=False
                .setRequestHeader "Content-Type", "application/json"
                .setRequestHeader "Accept", "*/*"
              
                Dim qpos  As String: qpos = " { ""identity"": """ + userpos + """,    ""password"": """ + PWDPOS + """ }"
                .send (qpos)
                ' MsgBox("token ap? lambdapi" + Chr(13) + Reqpos.ResponseText)
            End With
           ' Threading.Thread.Sleep (500)
           DoEvents
           
            '{"record":{"avatar":"","collectionId":"_pb_users_auth_","collectionName":"users","created":"2024-11-11 17:45:54.551Z",
            ''"email":"glagakis@gmail.com","emailVisibility":true,"id":"csxp3wqor6tc09m",
            ''"merchant":"4myj90okhn4v03p","name":"Ge?????? ?a?????","updated":"2024-11-11 17:45:54.551Z","username":"users99156","verified":true},
            ''"token":"eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjb2xsZWN0aW9uSWQiOiJfcGJfdXNlcnNfYXV0aF8iLCJleHAiOjE3MzI2NTYyNDUsImlkIjoiY3N4cDN3cW9yNnRjMDltIiwidHlwZSI6ImF1dGhSZWNvcmQifQ.F9qcn0U6uMyV6oaJgRHSIZyk5jsx0khhObMcJJI5QJk"}

            Dim JSON2 As String: JSON2 = Reqpos.responseText
            Dim ser As Object
            Set ser = JSON.parse(JSON2)
            Dim ctoken As String: ctoken = ser.Item("token")
            'qr = ser.SelectToken("signing")("qrCode").ToString
            'uid = ser.SelectToken("signing")("uid").ToString
            'authenticationCode = ser.SelectToken("signing")("authenticationCode").ToString
            request_pliromis = ctoken










        Next
    End Function



    Function POS_submitpayment(ByVal MARK As String, ByVal extID As String, ByVal signatureParochos As String, ByVal TID As String, ByVal AMOUNT As String, ByVal TransactionId As String) As String

      


On Error Resume Next
        Dim bearer As String: bearer = REQPAROCHOS()

 'bearer = gf_Bearer
 
' bearer = REQPAROCHOS()
' bearer = REQPAROCHOS()
 
'        Dim subm As New SubmitPayment
'        If mark.Substring(0, 3) = "ΔΟΚ" Then
'            subm.mark = mark.Substring(3, 15)
'        Else
'            subm.mark = mark
'
'        End If

'        subm.externalSystemId = extID
'        newpaymentInfos(1) = New paymentInfos
'
'        newpaymentInfos(1).amount = gVal(amount) '5.5
'        newpaymentInfos(1).terminalId = tid
'        newpaymentInfos(1).transactionId = transactionId   '"789260373"
'        newpaymentInfos(1).Type = 7
'        newpaymentInfos(1).signature = signatureParochos
'
'        subm.paymentInfo.ADD (newpaymentInfos(1))
'
'
'        Dim ResultJSONpos As String = JsonConvert.SerializeObject(subm).ToString











        Dim Reqpos As WinHttp.WinHttpRequest
       Set Reqpos = New WinHttp.WinHttpRequest



1200

        Dim SJ As String: SJ = ""
       ' bearer = REQPAROCHOS()
       ' AMOUNT = "1.5"
         SJ = SJ + Chr(13) + "   {  ""externalSystemId"":""" + Trim(LTrim(extID)) + ""","
        SJ = SJ + Chr(13) + "     ""mark"":" + MARK + ","
       
        SJ = SJ + Chr(13) + "     ""paymentInfo"":[{"
        SJ = SJ + Chr(13) + "     ""type"":7,"
        SJ = SJ + Chr(13) + "     ""amount"":" + Replace(AMOUNT, ",", ".") + ","
        'Sj = Sj + Chr(13) + "     ""paymentMethodInfo"":"""" ,"
        'Sj = Sj + Chr(13) + "     ""tipAmount"":0,"
        SJ = SJ + Chr(13) + "     ""transactionId"":""" + TransactionId + ""","
        SJ = SJ + Chr(13) + "     ""signature"":""" + signatureParochos + ""","
        
        SJ = SJ + Chr(13) + "     ""terminalId"":""" + TID + """}]}"
        
       ' Exit Function
        
        Open "C:\TXTFILES\SUBM" + extID + ".TXT" For Output As #1
          Write #1, SJ
        Close #1
        
MILSEC 1000

  Open "C:\TXTFILES\REQUESTAPOPAROXO" + extID + ".TXT" For Input As #1
          Input #1, SJ
          'Write #1, F_PAROX_SIGNATURE
  Close #1
 



       ' For KN = 1 To 1

            Dim uripos As String: uripos = gfURL1 + "api/submitPayment"
680:        With Reqpos
                .Open "POST", uripos, async:=False
                .setRequestHeader "Content-Type", "application/hal+json"
                .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
                .setRequestHeader "Authorization", "Bearer " + gf_Bearer
                .setRequestHeader "X-Version", "3"
                .send (SJ)
                 MsgBox (Reqpos.responseText)
            End With
            
         
            
           ' MsgBox Reqpos.responseText
            Open "C:\TXTFILES\SUBMANSW" + extID + ".TXT" For Output As #1
          Write #1, SJ
        Close #1
            
          ' GoTo 1200
            
            
            
            Dim posjson As String: posjson = Reqpos.responseText
           ' MsgBox ("submit:" + posjson)




            Dim posser As Object
            posser = JObject.parse(posjson)
            
            

            
           
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            '                        cMARK = ser.SelectToken("signing")("mark").ToString
            '                        qr = ser.SelectToken("signing")("qrCode").ToString
            '            posuid = posser.SelectToken("uid").ToString  '------------------


            '            Dim timestampStart As String = """timestamp"":"""  '-------------------------
            '            Dim timestampEnd As String = ""","

            '            Dim startIndex As Integer = posjson.IndexOf(timestampStart) + timestampStart.Length
            '            Dim endIndex As Integer = posjson.IndexOf(timestampEnd, startIndex)
            '740:        posTimestamp = posjson.Substring(startIndex, endIndex - startIndex)


            '            ' posTimestamp = posser.SelectToken("paymentToken")("timestamp").ToString("yyyy-MM-ddTHH:mm:ss.fffffffK")
            '            F_posSignature = posser.SelectToken("paymentToken")("signature").ToString  '------------------------

            Threading.Thread.Sleep (500)
      '  Next

POS_submitpayment = "ok"
    End Function



 ' Response Examp
    Sub CancelPosPayment(ByVal extID As String, ByVal signature As String)
        Dim Reqpos As WinHttp.WinHttpRequest
        Set Reqpos = New WinHttp.WinHttpRequest


        Dim SJ As String
        SJ = ""
        SJ = SJ + Chr(13) + "   {  ""externalSystemId"":""" + extID + ""","
        SJ = SJ + Chr(13) + "     ""signature"":""" + signature + """ } "


            'παιρνω την απαντηση απο τον παροχο
            Dim uripos As String: uripos = gfURL1 + "api/cancelPayment" '/api/cancelPayment
        With Reqpos
                .Open "POST", uripos, async:=False
                .setRequestHeader "Content-Type", "application/hal+json"
                .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
                .setRequestHeader "Authorization", "Bearer " + gf_Bearer
                .setRequestHeader "X-Version", "3"

               
                .send (SJ)

            End With

            Dim posjson As String: posjson = Reqpos.responseText
            MsgBox ("CancelPosPayment" + Chr(13) + posjson)
            
           '  MsgBox ("CancelPosPayment" + Chr(13) + posjson)
'
           ' Dim posser As Object: posser = JSON.parse(posjson)
  



    End Sub

