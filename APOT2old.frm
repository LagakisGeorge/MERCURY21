VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{CCA214C0-DFEB-4C91-9F0D-2665F77F6E23}#1.2#0"; "IDAutomationLinear.dll"
Begin VB.Form APOT2 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form2"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12165
   Icon            =   "APOT2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   12165
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command12 
      Caption         =   "Εττικέτα με περιγραφή"
      Height          =   405
      Left            =   10680
      TabIndex        =   74
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   3615
      TabIndex        =   73
      Top             =   1275
      Width           =   1785
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ΑΝΑΖΗΤΗΣΗ"
      Height          =   255
      Left            =   10680
      TabIndex        =   68
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   10680
      TabIndex        =   67
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CheckBox Scanner 
      Caption         =   "Από Scanner"
      Height          =   195
      Left            =   9135
      TabIndex        =   63
      ToolTipText     =   "Aπό αριθμό τιμ. 1234"
      Top             =   2160
      Width           =   1440
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Εττικέτα β"
      Height          =   225
      Left            =   9120
      TabIndex        =   61
      Top             =   1875
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Διαγραφή"
      Height          =   375
      Left            =   9375
      TabIndex        =   60
      Top             =   5760
      Width           =   1020
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   14
      Left            =   1800
      TabIndex        =   10
      Top             =   3060
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Εττικέτα"
      Height          =   255
      Left            =   9120
      TabIndex        =   58
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Height          =   375
      Left            =   3720
      TabIndex        =   57
      ToolTipText     =   "Αλλαγή κωδικού"
      Top             =   120
      Width           =   135
   End
   Begin VB.TextBox Text14 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   56
      Top             =   6240
      Width           =   3615
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      Left            =   7440
      TabIndex        =   55
      Top             =   2040
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7080
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   2640
      TabIndex        =   11
      Top             =   3060
      Width           =   1100
   End
   Begin MSDataGridLib.DataGrid DBGrid1 
      Bindings        =   "APOT2.frx":000C
      Height          =   2895
      Left            =   6960
      TabIndex        =   51
      Top             =   2520
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "kod"
         Caption         =   "Κωδικός"
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
         DataField       =   "ono"
         Caption         =   "Περιγραφή"
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
            ColumnWidth     =   2880
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   7080
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "DATA2"
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   7080
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "DATA1"
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
      DataField       =   " "
      Height          =   330
      Index           =   6
      Left            =   1680
      TabIndex        =   3
      Top             =   1030
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "αναζήτηση"
      Height          =   375
      Left            =   9120
      TabIndex        =   17
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   1815
      TabIndex        =   8
      Top             =   2430
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   12
      Top             =   3510
      Width           =   1335
   End
   Begin VB.TextBox xti 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   16
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox erg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CheckBox ypoloipo 
      Caption         =   "με υπόλοιπο"
      Height          =   375
      Left            =   9120
      TabIndex        =   29
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Καρτέλλα"
      Height          =   375
      Left            =   6960
      TabIndex        =   28
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Διόρθωση"
      Height          =   375
      Left            =   8235
      TabIndex        =   25
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox kod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   13
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox ono 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   14
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Εξοδος"
      Height          =   375
      Left            =   10470
      TabIndex        =   24
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   2655
      TabIndex        =   9
      Top             =   2430
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   135
      TabIndex        =   7
      Top             =   2430
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   3815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4650
      Left            =   0
      TabIndex        =   35
      Top             =   3960
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8202
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabHeight       =   529
      TabCaption(0)   =   "Προμηθευτής / ελάχιστο στόκ"
      TabPicture(0)   =   "APOT2.frx":0020
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(15)"
      Tab(0).Control(1)=   "Label1(16)"
      Tab(0).Control(2)=   "Label1(17)"
      Tab(0).Control(3)=   "Label1(18)"
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(5)=   "Label1(24)"
      Tab(0).Control(6)=   "Text1(9)"
      Tab(0).Control(7)=   "Text1(8)"
      Tab(0).Control(8)=   "Text1(7)"
      Tab(0).Control(9)=   "Text1(10)"
      Tab(0).Control(10)=   "Text1(15)"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Γραφική παράσταση"
      TabPicture(1)   =   "APOT2.frx":003C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "graph1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Παράμετροι Λογιστικής"
      TabPicture(2)   =   "APOT2.frx":0058
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "timokat"
      Tab(2).Control(1)=   "DataGrid2"
      Tab(2).Control(2)=   "Text1(12)"
      Tab(2).Control(3)=   "Text1(11)"
      Tab(2).Control(4)=   "Label1(20)"
      Tab(2).Control(5)=   "Label1(19)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Κινήσεις κατα μήνα"
      TabPicture(3)   =   "APOT2.frx":0074
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSFlexGrid1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Υπόλοιπα κατά Αποθήκη"
      TabPicture(4)   =   "APOT2.frx":0090
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "MSHFlexGrid1"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "BARCODES"
      TabPicture(5)   =   "APOT2.frx":00AC
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "DataGrid1"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Text3"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Command10"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Command11"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Adodc2"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).ControlCount=   5
      Begin MSAdodcLib.Adodc timokat 
         Height          =   330
         Left            =   -70065
         Top             =   2145
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Caption         =   "Adodc3"
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "APOT2.frx":00C8
         Height          =   1785
         Left            =   -74505
         TabIndex        =   75
         Top             =   2640
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   3149
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   4455
         Top             =   3360
         Width           =   2070
         _ExtentX        =   3651
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
         Caption         =   "BARCODES"
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
      Begin VB.CommandButton Command11 
         Caption         =   "ΔΙΑΓΡΑΦΗ BARCODE"
         Height          =   285
         Left            =   4425
         TabIndex        =   72
         Top             =   2700
         Width           =   2100
      End
      Begin VB.CommandButton Command10 
         Caption         =   "ΝΕΟ BARCODE"
         Height          =   285
         Left            =   4425
         TabIndex        =   71
         Top             =   2010
         Width           =   2100
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   4410
         TabIndex        =   70
         Top             =   1425
         Width           =   2070
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "APOT2.frx":00DE
         Height          =   2940
         Left            =   1035
         TabIndex        =   69
         Top             =   1455
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   5186
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         TabIndex        =   65
         Top             =   2970
         Width           =   1815
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Bindings        =   "APOT2.frx":00F3
         Height          =   3255
         Left            =   -74760
         TabIndex        =   54
         Top             =   960
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5741
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSChart20Lib.MSChart graph1 
         Height          =   3495
         Left            =   -74880
         OleObjectBlob   =   "APOT2.frx":0108
         TabIndex        =   50
         Top             =   720
         Width           =   6375
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   -72360
         TabIndex        =   47
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   -72360
         TabIndex        =   46
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   -72720
         TabIndex        =   41
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   -72720
         TabIndex        =   40
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   -72720
         TabIndex        =   39
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   -72720
         TabIndex        =   38
         Top             =   2040
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   37
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6588
         _Version        =   393216
         Rows            =   13
         Cols            =   5
         GridLines       =   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus σε κάρτα"
         Height          =   375
         Index           =   24
         Left            =   -74280
         TabIndex        =   66
         Top             =   2970
         Width           =   1455
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   -74160
         TabIndex        =   62
         Top             =   3375
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδ.Λογ.Αγορών"
         Height          =   375
         Index           =   20
         Left            =   -73920
         TabIndex        =   49
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδ.Λογ.Πωλήσεων"
         Height          =   375
         Index           =   19
         Left            =   -73920
         TabIndex        =   48
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ράφι-Θέση"
         Height          =   375
         Index           =   18
         Left            =   -74280
         TabIndex        =   45
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Προμηθευτής"
         Height          =   375
         Index           =   17
         Left            =   -74280
         TabIndex        =   44
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ποσότητα απογραφής"
         Height          =   375
         Index           =   16
         Left            =   -74280
         TabIndex        =   43
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ελάχιστο στόκ"
         Height          =   375
         Index           =   15
         Left            =   -74280
         TabIndex        =   42
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Διεύθυνση Επιχείρησης"
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
         Index           =   14
         Left            =   -74760
         TabIndex        =   36
         Top             =   1320
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Οικογένεια"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Index           =   25
      Left            =   3630
      TabIndex        =   77
      Top             =   1005
      Width           =   1695
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode1 
      Height          =   1065
      Left            =   7290
      TabIndex        =   76
      Top             =   8355
      Visible         =   0   'False
      Width           =   1995
      _cx             =   3519
      _cy             =   1879
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
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
      NarrowBarWidth  =   0,03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   0
      Code128CharSet  =   1
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   1
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0,2
      TopMarginCM     =   0,2
      SupplementToEncode=   ""
   End
   Begin VB.Label PROTEIN 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3795
      TabIndex        =   64
      Top             =   3075
      Width           =   720
   End
   Begin VB.Shape Shape1 
      Height          =   915
      Index           =   1
      Left            =   9060
      Top             =   1515
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MarkUP.% κέρδους"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   23
      Left            =   1650
      TabIndex        =   59
      Top             =   2835
      Width           =   840
   End
   Begin VB.Shape Shape1 
      Height          =   1290
      Index           =   0
      Left            =   45
      Top             =   2205
      Width           =   4545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Αποθήκη"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   22
      Left            =   6000
      TabIndex        =   53
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Λιανική τιμή"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   21
      Left            =   2640
      TabIndex        =   52
      Top             =   2820
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Βοηθ.Κλειδί(BARCODE)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   13
      Left            =   0
      TabIndex        =   34
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MarkUP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   12
      Left            =   1665
      TabIndex        =   33
      Top             =   2190
      Width           =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Υπόλοιπο"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   11
      Left            =   1080
      TabIndex        =   32
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   10
      Left            =   6000
      TabIndex        =   31
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Βοηθ.Κωδικός(Εργ)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   9
      Left            =   6000
      TabIndex        =   30
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδικός"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   27
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Περιγραφή"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   26
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Χονδρική τιμή"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   6
      Left            =   2685
      TabIndex        =   23
      Top             =   2205
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Τιμή Αγοράς"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   5
      Left            =   165
      TabIndex        =   22
      Top             =   2205
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Υποοικογένεια"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   4
      Left            =   3465
      TabIndex        =   21
      Top             =   1590
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κατηγ.ΦΠΑ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Μονάδα μέτρησης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Περιγραφή"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδικός"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "APOT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Option Explicit
Dim F_PONTOI
Dim PROTH_FORA
'






Function Update_Next_Pel(meidos As String) As Integer

Dim db As Database
Dim R As Recordset
'12- 9252  141

End Function

Function get_next_pel(meidos As String) As Integer
Dim db As Database
Dim R As Recordset


Set db = OpenDatabase(gDir)
Set db = OpenDatabase(gDir, False, False)
Set db = OpenDatabase(gDir, False, False, gConnect)




Set R = db.OpenRecordset("mem")
If meidos = "e" Then
   get_next_pel = R("epel") + 1
Else
   get_next_pel = R("rpel") + 1
End If
End Function

Private Sub Combo1_GotFocus()
   Combo1.BackColor = vbYellow
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub Combo1_LostFocus()
   Combo1.BackColor = vbWhite
End Sub

Private Sub Combo2_GotFocus()
  Combo2.BackColor = vbYellow
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub Combo2_LostFocus()
  Combo2.BackColor = vbWhite
End Sub

Private Sub Combo3_GotFocus()
  Combo3.BackColor = vbYellow
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub Combo3_LostFocus()
    Combo3.BackColor = vbWhite
End Sub

Private Sub Combo4_LostFocus()
' YPOOIKOGENEIES
Dim R As New ADODB.Recordset
If Len(Combo4.Text) > 0 Then
   R.Open "SELECT *FROM PINAKES WHERE TYPOS=3 AND TIMH=" + left(Combo4.Text, 2) + "ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
   Combo3.Clear


Do While Not R.EOF
  If R("typos") = 3 Then
     Combo3.AddItem Format(R("ayjon"), "00") + " " + left(R("PERIGRAFH") + Space(20), 20)
  Else
     Exit Do
  End If
  R.MoveNext
Loop
R.Close

End If


End Sub

Private Sub Command1_Click()
'exodos
Unload Me
End Sub

Private Sub Command10_Click()
   On Error Resume Next
  Dim d
  Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + Text1(0).Text + "','" + Text3.Text + "')", d
  Adodc2.Refresh
  If d = 0 Then
     MsgBox "Δεν καταχωρήθηκε"
  End If
  
End Sub

Private Sub Command11_Click()
  Gdb.Execute "DELETE  FROM BARCODES WHERE ERG='" + Text3.Text + "' AND KOD='" + Text1(0).Text + "'"
  Adodc2.Refresh
End Sub

Private Sub Command12_Click()
' ETIKETTA ME KEIMENO

Dim ll, apoSCANNER As Long
Dim db
Dim PSIFIA As Integer
Dim fores As Long, K As Integer
Dim X As Printer
Dim barc(200)
Dim mONOMA, mBarcode, mERG, mLTI5
Dim R As New ADODB.Recordset
Dim RD As DAO.Recordset
mONOMA = Text1(1).Text 'ONOMA
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
If Scanner Then

 ' μπουγουδης
Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
Set RD = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM= 1234")
If RD.RecordCount = 0 Then
   MsgBox "δεν υπάρχουν εγγραφές"
   Exit Sub
End If

  
  RD.MoveFirst
  K = 0
  Do While Not RD.EOF
     K = K + 1
     If K > 80 Then Exit Do
     If IsNull(RD("BARCODE")) Then
        barc(K) = " " 'Grid1.TextMatrix(K, f_k) = " "
     Else
        barc(K) = RD("barcode") 'Grid1.TextMatrix(K, f_k) = R("BARCODE")
     End If
     RD.MoveNext
  Loop
  RD.Close
  db.Close
  apoSCANNER = K
End If

fores = Val(InputBox("Πόσες εττικέτες; ", "", 1))
Dim mFORES As Long
For ll = 1 To apoSCANNER

mFORES = fores

If Scanner Then
   R.Close
   R.Open "select * FROM EID WHERE KOD='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   
  If R.EOF Or Len(Trim(barc(ll))) = 0 Then
     mFORES = 0
  Else
   mONOMA = R("ONO") 'ONOMA
   mBarcode = R("KOD") '  ' BARCODE
   mERG = R("ERG") ' erg
   mLTI5 = Format(R("LTI5"), "###0.00") ' LTI5
  End If
End If


Dim Z As Integer
Z = -5 ' InputBox("METATOPISH")

For K = 1 To mFORES



Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992

Printer.FontSize = 18
Printer.FontBold = True
Printer.Print Tab(4 + Z); left(mONOMA, 16) 'Text1(1).Text 'ONOMA
Printer.Print Tab(4 + Z); Mid$(mONOMA, 17, 16) 'Text1(1).Text 'ONOMA
Printer.Print Tab(4 + Z); Mid$(mONOMA, 33, 16) 'Text1(1).Text 'ONOMA
Printer.FontSize = 8
Printer.Print Tab(4 + Z); mBarcode; '  Text1(0).Text;  ' BARCODE




Printer.EndDoc
Next K




Next ll

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


End Sub


Private Sub Command2_Click()
'diortosi
Dim K As Integer, DUM As Integer, db As Database, R As Recordset
Dim X As String
       
      

  If Command2.Caption = "Διόρθωση" Then
      

       Command2.Caption = "Ενημέρωση"
       For K = 1 To 3
          Text1(K).Enabled = True
          Text1(K).ForeColor = vbBlack
       Next
       
          Text1(5).Enabled = True
          Text1(5).ForeColor = vbBlack
       
       
       For K = 7 To 13
          Text1(K).Enabled = True
          Text1(K).ForeColor = vbBlack
       Next
          Text14.Enabled = True
       
          Text14.ForeColor = vbBlack
       
       Combo1.Enabled = True
       Combo2.Enabled = True
       Combo3.Enabled = True
       Text1(1).SetFocus
       
  Else
       
       'Data2.Recordset.Edit
   
'       Data1.Recordset("kod") = Format(Text1(0), "0000")
       
    On Error Resume Next
    
    Data2.Recordset("ONO") = left(Text1(1).Text, Data2.Recordset("ONO").DefinedSize)
       
       Data2.Recordset("FPA") = Val(left(Combo2.Text, 2))
       
       Data2.Recordset("mon") = Combo1.Text
       
       Data2.Recordset("kodlog") = Val(left(Combo3.Text, 3)) ' Combo3.ListIndex  'Right(Combo3.Text, 3)
       
       Data2.Recordset("aeg") = Val(left(Combo4.Text, 3)) ' Combo3.ListIndex  'Right(Combo3.Text, 3)
       
       Data2.Recordset("kodSYNOD") = Text1(11).Text     'Right(Combo3.Text, 3)
       Data2.Recordset("kodlogAG") = Text1(12).Text     'Right(Combo3.Text, 3)
       
       Data2.Recordset("PROM") = Text1(7).Text     'Right(Combo3.Text, 3)
       Data2.Recordset("SPA") = Val(Text1(9).Text)     'Right(Combo3.Text, 3)
       
       Data2.Recordset("PROMHU") = Val(Text1(15).Text)     ' BONUS
       
       Data2.Recordset("UES") = Text1(10).Text     'Right(Combo3.Text, 3)
       Data2.Recordset("APOS") = Val(Text1(8).Text)     'Right(Combo3.Text, 3)
       
       
       
       
       
       Data2.Recordset("pos_kerd2") = Val(Replace(Text1(5).Text, ",", "."))
       Data2.Recordset("pos_kerd") = Val(Replace(Text1(14).Text, ",", ".")) ' Val(Text1(14).Text)
       
       Data2.Recordset("xti") = Val(Replace(Text1(2).Text, ",", ".")) ' Val(Text1(2).Text)
       Data2.Recordset("lti") = Val(Replace(Text1(3).Text, ",", ".")) 'Val(Text1(3).Text)
       Data2.Recordset("lti5") = Val(Replace(Text1(13).Text, ",", ".")) ' Val(Text1(13).Text)
       Data2.Recordset("ERG") = Text1(6).Text
       Data2.Recordset("MEMO") = Text14.Text
       Data2.Recordset.Update
       
       For K = 1 To 3
          Text1(K).Enabled = False
          Text1(K).ForeColor = vbBlack
       Next
       
          Text1(5).Enabled = False
          Text1(5).ForeColor = vbBlack
       
       
       For K = 7 To 13
          Text1(K).Enabled = False
          Text1(K).ForeColor = vbBlack
       Next
       kod.SetFocus
       Command2.Caption = "Διόρθωση"
  End If
End Sub


Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim a
a = 0
End Sub

Private Sub Command3_Click()
' kartella
  apot21.Text1.Text = Text1(0).Text
  apot21.Text2.Text = Text1(1).Text
  apot21.Text3.Text = " Αποθ." + Combo9.Text
  apot21.Label1.Caption = Str(Val(left(APOT2.Combo9.Text, 2)))
  apot21.Label2.Caption = Data1.Recordset("apos")
  'Me.Hide
  
  apot21.SHOW 1
  
End Sub

Private Sub Command4_Click()
Dim R As New ADODB.Recordset
Dim RECS As Long

R.Open "SELECT COUNT(*) FROM EGGTIM WHERE KODE='" + Text1(0).Text + "'", Gdb, adOpenForwardOnly, adLockReadOnly

If R(0) > 0 Then
   MsgBox "Αδύνατη η διαγραφή επειδή υπάρχουν κινησεις"
   R.Close
   Exit Sub
End If
Select Case MsgBox(Text1(1).Text + vbCrLf + "Είσαι σίγουρος για την διαγραφή;", vbYesNo Or vbExclamation Or vbSystemModal Or vbDefaultButton1, "Θα διαγραφεί το " + Text1(0).Text)

    Case vbYes

    Case vbNo
         Exit Sub
End Select








Gdb.Execute "DELETE FROM EID WHERE KOD='" + Text1(0).Text + "'", RECS
If RECS > 0 Then
   Call MsgBox("Διεγράφησαν " + Str(RECS) + "" _
               & vbCrLf & " εγγραφές " _
               , vbExclamation Or vbSystemModal, App.title)
Else
   Call MsgBox("Δεν ευρέθη η εγγραφή", vbCritical Or vbSystemModal, App.title)
End If



End Sub

Private Sub Command5_Click()
' anazitisi

If Len(ono.Text) > 0 Then
 If ypoloipo Then
    If Val(xti) > 0 Then
       Data2.RecordSource = "select *from EID where POS>0 and XTI>" + Str(Val(xti)) + " and ONO like '" + ono.Text + "%';"
    Else
       Data2.RecordSource = "select *from EID where POS>0 and ONO like '" + ono.Text + "%';"
    End If
    
  Else
    If Val(xti) > 0 Then
       Data2.RecordSource = "select *from EID where XTI>" + Str(Val(xti)) + " and ONO like '" + UCase(ono.Text) + "%';"
    Else
       Data2.RecordSource = "select *from EID where  ONO like '" + UCase(ono.Text) + "%';"
    End If
  End If
   Data2.Refresh
   DBGrid1.SetFocus
   mDBGrid1_RowColChange
  Exit Sub
End If


 If Len(kod) > 0 Then
    If ypoloipo Then
      If Val(xti) > 0 Then
          Data2.RecordSource = "select *from EID where POS>0 and XTI>" + Str(Val(xti)) + " and KOD like '" + kod.Text + "%';"
      Else
          Data2.RecordSource = "select *from EID where POS>0 and KOD like '" + kod.Text + "%';"
      End If
    Else
       If Val(xti) > 0 Then
           Data2.RecordSource = "select *from EID where  XTI>" + Str(Val(xti)) + " and KOD like '" + kod.Text + "%';"
       Else
           Data2.RecordSource = "select *from EID where KOD like '" + kod.Text + "%';"
       End If
    End If
       Data2.Refresh
'       xti.SetFocus
'       Data2.Refresh
       If Data2.Recordset.EOF Then ' δεν βρηκε κωδικούς
          kod.SetFocus
       Else
          DBGrid1.SetFocus
       End If
       
  End If

If Len(erg) > 0 Then
  
     Data2.RecordSource = UCase("select *from EID where ERG like '" + erg.Text + "%';")
     Data2.Refresh
     DBGrid1.SetFocus

End If

If Len(xti) > 0 Then
     On Error GoTo 0
     Data2.RecordSource = "select *from EID where KOD IN (SELECT KOD FROM BARCODES WHERE ERG like '" + xti.Text + "%')"
     Data2.Refresh
     DBGrid1.SetFocus

End If

mDBGrid1_RowColChange

End Sub

Private Sub Command6_Click()
Dim neos As String, RECS As Integer
  
  neos = InputBox("Δώσε τον νέο κωδικό ")
  Dim R As New ADODB.Recordset
  If Len(neos) = 0 Then
     Call MsgBox("Δεν δώσατε νέο κωδικό", vbCritical Or vbDefaultButton1, App.title)
     Exit Sub
  End If
     
  
  
  
  
  R.Open "select count(*) from EID WHERE KOD='" + neos + "'", Gdb, adOpenForwardOnly, adLockReadOnly
  If R(0) = 0 Then
     Gdb.Execute "UPDATE EID SET KOD='" + neos + "' WHERE KOD='" + Text1(0).Text + "'", RECS
     MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")
     
     Gdb.Execute "UPDATE EGGTIM SET KODE='" + neos + "' WHERE KODE='" + Text1(0).Text + "'", RECS
     MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ ΣΕ ΤΙΜΟΛΟΓΙΑ  " + Format(RECS, "#####")
          
     Gdb.Execute "UPDATE BARCODES SET KOD='" + neos + "' WHERE KOD='" + Text1(0).Text + "'", RECS
     MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ ΣΕ BARCODES  " + Format(RECS, "#####")
     
     
     
     
  Else
     MsgBox "Υπάρχει ήδη ο κωδικός " + neos
  End If
End Sub

Sub LAG_ETIK()
'========================etiketes ========================================================
Dim d1 As Date, d2 As Date, R As Integer, C As Integer, K As Integer, s As Single
Dim db As Database, DUM, rec As New ADODB.Recordset
Dim mlabel As String
Dim mper As String
Dim mtimes As Single
Dim TIMM As Single
Dim timologio As String

'On Error Resume NextN  N


Dim X As Printer

For Each X In Printers
    If InStr(X.DeviceName, "SATO") > 0 Then
        ' Set printer as system default.
        Set Printer = X
        ' Stop looking for a printer.
        Exit For
    End If
Next


On Error GoTo 0
rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

   
   If IsNull(Data2.Recordset("ERG")) Then
       GoSub FIND_N_NUMBER
   Else
       If Val(Data2.Recordset("ERG")) = 0 Then
          GoSub FIND_N_NUMBER
       Else
          mlabel = Data2.Recordset("erg")
       End If
   End If
   
   mper = Data2.Recordset("ono")
   mtimes = InputBox("Πόσες ετικέτες;", mper, 1)
   
   TIMM = Format(Data2.Recordset("lti") * (100 + g_Fpa(Data2.Recordset("fpa"))) / 100, "#####.00")
   
   'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

For K = 1 To Val(mtimes)

    If True Then
      Printer.FontName = "128"
      Printer.FontSize = 6
      Printer.Print Tab(12); mlabel
      Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
      Printer.FontSize = 8
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print


      Printer.Print Tab(20); mlabel
      Printer.Print Tab(20); mper
      Printer.Print Tab(20); timologio
      Printer.FontSize = 10
      Printer.Print Tab(20); Trim(LTrim(TIMM)) + " ΕΥΡΩ"
      Printer.EndDoc
    Else
      Printer.Print Tab(20); mper
      Printer.Print Tab(30); mlabel
      Printer.FontName = "ean13" ' "128"
      Printer.FontSize = 6
      Printer.Print Tab(18); Data2.Recordset("kod") ' mlabel
      Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
      Printer.FontSize = 8
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.EndDoc
    End If
Next
    

Exit Sub

FIND_N_NUMBER:
'************************
rec.MoveFirst
rec.Move 10
  'rec.Edit
  If IsNull(rec("epel")) Then
      rec("epel") = 1
  Else
      rec("epel") = rec("epel") + 1
  End If
  rec.Update
  mlabel = Right("00000000" + LTrim(Str(rec("EPEL"))), 8)
  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Data2.Recordset("KODE") + "'"
Return


End Sub
Private Sub Command7_Click()

Dim fores As Long, K As Integer
Dim X As Printer
Dim DUM
















On Error Resume Next


If Len(Dir("C:\LAGEURO\LAGAKIS.TXT", vbNormal)) > 0 Then
    LAG_ETIK
    Exit Sub
End If




If Len(Dir("C:\LAGEURO\MAKISGET.EXE", vbNormal)) > 0 Then
    DUM = Shell("c:\lageuro\makisget.exe " + Data2.Recordset("KOD") + " " + Replace(Data2.Recordset("ONO"), " ", "_") + " " + Format(Data2.Recordset("LTI5"), "##0.00") + " " + Format(Data2.Recordset("POS_KERD"), "##0") + " " + Data2.Recordset("ERG"))
    Exit Sub
End If


If Len(Dir("C:\DIONISIS.EXE", vbNormal)) > 0 Then
    fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    Dim param
    
    If Len(DATAGRID1.Text) < 12 Then
       MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
       Exit Sub
    End If
    
    If Len(DATAGRID1.Text) = 12 Then
         BarCode1.SymbologyID = UPCA
    Else
         BarCode1.SymbologyID = EAN13
    End If
    
    For Each X In Printers
        If InStr(X.DeviceName, "2844") > 0 Then
          ' Set printer as system default.
        
          Set Printer = X
          ' Stop looking for a printer
          Exit For
        End If
    Next
    
    BarCode1.DataToEncode = DATAGRID1.Text
    
    
   For K = 1 To fores
    Printer.CurrentY = 400
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 600
    Printer.FontSize = 12
    
    Printer.Print Text1(0).Text
    Printer.PaintPicture BarCode1.Picture, 348, 824
    Printer.EndDoc
   Next
    
    
   Exit Sub
End If



For Each X In Printers
    If InStr(X.DeviceName, "SATO") > 0 Then
        ' Set printer as system default.
        Set Printer = X
        ' Stop looking for a printer
        Exit For
    End If
Next



fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text))
If Len(Dir("C:\LAGEURO\MPOYGET.EXE", vbNormal)) > 0 Then
  Shell ("C:\LAGEURO\MPOYGET.EXE " + Text1(0).Text + " " + Str(fores) + " " + Replace(Text1(1).Text, " ", ".") + " " + "1")
  Exit Sub
End If
' MPOYGET.EXE 5201219046154 2 PERIGRAFH 3.56

For K = 1 To fores
Printer.Print Tab(20); Text1(1).Text 'ONOMA
Printer.Print Tab(30); Text1(6).Text ' mlabel
If Len(Trim(Text1(0).Text)) = 12 Then
   Printer.FontName = "UPCA" ' "128"
Else
   Printer.FontName = "ean13" ' "128"
End If

Printer.FontSize = 6
Printer.Print Tab(18); Text1(0).Text 'eggtim.Recordset("kode") ' mlabel
Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
Printer.FontSize = 8
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.EndDoc
Next
End Sub

Private Sub Command8_Click()
Dim ll, apoSCANNER As Long
Dim db
Dim PSIFIA As Integer
Dim fores As Long, K As Integer
Dim X As Printer
Dim barc(200)
Dim mONOMA, mBarcode, mERG, mLTI5
Dim R As New ADODB.Recordset
Dim RD As DAO.Recordset
mONOMA = Text1(1).Text 'ONOMA
mBarcode = Text1(0).Text '  ' BARCODE
mERG = Text1(6).Text ' erg
mLTI5 = Format(Val(Text1(13).Text), "###0.00") ' LTI5


If Len(Dir("C:\DIONISIS.EXE", vbNormal)) > 0 Then
    
    'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    Dim param
    
    If Len(DATAGRID1.Text) < 12 Then
       MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
       Exit Sub
    End If
    
    If Len(DATAGRID1.Text) = 12 Then
         BarCode1.SymbologyID = UPCA
    Else
         BarCode1.SymbologyID = EAN13
    End If
    
    For Each X In Printers
        If InStr(X.DeviceName, "2844") > 0 Then
          ' Set printer as system default.
        
          Set Printer = X
          ' Stop looking for a printer
          Exit For
        End If
    Next
    
    BarCode1.DataToEncode = DATAGRID1.Text
    
    
   'For K = 1 To fores
    Printer.CurrentY = 200
    Printer.FontSize = 10
   Printer.Print Text1(1).Text 'perigrafh
   Printer.Print Text1(0).Text 'kodikos
   
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 400
    'Printer.FontSize = 12
    Printer.FontSize = 10
    
    
    Printer.CurrentX = 200
        Printer.CurrentY = 600
    
    Printer.PaintPicture BarCode1.Picture, 0, 1124
    
    
    Printer.CurrentX = 1000
    Printer.FontSize = 24
    Printer.FontBold = True
    Printer.Print Format(Text1(13).Text, "##0.00") + " €" ' timh
    ' Printer.Print "4.55 €" ' timh
    Printer.EndDoc
   'Next
    
    
   Exit Sub
End If



























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
If Scanner Then

 ' μπουγουδης
Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
Set RD = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM= 1234")
If RD.RecordCount = 0 Then
   MsgBox "δεν υπάρχουν εγγραφές"
   Exit Sub
End If

  
  RD.MoveFirst
  K = 0
  Do While Not RD.EOF
     K = K + 1
     If K > 80 Then Exit Do
     If IsNull(RD("BARCODE")) Then
        barc(K) = " " 'Grid1.TextMatrix(K, f_k) = " "
     Else
        barc(K) = RD("barcode") 'Grid1.TextMatrix(K, f_k) = R("BARCODE")
     End If
     RD.MoveNext
  Loop
  RD.Close
  db.Close
  apoSCANNER = K
End If

fores = Val(InputBox("Πόσες εττικέτες; ", "", 1))
Dim mFORES As Long
For ll = 1 To apoSCANNER

mFORES = fores

If Scanner Then
   R.Close
   R.Open "select * FROM EID WHERE KOD='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   
  If R.EOF Or Len(Trim(barc(ll))) = 0 Then
     mFORES = 0
  Else
   mONOMA = R("ONO") 'ONOMA
   mBarcode = R("KOD") '  ' BARCODE
   mERG = R("ERG") ' erg
   mLTI5 = Format(R("LTI5"), "###0.00") ' LTI5
  End If
End If


Dim Z As Integer
Z = 10 ' InputBox("METATOPISH")

For K = 1 To mFORES



Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992

Printer.FontSize = 8
Printer.FontBold = True
'll = 45 - Len(Text1(1).Text)

' Printer.Print Tab(3 + ll / 2); Text1(1).Text 'ONOMA
Printer.Print Tab(4 + Z); mONOMA 'Text1(1).Text 'ONOMA
Printer.FontSize = 8
Printer.Print Tab(4 + Z); mBarcode; '  Text1(0).Text;  ' BARCODE


'Printer.Print Tab(30); Text1(6).Text ' mlabel
' Printer.Print
If Len(Trim(Text1(0).Text)) = 12 Then
   Printer.FontName = "UPCA" ' "128"
Else
   Printer.FontName = "EAN13SMALL" ' "128"
End If


'If Len(Dir("C:\LAGEURO\NOPONTOI.TXT", vbNormal)) > 0 Then
  
'Else

If F_PONTOI = 1 Then
   Printer.FontSize = 5
   Printer.Print Tab(15 + Z); mBarcode 'Text1(0).Text 'eggtim.Recordset("kode") ' mlabel
End If




Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
Printer.FontBold = True

Printer.FontSize = 8
Printer.Print Tab(4 + Z); mERG ' Text1(6).Text ' erg
Printer.FontSize = 24
If Len(mLTI5) = 6 Then ' 208.35
     Printer.Print Tab(7 + Z - 7); " € " + mLTI5; ""
ElseIf Len(mLTI5) = 5 Then '18.35
     Printer.Print Tab(8 + Z - 7); "€ " + mLTI5; ""
Else
     Printer.Print Tab(10 + Z - 7); "€ " + mLTI5; ""
End If
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 8
'  Printer.Print Tab(20); Text1(0).Text
 Printer.Print


'IF BONUS THEN

'Printer.Print "BONUS"
' Printer.Print ; Tab(4); "-1.20"
'ELSE
If F_PONTOI = 0 Then 'Len(Dir("C:\LAGEURO\NOPONTOI.TXT", vbNormal)) > 0 Then
   Printer.Print Tab(Z + 3); "5-" + Text1(14).Text
   
Else
    ' ΜΠΟΥΓΟΥΔΗΣ
  
    Printer.CurrentX = 200
    Printer.CurrentY = 500
    Printer.FontSize = 14
  
  If Val(mLTI5) > 2.99 Then
     Printer.Print Tab(Z - 2); "ΠΟΝΤΟΙ"
     PSIFIA = Len(Format(Int(Val(mLTI5) / 3), "##0"))
     If PSIFIA = 3 Then Printer.Print ; Tab(5 + Z - 5); Format(Int(Val(mLTI5) / 3), "##0")
     If PSIFIA = 1 Then Printer.Print ; Tab(6 + Z - 4); Format(Int(Val(mLTI5) / 3), "##0")
     If PSIFIA = 2 Then Printer.Print ; Tab(5 + Z - 4); Format(Int(Val(mLTI5) / 3), "##0")
  End If
End If



Printer.EndDoc
Next K




Next ll

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

End Sub

Private Sub Command9_Click()
  Data2.RecordSource = UCase(" SELECT * FROM EID WHERE MEMO LIKE '%" + Text2.Text + "%';")
  Data2.Refresh
  DBGrid1.SetFocus
  

'   SELECT * FROM EID WHERE MEMO LIKE '%DOKIM%'
End Sub

Private Sub DBGrid1_Click()
   mDBGrid1_RowColChange
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
       
       KeyAscii = 0
       mDBGrid1_RowColChange
       Command2.SetFocus
  End If
End Sub

Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 38 Or KeyCode = 40 Then 'ano=38 kai kato=40 belos
    mDBGrid1_RowColChange
 End If
End Sub

Private Sub mDBGrid1_RowColChange() '(LastRow As Variant, ByVal LastCol As Integer)
Dim G, M, a, GA, MA, P, K, SM, SG
Dim mPos01 As Single, mPos02 As Single

On Error Resume Next

'  Me.Caption = Now
Text1(0).Text = Data2.Recordset("kod")
Text1(1).Text = Data2.Recordset("ONO")


Text1(4).Text = IIf(IsNull(Data2.Recordset("pos")), " ", Data2.Recordset("pos"))


Combo2.Text = Data2.Recordset("FPA")
       
Combo1.Text = Data2.Recordset("mon")
Combo4.Text = Combo4.List(Val(Data2.Recordset("AEG")) - 1)

'Dim N As Integer
'For K = 0 To Combo4.ListCount - 1
'   If Data2.Recordset("AEG") = Val(left(Combo4.List(K), 2)) Then
'      Combo4.Text = Combo4.List(Val(Data2.Recordset("AEG")))
'   End If
'Next



Dim R As New ADODB.Recordset

If Len(Combo4.Text) > 0 Then

R.Open "SELECT *FROM PINAKES WHERE TYPOS=3 AND TIMH=" + left(Combo4.Text, 2) + "ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Combo3.Clear


Do While Not R.EOF
  If R("typos") = 3 Then
     Combo3.AddItem Format(R("ayjon"), "00") + " " + left(R("PERIGRAFH") + Space(20), 20)
  Else
     Exit Do
  End If
  R.MoveNext
Loop
R.Close

End If









Combo3.Text = Combo3.List(Val(Data2.Recordset("kodlog")) - 1)
       
Text1(2).Text = Data2.Recordset("xti")


'If IsNull(Data2.Recordset("pos_kerd2")) Then
   Text1(5).Text = nNull(Data2.Recordset("pos_kerd2"))
'End If


'If IsNull(Data2.Recordset("pos_kerd")) Then
   Text1(14).Text = nNull(Data2.Recordset("pos_kerd"))
'End If

Text1(6).Text = CNull(Data2.Recordset("ERG"))

Text1(3).Text = Data2.Recordset("lti")
Text1(13).Text = Data2.Recordset("lti5")
 
 If IsNull(Data2.Recordset("MEMO")) Then
    Text14.Text = ""
 Else
    Text14.Text = Data2.Recordset("MEMO")
 End If
 
 Text1(15).Text = Data2.Recordset("PROMHU")    ' BONUS
       
       Text1(11).Text = Data2.Recordset("kodSYNOD")
       Text1(12).Text = Data2.Recordset("kodlogAG")
       
       Text1(7).Text = Data2.Recordset("PROM")    'Right(Combo3.Text, 3)
        Text1(9).Text = Data2.Recordset("SPA")   'Right(Combo3.Text, 3)
       
       Text1(10).Text = Data2.Recordset("UES")    'Right(Combo3.Text, 3)
       Text1(8).Text = Data2.Recordset("APOS")



If IsNull(Data2.Recordset("POS01")) Then mPos01 = 0 Else mPos01 = Data2.Recordset("POS01")

If IsNull(Data2.Recordset("POS02")) Then mPos02 = 0 Else mPos02 = Data2.Recordset("POS02")

Label2.Caption = "ΤΕΛ.ΕΝΗΜΕΡΩΣΗ: KENΤΡ :" + Format(mPos01, "###0") + " ΥΠΟΚ 1 : " + Format(mPos02, "###0")





  MSFlexGrid1.row = 0
  MSFlexGrid1.col = 1
  MSFlexGrid1.Text = "Αγορές"
  MSFlexGrid1.col = 2
  MSFlexGrid1.Text = "Αξία Αγορ."
  
  MSFlexGrid1.col = 3
  MSFlexGrid1.Text = "Πωλήσεις"

  MSFlexGrid1.col = 4
  MSFlexGrid1.Text = "Αξία πωλήσεων"

'' Γραφική παράσταση
's = 0

'For l = 1 To 12
   ' Exit Sub
 '  G.row = 1 + addit: G.Column = l: G.Data = Stat(l, addit2 + 1) - Stat(l, addit2 + 2)
  ' G.row = 4 + addit: G.Column = l: G.Data = Stat(l, addit2 + 4) - Stat(l, addit2 + 5)
'Next
If Data2.Recordset.RecordCount = 0 Then Exit Sub


SM = 0: SG = 0

' Dim R As New ADODB.Recordset

On Error GoTo latos
R.Open "SELECT MONTH(HME),SUM(XRE),SUM(PIS),SUM(XRE*TIMM),SUM(PIS*TIMM) FROM EGGTIM WHERE KODE='" + Data2.Recordset("kod") + "' GROUP BY MONTH(HME) ", Gdb, adOpenForwardOnly, adLockReadOnly
Dim A_XRE(12) As Single
Dim A_PIS(12) As Single
Dim A_XREAJ(12) As Single
Dim A_PISAJ(12) As Single

For K = 1 To 12: A_XRE(K) = 0: A_PIS(K) = 0: A_XREAJ(K) = 0: A_PISAJ(K) = 0: Next
' Exit Sub
If Not R.EOF Then


On Error Resume Next
R.MoveFirst
Do While Not R.EOF
   A_XRE(R(0)) = R(1): A_XREAJ(R(0)) = R(3)
   A_PIS(R(0)) = R(2): A_PISAJ(R(0)) = R(4)
   R.MoveNext
Loop

End If

'graph1.top = 4680  'arxika
 '  graph1.height = 4485
   
  ' graph1.top = 46
   'graph1.height = 8485

MSFlexGrid1.Visible = False

For K = 1 To 12
'  G = "G" + Format(k, "00")
'  M = "M" + Format(k, "00")
'  GA = "g" + Format(k, "00")
'  MA = "M" + Format(k, "00")
  
  graph1.row = 1: graph1.Column = K: graph1.Data = A_XRE(K) ' DATA2.Recordset(G)
  graph1.row = 2: graph1.Column = K: graph1.Data = A_PIS(K) 'DATA2.Recordset(M)
   graph1.ColumnLabel = K
    graph1.Refresh
    
   
  
  MSFlexGrid1.row = K
  MSFlexGrid1.col = 0
  MSFlexGrid1.Text = K
    
  MSFlexGrid1.col = 1

  
  SM = SM + A_PIS(K) 'DATA2.Recordset(M)
  SG = SG + A_XRE(K) ' DATA2.Recordset(G)
  
  MSFlexGrid1.Text = Format(A_XRE(K), "#####.##") ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset(G)), "####.##")

  
  MSFlexGrid1.col = 2
  MSFlexGrid1.Text = Format(A_XREAJ(K), "#####.##") ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset("XTI") * DATA2.Recordset(G)), "#####,##")

  MSFlexGrid1.col = 3
  MSFlexGrid1.Text = Format(A_PIS(K), "#####.##") 'Format(IIf(IsNull(DATA2.Recordset(M)), 0, DATA2.Recordset(M)), "#####,##")


  MSFlexGrid1.col = 4
  MSFlexGrid1.Text = Format(A_PISAJ(K), "#####.##") 'Format(IIf(IsNull(DATA2.Recordset(MA)), 0, DATA2.Recordset("LTI") * DATA2.Recordset(MA)), "#####,##")

  
  
  
  
  
  
  
Next

If MSFlexGrid1.rows >= 14 Then
   
   MSFlexGrid1.rows = 13
End If

   

MSFlexGrid1.AddItem ("Σύνολα" & Chr(9) & Format(SG, "####.##") & Chr(9) & Format(0, "####.##") & Chr(9) & Format(SM, "####.##"))
   
MSFlexGrid1.Visible = True

   
   
   
Adodc1.ConnectionString = gConnect
Adodc1.RecordSource = "SELECT KODE,APOT,SUM(XRE) AS [ΕΙΣ],SUM(PIS) AS [ΕΞΑΓ]," _
& " (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" _
& " (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ) AS [ΥΠΟΛ]" _
& " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' GROUP BY KODE,APOT"

Adodc1.Refresh

Adodc2.ConnectionString = gConnect
Adodc2.RecordSource = "select ERG from BARCODES where  LEFT(KOD,14)='" + Text1(0).Text + "'"
Adodc2.Refresh

'DataGrid2.DataSource = timokat
timokat.ConnectionString = gConnect
timokat.RecordSource = "select * from TIMOKAT where  LEFT(KOD,14)='" + Text1(0).Text + "'"
timokat.Refresh









Exit Sub
latos:
MsgBox Err.Description

Resume Next


End Sub

Private Sub erg_GotFocus()
   erg.BackColor = vbYellow
End Sub

Private Sub erg_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub erg_LostFocus()
  erg.BackColor = vbWhite
End Sub

Private Sub Form_Load()
 Dim db, R As New ADODB.Recordset
 APOT2.Picture = LoadPicture(gPicture)
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

If gXEIRISTHS < 9 Then
  Command3.Visible = False
  Text1(2).Visible = False
  Text1(14).Visible = False
  Text1(5).Visible = False
  Text1(3).Visible = False
  Label1(5).Visible = False
  Label1(12).Visible = False
  Label1(6).Visible = False
  Label1(23).Visible = False
End If




 
 
 
 'Set db = OpenDatabase(gDir, False, False, gConnect)
   R.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
   

'R.Index = "pinakes"

'FPA
'R.Seek "=", 1, 1
Do While Not R.EOF
   If R("typos") = 1 Then
      Combo2.AddItem Str(R("AYJON")) + " -> " + Str(R("TIMH"))
   End If
   R.MoveNext
Loop
' mon.metrhshs
 R.Close
 
 
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=2 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
  If R("typos") = 2 Then
     Combo1.AddItem R("PERIGRAFH")
  Else
     Exit Do
  End If
  R.MoveNext
Loop
' kathgories

R.Close


'αποθηκη
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=4 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
  If R("typos") = 4 Then
     Combo9.AddItem Format(R("ayjon"), "##") + "." + R("PERIGRAFH")
  Else
     Exit Do
  End If
  R.MoveNext
Loop
Combo9.Text = Combo9.List(0)

R.Close





' OIKOGENEIES

R.Open "SELECT *FROM PINAKES WHERE TYPOS=11 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
'R.MoveFirst
Do While Not R.EOF
  If R("typos") = 11 Then
     Combo4.AddItem Format(R("ayjon"), "00") + " " + left(R("PERIGRAFH") + Space(20), 20)
  Else
     Exit Do
  End If
  R.MoveNext
Loop
R.Close





F_PONTOI = Val(FindParametroi("APOT2", "F_PONTOI", "1", "ΣΤΗΝ ΕΤΙΚΕΤΤΑ ΠΟΝΤΟΥΣ=1 ΟΧΙ=0"))


 
 
 
 

Data1.RecordSource = "SELECT TOP 10 *FROM EID"
Data1.ConnectionString = gConnect
Data1.Refresh

Data2.ConnectionString = gConnect
     Data2.RecordSource = "select TOP 20 * from EID where POS>0;"
     Data2.Refresh
     



PROTH_FORA = 0



End Sub

Private Sub Option1_Click(Index As Integer)
  
End Sub

Private Sub MSChart1_OLEStartDrag(Data As MSChart20Lib.DataObject, AllowedEffects As Long)

End Sub

Private Sub Form_Paint()
 
 If PROTH_FORA = 0 Then
    PROTH_FORA = 1
    kod.SetFocus
 End If
   
End Sub

Private Sub kod_GotFocus()
 kod.BackColor = vbYellow
End Sub

Private Sub kod_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub kod_LostFocus()
    kod.BackColor = vbWhite
End Sub

Private Sub ono_GotFocus()
  ono.BackColor = vbYellow
End Sub

Private Sub ono_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub ono_LostFocus()
 ono.BackColor = vbWhite
  'DBGrid1.SetFocus
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub Text1_GotFocus(Index As Integer)
  Text1(Index).BackColor = vbYellow
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       
       KeyAscii = 0 'suppress the beep
       If Index = 13 Then
           Command2.SetFocus
      
       
       Else
           keybd_event VK_TAB, 0, 0, 0 'send a tab
       End If
    End If
End Sub

Private Sub Text1_Change(Index As Integer)
  If Index = 5 Or Index = 4 Then
       Text1(3).Text = Val(Text1(2).Text) * (100 + Val(Text1(5).Text)) / 100
  End If
  
End Sub

Private Sub Text1_LostFocus(Index As Integer)
  
   
    
    
    Text1(Index).BackColor = vbWhite
    If Index = 14 Then
       PROTEIN.Caption = Val(Text1(2).Text) * (1 + g_Fpa(Val(Combo2.Text)) / 100) * (100 + Val(Text1(14).Text)) / 100
    End If
    
    End Sub

Private Sub Text14_GotFocus()
  Text14.BackColor = vbYellow
  Text14.ForeColor = vbBlack
End Sub

Private Sub Text14_LostFocus()
    Text14.BackColor = vbWhite
 '
End Sub

Private Sub xti_Change()
  ' Data2.RecordSource = "select *from EID where pos>0 and xti>" + Str(Val(xti)) + " order by xti;"
  ' Data2.Refresh
End Sub

Private Sub xti_GotFocus()
  xti.BackColor = vbYellow
End Sub

Private Sub xti_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub xti_LostFocus()
   xti.BackColor = vbWhite
End Sub





