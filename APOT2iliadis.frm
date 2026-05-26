VERSION 5.00
Object = "{CCA214C0-DFEB-4C91-9F0D-2665F77F6E23}#1.2#0"; "IDAutomationLinear.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form APOT2 
   BackColor       =   &H0080C0FF&
   ClientHeight    =   11196
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14592
   ControlBox      =   0   'False
   Icon            =   "APOT2iliadis.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11196
   ScaleWidth      =   14592
   Begin VB.ComboBox XRONIES 
      Height          =   288
      Left            =   12720
      TabIndex        =   131
      Text            =   "XRONIES"
      Top             =   2160
      Width           =   1212
   End
   Begin VB.CheckBox cXRONIES 
      Caption         =   "Παλιές Χρονιές"
      Height          =   372
      Left            =   11640
      TabIndex        =   130
      Top             =   2076
      Width           =   972
   End
   Begin VB.TextBox ONO2 
      Height          =   615
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   129
      Top             =   8640
      Width           =   4365
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Αναζήτηση απο ρυθμιζόμενες"
      Height          =   255
      Left            =   7920
      TabIndex        =   128
      Top             =   8880
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid DBGRID1 
      Bindings        =   "APOT2iliadis.frx":000C
      Height          =   2895
      Left            =   7080
      TabIndex        =   127
      Top             =   2520
      Width           =   7095
      _ExtentX        =   12510
      _ExtentY        =   5101
      _Version        =   393216
      FixedCols       =   0
      HighLight       =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc data22 
      Height          =   330
      Left            =   8640
      Top             =   8760
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
      Caption         =   "Υπολογισμός Mηνιαίων "
      Height          =   255
      Left            =   480
      TabIndex        =   126
      Top             =   8760
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox XondrmeFPA 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3915
      TabIndex        =   124
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   122
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   10080
      Picture         =   "APOT2iliadis.frx":0021
      ScaleHeight     =   444
      ScaleWidth      =   564
      TabIndex        =   119
      Top             =   165
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   12120
      TabIndex        =   118
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command15 
      Caption         =   "ΕΤΙΚΕΤΤΑ"
      Height          =   195
      Left            =   6480
      TabIndex        =   114
      Top             =   10920
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Εύρεση αρχείου εικόνας"
      Height          =   300
      Left            =   10695
      TabIndex        =   89
      Top             =   8805
      Width           =   2280
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7995
      Top             =   8925
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command13 
      Height          =   360
      Left            =   12885
      TabIndex        =   88
      Top             =   1605
      Width           =   120
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Height          =   330
      Index           =   16
      Left            =   4920
      TabIndex        =   78
      Top             =   3600
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   10680
      ScaleHeight     =   2484
      ScaleWidth      =   3624
      TabIndex        =   77
      Top             =   6240
      Width           =   3675
   End
   Begin VB.ComboBox Combo5 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4920
      TabIndex        =   75
      Top             =   3180
      Width           =   1800
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   450
      Left            =   13305
      ScaleHeight     =   20.4
      ScaleMode       =   2  'Point
      ScaleWidth      =   40.2
      TabIndex        =   74
      Top             =   8895
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Εττικέτα με περιγραφή"
      Height          =   405
      Left            =   11640
      TabIndex        =   68
      Top             =   1605
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3615
      TabIndex        =   6
      Top             =   1275
      Width           =   1785
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ΑΝΑΖΗΤΗΣΗ"
      Height          =   255
      Left            =   7995
      TabIndex        =   63
      Top             =   8445
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9435
      TabIndex        =   62
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CheckBox Scanner 
      Caption         =   "Από Scanner"
      Height          =   195
      Left            =   10095
      TabIndex        =   58
      ToolTipText     =   "Aπό αριθμό τιμ. 1234"
      Top             =   2205
      Width           =   1440
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Εττικέτα ράφι"
      Height          =   225
      Left            =   10080
      TabIndex        =   56
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Διαγραφή"
      Height          =   495
      Left            =   10845
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   " "
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   14
      Left            =   1995
      TabIndex        =   11
      Top             =   3045
      Width           =   900
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Εττικέτα"
      Height          =   255
      Left            =   10080
      TabIndex        =   54
      Top             =   1605
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Height          =   375
      Left            =   3720
      TabIndex        =   53
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
      TabIndex        =   52
      ToolTipText     =   "#ΚΩΔΙΚΟΣ,ΠΟΣΟΤΗΤΑ  : ΑΝ ΠΡΟΚΕΙΤΑΙ ΓΙΑ ΣΥΝΘΕΤΟ ΕΙΔΟΣ"
      Top             =   6240
      Width           =   3615
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      Left            =   8400
      TabIndex        =   51
      Top             =   2085
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7080
      Top             =   7680
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
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   " "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   3915
      TabIndex        =   12
      Top             =   3030
      Width           =   975
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   7080
      Top             =   6840
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
      Left            =   7080
      Top             =   7200
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
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
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
      Left            =   10080
      TabIndex        =   18
      Top             =   1125
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   1995
      TabIndex        =   9
      Top             =   2625
      Width           =   900
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
      Index           =   4
      Left            =   1800
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox xti 
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
      Left            =   8400
      TabIndex        =   17
      Top             =   1605
      Width           =   1575
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
      Left            =   8400
      TabIndex        =   16
      Top             =   1125
      Width           =   1575
   End
   Begin VB.CheckBox ypoloipo 
      Caption         =   "με υπόλοιπο"
      Height          =   375
      Left            =   10080
      TabIndex        =   28
      Top             =   645
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Καρτέλλα"
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Διόρθωση"
      Height          =   495
      Left            =   8805
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5520
      Width           =   1095
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
      Left            =   8400
      TabIndex        =   14
      ToolTipText     =   "Aναζήτηση"
      Top             =   165
      Width           =   1575
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
      Left            =   8400
      TabIndex        =   15
      ToolTipText     =   "Περιγραφή $ Κωδικός  Συνδυασμένη αναζήτηση π.χ. LL$456 η περιγραφή αρχίζει από LL και ο κωδικός αρχίζει με 456"
      Top             =   645
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   7
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
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   2955
      TabIndex        =   10
      Top             =   2610
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   1035
      TabIndex        =   8
      Top             =   2625
      Width           =   945
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
      TabIndex        =   2
      Top             =   600
      Width           =   3815
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
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4650
      Left            =   0
      TabIndex        =   33
      Top             =   3960
      Width           =   6855
      _ExtentX        =   12086
      _ExtentY        =   8192
      _Version        =   393216
      Tabs            =   9
      Tab             =   4
      TabHeight       =   529
      TabCaption(0)   =   "0.Προμηθ./ελαχ.στόκ"
      TabPicture(0)   =   "APOT2iliadis.frx":0463
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1(18)"
      Tab(0).Control(1)=   "Text1(17)"
      Tab(0).Control(2)=   "exoda"
      Tab(0).Control(3)=   "Text1(15)"
      Tab(0).Control(4)=   "Text1(10)"
      Tab(0).Control(5)=   "Text1(7)"
      Tab(0).Control(6)=   "Text1(8)"
      Tab(0).Control(7)=   "Text1(9)"
      Tab(0).Control(8)=   "Label1(29)"
      Tab(0).Control(9)=   "Label1(28)"
      Tab(0).Control(10)=   "Label3"
      Tab(0).Control(11)=   "Label1(24)"
      Tab(0).Control(12)=   "Label2"
      Tab(0).Control(13)=   "Label1(18)"
      Tab(0).Control(14)=   "Label1(17)"
      Tab(0).Control(15)=   "Label1(16)"
      Tab(0).Control(16)=   "Label1(15)"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "1.Γραφική παράσταση"
      TabPicture(1)   =   "APOT2iliadis.frx":047F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "graph1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "2.Παράμετροι Λογιστ."
      TabPicture(2)   =   "APOT2iliadis.frx":049B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "timokat"
      Tab(2).Control(1)=   "DataGrid2"
      Tab(2).Control(2)=   "Text1(12)"
      Tab(2).Control(3)=   "Text1(11)"
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(5)=   "Label1(20)"
      Tab(2).Control(6)=   "Label1(19)"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "3.Κινήσεις κατα μήνα"
      TabPicture(3)   =   "APOT2iliadis.frx":04B7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSFlexGrid1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "4.Υπόλοιπα κατά Αποθ."
      TabPicture(4)   =   "APOT2iliadis.frx":04D3
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label6"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "MSHFlexGrid1"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "5.BARCODES"
      TabPicture(5)   =   "APOT2iliadis.frx":04EF
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Adodc2"
      Tab(5).Control(1)=   "Command11"
      Tab(5).Control(2)=   "Command10"
      Tab(5).Control(3)=   "Text3"
      Tab(5).Control(4)=   "DataGrid1"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "6.Κίνηση περισυνή"
      TabPicture(6)   =   "APOT2iliadis.frx":050B
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "MSFlexGrid2"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "7.Γραφημα περισυνό"
      TabPicture(7)   =   "APOT2iliadis.frx":0527
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "graph2"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Πεδία Χρήστη"
      TabPicture(8)   =   "APOT2iliadis.frx":0543
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
      Tab(8).Control(12)=   "HM1"
      Tab(8).Control(13)=   "HM2"
      Tab(8).Control(14)=   "HM3"
      Tab(8).Control(15)=   "CH1"
      Tab(8).Control(16)=   "CH2"
      Tab(8).Control(17)=   "CH3"
      Tab(8).Control(18)=   "NUM1"
      Tab(8).Control(19)=   "NUM2"
      Tab(8).Control(20)=   "NUM3"
      Tab(8).Control(21)=   "ch6"
      Tab(8).Control(22)=   "ch5"
      Tab(8).Control(23)=   "ch4"
      Tab(8).ControlCount=   24
      Begin VB.TextBox ch4 
         Height          =   330
         Left            =   -70515
         TabIndex        =   110
         Text            =   " "
         Top             =   1935
         Width           =   2010
      End
      Begin VB.TextBox ch5 
         Height          =   330
         Left            =   -70515
         TabIndex        =   109
         Top             =   2730
         Width           =   2010
      End
      Begin VB.TextBox ch6 
         Height          =   330
         Left            =   -70515
         TabIndex        =   108
         Top             =   3480
         Width           =   2010
      End
      Begin VB.TextBox NUM3 
         Height          =   330
         Left            =   -72825
         TabIndex        =   98
         Top             =   4215
         Width           =   2250
      End
      Begin VB.TextBox NUM2 
         Height          =   330
         Left            =   -72825
         TabIndex        =   97
         Top             =   3855
         Width           =   2250
      End
      Begin VB.TextBox NUM1 
         Height          =   330
         Left            =   -72825
         TabIndex        =   96
         Top             =   3465
         Width           =   2250
      End
      Begin VB.TextBox CH3 
         Height          =   330
         Left            =   -72825
         TabIndex        =   95
         Top             =   3105
         Width           =   2250
      End
      Begin VB.TextBox CH2 
         Height          =   330
         Left            =   -72825
         TabIndex        =   94
         Top             =   2685
         Width           =   2250
      End
      Begin VB.TextBox CH1 
         Height          =   330
         Left            =   -72825
         TabIndex        =   93
         Text            =   " "
         Top             =   2295
         Width           =   2250
      End
      Begin VB.TextBox HM3 
         Height          =   330
         Left            =   -72825
         TabIndex        =   92
         Top             =   1935
         Width           =   2250
      End
      Begin VB.TextBox HM2 
         Height          =   330
         Left            =   -72825
         TabIndex        =   91
         Top             =   1530
         Width           =   2250
      End
      Begin VB.TextBox HM1 
         Height          =   330
         Left            =   -72825
         TabIndex        =   90
         Top             =   1095
         Width           =   2250
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   18
         Left            =   -72720
         TabIndex        =   85
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
         TabIndex        =   84
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox exoda 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72750
         TabIndex        =   72
         Top             =   4320
         Width           =   1830
      End
      Begin MSAdodcLib.Adodc timokat 
         Height          =   330
         Left            =   -70065
         Top             =   2445
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
         Bindings        =   "APOT2iliadis.frx":055F
         Height          =   1785
         Left            =   -74505
         TabIndex        =   69
         Top             =   2940
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   3154
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
         TabIndex        =   67
         Top             =   3000
         Width           =   2100
      End
      Begin VB.CommandButton Command10 
         Caption         =   "ΝΕΟ BARCODE"
         Height          =   285
         Left            =   -70575
         TabIndex        =   66
         Top             =   2310
         Width           =   2100
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   -70590
         TabIndex        =   65
         Top             =   1725
         Width           =   2070
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "APOT2iliadis.frx":0575
         Height          =   2940
         Left            =   -73965
         TabIndex        =   64
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
         TabIndex        =   60
         Top             =   2910
         Width           =   1815
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Bindings        =   "APOT2iliadis.frx":058A
         Height          =   3015
         Left            =   240
         TabIndex        =   50
         Top             =   1500
         Width           =   6135
         _ExtentX        =   10816
         _ExtentY        =   5313
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSChart20Lib.MSChart graph1 
         Height          =   3495
         Left            =   -74880
         OleObjectBlob   =   "APOT2iliadis.frx":059F
         TabIndex        =   48
         Top             =   1020
         Width           =   6375
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   -72360
         TabIndex        =   45
         Top             =   1860
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   -72360
         TabIndex        =   44
         Top             =   2340
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   " "
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   -72720
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   1980
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   35
         Top             =   1020
         Width           =   5295
         _ExtentX        =   9335
         _ExtentY        =   6583
         _Version        =   393216
         Rows            =   13
         Cols            =   5
         GridLines       =   2
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3450
         Left            =   -74340
         TabIndex        =   81
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
         OleObjectBlob   =   "APOT2iliadis.frx":23A8
         TabIndex        =   82
         Top             =   1050
         Width           =   6375
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   1800
         TabIndex        =   121
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Δεσμευμένα"
         Height          =   255
         Left            =   240
         TabIndex        =   120
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Περιγραφή 2η είδους"
         Height          =   390
         Left            =   -74055
         TabIndex        =   115
         Top             =   1290
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
         TabIndex        =   113
         Top             =   3195
         Width           =   1980
      End
      Begin VB.Label Lab5 
         Height          =   330
         Left            =   -70500
         TabIndex        =   112
         Top             =   2310
         Width           =   1995
      End
      Begin VB.Label Lab4 
         Height          =   330
         Left            =   -70470
         TabIndex        =   111
         Top             =   1530
         Width           =   1950
      End
      Begin VB.Label LNUM3 
         Height          =   285
         Left            =   -74970
         TabIndex        =   107
         Top             =   4215
         Width           =   2205
      End
      Begin VB.Label LNUM2 
         Height          =   285
         Left            =   -74925
         TabIndex        =   106
         Top             =   3855
         Width           =   2100
      End
      Begin VB.Label LNUM1 
         Height          =   255
         Left            =   -74955
         TabIndex        =   105
         Top             =   3480
         Width           =   2130
      End
      Begin VB.Label LAB3 
         Height          =   300
         Left            =   -74955
         TabIndex        =   104
         Top             =   3120
         Width           =   2100
      End
      Begin VB.Label Lab2 
         Height          =   285
         Left            =   -74940
         TabIndex        =   103
         Top             =   2715
         Width           =   2070
      End
      Begin VB.Label LaB1 
         Height          =   315
         Left            =   -74955
         TabIndex        =   102
         Top             =   2295
         Width           =   2130
      End
      Begin VB.Label lhm3 
         Height          =   285
         Left            =   -74940
         TabIndex        =   101
         Top             =   1950
         Width           =   2130
      End
      Begin VB.Label lhm2 
         Height          =   270
         Left            =   -74940
         TabIndex        =   100
         Top             =   1545
         Width           =   2115
      End
      Begin VB.Label lhm1 
         Height          =   315
         Left            =   -74955
         TabIndex        =   99
         Top             =   1110
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Συσκευασία"
         Height          =   375
         Index           =   29
         Left            =   -74280
         TabIndex        =   87
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Υποσυσκευασία"
         Height          =   375
         Index           =   28
         Left            =   -74280
         TabIndex        =   86
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Εξοδα"
         Height          =   210
         Left            =   -74205
         TabIndex        =   73
         Top             =   4320
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus σε κάρτα"
         Height          =   375
         Index           =   24
         Left            =   -74280
         TabIndex        =   61
         Top             =   2910
         Width           =   1455
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   -74400
         TabIndex        =   57
         Top             =   3960
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδ.Λογ.Αγορών"
         Height          =   375
         Index           =   20
         Left            =   -73920
         TabIndex        =   47
         Top             =   1860
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Κωδ.Λογ.Πωλήσεων"
         Height          =   375
         Index           =   19
         Left            =   -73920
         TabIndex        =   46
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ράφι-Θέση"
         Height          =   375
         Index           =   18
         Left            =   -74280
         TabIndex        =   43
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Προμηθευτής"
         Height          =   375
         Index           =   17
         Left            =   -74280
         TabIndex        =   42
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Επιθυμητό Υπόλοιπο"
         Height          =   375
         Index           =   16
         Left            =   -74280
         TabIndex        =   41
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ελάχιστο στόκ"
         Height          =   375
         Index           =   15
         Left            =   -74280
         TabIndex        =   40
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
         TabIndex        =   34
         Top             =   1320
         Width           =   2655
      End
   End
   Begin MSFlexGridLib.MSFlexGrid times 
      Height          =   1335
      Left            =   -15
      TabIndex        =   125
      Top             =   2145
      Width           =   4935
      _ExtentX        =   8700
      _ExtentY        =   2350
      _Version        =   393216
      Rows            =   3
      Cols            =   5
      RowHeightMin    =   400
      ScrollBars      =   0
   End
   Begin VB.Shape plaisio 
      BorderColor     =   &H8000000E&
      Height          =   6090
      Left            =   6900
      Top             =   75
      Width           =   7470
   End
   Begin VB.Label flag 
      Caption         =   "flag"
      Height          =   255
      Left            =   240
      TabIndex        =   123
      Top             =   8760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label proteinX 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4935
      TabIndex        =   117
      Top             =   2265
      Width           =   795
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode4 
      Height          =   900
      Left            =   -600
      TabIndex        =   116
      Top             =   9120
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
      TabIndex        =   83
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
      TabIndex        =   80
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
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   27
      Left            =   3720
      TabIndex        =   79
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Left            =   4935
      TabIndex        =   76
      Top             =   2880
      Width           =   1575
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
      Left            =   3630
      TabIndex        =   71
      Top             =   1005
      Width           =   1695
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode1 
      Height          =   1065
      Left            =   0
      TabIndex        =   70
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
      TabIndex        =   59
      Top             =   3075
      Width           =   720
   End
   Begin VB.Shape Shape1 
      Height          =   915
      Index           =   1
      Left            =   10020
      Top             =   1560
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000014&
      Height          =   255
      Index           =   22
      Left            =   6960
      TabIndex        =   49
      Top             =   2085
      Width           =   1575
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
      TabIndex        =   32
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   11
      Left            =   0
      TabIndex        =   31
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000014&
      Height          =   255
      Index           =   10
      Left            =   6960
      TabIndex        =   30
      Top             =   1725
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000014&
      Height          =   255
      Index           =   9
      Left            =   6960
      TabIndex        =   29
      Top             =   1245
      Width           =   1335
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
      Height          =   255
      Index           =   8
      Left            =   6960
      TabIndex        =   26
      Top             =   285
      Width           =   1335
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
      Height          =   255
      Index           =   7
      Left            =   6960
      TabIndex        =   25
      Top             =   765
      Width           =   1335
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
      Height          =   375
      Index           =   4
      Left            =   3465
      TabIndex        =   22
      Top             =   1590
      Width           =   1695
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
      TabIndex        =   21
      Top             =   1800
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
      TabIndex        =   20
      Top             =   1440
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
      TabIndex        =   19
      Top             =   600
      Width           =   1455
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
Dim f_tab 'TO TAB CONTROL ΣΕ ΠΟΙΟ  ΣΤΑΜΑΤΑΕΙ
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
Dim f_TIMOKAT
Dim F_add_barc As Integer
Dim F_YBARCODE, F_XBARCODE
Dim F_CHANGETIMH

Dim f_YPOL_MHN As Integer

Dim FGDB2 As New ADODB.Connection  ' fanaropoylos singular
Dim fConnect2 As String   ' fanaropoylos singular

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

'Dim f_Refresh As Boolean


Option Explicit

Dim F_ArPerEtik ' αριστερο περιθωριο ετικεττας



Dim F_2ArPerEtik '2o αριστερο περιθωριο ετικεττας
Dim F_PIChEIGHT ' Val(FindParametroi("APOT2", "F_PIChEIGHT", "2535", "Υψος εικόνας")) 'posa psifia tha exei h kathe seira
Dim F_PICwidth   ' = Val(FindParametroi("APOT2", "F_PICwidth", "2535", "Υψος εικόνας")) 'posa psifia tha exei h kathe seira
Dim F_sygx ' SYGXONEYSH KODIKON=1


Dim F_EIKONA  ' 1=KVDIKOS   2=ERG  3=TIPOTA DEN EXEI EIKONES



Dim F_KOD As Integer, F_ONO As Integer, F_ERG As Integer, F_XTI As Integer

Dim f_FontName   '"Arial"
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

Dim F_UPPER ' ΜΕΤΑΤΡΕΠΕΙ ΣΕ ΚΕΦΑΛΑΙΑ


Dim f_num1 As String
Dim f_num2 As String
Dim f_num3 As String









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

Private Sub Combo4_GotFocus()
      Call SendMessage(Combo4.hwnd, CB_SHOWDROPDOWN, 1, 0)
End Sub

Private Sub Combo4_LostFocus()
' YPOOIKOGENEIES
Dim R As New ADODB.Recordset
If Len(Combo4.Text) > 0 Then
   
   BALE_TSONTES "3", Left(Combo4.Text, 2)
   R.Open "SELECT *FROM PINAKES WHERE TYPOS=3 AND TIMH=" + Left(Combo4.Text, 2) + "ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
   
   
   Dim n As Integer
    n = Combo3.ListIndex
    Combo3.Clear


Do While Not R.EOF
  If R("typos") = 3 Then
     Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
  Else
     Exit Do
  End If
  R.MoveNext
Loop
R.Close

 On Error Resume Next
 ' Call SendMessage(Combo3.hWnd, CB_SHOWDROPDOWN, 1, 0)
If Combo3.ListCount < n Then
  
   Combo3.ListIndex = 0
Else
  Combo3.ListIndex = n   ' - 1 Else Combo3.ListIndex = 0
End If

End If


End Sub

Private Sub Command1_Click()
'exodos
Unload Me


If gApoMenu = True Then
   SendKeys "%"
   SendKeys "{RIGHT}+{DOWN}+{DOWN}"
End If



End Sub

Private Sub Command10_Click()
   
Dim a

Dim R0 As New ADODB.Recordset
 a = SSTab1.Tab
 
If a = 5 And F_EXO_XROMATA = 1 Then
   UGridEdit.Update.Visible = True
   
   UGridEdit.Text1.Visible = True
   UGridEdit.kod.Caption = Text1(0).Text
   UGridEdit.ono.Caption = Text1(1).Text
   UGridEdit.CH1.Caption = CH1.Text
   UGridEdit.LTI5.Caption = Text1(13).Text
   UGridEdit.XTI.Caption = Text1(2).Text
   
   
   
   
   
   
   
   Dim DUM
   DUM = ADD_FIELD("BARCODES", "PROELEYSH", "CHAR(15)")
   
   
   
     Dim ColS As Integer, ROWS As Integer
  ColS = 0: ROWS = 0
 
 
 'ΦΟΡΤΩΝΩ ΤΙΣ ΕΠΙΚΕΦΑΛΙΔΕΣ ΤΟΥ GRID ΧΡΩΜΑΤΩΝ
 Dim MEGETHOS As String, XROMA As String
  XROMA = Format(29 + (2 * Val(Left(Combo5.Text, 2)) - 1), "00")
  MEGETHOS = Format(29 + (2 * Val(Left(Combo5.Text, 2))), "00")
 
 
 
 'ΦΟΡΤΩΝΩ ΤΙΣ ΕΠΙΚΕΦΑΛΙΔΕΣ ΤΟΥ GRID ΧΡΩΜΑΤΩΝ
 R0.Open "SELECT *FROM PINAKES WHERE TYPOS=" + XROMA + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
 Do While Not R0.EOF
     UGridEdit.Grid1.TextMatrix(R0("AYJON"), 0) = R0("PERIGRAFH")
     ROWS = ROWS + 1
     R0.MoveNext
 Loop
 R0.Close
  
  
 'ΦΟΡΤΩΝΩ ΤΙΣ ΕΠΙΚΕΦΑΛΙΔΕΣ ΤΟΥ GRID ΜΕΓΕΘΩΝ
 
 
    
 R0.Open "SELECT *FROM PINAKES WHERE TYPOS=" + MEGETHOS + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
 Do While Not R0.EOF
     UGridEdit.Grid1.TextMatrix(0, R0("AYJON")) = R0("PERIGRAFH")
     ColS = ColS + 1
     R0.MoveNext
 Loop
 R0.Close
  
  UGridEdit.Grid1.ROWS = ROWS + 1
  UGridEdit.Grid1.ColS = ColS + 1
  
  
 'ΤΑ ΥΠΟΛΟΙΠΑ ANA XROMA ΑΠΟ ΤΟ MSFLEXGRID1 TA DEIXNO ΣΤΟ  GRID1 TOY UGRIDEDIT
 Dim k As Integer
 UGridEdit.Label1.Caption = "NoPaint" ' flag gia na mhn gemisei to grid1 apo to gm_str(i)
 On Error Resume Next
 Dim X As Long, Y As Long
 Dim n2 As Long
 Dim N3 As Long
 Dim n4
  For k = 1 To MSHFlexGrid1.ROWS - 1
     
     For n2 = 1 To UGridEdit.Grid1.ROWS - 1  ' ΨΑΧΝΩ ΣΤΙΣ ΣΕΙΡΕΣ ΝΑ ΒΡΩ ΤΟ ΧΡΩΜΑ
          If Left(MSHFlexGrid1.TextMatrix(k, 1), 3) = Left(UGridEdit.Grid1.TextMatrix(n2, 0), 3) Then
              X = n2
              Exit For
          End If
     Next
          
    'ΨΑΧΝΩ ΣΤΙΣ ΣΤΗΛΕΣ ΝΑ ΒΡΩ ΤΟ ΜΕΓΕΘΟΣ
    N3 = InStr(MSHFlexGrid1.TextMatrix(k, 1), "*")
    n4 = max(1, N3 - 4)
     For n2 = 1 To UGridEdit.Grid1.ColS - 1  ' ΨΑΧΝΩ ΣΤΙΣ ΣΕΙΡΕΣ ΝΑ ΒΡΩ ΤΟ ΧΡΩΜΑ
          If mID(MSHFlexGrid1.TextMatrix(k, 1), 4, n4) = Left(UGridEdit.Grid1.TextMatrix(0, n2), n4) Then
              Y = n2
              Exit For
          End If
     Next
     
     UGridEdit.Grid1.TextMatrix(X, Y) = Val(UGridEdit.Grid1.TextMatrix(X, Y)) + Val(MSHFlexGrid1.TextMatrix(k, 6))
  
  
  Next
  
 
  
  
  
  
  
  
  
  
  
  
  
  
   UGridEdit.SHOW 1
   Exit Sub
End If
   
   
   
   
   
   
   
   
   
   On Error Resume Next
  Dim d
  'On Error GoTo 0
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
Dim fores As Long, k As Integer
Dim X As Printer
Dim barc(200)
Dim mONOMA, mBarcode, mERG, mLTI5
Dim R As New ADODB.Recordset
Dim RD As DAO.Recordset
mONOMA = Text1(1).Text 'ONOMA
mBarcode = Text1(0).Text '  ' BARCODE
mERG = Text1(6).Text ' erg
mLTI5 = Format(Val(Text1(13).Text), "###0.00") ' LTI5




Dim Z As Integer

Z = InputBox("ΠΟΣΕΣ ΕΤΙΚΕΤΤΕΣ? ")

For k = 1 To Z
   If Len(Dir("C:\MPOYGZEB.EXE", vbNormal)) > 0 Then
      GoSub ZEBRAmonoBARCODE
   Else
      GoSub SATO
   End If
   
Next k

Printer.EndDoc

Exit Sub












SATO:
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
  k = 0
  Do While Not RD.EOF
     k = k + 1
     If k > 80 Then Exit Do
     If IsNull(RD("BARCODE")) Then
        barc(k) = " " 'Grid1.TextMatrix(K, f_k) = " "
     Else
        barc(k) = RD("barcode") 'Grid1.TextMatrix(K, f_k) = R("BARCODE")
     End If
     RD.MoveNext
  Loop
  RD.Close
  db.Close
  apoSCANNER = k
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



Z = -5 ' InputBox("METATOPISH")

For k = 1 To mFORES




Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992

Printer.FONTSIZE = 12
Printer.FontBold = True
Dim a As Integer
  a = InStr(1, mONOMA, ";")
  If a = 0 Then
    Printer.Print Tab(5); Left(mONOMA, 17) 'Text1(1).Text 'ONOMA
  Else
    Printer.Print Tab(5); Left(mONOMA, a - 1) 'Text1(1).Text 'ONOMA
    Printer.Print Tab(5); mID$(mONOMA, a + 1, 17) 'Text1(1).Text 'ONOMA
  End If

Printer.EndDoc
Next k




Next ll
Return





ZEBRAmonoBARCODE:

On Error Resume Next
For Each X In Printers
    If InStr(X.DeviceName, "MIKR") > 0 Then
        Set Printer = X
        Exit For
    End If
Next





  Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName454967325476-5200016000239
  
  
    Printer.FONTSIZE = 8
    
  
  
  
  
  Printer.CurrentY = 0
    Printer.CurrentX = 0
    'Printer.Print Left(Text1(6).Text, 12)
    

     BarCode2.DataToEncode = Text1(0).Text
     
     BarCode2.Code128CharSet = Set_A
    
    If Len(DataGrid1.Text) = 12 Then
         BarCode2.SymbologyID = UPCA
    Else
         BarCode2.SymbologyID = EAN13
    End If
    Printer.PaintPicture BarCode2.Picture, 200, -400
    Printer.EndDoc
    
       
   
Return











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


Private Sub Command13_Click()
    
Dim X
Dim rec As ADODB.Recordset
Dim ANS
Dim k
Dim fores
Dim FF, ARX
       
  
Dim REC2 As New ADODB.Recordset



Text4.Visible = True

  
  
  If Len(Dir("C:\FANAR.EXE", vbNormal)) > 0 Then
   
   For Each X In Printers
        If InStr(X.DeviceName, "TLP2844") > 0 Then
          ' Set printer as system default.
        
          Set Printer = X
          ' Stop looking for a printer
          Exit For
        End If
    Next
    
    GoSub FIND_N_NUMBER
    

    ARX = Text4.Text
    
    FF = InputBox(ARX, , ARX)
    
    Text4.Text = Right("000000000000" + LTrim(Str(FF)), 12)
    
    '   5212345678935
    
    BarCode2.DataToEncode = FF
       
          fores = Val(InputBox("Πόσα κομμάτια  "))
       
       
    For k = 1 To fores
         Printer.CurrentY = 400    '000000020404
         Printer.CurrentX = 600 + F_ArPerEtik * 100
         BarCode2.SymbologyID = EAN13
         Printer.PaintPicture BarCode2.Picture, F_XBARCODE, 100, , 1300
         Printer.EndDoc
    Next k
    
     If FF <> ARX Then

       ANS = MsgBox("Να αποθηκευτεί ο αριθμός " + FF, vbYesNo)
       If ANS = vbYes Then
          GoSub update_N_NUMBER
       End If
     End If
       
    
    
  ElseIf Len(Dir("C:\mercvb\zebra.txt", vbNormal)) > 0 Then
   
   For Each X In Printers
        If InStr(X.DeviceName, "TLP2844") > 0 Then
          ' Set printer as system default.
        
          Set Printer = X
          ' Stop looking for a printer
          Exit For
        End If
    Next
    
    GoSub FIND_N_NUMBER
    
    
    ARX = Text4.Text
    
    FF = InputBox(ARX, , ARX)
    
    Text4.Text = Right("000000000000" + LTrim(Str(FF)), 12)
    
    '   5212345678935
    
    BarCode2.DataToEncode = FF
       
    
       
       
       
      'For k = 1 To fores
         Printer.CurrentY = 400    '000000020404
         Printer.CurrentX = 600 + F_ArPerEtik * 100
         
     
   BarCode2.SymbologyID = EAN13
   Printer.PaintPicture BarCode2.Picture, F_XBARCODE, F_YBARCODE, , 1300
  
     Printer.EndDoc
     
     If FF <> ARX Then
     
       ANS = MsgBox("Να αποθηκευτεί ο αριθμός " + FF, vbYesNo)
       If ANS = vbYes Then
          GoSub update_N_NUMBER
       End If
     End If
       
  Else
        
          fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
          Dim PARAM
     
       
        For Each X In Printers
        If InStr(X.DeviceName, "SATO") > 0 Then
          ' Set printer as system default.
        
          Set Printer = X
          ' Stop looking for a printer
          Exit For
        End If
    Next
    
       
       
       
       
       BarCode2.DataToEncode = DataGrid1.Text
       
       For k = 1 To fores
        Printer.CurrentY = 400
        Printer.CurrentX = 600 + F_ArPerEtik * 100
        Printer.FONTSIZE = 12
        Printer.Print Text1(0).Text
        Printer.EndDoc
       
       Next
  End If
  

Exit Sub

    

FIND_N_NUMBER:
'************************


'rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
REC2.MoveFirst
REC2.Move 10
  'rec.Edit
  If IsNull(REC2("epel")) Then
      REC2("epel") = 1
  Else
      REC2("epel") = REC2("epel") + 1
  End If
  REC2.Update
  
   Text4.Text = Right("000000000000" + LTrim(Str(REC2("EPEL"))), 12)
  
Return



update_N_NUMBER:
'************************

REC2.Close
REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
REC2.MoveFirst
REC2.Move 10
REC2("epel") = Val(FF)
REC2.Update
  
   Text4.Text = Right("000000000000" + LTrim(Str(REC2("EPEL"))), 12)
  
Return


    
    
    
    

End Sub

Private Sub Command14_Click()
  CommonDialog1.ShowOpen
  FileCopy CommonDialog1.FILENAME, "c:\mercvb\images\" + Text1(0)
  
  
End Sub

Private Sub Command15_Click()


Dim STR_EKT As String
STR_EKT = data2.Recordset("kod") + " " + Format(data2.Recordset("hm2"), "DD/MM/YYYY") + " "

STR_EKT = STR_EKT + Left(data2.Recordset("ch2") + Space(25), 25) + " " + Left(data2.Recordset("ch4") + Space(20), 20)

STR_EKT = STR_EKT + " " + Left(data2.Recordset("ch3") + Space(22), 22) + _
" " + Left(data2.Recordset("ono") + Space(25), 25) + _
" " + Left(data2.Recordset("mon") + Space(4), 3) + _
" " + Format(1, "#0") + _
" " + Left(data2.Recordset("ch1") + Space(15), 15)








'Printer.FontName = "Courier New"
'Printer.FontSize = 8

Open "LPT1" For Output As 1
 Print #1, Chr(15) + to437(STR_EKT)
Close 1












End Sub

Private Sub Command16_Click()

Dim k As Integer, DUM As Integer, db As Database, R As New ADODB.Recordset
Dim X As String
Dim mlabel As String
Dim SQL3 As String
Dim RECS As Long

      

  If Command16.Caption = "+" Then
      

       Command16.Caption = "Ενημέρωση"
       For k = 1 To 3
          Text1(k).Enabled = True
          Text1(k).ForeColor = vbBlack
       Next
       
       For k = 6 To 7
          Text1(k).Enabled = True
          Text1(k).ForeColor = vbBlack
       Next
       For k = 11 To 13
          Text1(k).Enabled = True
          Text1(k).ForeColor = vbBlack
       Next
       
       Text1(9).Enabled = True
       Text1(5).Enabled = True
       Text1(5).ForeColor = vbBlack
       
       
       Combo1.Enabled = True
       Combo2.Enabled = True
       Combo3.Enabled = True
       
        Combo2.Text = Combo2.List(1)
        Text1(0).Enabled = True
        Text1(0).SetFocus
        
  Else
      
      'ελεγχος μην ξαναυπάρχει ο κωδικός
   'Set db = OpenDatabase(gDir, False, False, gConnect)
   R.Open "select count(*) as ar from EID where  KOD='" + Text1(0).Text + "';", _
       Gdb, adOpenDynamic, adLockOptimistic
   
   If R("ar") > 0 And Len(Text1(0).Text) > 0 Then
       MsgBox "Υπάρχει ο κωδικός"
       Exit Sub
   End If
  
  If Val(Left(Combo2.Text, 2)) = 0 Then
       MsgBox "ΦΠΑ ;"
       Exit Sub
  End If
  
  
  
  
  ' ενημέρωση μετρητή
          
       '   dum = Update_Next_Pel(x)
       On Error GoTo NoUpdate
       DATA1.Recordset.AddNew
       On Error Resume Next
       DATA1.Recordset("kod") = UCase(Text1(0))  ', "0000")
       DATA1.Recordset("ONO") = Left(Text1(1).Text, _
           DATA1.Recordset("ONO").DefinedSize)
       
       DATA1.Recordset("FPA") = Val(Left(Combo2.Text, 2))
       
       DATA1.Recordset("mon") = Combo1.Text
       
      
       
       
'       DATA2.Recordset("kodlog") = Val(left(Combo3.Text, 3)) ' Combo3.ListIndex  'Right(Combo3.Text, 3)
       
       DATA1.Recordset("aeg") = Val(Left(Combo4.Text, 3)) ' Combo3.ListIndex  'Right(Combo3.Text, 3)
        'Data1.Recordset("kodlog") = Right(Combo3.Text, 3)
       
       
       
       
'       'αν εχω προσθέσει υπο οικογένεια την καταχωρεί αυτόματα
'       If Val(Combo3.Text) = 0 Then
'           SQL3 = "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (3," + Str(Combo3.ListCount + 1) + ",'" + Combo3.Text + "'," + Left(Combo4.Text, 2) + ")"
'           Gdb.Execute SQL3
'       End If
       
       
      DATA1.Recordset("kodlog") = Val(Left(Combo3.Text, 3))
           
       
       
       
       DATA1.Recordset("xti") = Val(Replace(Text1(2).Text, ",", "."))
       DATA1.Recordset("lti") = Val(Replace(Text1(3).Text, ",", ".")) 'Val(Text1(3).Text)
       DATA1.Recordset("lti5") = Val(Replace(Text1(10).Text, ",", ".")) ' Val(Text1(10).Text)
       DATA1.Recordset("pos_kerD") = Val(Replace(Text1(5).Text, ",", ".")) ' Val(Text1(5).Text)
       
       DATA1.Recordset("prom") = Text1(7).Text
       DATA1.Recordset("spa") = Val(Text1(9).Text)
       ' DATA1.Recordset("apos") = Val(Text1(8).Text)
       DATA1.Recordset("ues") = Text1(6).Text
       
       DATA1.Recordset("kodsynod") = Text1(12).Text
       DATA1.Recordset("kodlogag") = Text1(11).Text
       
       DATA1.Recordset("CH1") = CH1.Text
       
       DATA1.Recordset("ERG") = Text1(4).Text 'ΕΡΓ
       DATA1.Recordset("KODERG") = Text1(13).Text ' ΚΩΔ.ΠΡΟΜΗΘΕΥΤΗ
       DATA1.Recordset("MEMO") = Text14.Text
       DATA1.Recordset("KATHGORIA") = Val(Left(Combo5.Text, 3)) ' Combo3.ListIndex  'Right(Combo3.Text, 3)
 
        
        On Error GoTo NoUpdate
       DATA1.Recordset.Update
       On Error Resume Next
       
  
       
  
  
  If F_add_barc = 1 Then   ' file("C:\LAGEURO\MPOYG.TXT") Or
      'ΤΟ ΒΑRCODE ΕΙΝΑΙ Ο ΚΩΔΙΚΟΣ
         Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + UCase(Text1(0).Text) + "','" + UCase(Text1(0).Text) + "')"
  ElseIf F_add_barc = 2 Then
        'ΤΟ ΒΑRCODE ΕΙΝΑΙ TO BOHΘ.ΚΛΕΙΔΙ
        If Val(Text1(4).Text) = 0 Then
           GoSub FIND_N_NUMBER
        End If
RECS = 0
  Do While RECS = 0
       
        Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + UCase(Text1(0).Text) + "','" + UCase(Text1(4).Text) + "')", RECS
       
       If RECS = 0 Then GoSub FIND_N_NUMBER
  Loop
        
        
  End If
       
  f_TIMOKAT = Val(FindParametroi("PELAT2", "F_TIMOKAT", "0", "Aριθμός Τιμοκαταλόγων"))
       If f_TIMOKAT > 0 Then
         For k = 1 To f_TIMOKAT
              Gdb.Execute "INSERT INTO TIMOKAT SELECT KOD,0 AS EKPT, " + Str(k) + " AS TIMOK,ONO FROM EID WHERE KOD='" + UCase(Text1(0)) + "'"
         Next
       End If
       If FLAG.Caption = "1" Then  ' ΕΡΧΕΤΑΙ ΑΠΟ ΤΑ ΤΙΜΟΛΟΓΙΑ
          MDIForm1.StatusBar1.Panels(6).Text = Text1(0).Text
          Unload Me
       End If
       
       MHNYMA2.Timer1.Interval = 1000
       MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " αποθηκεύτηκε με επιτυχία"
       MHNYMA2.SHOW 1
       
       Command2.Caption = "Νέα Εγγραφή"
End If
  ' Command2.SetFocus
  
  
  
  
  
Exit Sub

NoUpdate:
MsgBox "Δεν έγινε η αποθήκευση"
Resume Next
Exit Sub


FIND_N_NUMBER:
'************************
Dim rec As ADODB.Recordset

'rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
Dim REC2 As New ADODB.Recordset
REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
REC2.MoveFirst
REC2.Move 10
  'rec.Edit
  If IsNull(REC2("epel")) Then
      REC2("epel") = 1
  Else
      REC2("epel") = REC2("epel") + 1
  End If
  REC2.Update
  
   Text1(4).Text = Right("00000000" + LTrim(Str(REC2("EPEL"))), 8)
   Gdb.Execute "UPDATE EID SET ERG='" + Text1(4).Text + "' WHERE KOD='" + Text1(0).Text + "'"
Return
















End Sub

Private Sub Command2_Click()
'diortosi
Dim k As Integer, DUM As Integer, db As Database, R As Recordset
Dim X As String
       
      

  If Command2.Caption = "Διόρθωση" Then
      

       Command2.Caption = "Ενημέρωση"
       For k = 1 To 3
          Text1(k).Enabled = True
          Text1(k).ForeColor = vbBlack
       Next
       
          Text1(5).Enabled = True
          Text1(5).ForeColor = vbBlack
          
          Text1(6).Enabled = True
          Text1(6).ForeColor = vbBlack
       
       For k = 7 To 13
          Text1(k).Enabled = True
          Text1(k).ForeColor = vbBlack
       Next
          Text14.Enabled = True
       
          Text14.ForeColor = vbBlack
       
        Combo1.Enabled = True
        Combo2.Enabled = True
        Combo3.Enabled = True
        Combo4.Enabled = True
        Combo5.Enabled = True
        Text1(14).Enabled = True
       
       
       
       
       
       Text1(1).SetFocus
       
  Else
       
       'Data2.Recordset.Edit
   
'       Data1.Recordset("kod") = Format(Text1(0), "0000")
       
    On Error Resume Next
    
       data2.Recordset("ONO") = Left(Text1(1).Text, data2.Recordset("ONO").DefinedSize)
       
       data2.Recordset("ONO2") = Left(ONO2.Text, data2.Recordset("ONO2").DefinedSize)
       
       
       data2.Recordset("FPA") = Val(Left(Combo2.Text, 2))
       
       data2.Recordset("mon") = Combo1.Text
       
       'YPOOIKOGENEIA
       data2.Recordset("kodlog") = Left(Combo3.Text, 3) ' Combo3.ListIndex  'Right(Combo3.Text, 3)
       
       
       'OIKOGENEIA
       data2.Recordset("aeg") = Val(Left(Combo4.Text, 3)) ' Combo3.ListIndex  'Right(Combo3.Text, 3)
       
       
       data2.Recordset("KATHGORIA") = Val(Left(Combo5.Text, 3)) ' Combo3.ListIndex  'Right(Combo3.Text, 3)
       
       
       data2.Recordset("kodSYNOD") = Text1(11).Text     'Right(Combo3.Text, 3)
       data2.Recordset("kodlogAG") = Text1(12).Text     'Right(Combo3.Text, 3)
       
       data2.Recordset("PROM") = Text1(7).Text     'Right(Combo3.Text, 3)
       data2.Recordset("SPA") = Val(Text1(9).Text)     'Right(Combo3.Text, 3)
       
       data2.Recordset("PROMHU") = Val(Text1(15).Text)     ' BONUS
       
       data2.Recordset("UES") = Text1(10).Text     'Right(Combo3.Text, 3)
       data2.Recordset("EPIUYP") = Val(Text1(8).Text)     'Right(Combo3.Text, 3)
       
        data2.Recordset("koderg") = Text1(16).Text
       
       
       
       data2.Recordset("SYSKMAX") = Text1(18).Text 'SYSKEYASIA
       
       data2.Recordset("SYSKMIN") = Text1(17).Text ' YPOSYSKEYASIA
       
       
       data2.Recordset("pos_kerd2") = Val(Replace(Text1(5).Text, ",", "."))
       data2.Recordset("pos_kerd") = Val(Replace(Text1(14).Text, ",", ".")) ' Val(Text1(14).Text)
       
       data2.Recordset("xti") = Val(Replace(Text1(2).Text, ",", ".")) ' Val(Text1(2).Text)
       data2.Recordset("lti") = Val(Replace(Text1(3).Text, ",", ".")) 'Val(Text1(3).Text)
       data2.Recordset("lti5") = Val(Replace(Text1(13).Text, ",", ".")) ' Val(Text1(13).Text)
       data2.Recordset("ERG") = Text1(6).Text
       data2.Recordset("MEMO") = Text14.Text
       
       
       
    data2.Recordset("NUM1") = NUM1
   data2.Recordset("NUM2") = NUM2
   data2.Recordset("NUM3") = NUM3
  
   data2.Recordset("CH1") = Trim(CH1)
   data2.Recordset("CH2") = Trim(CH2)
   data2.Recordset("CH3") = Trim(CH3)
   
    data2.Recordset("CH4") = Trim(ch4)
   data2.Recordset("CH5") = Trim(ch5)
   data2.Recordset("CH6") = Trim(ch6)
   
   

   data2.Recordset("HM1") = HM1
   data2.Recordset("HM2") = HM2
   data2.Recordset("HM3") = HM3
       
       
On Error GoTo no_update
       
       data2.Recordset.Update
       
On Error Resume Next

       
       For k = 1 To 3
          Text1(k).Enabled = False
          Text1(k).ForeColor = vbBlack
          
       Next
       
          Text1(5).Enabled = False
          Text1(6).Enabled = False
          Text1(5).ForeColor = vbBlack
       
        Combo1.Enabled = False
        Combo2.Enabled = False
        Combo3.Enabled = False
        Combo4.Enabled = False
        Combo5.Enabled = False
        Text1(14).Enabled = False
        
       For k = 7 To 13
          Text1(k).Enabled = False
          Text1(k).ForeColor = vbBlack
       Next
       
       
       
       



       
       
       
       
       
       
       
       
       
       
       
       
       
       
       MHNYMA2.Timer1.Interval = 1000
       MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " αποθηκεύτηκε με επιτυχία"
       
       MHNYMA2.SHOW 1
       
       
       
       kod.SetFocus
       Command2.Caption = "Διόρθωση"
  End If
Exit Sub

no_update:

       MHNYMA2.Timer1.Interval = 1000
       MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " ΔΕΝ αποθηκεύτηκε"
       MHNYMA2.BackColor = vbRed
       MHNYMA2.SHOW 1




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
  apot21.Label1.Caption = Str(Val(Left(Me.Combo9.Text, 2)))
  apot21.Label2.Caption = IIf(IsNull(DATA1.Recordset("apos")), "", DATA1.Recordset("apos"))
  'Me.Hide
  
  
If cXRONIES.value = vbChecked Then
  Gdb.Close
  'f_PROPERSI = Trim(mID(XRONIES.Text, 21, 20))
   Gdb.Open gConnect + ";DATABASE=" + Trim(mID(XRONIES.Text, 21, 20))
   apot21.DTPicker1.value = CDate("01/01/2000")
 ' synt = " YEAR(HME)>1900  "
End If


  
  
  
  
  
  apot21.SHOW 1
  
  
  On Error Resume Next
  
If cXRONIES.value = vbChecked Then Gdb.Close: Gdb.Open gConnect
  
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
        Gdb.Execute "DELETE FROM BARCODES WHERE KOD='" + Text1(0).Text + "'"
                  
       MHNYMA2.Timer1.Interval = 2000
       MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " διεγράφη."
       MHNYMA2.Label1.Caption = "Διαγραφή εγγραφής"
       MHNYMA2.SHOW 1





   'Call MsgBox("Διεγράφησαν " + Str(RECS) + "" _
               & vbCrLf & " εγγραφές " _
               , vbExclamation Or vbSystemModal, App.title)



Else
   Call MsgBox("Δεν ευρέθη η εγγραφή", vbCritical Or vbSystemModal, App.title)
End If



End Sub

Private Sub Command5_Click()
' anazitisi
  Dim mk As String
  Dim mO As String
  
  On Error Resume Next
  
DBGrid1.ColWidth(1) = 3300

DBGrid1.ColWidth(2) = 1250
DBGrid1.ColAlignment(2) = 7

DBGrid1.ColWidth(3) = 1250
DBGrid1.ColAlignment(3) = 7


  
If Len(ono.Text) > 0 Then
 If ypoloipo Then
                    
                    
                    If F_UPPER = 1 Then
                      mO = SameLetters(ono.Text)
                   Else
                      mO = ono.Text
                   End If

'     mO = SameLetters(ono.Text)
    If Val(XTI) > 0 Then
       data22.RecordSource = "select  KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where POS>0 and CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ]>" + Str(Val(Round(XTI, 2))) + " and ONO like '" + mO + "%';"
    Else
       data22.RecordSource = "select   KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where POS>0 and ONO like '" + mO + "%';"
    End If
    
  Else  'If ypoloipo Then
               Dim ONOM As String
               Dim KODI As String
               Dim DOL As Long
               'SYNDYASMENH ANAZHTHSH ME $  DHLADH   ΣΚΙΑ$99 ΟΛΕΣ ΤΙΣ "ΣΚΙΕΣ" ΜΕ ΚΩΔΙΚΟ ΠΟΥ ΑΡΧΙΖΕΙ ΑΠΟ "99"
               ONOM = ono.Text
               DOL = InStr(1, ONOM, "$")
               KODI = ""
               If DOL > 1 Then
                   KODI = Trim(mID(ONOM, DOL + 1, 14))
                   ONOM = Left(ONOM, DOL - 1)
                   If F_UPPER = 1 Then
                      mO = SameLetters(ONOM)
                   Else
                      mO = ONOM
                   End If
                   
                   data22.RecordSource = "select  KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where KOD LIKE '" + KODI + "%' AND ONO like '" + UCase(mO) + "%';"
               Else
               
                   If F_UPPER = 1 Then
                      mO = SameLetters(ono.Text)
                   Else
                      mO = ono.Text
                   End If

               
               
'                   mO = SameLetters(ono.Text)
                   If Val(XTI) > 0 Then
                      data22.RecordSource = "Select   KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ]>" + Str(Val(Round(XTI, 2))) + " and ONO like '" + UCase(mO) + "%';"
                   Else
                      data22.RecordSource = "select  KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where  ONO like '" + mO + "%';"
                   End If
               End If
   End If 'If ypoloipo Then
   ' f_Refresh = False
   data22.Refresh
   ' f_Refresh = True
   
   If data22.Recordset.EOF Then
      Text1(0).Text = " "
      Text1(1).Text = " "
   
   End If
   
   DBGrid1.SetFocus
   mDBGrid1_RowColChange
  Exit Sub
End If  ' If Len(ono.Text) > 0 Then


 If Len(kod) > 0 Then
  
    mk = SameLetters(kod.Text)
    If ypoloipo Then
      If Val(XTI) > 0 Then
          data22.RecordSource = "select  KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where POS>0 and CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ]>" + Str(Val(Round(XTI.Text, 2))) + " and KOD like '" + mk + "%' ESCAPE '$';"
      Else
          data22.RecordSource = "select   KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where POS>0 and KOD like '" + mk + "%' ESCAPE '$';"
      End If
    Else
       If Val(XTI) > 0 Then
           data22.RecordSource = "select  KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ]  from EID where  CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ]>" + Str(Val(Round(XTI, 2))) + " and KOD like '" + mk + "%'  ESCAPE '$';"
       Else
           data22.RecordSource = "select  KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where KOD like '" + mk + "%' ESCAPE '$' ;"
       End If
    End If
       
 '  f_Refresh = False
     data22.Refresh
 '  f_Refresh = True
       
       
       
  '     DATA2.Refresh
'       CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ].SetFocus
'       Data2.Refresh
       
       If data22.Recordset.EOF Then ' δεν βρηκε κωδικούς
          Text1(0).Text = " "
          Text1(1).Text = " "
          kod.SetFocus
       Else
          DBGrid1.SetFocus
       End If
       
  End If

If Len(erg) > 0 Then
     Dim mERG
     mERG = SameLetters(erg.Text)
     
     
 If Left(mERG, 1) = "+" Then
     data22.RecordSource = UCase("select  KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where KODERG = '" + Trim(mID(erg.Text, 2, 10)) + "'   ;")
 Else
     data22.RecordSource = UCase("select  KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where ERG like '" + mERG + "%'  ESCAPE '$' ;")
 End If
 
     data22.Refresh
     
     If data22.Recordset.EOF Then
        Text1(0).Text = " "
        Text1(1).Text = " "
     End If

     
     
     
     DBGrid1.SetFocus

End If

If Len(XTI) > 0 Then
     On Error GoTo 0
     data22.RecordSource = "select  KOD,ONO,CONVERT(decimal(10,2),LTI5) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [ΤΙΜΗ ΑΓΟΡΑΣ] from EID where KOD IN (SELECT KOD FROM BARCODES WHERE ERG like '" + XTI.Text + "%' ESCAPE '$' )"
     data22.Refresh
   
 If data22.Recordset.EOF Then
        Text1(0).Text = " "
        Text1(1).Text = " "
     End If
     
     
     
     
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
  If neos = Text1(0).Text Then
      MsgBox "Ο νέος κωδικός είναι ίδιος με τον παλιό"
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
     Dim ANS
     If F_sygx = 0 Then
          MsgBox "Υπάρχει ήδη ο κωδικός "
          Exit Sub
     End If
     
     
     
     
     
     
     ANS = MsgBox("Υπάρχει ήδη ο κωδικός " + neos + Chr(13) + " Nα συγχωνευθεί ο παλιός με τον νέο κωδικό;", vbYesNo)
     If ANS = vbYes Then
     
        Gdb.Execute "delete FROM  EID  WHERE KOD='" + Text1(0).Text + "'", RECS
        'MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ   " + Format(RECS, "#####")
     
        Gdb.Execute "UPDATE EGGTIM SET KODE='" + neos + "' WHERE KODE='" + Text1(0).Text + "'", RECS
        MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ ΣΕ ΤΙΜΟΛΟΓΙΑ  " + Format(RECS, "#####")
          
        Gdb.Execute "UPDATE BARCODES SET KOD='" + neos + "' WHERE KOD='" + Text1(0).Text + "'", RECS
        MsgBox "ΕΝΗΜΕΡΏΘΗΚΑΝ ΣΕ BARCODES  " + Format(RECS, "#####")
     
     End If
     
  End If
End Sub

Sub LAG_ETIK()
'========================etiketes ========================================================
Dim d1 As Date, d2 As Date, R As Integer, C As Integer, k As Integer, S As Single
Dim db As Database, DUM, rec As New ADODB.Recordset
Dim mlabel As String
Dim MPER As String
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

   
Dim R1 As New ADODB.Recordset
R1.Open "SELECT * FROM BARCODES WHERE KOD='" + data2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly




   If R1.EOF Then
       GoSub FIND_N_NUMBER
   Else
       If Val(R1("ERG")) = 0 Then
          GoSub FIND_N_NUMBER
       Else
          mlabel = R1("erg")
       End If
   End If
   
   MPER = data2.Recordset("ono")
   mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)
   
   TIMM = Format(data2.Recordset("lti5"), "#####.00")
   
   'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

For k = 1 To Val(mtimes)

    If True Then   ' Len(mlabel) <> 13 Then
      Printer.FontName = "128"
      Printer.FONTSIZE = 6
      Printer.Print Tab(F_ArPerEtik + 5); mlabel
      Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
      Printer.FONTSIZE = 8
      Printer.Print
'      Printer.Print
'      Printer.Print
'      Printer.Print


      Printer.Print Tab(F_ArPerEtik + 10); mlabel + "  Kωδ." + Text1(0).Text
      Printer.Print Tab(F_ArPerEtik + 10); Left(MPER, 25)
      Printer.Print Tab(F_ArPerEtik + 10); timologio
      Printer.FONTSIZE = 12
      Printer.FontBold = True
      Printer.Print Tab(F_ArPerEtik + 10); Format(Val(Trim(LTrim(TIMM))), "00.00") + " €"
      Printer.FontBold = False
      
      Printer.EndDoc
    Else
      Printer.Print Tab(F_ArPerEtik + 10); Left(MPER, 25)
     
      Printer.FontName = "EAN13" ' "128"
       Printer.Print Tab(F_ArPerEtik + 10); mlabel
        Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
      Printer.FONTSIZE = 6
      Printer.Print Tab(F_ArPerEtik + 8); data2.Recordset("kod") ' mlabel
     
      Printer.FONTSIZE = 8
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
  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + data2.Recordset("KODE") + "'"
Return


End Sub


Sub diesel_ETIK()
'========================etiketes ========================================================
Dim d1 As Date, d2 As Date, R As Integer, C As Integer, k As Integer, S As Single
Dim db As Database, DUM, rec As New ADODB.Recordset
Dim mlabel As String
Dim MPER As String
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

   
Dim R1 As New ADODB.Recordset
R1.Open "SELECT * FROM BARCODES WHERE KOD='" + data2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly




   If R1.EOF Then
       GoSub FIND_N_NUMBER
   Else
       If Val(R1("ERG")) = 0 Then
          GoSub FIND_N_NUMBER
       Else
          mlabel = R1("erg")
       End If
   End If
   
   MPER = data2.Recordset("ono")
   mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)
   
   TIMM = Format(data2.Recordset("lti5"), "#####.00")
   
   'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

For k = 1 To Val(mtimes)

    If True Then
    
    ' BarCode2.SymbologyID = CODE128
    ' BarCode2.DataToEncode = mlabel
     
     
    ' Picture2.height = F_PIChEIGHT
     'Picture2.width = F_PICwidth
 
     
     
     'Printer.PaintPicture BarCode2.Picture, F_XBARCODE, F_YBARCODE
     
      Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
      Printer.FONTSIZE = 8
      Printer.Print
'      Printer.Print
'      Printer.Print10002357

'      Printer.Print


      Printer.Print Tab(F_ArPerEtik + 10); mlabel + "  Kωδ." + Text1(0).Text
      Printer.Print Tab(F_ArPerEtik + 10); Left(MPER, 25)
      Printer.Print Tab(F_ArPerEtik + 10); timologio
      Printer.FONTSIZE = 12
      Printer.FontBold = True
      Printer.Print Tab(F_ArPerEtik + 10); Format(Val(Trim(LTrim(TIMM))), "00.00") + " €"
      Printer.FontBold = False
      
      Printer.CurrentY = F_YBARCODE
      
     Printer.FontName = "128"
     Printer.FONTSIZE = 6
     Printer.Print Tab(12); mlabel
      
      
      
      
      Printer.EndDoc
    Else
      Printer.Print Tab(F_ArPerEtik + 10); Left(MPER, 25)
      Printer.Print Tab(F_ArPerEtik + 20); mlabel
      Printer.FontName = "ean13" ' "128"
      Printer.FONTSIZE = 6
      Printer.Print Tab(F_ArPerEtik + 8); data2.Recordset("kod") ' mlabel
      Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
      Printer.FONTSIZE = 8
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
  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + data2.Recordset("KOD") + "'"
  
  Gdb.Execute "INSERT INTO BARCODES (ERG,KOD) VALUES('" + mlabel + "','" + data2.Recordset("KOD") + "')"
  
  
Return


End Sub

Sub LAG2_ETIK()
'========================etiketes ========================================================
Dim d1 As Date, d2 As Date, R As Integer, C As Integer, k As Integer, S As Single
Dim db As Database, DUM, rec As New ADODB.Recordset
Dim mlabel As String
Dim MPER As String
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

   
Dim R1 As New ADODB.Recordset
R1.Open "SELECT * FROM BARCODES WHERE KOD='" + data2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly




   If R1.EOF Then
       GoSub FIND_N_NUMBER
   Else
       If Val(R1("ERG")) = 0 Then
          GoSub FIND_N_NUMBER
       Else
          mlabel = R1("erg")
       End If
   End If
   
   MPER = data2.Recordset("ono")
   mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)
   
   TIMM = Format(data2.Recordset("lti5"), "#####.00")
   
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

















F_ArPerEtik = 5

'------------------------------------------------------------------------
For k = 1 To Val(mtimes)

    If True Then
      Printer.FontName = "128"
      Printer.FONTSIZE = 6
      Printer.Print Tab(12); mlabel
      Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
      Printer.FONTSIZE = 8
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print


      Printer.Print Tab(F_ArPerEtik + 10); mlabel
      Printer.Print Tab(F_ArPerEtik + 10); MPER
      Printer.Print Tab(F_ArPerEtik + 10); timologio
      Printer.FONTSIZE = 14
      Printer.FontBold = True
      Printer.Print Tab(F_ArPerEtik + 10); Format(Val(Trim(LTrim(TIMM))), "00.00") + " €"
      Printer.FontBold = False
      
      Printer.EndDoc
    Else
      Printer.Print Tab(F_ArPerEtik + 10); MPER
      Printer.Print Tab(F_ArPerEtik + 20); mlabel
      
      
         BarCode2.SymbologyID = 128
         BarCode2.DataToEncode = DataGrid1.Text
         Printer.PaintPicture BarCode2.Picture, 348, 824
      
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
      Printer.FONTSIZE = 8
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
  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + data2.Recordset("KODE") + "'"
Return




End Sub





Sub patrick_etik()
'========================etiketes ========================================================
Dim d1 As Date, d2 As Date, R As Integer, C As Integer, k As Integer, S As Single
Dim db As Database, DUM, rec As New ADODB.Recordset
Dim mlabel As String
Dim MPER As String
Dim mtimes As Single
Dim TIMM As Single
Dim timologio As String

'On Error Resume NextN  N

 Dim nn As Long
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

   
Dim R1 As New ADODB.Recordset
R1.Open "SELECT * FROM BARCODES WHERE KOD='" + data2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly




   If R1.EOF Then
       GoSub FIND_N_NUMBER  ' EPISTREFEI TO mlabel
   Else
       If Val(R1("ERG")) = 0 Then
          GoSub FIND_N_NUMBER
       Else
          mlabel = R1("ERG")
       End If
   End If
  
   
   MPER = data2.Recordset("ono")
   mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)
   
   TIMM = Format(data2.Recordset("lti5"), "#####.00")
   
   'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")







For k = 1 To Val(mtimes)
    
    
  Printer.FONTSIZE = 8
Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName

Printer.Print Tab(0); MPER


'Printer.FontName = "EAN8"
Printer.FontName = "EAN8"
Printer.FONTSIZE = 8
Printer.Print Tab(0); mlabel

Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
Printer.Print
Printer.Print
Printer.Print

'If Val(TIMM) < 10 Then
'    Printer.FontSize = 14
'Else
    Printer.FONTSIZE = 12
'End If

Printer.FontBold = True

Printer.Print Tab(1); Format(Val(TIMM), "###0.00") + " €"

Printer.FONTSIZE = 8
Printer.FontBold = False


Printer.EndDoc
    
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
  
  Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + data2.Recordset("KOD") + "','" + mlabel + "')", nn
  
   
   If nn = 0 Then
      MsgBox "ΔΕΝ ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ BARCODE"
   End If
  
  
  
'  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Text1(0).Text + "'"
Return

End Sub

Sub SOKRATIS_etik()
'========================etiketes ========================================================
Dim d1 As Date, d2 As Date, R As Integer, C As Integer, k As Integer, S As Single
Dim db As Database, DUM, rec As New ADODB.Recordset
Dim mlabel As String
Dim MPER As String
Dim mtimes As Single
Dim TIMM As Single
Dim timologio As String

'On Error Resume NextN  N

 Dim nn As Long
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

   
Dim R1 As New ADODB.Recordset
R1.Open "SELECT * FROM BARCODES WHERE KOD='" + data2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly




   If R1.EOF Then
       GoSub FIND_N_NUMBER  ' EPISTREFEI TO mlabel
   Else
       If Val(R1("ERG")) = 0 Then
          GoSub FIND_N_NUMBER
       Else
          mlabel = R1("ERG")
       End If
   End If
  
   
   MPER = data2.Recordset("ono")
   mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)
   
   TIMM = Format(data2.Recordset("lti5"), "#####.00")
   
   'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")







For k = 1 To Val(mtimes)
    
    
  Printer.FONTSIZE = 6
Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
Printer.FontBold = True


Printer.CurrentY = 12

Printer.Print Tab(0); MPER


    'If Len(DataGrid1.Text) = 12 Then
        ' BarCode1.SymbologyID = CODE128
    'Else
    '     BarCode1.SymbologyID = EAN13
    'End If
    BarCode1.ShowText = NO
     BarCode1.NarrowBarWidth = 0.025  '0.03
    BarCode1.DataToEncode = DataGrid1.Text
    Printer.PaintPicture BarCode1.Picture, -18, F_YBARCODE  ' '-88,124




'Printer.FontName = "EAN8"
'Printer.FontName = "EAN8"
'Printer.FontSize = 8
Printer.CurrentY = 900
Printer.Print ; Text1(0).Text; " "; mlabel
Printer.FONTSIZE = 8
Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
'Printer.Print
'Printer.Print
'Printer.Print
'Printer.Print

Printer.CurrentY = 900
'Printer.Print Tab(5); Format(Val(TIMM), "###0.00") + " ΕΥΡΩ "
Printer.EndDoc
    
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
  
  Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + data2.Recordset("KOD") + "','" + mlabel + "')", nn
  
   
   If nn = 0 Then
      MsgBox "ΔΕΝ ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ BARCODE"
   End If
  
  
  
'  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Text1(0).Text + "'"
Return

End Sub

Sub GATIDIS_ETIK()

'========================etiketes ========================================================
Dim d1 As Date, d2 As Date, R As Integer, C As Integer, k As Integer, S As Single
Dim db As Database, DUM, rec As New ADODB.Recordset
Dim mlabel As String
Dim MPER As String
Dim mtimes As Single
Dim TIMM As Single
Dim timologio As String

'On Error Resume NextN  N

 Dim nn As Long
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

   
Dim R1 As New ADODB.Recordset
R1.Open "SELECT * FROM BARCODES WHERE KOD='" + data2.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly




   If R1.EOF Then
       GoSub FIND_N_NUMBER  ' EPISTREFEI TO mlabel
   Else
       If Val(R1("ERG")) = 0 Then
          GoSub FIND_N_NUMBER
       Else
          mlabel = R1("ERG")
       End If
   End If
  
   
   MPER = data2.Recordset("ono")
   mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)
   
   TIMM = Format(data2.Recordset("lti5"), "#####.00")
   
   'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")







For k = 1 To Val(mtimes)
    
    
  Printer.FONTSIZE = 6
Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
Printer.FontBold = True


Printer.CurrentY = 12

'Printer.Print Tab(0); MPER


    'If Len(DataGrid1.Text) = 12 Then
        ' BarCode1.SymbologyID = CODE128
    'Else
    '     BarCode1.SymbologyID = EAN13
    'End If
    BarCode1.ShowText = NO
     BarCode1.NarrowBarWidth = 0.025  '0.03
    BarCode1.DataToEncode = DataGrid1.Text
    Printer.PaintPicture BarCode1.Picture, F_XBARCODE, F_YBARCODE  ' '-88,124




'Printer.FontName = "EAN8"
'Printer.FontName = "EAN8"
'Printer.FontSize = 8
Printer.CurrentY = 850
Printer.Print Tab(10); Text1(0).Text
Printer.Print Tab(10); mlabel
Printer.FONTSIZE = 8
Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
'Printer.Print
'Printer.Print
'Printer.Print
'Printer.Print

Printer.CurrentY = 900
'Printer.Print Tab(5); Format(Val(TIMM), "###0.00") + " ΕΥΡΩ "
Printer.EndDoc
    
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
  
  Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + data2.Recordset("KOD") + "','" + mlabel + "')", nn
  
   
   If nn = 0 Then
      MsgBox "ΔΕΝ ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ BARCODE"
   End If
  
  
  
'  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Text1(0).Text + "'"
Return





End Sub

Sub gazoros()
'---------------------------------------------------------------
Dim kod, per, MON, pos, TIMH
Dim f As String, RECS As Integer
Dim k, X
Dim barc(200)
    
 Dim apoSCANNER
    Dim fores
    Dim a$
    'yupie
    fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    
    
If Scanner Then
'
      Open "C:\APOG.TXT" For Input As #1
      X = 0
      k = 0
      Do While Not EOF(1) 'True
          On Error Resume Next
          Line Input #1, a$
          X = X + 1
          kod = mID(a$, 1, 13)
          pos = Val(mID$(a$, 15, 20))
          barc(X) = kod
       Loop
       Close #1

          If X = 0 Then
               MsgBox "δεν υπάρχουν εγγραφές"
               Exit Sub
           End If
         apoSCANNER = X
End If
    
' ΑΝ ΔΕΝ ΕΙΝΑΙ ΑΠΟ ΣΚΑΝΕΡ
If X = 0 Then
   apoSCANNER = 1
   barc(1) = Text1(6).Text
End If


Dim ll
Dim mFORES As Long
Dim R As New ADODB.Recordset
Dim mONOMA, mBarcode, mERG, mLTI5


For ll = 1 To apoSCANNER
'
'mFORES = fores
'
'If Scanner Then
   R.Open "select * FROM EID WHERE ERG='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   If R.EOF Or Len(Trim(barc(ll))) = 0 Then
     mFORES = 0
   Else
     mONOMA = R("ONO") 'ONOMA
     mBarcode = R("KOD") '  ' BARCODE
     mERG = R("ERG") ' erg
     mLTI5 = Format(R("LTI5"), "###0.00") ' LTI5
   End If
   R.Close
'End If
    
    
    
    
    
    If Len(Text1(6).Text) < 13 Then
         BarCode3.SymbologyID = CODE128 ' UPCA
    Else
         BarCode3.SymbologyID = EAN13
    End If
    
    For Each X In Printers
        If InStr(X.DeviceName, "2-8-4-4") Or InStr(X.DeviceName, "SATO") > 0 Then
          ' Set printer as system default.
        
          Set Printer = X
          ' Stop looking for a printer
          Exit For
        End If
    Next
    
    BarCode3.DataToEncode = mERG '  Text1(6).Text 'DataGrid1.Text
    
    
  For k = 1 To fores
    
    Printer.CurrentX = F_ArPerEtik * 100
    Printer.CurrentY = 0
    Printer.FONTSIZE = 10
     Printer.FontBold = False
    Printer.FontName = "Arial Greek"
    
    
    Printer.Print mONOMA 'Text1(1).Text 'perigrafh
    
    Printer.CurrentX = F_ArPerEtik * 100
    
    Printer.Print " " ' Text1(6).Text ' barcode Text1(0).Text 'kodikos
   
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 400 + F_ArPerEtik * 100
    'Printer.FontSize = 12
    Printer.FONTSIZE = 10
    
    
    Printer.CurrentX = 200 + F_ArPerEtik * 100
        Printer.CurrentY = 50
    
    Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 300  '500
    'Printer.CurrentY = 1200
    
     Printer.CurrentY = 1000
   'Printer.Print MERG  ' Text1(6).Text ' barcode Text1(0).Text 'kodikos
   
    Printer.EndDoc
   
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
'    'Printer.Print Format(Text1(13).Text, "##0.00") + " €" ' timh
'    ' Printer.Print "4.55 €" ' timh
'    Printer.EndDoc
'
'  Next
'



End Sub

Sub gazorosTIMES()
'---------------------------------------------------------------
Dim kod, per, MON, pos, TIMH
Dim f As String, RECS As Integer
Dim k, X
Dim barc(200)
    
 Dim apoSCANNER
    Dim fores
    Dim a$
    'yupie
    fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    
    
If Scanner Then
'
      Open "C:\APOG.TXT" For Input As #1
      X = 0
      k = 0
      Do While Not EOF(1) 'True
          On Error Resume Next
          Line Input #1, a$
          X = X + 1
          kod = mID(a$, 1, 13)
          pos = Val(mID$(a$, 15, 20))
          barc(X) = kod
       Loop
       Close #1

          If X = 0 Then
               MsgBox "δεν υπάρχουν εγγραφές"
               Exit Sub
           End If
         apoSCANNER = X
End If
    
' ΑΝ ΔΕΝ ΕΙΝΑΙ ΑΠΟ ΣΚΑΝΕΡ
If X = 0 Then
   apoSCANNER = 1
   barc(1) = Text1(6).Text
End If


Dim ll
Dim mFORES As Long
Dim R As New ADODB.Recordset
Dim mONOMA, mBarcode, mERG, mLTI5


For ll = 1 To apoSCANNER
'
'mFORES = fores
'
'If Scanner Then
   R.Open "select * FROM EID WHERE ERG='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   If R.EOF Or Len(Trim(barc(ll))) = 0 Then
     mFORES = 0
   Else
     mONOMA = R("ONO") 'ONOMA
     mBarcode = R("KOD") '  ' BARCODE
     mERG = R("ERG") ' erg
     mLTI5 = Format(R("LTI5"), "###0.00") ' LTI5
   End If
   R.Close
'End If
    
    
    
    
    
    If Len(Text1(6).Text) < 13 Then
         BarCode3.SymbologyID = CODE128 ' UPCA
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
    
    BarCode3.DataToEncode = mERG '  Text1(6).Text 'DataGrid1.Text
    
   '   BarCode3.DataToEncode = Text1(6).Text 'DataGrid1.Text
    
    
  For k = 1 To fores
    
    Printer.CurrentX = F_ArPerEtik * 100
    Printer.CurrentY = 200
    Printer.FONTSIZE = 12
     Printer.FontBold = False
    Printer.FontName = "Arial Greek"
    
    
    Printer.Print mONOMA ' Text1(1).Text 'perigrafh
    
    Printer.CurrentX = F_ArPerEtik * 100
    
    
     Printer.CurrentY = 0
    Printer.CurrentX = 800
    Printer.FONTSIZE = 8
   ' Printer.Print Tab(10); MERG ' Text1(6).Text ' barcode    kodikos
   
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 400 + F_ArPerEtik * 100
    'Printer.FontSize = 12
    Printer.FONTSIZE = 10
    
    
    Printer.CurrentX = 200 + F_ArPerEtik * 100
        Printer.CurrentY = 100
    
    Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 500 '1124
    Printer.CurrentY = 1200
    
    Printer.CurrentX = 1900 + F_ArPerEtik * 100
    Printer.FONTSIZE = 24
    Printer.FontBold = True
    Printer.Print mLTI5 + " €" ' timh
    ' Printer.Print "4.55 €" ' timh
    Printer.EndDoc
   
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
'    'Printer.Print Format(Text1(13).Text, "##0.00") + " €" ' timh
'    ' Printer.Print "4.55 €" ' timh
'    Printer.EndDoc
'
'  Next
'



End Sub

Private Sub Command7_Click()

Dim fores As Long, k As Integer
Dim X As Printer
Dim DUM



'========================================================================00002357

If Len(Dir("C:\GAZOROS2.EXE", vbNormal)) > 0 Then
    gazoros
    
   Exit Sub
End If













On Error Resume Next


If Len(Dir("C:\LAGEURO\LAGAKIS.TXT", vbNormal)) > 0 Then
    LAG_ETIK
    Exit Sub
End If

If Len(Dir("C:\LAGEURO\diesel.TXT", vbNormal)) > 0 Then
    diesel_ETIK
    Exit Sub
End If



If Len(Dir("C:\LAGEURO\LAGAKIS2.TXT", vbNormal)) > 0 Then
    LAG2_ETIK
    Exit Sub
End If




If Len(Dir("C:\LAGEURO\patrick.TXT", vbNormal)) > 0 Then
    patrick_etik
    Exit Sub
End If


If Len(Dir("C:\LAGEURO\SOKRATIS.TXT", vbNormal)) > 0 Then
    SOKRATIS_etik
    Exit Sub
End If

If Len(Dir("C:\LAGEURO\GATIDIS.TXT", vbNormal)) > 0 Then
    GATIDIS_ETIK
    Exit Sub
End If


If Len(Dir("C:\LAGEURO\MAKISGET.EXE", vbNormal)) > 0 Then
    'DUM = Shell("c:\lageuro\makisget.exe " + DATA2.Recordset("KOD") + " " + Replace(DATA2.Recordset("ONO"), " ", "_") + " " + Format(DATA2.Recordset("LTI5"), "##0.00") + " " + Format(DATA2.Recordset("POS_KERD"), "##0") + " " + DATA2.Recordset("ERG"))
    MAKIS_ETIK
    Exit Sub
End If


If Len(Dir("C:\DIONISIS.EXE", vbNormal)) > 0 Then
    fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    Dim PARAM
    
    If Len(DataGrid1.Text) < 12 Then
       MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
       Exit Sub
    End If
    
    If Len(DataGrid1.Text) = 12 Then
         BarCode2.SymbologyID = UPCA
    Else
         BarCode2.SymbologyID = EAN13
    End If
    
    For Each X In Printers
        If InStr(X.DeviceName, "2844") > 0 Then
          ' Set printer as system default.
        
          Set Printer = X
          ' Stop looking for a printer
          Exit For
        End If
    Next
    
    BarCode2.DataToEncode = DataGrid1.Text
    
    
   For k = 1 To fores
    Printer.CurrentY = 400
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 600 + F_ArPerEtik * 100
    Printer.FONTSIZE = 12
    
    Printer.Print Text1(0).Text
    Printer.PaintPicture BarCode2.Picture, F_XBARCODE, F_YBARCODE
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
  'προγραμμα dos για εκτυπωτη epson tml-60
  Shell ("C:\LAGEURO\MPOYGET.EXE " + Text1(0).Text + " " + Str(fores) + " " + Replace(Text1(1).Text, " ", ".") + " " + "1")
  Exit Sub
End If
' MPOYGET.EXE 5201219046154 2 PERIGRAFH 3.56

 For k = 1 To fores

  If Len(Dir("C:\mercvb\zebra.txt", vbNormal)) > 0 Then
    
    
'    Printer.Print Tab(20); Text1(1).Text 'ONOMA
 '   Printer.Print Tab(30); Text1(6).Text ' mlabel  5200016000277
 
 
    'Printer.FontName = "Arial" ' "128"
  Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName454967325476-5200016000239
  
  
    Printer.FONTSIZE = 8
    
  
  
  
  
  Printer.CurrentY = 0
    Printer.CurrentX = 0
    'Printer.Print Left(Text1(6).Text, 12)
    

     BarCode2.DataToEncode = Text1(0).Text
     BarCode2.Code128CharSet = Set_A
 BarCode2.SymbologyID = EAN13
 Printer.PaintPicture BarCode2.Picture, 200, -500

 
    
'    If Len(Trim(Text1(0).Text)) = 12 Then
'       Printer.FontName = "UPCA" ' "128"
'    Else
'       Printer.FontName = "C128" '  "EAN13" ' "128"
'    End If
'   ' Printer.FontSize = 6
'    Printer.CurrentY = 0
'    Printer.CurrentX = 800
'    Printer.Print Left(Text1(6).Text, 13) '  "123456789012" '  Text1(6).Text 'eggtim.Recordset("kode") ' mlabel
  
    Printer.EndDoc
  
  
  
  
  
  Else
    Printer.Print Tab(20); Text1(1).Text 'ONOMA
    Printer.Print Tab(30); Text1(6).Text ' mlabel
    If Len(Trim(Text1(0).Text)) = 12 Then
       Printer.FontName = "UPCA" ' "128"
    Else
       Printer.FontName = "ean13" ' "128"
    End If
    Printer.FONTSIZE = 6
    Printer.Print Tab(18); Text1(0).Text 'eggtim.Recordset("kode") ' mlabel
    Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
    Printer.FONTSIZE = 8
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.EndDoc
  End If
Next
End Sub

Private Sub Command8_Click()
Dim nn As Long

Dim ll, apoSCANNER As Long
Dim db
Dim PSIFIA As Integer
Dim fores As Long, k As Integer
Dim X As Printer
Dim barc(2600)
Dim mONOMA, mBarcode, mERG, mLTI5
Dim R As New ADODB.Recordset
Dim RD As DAO.Recordset
mONOMA = Text1(1).Text 'ONOMA
mBarcode = Text1(0).Text '  ' BARCODE
mERG = Text1(6).Text ' erg
mLTI5 = Format(Val(Text1(13).Text), "###0.00") ' LTI5
    Dim PARAM


If Len(Dir("C:\FANAR.EXE", vbNormal)) > 0 Then
   fanar_etik
   Exit Sub
End If



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
'    FGDB2.Execute "update Material  set retailprice=" + Text1(13).Text + "  WHERE  Code='" + Text1(0).Text + "'", nn
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
'    Printer.Print Format(Text1(13).Text, "##0.00") + " €" ' timh
'    ' Printer.Print "4.55 €" ' timh
'    Printer.EndDoc
'   'Next
'
'
'   Exit Sub
'End If


If Len(Dir("C:\DION2.EXE", vbNormal)) > 0 Then
    'yupie
    'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    
    If Len(DataGrid1.Text) < 12 Then
       MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
       Exit Sub
    End If
    
    If Len(DataGrid1.Text) = 12 Then
         BarCode3.SymbologyID = UPCA
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
    
    BarCode3.DataToEncode = DataGrid1.Text
    
    
   'For K = 1 To fores
    
    Printer.CurrentX = F_ArPerEtik * 100
    Printer.CurrentY = 200
    Printer.FONTSIZE = 10
     Printer.FontBold = False
    Printer.FontName = "Arial Greek"
    
    
    Printer.Print Text1(1).Text 'perigrafh
    
    Printer.CurrentX = F_ArPerEtik * 100
    
    Printer.Print Text1(0).Text 'kodikos
   
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 400 + F_ArPerEtik * 100
    'Printer.FontSize = 12
    Printer.FONTSIZE = 10
    
    
    Printer.CurrentX = 200 + F_ArPerEtik * 100
        Printer.CurrentY = 600
    
    Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 1124
    
    
    Printer.CurrentX = 1000 + F_ArPerEtik * 100
    Printer.FONTSIZE = 24
    Printer.FontBold = True
    Printer.Print Format(Text1(13).Text, "##0.00") + " €" ' timh
    ' Printer.Print "4.55 €" ' timh
    Printer.EndDoc
   'Next
    
    
   Exit Sub
End If



'========================================================================
If Len(Dir("C:\GAZOROS.EXE", vbNormal)) > 0 Then
     gazorosTIMES
     Exit Sub
    'yupie
    fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    
    'If Len(DataGrid1.Text) < 12 Then
    '   MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
    '   Exit Sub
    'End If
    
    If Len(Text1(6).Text) < 13 Then
         BarCode3.SymbologyID = CODE128 ' UPCA
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
    
    BarCode3.DataToEncode = Text1(6).Text 'DataGrid1.Text
    
    
  For k = 1 To fores
    
    Printer.CurrentX = F_ArPerEtik * 100
    Printer.CurrentY = 200
    Printer.FONTSIZE = 12
     Printer.FontBold = False
    Printer.FontName = "Arial Greek"
    
    
    Printer.Print Text1(1).Text 'perigrafh
    
    Printer.CurrentX = F_ArPerEtik * 100
    
    
     Printer.CurrentY = 450
    Printer.CurrentX = 200
    Printer.FONTSIZE = 8
    Printer.Print Tab(10); Text1(6).Text ' barcode    kodikos
   
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 400 + F_ArPerEtik * 100
    'Printer.FontSize = 12
    Printer.FONTSIZE = 10
    
    
    Printer.CurrentX = 200 + F_ArPerEtik * 100
        Printer.CurrentY = 100
    
    Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 500 '1124
    Printer.CurrentY = 1200
    
    Printer.CurrentX = 1900 + F_ArPerEtik * 100
    Printer.FONTSIZE = 24
    Printer.FontBold = True
    Printer.Print Format(Text1(13).Text, "##0.00") + " €" ' timh
    ' Printer.Print "4.55 €" ' timh
    Printer.EndDoc
   
  Next
    
    
   Exit Sub
End If






If Len(Dir("C:\DION3.EXE", vbNormal)) > 0 Then
    'yupie
    'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    
    If Len(DataGrid1.Text) < 12 Then
       MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
       Exit Sub
    End If
    
    If Len(DataGrid1.Text) = 12 Then
         BarCode3.SymbologyID = UPCA
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
    
    BarCode3.DataToEncode = DataGrid1.Text
    
    
   'For K = 1 To fores
    
    Printer.CurrentX = F_ArPerEtik * 100
    
    Printer.FONTSIZE = 10
     Printer.FontBold = False
    Printer.FontName = "Arial Greek"
    
    
    'Printer.CurrentY = 1100
    'Printer.Print DataGrid1.Text
    
     Printer.CurrentY = 200
   
    
    
    Printer.Print Text1(1).Text 'perigrafh
    
    
    
    
    
    
    
    Printer.CurrentX = F_ArPerEtik * 100
    
    Printer.Print Text1(0).Text 'kodikos
   
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 400 + F_ArPerEtik * 100
    'Printer.FontSize = 12
    Printer.FONTSIZE = 10
    
    
    Printer.CurrentX = 200 + F_ArPerEtik * 100
      Printer.CurrentY = 800
    
    Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 100, 1224
    
    Printer.CurrentY = 600
    Printer.CurrentX = 1000 + F_ArPerEtik * 100
    Printer.FONTSIZE = 24
    Printer.FontBold = True
    Printer.Print Format(Text1(13).Text, "##0.00") + " €" ' timh
    ' Printer.Print "4.55 €" ' timh
    Printer.EndDoc
   'Next
    
    
   Exit Sub
End If

'--------------------------------------------------------------------------------
If Len(Dir("C:\DIONXONDR.TXT", vbNormal)) > 0 Then
    
    'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    
    
    If Len(DataGrid1.Text) < 12 Then
       MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
       Exit Sub
    End If
    
    If Len(DataGrid1.Text) = 12 Then
         BarCode2.SymbologyID = UPCA
    Else
         BarCode2.SymbologyID = EAN13
    End If
    
    For Each X In Printers
        If InStr(X.DeviceName, "2844") > 0 Then
          ' Set printer as system default.
        
          Set Printer = X
          ' Stop looking for a printer
          Exit For
        End If
    Next
    
    BarCode2.DataToEncode = DataGrid1.Text
    
    
   'For K = 1 To fores
   
   Printer.FontBold = False
    Printer.FontName = Text1(1).FontName  '  "Courier New Greek"
    
    Printer.CurrentX = F_2ArPerEtik * 100
    Printer.CurrentY = 200
    Printer.FONTSIZE = 10
    Printer.Print Text1(1).Text 'perigrafh
    
    Printer.CurrentX = F_2ArPerEtik * 100
    
    Printer.Print Text1(0).Text 'kodikos
   
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 400 + F_2ArPerEtik * 100
    'Printer.FontSize = 12
    Printer.FONTSIZE = 10
    
    
    Printer.CurrentX = 200 + F_2ArPerEtik * 100
        Printer.CurrentY = 600
    
    Printer.PaintPicture BarCode2.Picture, F_2ArPerEtik * 100, 1124
    
    
    Printer.CurrentX = 1000 + F_2ArPerEtik * 100
    Printer.FONTSIZE = 16
    Printer.FontBold = True
    Printer.FontName = "Courier New"
    
    Printer.Print Format(Text1(3).Text, "##0.00") + " € (" + Format((100 + g_Fpa(Combo2.Text)) / 100 * Text1(3).Text, "##0.00") + ")" ' timh
    
    
    ' Printer.Print "4.55 €" ' timh
    Printer.EndDoc
   'Next
    
    
   Exit Sub
End If
'-----------------------------------------------------------If Len(Dir("C:\DIONISIS.EXE", vbNormal)) > 0 Then


'--------------------------------------------------------------------------------
If Len(Dir("C:\DIONISIS.EXE", vbNormal)) > 0 Then
    
    'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    
    
    If Len(DataGrid1.Text) < 12 Then
       MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
       Exit Sub
    End If
    
    If Len(DataGrid1.Text) = 12 Then
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
    
    BarCode1.DataToEncode = DataGrid1.Text
    
    
   'For K = 1 To fores
   
   Printer.FontBold = False
    Printer.FontName = "Arial"
    
    Printer.CurrentX = F_ArPerEtik * 100
    Printer.CurrentY = 200
    Printer.FONTSIZE = 10
    Printer.Print Text1(1).Text 'perigrafh
    
    Printer.CurrentX = F_ArPerEtik * 100
    
    Printer.Print Text1(0).Text 'kodikos
   
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 400 + F_ArPerEtik * 100
    'Printer.FontSize = 12
    Printer.FONTSIZE = 10
    
    
    Printer.CurrentX = 200 + F_ArPerEtik * 100
        Printer.CurrentY = 600
    
    Printer.PaintPicture BarCode1.Picture, F_ArPerEtik * 100, 1124
    
    
    Printer.CurrentX = 1000 + F_ArPerEtik * 100
    Printer.FONTSIZE = 24
    Printer.FontBold = True
    Printer.Print Format(Text1(13).Text, "##0.00") + " €" ' timh
    ' Printer.Print "4.55 €" ' timh
    Printer.EndDoc
   'Next
    
    
   Exit Sub
End If
'-----------------------------------------------------------If Len(Dir("C:\DIONISIS.EXE", vbNormal)) > 0 Then


On Error Resume Next
For Each X In Printers
    If InStr(X.DeviceName, "SATO") > 0 Then
        ' Set printer as system default.
        Set Printer = X
        ' Stop looking for a printer  5200016000239=== 5200016000192
        Exit For
    End If
Next

apoSCANNER = 1
If Scanner Then

 ' μπουγουδης
Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
Set RD = db.OpenRecordset("SELECT * FROM TIMAGOR ")
If RD.RecordCount = 0 Then
   MsgBox "δεν υπάρχουν εγγραφές"
   Exit Sub
End If

  
  RD.MoveFirst
  k = 0
  Do While Not RD.EOF
     k = k + 1
     If k > 2000 Then Exit Do
     If IsNull(RD("BARCODE")) Then
        barc(k) = " " 'Grid1.TextMatrix(K, f_k) = " "
     Else
        barc(k) = RD("barcode") 'Grid1.TextMatrix(K, f_k) = R("BARCODE")
     End If
     RD.MoveNext
  Loop
  RD.Close
  db.Close
  apoSCANNER = k
End If
'----------------------
fores = 1

fores = Val(InputBox("Πόσες εττικέτες; ", "", 1))
Dim mFORES As Long
Dim R1 As New ADODB.Recordset
Dim C0 As String

For ll = 1 To apoSCANNER

mFORES = fores

If Scanner Then
   R.Close
   ' R.Open "SELECT * FROM BARCODES WHERE ERG = '" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
    'C0 = R("KOD")
    'R.Close
    
   R.Open "select * FROM EID WHERE KOD LIKE '%" + barc(ll) + "%'", Gdb, adOpenForwardOnly, adLockReadOnly
   
  If R.EOF Or Len(Trim(barc(ll))) = 0 Then
     mFORES = 0
  Else
      'το r1 xrhsimopoihuhke για να τυπωνει to barcode 15/02/2010
      'R1.Open "SELECT top 1 * FROM BARCODES WHERE KOD='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
      mBarcode = R("ERG") '  R("KOD") '  ' BARCO
      'R1.Close
      mONOMA = R("ONO") 'ONOMA
   'DE
      Text1(6).Text = R("ERG")
      
      mERG = R("ERG") ' erg
      mLTI5 = Format(R("LTI5"), "###0.00")  ' LTI5
  End If
Else
      R1.Open "SELECT top 1 ERG FROM BARCODES WHERE KOD='" + Text1(0).Text + "'", Gdb, adOpenForwardOnly, adLockReadOnly
      mBarcode = R1("ERG")
      R1.Close
      
  
End If


Dim Z As Integer

Z = F_ArPerEtik ' InputBox("METATOPISH")

For k = 1 To mFORES
   If Len(Dir("C:\CMPOYGZEB.EXE", vbNormal)) > 0 Then
      GoSub ZEBRA
   Else
      GoSub SATO
   End If  ' 078257586417
   
   


Next k

Next ll


Exit Sub




'------------------- ZEBRA-------------------------------------
ZEBRA:
Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992  123456789012

'Printer.CurrentY = 0

Printer.FONTSIZE = 8
Printer.FontBold = True
Printer.Print Tab(4); mONOMA 'Text1(1).Text 'ONOMA
Printer.FONTSIZE = 8
Printer.Print Tab(4); mBarcode '  Text1(0).Text;  ' BARCODE
Printer.CurrentY = 350
'boithitikos kodikos
Printer.Print Tab(4); Text1(6).Text ' mlabel
 Printer.Print
 Printer.CurrentY = 200
    
       
    If Len(DataGrid1.Text) = 12 Then
         BarCode4.SymbologyID = UPCA
    Else
         BarCode4.SymbologyID = EAN13
    End If
'    BarCode4.ShowText = NO
    BarCode4.DataToEncode = mBarcode
     
    Printer.PaintPicture BarCode4.Picture, 1900, 150, , 400 ' PICTURE ,X,Y,PLATOS,YPSOS  500ΥΨΟΣ



'If Len(Trim(Text1(0).Text)) = 12 Then
'   Printer.FontName = "UPCA" ' "128"
'Else
'   Printer.FontName = "EAN13SMALL" ' "128"
'End If
'Printer.Print Tab(14 + Z); mBarcode '  Text1(0).Text;  ' BARCODE

If F_PONTOI = 1 Then
   Printer.FONTSIZE = 5
'   Printer.Print Tab(15 + Z); mBarcode 'Text1(0).Text 'eggtim.Recordset("kode") ' mlabel
End If
Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
Printer.FontBold = True
Printer.FONTSIZE = 8
Printer.CurrentY = 570
Printer.FONTSIZE = 24
If Len(mLTI5) = 6 Then ' 208.35
     Printer.Print Tab(1 + 7); "" + mLTI5; " €"   '-7 EIXE STIS 17-8-2006
ElseIf Len(mLTI5) = 5 Then '18.35
     Printer.Print Tab(2 + 7); "" + mLTI5; " €"
Else
     Printer.Print Tab(2 + 8); "" + mLTI5; " €"
End If
Printer.FONTSIZE = 6
Printer.Print
Printer.FONTSIZE = 8
Printer.Print
If F_PONTOI = 0 Then 'Len(Dir("C:\LAGEURO\NOPONTOI.TXT", vbNormal)) > 0 Then
Else
    ' ΜΠΟΥΓΟΥΔΗΣ
    Printer.CurrentX = 200
    Printer.CurrentY = 550
    Printer.FONTSIZE = 14
  If Val(mLTI5) > 2.99 Then
     Printer.Print Tab(Z - 4); "ΠΟΝΤΟΙ"
     PSIFIA = Len(Format(Int(Val(mLTI5) / 3), "##0"))
     If PSIFIA = 3 Then Printer.Print ; Tab(5 + Z - 7); Format(Int(Val(mLTI5) / 3), "##0")
     If PSIFIA = 1 Then Printer.Print ; Tab(6 + Z - 6); Format(Int(Val(mLTI5) / 3), "##0")
     If PSIFIA = 2 Then Printer.Print ; Tab(5 + Z - 6); Format(Int(Val(mLTI5) / 3), "##0")
  End If
End If

Printer.EndDoc
Return

'--------------------------------------------SATO ----------------------
SATO:
Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName  045496370992
Printer.FONTSIZE = 8
Printer.FontBold = True
Printer.Print Tab(4 + Z); mONOMA 'Text1(1).Text 'ONOMA
Printer.FONTSIZE = 8
'Printer.Print Tab(4 + Z); mBarcode '  Text1(0).Text;  ' BARCODE
Printer.CurrentY = 350
'boithitikos kodikos
Printer.Print Tab(2 + Z); Text1(6).Text ' mlabel
 Printer.Print
 Printer.CurrentY = 200
If Len(mBarcode) = 12 Then ' If Len(Trim(Text1(0).Text)) = 12 Then
 '  Printer.FontName = "UPCA" ' "128"
Else
'   Printer.FontName = "EAN13" ' "128"
End If
Printer.Print Tab(10 + Z); mBarcode '  Text1(0).Text;  ' BARCODE

If F_PONTOI = 1 Then
   Printer.FONTSIZE = 5
'   Printer.Print Tab(15 + Z); mBarcode 'Text1(0).Text 'eggtim.Recordset("kode") ' mlabel
End If
Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
Printer.FontBold = True
Printer.FONTSIZE = 8
Printer.CurrentY = 500
Printer.FONTSIZE = 24
If Len(mLTI5) = 6 Then ' 208.35
     Printer.Print Tab(1 + Z); "" + mLTI5; " €"   '-7 EIXE STIS 17-8-2006
ElseIf Len(mLTI5) = 5 Then '18.35
     Printer.Print Tab(2 + Z); "" + mLTI5; " €"
Else
     Printer.Print Tab(2 + Z); "" + mLTI5; " €"
End If
Printer.FONTSIZE = 6
Printer.Print
Printer.FONTSIZE = 8
Printer.Print
If F_PONTOI = 0 Then 'Len(Dir("C:\LAGEURO\NOPONTOI.TXT", vbNormal)) > 0 Then
Else
    ' ΜΠΟΥΓΟΥΔΗΣ
    Printer.CurrentX = 200
    Printer.CurrentY = 500
    Printer.FONTSIZE = 14
  If Val(mLTI5) > 2.99 Then
     Printer.Print Tab(Z - 2); "ΠΟΝΤΟΙ"
     PSIFIA = Len(Format(Int(Val(mLTI5) / 3), "##0"))
     If PSIFIA = 3 Then Printer.Print ; Tab(5 + Z - 5); Format(Int(Val(mLTI5) / 3), "##0")
     If PSIFIA = 1 Then Printer.Print ; Tab(6 + Z - 4); Format(Int(Val(mLTI5) / 3), "##0")
     If PSIFIA = 2 Then Printer.Print ; Tab(5 + Z - 4); Format(Int(Val(mLTI5) / 3), "##0")
  End If
End If

Printer.EndDoc
Return

End Sub

Private Sub Command9_Click()
  data22.RecordSource = UCase(" SELECT * FROM EID WHERE MEMO LIKE '%" + Text2.Text + "%';")
  data2.Refresh
  DBGrid1.SetFocus
  

'   SELECT * FROM EID WHERE MEMO LIKE '%DOKIM%'
End Sub

Private Sub DataGrid1_Click()
' On Error Resume Next
 Text3.Text = DataGrid1.Text
End Sub

Private Sub DBGrid1_Click()
  ' mDBGrid1_RowColChange
  ' MILSEC 3000
  ' mDBGrid1_RowColChange
End Sub

Private Sub DBGRID1_KeyDown(KeyCode As Integer, Shift As Integer)
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
Dim G, m, a, GA, MA, P, k, SM, SG, SM2, SG2
Dim mPos01 As Single, mPos02 As Single
Dim mPos03 As Single
Dim mPos04 As Single
On Error Resume Next
Dim Sql2 As String

DBGrid1.BackColorSel = vbBlue
Me.MousePointer = vbHourglass
data2.RecordSource = " select * from EID WHERE KOD='" + DBGrid1.Text + "'"
data2.Refresh

'  Me.Caption = Now
Text1(0).Text = data2.Recordset("kod")
Text1(1).Text = data2.Recordset("ONO")

Text1(16).Text = data2.Recordset("koderg")


Text1(4).Text = IIf(IsNull(data2.Recordset("pos")), " ", data2.Recordset("pos"))


Combo2.Text = data2.Recordset("FPA")

Dim n As Integer
For n = 0 To Combo5.ListCount - 1
   If Val(Left(Combo5.List(n), 3)) = Val(data2.Recordset("kathgoria")) Then
      Combo5.Text = Combo5.List(n) 'OIKOGENEIA
      Exit For
   End If
Next







       
'Combo5.Text = Combo5.List(Val(DATA2.Recordset("KATHGORIA")) - 1)
       
Combo1.Text = data2.Recordset("mon")

' Combo4.Text = Combo4.List(Val(Data2.Recordset("AEG")) - 1) 'OIKOGENEIA
'Dim n As Integer
For n = 0 To Combo4.ListCount - 1
   If Val(Left(Combo4.List(n), 3)) = Val(data2.Recordset("AEG")) Then
      Combo4.Text = Combo4.List(n) 'OIKOGENEIA
      Exit For
   End If
Next





Dim R As New ADODB.Recordset
Dim NYP As Integer
NYP = 0
Dim NYP0 As Integer

If Len(Combo4.Text) > 0 Then
 
' On Error GoTo LATOSEXo
On Error Resume Next

 R.Open "SELECT *FROM PINAKES WHERE TYPOS=3 AND TIMH=" + Left(Combo4.Text, 2) + "ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
 Combo3.Clear


Do While Not R.EOF

     Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
     If R("AYJON") = Val(data2.Recordset("kodlog")) Then
        NYP0 = NYP
     End If

     
  R.MoveNext
  NYP = NYP + 1
Loop
R.Close

End If

On Error Resume Next






Combo3.Text = Combo3.List(NYP0)


'Combo3.Text = Combo3.List(Val(Data2.Recordset("kodlog")) - 1)
       
Text1(2).Text = data2.Recordset("xti")


'If IsNull(Data2.Recordset("pos_kerd2")) Then
   Text1(5).Text = nNull(data2.Recordset("pos_kerd2"))
'End If


'If IsNull(Data2.Recordset("pos_kerd")) Then
   Text1(14).Text = nNull(data2.Recordset("pos_kerd"))
'End If

Text1(6).Text = CNull(data2.Recordset("ERG"))
ONO2.Text = CNull(data2.Recordset("ono2"))

Text1(3).Text = data2.Recordset("lti")
Text1(13).Text = data2.Recordset("lti5")


XondrmeFPA.Text = data2.Recordset("lti") * (1 + g_Fpa(data2.Recordset("fpa")) / 100)
 
 If IsNull(data2.Recordset("MEMO")) Then
    Text14.Text = ""
 Else
    Text14.Text = data2.Recordset("MEMO")
 End If
 
 Text1(15).Text = data2.Recordset("PROMHU")    ' BONUS
       
       Text1(11).Text = data2.Recordset("kodSYNOD")
       Text1(12).Text = data2.Recordset("kodlogAG")
       
       Text1(7).Text = data2.Recordset("PROM")    'Right(Combo3.Text, 3)
        Text1(9).Text = data2.Recordset("SPA")   'Right(Combo3.Text, 3)
       
       Text1(10).Text = data2.Recordset("UES")    'Right(Combo3.Text, 3)
       Text1(8).Text = data2.Recordset("EPIUYP")

       exoda.Text = data2.Recordset("pos_kerd3")



If IsNull(data2.Recordset("POS01")) Then mPos01 = 0 Else mPos01 = data2.Recordset("POS01")

If IsNull(data2.Recordset("POS02")) Then mPos02 = 0 Else mPos02 = data2.Recordset("POS02")

If IsNull(data2.Recordset("POS03")) Then mPos03 = 0 Else mPos03 = data2.Recordset("POS03")
If IsNull(data2.Recordset("POS04")) Then mPos04 = 0 Else mPos04 = data2.Recordset("POS04")

' Label2.Caption = "ΤΕΛ.ΕΝΗΜΕΡΩΣΗ: KENΤΡ :" + Format(mPos01, "###0") + " ΥΠ.1: " + Format(mPos02, "###0") + " ΥΠ.3: " + Format(mPos03, "###0")

If Combo9.ListCount = 2 + 1 Then
   Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " + Trim(mID(Combo9.List(0), 3, 10)) + " :" _
   + Format(mPos01, "###0") + "   " + Trim(mID(Combo9.List(1), 3, 10)) + " :" _
   + Format(mPos02, "###0")
ElseIf Combo9.ListCount = 1 + 1 Then
   Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " + Trim(mID(Combo9.List(0), 3, 10)) + " :" _
   + Format(mPos01, "###0")
ElseIf Combo9.ListCount = 3 + 1 Then
   Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " + Trim(mID(Combo9.List(0), 3, 10)) + _
   " :" + Format(mPos01, "###0") + " - " + Trim(mID(Combo9.List(1), 3, 10)) + " :" + _
   "   " + Format(mPos02, "###0") + " - " + Trim(mID(Combo9.List(2), 3, 10)) + " :" _
   + Format(mPos03, "###0")
   
ElseIf Combo9.ListCount = 4 + 1 Then
   Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " _
   + Trim(mID(Combo9.List(0), 3, 10)) + " :" + Format(mPos01, "###0") + "   " _
   + Trim(mID(Combo9.List(1), 3, 10)) + " :  " + Format(mPos02, "###0") + "   " _
   + Trim(mID(Combo9.List(2), 3, 10)) + " :" + Format(mPos03, "###0") + "   " _
   + Trim(mID(Combo9.List(3), 3, 10)) + " :" + Format(mPos04, "###0")
   
End If




  Text1(18).Text = data2.Recordset("SYSKMAX") 'SYSKEYASIA
  Text1(17).Text = data2.Recordset("SYSKMIN") ' YPOSYSKEYASIA
  MSFlexGrid1.row = 0
  MSFlexGrid1.Col = 1
  MSFlexGrid1.Text = "Αγορές"
  MSFlexGrid1.Col = 2
  MSFlexGrid1.Text = "Αξία Αγορ."
  
  MSFlexGrid1.Col = 3
  MSFlexGrid1.Text = "Πωλήσεις"

  MSFlexGrid1.Col = 4
  MSFlexGrid1.Text = "Αξία πωλήσεων"

'' Γραφική παράσταση
's = 0

'For l = 1 To 12
   ' Exit Sub
 '  G.row = 1 + addit: G.Column = l: G.Data = Stat(l, addit2 + 1) - Stat(l, addit2 + 2)
  ' G.row = 4 + addit: G.Column = l: G.Data = Stat(l, addit2 + 4) - Stat(l, addit2 + 5)
'Next
If data2.Recordset.RecordCount = 0 Then
   Me.MousePointer = vbNormal
   If cXRONIES.value = vbChecked Then Gdb.Close: Gdb.Open gConnect
   Exit Sub
End If

SM = 0: SG = 0
SM2 = 0: SG2 = 0
' Dim R As New ADODB.Recordset

On Error GoTo LATOS

Dim synt
synt = " HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND HME<'" + Format(gLHJH1, "MM/DD/YYYY") + "' "

If cXRONIES.value = vbChecked Then
  Gdb.Close
  'f_PROPERSI = Trim(mID(XRONIES.Text, 21, 20))
  Gdb.Open gConnect + ";DATABASE=" + Trim(mID(XRONIES.Text, 21, 20))
  synt = " YEAR(HME)>1900  "
End If







If Check1.value = vbUnchecked Then GoTo 700

R.Open "SELECT MONTH(HME),SUM(XRE),SUM(PIS),SUM(XRE*TIMM),SUM(PIS*TIMM) FROM EGGTIM WHERE " + synt + " AND KODE='" + data2.Recordset("kod") + "' GROUP BY MONTH(HME) ", Gdb, adOpenForwardOnly, adLockReadOnly
Dim A_XRE(12) As Single
Dim A_PIS(12) As Single
Dim A_XREAJ(12) As Single
Dim A_PISAJ(12) As Single

For k = 1 To 12: A_XRE(k) = 0: A_PIS(k) = 0: A_XREAJ(k) = 0: A_PISAJ(k) = 0: Next
' Exit Sub
If Not R.EOF Then

 On Error Resume Next 'GoTo NOCONNECT

R.MoveFirst
Do While Not R.EOF
   'On Error GoTo NOCONNECT
   A_XRE(R(0)) = R(1): A_XREAJ(R(0)) = R(3)
   A_PIS(R(0)) = R(2): A_PISAJ(R(0)) = R(4)
   R.MoveNext
Loop
R.Close
On Error Resume Next
End If

'graph1.top = 4680  'arxika
 '  graph1.height = 4485
   
  ' graph1.top = 46
   'graph1.height = 8485

MSFlexGrid1.Visible = False

For k = 1 To 12
'  G = "G" + Format(k, "00")
'  M = "M" + Format(k, "00")
'  GA = "g" + Format(k, "00")
'  MA = "M" + Format(k, "00")
  
  graph1.row = 1: graph1.Column = k: graph1.DATA = A_XRE(k) ' DATA2.Recordset(G)
  graph1.row = 2: graph1.Column = k: graph1.DATA = A_PIS(k) 'DATA2.Recordset(M)
   graph1.ColumnLabel = k
    graph1.Refresh
    
   
  
  MSFlexGrid1.row = k
  MSFlexGrid1.Col = 0
  MSFlexGrid1.Text = k
    
  MSFlexGrid1.Col = 1

  
  SM = SM + A_PIS(k) 'DATA2.Recordset(M)
  SG = SG + A_XRE(k) ' DATA2.Recordset(G)
  
  
  SM2 = SM2 + A_PISAJ(k) 'DATA2.Recordset(M)
  SG2 = SG2 + A_XREAJ(k) ' DATA2.Recordset(G)
  
  
  
  
  MSFlexGrid1.Text = Format(A_XRE(k), "#####.##") ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset(G)), "####.##")

  
  MSFlexGrid1.Col = 2
  MSFlexGrid1.Text = Format(A_XREAJ(k), "#####.##") ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset("XTI") * DATA2.Recordset(G)), "#####,##")

  MSFlexGrid1.Col = 3
  MSFlexGrid1.Text = Format(A_PIS(k), "#####.##") 'Format(IIf(IsNull(DATA2.Recordset(M)), 0, DATA2.Recordset(M)), "#####,##")


  MSFlexGrid1.Col = 4
  MSFlexGrid1.Text = Format(A_PISAJ(k), "#####.##") 'Format(IIf(IsNull(DATA2.Recordset(MA)), 0, DATA2.Recordset("LTI") * DATA2.Recordset(MA)), "#####,##")

  
  
  
  
  
  
  
Next

If MSFlexGrid1.ROWS >= 14 Then
   
   MSFlexGrid1.ROWS = 13
End If

   

'MSFlexGrid1.AddItem ("Σύνολα" & Chr(9) & Format(SG, "####.##") & Chr(9) & Format(0, "####.##") & Chr(9) & Format(SM, "####.##"))
 MSFlexGrid1.AddItem ("Σύνολα" & Chr(9) & Format(SG, "####.##") & Chr(9) & Format(SG2, "####.##") & Chr(9) & Format(SM, "####.##") & Chr(9) & Format(SM2, "####.##"))
MSFlexGrid1.Visible = True

   
700
   
Adodc1.ConnectionString = gConnect


If F_EXO_XROMATA = 0 Then




'ΥΠΟΛΟΓΙΖΕΙ ΤΟ ΥΠΟΛΟΙΠΟ / ΑΝΤΙΑΚΑΤΑΣΤΑΘΗΚΕ ΜΕ ΕΝΑ LOOP ΠΑΡΑΚΑΤΩ ΓΙΑ ΝΑ ΜΗΝ ΑΡΓΕΙ
'Adodc1.RecordSource = "SELECT  round( (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" _
'& "  (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ),1) AS [ΥΠΟΛ]" _
'& " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' AND " + synt
'Adodc1.Refresh
' Text1(4).Text = Adodc1.Recordset(0)


Adodc1.RecordSource = "SELECT KODE,APOT,(select PERIGRAFH FROM PINAKES WHERE TYPOS=4 AND AYJON=APOT) AS [ΑΠΟΘ],ROUND(SUM(XRE),0) AS [ΕΙΣ],ROUND(SUM(PIS),0) AS [ΕΞΑΓ]," _
& " round( (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" _
& "  (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ),1) AS [ΥΠΟΛ]" _
& " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' AND " + synt + "  GROUP BY KODE,APOT"





Else

Adodc1.RecordSource = "SELECT KODE,LEFT(PROELEYSH,10) AS [ΧΡ-ΜΕΓ],APOT,(select PERIGRAFH FROM PINAKES WHERE TYPOS=4 AND AYJON=APOT) AS [ΑΠΟΘ],ROUND(SUM(XRE),0) AS [ΕΙΣ],ROUND(SUM(PIS),0) AS [ΕΞΑΓ]," _
& " round( (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" _
& "  (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ),1) AS [ΥΠΟΛ]" _
& " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' AND " + synt + " GROUP BY APOT,KODE,LEFT(PROELEYSH,10) ORDER BY APOT,KODE,LEFT(PROELEYSH,10)"
End If

On Error GoTo 0


Adodc1.Refresh
MSHFlexGrid1.ColWidth(0) = 1000


    Text1(4).Text = 0
Do While Not Adodc1.Recordset.EOF
    Text1(4).Text = Val(Text1(4).Text) + Adodc1.Recordset("ΥΠΟΛ")
    Adodc1.Recordset.MoveNext
Loop









 Sql2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" _
                   & "  (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Text1(0).Text + "'"
 On Error GoTo LATOS
 
On Error Resume Next
  R.Close
 
 
 
 R.Open Sql2, Gdb, adOpenForwardOnly, adLockReadOnly
 Label6.Caption = CNull(R(0))
 R.Close









MSHFlexGrid1.ColWidth(0) = 1




Adodc2.ConnectionString = gConnect
If F_EXO_XROMATA = 0 Then
    Adodc2.RecordSource = "select ERG from BARCODES where LEFT(KOD,14)='" + Trim(Text1(0).Text) + "'"
Else
    Adodc2.RecordSource = "select ERG,PROELEYSH from BARCODES where  LEFT(KOD,14)='" + Text1(0).Text + "'"
End If

Adodc2.Refresh

'DataGrid2.DataSource = timokat

' timokat.ConnectionString = gConnect
' timokat.RecordSource = "select * from TIMOKAT where  LEFT(KOD,14)='" + Text1(0).Text + "'"
' timokat.Refresh


Dim PI

If F_EIKONA > 0 Then
  On Error Resume Next
  If F_EIKONA = 2 Then 'βοηθητικοσ κωδικοσ
     PI = "C:\MERCVB\IMAGES\" + Text1(6).Text
  ElseIf F_EIKONA = 1 Then ' kvdikos
     PI = "C:\MERCVB\IMAGES\" + Text1(0).Text
  ElseIf F_EIKONA = 12 Then 'βοηθητικοσ κωδικοσ +.jpg
     PI = "C:\MERCVB\IMAGES\" + Text1(6).Text + ".JPG"
  ElseIf F_EIKONA = 11 Then ' kodikos .jpg
     PI = "C:\MERCVB\IMAGES\" + Text1(0).Text + ".JPG"
  ElseIf F_EIKONA = 111 Then ' kodikos .jpg
     PI = "C:\MERCVB\IMAGES\" + Trim(Text1(0).Text) + "*" + ".JPG"
  End If
  
 
 If Len(Dir(PI)) > 0 Then
    Picture1.Picture = LoadPicture("c:\mercvb\images\" + Dir(PI))
    Picture2.PaintPicture Picture1.Picture, 0, 0, Picture2.width, Picture2.height, 0, 0, Picture1.width, Picture1.height
 Else
    Picture1.Picture = LoadPicture()
    Picture2.Picture = LoadPicture()
 End If
   
End If

 
 
 
 




'DBGRID1.Columns(1).width = 4000




  NUM1 = IIf(IsNull(data2.Recordset("NUM1")), "", data2.Recordset("NUM1"))
  NUM2 = IIf(IsNull(data2.Recordset("NUM2")), "", data2.Recordset("NUM2"))
  NUM3 = IIf(IsNull(data2.Recordset("NUM3")), "", data2.Recordset("NUM3"))
  
  
  
  CH1 = IIf(IsNull(data2.Recordset("CH1")), "", data2.Recordset("CH1"))
  CH2 = IIf(IsNull(data2.Recordset("CH2")), "", data2.Recordset("CH2"))
  CH3 = IIf(IsNull(data2.Recordset("CH3")), "", data2.Recordset("CH3"))
  
  
  ch4 = IIf(IsNull(data2.Recordset("CH4")), "", data2.Recordset("CH4"))
  ch5 = IIf(IsNull(data2.Recordset("CH5")), "", data2.Recordset("CH5"))
  If IsNull(data2.Recordset("CH6")) Then ch6 = "" Else ch6 = data2.Recordset("CH6")
  

  
  
  
  
  

  HM1 = IIf(IsNull(data2.Recordset("HM1")), "", data2.Recordset("HM1"))
  HM2 = IIf(IsNull(data2.Recordset("HM2")), "", data2.Recordset("HM2"))
  HM3 = IIf(IsNull(data2.Recordset("HM3")), "", data2.Recordset("HM3"))



'===============================================================================
'==================== PALIA XRONIA =============================================
'========================================================================

If Check1.value = vbUnchecked Then
  Me.MousePointer = vbNormal
  If cXRONIES.value = vbChecked Then Gdb.Close: Gdb.Open gConnect
  Exit Sub
End If



Dim db2 As New ADODB.Connection
Dim PaLIAXRONIA As String
'On Error GoTo 0
PaLIAXRONIA = Trim(FindParametroi("PAR1", "F_PALIAXRONIA", "2005", "Γ.18 DSN ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ"))

If Len(PaLIAXRONIA) = 0 Then
     Me.MousePointer = vbNormal
     If cXRONIES.value = vbChecked Then Gdb.Close: Gdb.Open gConnect
     Exit Sub
End If


   On Error GoTo oxi_palia

   db2.Open gConnect + ";DATABASE=" + PaLIAXRONIA
   On Error Resume Next

   R.Close
   R.Open "SELECT MONTH(HME),SUM(XRE),SUM(PIS),SUM(XRE*TIMM),SUM(PIS*TIMM) FROM EGGTIM WHERE KODE='" + data2.Recordset("kod") + "' GROUP BY MONTH(HME) ", db2, adOpenForwardOnly, adLockReadOnly

   For k = 1 To 12
        A_XRE(k) = 0: A_PIS(k) = 0: A_XREAJ(k) = 0: A_PISAJ(k) = 0
   Next
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
        MSFLEXGRID2.Visible = False


SM2 = 0: SG2 = 0
SM = 0: SG = 0
For k = 1 To 12
  graph2.row = 1: graph2.Column = k: graph2.DATA = A_XRE(k) ' DATA2.Recordset(G)
  graph2.row = 2: graph2.Column = k: graph2.DATA = A_PIS(k) 'DATA2.Recordset(M)
   graph2.ColumnLabel = k
    graph2.Refresh
    
   
  
  MSFLEXGRID2.row = k
  MSFLEXGRID2.Col = 0
  MSFLEXGRID2.Text = k
    
  MSFLEXGRID2.Col = 1

  
  SM = SM + A_PIS(k) 'DATA2.Recordset(M)
  SG = SG + A_XRE(k) ' DATA2.Recordset(G)
  
  SM2 = SM2 + A_PISAJ(k) 'DATA2.Recordset(M)
  SG2 = SG2 + A_XREAJ(k) ' DATA2.Recordset(G)
  
  
  
  
  MSFLEXGRID2.Text = Format(A_XRE(k), "#####.##") ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset(G)), "####.##")

  
  MSFLEXGRID2.Col = 2
  MSFLEXGRID2.Text = Format(A_XREAJ(k), "#####.##") ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset("XTI") * DATA2.Recordset(G)), "#####,##")

  MSFLEXGRID2.Col = 3
  MSFLEXGRID2.Text = Format(A_PIS(k), "#####.##") 'Format(IIf(IsNull(DATA2.Recordset(M)), 0, DATA2.Recordset(M)), "#####,##")


  MSFLEXGRID2.Col = 4
  MSFLEXGRID2.Text = Format(A_PISAJ(k), "#####.##") 'Format(IIf(IsNull(DATA2.Recordset(MA)), 0, DATA2.Recordset("LTI") * DATA2.Recordset(MA)), "#####,##")
  
Next

If MSFLEXGRID2.ROWS >= 14 Then
   
   MSFLEXGRID2.ROWS = 13
End If
   MSFLEXGRID2.AddItem ("Σύνολα" & Chr(9) & Format(SG, "####.##") & Chr(9) & Format(SG2, "####.##") & Chr(9) & Format(SM, "####.##") & Chr(9) & Format(SM2, "####.##"))
   MSFLEXGRID2.Visible = True





















Me.MousePointer = vbNormal



If cXRONIES.value = vbChecked Then Gdb.Close: Gdb.Open gConnect

Exit Sub



NOCONNECT:
 If cXRONIES.value = vbChecked Then Gdb.Close: Gdb.Open gConnect
  Exit Sub


LATOSEXo:
If cXRONIES.value = vbChecked Then Gdb.Close: Gdb.Open gConnect
Exit Sub




LATOS:
MsgBox Err.Description

Resume Next


oxi_palia:
If cXRONIES.value = vbChecked Then Gdb.Close: Gdb.Open gConnect
Exit Sub




End Sub

Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mDBGrid1_RowColChange

End Sub

'Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 '  Dim d
'   d = DBGrid1.Text
' If f_Refresh Then
  ' mDBGrid1_RowColChange
' End If

'End Sub

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
  
      ono.Text = ""
      kod.Text = ""
      XTI.Text = ""
      
  
  
  
  
  Command5_Click
  
End Sub

Private Sub Form_DblClick()
   PARAMETROI.PARAM.Caption = "APOT2"
   PARAMETROI.SHOW 1
    'F_TAB = Val(FindParametroi("APOT2", "F_TAB", "3", "Σε ποιό TAB είναι σταματημένο"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim a
 a = 1
If KeyCode = 27 Then
   Command1_Click
End If

End Sub

Private Sub Form_Load()

mForm_Load Me, fh, fw, ft, fl

 Dim db, R As New ADODB.Recordset
 KeyPreview = True
Me.Picture = LoadPicture(gPicture)
' Command3.Picture = LoadPicture("c:\mercvb\page.gif")
alignGridTimon

'APOT2.SHOW

' f_Refresh = True

''SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", "NormalOffice2007.ini"

''SkinFramework1.ApplyWindow Me.hWnd
''SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
On Error Resume Next



If Len(Dir("c:\mercfan.txt")) > 0 Then
   Open "c:\mercfan.txt" For Input As #1
    Input #1, fConnect2 'DUMMY
    Input #1, fConnect2
   Close #1
   FGDB2.Open fConnect2

End If


' Dim R As New ADODB.Recordset
R.Open "select DSN,TITLOS FROM ETAIREIES ORDER BY TITLOS DESC", Gdb, adOpenDynamic, adLockOptimistic
R.MoveFirst
XRONIES.Clear
Do While Not R.EOF
     If Not IsNull(R("TITLOS")) And (Not IsNull(R("DSN"))) Then
           XRONIES.AddItem Left(R("TITLOS") + Space(20), 20) + R("DSN")
     End If
     R.MoveNext
Loop
R.Close








If Len(Dir("C:\LAGEURO\LAGAKIS.TXT", vbNormal)) > 0 Then
    Command15.Visible = True
    
    'Exit Sub
End If

F_CHANGETIMH = Val(FindParametroi("APOT2", "F_CHANGETIMH", "0", "ΑΛΛΑΖΕΙ ΑΥΤΟΜΑΤΑ ΤΗΝ ΧΟΝΔ.ΤΙΜΗ=1/0"))
F_UPPER = Val(FindParametroi("APOT2", "F_UPPER", "1", "ΜΕΤΑΤΡΟΠΗ ΠΕΡΙΓΡΑΦΩΝ ΣΕ ΚΕΦΑΛΑΙΑ ΕΛΛΗΝΙΚΑ=1/0"))


F_EIKONA = Val(FindParametroi("APOT2", "F_EIKONA", "1", "1=EIKONA=>ΚΩΔΙΚΟΣ 2=>ΜΕ ΒΟΗΘ.ΚΛΕΙΔΙ 11=ΚVDIKOS.JPG 12=BOH.JPG 0=ΟΧΙ ΕΙΚΟΝΕΣ"))


F_add_barc = Val(FindParametroi("PAR1", "add_barc", "1", "Γ.16 Barcode έιναι 1.Κωδ 2.Βοηθ.Κλειδί"))


f_YPOL_MHN = Val(FindParametroi("APOT2", "f_YPOL_MHN", "1", "Γ.1 ΠΡΟΕΠΙΛΟΓΗ ΥΠΟΛ.ΜΗΝΙΑΙΩΝ=1 / 0"))
 'As Integer

If f_YPOL_MHN = 1 Then
   Check1.value = vbChecked
Else
   Check1.value = vbUnchecked
   
End If




F_EXO_XROMATA = Val(FindParametroi("PAR1", "F_EXO_XROMATA", "0", "Γ.1 ΧΡΗΣΙΜΟΠΟΙΩ ΜΕΓΕΘΗ-ΧΡΩΜΑΤΑ=1")) 'posa psifia tha exei h kathe seira

f_FontName = FindParametroi("APOT2", "F_FONTNAME", "Arial", "FONT-NAME BARCODE") '"Arial"

   f_tab = Val(FindParametroi("APOT2", "F_TAB", "3", "Προεπιλεγμένο TAB "))
   SSTab1.Tab = IIf(f_tab > SSTab1.Tabs - 1 Or f_tab < 0, SSTab1.Tabs - 1, f_tab)
   
 Me.Picture = LoadPicture(gPicture)
 
  F_ArPerEtik = Val(FindParametroi("APOT2", "F_ArPerEtik", "2", "Αριστερό περιθώριο ετικέττας"))  'posa psifia tha exei h kathe seira
 
  F_2ArPerEtik = Val(FindParametroi("APOT2", "F_2ArPerEtik", "2", "2o Αριστερό περιθώριο ετικέττας"))  'posa psifia tha exei h kathe seira
 
 
  F_PIChEIGHT = Val(FindParametroi("APOT2", "F_PIChEIGHT", "2535", "Υψος εικόνας")) 'posa psifia tha exei h kathe seira
  F_PICwidth = Val(FindParametroi("APOT2", "F_PICwidth", "3675", "Πλάτος εικόνας")) 'posa psifia tha exei h kathe seira
 
  F_XBARCODE = Val(FindParametroi("APOT2", "F_XBARCODE", "348", "barcode Χ")) 'posa psifia tha exei h kathe seira
  F_YBARCODE = Val(FindParametroi("APOT2", "F_YBARCODE", "824", "barcode Y")) 'posa psifia tha exei h kathe seira
 
 
 Picture2.height = F_PIChEIGHT
 Picture2.width = F_PICwidth
 
  F_sygx = Val(FindParametroi("APOT2", "F_SYGX", "1", "ΕΠΙΤΡΕΠΕΤΑΙ ΣΥΓΧΩΝΕΥΣΗ ΚΩΔΙΚΩΝ=1 0=ΟΧΙ")) 'posa psifia tha exei h kathe seira
 
 
times.TextMatrix(0, 1) = "Tιμή Αγοράς"
times.TextMatrix(0, 2) = "Mark Up"
times.TextMatrix(0, 3) = "Χωρίς ΦΠΑ"
times.TextMatrix(0, 4) = "Με ΦΠΑ"

times.TextMatrix(1, 0) = "Χονδρ.Τιμή"
times.TextMatrix(2, 0) = "Λιαν.Τιμή"


 
 
 
 
 
F_KOD = Val(FindParametroi("APOT2", "F_KOD", "1", "Εμφανίζεται η αναζήτηση με κωδικό=1"))
F_ONO = Val(FindParametroi("APOT2", "F_ONO", "1", "Εμφανίζεται η αναζήτηση με ΟΝΟΜΑ=1"))
F_ERG = Val(FindParametroi("APOT2", "F_ERG", "1", "Εμφανίζεται η αναζήτηση με ΒΟΗΘ.ΚΛΕΙΔΙ=1"))
F_XTI = Val(FindParametroi("APOT2", "F_XTI", "1", "Εμφανίζεται η αναζήτηση με BARCODE=1"))

 
 
If F_KOD = 0 Then kod.Visible = False
If F_ONO = 0 Then ono.Visible = False
If F_ERG = 0 Then erg.Visible = False
If F_XTI = 0 Then XTI.Visible = False
 
 
 
 
 
 F_LHM1 = FindParametroi("APOT2", "F_lhm1", "Ημερ1", "Ετικέτα ημερ.1")
 F_LHM2 = FindParametroi("APOT2", "F_lhm2", "Ημερ2", "Ετικέτα ημερ.2")
 f_lHM3 = FindParametroi("APOT2", "F_lhm3", "Ημερ3", "Ετικέτα ημερ.3")
 lhm1.Caption = F_LHM1
 lhm2.Caption = F_LHM2
 lhm3.Caption = f_lHM3
 f_lab1 = FindParametroi("APOT2", "F_lab1", "Κειμ1", "Ετικέτα κειμ.1")
 f_lab2 = FindParametroi("APOT2", "F_lab2", "Κειμ2", "Ετικέτα κειμ.2")
 f_lab3 = FindParametroi("APOT2", "F_lab3", "Κειμ3", "Ετικέτα κειμ.3")
 
 f_lab4 = FindParametroi("APOT2", "F_lab4", "Κειμ4", "Ετικέτα κειμ.4")
 f_lab5 = FindParametroi("APOT2", "F_lab5", "Κειμ5", "Ετικέτα κειμ.5")
 f_lab6 = FindParametroi("APOT2", "F_lab6", "Κειμ6", "Ετικέτα κειμ.6")
 
 
 
 lab1.Caption = f_lab1
 lab2.Caption = f_lab2
 lab3.Caption = f_lab3
 
 lab4.Caption = f_lab4
 lab5.Caption = f_lab5
 lab6.Caption = f_lab6
 
 
 
 f_num1 = FindParametroi("APOT2", "F_num1", "Αριθ1", "Ετικέτα κειμ.1")
 f_num2 = FindParametroi("APOT2", "F_num2", "Αριθ2", "Ετικέτα κειμ.2")
 f_num3 = FindParametroi("APOT2", "F_num3", "Αριθ3", "Ετικέτα κειμ.3")
 Lnum1.Caption = f_num1
 Lnum2.Caption = f_num2
 Lnum3.Caption = f_num3
  
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
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

If gXEIRISTHS < 8 Then
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
     Combo4.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
  Else
     Exit Do
  End If
  R.MoveNext
Loop
R.Close



' ΚΑΤΗΓΟΡΙΕΣ

R.Open "SELECT *FROM PINAKES WHERE TYPOS=15 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
'R.MoveFirst

If R.EOF Then
   Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH) VALUES (15,1,'ΕΜΠΟΡΕΥΜΑ')"
End If

Do While Not R.EOF
  If R("typos") = 15 Then
     Combo5.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
  Else
     Exit Do
  End If
  R.MoveNext
Loop
R.Close



  




F_PONTOI = Val(FindParametroi("APOT2", "F_PONTOI", "1", "ΣΤΗΝ ΕΤΙΚΕΤΤΑ ΠΟΝΤΟΥΣ=1 ΟΧΙ=0"))



 DBGrid1.ColWidth(1) = 3300

DBGrid1.ColWidth(3) = 1250
DBGrid1.ColAlignment(3) = 7

DBGrid1.ColWidth(2) = 1250
DBGrid1.ColAlignment(2) = 7

 
 

DATA1.RecordSource = "SELECT TOP 100 *FROM EID"
DATA1.ConnectionString = gConnect
DATA1.Refresh

data2.ConnectionString = gConnect

data22.ConnectionString = gConnect
     data22.RecordSource = "select TOP 30  KOD,ONO,CONVERT(decimal(10,2),LTI) AS [ΛΙΑΝ.ΤΙΜΗ],CONVERT(decimal(10,2),XTI) AS [XONΔ.ΤΙΜΗ]  from EID where POS>0;"
     data22.Refresh
     



PROTH_FORA = 0



End Sub

Private Sub Option1_Click(Index As Integer)
  
End Sub

Private Sub Form_Paint()
 
 If PROTH_FORA = 0 Then
    PROTH_FORA = 1
On Error Resume Next
kod.SetFocus
 End If
   
End Sub



Private Sub Form_Resize()
      ' mForm_Resize Me, 14715, 11310,fh,fw,ft,fl
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
    
 ' If Len(Trim(kod.Text)) > 0 Then
      
      
      ono.Text = ""
      erg.Text = ""
      XTI.Text = ""
      
      Command5_Click
      
      
      
      
  'End If
  
End Sub

Private Sub Lab6_Click()
 
 Dim l_lngRetVal As Long
   Const SW_SHOWNORMAL = 1
   l_lngRetVal = ShellExecute(0&, vbNullString, ch6.Text, vbNullString, "C:\", SW_SHOWNORMAL)





End Sub

Private Sub Label1_Click(Index As Integer)
'----------- OIKOGENEIES ---------------------------
  Dim f As New DataEntry2, DD
  Dim R As New ADODB.Recordset
 
 
 'Sub UPDATE_PINAKES(MTYPOS As Integer, OO As Object, MCOMBO As Object)
'mtypos  αριθμος πίνακα
'οο: to object που θέλω για το κεντράρισμα της οθόνης DataEntry
'mcombo : το combo που θα γεμίσει με τις τιμές
 
 If Index = 2 Then 'monades metrhshs
    
    UPDATE_PINAKES 2, Label1(Index), Combo1
 End If
 If Index = 25 Then ' oikogeneies
    UPDATE_PINAKES 11, Label1(Index), Combo4
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
If Index = 4 Then 'YPO - OIKOGENEIES
  If Val(Left(Combo4.Text, 2)) = 0 Then
     Exit Sub
  End If
  

  'sql οπου θα γίνεται το update ή το addnew
  f.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=3 AND TIMH= " + Left(Combo4.Text, 2) '    sql
'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
f.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " _
& " FROM PINAKES WHERE TYPOS=3 AND TIMH=" + Left(Combo4.Text, 2) + " ORDER BY AYJON " 'ParamGrid   sql2

'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
f.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=3  AND TIMH=" + Left(Combo4.Text, 2) + " AND " 'ParamGrid              DelSQL

'ποιό update query θα δημιουργείται (προαιρετικά )
'αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
'                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
 

f.ParamGrid.TextMatrix(0, 6) = "TYPOS" 'EPIPLEON FIELD
f.ParamGrid.TextMatrix(0, 7) = 3 ' TIMH TOY EPIPLEON FIELD
 
 
f.ParamGrid.TextMatrix(0, 8) = "TIMH" '2 EPIPLEON FIELD
f.ParamGrid.TextMatrix(0, 9) = Left(Combo4.Text, 2) ' 2 TIMH TOY EPIPLEON FIELD
 
 
 
 
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
f.ParamGrid.TextMatrix(0, 4) = "" ' DD
f.ParamGrid.TextMatrix(1, 1) = "A/A"
f.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "

'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
f.Labels(0).width = f.Labels(0).width + 1000 '* TextWidth("A")
f.Textboxes(0).Left = f.Textboxes(0).Left + 1000 ' * TextWidth("A")
'f.Left = 0
'f.Top = 0
'f.height = 18000

f.Left = Label1(Index).Left
f.Top = Label1(Index).Top + MDIForm1.Top
f.SHOW 1



Combo3.Clear
 R.Open "SELECT * FROM PINAKES WHERE TYPOS=3 AND TIMH= " + Left(Combo4.Text, 2) + " ORDER BY AYJON ", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
     Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
     R.MoveNext
Loop
R.Close


End If


'  Me.WindowState = 2




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
 
      kod.Text = ""
      erg.Text = ""
      XTI.Text = ""
      
 
 
 
 
 
 
 Command5_Click
  'DBGrid1.SetFocus
End Sub

Private Sub TabStrip1_Click()
Dim a
 a = SSTab1.TabIndex
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
 ' If Index = 5 Or Index = 4 Then
 '      Text1(3).Text = Val(Text1(2).Text) * (100 + Val(Text1(5).Text)) / 100
 ' End If
  
 '  If Index = 3 And Val(Text1(2).Text) > 0 Then 'ΟΤΑΝ ΑΛΛΑΖΩ ΤΙΜΗ XONTRIKHS (3)  ΑΛΛΑΖΕΙ ΤΟ ΜΑΡΚΑΠ (5)      Η ΑΓΟΡΑ ΕΙΝΑΙ ΤΟ (2)
 '      Text1(5).Text = Val(Text1(3).Text) / Val(Text1(2).Text) * 100
 ' End If
  
  
  
 '  If Index = 13 And Val(Text1(2).Text) > 0 Then 'ΟΤΑΝ ΑΛΛΑΖΩ ΤΙΜΗ ΛΙΑΝΙΚΗΣ (13)  ΑΛΛΑΖΕΙ ΤΟ ΜΑΡΚΑΠ (14)      Η ΑΓΟΡΑ ΕΙΝΑΙ ΤΟ (2)
 '      Text1(14).Text = Val(Text1(13).Text) / (1 + g_Fpa(Val(Combo2.Text))) / Val(Text1(2).Text) * 100
 ' End If
  
End Sub

Private Sub Text1_LostFocus(Index As Integer)
  
    
   If Index = 13 And Val(Text1(2).Text) > 0 Then 'ΟΤΑΝ ΑΛΛΑΖΩ ΤΙΜΗ ΛΙΑΝΙΚΗΣ (13)  ΑΛΛΑΖΕΙ ΤΟ ΜΑΡΚΑΠ (14)      Η ΑΓΟΡΑ ΕΙΝΑΙ ΤΟ (2)
       'Text1(14).Text = (Val(Text1(13).Text) / Val(Text1(2).Text) - 1) * 100
       Text1(14).Text = Val(Text1(13).Text) / (1 + g_Fpa(Val(Combo2.Text)) / 100) / Val(Text1(2).Text) * 100 - 100
  End If
    
   If Index = 3 And Val(Text1(2).Text) > 0 Then 'ΟΤΑΝ ΑΛΛΑΖΩ ΤΙΜΗ XONTRIKHS (3)  ΑΛΛΑΖΕΙ ΤΟ ΜΑΡΚΑΠ (5)      Η ΑΓΟΡΑ ΕΙΝΑΙ ΤΟ (2)
       Text1(5).Text = (Val(Text1(3).Text) / Val(Text1(2).Text) - 1) * 100
       
       XondrmeFPA.Text = Val(Text1(3).Text) * (1 + g_Fpa(Val(Combo2.Text)) / 100)
  End If
      
     If Index = 5 Then
       If Val(Combo2.Text) > 0 Then
           proteinX.Caption = Val(Text1(2).Text) * (100 + Val(Text1(5).Text)) / 100
       End If
       
       
       If F_CHANGETIMH = 1 Then
          Text1(3).Text = Format(Val(Text1(2).Text) * (100 + Val(Text1(5).Text)) / 100, "###0.00")
          XondrmeFPA.Text = Format(Val(Text1(3).Text) * (1 + g_Fpa(Val(Combo2.Text)) / 100), "###0.00")
       End If
       
           
    
    
    End If
   
    Text1(Index).BackColor = vbWhite
    If Index = 14 Then
       If Val(Combo2.Text) > 0 Then
           proteinX.Caption = Val(Text1(2).Text) * (1 + g_Fpa(Val(Combo2.Text)) / 100) * (100 + Val(Text1(14).Text)) / 100
       End If
    End If
    
    End Sub
Sub alignGridTimon()
   'timh agoras
   Text1(2).Left = times.Left + times.CellWidth + 60
   Text1(2).Top = times.Top + times.CellHeight + 70
   Text1(2).width = times.CellWidth
   
   'markup xontrikhs
   Text1(5).Left = times.Left + 2 * times.CellWidth + 60
   Text1(5).Top = times.Top + times.CellHeight + 70
   Text1(5).width = times.CellWidth
   'timh xontrikhs
   Text1(3).Left = times.Left + 3 * times.CellWidth + 80
   Text1(3).Top = times.Top + times.CellHeight + 70
   Text1(3).width = times.CellWidth
   'timh xontrikhs me fpa
   XondrmeFPA.Left = times.Left + 4 * times.CellWidth + 90
   XondrmeFPA.Top = times.Top + times.CellHeight + 70
   XondrmeFPA.width = times.CellWidth
   
   'markup lianikhs
   Text1(14).Left = times.Left + 2 * times.CellWidth + 80
   Text1(14).Top = times.Top + 2 * times.CellHeight + 70
   Text1(14).width = times.CellWidth
   'timh lianikhs me fpa
   Text1(13).Left = times.Left + 4 * times.CellWidth + 90
   Text1(13).Top = times.Top + 2 * times.CellHeight + 50
   Text1(13).width = times.CellWidth
 
   Text1(13).height = times.CellHeight

End Sub
Private Sub Text14_GotFocus()
  Text14.BackColor = vbYellow
  Text14.ForeColor = vbBlack
End Sub

Private Sub Text14_LostFocus()
    Text14.BackColor = vbWhite
 '
End Sub

Private Sub XondrmeFPA_LostFocus()
 'YPOLOGIZEI TO POSOSTO OTAN ALLAZO THN TIMH ME FPA
    Text1(5).Text = Val(XondrmeFPA.Text) / (1 + g_Fpa(Val(Combo2.Text)) / 100) / Val(Text1(2).Text) * 100 - 100
End Sub

Private Sub xti_Change()
  ' data22.recordsource = "select *from EID where pos>0 and xti>" + Str(Val(xti)) + " order by xti;"
  ' Data2.Refresh
End Sub

Private Sub xti_GotFocus()
  XTI.BackColor = vbYellow
End Sub

Private Sub xti_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub xti_LostFocus()
   XTI.BackColor = vbWhite
   
      ono.Text = ""
      erg.Text = ""
      kod.Text = ""
      
   
   On Error Resume Next
   
   
   
   Command5_Click
   
End Sub





Sub MAKIS_ETIK()

Dim k, timologio

 Dim rec As New ADODB.Recordset



't = "c:\lageuro\makisget.exe "
't = t + par1.EID.Recordset("KOD") + " "
't = t + Replace(par1.EID.Recordset("ONO"), " ", "_") + " "
't = t + Format(Val(Text1(4).Text), "##0.00") + " "   ' LTI5  LIANIKI
't = t + Format(Val(Text1(2).Text), "##0") + " "  ' POS_KERD
'If Len(Trim(mERG)) > 0 Then
'   t = t + mERG
'Else
'   t = t + "00"
'End If




Dim MPER
MPER = data2.Recordset("ONO")



Dim mtimes, TIMM, pOS_KERD

   
   mtimes = 1
   TIMM = Format(Val(Text1(13).Text), "##0.00")
   MPER = Text1(1).Text
   
   pOS_KERD = Format(Val(Text1(14).Text), "##0") + " -" ' POS_KERD
   
'   timologio = par1.EID.Recordset("KOD") + Left(par1.Text1.Text, 6) + "=" + par1.PEL.Recordset("KOD")
 '  timologio = timologio + Format(Val(Label11.Caption) * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(par1.DTPicker1, "dd/mm/yyyy")
Dim mlabel

  Printer.FontBold = True

Printer.CurrentY = 30

  Printer.FONTSIZE = 10
  Printer.Print Tab(2); MPER

  Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
  Printer.FONTSIZE = 26

  Printer.CurrentY = 600
  Printer.Print "      " + TIMM + " €"
  Printer.FONTSIZE = 10
  Printer.CurrentY = 1450
  Printer.Print "  " + Text1(6).Text + "-" + Format(500 + Val(Text1(14).Text), "####0") + " -" + Text1(0).Text
  


  


  Printer.EndDoc





Exit Sub


FIND_N_NUMBER:
'************************

On Error GoTo 0
rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

rec.MoveFirst
rec.Move 10
  'rec.Edit
  If IsNull(rec("epel")) Then
      rec("epel") = 1
  Else
      rec("epel") = rec("epel") + 1
  End If
  rec.Update
rec.Close

  
  mlabel = Right("00000000" + LTrim(Str(rec("EPEL"))), 8)
  Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + DATA1.Recordset("KOD") + "'"
Return








End Sub

Private Sub mDBGrid1old_RowColChange() '(LastRow As Variant, ByVal LastCol As Integer)
Dim G, m, a, GA, MA, P, k, SM, SG, SM2, SG2
Dim mPos01 As Single, mPos02 As Single
Dim mPos03 As Single
Dim mPos04 As Single
On Error Resume Next
Dim Sql2 As String

' DATA2.Refresh

'  Me.Caption = Now
Text1(0).Text = data2.Recordset("kod")
Text1(1).Text = data2.Recordset("ONO")

Text1(16).Text = data2.Recordset("koderg")


Text1(4).Text = IIf(IsNull(data2.Recordset("pos")), " ", data2.Recordset("pos"))


Combo2.Text = data2.Recordset("FPA")
       
Combo5.Text = Combo5.List(Val(data2.Recordset("KATHGORIA")) - 1)
       
Combo1.Text = data2.Recordset("mon")

' Combo4.Text = Combo4.List(Val(Data2.Recordset("AEG")) - 1) 'OIKOGENEIA
Dim n As Integer
For n = 0 To Combo4.ListCount - 1
   If Val(Left(Combo4.List(n), 3)) = Val(data2.Recordset("AEG")) Then
      Combo4.Text = Combo4.List(n) 'OIKOGENEIA
      Exit For
   End If
Next





Dim R As New ADODB.Recordset

If Len(Combo4.Text) > 0 Then
 
 On Error GoTo LATOSEXo
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=3 AND TIMH=" + Left(Combo4.Text, 2) + "ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
 Combo3.Clear


Do While Not R.EOF

  If R("typos") = 3 Then
     Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
  Else
     Exit Do
  End If
  R.MoveNext
Loop
R.Close

End If

On Error Resume Next








Combo3.Text = Combo3.List(Val(data2.Recordset("kodlog")) - 1)
       
Text1(2).Text = data2.Recordset("xti")


'If IsNull(Data2.Recordset("pos_kerd2")) Then
   Text1(5).Text = nNull(data2.Recordset("pos_kerd2"))
'End If


'If IsNull(Data2.Recordset("pos_kerd")) Then
   Text1(14).Text = nNull(data2.Recordset("pos_kerd"))
'End If

Text1(6).Text = CNull(data2.Recordset("ERG"))
ONO2.Text = CNull(data2.Recordset("ono2"))

Text1(3).Text = data2.Recordset("lti")
Text1(13).Text = data2.Recordset("lti5")


XondrmeFPA.Text = data2.Recordset("lti") * (1 + g_Fpa(data2.Recordset("fpa")) / 100)
 
 If IsNull(data2.Recordset("MEMO")) Then
    Text14.Text = ""
 Else
    Text14.Text = data2.Recordset("MEMO")
 End If
 
 Text1(15).Text = data2.Recordset("PROMHU")    ' BONUS
       
       Text1(11).Text = data2.Recordset("kodSYNOD")
       Text1(12).Text = data2.Recordset("kodlogAG")
       
       Text1(7).Text = data2.Recordset("PROM")    'Right(Combo3.Text, 3)
        Text1(9).Text = data2.Recordset("SPA")   'Right(Combo3.Text, 3)
       
       Text1(10).Text = data2.Recordset("UES")    'Right(Combo3.Text, 3)
       Text1(8).Text = data2.Recordset("EPIUYP")

       exoda.Text = data2.Recordset("pos_kerd3")



If IsNull(data2.Recordset("POS01")) Then mPos01 = 0 Else mPos01 = data2.Recordset("POS01")

If IsNull(data2.Recordset("POS02")) Then mPos02 = 0 Else mPos02 = data2.Recordset("POS02")

If IsNull(data2.Recordset("POS03")) Then mPos03 = 0 Else mPos03 = data2.Recordset("POS03")
If IsNull(data2.Recordset("POS04")) Then mPos04 = 0 Else mPos04 = data2.Recordset("POS04")

' Label2.Caption = "ΤΕΛ.ΕΝΗΜΕΡΩΣΗ: KENΤΡ :" + Format(mPos01, "###0") + " ΥΠ.1: " + Format(mPos02, "###0") + " ΥΠ.3: " + Format(mPos03, "###0")

If Combo9.ListCount = 2 + 1 Then
   Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " + Trim(mID(Combo9.List(0), 3, 10)) + " :" _
   + Format(mPos01, "###0") + "   " + Trim(mID(Combo9.List(1), 3, 10)) + " :" _
   + Format(mPos02, "###0")
ElseIf Combo9.ListCount = 1 + 1 Then
   Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " + Trim(mID(Combo9.List(0), 3, 10)) + " :" _
   + Format(mPos01, "###0")
ElseIf Combo9.ListCount = 3 + 1 Then
   Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " + Trim(mID(Combo9.List(0), 3, 10)) + _
   " :" + Format(mPos01, "###0") + " - " + Trim(mID(Combo9.List(1), 3, 10)) + " :" + _
   "   " + Format(mPos02, "###0") + " - " + Trim(mID(Combo9.List(2), 3, 10)) + " :" _
   + Format(mPos03, "###0")
   
ElseIf Combo9.ListCount = 4 + 1 Then
   Label2.Caption = "ΤΕΛ.ΕΝΗΜ:  " _
   + Trim(mID(Combo9.List(0), 3, 10)) + " :" + Format(mPos01, "###0") + "   " _
   + Trim(mID(Combo9.List(1), 3, 10)) + " :  " + Format(mPos02, "###0") + "   " _
   + Trim(mID(Combo9.List(2), 3, 10)) + " :" + Format(mPos03, "###0") + "   " _
   + Trim(mID(Combo9.List(3), 3, 10)) + " :" + Format(mPos04, "###0")
   
End If




  Text1(18).Text = data2.Recordset("SYSKMAX") 'SYSKEYASIA
  Text1(17).Text = data2.Recordset("SYSKMIN") ' YPOSYSKEYASIA
  MSFlexGrid1.row = 0
  MSFlexGrid1.Col = 1
  MSFlexGrid1.Text = "Αγορές"
  MSFlexGrid1.Col = 2
  MSFlexGrid1.Text = "Αξία Αγορ."
  
  MSFlexGrid1.Col = 3
  MSFlexGrid1.Text = "Πωλήσεις"

  MSFlexGrid1.Col = 4
  MSFlexGrid1.Text = "Αξία πωλήσεων"

'' Γραφική παράσταση
's = 0

'For l = 1 To 12
   ' Exit Sub
 '  G.row = 1 + addit: G.Column = l: G.Data = Stat(l, addit2 + 1) - Stat(l, addit2 + 2)
  ' G.row = 4 + addit: G.Column = l: G.Data = Stat(l, addit2 + 4) - Stat(l, addit2 + 5)
'Next
If data2.Recordset.RecordCount = 0 Then Exit Sub


SM = 0: SG = 0
SM2 = 0: SG2 = 0
' Dim R As New ADODB.Recordset

On Error GoTo LATOS
R.Open "SELECT MONTH(HME),SUM(XRE),SUM(PIS),SUM(XRE*TIMM),SUM(PIS*TIMM) FROM EGGTIM WHERE KODE='" + data2.Recordset("kod") + "' GROUP BY MONTH(HME) ", Gdb, adOpenForwardOnly, adLockReadOnly
Dim A_XRE(12) As Single
Dim A_PIS(12) As Single
Dim A_XREAJ(12) As Single
Dim A_PISAJ(12) As Single

For k = 1 To 12: A_XRE(k) = 0: A_PIS(k) = 0: A_XREAJ(k) = 0: A_PISAJ(k) = 0: Next
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

For k = 1 To 12
'  G = "G" + Format(k, "00")
'  M = "M" + Format(k, "00")
'  GA = "g" + Format(k, "00")
'  MA = "M" + Format(k, "00")
  
  graph1.row = 1: graph1.Column = k: graph1.DATA = A_XRE(k) ' DATA2.Recordset(G)
  graph1.row = 2: graph1.Column = k: graph1.DATA = A_PIS(k) 'DATA2.Recordset(M)
   graph1.ColumnLabel = k
    graph1.Refresh
    
   
  
  MSFlexGrid1.row = k
  MSFlexGrid1.Col = 0
  MSFlexGrid1.Text = k
    
  MSFlexGrid1.Col = 1

  
  SM = SM + A_PIS(k) 'DATA2.Recordset(M)
  SG = SG + A_XRE(k) ' DATA2.Recordset(G)
  
  
  SM2 = SM2 + A_PISAJ(k) 'DATA2.Recordset(M)
  SG2 = SG2 + A_XREAJ(k) ' DATA2.Recordset(G)
  
  
  
  
  MSFlexGrid1.Text = Format(A_XRE(k), "#####.##") ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset(G)), "####.##")

  
  MSFlexGrid1.Col = 2
  MSFlexGrid1.Text = Format(A_XREAJ(k), "#####.##") ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset("XTI") * DATA2.Recordset(G)), "#####,##")

  MSFlexGrid1.Col = 3
  MSFlexGrid1.Text = Format(A_PIS(k), "#####.##") 'Format(IIf(IsNull(DATA2.Recordset(M)), 0, DATA2.Recordset(M)), "#####,##")


  MSFlexGrid1.Col = 4
  MSFlexGrid1.Text = Format(A_PISAJ(k), "#####.##") 'Format(IIf(IsNull(DATA2.Recordset(MA)), 0, DATA2.Recordset("LTI") * DATA2.Recordset(MA)), "#####,##")

  
  
  
  
  
  
  
Next

If MSFlexGrid1.ROWS >= 14 Then
   
   MSFlexGrid1.ROWS = 13
End If

   

'MSFlexGrid1.AddItem ("Σύνολα" & Chr(9) & Format(SG, "####.##") & Chr(9) & Format(0, "####.##") & Chr(9) & Format(SM, "####.##"))
 MSFlexGrid1.AddItem ("Σύνολα" & Chr(9) & Format(SG, "####.##") & Chr(9) & Format(SG2, "####.##") & Chr(9) & Format(SM, "####.##") & Chr(9) & Format(SM2, "####.##"))
MSFlexGrid1.Visible = True

   
   
   
Adodc1.ConnectionString = gConnect

If F_EXO_XROMATA = 0 Then
Adodc1.RecordSource = "SELECT KODE,APOT,(select PERIGRAFH FROM PINAKES WHERE TYPOS=4 AND AYJON=APOT) AS [ΑΠΟΘ],ROUND(SUM(XRE),0) AS [ΕΙΣ],ROUND(SUM(PIS),0) AS [ΕΞΑΓ]," _
& " round( (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" _
& "  (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ),1) AS [ΥΠΟΛ]" _
& " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' GROUP BY KODE,APOT"
Else

Adodc1.RecordSource = "SELECT KODE,PROELEYSH,APOT,(select PERIGRAFH FROM PINAKES WHERE TYPOS=4 AND AYJON=APOT) AS [ΑΠΟΘ],ROUND(SUM(XRE),0) AS [ΕΙΣ],ROUND(SUM(PIS),0) AS [ΕΞΑΓ]," _
& " round( (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" _
& "  (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ),1) AS [ΥΠΟΛ]" _
& " FROM EGGTIM  WHERE KODE='" + Text1(0).Text + "' GROUP BY APOT,KODE,PROELEYSH ORDER BY APOT,KODE,PROELEYSH"
End If
 Adodc1.Refresh



 Sql2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" _
                   & "  (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Text1(0).Text + "'"
 On Error GoTo LATOS
 
 R.Close
 
 R.Open Sql2, Gdb, adOpenForwardOnly, adLockReadOnly
 Label6.Caption = CNull(R(0))
 R.Close









MSHFlexGrid1.ColWidth(0) = 1




Adodc2.ConnectionString = gConnect
If F_EXO_XROMATA = 0 Then
    Adodc2.RecordSource = "select ERG from BARCODES where  LEFT(KOD,14)='" + Text1(0).Text + "'"
Else
    Adodc2.RecordSource = "select ERG,PROELEYSH from BARCODES where  LEFT(KOD,14)='" + Text1(0).Text + "'"
End If

Adodc2.Refresh

'DataGrid2.DataSource = timokat

' timokat.ConnectionString = gConnect
' timokat.RecordSource = "select * from TIMOKAT where  LEFT(KOD,14)='" + Text1(0).Text + "'"
' timokat.Refresh


Dim PI
PI = "C:\MERCVB\IMAGES\" + Text1(0).Text
If Len(Dir(PI)) > 0 Then
   Picture1.Picture = LoadPicture(PI)
   Picture2.PaintPicture Picture1.Picture, 0, 0, Picture2.width, Picture2.height, 0, 0, Picture1.width, Picture1.height
Else
   PI = "C:\MERCVB\IMAGES\" + Text1(0).Text + ".JPG"
   If Len(Dir(PI)) > 0 Then
      Picture1.Picture = LoadPicture(PI)
      Picture2.PaintPicture Picture1.Picture, 0, 0, Picture2.width, Picture2.height, 0, 0, Picture1.width, Picture1.height
   Else
      Picture1.Picture = LoadPicture()
      Picture2.Picture = LoadPicture()
   End If
   
   ' Picture2.PaintPicture Picture1.Picture, 0, 0, Picture2.width, Picture2.height, 0, 0, Picture1.width, Picture1.height
End If


'DBGRID1.Columns(1).width = 4000




  NUM1 = IIf(IsNull(data2.Recordset("NUM1")), "", data2.Recordset("NUM1"))
  NUM2 = IIf(IsNull(data2.Recordset("NUM2")), "", data2.Recordset("NUM2"))
  NUM3 = IIf(IsNull(data2.Recordset("NUM3")), "", data2.Recordset("NUM3"))
  
  
  
  CH1 = IIf(IsNull(data2.Recordset("CH1")), "", data2.Recordset("CH1"))
  CH2 = IIf(IsNull(data2.Recordset("CH2")), "", data2.Recordset("CH2"))
  CH3 = IIf(IsNull(data2.Recordset("CH3")), "", data2.Recordset("CH3"))
  
  
  ch4 = IIf(IsNull(data2.Recordset("CH4")), "", data2.Recordset("CH4"))
  ch5 = IIf(IsNull(data2.Recordset("CH5")), "", data2.Recordset("CH5"))
  If IsNull(data2.Recordset("CH6")) Then ch6 = "" Else ch6 = data2.Recordset("CH6")
  

  
  
  
  
  

  HM1 = IIf(IsNull(data2.Recordset("HM1")), "", data2.Recordset("HM1"))
  HM2 = IIf(IsNull(data2.Recordset("HM2")), "", data2.Recordset("HM2"))
  HM3 = IIf(IsNull(data2.Recordset("HM3")), "", data2.Recordset("HM3"))



'===============================================================================
'==================== PALIA XRONIA =============================================
'========================================================================

Dim db2 As New ADODB.Connection
Dim PaLIAXRONIA As String
On Error GoTo 0
PaLIAXRONIA = Trim(FindParametroi("PAR1", "F_PALIAXRONIA", "2005", "Γ.18 DSN ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ"))


On Error GoTo oxi_palia

db2.Open gConnect + ";DATABASE=" + PaLIAXRONIA
On Error Resume Next

R.Close
R.Open "SELECT MONTH(HME),SUM(XRE),SUM(PIS),SUM(XRE*TIMM),SUM(PIS*TIMM) FROM EGGTIM WHERE KODE='" + data2.Recordset("kod") + "' GROUP BY MONTH(HME) ", db2, adOpenForwardOnly, adLockReadOnly

For k = 1 To 12: A_XRE(k) = 0: A_PIS(k) = 0: A_XREAJ(k) = 0: A_PISAJ(k) = 0: Next
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
MSFLEXGRID2.Visible = False


SM2 = 0: SG2 = 0
SM = 0: SG = 0
For k = 1 To 12
  graph2.row = 1: graph2.Column = k: graph2.DATA = A_XRE(k) ' DATA2.Recordset(G)
  graph2.row = 2: graph2.Column = k: graph2.DATA = A_PIS(k) 'DATA2.Recordset(M)
   graph2.ColumnLabel = k
    graph2.Refresh
    
   
  
  MSFLEXGRID2.row = k
  MSFLEXGRID2.Col = 0
  MSFLEXGRID2.Text = k
    
  MSFLEXGRID2.Col = 1

  
  SM = SM + A_PIS(k) 'DATA2.Recordset(M)
  SG = SG + A_XRE(k) ' DATA2.Recordset(G)
  
  SM2 = SM2 + A_PISAJ(k) 'DATA2.Recordset(M)
  SG2 = SG2 + A_XREAJ(k) ' DATA2.Recordset(G)
  
  
  
  
  MSFLEXGRID2.Text = Format(A_XRE(k), "#####.##") ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset(G)), "####.##")

  
  MSFLEXGRID2.Col = 2
  MSFLEXGRID2.Text = Format(A_XREAJ(k), "#####.##") ' Format(IIf(IsNull(DATA2.Recordset(G)), 0, DATA2.Recordset("XTI") * DATA2.Recordset(G)), "#####,##")

  MSFLEXGRID2.Col = 3
  MSFLEXGRID2.Text = Format(A_PIS(k), "#####.##") 'Format(IIf(IsNull(DATA2.Recordset(M)), 0, DATA2.Recordset(M)), "#####,##")


  MSFLEXGRID2.Col = 4
  MSFLEXGRID2.Text = Format(A_PISAJ(k), "#####.##") 'Format(IIf(IsNull(DATA2.Recordset(MA)), 0, DATA2.Recordset("LTI") * DATA2.Recordset(MA)), "#####,##")

  
  
  
  
  
  
  
Next

If MSFLEXGRID2.ROWS >= 14 Then
   
   MSFLEXGRID2.ROWS = 13
End If
   MSFLEXGRID2.AddItem ("Σύνολα" & Chr(9) & Format(SG, "####.##") & Chr(9) & Format(SG2, "####.##") & Chr(9) & Format(SM, "####.##") & Chr(9) & Format(SM2, "####.##"))
   MSFLEXGRID2.Visible = True



























Exit Sub


LATOSEXo:
Exit Sub




LATOS:
MsgBox Err.Description

Resume Next


oxi_palia:
Exit Sub




End Sub


'If Len(Dir("C:\FANAR.EXE", vbNormal)) > 0 Then
'   fanar_etik
'End If

Sub fanar_etik()
Dim X, X2
Dim barc(500)
Dim apoSCANNER



    'yupie
    'fores = Val(InputBox("Πόσα κομμάτια από το " + Text1(1).Text, , 1))
    Dim a$
    
    X = 0
    
  If Scanner Then
'
      Open "C:\MERCVB\APOG.TXT" For Input As #1
      X = 0
      
      Do While Not EOF(1) 'True
          On Error Resume Next
          Line Input #1, a$
          
          kod = mID(a$, 1, 13)
         ' pos = Val(mID$(a$, 15, 20))
         If Len(Trim(Left(kod, 3))) = 0 Then
          ' TIPOTA
        Else
          X = X + 1
          barc(X) = kod
        End If
        
        
       Loop
       Close #1

          If X = 0 Then
               MsgBox "δεν υπάρχουν εγγραφές"
               Exit Sub
           End If
         apoSCANNER = X
  End If
  
       
' ΑΝ ΔΕΝ ΕΙΝΑΙ ΑΠΟ ΣΚΑΝΕΡ
If X = 0 Then
   apoSCANNER = 1
   barc(1) = DataGrid1.Text  ' Text1(6).Text
End If
    
    
    
    
    Dim mkod
    

Dim mFORES As Long


Dim R As New ADODB.Recordset
Dim nn
Dim ll
Dim mONOMA, mERG, mLTI5
For ll = 1 To apoSCANNER
    
    
   R.Open "select BARCODES.ERG,EID.* FROM BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD  WHERE BARCODES.ERG='" + barc(ll) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   
   If R.EOF Or Len(Trim(barc(ll))) = 0 Then
     mFORES = 0
   Else
     mONOMA = R("ONO") 'ONOMA
     mkod = R("KOD") '  ' BARCODE
     
     mLTI5 = Format(R("LTI5"), "###0.00") ' LTI5
     mERG = R(0) ' erg
     
   End If
   R.Close

    
    
    
    
    
    
    
    If Len(mERG) < 12 Then
      ' MsgBox "ΑΚΑΤΑΛΛΗΛΟ BARCODE"
      ' Exit Sub
    End If
    
    If Len(mERG) = 12 Then
         BarCode3.SymbologyID = UPCA
             
    ElseIf Len(mERG) = 13 Then
         BarCode3.SymbologyID = EAN13
    Else
         BarCode3.SymbologyID = CODE128 ' UCC128 '  EAN8 '  ' "128"  ' UPCA
         BarCode3.Code128CharSet = Set_C
    End If
    
    For Each X2 In Printers
        If InStr(X2.DeviceName, "2844") > 0 Then
          ' Set printer as system default.
        
          Set Printer = X2
          ' Stop looking for a printer
          Exit For
        End If
    Next
    
    BarCode3.DataToEncode = mERG
    
    
   'For K = 1 To fores
    
    Printer.CurrentX = F_ArPerEtik * 100
    Printer.CurrentY = -150
    Printer.FONTSIZE = 10
     Printer.FontBold = True
    Printer.FontName = "COURIER" ' "Arial Greek"
    
    
 '   Printer.Print ""
    Printer.Print Tab(3); Left(UCase(mONOMA), 24) 'perigrafh
    Printer.Print Tab(3); mID(UCase(mONOMA), 25, 25) 'perigrafh
    
    
     Printer.CurrentY = 1300 ' 1600 ΟΚ 10-4-2010
    Printer.CurrentX = 2000
    Printer.FONTSIZE = 8
    Printer.Print mERG 'kodikos
    
    
    Printer.CurrentY = 1600 ' 1600 ΟΚ 10-4-2010
    Printer.CurrentX = 2000
    Printer.FONTSIZE = 8
    
    Printer.Print mkod  'kodikos
   
   
   If Scanner = False And ll = 1 Then
      FGDB2.Execute "update Material  set retailprice=" + Text1(13).Text + "  WHERE  Code='" + Text1(0).Text + "'", nn
      'FGDB2.Execute "update Material  set retailprice=" + mLti5 + "  WHERE  Code='" + mKod + "'", nn
      If nn = 0 Then
         MsgBox "δεν δενημερωθηκε το singular"
      End If
   End If
   
    
   ' Printer.CurrentX = 2048
    'Printer.Print s
    Printer.CurrentX = 400 + F_ArPerEtik * 100
    'Printer.FontSize = 12
    Printer.FONTSIZE = 10
    
    
    Printer.CurrentX = 200 + F_ArPerEtik * 100
        Printer.CurrentY = 700 ' 600
    
    Printer.PaintPicture BarCode3.Picture, F_ArPerEtik * 20, 1200, , 500
    
      Printer.CurrentY = 600 ' 600
    Printer.CurrentX = 1000 + F_ArPerEtik * 100
    Printer.FONTSIZE = 30
    Printer.FontBold = True
    Printer.Print mLTI5 + " €" ' timh  ' Format(Text1(13).Text, "##0.00")
    ' Printer.Print "4.55 €" ' timh
    Printer.EndDoc
  Next
 If Scanner Then
    Kill "C:\MERCVB\APOG.TXT"
 End If

 
    
   Exit Sub

End Sub

