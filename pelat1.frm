VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form pelat1 
   BackColor       =   &H00FF0000&
   Caption         =   "Õ›· ≈„„Ò·ˆﬁ"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18945
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   18945
   WindowState     =   2  'Maximized
   Begin VB.TextBox BRANCH 
      Height          =   375
      Left            =   7320
      TabIndex        =   65
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox STREETNUMBER 
      Height          =   495
      Left            =   7320
      TabIndex        =   64
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox ARKYKL 
      Height          =   375
      Left            =   3600
      TabIndex        =   58
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox VCAT 
      Height          =   375
      Left            =   5280
      TabIndex        =   57
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox VFACT 
      Height          =   375
      Left            =   6960
      TabIndex        =   56
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdKatax 
      Caption         =   "Katax"
      Height          =   360
      Left            =   8880
      TabIndex        =   55
      Top             =   4200
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmd–ÂÒÈÛÛ¸ÙÂÒ· 
      Caption         =   "–ÂÒÈÛÛ¸ÙÂÒ· =>"
      Height          =   360
      Left            =   8520
      TabIndex        =   54
      Top             =   7200
      Width           =   1500
   End
   Begin VB.TextBox C1 
      Height          =   375
      Left            =   1680
      TabIndex        =   52
      Top             =   8280
      Width           =   2175
   End
   Begin VB.TextBox COUNTRY 
      Height          =   285
      Left            =   5400
      TabIndex        =   50
      Text            =   "GR"
      Top             =   7080
      Width           =   495
   End
   Begin VB.CommandButton cmdCommand3 
      Caption         =   "≈’—≈”« ¡÷Ã Ã≈”Ÿ √√–”"
      Height          =   360
      Left            =   3600
      TabIndex        =   48
      Top             =   3240
      Width           =   2295
   End
   Begin VB.ComboBox PELPROM 
      Height          =   315
      ItemData        =   "pelat1.frx":0000
      Left            =   4560
      List            =   "pelat1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   330
      Width           =   3012
   End
   Begin VB.CommandButton cmdEYRESHAfm 
      Caption         =   "≈’—≈”« ¡÷Ã Ã≈”Ÿ √√–”"
      Height          =   375
      Left            =   8880
      TabIndex        =   46
      Top             =   10080
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.TextBox PATR 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6135
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox MIKRO 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox EPO1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "pelat1.frx":0024
      Left            =   5760
      List            =   "pelat1.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   7920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   612
      Left            =   5775
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   14
      Left            =   1680
      TabIndex        =   39
      Top             =   7800
      Width           =   2175
   End
   Begin VB.ComboBox KATHGORIA 
      Height          =   315
      ItemData        =   "pelat1.frx":0041
      Left            =   5415
      List            =   "pelat1.frx":0043
      TabIndex        =   24
      Top             =   6585
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "pelat1.frx":0045
      Left            =   1680
      List            =   "pelat1.frx":0055
      TabIndex        =   23
      Top             =   6600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataSource      =   "DATA1"
      Height          =   375
      Index           =   13
      Left            =   1680
      TabIndex        =   25
      Top             =   6960
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1680
      TabIndex        =   26
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CheckBox kenoi 
      BackColor       =   &H00FF0000&
      Caption         =   "≈ıÒÂÛÁ ÍÂÌÔ˝ ·ÒÈËÏÔ˝"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   5775
      TabIndex        =   34
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   4350
      MaxLength       =   5
      TabIndex        =   11
      Top             =   2280
      Width           =   1536
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   6120
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   1680
      TabIndex        =   22
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   1680
      TabIndex        =   21
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   1680
      TabIndex        =   20
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   1680
      TabIndex        =   19
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   18
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ›· ≈„„Ò·ˆﬁ"
      Height          =   612
      Left            =   5775
      TabIndex        =   0
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   17
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   15
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   13
      Text            =   "GR"
      Top             =   2760
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1680
      MaxLength       =   200
      TabIndex        =   6
      Top             =   960
      Width           =   4212
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6975
      Top             =   9210
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5895
      Top             =   5280
      _ExtentX        =   979
      _ExtentY        =   979
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pelat1.frx":0094
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ˘‰.’ÔÍ·Ù."
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
      Height          =   195
      Index           =   21
      Left            =   6120
      TabIndex        =   66
      Top             =   2280
      Width           =   1125
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡ÒÈËÏ¸Ú"
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
      Index           =   19
      Left            =   6240
      TabIndex        =   63
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LARKYKL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡—. ’ À"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3600
      TabIndex        =   62
      Top             =   3840
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblCAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ¡‘«√œ—…¡"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5280
      TabIndex        =   61
      Top             =   3840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ã¡— ¡ ¡’‘."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6960
      TabIndex        =   60
      Top             =   3840
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblvID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "........"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7080
      TabIndex        =   59
      Top             =   4800
      Width           =   360
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ƒ≈«"
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
      Height          =   195
      Index           =   15
      Left            =   120
      TabIndex        =   53
      Top             =   8400
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"pelat1.frx":01A6
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1950
      Left            =   9840
      TabIndex        =   51
      Top             =   1920
      Width           =   1920
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "◊˛Ò·(2)"
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
      Left            =   4320
      TabIndex        =   49
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label LEpo1 
      BackStyle       =   0  'Transparent
      Caption         =   "≈˛ÌıÏÔ"
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
      Left            =   1560
      TabIndex        =   45
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Lpatr 
      BackStyle       =   0  'Transparent
      Caption         =   "–·ÙÒ˛ÌıÏÔ"
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
      Left            =   5280
      TabIndex        =   44
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lmikro 
      BackStyle       =   0  'Transparent
      Caption         =   "OÌÔÏ·"
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
      Left            =   3840
      TabIndex        =   43
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   40
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      BackStyle       =   0  'Transparent
      Caption         =   " ·ËÂÛÙ˛Ú ÷–¡"
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
      Index           =   16
      Left            =   120
      TabIndex        =   38
      Top             =   6690
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   17
      Left            =   120
      TabIndex        =   37
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      BackStyle       =   0  'Transparent
      Caption         =   "‘Ò¸ÔÚ –ÎÁÒ˘ÏﬁÚ"
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
      Index           =   18
      Left            =   120
      TabIndex        =   36
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      BackStyle       =   0  'Transparent
      Caption         =   " ·ÙÁ„ÔÒﬂ·"
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
      Index           =   20
      Left            =   4320
      TabIndex        =   35
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      BackStyle       =   0  'Transparent
      Caption         =   "‘. ."
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
      Left            =   3945
      TabIndex        =   33
      Top             =   2325
      Width           =   450
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   32
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   120
      TabIndex        =   31
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   30
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      BackStyle       =   0  'Transparent
      Caption         =   "≈ÍÙ˘ÛÁ %"
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
      TabIndex        =   29
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   28
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      BackStyle       =   0  'Transparent
      Caption         =   "≈‹„„ÂÎÏ·"
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
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
      BackStyle       =   0  'Transparent
      Caption         =   "ƒÈÂ˝ËıÌÛÁ"
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
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lbl¡ÒÈËÏ¸Ú 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "pelat1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PELPROM_STRING As String
Dim F_FORMAT_PEL_KOD As String

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Dim f_TIMOKAT As Integer

Dim f_anal_epon As Integer

Dim F_PALIO_FINDAFM As Integer


Dim f_MONO_PELATES_KATHGORIAS As Integer

Dim f_EYRESHAFM_USER As String
Dim f_EYRESHAFM_PWD As String
Dim f_EYRESHAFM_AFM As String

' Dim f_anal_epon As Integer
' Dim f_anal_epon As Integer


Dim f_lab1 As String
Dim F_ISCAR As Integer

Dim F_artam As String


Dim f_KENTRA_ADYNATISMATOS As Integer


Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.key
        Case "Save"
            'ToDo: Add 'Save' button code.
            MsgBox "Add 'Save' button code."
    End Select
End Sub


Function Update_Next_Pel(meidos As String, ByVal KODIK As String) As Integer

        '<EhHeader>
        On Error GoTo Update_Next_Pel_Err

        '</EhHeader>

        Dim DB As Database

        Dim R  As New ADODB.Recordset

        On Error Resume Next

        'Set db = OpenDatabase(gDir, False, False, gConnect)
        ' R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic

        'R.Edit
100     If meidos = "e" Then
110         Gdb.Execute "update MEM SET EPEL=" + str(Val(KODIK)) + " WHERE ID=1"
            ' R("epel") = Val(Text1(0))
        Else
    
120         Gdb.Execute "update MEM SET RPEL=" + str(Val(KODIK)) + " WHERE ID=1"
            ' R("rpel") = Val(Text1(0))
        End If

130     Update_Next_Pel = Val(KODIK) ' mForm_Resize Me,8, 14715, 11310, fh, fw, ft, fl

        '<EhFooter>
        Exit Function

Update_Next_Pel_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.Update_Next_Pel " & _
         "at line "7105870 & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Update_Next_Pel " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function get_next_pel(meidos As String, ByVal BriskoKena As Integer, ByVal FORMAT_PEL_KOD As String) As Long

'documentaries


        '<EhHeader>
        On Error GoTo get_next_pel_Err

        '</EhHeader>
        Dim DB As Database

        Dim R  As New ADODB.Recordset

        Dim r0 As New ADODB.Recordset

        Dim k  As Long

        On Error Resume Next

        'Set db = OpenDatabase(gDir)
        'Set db = OpenDatabase(gDir, False, False)
        'Set db = OpenDatabase(gDir, False, False, gConnect)




 ' ◊Ÿ—…Ú ≈’—≈”«  ≈Õœ’
If BriskoKena = 0 Then '  kenoi.Value = vbUnchecked Then

400     R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
410     If meidos = "e" Then
420         get_next_pel = R("epel") + 1
        Else
430         get_next_pel = R("rpel") + 1
        End If
        R.Close
        
       ' F_FORMAT_PEL_KOD = "0000"
        
        R.Open "SELECT * FROM PEL WHERE KOD='" + Format(get_next_pel, FORMAT_PEL_KOD) + "' AND EIDOS='" + meidos + "'", Gdb, adOpenDynamic, adLockOptimistic
        If R.EOF Then
           R.Close
           Exit Function
        End If
        R.Close
        
        '‚ÒÂËÁÍÂ ÙÂÙÔÈÛ Í˘‰ÈÍÔÛ
                        ' SELECT KOD FROM PEL WHERE LEFT(KOD,1)='" + Left(FORMAT_PEL_KOD, 1) + "' AND LEN(KOD)=" + str(Len(FORMAT_PEL_KOD)) + " AND EIDOS='" + meidos + "'   ORDER BY KOD DESC "
             Dim CC As String
             
             If FORMAT_PEL_KOD = "0000" Then
                CC = "4"
             Else
                CC = Format(Len(FORMAT_PEL_KOD), "0")
             End If
                
                        
           R.Open "select MAX(KOD) FROM PEL WHERE EIDOS='e' AND LEN(KOD)=" + CC + " AND LEFT(KOD,1) IN ('1','2','3','4','5','6','7,','8','9','0') ", Gdb, adOpenDynamic, adLockOptimistic
        If R.EOF Then
           R.Close
           Exit Function
        End If
        
        get_next_pel = Val(R(0)) + 1
        
        
        R.Close
        
        
        Exit Function
        
        
        
        
   '≈¡Õ œ ≈–œÃ≈Õœ ¡–œ ‘œ ≈–≈À ’–¡—◊≈… ¬À≈–≈… ‘œÕ Ã≈√…”‘œ  ¡… –¡≈… +1
   
        
        
        
        Dim N34 As Integer
        
        R.Open "SELECT TOP 1 KOD FROM PEL WHERE (KOD NOT LIKE '9%')  AND LEN(KOD)=" + str(Len(FORMAT_PEL_KOD)) + " ORDER BY KOD DESC"
        If IsNull(R(0)) Then
           get_next_pel = 0
        Else
           get_next_pel = Val(R(0)) + 1
           
        End If
        
       ' get_next_pel = N34 ' Right$("000000" + Format(N34, "000000000") + LTrim$(str(N34)), Len(F_FORMAT_PEL_KOD))
        
        'AFOY YPARXEI CAJE KENO KVDIKO
      '  kenoi.Value = vbChecked


End If












100     If BriskoKena = 1 Then    'kenoi.Value = vbChecked Then
110         R.Open "select KOD FROM PEL WHERE LEN(KOD)=4 AND EIDOS='" + meidos + "' order by KOD", Gdb, adOpenForwardOnly, adLockReadOnly

120         If R.EOF Then
130             get_next_pel = FORMAT_PEL_KOD  '"0000"
140             R.Close: Exit Function
            End If

150         Me.MousePointer = vbHourglass

160         R.MoveFirst
170         k = 0

180         Do While Not R.EOF

190             If Val(R(0)) = 0 Then
200                 R.MoveNext
                Else
210                 k = k + 1

220                 If Left(R(0), 4) = Format(k, "0000") Then
230                     R.MoveNext
                    Else

                        On Error GoTo get_next_pel_Err

240                     r0.Open "select COUNT(*) FROM PEL WHERE KOD='" + Format(k, "0000") + "' AND EIDOS='" + meidos + "' ", Gdb, adOpenForwardOnly, adLockReadOnly

250                     If r0(0) > 0 Then    '«ƒ« ’–¡—◊≈…
260                         r0.Close

                            Exit Do

                        Else
270                         get_next_pel = Format(k, "0000")
280                         R.Close
290                         Me.MousePointer = vbNormal

                            Exit Function

                        End If

                    End If
                End If

            Loop

300         R.Close

            Dim ll, L

            '

310         For L = k + 1 To 9999
320             R.Open "select COUNT(*) FROM PEL WHERE KOD='" + Format(L, "0000") + "' AND EIDOS='" + meidos + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
330             ll = R(0)
340             R.Close

350             DoEvents
360             Me.Caption = k

370             If ll = 0 Then

                    Exit For

                End If

            Next

380         get_next_pel = Format(L, "0000")
390         Me.MousePointer = vbNormal

            Exit Function

        End If

        
        
        
        
        
        
        

        '<EhFooter>
        Exit Function

get_next_pel_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.get_next_pel " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.get_next_pel " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Sub PALIO_FIND_AFM()
'On Error Resume Next
'
'   If Len(Trim(Text1(5).Text)) = 9 Then
'      'OK       Form1.parametros.Caption = "http://lagakis.gr/greek-vat/web-vat-details.php?myafm=028783755&afm=" + Trim(Text1(5).Text) ' 028783755"
'      Dim u As String, p As String, AFM As String
'      u = Trim(f_EYRESHAFM_USER)   ' = "LAGAKIS23041960"  ' Ww367868u916
'      p = Trim(f_EYRESHAFM_PWD)  '  = "d389833921"   '080353
'      AFM = Trim(f_EYRESHAFM_AFM)    ' = "028783755"   '051103283
'
'
'
'   '   u = "Ww367868u916"
'    '  p = "080353"
'    '  AFM = "051103283"
'
'
'
'
'
'       Form1.parametros.Caption = "http://lagakis.gr/greek-vat/web-vat-details.php?u=" + u + "&p=" + p + "&myafm=" + AFM + "&afm=" + Trim(Text1(5).Text) ' 028783755"
'
'
'      ' Form1.SHOW
'       MILSEC 2000
'       Form1.Command4_Click
'       MILSEC 2000
'       Form1.Command1_Click
'       Form1.GEMISE_TEXT 1, Me
'
'   End If
'   'Form1.SHOW
'
'End Sub


Public Sub neo_FIND_AFM(ByVal askafm As String, ByVal ISNEW As Integer)
      Dim u As String, p As String, AFM As String
100   u = Trim(f_EYRESHAFM_USER)   ' = "LAGAKIS23041960"  ' Ww367868u916
102   p = Trim(f_EYRESHAFM_PWD)  '  = "d389833921"   '080353
104  AFM = Trim(f_EYRESHAFM_AFM)    ' = "028783755"   '051103283
      Dim EPO As String, DIE As String, EPA As String, pol As String, tk As String, doy As String
      Dim c As String
      c = askafm
      Dim ENERGOS As Integer
      ENERGOS = 1
      
      Par7MyData.requestafm2 u, p, c, EPO, DIE, EPA, pol, tk, doy, ENERGOS
      
      
  Dim sos As Integer
  If ENERGOS = 0 Then
     sos = 1
  End If
  If ENERGOS = -1 Then
     MsgBox "À¡»œ” ¡÷Ã"
     Exit Sub
  End If
      
If sos = 1 Then
    pelat1.Text1(1) = "ÏÁ Â„ÍıÒÔ ¡÷Ã"
 Else
    If ISNEW = 1 Then
       Me.Text1(1) = Replace(EPO, """", "")
       Me.Text1(2) = Replace(DIE, """", "")
       Me.Text1(3) = Replace(pol, """", "")
       Me.Text1(4) = Replace(EPA, """", "")
       Me.Text1(7) = Replace(doy, """", "")
       Me.Text1(7) = Replace(Me.Text1(7), "'", "`")
       Me.Text1(12) = Replace(tk, """", "")
    Else
'      Dim ff As New pelat2
'       ff.Text1(1) = Replace(EPO, """", "")
'       ff.Text1(2) = Replace(DIE, """", "")
'       ff.Text1(3) = Replace(pol, """", "")
'       ff.Text1(4) = Replace(EPA, """", "")
'       ff.Text1(7) = Replace(doy, """", "")
'       ff.Text1(7) = Replace(ff.Text1(7), "'", "`")
'       ff.Text1(13) = Replace(tk, """", "")
    End If
    
 End If
      
      
      

End Sub

Private Sub cmdCommand3_Click()

'
'  If F_PALIO_FINDAFM = 1 Then
'
'      PALIO_FIND_AFM
'      Exit Sub
'
'  End If
   
       Dim u As String, p As String, AFM As String
100   u = Trim(f_EYRESHAFM_USER)   ' = "LAGAKIS23041960"  ' Ww367868u916
102   p = Trim(f_EYRESHAFM_PWD)  '  = "d389833921"   '080353
104   AFM = Trim(f_EYRESHAFM_AFM)    ' = "028783755"   '051103283
      PutGGPS u

 ' If F_PALIO_FINDAFM = 0 Then
  
      neo_FIND_AFM Trim(Text1(5).Text), 1
      Exit Sub
  
'  End If





       Dim XMLServer As Object
       Dim XMLReceive As Object
       Dim v As String
   
   
    On Error GoTo cmdCommand3_err

  ' MsgBox "99"

      
       Dim url2 As String
   
    '  MsgBox "104"
   
   
       '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
        '' Url2 = UrlAADE + "SendInvoices" '/RequestDocs"
106   url2 = "https://lagakis.gr/greek-vat/web-vat-details.php?u=" + u + "&p=" + p + "&myafm=" + AFM + "&afm=" + Trim(Text1(5).Text) ' 028783755"
       'initialize
108     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
110     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
112     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
114    ' XMLServer.Option(9) = 2048
116     XMLServer.Option(6) = True
    
        ' XMLServer.Open "POST", Url2, False
118     XMLServer.Open "GET", url2, False  ' & "?mark=400000019698028", False
       ' XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
       ' XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
120     XMLServer.send 'TXML 'strData  '”‘≈ÀÕŸ ‘œ STRING TXML STO MYDATA  'XMLServer.send TXML '  ¡Õ »≈ÀŸ Õ¡ ‘œ ”‘≈…ÀŸ ¡–œ ¡—◊≈…œ ‘œ‘≈ ¬¡∆Ÿ strData
       ' Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
122  v = XMLServer.responseText

' MsgBox "122"
    '
    Dim aaa As String
    '
124 aaa = UnicodeTOGreek(v)
 'MsgBox "124"
126 gemiseFields 1, aaa, Me
 'MsgBox "126"

    '"RGWSPUBLICBASICRT_OUT":"AFM":"028783755   ",
    '"STOPDATE":NULL,"POSTALADDRESSNO":"21       ",
    '"DOYDESCR":"ƒ—¡Ã¡”","DOY":"5111",
    '"ONOMASIA":"À¡√¡ «”  √≈Ÿ—√…œ” Õ… œÀ¡œ”",
    '"LEGALSTATUSDESCR":NULL,"REGISTDATE":
    '"1900-01-01T00:00:00.000+02:00",
    '"DEACTIVATIONFLAG":"1",
    '"DEACTIVATIONFLAGDESCR":"≈Õ≈—√œ” ¡÷Ã          ",
    '"POSTALADDRESS":"  –¡À¡…œÀœ√œ’","FIRMFLAGDESCR":"≈–…‘«ƒ≈’Ã¡‘…¡”      ","COMMERTITLE":NULL,"POSTALAREADESCRIPTION":"ƒ—¡Ã¡",
    '"INIFLAGDESCR":"÷–   ","POSTALZIPCODE":"66100","ARRAYOFRGWSPUBLICFIRMACTRT_OUT":"RGWSPUBLICFIRMACTRTUSER":["FIRMACTDESCR":"◊œÕƒ—… œ ≈Ã–œ—…œ Ã≈—ŸÕ  ¡… ≈Œ¡—‘«Ã¡‘ŸÕ ’–œÀœ√…”‘… ŸÕ Ã«◊¡ÕŸÕ","FIRMACTKIND":"1","FIRMACTKINDDESCR":" ’—…¡",
    '"FIRMACTCODE":"46511005","FIRMACTDESCR":"’–«—≈”…≈” ≈√ ¡‘¡”‘¡”«” –¡ ≈‘ŸÕ Àœ√…”Ã… œ’ « ≈–≈ ‘¡”≈ŸÕ","FIRMACTKIND":"2","FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"95111001","FIRMACTDESCR":"’ƒ—¡’À… ≈”  ¡…  À…Ã¡‘…”‘… ≈” ≈√ ¡‘¡”‘¡”≈…” »≈—Ã¡Õ”«”  ¡… ÿ’Œ«”"
    ',"FIRMACTKIND":"2","FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"43220000","FIRMACTDESCR":"◊œÕƒ—… œ ≈Ã–œ—…œ Ã≈À
    '¡Õ«” √—¡÷«” « ”◊≈ƒ…¡”«”  ¡… ¡ÀÀŸÕ Ã≈À¡Õ…ŸÕ","FIRMACTKIND":"2","FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":
    '"46492315","FIRMACTDESCR":"◊œÕƒ—… œ ≈Ã–œ—…œ Ã≈À¡Õ…ŸÕ, Ã≈ ¡Õ¡√œÃŸ”« ÷’”…√√…ŸÕ (CARTRIDGES) ≈ ‘’–Ÿ‘ŸÕ",
    '"FIRMACTKIND":"2","FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"46492316","FIRMACTDESCR":
    '"◊œÕƒ—… œ ≈Ã–œ—…œ ≈–…–ÀŸÕ √—¡÷≈…œ’","FIRMACTKIND":"2","FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":
    '"46650000","FIRMACTDESCR":"◊œÕƒ—… œ ≈Ã–œ—…œ Ã«◊¡Õ«Ã¡‘ŸÕ Ã≈‘¡‘—œ–«” «À…¡ «” ≈Õ≈—√≈…¡” ”≈ «À≈ ‘—… «",
    '"FIRMACTKIND":"2","FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"46691551","FIRMACTDESCR":
    '"À…¡Õ… œ ≈Ã–œ—…œ Ã≈—ŸÕ  ¡… ≈Œ¡—‘«Ã¡‘ŸÕ ’–œÀœ√…”‘… ŸÕ Ã«◊¡ÕŸÕ","FIRMACTKIND":"2","FIRMACTKINDDESCR":
    '"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"47413106","FIRMACTDESCR":"À…¡Õ… œ ≈Ã–œ—…œ ÷Ÿ‘œ¬œÀ‘¡⁄ ŸÕ ”‘œ…◊≈…ŸÕ",
    '"FIRMACTKIND":"2","FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"47545441","FIRMACTDESCR":
    '"À…¡Õ… œ ≈Ã–œ—…œ ≈–…–ÀŸÕ","FIRMACTKIND":"2","FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"47595500",
    '"FIRMACTDESCR":"’–«—≈”…≈” ≈√ ¡‘¡”‘¡”
    '«” «À≈ ‘—œÕ… ŸÕ ’–œÀœ√…”‘ŸÕ  ¡… –≈—…÷≈—≈…¡ œ’ ≈Œœ–À…”Ãœ’","FIRMACTKIND":"2",
    '"FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"62091000","FIRMACTDESCR":
    '"’–«—≈”…≈” ≈Õœ… …¡”«”  ¡… ◊—«Ã¡‘œƒœ‘… «” Ã…”»Ÿ”«” «À≈ ‘—œÕ… ŸÕ ’–œÀœ√…”‘ŸÕ","FIRMACTKIND":"2",
    '"FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"77331200","FIRMACTDESCR":
    '"’–«—≈”…≈” ≈–…” ≈’«” «À≈ ‘—œÕ… ŸÕ ’–œÀœ√…”‘ŸÕ  ¡… –≈—…÷≈—≈…¡ œ’ ≈Œœ–À…”Ãœ’",
    '"FIRMACTKIND":"2","FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"95111000",
    '"FIRMACTDESCR":"’–«—≈”…≈” ≈√ ¡‘¡”‘¡”«” «À≈ ‘—… œ’ ≈Œœ–À…”Ãœ’","FIRMACTKIND":"2",
    '"FIRMACTKINDDESCR":"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"33205000","FIRMACTDESCR":
    '"’–«—≈”…≈” ≈√ ¡‘¡”‘¡”«” ¡ÀÀŸÕ ≈…ƒŸÕ –.ƒ. .¡.","FIRMACTKIND":"2","FIRMACTKINDDESCR":
    '"ƒ≈’‘≈—≈’œ’”¡","FIRMACTCODE":"33207000"],"PCALLSEQID_OUT":"1808311808","PERRORREC_OUT":
    '"ERRORDESCR":NULL,"ERRORCODE":NULL
    'VIES RESULTS:
    '"COUNTRYCODE":"EL","VATNUMBER":"028783755","REQUESTDATE":"2021-04-23+02:00","VALID":TRUE,
    '"NAME":"À¡√¡ «” √≈Ÿ—√…œ” Õ… œÀ¡œ”","ADDRESS":"  –¡À¡…œÀœ√œ’ 21        66100 - ƒ—¡Ã¡"


cmdCommand3_err:
128    ' MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.CmdCommand3_err_Err " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
130     SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.CmdCommand3_err_Err " & "at line " & Erl

132     Resume Next



End Sub

Public Sub gemiseFields(ISNEW As Integer, aaa As String, FF As Form)


     Dim a() As String
     Dim D() As String
100  a = Split(aaa, ",")
 
     Dim c As String
102  c = a(1)
    ' Dim isnew As Integer
     'isnew = 1
 
 
    '
    '
    '


    Dim doy, EPO, EPA, ENERGOS, DIE, dien
    Dim sos As Integer
    Dim tk, pol As String

    Dim k As Integer
    Dim DD As String
104 For k = 1 To UBound(a)

    '
    '
106     DD = a(k)
108     D = Split(DD, ":")
    
110     If UBound(D) > 0 Then
112        If InStr(1, D(0), "DOYDESCR") > 0 Then
114          doy = D(1)
           End If

116         If InStr(1, D(0), "ONOMASIA") > 0 Then
118            EPO = D(1)
            End If

    '
    '       'LEGALSTATUSDESCR
    '
120         If InStr(1, D(0), "ARRAYOFRGWSPUBLICFIRMACTRT_OUT") > 0 Then
122            If UBound(D) > 2 Then
124               EPA = D(3)
               Else
126               EPA = ""
               End If
              
            End If

    '       On Error GoTo LATOS
    '
    ' '"DEACTIVATIONFLAGDESCR":"≈Õ≈—√œ” ¡÷Ã "
    '
128         If InStr(1, D(0), "DEACTIVATIONFLAGDESCR") > 0 Then
130            If Left(D(1), 7) = Left("""≈Õ≈—√œ” ¡÷Ã """, 7) Then
132                sos = 0
                  ' MsgBox "ok ¡÷Ã"
               Else
134                sos = 1
136                MsgBox "ÒÔÛÔ˜Á !!!!!!! ÏÁ ÂÌÂÒ„Ô ¡÷Ã"
               End If

138            ENERGOS = D(1)
            End If
    '
           'POSTALADDRESS
140         If InStr(1, D(0), "POSTALADDRESS") > 0 Then
142            DIE = D(1) + dien
            End If

           'POSTALADDRESSNO
144         If InStr(1, D(0), "POSTALADDRESSNO") > 0 Then
146            dien = " " + D(1)
            End If

           ' "POSTALZIPCODE":"66100"
148         If InStr(1, D(0), "POSTALZIPCODE") > 0 Then
150            tk = D(1)
            End If




          'POSTALAREADESCRIPTION
152        If InStr(1, D(0), "POSTALAREADESCRIPTION") > 0 Then
154            pol = D(1)
            End If
    '
    '
        End If
    '
    '
    ' ' Line Input #1, gConnect
    Next
    '



156 If sos = 1 Then
158     pelat1.Text1(1) = "ÏÁ Â„ÍıÒÔ ¡÷Ã"
     Else
160     If ISNEW = 1 Then
162        FF.Text1(1) = Replace(EPO, """", "")
164        FF.Text1(2) = Replace(DIE, """", "")
166        FF.Text1(3) = Replace(pol, """", "")
168        FF.Text1(4) = Replace(EPA, """", "")
170        FF.Text1(7) = Replace(doy, """", "")
172        FF.Text1(7) = Replace(FF.Text1(7), "'", "`")
174        FF.Text1(12) = Replace(tk, """", "")
        Else
176        FF.Text1(1) = Replace(EPO, """", "")
178        FF.Text1(2) = Replace(DIE, """", "")
180        FF.Text1(3) = Replace(pol, """", "")
182        FF.Text1(4) = Replace(EPA, """", "")
184        FF.Text1(7) = Replace(doy, """", "")
186        FF.Text1(7) = Replace(FF.Text1(7), "'", "`")
188        FF.Text1(13) = Replace(tk, """", "")
        End If
    
     End If
 



     'Form1.GEMISE_TEXT 1, Me

End Sub

'Private Sub cmdEYRESHAfm_Click()
'On Error Resume Next
'
'   If Len(Trim(Text1(5).Text)) = 9 Then
''OK       Form1.parametros.Caption = "http://lagakis.gr/greek-vat/web-vat-details.php?myafm=028783755&afm=" + Trim(Text1(5).Text) ' 028783755"
'      Dim u As String, p As String, AFM As String
'  u = Trim(f_EYRESHAFM_USER)   ' = "LAGAKIS23041960"  ' Ww367868u916
'  p = Trim(f_EYRESHAFM_PWD)  '  = "d389833921"   '080353
' AFM = Trim(f_EYRESHAFM_AFM)    ' = "028783755"   '051103283
'
'
'
'   '   u = "Ww367868u916"
'    '  p = "080353"
'    '  AFM = "051103283"
'
'
'
'
'
'       Form1.parametros.Caption = "https://lagakis.gr/greek-vat/web-vat-details.php?u=" + u + "&p=" + p + "&myafm=" + AFM + "&afm=" + Trim(Text1(5).Text) ' 028783755"
'
'
'      ' Form1.SHOW
'       MILSEC 2000
'       Form1.Command4_Click
'       MILSEC 2000
'       Form1.Command1_Click
'       Form1.GEMISE_TEXT 1, Me
'
'   End If
'   'Form1.SHOW
'End Sub


Public Function UnicodeTOGreek(s As String) As String


Dim JS, MASCI



JS = "\u0391\u0392\u0393\u0394\u0395\u0396\u0397\u0398\u0399\u039A\u039B\u039C\u039D\u039E\u039F\u03A0\u03A1\u03A3\u03A4\u03A5\u03A6\u03A7\u03A8\u03A9\u0386\u0388\u0389\u038A\u038E\u038F\u038C\u03B1\u03B2\u03B3\u03B4\u03B5\u03B6\u03B7\u03B8\u03B9\u03BA\u03BB\u03BC\u03BD\u03BE\u03BF\u03C0\u03C1\u03C3\u03C4\u03C5\u03C6\u03C7\u03C8\u03C9\u03C2\u03AC\u03AD\u03AE\u03AF\u03CC\u03CE\u03CD\u03CB\u03B0\U03AA\U0384"
MASCI = "¡¬√ƒ≈∆«»… ÀÃÕŒœ–—”‘’÷◊ÿŸ¢∏π∫æøº·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÛÙıˆ˜¯˘Ú‹›ﬁﬂ¸˛˝˚‡⁄'"

'Dim DD(300)
'FETES_DELIM JS, DD
JS = UCase(JS)

Dim output() As String
output = Split(JS, "\")

'65 GRAMMATA

Dim L As Integer

Dim C1, c2

s = UCase(s)
For L = 1 To 67
   C1 = "\" + output(L)
   c2 = mID(MASCI, L, 1)
   s = Replace(s, C1, c2)
 

Next


s = Replace(s, "{", "")
s = Replace(s, "}", "")

Dim OU() As String
OU = Split(s, ",")


Open "C:\MERCVB\AA2.TXT" For Output As #1
   

For L = 0 To UBound(OU) - 1
   Print #1, OU(L) + Chr(13)
Next
Close #1

UnicodeTOGreek = s

End Function










Private Sub cmdKatax_Click()
Dim vid As String
'Par7MyData.loadpar7
'vid = Par7MyData.newclient(3, ARKYKL.Text, VFACT.Text, VCAT.Text)


lblvID.Caption = vid

'vid = Par7MyData.newclient(3, ARKYKL.Text, VFACT.Text, VCAT.Text)




End Sub

'Private Sub Command1_Click()
'   Dim r As Long
'   r = ShellExecute(0, "open", "http://www.microsoft.com", 0, 0, 1)
'End Sub


Private Sub cmd–ÂÒÈÛÛ¸ÙÂÒ·_Click()
 Dim R As Long
   On Error Resume Next
   R = ShellExecute(0, "open", "https://el.wikipedia.org/wiki/ISO_3166-1", 0, 0, 1)
' WebBrowser1.Navigate2 "https://el.wikipedia.org/wiki/ISO_3166-1"
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
         "in ADOMERCNEW.pelat1.Combo1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Combo1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Combo2_KeyPress_Err

        '</EhHeader>

100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
            ' keybd_event VK_TAB, 0, 0, 0 'send a tab
120         Command2.SetFocus
        End If

        '<EhFooter>
        Exit Sub

Combo2_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.Combo2_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Combo2_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "~"
        End If

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim k As Integer, DUM As Long, DB As Database, R As New ADODB.Recordset

        Dim x As String, ANS As Integer

        '      If Option1(1) Then X = "e" Else X = "r"
100     x = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)

110     If Command2.Caption = "Õ›· ≈„„Ò·ˆﬁ" Then
            If kenoi.Value = vbChecked Then
                DUM = get_next_pel(x, 1, F_FORMAT_PEL_KOD)
            Else
                 DUM = get_next_pel(x, 0, F_FORMAT_PEL_KOD)
            End If
120
            
            If f_KENTRA_ADYNATISMATOS = 1 Then
                If Combo3.ListIndex = 1 Then
                    DUM = 100000 + DUM
                    Text1(0).Text = Format(DUM, "000000")
                Else
                    Text1(0).Text = Format(DUM, "0000")
                End If
               
            Else
                Text1(0).Text = Format(DUM, F_FORMAT_PEL_KOD)
            End If
            
130
140         Command2.Caption = "≈ÌÁÏ›Ò˘ÛÁ"

150         For k = 0 To 12
160             Text1(k).Enabled = True
170             Text1(k).ForeColor = vbBlack
            Next

180         For k = 1 To 12
190             Text1(k).Text = ""
            Next

200         Text1(0).SetFocus
        Else
        
        
        
        
        
        '·Ì ÛÙÁÌ ÙÈÏÔÎ¸„ÁÛÁ ˆÈÎÙÒ·ÒÂÈ ÙÔıÚ ÂÎ‹ÙÂÚ, Ù¸ÙÂ Ò›ÂÈ Ì· ‚‹ÎÂÈÚ Í·ÙÁ„ÔÒﬂ·
                If f_MONO_PELATES_KATHGORIAS = 1 And x = "e" Then
                    If Val(Left(KATHGORIA.Text, 2)) = 0 Then
                         MsgBox "–Ò›ÂÈ Ì· ‚‹ÎÂÙÂ Í·ÙÁ„ÔÒﬂ· ◊ÔÌ‰ÒÈÍﬁÚ/ÀÈ·ÌÈÍﬁÚ"
                         Exit Sub
                    End If
                End If
                
        
        
        
        
        
        
        
        
        
        
        
        
            'ÂÎÂ„˜ÔÚ ÏÁÌ Ó·Ì·ı‹Ò˜ÂÈ Ô Í˘‰ÈÍ¸Ú
            'Set db = OpenDatabase(gDir, False, False, gConnect)
210         R.Open "select count(*) as ar from PEL where EIDOS='" + x + "' and KOD='" + Text1(0).Text + "';", Gdb, adOpenDynamic, adLockOptimistic

220         If R("ar") > 0 Then
230             MsgBox "’‹Ò˜ÂÈ Ô Í˘‰ÈÍ¸Ú"

                'Command2.Caption = "Õ›· ≈„„Ò·ˆﬁ"
                'Command2.SetFocus
                Exit Sub

            End If

240         R.Close
250         R.Open "select count(*) as ar,KOD,EPO from PEL where EIDOS='" + x + "' and AFM='" + Text1(5).Text + "' group by KOD,EPO;", Gdb, adOpenDynamic, adLockOptimistic

260         If Not R.EOF And Len(Text1(5).Text) > 3 Then
270             If R("ar") > 0 Then
280                 ANS = MsgBox("’‹Ò˜ÂÈ Ó·Ì‹ ÙÔ ¡÷Ã : " + R(1) + " - " + R(2) + ". Õ· „ﬂÌÂÈ Á Í·Ù·˜˛ÒÁÛÁ;", vbYesNo)

290                 If ANS = vbNo Then

                        Exit Sub

                    End If

                End If
            End If

            ' ÂÌÁÏ›Ò˘ÛÁ ÏÂÙÒÁÙﬁ
            If f_KENTRA_ADYNATISMATOS = 1 Then
                DUM = Update_Next_Pel(x, Right(Text1(0), 4))
            Else
                DUM = Update_Next_Pel(x, Text1(0))
            End If
            
300

            On Error Resume Next

            On Error GoTo Command2_Click_Err
  Dim MKOD As String
  
  MKOD = Text1(0).Text
310        ' data1.Recordset.AddNew
             
            Gdb.Execute "insert into PEL (EIDOS,KOD) VALUES ('" + x + "','" + MKOD + "')"

320         'data1.Recordset("EIDOS") = X
330         'data1.Recordset("kod") = Text1(0).Text   ' Format(Text1(0), "0000")
340       '  data1.Recordset("epo") = Left(Text1(1).Text, data1.Recordset("epo").DefinedSize)
    
    Dim MDIE As String
  
    
    
    
  
350         If Len(Text1(2).Text) = 0 Then
360             MDIE = "-"
            Else
370             MDIE = Left(Text1(2).Text, data1.Recordset("DIE").DefinedSize)
            End If
            
            
            
            
        If f_anal_epon = 1 Then
            Text1(1).Text = Left(EPO1 + Space(14), 14) + " " + Left(MIKRO.Text + Space(14), 10) + " " + Left(PATR.Text + Space(14), 9)
        End If
            
            
            Gdb.Execute "UPDATE PEL SET EPO='" + Left(Text1(1).Text, data1.Recordset("epo").DefinedSize) + "' WHERE EIDOS='" + x + "' AND KOD='" + MKOD + "'"

 
 UPDATE_PEL "DIE", Left(Text1(2).Text, data1.Recordset("DIE").DefinedSize), 0, x, MKOD
 UPDATE_PEL "POL", Left(Text1(3).Text, data1.Recordset("POL").DefinedSize), 0, x, MKOD
 UPDATE_PEL "EPA", Left(Text1(4).Text, data1.Recordset("EPA").DefinedSize), 0, x, MKOD
 UPDATE_PEL "AFM", Left(Text1(5).Text, data1.Recordset("AFM").DefinedSize), 0, x, MKOD
 UPDATE_PEL "THL", Left(Text1(6).Text, data1.Recordset("THL").DefinedSize), 0, x, MKOD
 
 UPDATE_PEL "CH1", Left(C1.Text, data1.Recordset("CH1").DefinedSize), 0, x, MKOD
 
 
If F_ISCAR = 1 Then
   
   
   UPDATE_PEL "ARKYKL", Left(ARKYKL.Text, data1.Recordset("ARKYKL").DefinedSize), 0, x, MKOD
   UPDATE_PEL "VCAT", Left(VCAT.Text, data1.Recordset("VCAT").DefinedSize), 0, x, MKOD
   UPDATE_PEL "VFACT", Left(VFACT.Text, data1.Recordset("VFACT").DefinedSize), 0, x, MKOD
   UPDATE_PEL "VID", Left(lblvID.Caption, data1.Recordset("VID").DefinedSize), 0, x, MKOD
   
   
 End If
 
 
 
 
 
 
 
 
 
 
 
 
 

'380         data1.Recordset("pol") = Left(Text1(3).Text, data1.Recordset("POL").DefinedSize)  'Text1(3).Text
'390         data1.Recordset("epa") = Left(Text1(4).Text, data1.Recordset("epA").DefinedSize)  ' Text1(4).Text
'400         data1.Recordset("afm") = Left(Text1(5).Text, data1.Recordset("AFM").DefinedSize)  ' Text1(5).Text
'410         data1.Recordset("thl") = Left(Text1(6).Text, data1.Recordset("THL").DefinedSize)  ' Text1(6).Text


UPDATE_PEL "COUNTRY", Left(COUNTRY.Text, 2), 0, x, MKOD

UPDATE_PEL "DOY", Left(Text1(7).Text, data1.Recordset("DOY").DefinedSize), 0, x, MKOD

UPDATE_PEL "ENERGOS", "1", 1, x, MKOD
UPDATE_PEL "PEK", Left(Text1(8).Text, data1.Recordset("PEK").DefinedSize), 1, x, MKOD

UPDATE_PEL "PLAFON", Left(Text1(9).Text, data1.Recordset("PLAFON").DefinedSize), 1, x, MKOD

'420         data1.Recordset("doy") = Left(Text1(7).Text + " ", data1.Recordset("DOY").DefinedSize)    ' Text1(7).Text
'430         data1.Recordset("pek") = Val(Text1(8).Text)
'440         data1.Recordset("plafon") = Val(Text1(9).Text)



UPDATE_PEL "SHM2", Left(Text1(10).Text, data1.Recordset("SHM2").DefinedSize), 0, x, MKOD
'450         data1.Recordset("shm2") = Left(Text1(10).Text + " ", data1.Recordset("SHM2").DefinedSize)    ' Text1(10).Text



UPDATE_PEL "AYP", Left(Text1(11).Text, data1.Recordset("AYP").DefinedSize), 1, x, MKOD
460         data1.Recordset("ayp") = Val(Text1(11).Text)



UPDATE_PEL "XRVMA", Left(Text1(12).Text, data1.Recordset("XRVMA").DefinedSize), 0, x, MKOD
'470         data1.Recordset("xrvma") = Val(Text1(12).Text)




UPDATE_PEL "BRANCH", BRANCH.Text, 0, x, MKOD
UPDATE_PEL "STREETNUMBER", STREETNUMBER.Text, 0, x, MKOD



'UPDATE_PEL "TYP", Left(Text1(11).Text, data1.Recordset("TYP").DefinedSize), 1, X, MKOD
'480         data1.Recordset("typ") = Val(Text1(11).Text)


UPDATE_PEL "KODGAL", Left(KATHGORIA.Text, 2), 1, x, MKOD
'490         data1.Recordset("KODGAL") = Left(KATHGORIA.Text, 2)    'KATHGORIA


UPDATE_PEL "PLAISIO", Left(Text1(13).Text, data1.Recordset("PLAISIO").DefinedSize), 0, x, MKOD

'500         data1.Recordset("PLAISIO") = Text1(13).Text    'METAFORIKH

UPDATE_PEL "TYPOS", Left(Combo1.Text, data1.Recordset("TYPOS").DefinedSize), 0, x, MKOD

'510         data1.Recordset("TYPOS") = Combo1.Text


UPDATE_PEL "ARPARAG", Left(Combo2.Text, data1.Recordset("ARPARAG").DefinedSize), 0, x, MKOD
'520         data1.Recordset("ARPARAG") = Combo2.Text    'PLHROMI



UPDATE_PEL "ADT", Left(Text1(14).Text, data1.Recordset("ADT").DefinedSize), 0, x, MKOD
'UPDATE_PEL "CH1", Left(Text1(14).Text, DATA1.Recordset("CH1").DefinedSize), 0, X, mkod
'530         data1.Recordset("CH1") = Text1(14).Text    'DEH


           


            'f_KENTRA_ADYNATISMATOS = Val(FindParametroi(1,"PELAT2", "f_KENTRA_ADYNATISMATOS", "0", "¡Ì ÂﬂÌ·È Í›ÌÙÒÔ ·‰ıÌ·ÙﬂÛÏ·ÙÔÚ=1"))
            If f_KENTRA_ADYNATISMATOS = 1 Then
                data1.Recordset("num3") = Combo3.ListIndex
                UPDATE_PEL "NUM3", str(Combo3.ListIndex), 1, x, MKOD
            End If

            '  Data1.Recordset("ayp") = 0
540        ' data1.Recordset.Update

550         MHNYMA2.Timer1.Interval = 2000
560         MHNYMA2.Label2.Caption = "« Â„„Ò·ˆﬁ " + Text1(1).Text + " ·ÔËÁÍÂ˝ÙÁÍÂ ÏÂ ÂÈÙı˜ﬂ·"
570         MHNYMA2.SHOW 1

580         Command2.Caption = "Õ›· ≈„„Ò·ˆﬁ"
        End If

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub UPDATE_PEL(ByVal PEDIO As String, ByVal TIMH As String, ByVal ARITMOS As Integer, ByVal EIDOS As String, ByVal kod As String)
   
   If ARITMOS = 1 Then
            Gdb.Execute "UPDATE PEL SET " + PEDIO + "=" + TIMH + "   WHERE EIDOS='" + EIDOS + "' AND KOD='" + kod + "'"
   Else
            Gdb.Execute "UPDATE PEL SET " + PEDIO + "='" + TIMH + "' WHERE EIDOS='" + EIDOS + "' AND KOD='" + kod + "'"

   End If



End Sub






Private Sub Form_Activate()

        '<EhHeader>
        On Error GoTo Form_Activate_Err

        '</EhHeader>

100     Me.Picture = LoadPicture(gPicture)

110     data1.ConnectionString = gConnect
120     data1.RecordSource = "select top 10 *from PEL"
130     data1.Refresh

        '<EhFooter>
        Exit Sub

Form_Activate_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.Form_Activate " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Form_Activate " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

100     PARAMETROI.PARAM.Caption = "PELAT1"
110     PARAMETROI.SHOW 1


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command1_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '   Text1(0).Text = get_next_pel("e")
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl
110     Me.KeyPreview = True

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        ' PELPROM.AddItem "–≈À¡‘«”"
        ' PELPROM.AddItem "–—œÃ«»≈’‘«”"
120     PELPROM_STRING = "er"
        ' PELPROM.Text = PELPROM.List(0)
130     PELPROM.ListIndex = 0

        'tropos plhromhs
        Dim R As New ADODB.Recordset

140     Adodc2.ConnectionString = gConnect

150     Adodc2.RecordSource = "SELECT STR(AYJON,2,0) AS AYJ,PERIGRAFH FROM PINAKES WHERE TYPOS=13 ORDER BY PERIGRAFH"
160     Adodc2.Refresh

170     R.Open "SELECT *FROM PINAKES WHERE TYPOS=12 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

180     Do While Not R.EOF

190         If R("typos") = 12 Then
200             Combo2.AddItem str(R("AYJON")) + "." + R("PERIGRAFH")
            End If

210         R.MoveNext
        Loop

       
       
     
   f_lab1 = FINDPARAMETROI(3, "PELAT2", "F_lab1", " ÂÈÏ1", "≈ÙÈÍ›Ù· ÍÂÈÏ.1")
'Label1(15).Caption = f_lab1

F_artam = Val(FINDPARAMETROI(1, "PELAT1", "F_ARTAM", "", "¡—…»Ãœ” SN TAMEIAKH” /÷œ—OÀO√IKOY ”’Õ≈—√≈…œ’")) '

f_EYRESHAFM_USER = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_USER", "LAGAKIS23041960", "◊—«”‘«” ≈’—≈”«” ¡÷Ã")
f_EYRESHAFM_PWD = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_PWD", "d389833921", " Ÿƒ… œ” ≈’—≈”«” ¡÷Ã")
f_EYRESHAFM_AFM = FINDPARAMETROI(1, "PELAT1", "f_EYRESHAFM_AFM", "028783755", "¡÷Ã ≈–…◊≈…—«”«")
       
       
       ' F_ISCAR
        F_ISCAR = Val(FINDPARAMETROI(1, "PELAT1", "F_ISCAR", "0", "”YNER√EIA=1 œ◊…=0")) '
        
        
        
        
       If F_ISCAR = 1 Then
         ARKYKL.Visible = True
         VCAT.Visible = True
         VFACT.Visible = True
         cmdKatax.Visible = True
         LARKYKL.Visible = True
         lblCAT.Visible = True
         Label3.Visible = True
       
       End If
        
        
        
        
        
        
        
        
        f_MONO_PELATES_KATHGORIAS = Val(FINDPARAMETROI(1, "PAR1", "f_MONO_PELATES_KATHGORIAS", "0", "√.45.TÔ ·Ò/ÙÈÍ¸ ˛ÎÁÛÁÚ Ï¸ÌÔ ·¸ Í·ÙÁ„.ÂÎ·Ù˛Ì=1 œ◊…=0")) '·„ÔÒ=1 ˘Î=2 ·ÌÙÔı=3 œ◊…=0"))
        
        F_PALIO_FINDAFM = Val(FINDPARAMETROI(1, "PELAT1", "F_PALIO_FINDAFM", "0", "¡Ì·ÊﬁÙÁÛÁ ÏÂ ÙÔÌ ·Î·È¸ ÙÒ¸Ô=1 ÃÂ Ì›Ô=0 ")) '
        
       
        f_anal_epon = Val(FINDPARAMETROI(1, "PELAT1", "F_ANAL_EPON", "0", "1=≈–ŸÕ’Ãœ,œÕœÃ¡ –¡‘—ŸÕ’Ãœ 0=≈–ŸÕ’Ã…¡"))
       
        F_FORMAT_PEL_KOD = FINDPARAMETROI(1, "PELAT1", "F_FORMAT_PEL_KOD", "0000", "ÿ«÷…¡  Ÿƒ… œ’ –≈À¡‘ŸÕ")

        f_KENTRA_ADYNATISMATOS = Val(FINDPARAMETROI(1, "PELAT2", "f_KENTRA_ADYNATISMATOS", "0", "¡Ì ÂﬂÌ·È Í›ÌÙÒÔ ·‰ıÌ·ÙﬂÛÏ·ÙÔÚ=1"))

        If f_KENTRA_ADYNATISMATOS = 1 Then
            Combo3.Visible = True
            Combo3.ListIndex = 0 ' Combo3.List(0)
        End If

        'KATHGORIA
220     R.Close
230     R.Open "SELECT *FROM PINAKES WHERE TYPOS=13 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

240     If R.EOF Then
250         Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH) VALUES (13,1,'≈Ã–œ—œ”')"
        End If

260     R.MoveFirst

270     Do While Not R.EOF

280         If R("typos") = 13 Then
290             KATHGORIA.AddItem str(R("AYJON")) + "." + R("PERIGRAFH")
            End If

300         R.MoveNext
        Loop

310     R.Close

        ' R.Close


      If f_anal_epon = 1 Then
          EPO1.Visible = True
          MIKRO.Visible = True
          PATR.Visible = True
          
          LEpo1.Visible = True
          lmikro.Visible = True
          Lpatr.Visible = True
          
          Text1(1).Visible = False
          
          
      End If







320     f_TIMOKAT = Val(FINDPARAMETROI(1, "PELAT2", "F_TIMOKAT", "0", "AÒÈËÏ¸Ú ‘ÈÏÔÍ·Ù·Î¸„˘Ì"))

330     If f_TIMOKAT = 1 Then
340       '  Label1(11).Caption = "‘ÈÏÔÍ·Ù‹ÎÔ„ÔÚ"
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
'ResizeFormFor Me
    mForm_Resize Me, 8, 10860, 9060, fh, fw, ft, fl

End Sub

Private Sub KATHGORIA_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo KATHGORIA_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

KATHGORIA_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.KATHGORIA_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.KATHGORIA_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub PELPROM_Change()

    '<EhHeader>
    On Error GoTo PELPROM_Change_Err

    '</EhHeader>
    On Error Resume Next

    'Text1(0).Text = get_next_pel(mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1))
    '<EhFooter>
    Exit Sub

PELPROM_Change_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.pelat1.PELPROM_Change " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.PELPROM_Change " & "at line " & Erl

    Resume Next

    '</EhFooter>

End Sub

'Private Sub Option1_Click(Index As Integer)
' ' If Option1(1) Then
'     Text1(0).Text = get_next_pel(mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1))
' ' Else
' '    Text1(0).Text = get_next_pel("r")
' ' End If
'
'End Sub

Private Sub Text1_GotFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text1_GotFocus_Err

        '</EhHeader>

100     Text1(index).BackColor = vbYellow

        '<EhFooter>
        Exit Sub

Text1_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.Text1_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Text1_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)

        '  If KeyAscii = 13 Then KeyAscii = 0
        '<EhHeader>
        On Error GoTo Text1_KeyPress_Err

        '</EhHeader>

100     If KeyAscii = vbKeyReturn Then
            ' If Index = 11 Then
            '    Command2.SetFocus
            ' Else
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
            '  End If
        End If

        '<EhFooter>
        Exit Sub

Text1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat1.Text1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Text1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    ' If KeyCode = 13 Then 'enter
    '   If Index < 7 Then
    '
    '      Text1(Index + 1).SetFocus
    '   Else
    '      Command2.SetFocus
    '   End If
    'End If
    '
    '
    '

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
         "in ADOMERCNEW.pelat1.Text1_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat1.Text1_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
