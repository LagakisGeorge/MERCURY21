VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form bohu5 
   BackColor       =   &H00FF0000&
   Caption         =   "Imports"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H0000C000&
      Caption         =   "Frame4"
      Height          =   2175
      Left            =   12000
      TabIndex        =   44
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command15 
         Caption         =   "ÂÈÛÒÔ˛Ì"
         Height          =   360
         Left            =   240
         TabIndex        =   50
         Top             =   1800
         Width           =   990
      End
      Begin VB.CommandButton Command8 
         Caption         =   "…mport –ÂÎ·Ù˛Ì/–ÒÔÏÁËÂıÙ˛Ì  ·¸ Excel"
         Height          =   495
         Left            =   240
         TabIndex        =   46
         Top             =   1080
         Width           =   2550
      End
      Begin VB.CommandButton Command7 
         Caption         =   "…mport ≈È‰˛Ì ·¸ Excel"
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   480
         Width           =   2550
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Caption         =   "Frame3"
      Height          =   2175
      Left            =   4560
      TabIndex        =   35
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton IMPORT_MHLIOY 
         Caption         =   "IMPORT ƒ≈À‘…ŸÕ"
         Height          =   372
         Left            =   720
         TabIndex        =   53
         Top             =   1560
         Width           =   2172
      End
      Begin VB.CommandButton karam2 
         BackColor       =   &H00FFFF00&
         Caption         =   "¡–œ 2o ÷œ—‘«√œ  ¡—¡Ã"
         Height          =   360
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton Command14 
         Caption         =   "¡–œ 1o ÷œ—‘«√œ  ¡—¡Ã"
         Height          =   360
         Left            =   5160
         TabIndex        =   48
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton KARAM 
         Caption         =   " ¡—¡Ã≈”…Õ«” ÍÒÈ-ÍÒÈ ·ÔÛÙÔÎﬁ ·Ò˜Âﬂ˘Ì"
         Height          =   375
         Left            =   5160
         TabIndex        =   47
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "≈·ÌıÔÎÔ„ÈÛÏ¸Ú –·Ò·ÛÙ·ÙÈÍ˛Ì"
         Height          =   375
         Left            =   3180
         TabIndex        =   43
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton EXP_SMARTWARE 
         Caption         =   "ÍÒÈ-ÍÒÈ ·ÔÛÙÔÎﬁ ·Ò˜Âﬂ˘Ì"
         Height          =   375
         Left            =   720
         TabIndex        =   39
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command11 
         Caption         =   "”ı„ÍÂÌÙÒ˘ÙÈÍ¸ ÂÈÛ"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3180
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command12 
         Caption         =   "”ı„Í.÷¸ÒÙ˘ÛÁÚ"
         Height          =   255
         Left            =   3180
         TabIndex        =   37
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         Caption         =   "·Ò·Î·‚Á ÙÈÏÔÎÔ„È˘Ì"
         Height          =   495
         Left            =   3180
         TabIndex        =   36
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "√Ò·ÏÏÔ„Ò‹ˆÁÛÁ ÂÓ·„.–ÂÎ·Ù˘Ì"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   42
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "√Ò·ÏÏÔ„Ò‹ˆÁÛÁ ÂÓ·„.≈È‰˛Ì"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   41
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "√Ò·ÏÏÔ„Ò‹ˆÁÛÁ ‘ÈÏÔÍ.≈È‰˛Ì"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   40
         Top             =   1200
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Caption         =   "Frame2"
      Height          =   8775
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   4455
      Begin VB.TextBox ARPAR 
         Height          =   285
         Left            =   480
         TabIndex        =   52
         Text            =   "10"
         Top             =   7800
         Width           =   735
      End
      Begin VB.CommandButton IMPORTSERV 
         Caption         =   "IMPORT ¡–œ SERVER"
         Height          =   360
         Left            =   360
         TabIndex        =   51
         Top             =   7320
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Import"
         Height          =   465
         Left            =   360
         TabIndex        =   32
         Top             =   600
         Width           =   3840
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   2115
         TabIndex        =   31
         Text            =   "C:\MERCVB\F35.TXT"
         Top             =   1245
         Width           =   2085
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   2130
         TabIndex        =   30
         Text            =   "z:\kef4\main.asc"
         Top             =   1785
         Width           =   2085
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "bohu5b.frx":0000
         Left            =   360
         List            =   "bohu5b.frx":000A
         TabIndex        =   29
         Top             =   2520
         Width           =   3840
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Import ‘ÈÏÔÎÔ„ﬂ˘Ì"
         Height          =   465
         Left            =   360
         TabIndex        =   28
         Top             =   3960
         Width           =   3840
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Import ¡–œ EXCEL ""C:\MODELO2.xls / ÷˝ÎÎÔ1"""
         Height          =   465
         Left            =   360
         TabIndex        =   27
         Top             =   3120
         Width           =   3840
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   4560
         Width           =   3840
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Left            =   360
         TabIndex        =   25
         Top             =   5520
         Width           =   3840
      End
      Begin VB.CommandButton Command9 
         Caption         =   "À«ÿ« ¡–œ PDA"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   3600
         Width           =   3840
      End
      Begin VB.CommandButton cmd≈–…”‘—œ÷«÷œ—‘«√œ’ 
         BackColor       =   &H0000C0C0&
         Caption         =   "≈–…”‘—œ÷« ÷œ—‘«√œ’ ANDROID"
         Height          =   360
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5040
         Width           =   3840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "÷œ—Ã¡"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   360
         TabIndex        =   34
         Top             =   1275
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "¡—◊≈…œ"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   360
         TabIndex        =   33
         Top             =   1815
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   6495
      Left            =   4560
      TabIndex        =   1
      Top             =   2280
      Width           =   10335
      Begin VB.CommandButton Command4 
         Caption         =   " ¡‘¡◊Ÿ—«”« ÷œ—‘Ÿ‘… «”"
         Height          =   405
         Left            =   2565
         TabIndex        =   13
         Top             =   3165
         Width           =   2910
      End
      Begin VB.TextBox POSO 
         Height          =   330
         Left            =   2595
         TabIndex        =   11
         Top             =   1455
         Width           =   1725
      End
      Begin VB.ComboBox AFM 
         Height          =   315
         ItemData        =   "bohu5b.frx":0018
         Left            =   2595
         List            =   "bohu5b.frx":001A
         TabIndex        =   10
         Top             =   1005
         Width           =   1710
      End
      Begin VB.TextBox MATIM 
         Height          =   345
         Left            =   2595
         TabIndex        =   9
         Top             =   1890
         Width           =   1710
      End
      Begin VB.CommandButton Command5 
         Caption         =   " ·Ù·˜˛ÒÁÛÁ ∆ Ù·ÏÂÈ·ÍﬁÚ"
         Height          =   375
         Left            =   2595
         TabIndex        =   8
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox FPA 
         Height          =   375
         Left            =   7005
         TabIndex        =   7
         Top             =   3840
         Width           =   855
      End
      Begin VB.ListBox Z 
         Height          =   840
         Left            =   1080
         TabIndex        =   6
         Top             =   4680
         Width           =   6975
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   6285
         TabIndex        =   5
         Top             =   1455
         Width           =   1590
      End
      Begin VB.CommandButton SYNOLO 
         Caption         =   "’–œÀœ√…”Ãœ” ”’ÕœÀœ - –’"
         Height          =   540
         Left            =   6285
         TabIndex        =   4
         Top             =   2325
         Width           =   1590
      End
      Begin VB.TextBox SYNOLIKO 
         Height          =   285
         Left            =   6285
         TabIndex        =   3
         Top             =   1995
         Width           =   1590
      End
      Begin VB.TextBox PERIB 
         Height          =   330
         Left            =   2610
         TabIndex        =   2
         Top             =   2505
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   2595
         TabIndex        =   12
         Top             =   630
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   503
         _Version        =   393216
         Format          =   308740097
         CurrentDate     =   39117
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "–œ”œ ME ÷–¡"
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   1185
         TabIndex        =   21
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "A÷M META÷"
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   1185
         TabIndex        =   20
         Top             =   990
         Width           =   1110
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "AP.TIM."
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   1260
         TabIndex        =   19
         Top             =   1890
         Width           =   1140
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "%÷–¡ ıÔÎÔ„ÈÛÏÔı"
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   5745
         TabIndex        =   18
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "√—«√œ—«  ¡‘¡◊Ÿ—«”« ƒ≈À‘…ŸÕ ∆  ¡… Ã≈‘¡÷œ—… ŸÕ (Ã≈ Ã…¡  ¡‘«√œ—…¡ ÷–¡)"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1305
         TabIndex        =   17
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "–œ”œ  Z ME ÷–¡ –’"
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   4515
         TabIndex        =   16
         Top             =   1500
         Width           =   1665
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "”’ÕœÀ… œ –œ”œ "
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   4815
         TabIndex        =   15
         Top             =   2040
         Width           =   1350
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   " ¡».¡Œ…¡.–≈—…¬"
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   2490
         Width           =   1200
      End
   End
   Begin MSAdodcLib.Adodc KIN1 
      Height          =   330
      Left            =   5520
      Top             =   8400
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VanSales\pol.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VanSales\pol.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM KIN1"
      Caption         =   "kin1"
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   9000
      Width           =   4695
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   420
      Left            =   480
      Top             =   8280
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   741
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
   Begin VB.CommandButton Command2 
      Caption         =   "EÓÔ‰ÔÚ"
      Height          =   300
      Left            =   9570
      TabIndex        =   0
      Top             =   7680
      Width           =   1740
   End
End
Attribute VB_Name = "bohu5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim F_LINKEDSERVER As String '
Dim F_LINKED_PARAST As String





Dim F_DIR_APOS_KARAM As String '
Dim F_DIR_EPIS_KARAM As String

Dim F_DOUTSIOSREM As String
'Dim F_DOUTSIOSPARAS As String

Dim F_TIM    ' Val(FindParametroi(1,"APOT2", "F_PIChEIGHT", "2535", "’¯ÔÚ ÂÈÍ¸Ì·Ú")) 'posa psifia tha exei h kathe seira

Dim ELEM(100)

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Dim F_PY As String
Dim F_KATHG_METAG As Integer

Private Sub cmd≈–…”‘—œ÷«÷œ—‘«√œ’_Click()
        '<EhHeader>
        On Error GoTo cmd≈–…”‘—œ÷«÷œ—‘«√œ’_Click_Err
        '</EhHeader>
   ' On Error GoTo Rollback
    Dim RECS As Integer



    'ƒ≈Õ  —¡‘¡≈… ≈ –‘Ÿ”« EGGTIM
    ' DEN YPOLOGIZEI KAU_AJIA,MIK_AJIA
    'AJ1 DEN YPOLOGIZEI KAI FPA1 STO TIM




100 Gdb.BeginTrans

102 Gdb.Execute "DELETE FROM PTIM WHERE LEFT(ATIM,1)='Ù' "
104 Gdb.Execute "DELETE FROM PEGGTIM WHERE LEFT(ATIM,1)='Ù' "

    '-----------------1A,1B,1D----------------
    '  Gdb.Execute "UPDATE  PEGGTIM SET EKPT=(SELECT TOP 1  ISNULL(NUM1,0) FROM PEL WHERE KOD=PEGGTIM.PELKOD AND EIDOS='e')   WHERE ATIM LIKE 'T%' "
    '  Gdb.Execute "UPDATE  PEGGTIM SET EKPT=0   WHERE ATIM LIKE 'Ò%'"
106 Gdb.Execute "UPDATE PEGGTIM SET  KAU_AJIA=POSO*TIMM*(100-EKPT)/100"
    '-----------------1A,1B,1D----------------
    Dim N As Integer

108 For N = 1 To 7
110    Gdb.Execute "UPDATE PEGGTIM SET  MIK_AJIA=KAU_AJIA* " + str((g_Fpa(N) + 100) / 100) + " WHERE FPA= " + str(N)
    Next

112 Gdb.Execute "UPDATE PTIM SET FPA2 =0,FPA3 =0,FPA4 =0,FPA6 =0,FPA7=0,FPA8=0,FPA9=0,TYP=0,AJ2=0,AJ3=0,AJ4=0,AJ5=0,AJ6=0,AJ7=0,AJ8=0,AJ9=0,EKPT1=0,EKPT2=0,EKPT3=0,EKPT4=0,EKPT5=0"




    '----------------------2
    'Gdb.Execute "UPDATE PEGGTIM SET FPA=1"   '   ISNULL(FPA,1)  "

114 Gdb.Execute "UPDATE PTIM SET AJ2=ISNULL(AJ2,0)   "

116 Gdb.Execute "UPDATE PTIM SET  B_N1=1  "




118 Gdb.Execute "UPDATE PTIM SET  AJ1=(SELECT SUM(KAU_AJIA) FROM PEGGTIM WHERE FPA=1 AND PEGGTIM.ATIM=PTIM.ATIM )  "
120 Gdb.Execute "UPDATE PTIM SET  AJ2=(SELECT SUM(ISNULL(KAU_AJIA,0)) FROM PEGGTIM WHERE FPA=2 AND PEGGTIM.ATIM=PTIM.ATIM)  "

122 Gdb.Execute "UPDATE PTIM SET  FPA1=ISNULL(AJ1,0)*" + str(g_Fpa(1) / 100)
124 Gdb.Execute "UPDATE PTIM SET  FPA2=ISNULL(AJ2,0)*" + str(g_Fpa(2) / 100)
126 Gdb.Execute "UPDATE PTIM SET  FPA3=ISNULL(AJ3,0)*" + str(g_Fpa(3) / 100)
128 Gdb.Execute "UPDATE PTIM SET  FPA4=ISNULL(AJ4,0)*" + str(g_Fpa(4) / 100)

130 Gdb.Execute "UPDATE PTIM SET  FPA4=ISNULL(AJ6,0)*" + str(g_Fpa(6) / 100)
132 Gdb.Execute "UPDATE PTIM SET  FPA4=ISNULL(AJ7,0)*" + str(g_Fpa(7) / 100)


134 Gdb.Execute "UPDATE PTIM SET  AJI=ISNULL(AJ1,0)+ISNULL(FPA1,0)+ISNULL(AJ2,0)+ISNULL(FPA2,0)+ISNULL(AJ3,0)+ISNULL(FPA3,0)+ISNULL(AJ4,0)+ISNULL(FPA4,0)+ISNULL(AJ6,0)+ISNULL(FPA6,0)+ISNULL(AJ5,0)+ISNULL(AJ7,0)+ISNULL(FPA7,0)"

    '----------------------2

136  Gdb.Execute "UPDATE PTIM SET TRP='1;Ã≈' WHERE TRP LIKE 'Ã≈‘%';"
138  Gdb.Execute "UPDATE PTIM SET TRP='2;–…' WHERE TRP LIKE '–…”%';"



140 Gdb.Execute "INSERT INTO TIM (B_N1,EIDOS,KPE,HME,TRP,ATIM,ART,AJI,EKPT,EIDPAR,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,TYP,AJ1,AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9,EKPT1,EKPT2,EKPT3,EKPT4,EKPT5,KLEIDI) SELECT 1,'e',KPE,HME,TRP,ATIM,ART,AJI,EKPT,EIDPAR,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,TYP,AJ1,AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9,EKPT1,EKPT2,EKPT3,EKPT4,EKPT5,ATIM From PTIM", RECS

142 Gdb.Execute "UPDATE PEGGTIM SET APOT=1,EIDOS='e',XRE=0,PIS=POSO,KAU_AJIA=POSO*TIMM*(100-EKPT)/100,ID_NUM =(SELECT top 1 ID_NUM FROM TIM WHERE ATIM=PEGGTIM.ATIM),ONOMA = (select top 1 ONO FROM EID WHERE KOD=PEGGTIM.KODE)"





144 Gdb.Execute "INSERT INTO EGGTIM (FPA,PELKOD,ONOMA,ID_NUM,ATIM,HME,KODE,POSO,TIMM,APOT,EIDOS,XRE,PIS,KAU_AJIA,MIK_AJIA,EKPT) SELECT FPA,PELKOD,ONOMA,ID_NUM,ATIM,HME,KODE,POSO,TIMM,APOT,EIDOS,XRE,PIS,KAU_AJIA,MIK_AJIA,EKPT FROM PEGGTIM"

146 Gdb.Execute "INSERT INTO EGG (AIT,IDTIM,EIDOS,ATIM,HME,KOD,XRE,XREOSI,PISTOSI)  SELECT '‘…ÃœÀœ√…œ –ŸÀ«”«”',ID_NUM,'e',ATIM,HME,KPE,AJI,AJI,  0 FROM PTIM "   ' WHERE LEFT(TRP,1)='–'"
148 Gdb.Execute "INSERT INTO EGG (AIT,IDTIM,EIDOS,ATIM,HME,KOD,XRE,XREOSI,PISTOSI)  SELECT 'Ã≈‘—«‘¡' ,ID_NUM,'e',ATIM,HME,KPE,AJI,0  ,AJI FROM PTIM WHERE LEFT(TRP,1)='1'"

    'Gdb.Execute "delete from PTIM"

    'Gdb.Execute "delete from PEGGTIM"



150 Gdb.CommitTrans

152 MsgBox "OK TÈÏÔÎ¸„È· " + str(RECS)

























        '<EhFooter>
        Exit Sub

cmd≈–…”‘—œ÷«÷œ—‘«√œ’_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.bohu5.cmd≈–…”‘—œ÷«÷œ—‘«√œ’_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
               
          Gdb.RollbackTrans

156 MsgBox "ƒ≈Õ Ã≈‘¡÷≈—»« ¡Õ " + Chr(13) + Err.Description
               
               
               
       
        '</EhFooter>
End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim a

        Dim F928 As Integer

100     F928 = 1

110     If Len(Combo1.Text) = 0 Then
120         MsgBox "≈–…À≈Œ‘≈ ¡—◊≈…œ EID / PEL"

            Exit Sub

        End If

130     Adodc2.ConnectionString = gConnect
140     Adodc2.RecordSource = "select * FROM " + Combo1.Text    'PEL WHERE  EIDOS='e' ORDER BY KOD"
150     Adodc2.Refresh

        'Set fSCR = CreateObject("MSScriptControl.ScriptControl")
        'fSCR.language = "vbscript"
        'fSCR.AddObject "ADODC2", Adodc2
        'fSCR.AddObject "mactext", MACtEXT

160     a = toascii2(Text1.Text, Text2.Text)

170     Adodc2.RecordSource = "select * FROM EID ORDER BY KOD"
180     Adodc2.Refresh

        ' A = toascii2("C:\MERCVB\F33.TXT", "C:\EID.TXT")

190     MsgBox "‘›ÎÔÚ Àπÿ«”"

        'Me.Caption = mac("Adodc2.Recordset(3)")

        'Set fSCR = Nothing

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function toascii2(ByVal arxeio As String, moutput As String)

        ' arxeio .˜. f90.txt  Á ˆ¸ÒÏ· ÙÁÚ ÂÍÙ˝˘ÛÁÚ    host.txt to arxeio me ejagogi ascii
        '<EhHeader>
        On Error GoTo toascii2_Err

        '</EhHeader>

        Dim DUM

        Dim eField(1 To 50), ePict(1 To 50), eProw(1 To 50), ePcol(1 To 50), eFonts(1 To 50)

        Dim mField(1 To 50), mPict(1 To 50), mProw(1 To 50), mPcol(1 To 50), mFonts(1 To 50)

        Dim sField(1 To 50), sPict(1 To 50), sProw(1 To 50), sPcol(1 To 50), sFonts(1 To 50)

        Dim m_No_of_seir, i, k, m, c

        Dim m_seir_synol, m_seir_eid, sf

        Dim a(1 To 250), ar_ped(1 To 250)

        Dim xa(1 To 250), m_m, npic

        Dim m1

        ' dim gm_str(1 To 250)
        Dim DB     As Database

        Dim TIM    As Recordset

        Dim PEL    As Recordset

        Dim EGGTIM As Recordset

        Dim m_entol_ektyp

        Dim xart11, kod_tim

        Dim ejodos

        Dim e, MM, s

        Dim ektypoths

        Dim m_syn_row, syn_row

        Dim CDOK

        Dim h$

        ' spacing=1 klassiko poy ta ypologizei me +1 thn stili toy pedioy (x1)
        'spacing=0 kanonika
        '*********************************************************
        '-----------    GLOBAL Ã≈‘¡¬À«‘≈”  ----------------
        'gm_str() ÔÈ ÛÂÈÒ›Ú ÙÔı f99.txt
        'gpic(i) ÙÔ ÛÙÒÈÌ„Í ÙÔı Â‰ﬂÔı
        'gm_r(i),gm_c(i) ÛÂÈÒ‹ Í·È ÛÙﬁÎÁ ÙÔı Â‰ﬂÔı
        'gm_f(i) ÙÔ ¸ÌÔÏ· ÙÔı Â‰ﬂÔı

        'gm_str(6)="     ^XXXXXXXXXXXXXXXXXXXXXXXXXXX               ^xxxxxxxx    @XXXXXXXX  ~ono_par  ~SUBS(tim->atim,2,5)  ~TIM->HME
        '        gpic(2)='XXXXXXXXXXXXXXXXXXXXXXXXXXX
        'gm_r(2)=5  gm_c(2)=6    gm_f(2)=ono_par
        '========================================================
        'Dim gm_str(1 To 250) As String

100     If Len(Dir(Text2.Text, vbNormal)) = 0 Then
110         MsgBox "ƒ≈Õ ’–¡—◊≈… ‘œ ¡—◊≈…œ " + Text2.Text

            Exit Function

        End If

120     If Len(Dir(arxeio, vbNormal)) = 0 Then
130         MsgBox "ƒ≈Õ ’–¡—◊≈… ‘œ ¡—◊≈…œ " + arxeio

            Exit Function

        End If

        Dim ANS, meidos

140     DUM = bohu6.load2_forma(arxeio, 1)

150     If Combo1.Text = "PEL" Then
160         ANS = MsgBox("AıÙÔﬂ Ôı Ë· ÂÈÛ·˜ËÔ˝Ì ÂﬂÌ·È ÂÎ‹ÙÂÚ;", vbYesNo)

170         If ANS = vbYes Then meidos = "e" Else meidos = "r"
        End If

180     Open Text2.Text For Input As #1

        Dim S2 As String

190     Do While Not EOF(1)

200         Line Input #1, h$
210         S2 = ""
220         h$ = to928(h$)

230         For k = 1 To 30

240             If Trim(UCase(Trim(gm_f(k)))) = "KOD" Then    '  Then
250                 If Combo1.Text = "PEL" Then
260                     Adodc2.RecordSource = "SELECT * FROM " + Combo1.Text + " WHERE KOD='" + Trim(mID$(h$, gm_c(k), Len(gpic(k)))) + "' and EIDOS='" + meidos + "'"
                    Else
270                     Adodc2.RecordSource = "SELECT * FROM " + Combo1.Text + " WHERE KOD='" + Trim(mID$(h$, gm_c(k), Len(gpic(k)))) + "'"
                    End If

280                 Adodc2.Refresh

290                 If Adodc2.Recordset.EOF Then
300                     Adodc2.Recordset.AddNew

310                     If Combo1.Text = "PEL" Then
320                         Adodc2.Recordset("EIDOS") = meidos
                        End If
                    End If

                    Exit For

                End If

            Next

330         For k = 1 To 30

340             If Len(gm_f(k)) > 0 Then

350                 If Adodc2.Recordset(Trim(gm_f(k))).Type = adVarChar Then    ' adLongVarBinary Then 'adLongVarChar Then 'adLongVarWChar  Then '  adChar Then
360                     Adodc2.Recordset(Trim(gm_f(k))) = mID$(h$, gm_c(k), Len(Trim(gpic(k))))
                    End If

                    'adDouble

370                 If Adodc2.Recordset(Trim(gm_f(k))).Type = adDouble Then    ' adLongVarBinary Then 'adLongVarChar Then 'adLongVarWChar  Then '  adChar Then
380                     Adodc2.Recordset(Trim(gm_f(k))) = Val(mID$(h$, gm_c(k), Len(gpic(k))))
                    End If

                End If

            Next

390         DoEvents
400         Me.Caption = h$
410         Adodc2.Recordset.Update

            '   Print #1, S2
420         Adodc2.Recordset.MoveNext
        Loop

430     Close #1
440     Adodc2.Recordset.Close

450     If Combo1.Text = "EID" Then
460         Gdb.Execute "UPDATE EID SET FPA=2 WHERE FPA=19"
470         Gdb.Execute "UPDATE EID SET FPA=1 WHERE FPA=9"
        End If

        '<EhFooter>
        Exit Function

toascii2_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.toascii2 " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.toascii2 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

' Function find_eid_parastat()
'
'  Dim R As New ADODB.Recordset
'
'  R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
'  R.MoveFirst
'  Do While Not R.EOF
'     If Right(PARAS.Text, 1) = R("EIDOS") Then
'       F_TITLOS = R("TITLOS")
'       F_POS_APOU = R("POS_APOU")
'       F_AJIA_APOU = R("AJIA_APOU")
'       f_pel = R("pel")
'       f_pol = R("pol") ' 1=POLISEIS 2=AGORES
'       F_STADIO = R("STADIO") ' 1=–—œ”÷ 2=–¡—¡√√ 3=ƒ¡ 4=‘…Ã.–œÀ 5=≈Õƒœƒ…¡ …Õ…”…
'       If IsNull(R("GEF_P")) Then
'           F_rec_p = 0
'       Else
'           F_rec_p = R("GEF_P") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
'       End If
'       If IsNull(R("XONDR")) Then
'           F_XONDR = 0
'       Else
'           F_XONDR = R("XONDR") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
'       End If
'     End If
'     R.MoveNext
'  Loop
'
'
'End Function

Private Sub Command10_Click()

        '=============================================================================
        '<EhHeader>
        On Error GoTo Command10_Click_Err

        '</EhHeader>
        Dim R     As New ADODB.Recordset

        Dim s     As String

        Dim r2    As New ADODB.Recordset

        Dim EIDOS As String

100     EIDOS = "G','g','t','T','L','l"
110     EIDOS = InputBox("ƒŸ”≈ ‘¡ –¡—¡”‘¡‘… ¡ –œ’ »¡ ≈Õ«Ã≈—Ÿ»œ’Õ ", , EIDOS)

120     Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA= POSO*TIMM*(100-EKPT)/100*" + str((100 - g_Fpa(1)) / 100) + " where FPA=1 AND LEFT(ATIM,1) IN ('" + EIDOS + "')  "
130     Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA= POSO*TIMM*(100-EKPT)/100*" + str((100 - g_Fpa(2)) / 100) + " where FPA=2 AND LEFT(ATIM,1) IN ('" + EIDOS + "')  "
140     Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA= POSO*TIMM*(100-EKPT)/100*" + str((100 - g_Fpa(3)) / 100) + " where FPA=3 AND LEFT(ATIM,1) IN ('" + EIDOS + "')  "
150     Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA= POSO*TIMM*(100-EKPT)/100*" + str((100 - g_Fpa(4)) / 100) + " where FPA=4 AND LEFT(ATIM,1) IN ('" + EIDOS + "')  "
160     Gdb.Execute "UPDATE EGGTIM SET KAU_AJIA=POSO*TIMM*(100-EKPT)/100,MIK_AJIA= POSO*TIMM*(100-EKPT)/100*" + str((100 - g_Fpa(5)) / 100) + " where FPA=5 AND LEFT(ATIM,1) IN ('" + EIDOS + "')  "

170     s = "select sum(CASE WHEN FPA=1 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S1,"
180     s = s + " sum(CASE WHEN FPA=2 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S2,"
190     s = s + " sum(CASE WHEN FPA=3 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S3,"
200     s = s + " sum(CASE WHEN FPA=4 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S4,"
210     s = s + " sum(CASE WHEN FPA=5 THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END ) AS S5, "
220     s = s + "ATIM,HME FROM EGGTIM where LEFT(ATIM,1) IN ('" + EIDOS + "') GROUP BY ATIM,HME "

230     R.Open s, Gdb, adOpenDynamic, adLockOptimistic

        Dim Z As Single

240     R.MoveFirst

250     Do While Not R.EOF
260         Z = R("S1") * (100 + g_Fpa(1)) / 100
270         Z = Z + R("S2") * (100 + g_Fpa(2)) / 100
280         Z = Z + R("S3") * (100 + g_Fpa(3)) / 100
290         Z = Z + R("S4") * (100 + g_Fpa(4)) / 100
300         Z = Z + R("S5") * (100 + g_Fpa(5)) / 100

310         s = "UPDATE TIM SET AJ1=" + str(R("S1")) + ","
320         s = s + " AJ2=" + str(R("S2")) + ","
330         s = s + " AJ3=" + str(R("S3")) + ","
340         s = s + " AJ4=" + str(R("S4")) + ","
350         s = s + " AJ5=" + str(R("S5")) + ","
360         s = s + "AJI=" + str(Z) + " WHERE ATIM='" + R("ATIM") + "' AND HME='" + Format(R("HME"), "MM/DD/YYYY") + "'"
370         Gdb.Execute s

            'R2.Open "SELECT * FROM EGG  WHERE ATIM='" + R("ATIM") + "' AND HME='" + Format(R("HME"), "MM/DD/YYYY") + "'"", Gdb, adOpenDynamic, adLockOptimistic"

            'ÃœÕœ –…”‘œ”«  ¡Õ≈…
            '  S = "UPDATE EGG SET PISTOSI=" + Str(Z) + " WHERE ATIM='" + R("ATIM") + "' AND HME='" + Format(R("HME"), "MM/DD/YYYY") + "'"
            '  Gdb.Execute S

380         R.MoveNext
        Loop

390     R.Close

        '<EhFooter>
        Exit Sub

Command10_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Command10_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Command10_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command11_Click()

    On Error Resume Next

    Gdb.Execute "drop table DOKFORTHGO;"
    Gdb.Execute "DROP TABLE DOKSYGKENTROTIKO;"

    Gdb.Execute "DROP TABLE DOKSYGKTIM;"
 
    Dim parast_pol As String, sql As String

    parast_pol = "'t','T','A'"

    Dim MDAY As String

    On Error GoTo 0

    Dim m As Long

    MDAY = InputBox("ƒŸ”≈ «Ã≈—œÃ«Õ…¡ –—Ÿ‘œ’ –¡—¡”‘¡‘… œ’ ÷œ—‘«√œ’", , Format(Now, "DD/MM/YYYY"))

    If Not IsDate(MDAY) Then
        MsgBox "À¡»œ” «Ã≈—œÃ«Õ…¡"
    End If

    Dim MDATE As Date

    MDATE = CDate(MDAY)

    'Dim MATIM: MATIM = InputBox("ƒŸ”≈ ‘’–œ & ¡—…»Ãœ –—Ÿ‘œ’ –¡—¡”‘¡‘… œ’ ÷œ—‘«√œ’ –.◊. ‘15", , Format(Now, "DD/MM/YYYY"))
    'If InStr(Left(MATIM, 1), "T‘Ùt") > 0 Then
    '      MATIM = "T" + Format(Val(mID(MATIM, 2, 10)), "000000")
    'Else
    '      MATIM = "A" + Format(Val(mID(MATIM, 2, 10)), "000000")
    'End If

    'ƒ«Ã…œ’—√…¡ DOKFORTHGO ¡»—œ…”Ã¡‘œ” ‘…ÃœÀœ√…ŸÕ –œ’ –ŸÀ«»« ¡Õ
    'Gdb.Execute "DROP TABLE DOKSYGKENTROTIKO;"
    sql = "SELECT SUM(XRE) AS FORT ,SUM(PIS) AS POL , SUM(XRE)-SUM(PIS) AS YPOL,"
    sql = sql + " KODE INTO DOKFORTHGO  FROM EGGTIM  "
    sql = sql + " where LEFT(ATIM,1) IN (" + parast_pol + ")   AND HME>='" + Format(MDATE, "MM/DD/YYYY") + "'    GROUP BY KODE;"
    Gdb.Execute sql, m

    Dim RD As New ADODB.Recordset

    RD.Open "SELECT SUM(POL) FROM DOKFORTHGO ", Gdb, adOpenDynamic, adLockOptimistic
    MsgBox RD(0)
 
    RD.Close
 
    '000004
 
    '¡Õ‘…√—¡÷Ÿ ‘œ ‘≈À≈’‘¡…œ ”’√ ≈Õ‘—Ÿ‘… œ ”‘œ DOKSYGKENTRVTIKO APO TO EGGTIM
    Dim TEL_SYGK  As Integer

    Dim CTEL_SYGK As String
 
    TEL_SYGK = InputBox("ƒŸ”≈ ‘≈À≈’‘¡…œ ”’√ ≈Õ‘—Ÿ‘… œ", , "0")
    CTEL_SYGK = Right("000000" + Trim(LTrim(TEL_SYGK)), 6)

    Dim FF

    'Gdb.Execute "SELECT  *  INTO DOKSYGKENTROTIKO   FROM  EGGTIM WHERE ATIM='µ" + Format(TEL_SYGK, "000000") + "';", M

    'Gdb.Execute "DROP TABLE DOKSYGKENTROTIKO;"

    'Gdb.Execute "DROP TABLE DOKSYGKENTROTIKO;"
    Gdb.Execute "SELECT  *  INTO DOKSYGKENTROTIKO   FROM  EGGTIM WHERE ATIM='F" + CTEL_SYGK + "'", m

    Gdb.Execute "update  DOKSYGKENTROTIKO set  POSO=POSO-MIKTA WHERE MIKTA>0;", m
    Gdb.Execute "update  DOKSYGKENTROTIKO set  MIKTA=0"

    'TO DOKSYGKENTROTIKO ≈Õ«Ã≈—ŸÕ≈‘¡… ¡–œ ‘…” –ŸÀ«”≈…”  ¡… ‘…” ¬¡∆≈… ”‘¡ Ã… ‘¡
    Gdb.Execute "update  DOKSYGKENTROTIKO set MIKTA=DOKFORTHGO.POL    FROM  DOKSYGKENTROTIKO LEFT JOIN DOKFORTHGO ON DOKSYGKENTROTIKO.KODE=DOKFORTHGO.KODE;", m

    'Gdb.Execute "update  DOKSYGKENTROTIKO set  POSO=POSO-MIKTA WHERE MIKTA>0;"
    'Gdb.Execute "update  DOKSYGKENTROTIKO set   MIKTA=0;"

    'ANEBAZAI KATA 1 TON A—I»MO TOY ”’√ ≈Õ‘—Ÿ‘… œ’
    Gdb.Execute "update  DOKSYGKENTROTIKO set MIKTA=DOKFORTHGO.POL    FROM  DOKSYGKENTROTIKO INNER JOIN DOKFORTHGO ON DOKSYGKENTROTIKO.KODE=DOKFORTHGO.KODE;", m

    Gdb.Execute "UPDATE DOKSYGKENTROTIKO SET ATIM= SUBSTRING(ATIM,1,1)+RTRIM(RIGHT(CONVERT(VARCHAR(7),1000000+1+CONVERT(INT,SUBSTRING(ATIM,2,6))) ,6))+' '"

    Gdb.Execute "UPDATE DOKSYGKENTROTIKO SET HME='" + Format(Now, "MM/DD/YYYY") + "'"

    ' ≈…”¡√Ÿ ‘…” ≈√√—¡÷≈” ”‘œ EGGTIM
    Gdb.Execute "ALTER TABLE  EGGTIM DROP COLUMN ID"

    Gdb.Execute "ALTER TABLE  DOKSYGKENTROTIKO DROP COLUMN ID"

    Gdb.Execute "INSERT INTO EGGTIM  SELECT * FROM  DOKSYGKENTROTIKO"

    'ƒ«Ã…œ’—√Ÿ ¡Õ‘…√—¡÷œ ‘œ’ ”’√ ≈Õ‘—Ÿ‘… œ’ ¡––œ ‘œ ‘…Ã ”‘œ DOKSYGKTIM
    'Gdb.Execute "DROP TABLE DOKSYGKTIM;"
    Gdb.Execute "SELECT  *  INTO DOKSYGKTIM   FROM  TIM WHERE ATIM='F" + CTEL_SYGK + "'", m  '   'µ" + Format(TEL_SYGK, "000000") + "';"

    Gdb.Execute "UPDATE DOKSYGKTIM SET HME='" + Format(Now, "MM/DD/YYYY") + "'"
    Gdb.Execute "alter table  DOKSYGKTIM drop COLUMN ID_NUM"
    Gdb.Execute "UPDATE DOKSYGKTIM SET ATIM= SUBSTRING(ATIM,1,1)+RTRIM(RIGHT(CONVERT(VARCHAR(7),1000000+1+CONVERT(INT,SUBSTRING(ATIM,2,6))) ,6))+' '"

    Gdb.Execute "INSERT INTO TIM SELECT *  FROM   DOKSYGKTIM"
 
End Sub

Private Sub Command12_Click()

    '”’√ . ¡Õ¡◊Ÿ—«”«”
    Dim R    As New ADODB.Recordset

    Dim REID As New ADODB.Recordset

    REID.Open "select * FROM EID WHERE POS>0", Gdb, adOpenDynamic, adLockOptimistic

    R.Open "SELECT TOP 1 * FROM EGGTIM", Gdb, adOpenDynamic, adLockOptimistic

    Dim MDAY As String

    Dim m    As Long

    MDAY = InputBox("ƒŸ”≈ «Ã≈—œÃ«Õ…¡ ”’√ .÷œ—‘Ÿ”«”", , Format(Now, "DD/MM/YYYY"))

    If Not IsDate(MDAY) Then
        MsgBox "À¡»œ” «Ã≈—œÃ«Õ…¡"
    End If

    Dim MDATE As Date

    MDATE = CDate(MDAY)

    Dim cHME As String

    cHME = Format(MDATE, "MM/DD/YYYY")
 
    Dim M_ID_ARITMISI As Long

    Dim CAR_SYGK      As String

    Dim RD            As New ADODB.Recordset

    RD.Open "SELECT A.ARITMISI,A.ID FROM PARASTAT P INNER JOIN ARITMISI A ON P.ARITMISI=A.ID WHERE P.EIDOS='F'", Gdb, adOpenDynamic, adLockOptimistic
    CAR_SYGK = Format(RD(0) + 1, "000000")
    M_ID_ARITMISI = RD(1)
 
    RD.Close

    Gdb.Execute "UPDATE ARITMISI SET ARITMISI=" + CAR_SYGK + " WHERE ID=" + str(M_ID_ARITMISI)

    Do While Not REID.EOF
    
        R.AddNew
        R!ATIM = "F" + CAR_SYGK
        R!hme = MDATE
        R!KODE = REID!kod
        R!ONOMA = REID("ONO")
        R!POSO = REID!POS
        R!PIS = 0
        R!XRE = 0
        R!TIMM = 0
        R!FPA = REID!FPA
        R!apot = 1
        R!EIDOS = "e"
        R.Update
        
        REID.MoveNext
    Loop

    REID.Close
        
    R.Close
    
    Dim r3 As New ADODB.Recordset

    r3.Open "SELECT TOP 1 * FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
        
    r3.AddNew
    r3("HME") = MDATE
    r3!ATIM = "F" + CAR_SYGK
    r3!B_N1 = 1
    r3!trp = "Ã≈"
    r3!EIDOS = "e"
    r3!KPE = "9999"  '–≈—–¡‘œ’À«” ≈Ã–œ—… «
    r3!aj1 = 0
    r3!FPA1 = 0
    r3!aj2 = 0: r3!aj3 = 0: r3!aj4 = 0: r3!aj5 = 0: r3!aj6 = 0
    r3!fpa2 = 0: r3!fpa3 = 0: r3!FPA4 = 0
    r3!EKPT1 = 0
    r3!aji = 0
    r3.Update

    Exit Sub

    '------------------------------------------------------------- old
    On Error Resume Next

    Gdb.Execute "drop table DOKFORTHGO;"
    Gdb.Execute "DROP TABLE DOKSYGKENTROTIKO;"

    Gdb.Execute "DROP TABLE DOKSYGKTIM;"
 
    Dim parast_pol As String, sql As String

    parast_pol = "'t','T','A'"

    'Dim MDAY As String
    On Error GoTo 0

    'Dim m As Long

    MDAY = InputBox("ƒŸ”≈ «Ã≈—œÃ«Õ…¡", , Format(Now, "DD/MM/YYYY"))

    If Not IsDate(MDAY) Then
        MsgBox "À¡»œ” «Ã≈—œÃ«Õ…¡"
    End If

    'Dim MDATE As Date
    MDATE = CDate(MDAY)
 
    '¡Õ‘…√—¡÷Ÿ ‘œ ‘≈À≈’‘¡…œ ”’√ ≈Õ‘—Ÿ‘… œ ”‘œ DOKSYGKENTRVTIKO APO TO EGGTIM
    ' Dim TEL_SYGK As Integer
    'TEL_SYGK = InputBox("ƒŸ”≈ ‘≈À≈’‘¡…œ ”’√ ≈Õ‘—Ÿ‘… œ", , "0")
    ' TEL_SYGK = Right("000000" + LTrim(TEL_SYGK), 6)

    Dim TEL_SYGK  As Integer

    Dim CTEL_SYGK As String
 
    TEL_SYGK = InputBox("ƒŸ”≈ ‘≈À≈’‘¡…œ ”’√ ≈Õ‘—Ÿ‘… œ", , "0")
    CTEL_SYGK = Right("000000" + Trim(LTrim(TEL_SYGK)), 6)

    Dim FF

    Gdb.Execute "SELECT  *  INTO DOKSYGKENTROTIKO   FROM  EGGTIM WHERE ATIM='F" + CTEL_SYGK + "'"

    Gdb.Execute "update  DOKSYGKENTROTIKO set  POSO=POSO-MIKTA WHERE MIKTA>0;", m
    Gdb.Execute "update  DOKSYGKENTROTIKO set  MIKTA=0"

    'ANEBAZAI KATA 1 TON A—I»MO TOY ”’√ ≈Õ‘—Ÿ‘… œ’
    Gdb.Execute "UPDATE DOKSYGKENTROTIKO SET ATIM= SUBSTRING(ATIM,1,1)+RTRIM(RIGHT(CONVERT(VARCHAR(7),1000000+1+CONVERT(INT,SUBSTRING(ATIM,2,6))) ,6))+' '", m

    Gdb.Execute "UPDATE DOKSYGKENTROTIKO SET HME='" + Format(Now, "MM/DD/YYYY") + "'", m

    Gdb.Execute "DELETE FROM  DOKSYGKENTROTIKO WHERE POSO=0", m

    ' ≈…”¡√Ÿ ‘…” ≈√√—¡÷≈” ”‘œ EGGTIM
    Gdb.Execute "INSERT INTO EGGTIM  SELECT * FROM  DOKSYGKENTROTIKO", m

    'ƒ«Ã…œ’—√Ÿ ¡Õ‘…√—¡÷œ ‘œ’ ”’√ ≈Õ‘—Ÿ‘… œ’ ¡––œ ‘œ ‘…Ã ”‘œ DOKSYGKTIM
    'Gdb.Execute "DROP TABLE DOKSYGKTIM;"
    Gdb.Execute "SELECT  *  INTO DOKSYGKTIM   FROM  TIM WHERE ATIM='F" + CTEL_SYGK + "'", m '   'µ" + Format(TEL_SYGK, "000000") + "';"

    Gdb.Execute "UPDATE DOKSYGKTIM SET HME='" + Format(Now, "MM/DD/YYYY") + "'"
    Gdb.Execute "alter table  DOKSYGKTIM drop COLUMN ID_NUM"
    Gdb.Execute "UPDATE DOKSYGKTIM SET ATIM= SUBSTRING(ATIM,1,1)+RTRIM(RIGHT(CONVERT(VARCHAR(7),1000000+1+CONVERT(INT,SUBSTRING(ATIM,2,6))) ,6))+' '"

    Gdb.Execute "INSERT INTO TIM SELECT *  FROM   DOKSYGKTIM", m

End Sub

Private Sub Command13_Click()
Dim id_num(500) As Long ' ‘¡ ID_NUM ‘ŸÕ ‘…ÃœÀœ√…ŸÕ –œ’ –—œ”‘…»≈Õ‘¡…
    
    Dim mLet As String

    Gdb.BeginTrans

    On Error GoTo rollover
   
   
   
    Dim knok As Integer
   
   
    Dim a          As String

    Dim ATIM       As String, POSO As String, TIMM As String, ONOMA As String, kod As String

    Dim R          As New ADODB.Recordset

    Dim r2         As New ADODB.Recordset

    Dim aitiol     As String
   
    ' GoTo 1000
   
    '                                                                                       AJI+EKPT         AJ1            FPA1                             EKPT           AJI
    '3              19-03-2014 18:01:40 TI  B     03883                2   0.00            500.09          375.03          48.76           0.00            125.06          423.79          0.00            ?
    '4              19-03-2014 19:54:58 TI  B     03631                2   0.00            152.62          152.62          19.84           0.00            0.00            172.46          0.00            ?
    '5              20-03-2014 00:13:23 TI  B     60009                1   0.00            465.24          370.90          48.21           0.00            94.34           419.11          419.11          ?
    '6              20-03-2014 00:59:26 TI  B     60001                1   0.00            361.59          259.45          33.74           0.00            102.14          293.19          293.19          ?
    '7              20-03-2014 01:43:43 TI  B     05685                2   0.00            243.25          187.00          24.32           0.00            56.25           211.32          0.00            ?
    '8              20-03-2014 02:32:16 TI  B     05308                1   0.00            267.24          236.04          30.68           0.00            31.20           266.72          266.72          ?
   
    Dim MAX_ID_NUM As Long

    Dim MAX_ID     As Long
   
    R.Open "SELECT TOP 1 * FROM EGGTIM ORDER BY ID DESC ", Gdb, adOpenDynamic, adLockOptimistic
    MAX_ID = R("ID")
    R.Close
   
    R.Open "SELECT TOP 1 * FROM TIM ORDER BY ID_NUM DESC ", Gdb, adOpenDynamic, adLockOptimistic
    MAX_ID_NUM = R("ID_NUM")
    R.Close
   
    On Error GoTo 0

    Open "C:\VanSales\Mobile1\From\kin1.asc" For Input As #1
    
    R.Open "SELECT TOP 1 * FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
    r2.Open "SELECT TOP 1 * FROM EGG", Gdb, adOpenDynamic, adLockOptimistic
    
Dim j As Integer

    
Dim R6 As New ADODB.Recordset
    
    
    Dim akyr      As Integer

    Dim da        As Integer

    Dim pistotiko As Integer
    
    Do While Not EOF(1)  'KIN2.Recordset.EOF
        Input #1, a
        da = 0
        akyr = 0
        
        pistotiko = 0
        
        R.AddNew
        R("HME") = CDate(mID(a, 17, 10))

        If mID(a, 37, 2) = "‘…" Then
            mLet = "T"
            aitiol = "‘ÈÏÔÎ¸„ÈÔ ƒ¡ ”ÂÈÒ·Ú ¬"
        ElseIf mID(a, 37, 2) = "¡ " Then
            mLet = "ˆ"
            aitiol = "¡ÍıÒ˘ÙÈÍ¸"
            akyr = 1
        ElseIf mID(a, 37, 2) = "ƒ¡" Then
            mLet = "A"
            aitiol = "ƒÂÎÙÈÔ ¡ÔÛÙ.”ÂÈÒ·Ú ¬"
            da = 1
        
        ElseIf mID(a, 37, 2) = "ƒ–" Then
            mLet = ""
            aitiol = "ƒÂÎÙÈÔ –·Ò·Î·‚ﬁÚ"
            da = 1
        ElseIf mID(a, 37, 2) = "–…" Then
            mLet = "Ù"
            aitiol = "–ÈÛÙ˘ÙÈÍ¸ ‘ÈÏ.ƒ.–."
            pistotiko = 1
            
        Else
            MsgBox "–—œ”œ◊« ¡√ÕŸ”‘œ” ‘’–œ” –¡—¡”‘¡‘… œ’ ¡ ’—ŸÕ≈‘¡… « ≈…”¡√Ÿ√«"
            GoTo rollover
        End If
        
        R!ATIM = mLet + Format(Val(mID(a, 1, 6)), "000000")
        
        R!B_N1 = 1
        
        If mID(a, 68, 1) = "1" Then
            R!trp = "1.Ã≈"
        Else
            R!trp = "2.–…"
        End If

        R!EIDOS = "e"
        
        R!KPE = Trim(mID(a, 47, 7))
        
        'R2.Open "select * from EID WHERE KOD='" + Trim(mID(a, 27, 5)) + "'", Gdb, adOpenDynamic, adLockOptimistic
        'If R2.EOF Then
        '   R!ONOMA = ""
        'Else
        '    R!ONOMA = R2("ONO")
        'End If
        'R2.Close
        
     
        R!EKPT1 = Val(mID(a, 152, 8))
          R!aj1 = Val(mID(a, 104, 8))
        R!FPA1 = Val(mID(a, 120, 7))
        
        
        R!aji = Val(mID(a, 168, 8))
        
        If da = 1 Then
            R!aj6 = 0 ' Val(mID(a, 104, 8))
            R!FPA6 = 0 '  Val(mID(a, 120, 7))
            R!EKPT1 = 0 ' Val(mID(a, 152, 8))
            R!aji = 0 ' Val(mID(a, 168, 8))
        End If
        
        If akyr = 1 Then
            R!aj1 = -Val(mID(a, 104, 8))
            R!FPA1 = -Val(mID(a, 120, 7))
            R!EKPT1 = -Val(mID(a, 152, 8))
            R!aji = -Val(mID(a, 168, 8))
        End If
        
        R!aj2 = 0: R!aj3 = 0: R!aj4 = 0: R!aj5 = 0: R!aj6 = 0
        R!fpa2 = 0: R!fpa3 = 0: R!FPA4 = 0: R!FPA6 = 0
        
        
        
        
        R.Update
        
        
        
    R6.Open "SELECT MAX(ID_NUM) FROM TIM ", Gdb, adOpenDynamic, adLockOptimistic
    j = j + 1: id_num(j) = R6(0)
    R6.Close
        
        
        
        
        
        ' egg xreosi
        r2.AddNew
        r2!kod = Trim(mID(a, 47, 7))
        r2!EIDOS = "e"
        r2!hme = CDate(mID(a, 17, 10))
         
        r2!XRE = Val(mID(a, 168, 8))
        r2!PISTOSI = 0
        r2!XREOSI = Val(mID(a, 168, 8))
         
        If da = 1 Then
            r2!XRE = 0 ' Val(mID(a, 168, 8))
            r2!PISTOSI = 0
            r2!XREOSI = 0 ' Val(mID(a, 168, 8))
        End If
         
        If akyr = 1 Or pistotiko = 1 Then
            r2!XRE = -Val(mID(a, 168, 8))
            r2!PISTOSI = 0
            r2!XREOSI = -Val(mID(a, 168, 8))
        End If
         
        r2!ATIM = mLet + Format(Val(mID(a, 1, 6)), "000000")
        r2!AIT = aitiol
        
        r2!IDtim = id_num(j)
        
        r2.Update
         
        'metrhta
        If mID(a, 68, 1) = "1" Then
        
            ' egg xreosi
            r2.AddNew
            r2!kod = Trim(mID(a, 47, 7))
            r2!EIDOS = "e"
            r2!hme = CDate(mID(a, 17, 10))
          
            r2!XRE = Val(mID(a, 168, 8))
            r2!PISTOSI = Val(mID(a, 168, 8))
            r2!XREOSI = 0
          
            If da = 1 Then
                r2!XRE = 0 'Val(mID(a, 168, 8))
                r2!PISTOSI = 0
                r2!XREOSI = 0 ' Val(mID(a, 168, 8))
            End If
         
            If akyr = 1 Or pistotiko = 1 Then
                r2!XRE = -Val(mID(a, 168, 8))
                r2!PISTOSI = -Val(mID(a, 168, 8))
            End If
          
            r2!ATIM = mLet + Format(Val(mID(a, 1, 6)), "000000")
            r2!AIT = "ÏÂÙÒÁÙ·"
            
            r2!IDtim = id_num(j)
            
            r2.Update
        
        End If
        
    Loop

    R.Close
    r2.Close
    
    Close #1

    'C::\VanSales\Mobile1\From\kin2.asc
    '7              ÙÈ  ‚     219                  1               .8200           23.11           0               0.00         0.00         20
    '7              ‘…  ¬     117                  1               1.0700          23.13           0               0.00         0.00         20
   
    Open "C:\VanSales\Mobile1\From\kin2.asc" For Input As #1
    
    R.Open "SELECT TOP 1 * FROM EGGTIM", Gdb, adOpenDynamic, adLockOptimistic
    
     Dim KN As Integer
    
    Do While Not EOF(1)  'KIN2.Recordset.EOF
        
        da = 0
        akyr = 0
        
        Input #1, a
             
        R.AddNew
        
        If mID(a, 17, 2) = "‘…" Then
            mLet = "T"
        ElseIf mID(a, 17, 2) = "–…" Then
            mLet = "Ù"
            akyr = 0
        ElseIf mID(a, 17, 2) = "¡ " Then
            mLet = "ˆ"
            akyr = 1
            'Else
        ElseIf mID(a, 17, 2) = "ƒ¡" Then
            mLet = "A"
            da = 1
            
        ElseIf mID(a, 17, 2) = "ƒ–" Then
            mLet = ""
            da = 1
            
            
        Else   ' If mID(a, 17, 2) = "ƒ¡" Then
        
            MsgBox "–—œ”œ◊« ¡√ÕŸ”‘œ” ‘’–œ” –¡—¡”‘¡‘… œ’ ”≈  …Õ«”« " + mID(a, 17, 2)
            GoTo rollover

        
            'mLet = "A"
            'da = 1
            
        End If
        
        R!ATIM = mLet + Format(Val(mID(a, 1, 6)), "000000")
        
        ' r2.Close
        
        r2.Open "select * from TIM WHERE ATIM='" + mLet + Format(Val(mID(a, 1, 6)), "000000") + "'", Gdb, adOpenDynamic, adLockOptimistic
        knok = 0
        For KN = 1 To j
           If r2!id_num = id_num(KN) Then
               'ok  ·ÌÁÍÂÈ ÏÂÛ· ÛÙ· ÛÁÏÂÒÈÌ· ÙÈÏÔÎÔ„È·
               knok = r2!id_num
           End If
        Next
        
        If knok = 0 Then
            MsgBox "Î·ËÔÚ ÛÙÁÌ ‰ÔÏﬁ Ù˘Ì ÂÈÛÂÒ˜ÔÏ›Ì˘Ì"
            GoTo rollover
        End If
        


        If r2.EOF Then
            R("HME") = CDate(Format(Now, "MM/DD/YYYY"))
        Else
            R("HME") = r2!hme
        End If

        r2.Close
        
        R!KODE = Trim(mID(a, 27, 5))
        r2.Open "select * from EID WHERE KOD='" + Trim(mID(a, 27, 5)) + "'", Gdb, adOpenDynamic, adLockOptimistic

        If r2.EOF Then
            R!ONOMA = ""
            R!FPA = 1
        Else
            R!ONOMA = r2("ONO")
            R!FPA = r2!FPA
        End If
        
        r2.Close
        
        R!POSO = Val(mID(a, 138, 4))
        R!PIS = Val(mID(a, 138, 4))
        
        If akyr = 1 Then
            R!POSO = -Val(mID(a, 138, 4))
            R!PIS = -Val(mID(a, 138, 4))
        End If
        
        R!XRE = 0
        
        R!TIMM = Val(mID(a, 64, 6))
        
        R!EKPT = Val(mID(a, 79, 6))
        R!kau_ajia = R!TIMM * R!POSO * (100 - R!EKPT) / 100
        R!MIK_AJIA = R!kau_ajia * (100 + GGET_NVALUE("SELECT TOP 1 TIMH FROM PINAKES WHERE TYPOS=1 AND AYJON=" + str(R!FPA))) / 100
        R!apot = 1
        
        R!id_num = knok ' TO ID_NUM TOY TIMOLOGIOY
        
        R!EIDOS = "e"
        R.Update
        
        ' Gdb.Execute "INSERT INTO EGGTIM (ATIM,KODE,POSO,TIMM,ONOMA,"
         
    Loop
    
    R.Close
    
    Close #1
    
    'Do While Not KIN1.Recordset.EOF  'Flag_Ekk
    'KIN1.Recordset.Edit
    '  KIN1.Recordset("Flag_Ekk") = 0
    ' KIN1.Recordset.Update
    ' KIN1.Recordset.MoveNext
    ' Loop
   
   Dim SQL6 As String
   
    
   
   
   
    For KN = 1 To j
      'SQL6 = "select sum(CASE WHEN FPA=1 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA1,"
      'SQL6 = SQL6 + " sum(CASE WHEN FPA=2 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA2 ,"
      'SQL6 = SQL6 + " sum(CASE WHEN FPA=3 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA3 ,"
      'SQL6 = SQL6 + " sum(CASE WHEN FPA=4 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA4 ,"
      'SQL6 = SQL6 + " sum(CASE WHEN FPA=5 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA5 "
      'SQL6 = SQL6 + "  FROM EGGTIM WHERE ID_NUM=" + str(ID_NUM(kn))
      'R6.Open SQL6, Gdb, adOpenDynamic, adLockOptimistic
      
      SQL6 = "select sum(CASE WHEN FPA=1 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA1,"
      SQL6 = SQL6 + " sum(CASE WHEN FPA=2 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA2 ,"
      SQL6 = SQL6 + " sum(CASE WHEN FPA=3 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA3 ,"
      SQL6 = SQL6 + " sum(CASE WHEN FPA=4 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA4 ,"
      
      SQL6 = SQL6 + " sum(CASE WHEN FPA=6 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA6 ,"
      
      SQL6 = SQL6 + " sum(CASE WHEN FPA=5 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA5 "
      SQL6 = SQL6 + "  FROM EGGTIM WHERE ID_NUM=" + str(id_num(KN))
      R6.Open UCase(SQL6), Gdb, adOpenDynamic, adLockOptimistic
      
      
      
      
      
      
      
      
      
      
      
   
      r2.Open "SELECT * FROM TIM WHERE ID_NUM=" + str(id_num(KN)), Gdb, adOpenDynamic, adLockOptimistic
      
      If Abs(NNUL(R6!FPA1) + NNUL(R6!fpa2) + NNUL(R6!fpa3) + NNUL(R6!FPA4) + NNUL(R6!FPA5) + NNUL(R6!FPA6) - (r2!aj6 + r2!aj1 + r2!aj2 + r2!aj3 + r2!aj4 + r2!aj5)) < 0.02 Then
   
          Gdb.Execute "UPDATE  TIM SET AJ2=" + Replace(str(Round(NNUL(R6!fpa2), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
          Gdb.Execute "UPDATE  TIM SET AJ3=" + Replace(str(Round(NNUL(R6!fpa3), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
          Gdb.Execute "UPDATE  TIM SET AJ4=" + Replace(str(Round(NNUL(R6!FPA4), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
          
          Gdb.Execute "UPDATE  TIM SET AJ6=" + Replace(str(Round(NNUL(R6!FPA6), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
          Gdb.Execute "UPDATE  TIM SET AJ1=" + Replace(str(Round(NNUL(R6!FPA1), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
          
          Gdb.Execute "UPDATE  TIM SET AJ5=" + Replace(str(Round(NNUL(R6!FPA5), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
          'Gdb.Execute "UPDATE  TIM SET AJ1=AJ1-AJ2-AJ3-AJ4-AJ5 WHERE ID_NUM=" + str(ID_NUM(KN))
      
      End If
      R6.Close
      r2.Close
      
    Next
    
    Gdb.CommitTrans
    MsgBox "ok"
    
    create_sygk_epis MAX_ID_NUM, MAX_ID
    
    Exit Sub
    
rollover:
    Gdb.RollbackTrans
    MsgBox "‰ÂÌ ·ÔËÁÍÂıÙÁÍ·Ì"
    
    'PAR2_NUMBER   PAR2_APO_CODE   PAR2_APO_PERI                  PAR2_POSO1  PAR2_POSO2  PAR2_TIMH   PAR2_EKP    PAR2_TIMH_EFK     PAR2_KEY
    '  ‘…      13  0104026         MARS KING SIZE –¡√Ÿ‘œ 24 ‘≈/ ¬  1                       1.4200         0           0              ‘…¬13

End Sub



Function NNUL(ByVal c) As Single

If IsNull(c) Then
    NNUL = 0
Else
    NNUL = c
End If



End Function







Sub create_sygk_epis(ByVal IDtim As Long, ByVal IDeggtim As Long)

    On Error Resume Next

    Gdb.Execute "drop table DOKFORTHGO;"
    Gdb.Execute "DROP TABLE DOKSYGKENTROTIKO;"

    Gdb.Execute "DROP TABLE DOKSYGKTIM;"
 
    Dim parast_pol As String, sql As String

    parast_pol = "'t','T','A'"

    Dim MDAY As String

    On Error GoTo 0

    Dim m As Long

    'MDAY = InputBox("ƒŸ”≈ «Ã≈—œÃ«Õ…¡ –—Ÿ‘œ’ –¡—¡”‘¡‘… œ’ ÷œ—‘«√œ’", , Format(Now, "DD/MM/YYYY"))
    'If Not IsDate(MDAY) Then
    '   MsgBox "À¡»œ” «Ã≈—œÃ«Õ…¡"
    'End If
    'Dim MDATE As Date
    'MDATE = CDate(MDAY)

    'Dim MATIM: MATIM = InputBox("ƒŸ”≈ ‘’–œ & ¡—…»Ãœ –—Ÿ‘œ’ –¡—¡”‘¡‘… œ’ ÷œ—‘«√œ’ –.◊. ‘15", , Format(Now, "DD/MM/YYYY"))
    'If InStr(Left(MATIM, 1), "T‘Ùt") > 0 Then
    '      MATIM = "T" + Format(Val(mID(MATIM, 2, 10)), "000000")
    'Else
    '      MATIM = "A" + Format(Val(mID(MATIM, 2, 10)), "000000")
    'End If

    'ƒ«Ã…œ’—√…¡ DOKFORTHGO ¡»—œ…”Ã¡‘œ” ‘…ÃœÀœ√…ŸÕ –œ’ –ŸÀ«»« ¡Õ
    'Gdb.Execute "DROP TABLE DOKSYGKENTROTIKO;"
    sql = "SELECT SUM(XRE) AS FORT ,SUM(PIS) AS POL , SUM(XRE)-SUM(PIS) AS YPOL,"
    sql = sql + " KODE INTO DOKFORTHGO  FROM EGGTIM  "
    sql = sql + " where LEFT(ATIM,1) IN (" + parast_pol + ")   AND ID>" + str(IDeggtim) + "  GROUP BY KODE;"
    Gdb.Execute sql, m

    Dim RD As New ADODB.Recordset

    RD.Open "SELECT SUM(POL) FROM DOKFORTHGO ", Gdb, adOpenDynamic, adLockOptimistic
    MsgBox RD(0)
 
    RD.Close
 
    '000004
 
    '¡Õ‘…√—¡÷Ÿ ‘œ ‘≈À≈’‘¡…œ ”’√ ≈Õ‘—Ÿ‘… œ ”‘œ DOKSYGKENTRVTIKO APO TO EGGTIM
    Dim TEL_SYGK      As Integer

    Dim CTEL_SYGK     As String
 
    'RD.Open "SELECT A.ARITMISI FROM PARASTAT P INNER JOIN ARITMISI A ON P.ARITMISI=A.ID WHERE P.EIDOS='F'", Gdb, adOpenDynamic, adLockOptimistic
    'CTEL_SYGK = Format(RD(0), "000000")
    'RD.Close
 
    Dim M_ID_ARITMISI As Long

    Dim CAR_SYGK      As String
 
    RD.Open "SELECT A.ARITMISI,A.ID FROM PARASTAT P INNER JOIN ARITMISI A ON P.ARITMISI=A.ID WHERE P.EIDOS='F'", Gdb, adOpenDynamic, adLockOptimistic
    CTEL_SYGK = Format(RD(0), "000000")
    CAR_SYGK = Format(RD(0) + 1, "000000")
    M_ID_ARITMISI = RD(1)
    RD.Close
    Gdb.Execute "UPDATE ARITMISI SET ARITMISI=" + CAR_SYGK + " WHERE ID=" + str(M_ID_ARITMISI)
 
    'TEL_SYGK = InputBox("ƒŸ”≈ ‘≈À≈’‘¡…œ ”’√ ≈Õ‘—Ÿ‘… œ", , "0")
    'CTEL_SYGK = Right("000000" + Trim(LTrim(TEL_SYGK)), 6)

    Dim FF

    Gdb.Execute "SELECT  *  INTO DOKSYGKENTROTIKO   FROM  EGGTIM WHERE ATIM='F" + CTEL_SYGK + "'", m
    Gdb.Execute "update  DOKSYGKENTROTIKO set  POSO=POSO-MIKTA WHERE MIKTA>0;", m
    Gdb.Execute "update  DOKSYGKENTROTIKO set  MIKTA=0"

    'TO DOKSYGKENTROTIKO ≈Õ«Ã≈—ŸÕ≈‘¡… ¡–œ ‘…” –ŸÀ«”≈…”  ¡… ‘…” ¬¡∆≈… ”‘¡ Ã… ‘¡
    Gdb.Execute "update  DOKSYGKENTROTIKO set MIKTA=DOKFORTHGO.POL    FROM  DOKSYGKENTROTIKO LEFT JOIN DOKFORTHGO ON DOKSYGKENTROTIKO.KODE=DOKFORTHGO.KODE;", m

    'ANEBAZAI KATA 1 TON A—I»MO TOY ”’√ ≈Õ‘—Ÿ‘… œ’
    Gdb.Execute "update  DOKSYGKENTROTIKO set MIKTA=DOKFORTHGO.POL    FROM  DOKSYGKENTROTIKO INNER JOIN DOKFORTHGO ON DOKSYGKENTROTIKO.KODE=DOKFORTHGO.KODE;", m
    Gdb.Execute "UPDATE DOKSYGKENTROTIKO SET ATIM= SUBSTRING(ATIM,1,1)+RTRIM(RIGHT(CONVERT(VARCHAR(7),1000000+1+CONVERT(INT,SUBSTRING(ATIM,2,6))) ,6))+' '"
    Gdb.Execute "UPDATE DOKSYGKENTROTIKO SET HME='" + Format(Now, "MM/DD/YYYY") + "'"

    ' ≈…”¡√Ÿ ‘…” ≈√√—¡÷≈” ”‘œ EGGTIM
    'Gdb.Execute "ALTER TABLE  EGGTIM DROP COLUMN ID"

    Gdb.Execute "ALTER TABLE  DOKSYGKENTROTIKO DROP COLUMN ID"

    Dim fields_eggtim As String

    'fields_eggtim = "(EIDOS,ATIM,POSO,MONA,TIMM,KERDOS,KODE,HME,ERGO,FPA,PROOD,PROOD_AJ,EKPT,KAU_AJIA,MIK_AJIA,ONOMA,MIKTA,KOLA,PELKOD,PROELEYSH,XRE,PIS,APOT,ATIM2,FCURRENCY,EKPT2)"

    'Gdb.Execute "INSERT INTO EGGTIM " + fields_eggtim + "   SELECT " + fields_eggtim + " FROM  DOKSYGKENTROTIKO"'

    fields_eggtim = "EIDOS,ATIM,POSO,MONA,TIMM,KERDOS,KODE,HME,ERGO,FPA,PROOD,PROOD_AJ,EKPT,KAU_AJIA,MIK_AJIA,ONOMA,MIKTA,KOLA,PELKOD,PROELEYSH,XRE,PIS,APOT,ATIM2,FCURRENCY,EKPT2"

    Gdb.Execute "INSERT INTO EGGTIM (" + fields_eggtim + ")   SELECT " + fields_eggtim + " FROM  DOKSYGKENTROTIKO"

    'ƒ«Ã…œ’—√Ÿ ¡Õ‘…√—¡÷œ ‘œ’ ”’√ ≈Õ‘—Ÿ‘… œ’ ¡––œ ‘œ ‘…Ã ”‘œ DOKSYGKTIM
    'Gdb.Execute "DROP TABLE DOKSYGKTIM;"
    Gdb.Execute "SELECT  *  INTO DOKSYGKTIM   FROM  TIM WHERE ATIM='F" + CTEL_SYGK + "'", m  '   'µ" + Format(TEL_SYGK, "000000") + "';"

    Gdb.Execute "UPDATE DOKSYGKTIM SET HME='" + Format(Now, "MM/DD/YYYY") + "'"
    Gdb.Execute "alter table  DOKSYGKTIM drop COLUMN ID_NUM"
    Gdb.Execute "UPDATE DOKSYGKTIM SET ATIM= SUBSTRING(ATIM,1,1)+RTRIM(RIGHT(CONVERT(VARCHAR(7),1000000+1+CONVERT(INT,SUBSTRING(ATIM,2,6))) ,6))+' '"

    Gdb.Execute "INSERT INTO TIM SELECT *  FROM   DOKSYGKTIM"

End Sub

Private Sub Command14_Click()
    '-------------------------------------¡–œ  ¡—¡Ã≈”…Õ«  ---------
      APOKARAM 1
End Sub

Sub APOKARAM(ByVal FORT As Integer)
        '<EhHeader>
        On Error GoTo APOKARAM_Err
        '</EhHeader>


        Dim cnn           As ADODB.Connection

        Dim RDB           As ADODB.Recordset

        Dim strConnection As String, strProvider As String, strSource As String

100     strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
        ' strSource = "Data Source=\ADOPROG2\GIORAN\DB.mdb;"  '& App.Path &
102     strSource = "Data Source=" + F_DIR_EPIS_KARAM + "\DB.mdb;" '& App.Path &

104     Set cnn = New ADODB.Connection
106     strConnection = strProvider & strSource & "Persist Security Info=False"
108     cnn.Open strConnection
110     Set RDB = New ADODB.Recordset

        Dim id_num(500) As Long ' ‘¡ ID_NUM ‘ŸÕ ‘…ÃœÀœ√…ŸÕ –œ’ –—œ”‘…»≈Õ‘¡…

        Dim DB          As Database

        'Dim rDB       As Recordset
               
        'Set RDB = dB.OpenRecordset("eggtim")

        '  PaperCode   PaperAA PaperDate   PaperTime   CustomerCodeWhoIsCharged    BranchCode  WayOfPayment    PaperDiscount   AmountToBeSubtractedFromCash    Printed  PaperValue  Synolo_FPA    Sxolio_gia_ERP  CustomerCode    CustomerCentralShopCode CustomerOmilosCode  RelatedPaperCode    RelatedPaperAA  RelatedPaperSeira   Canceled
        '0               23488   30-12-2020  08:10:55    0297                                             2   0                          0                          Õ·È 34.80                        8.35                    0297    0521067688                  º˜È
        '0               23490   30-12-2020  09:24:33    0307                                             2   0                          0                          Õ·È 81.05                       13.84                    0307    6942472665                  º˜È
        '0               23489   30-12-2020  08:39:20    0341                                             2   0                          0                          Õ·È 59.00                        8.55                    0341    2510-517 264                    º˜È
        '8               2468    30-12-2020  08:40:36    0341                                             2   0                          0                          Õ·È 33.00                           0                    0341    2510-517 264                    º˜È
        '
    
        Dim mLet        As String

112     Gdb.BeginTrans

        On Error GoTo rollover
   
        '  Set dB = OpenDatabase("c:\CL\DATA\DB.mdb")
   
        Dim knok       As Long
   
        Dim a          As String

        Dim ATIM       As String, POSO As String, TIMM As String, ONOMA As String, kod As String

        Dim R          As New ADODB.Recordset

        Dim r2         As New ADODB.Recordset

        Dim aitiol     As String
   
        Dim MAX_ID_NUM As Long

        Dim MAX_ID     As Long
   
    '    R.Open "SELECT TOP 1 * FROM EGGTIM ORDER BY ID DESC ", Gdb, adOpenDynamic, adLockOptimistic
     '   MAX_ID = R("ID")
      '  R.Close
   
114     R.Open "SELECT TOP 1 *  FROM TIM ORDER BY ID_NUM DESC ", Gdb, adOpenDynamic, adLockOptimistic
       ' MAX_ID_NUM = R("ID_NUM")
116     R.Close
   
        On Error GoTo APOKARAM_Err

        ' Open "C:\VanSales\Mobile1\From\kin1.asc" For Input As #1
       ' cnn.Execute "update Papers set PaperValue=0 where PaperValue is null "
    
    
    'where papercode<>'12'
118     RDB.Open "SELECT * FROM Papers  ", cnn, adOpenDynamic, adLockOptimistic  ' where canceled=false and papercode<>12
    
120     R.Open "SELECT TOP 1 * FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
122     r2.Open "SELECT TOP 1 * FROM EGG", Gdb, adOpenDynamic, adLockOptimistic
    
        Dim j         As Integer

        Dim N         As Long
    
        Dim R6        As New ADODB.Recordset
    
        Dim akyr      As Integer

        Dim da        As Integer

        Dim pistotiko As Integer
    
        Dim mSS As String
124     mSS = mID$("¬√", FORT, 1)
    
126     Do While Not RDB.EOF  'KIN2.Recordset.EOF
        
128         da = 0
130         akyr = 0
        
132         pistotiko = 0
        
            ' R.AddNew
            ' R("HME") = RDB!PaperDate ' CDate(mID(a, 17, 10))

134         If RDB!paperCode = 0 Then
                ' mLet = "T"   ' MLET=MID("TQ",FORT,1)
136             mLet = mID("TQ", FORT, 1)
138             aitiol = "‘ÈÏÔÎ¸„ÈÔ ƒ¡ ”ÂÈÒ·Ú " + mSS
                '        ElseIf mID(a, 37, 2) = "¡ " Then
                '            mLet = "ˆ"
                '            aitiol = "¡ÍıÒ˘ÙÈÍ¸"
                '            akyr = 1
140         ElseIf RDB!paperCode = 1 Then
                ' mLet = "A"
142             mLet = mID("AW", FORT, 1)
144             aitiol = "ƒÂÎÙÈÔ ¡ÔÛÙ.”ÂÈÒ·Ú " + mSS
146             da = 1
        
148         ElseIf RDB!paperCode = 3 Then
                ' mLet = ""
150             mLet = mID("q", FORT, 1)  ' htan U kai egine : q
152             aitiol = "ƒÂÎÙÈÔ –·Ò·Î·‚ﬁÚ ”ÂÈÒ‹Ú " + mSS
154             da = 1
156         ElseIf RDB!paperCode = 2 Then
                'mLet = "Ù"
158             mLet = mID("ÙI", FORT, 1)
160             aitiol = "–ÈÛÙ˘ÙÈÍ¸ ‘ÈÏ.ƒ.–. ”ÂÈÒ‹Ú " + mSS
162             pistotiko = 1

            ElseIf Val(RDB!paperCode) = 12 Then
                'mLet = "Ù"
                 mLet = mID("UO", FORT, 1)
                 aitiol = "¡ÍıÒ˘ÙÈÍÔ " + mSS
                 akyr = 1
164         ElseIf RDB!paperCode = 8 Then  ' ÎÁÒ˘ÏÁ ÂÌ·ÌÙÈ
        
              ' egg xreosi
166             r2.AddNew
168             r2!kod = RDB!CustomerCodeWhoIsCharged  ' Trim(mID(a, 47, 7))
170             r2!EIDOS = "e"
172             r2!hme = RDB!PaperDate ' CDate(mID(a, 17, 10))
          
174             r2!XRE = Val(RDB!PaperValue)  ' Val(mID(a, 168, 8))
176             r2!PISTOSI = Val(RDB!PaperValue)   'Val(mID(a, 168, 8))
178             r2!XREOSI = 0
          
180             If da = 1 Then
182                 r2!XRE = 0 'Val(mID(a, 168, 8))
184                 r2!PISTOSI = 0
186                 r2!XREOSI = 0 ' Val(mID(a, 168, 8))
                End If
         
           
          
188             r2!ATIM = mSS + Format(Val(RDB!paperaa), "000000")
190             r2!AIT = mSS + " ÏÂÙÒÁÙ·"
            
192             r2!IDtim = id_num(j)
            
194             r2.Update
196             GoTo parakato
            Else
198             MsgBox "–—œ”œ◊« ¡√ÕŸ”‘œ” ‘’–œ” –¡—¡”‘¡‘… œ’ " + RDB!paperCode + "  ¡ ’—ŸÕ≈‘¡… « ≈…”¡√Ÿ√«"
200             GoTo rollover
            
            End If
       
            ' R!ATIM = mLet + Format(RDB!paperaa, "000000")
            Dim NLONG As Long
        
202         If GGET_NVALUE("select count(*) from TIM where ATIM='" + mLet + Format(RDB!paperaa, "000000") + "'") > 0 Then
204             MsgBox "’–¡—◊œ’Õ «ƒ« ‘¡ –¡—¡”‘¡‘… ¡"
206             Gdb.RollbackTrans
            
                Exit Sub
            End If
        
            
              ' Exit Sub

208         Gdb.Execute "insert into TIM (HME,ATIM,KLEIDI) VALUES ('" + Format(RDB!PaperDate, "mm/dd/yyyy") + "','" + mLet + Format(RDB!paperaa, "000000") + "','" + mLet + Format(RDB!paperaa, "000000") + "')", N

210         If N > 0 Then
212             NLONG = GGET_NVALUE("SELECT MAX(ID_NUM) FROM TIM")  ' CAST(scope_identity() AS int)")
214             j = j + 1: id_num(j) = NLONG


216              Gdb.Execute "UPDATE TIM SET ENTITYMARK='" + RDB!AADEMARK + "',ENTITYUID='" + RDB!AADEUID + "',QRURL='" + RDB!AADEURL + "' WHERE ID_NUM=" + str(NLONG)


                Dim cTrp As String

218             If RDB!WayOfPayment = 2 Then
220                 cTrp = "1.Ã≈"
                Else
222                 cTrp = "2.–…"
                End If
             
224             Gdb.Execute "UPDATE TIM SET KPE='" + CNull(RDB!CustomerCodeWhoIsCharged) + "', B_N1=1,EIDOS='e',TRP='" + cTrp + "' WHERE ID_NUM=" + str(NLONG)
         
226             Gdb.Execute "UPDATE TIM SET EKPT1=0,AJ1=" + str(Val(RDB!PaperValue)) + ",FPA1=" + str(Val(RDB!SynoloFPA)) + " WHERE ID_NUM=" + str(NLONG)

228             If da = 1 Then
                    ' Gdb.Execute "UPDATE TIM SET ... WHERE ID_NUM=" + str(NLONG)
230                 Gdb.Execute "UPDATE TIM SET AJ6=0,FPA6=0,AJI=0 WHERE ID_NUM=" + str(NLONG)
            
                End If
        
                '        If akyr = 1 Then
                '            R!aj1 = -Val(mID(a, 104, 8))
                '            R!fpa1 = -Val(mID(a, 120, 7))
                '            R!EKPT1 = -Val(mID(a, 152, 8))
                '            R!aji = -Val(mID(a, 168, 8))
                '        End If
232             Gdb.Execute "UPDATE TIM SET AJ2 = 0, AJ3 = 0, AJ4 = 0, AJ5 = 0, AJ6 = 0,FPA2 = 0,FPA3 = 0, FPA4 = 0, FPA6 = 0 WHERE ID_NUM=" + str(NLONG)
234             Gdb.Execute "UPDATE   TIM SET AJI=AJ1+FPA1,B_N1=1,AJ7=0,AJ8=0,AJ9=0,FPA7=0,FPA8=0,FPA9=0   WHERE ID_NUM=" + str(NLONG)
            Else
                ' ƒ≈Õ ¡–œ»« ≈’”≈ ‘…–œ‘¡
            End If
        
            ' egg xreosi
236         r2.AddNew
238         r2!kod = RDB!CustomerCodeWhoIsCharged  ' Trim(mID(a, 47, 7))
240         r2!EIDOS = "e"
242         r2!hme = RDB!PaperDate ' CDate(mID(a, 17, 10))
         
244         r2!XRE = Val(RDB!SynoloFPA) + Val(RDB!PaperValue)   'Val(mID(a, 168, 8))
246         r2!PISTOSI = 0
248         r2!XREOSI = Val(RDB!SynoloFPA) + Val(RDB!PaperValue)   ' Val(mID(a, 168, 8))
         
250         If da = 1 Then
252             r2!XRE = 0 ' Val(mID(a, 168, 8))
254             r2!PISTOSI = 0
256             r2!XREOSI = 0 ' Val(mID(a, 168, 8))
            End If
         
258         If akyr = 1 Or pistotiko = 1 Then
260             r2!XRE = -Val(RDB!SynoloFPA) - Val(RDB!PaperValue)   ' Val(mID(a, 168, 8))
262             r2!PISTOSI = 0
264             r2!XREOSI = -Val(RDB!SynoloFPA) - Val(RDB!PaperValue)  ' Val(mID(a, 168, 8))
            End If
         
266         r2!ATIM = mLet + Format(Val(RDB!paperaa), "000000")
268         r2!AIT = aitiol
        
270         r2!IDtim = id_num(j)
        
272         r2.Update
         
            'metrhta
274         If RDB!WayOfPayment = 2 Then
        
                ' egg xreosi
276             r2.AddNew
278             r2!kod = RDB!CustomerCodeWhoIsCharged  ' Trim(mID(a, 47, 7))
280             r2!EIDOS = "e"
282             r2!hme = RDB!PaperDate ' CDate(mID(a, 17, 10))
          
284             r2!XRE = Val(RDB!SynoloFPA) + Val(RDB!PaperValue) ' Val(mID(a, 168, 8))
286             r2!PISTOSI = Val(RDB!SynoloFPA) + Val(RDB!PaperValue)  'Val(mID(a, 168, 8))
288             r2!XREOSI = 0
          
290             If da = 1 Then
292                 r2!XRE = 0 'Val(mID(a, 168, 8))
294                 r2!PISTOSI = 0
296                 r2!XREOSI = 0 ' Val(mID(a, 168, 8))
                End If
         
298             If akyr = 1 Or pistotiko = 1 Then
300                 r2!XRE = -Val(RDB!SynoloFPA) - Val(RDB!PaperValue)
302                 r2!PISTOSI = -Val(RDB!SynoloFPA) - Val(RDB!PaperValue)
                End If
          
304             r2!ATIM = mLet + Format(Val(RDB!paperaa), "000000")
306             r2!AIT = "ÏÂÙÒÁÙ·"
            
308             r2!IDtim = id_num(j)
            
310             r2.Update
        
            End If
parakato:
312         RDB.MoveNext
        
        Loop
    
314     RDB.Close

316     R.Close
318     r2.Close
    
         'Gdb.CommitTrans
  
        ' Open "C:\VanSales\Mobile1\From\kin2.asc" For Input As #1
   
       '----------------------------- EGGTIM ---------------------------------------------------------------------------
    
320     RDB.Open "select * from PaperProducts     ", cnn, adOpenDynamic, adLockOptimistic ' where canceled=false and papercode<>12 ORDER BY PaperAA
     
322     R.Open "SELECT TOP 1 * FROM EGGTIM", Gdb, adOpenDynamic, adLockOptimistic
    
        Dim KN As Integer
    
324 RDB.MoveFirst


326 Do While Not RDB.EOF  'KIN2.Recordset.EOF
        
         '   On Error GoTo APOKARAM_Err
        
        
328         da = 0
330         akyr = 0
        
            '  Input #1, a
             
332         R.AddNew
        
334         If RDB!paperCode = 0 Then
               ' mLet = "T"
336              mLet = mID("TQ", FORT, 1)
338             aitiol = "‘ÈÏÔÎ¸„ÈÔ ƒ¡ ”ÂÈÒ·Ú ¬"
                '        ElseIf mID(a, 37, 2) = "¡ " Then
                '            mLet = "ˆ"
                '            aitiol = "¡ÍıÒ˘ÙÈÍ¸"
                '            akyr = 1
340         ElseIf RDB!paperCode = 1 Then
                 ' mLet = "A"
342              mLet = mID("AW", FORT, 1)
            
344             aitiol = "ƒÂÎÙÈÔ ¡ÔÛÙ.”ÂÈÒ·Ú ¬"
346             da = 1
        
348         ElseIf RDB!paperCode = 3 Then
                ' mLet = ""
350              mLet = mID("q", FORT, 1)
352             aitiol = "ƒÂÎÙÈÔ –·Ò·Î·‚ﬁÚ"
354             da = 1
356         ElseIf RDB!paperCode = 2 Then
                ' mLet = "Ù"
358              mLet = mID("ÙI", FORT, 1)
360             aitiol = "–ÈÛÙ˘ÙÈÍ¸ ‘ÈÏ.ƒ.–."
362             pistotiko = 1


            ElseIf RDB!paperCode = 12 Then
                ' mLet = "Ù"
             ' mLet = mID("ÙI", FORT, 1)
             ' aitiol = "–ÈÛÙ˘ÙÈÍ¸ ‘ÈÏ.ƒ.–."
             ' pistotiko = 1
                 mLet = mID("UO", FORT, 1)
                 aitiol = "¡ÍıÒ˘ÙÈÍÔ " + mSS
                 akyr = 1


            
            Else
364             MsgBox "–—œ”œ◊« ¡√ÕŸ”‘œ” ‘’–œ” –¡—¡”‘¡‘… œ’ ¡ ’—ŸÕ≈‘¡… « ≈…”¡√Ÿ√«"
366             GoTo rollover
            End If
        
368         R!ATIM = mLet + Format(RDB!paperaa, "000000")
        
370         r2.Open "select * from TIM WHERE ATIM='" + mLet + Format(RDB!paperaa, "000000") + "'", Gdb, adOpenDynamic, adLockOptimistic
372         knok = 0

374         For KN = 1 To j

376             If r2!id_num = id_num(KN) Then
                    'ok  ·ÌÁÍÂÈ ÏÂÛ· ÛÙ· ÛÁÏÂÒÈÌ· ÙÈÏÔÎÔ„È·
378                 knok = r2!id_num
                End If

            Next
        
380         If knok = 0 Then
382             MsgBox "Î·ËÔÚ ÛÙÁÌ ‰ÔÏﬁ Ù˘Ì ÂÈÛÂÒ˜ÔÏ›Ì˘Ì. ƒÂÌ ‚Ò›ËÁÍÂ ÙÔ master ÙÔı  " + R!ATIM
384             GoTo rollover
            End If

386         If r2.EOF Then
388             R("HME") = CDate(Format(Now, "MM/DD/YYYY"))
            Else
390             R("HME") = r2!hme
            End If

392         r2.Close
        
394         R!KODE = RDB!ProductCode ' Trim(mID(a, 27, 5))
396         r2.Open "select * from EID WHERE KOD='" + RDB!ProductCode + "'", Gdb, adOpenDynamic, adLockOptimistic

398         If r2.EOF Then
400             R!ONOMA = ""
402             R!FPA = 1
            Else
404             R!ONOMA = r2("ONO")
406             R!FPA = r2!FPA
            End If
        
408         r2.Close
        
410         R!POSO = Val(RDB!ProductQuantity) ' Val(mID(a, 138, 4))
412         R!PIS = Val(RDB!ProductQuantity)   ' Val(mID(a, 138, 4))
        
414         If akyr = 1 Then
416             R!POSO = -Val(RDB!ProductQuantity)  'Val(mID(a, 138, 4))
418             R!PIS = -Val(RDB!ProductQuantity)  'Val(mID(a, 138, 4))
            End If
        
420         R!XRE = 0
        
422         R!TIMM = Val(RDB!ProductPrice) ' Val(mID(a, 64, 6))
        
424         R!EKPT = Val(RDB!ProductDiscount) ' Val(mID(a, 79, 6))
426         R!kau_ajia = R!TIMM * R!POSO * (100 - R!EKPT) / 100
428         R!MIK_AJIA = R!kau_ajia * (100 + g_Fpa(R!FPA) / 100) ' GGET_NVALUE("SELECT TOP 1 TIMH FROM PINAKES WHERE TYPOS=1 AND AYJON=" + str(R!FPA))) / 100
430         R!apot = 1
        
432         R!id_num = knok ' TO ID_NUM TOY TIMOLOGIOY
        
434         R!EIDOS = "e"
436         R.Update
        
438         RDB.MoveNext
            ' Gdb.Execute "INSERT INTO EGGTIM (ATIM,KODE,POSO,TIMM,ONOMA,"
         
        Loop
    
440     R.Close
    
442     RDB.Close
    
        'Do While Not KIN1.Recordset.EOF  'Flag_Ekk
        'KIN1.Recordset.Edit
        '  KIN1.Recordset("Flag_Ekk") = 0
        ' KIN1.Recordset.Update
        ' KIN1.Recordset.MoveNext
        ' Loop
   
        Dim SQL6 As String
   
444     For KN = 1 To j
            'SQL6 = "select sum(CASE WHEN FPA=1 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA1,"
            'SQL6 = SQL6 + " sum(CASE WHEN FPA=2 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA2 ,"
            'SQL6 = SQL6 + " sum(CASE WHEN FPA=3 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA3 ,"
            'SQL6 = SQL6 + " sum(CASE WHEN FPA=4 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA4 ,"
            'SQL6 = SQL6 + " sum(CASE WHEN FPA=5 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA5 "
            'SQL6 = SQL6 + "  FROM EGGTIM WHERE ID_NUM=" + str(ID_NUM(kn))
            'R6.Open SQL6, Gdb, adOpenDynamic, adLockOptimistic
      
446         SQL6 = "select sum(CASE WHEN FPA=1 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA1,"
448         SQL6 = SQL6 + " sum(CASE WHEN FPA=2 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA2 ,"
450         SQL6 = SQL6 + " sum(CASE WHEN FPA=3 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA3 ,"
452         SQL6 = SQL6 + " sum(CASE WHEN FPA=4 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA4 ,"
      
454         SQL6 = SQL6 + " sum(CASE WHEN FPA=6 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA6 ,"
      
456         SQL6 = SQL6 + " sum(CASE WHEN FPA=5 THEN poso*timm*(100-ekpt)/100 ELSE 0 END) as FPA5 "
458         SQL6 = SQL6 + "  FROM EGGTIM WHERE ID_NUM=" + str(id_num(KN))
460         R6.Open UCase(SQL6), Gdb, adOpenDynamic, adLockOptimistic
   
462         r2.Open "SELECT * FROM TIM WHERE ID_NUM=" + str(id_num(KN)), Gdb, adOpenDynamic, adLockOptimistic
      
464         If Abs(NNUL(R6!FPA1) + NNUL(R6!fpa2) + NNUL(R6!fpa3) + NNUL(R6!FPA4) + NNUL(R6!FPA5) + NNUL(R6!FPA6) - (r2!aj6 + r2!aj1 + r2!aj2 + r2!aj3 + r2!aj4 + r2!aj5)) < 0.02 Then
   
466             Gdb.Execute "UPDATE  TIM SET AJ2=" + Replace(str(Round(NNUL(R6!fpa2), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
468             Gdb.Execute "UPDATE  TIM SET AJ3=" + Replace(str(Round(NNUL(R6!fpa3), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
470             Gdb.Execute "UPDATE  TIM SET AJ4=" + Replace(str(Round(NNUL(R6!FPA4), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
          
472             Gdb.Execute "UPDATE  TIM SET AJ6=" + Replace(str(Round(NNUL(R6!FPA6), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
474             Gdb.Execute "UPDATE  TIM SET AJ1=" + Replace(str(Round(NNUL(R6!FPA1), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
          
476             Gdb.Execute "UPDATE  TIM SET AJ5=" + Replace(str(Round(NNUL(R6!FPA5), 2)), ",", ".") + " WHERE ID_NUM=" + str(id_num(KN))
                'Gdb.Execute "UPDATE  TIM SET AJ1=AJ1-AJ2-AJ3-AJ4-AJ5 WHERE ID_NUM=" + str(ID_NUM(KN))
      
            End If

478         R6.Close
480         r2.Close
      
        Next
    
482     Gdb.CommitTrans
484     MsgBox "ok"
    
    '    create_sygk_epis MAX_ID_NUM, MAX_ID
    
        Exit Sub
    
rollover:
486     Gdb.RollbackTrans
488     MsgBox "‰ÂÌ ·ÔËÁÍÂıÙÁÍ·Ì"
    
        'PAR2_NUMBER   PAR2_APO_CODE   PAR2_APO_PERI                  PAR2_POSO1  PAR2_POSO2  PAR2_TIMH   PAR2_EKP    PAR2_TIMH_EFK     PAR2_KEY
        '  ‘…      13  0104026         MARS KING SIZE –¡√Ÿ‘œ 24 ‘≈/ ¬  1                       1.4200         0           0              ‘…¬13

        '<EhFooter>
        Exit Sub

APOKARAM_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ADOMERCNEW.bohu5.APOKARAM " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub Command15_Click()
   EISROON
End Sub



Public Sub EISROON()

   
   
   
   

        'EISROON================================================================
        '<EhHeader>
        ' On Error GoTo UpdateTIMwithMARK_Err
        '</EhHeader>

        '<?xml version="1.0" encoding="utf-8"?>
        '<Country xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
        '<city>
        '<index>1</index>
        '<StreetName>Street name</StreetName>
        '<StreetNumber>1050</StreetNumber>
        '<status>Success</status>
        '</city>
        '</Country>

        Dim Element As MSXML2.IXMLDOMElement

        Dim FAK     As String

        FAK = "c:\EXPORT\"

        'Print (Dir(App.Path + "\*.TXT"))
        'ADOMERCNEW Statistics.txt
        'MERCPATH.txt
        '12
        Dim ARX As String

        If Len(Dir("c:\EXPORT\*.XML")) > 0 Then
            ' –…¡”≈ ‘œ –—Ÿ‘œ ¡—◊≈…œ
            ARX = Dir("c:\EXPORT\*.XML")

            '÷œ—‘Ÿ”≈ ‘…” Ã≈‘¡¬À«‘≈”
   
            With New MSXML2.DOMDocument
                .Load FAK + ARX '"sample.xml" 'Local file name here, could also be HTTP or HTTPS URL.

                'Note that despite the name, the getElementsByTagName() method returns a
                'node list and not an elements list.  So we must query for the element
                'interface of the node in order to call getElementsByTagName() as we
                'descend the tree:
                ' Set Element = .documentElement.getElementsByTagName("city")(0)
                ' MsgBox Element.getElementsByTagName("StreetNumber")(0).Text
                'Set Element = .documentElement.getElementsByTagName("Product")(0)
                Dim MoneyEuClean As String: MoneyEuClean = .documentElement.getElementsByTagName("MoneyEuClean")(0).Text

                Dim TransSeqNum  As String: TransSeqNum = .documentElement.getElementsByTagName("TransSeqNum")(0).Text

                Dim GradeID      As String: GradeID = .documentElement.getElementsByTagName("GradeID")(0).Text

                Dim Product      As String: Product = .documentElement.getElementsByTagName("Product")(0).Text

                Dim VolumeLt     As String: VolumeLt = .documentElement.getElementsByTagName("VolumeLt")(0).Text

                Dim MoneyEu      As String: MoneyEu = .documentElement.getElementsByTagName("MoneyEu")(0).Text

                Dim MoneyEuVat   As String: MoneyEuVat = .documentElement.getElementsByTagName("MoneyEuVat")(0).Text

                Dim VatPerc      As String: VatPerc = .documentElement.getElementsByTagName("VatPerc")(0).Text

                Dim FinTime      As String: FinTime = .documentElement.getElementsByTagName("FinTime")(0).Text

                Dim Price        As String: Price = .documentElement.getElementsByTagName("Price")(0).Text

                Dim Series       As String: Series = .documentElement.getElementsByTagName("Series")(0).Text

            End With

100         Gdb.BeginTrans

140         Dim EXTIM As String: EXTIM = "INSERT INTO TIM (B_N1,EIDOS,KPE,HME,TRP, ATIM,ART,AJI,EKPT,EIDPAR, FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,   TYP, AJ1,AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9, EKPT1,EKPT2,EKPT3,EKPT4,EKPT5 ,KLEIDI) "

            EXTIM = EXTIM + "VALUES(1,'e','1111','" + Left(FinTime, 10) + "','1.Ã≈‘—',L'" + TransSeqNum + "',''," + MoneyEu + ",0,'1',0,0," + MoneyEuVat + ",0,0,0,0,0,0,0,0," + MoneyEuClean + ",0,0,0,0,0,0,0,0,0,0,0,'" + "L" + TransSeqNum + "') "

            'Gdb.Execute EXTIM

            'EXTIM = "INSERT INTO TIM (B_N1,EIDOS,KPE,HME,TRP, ATIM,ART,AJI,EKPT,EIDPAR, FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,   TYP, AJ1,AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9, EKPT1,EKPT2,EKPT3,EKPT4,EKPT5 ,KLEIDI) "
            'EXTIM = EXTIM + "VALUES(1,'e','1111','" + Left(FinTime, 10) + "','1.Ã≈‘—','L" + TransSeqNum + "',''," + MoneyEu + ",0,'1',0,0," + MoneyEuVat + ",0,0,0,0,0,0,0,0," + MoneyEuClean + ",0,0,0,0,0,0,0,0,0,0,0,'" + "L" + TransSeqNum + "') "
            'Gdb.Execute EXTIM
            EXTIM = "INSERT INTO TIM (B_N1,EIDOS,KPE,HME,TRP, ATIM,ART,AJI,EKPT,EIDPAR, FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,   TYP, AJ1,AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9, EKPT1,EKPT2,EKPT3,EKPT4,EKPT5 ,KLEIDI) "
            EXTIM = EXTIM + "VALUES(1,'e','1111','" + Left(FinTime, 10) + "','1.Ã≈‘—','L" + TransSeqNum + "',''," + MoneyEu + ",0,'1',0," + MoneyEuVat + ",0,0,0,0,0,0,0,0," + MoneyEuClean + ",0,0,0,0,0,0,0,0,0,0,0,0,'" + "L" + TransSeqNum + "') "
            Gdb.Execute EXTIM

            Dim id_num As String

            id_num = GGET_CVALUE("select str(max(ID_NUM)) from TIM")

144         Dim EXEGGT As String: EXEGGT = "INSERT INTO EGGTIM (FPA,PELKOD,ONOMA,ID_NUM,ATIM,HME,KODE,POSO,TIMM,APOT,EIDOS,XRE,PIS,KAU_AJIA,MIK_AJIA,EKPT)"

            If Len(id_num) = 0 Then id_num = "-" + TransSeqNum

            EXEGGT = EXEGGT + "VALUES(2,'1111','" + Product + "'," + id_num + ",'L" + TransSeqNum + "','" + Left(FinTime, 10) + "','" + GradeID + "'," + VolumeLt + "," + Price + ",1,'e',0," + VolumeLt + "," + MoneyEuClean + "," + MoneyEu + ",0)"
146         Gdb.Execute EXEGGT

150         Gdb.CommitTrans

            Dim PROOR As String

            PROOR = "C:\EXPORTBAK\" + Format(Now, "yyyy-MM-dd-HH-mm") + "OK" + ARX
            FileCopy FAK + ARX, PROOR
  
            ' ·Ì ÙÔ ·ÌÙÂ„Ò·¯Â Û‚ÁÛÙÔ
            If Len(Dir(PROOR)) > 0 Then
                Kill FAK + ARX
            End If
            
            Dim mStelno As Integer: mStelno = 1
            If mStelno = 1 Then
                gisOnlineMydata = 1
                Par7MyData.where.Caption = " where ID_NUM=" + str(id_num) + "  " ' M_idnum
                Par7MyData.Visible = False
                     
'                If f_parochos = 1 Then
'                    Par7MyData.loadpar7
'
'                    Dim ubl As String: ubl = GGET_CVALUE("SELECT left(ISNULL(UBL,'000'),3) AS UBL2 FROM PARASTAT WHERE EIDOS='" & Left(f_matim, 1) & "'")
'
'                    'e?? paroxo mono sthn xondriki kai to parastattiko anti na exei 380 exei "mydata"
'                    If ubl = "000" Then 'mydata
'                        Par7MyData.ToXMLsub 0
'                    Else
'                        Par7MyData.ToJason 0
'                    End If
'
'                    'Par7MyData.ToJason 0
'                Else
                     
                    Par7MyData.ToXMLsub 0, ""
'                End If
                     
            End If

        Else ' ƒ≈Õ ≈–≈”≈ ¡—◊≈…œ

        End If

        'id_num = GGET_CVALUE("select str(max(ID_NUM)) from TIM")

        '<?xml version="1.0" encoding="windows-1253"?>
        '<FuelSalesExport xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
        '*  <TransSeqNum>586543</TransSeqNum>  ar.apodixis
        '  *<GradeID>10</GradeID>   /kaysimo kodikos
        ' * <Product>UNL95 ???????</Product>  perigr kaysimoy
        '*  <VolumeLt>11.570000</VolumeLt>   litra
        '*  <MoneyEu>20.000000</MoneyEu>  euro axia
        '*  <MoneyEuClean>16.130000</MoneyEuClean>  kau.axia
        '* <MoneyEuVat>3.870000</MoneyEuVat>   fpa
        '*  <VatPerc>24.00</VatPerc>  %fpa
        '  <FpID>1</FpID>   apo 1 h antlia
        '  <NozzleID>1</NozzleID>  1o nozzle
        '*  <FinTime>2024-09-23T12:28:28</FinTime> hmer ora
        ' * <Price>1.729000</Price>  timh litroy me fpa
        '  <VatNumber />
        '  <Tank>2</Tank>  2 dexamenh
        '  <LicensePlate />
        ' * <Series>EE</Series>  seira par/koy
        '</FuelSalesExport>

End Sub

Private Sub IMPORTSERV_Click()
 ' Dim GDBDRAMA As New ADODB.Connection
 ' GDBDRAMA.Open "DSN=doytsios;uid=sa;pwd=p@ssw0rd"
 F_DOUTSIOSREM = Trim(F_DOUTSIOSREM)
 
 If InStr(UCase(F_DOUTSIOSREM), "DOUTSIOS") = 0 Then
     
     MsgBox "À¡»œ”  254"
     Exit Sub
 End If
 If Len(Trim(F_DOUTSIOSREM)) < 10 Then
     MsgBox "ƒ≈Õ ≈◊≈‘≈ œ—…”≈… ”≈—¬≈—"
     Exit Sub
 End If
 
 
 
 ' F_DOUTSIOSPARAS = Trim(F_DOUTSIOSPARAS)
 On Error GoTo SFALMA1
  Dim R As New ADODB.Recordset
  R.Open "select top " + ARPAR.Text + " ATIM,HME,ID_NUM  from " + F_DOUTSIOSREM + ".DOUTSIOSQ.dbo.TIM WHERE LEFT(ATIM,1)='" + F_LINKED_PARAST + "' order by ID_NUM DESC", Gdb, adOpenDynamic, adLockOptimistic
  Me.Caption = "OK OPEN " + F_DOUTSIOSREM + ".DOUTSIOSQ.dbo.TIM"
  
  Dim COLTIM As String
  On Error GoTo SFALMA
  Dim N As Integer
  Dim MATIM As String
  Dim mID_NUM As Long
  Dim FEGGTIM As String
  FEGGTIM = "EIDOS,ATIM,POSO,MONA,TIMM,KERDOS,KODE,HME,ERGO,FPA,PROOD,PROOD_AJ,EKPT,KAU_AJIA,MIK_AJIA,ONOMA,MIKTA,KOLA,PELKOD,PROELEYSH,XRE,PIS,APOT,ATIM2,FCURRENCY,EKPT2,MIKTAKILA"
  COLTIM = "KLEIDI,KPE,HME,TRP,ATIM,ART,AJI,EIDOS,METAF,EKPT,EIDPAR,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,TYP,AJ1, AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9,EKPT1,EKPT2,EKPT3,EKPT4,EKPT5,HME_KATAX,KERDOS,SKOPOS,SXETIKO,PARAT,ELGA,SYNPOS,SKOPOS2,FORTOSH,PROOR,AYTOK,B_C1,B_C2,B_N1,B_N2,KR1,KR2,ATIM2"
  Dim PROTO As Integer
  PROTO = 0
  Dim eis As Integer
  eis = 0
  Dim r0 As New ADODB.Recordset
  Dim DOKTIM As String
  Dim TEL_ID_NUM As Long
  Do While Not R.EOF
  
   
   If PROTO = 0 Then
        Me.Caption = "PROTO=0 OK "
       Gdb.Execute "delete from BARCODES"
           Gdb.Execute "  INSERT INTO BARCODES (KOD,ERG) SELECT KOD,ERG FROM " + F_DOUTSIOSREM + ".DOUTSIOSQ.dbo.BARCODES"
           Me.Caption = "BARCODES OK"
          MILSEC 1000
           Gdb.Execute "delete from EID"
           Dim TROPOS As Integer
           TROPOS = 2
           
           If TROPOS = 1 Then
           
                Gdb.Execute "SELECT KOD,ONO,FPA,NUM1,LTI,MON into DOKEIDD FROM " + F_DOUTSIOSREM + ".DOUTSIOSQ.dbo.EID "
                Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,NUM1,LTI,MON) SELECT KOD,ONO,FPA,NUM1,LTI,MON FROM DOKEIDD"
                
                Gdb.Execute "DROP TABLE DOKEIDD"
                
           ElseIf TROPOS = 2 Then
                 Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,NUM1,LTI,MON) SELECT KOD,ONO,FPA,NUM1,LTI,MON FROM " + F_DOUTSIOSREM + ".DOUTSIOSQ.dbo.EID "
           ElseIf TROPOS = 3 Then
               r0.Open "SELECT ID  from " + F_DOUTSIOSREM + ".DOUTSIOSQ.dbo.EID ", Gdb, adOpenDynamic, adLockOptimistic
           
               N = 0:
               Do While Not r0.EOF
                   N = N + 1
                  Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,NUM1,LTI,MON) SELECT KOD,ONO,FPA,NUM1,LTI,MON FROM " + F_DOUTSIOSREM + ".DOUTSIOSQ.dbo.EID where ID=" + str(r0(0))
                  Me.Caption = str(N)
                  If N Mod 20 = 0 Then DoEvents
                  r0.MoveNext
               Loop
               r0.Close
           End If
           
           
           
           
          MILSEC 1000
       Gdb.Execute "UPDATE EID SET LTI5=NUM1"
       Gdb.Execute "UPDATE EID SET FPA=3 WHERE FPA=7"
          Me.Caption = "EIƒH OK"
   
   End If
   
  
    MATIM = R!ATIM
    mID_NUM = R!id_num  ' TO ID_NUM TOY SERVER
    PROTO = PROTO + 1
    
    '‰ÂÌ ı·Ò˜ÂÈ ÙÔ ·Ò·ÛÙ·ÙÈÍÔ Â‰˘ ·Ò· ÙÔ ˆÔÒÙ˘Ì˘
     If GGET_NVALUE("select count(*) from TIM WHERE ATIM='" + MATIM + "'") = 0 Then
        eis = eis + 1
          Me.Caption = "PROTO=1 OK "
        
        Gdb.Execute "INSERT INTO TIM (" + COLTIM + ") SELECT " + COLTIM + " FROM " + F_DOUTSIOSREM + ".DOUTSIOSQ.dbo.TIM  WHERE  ID_NUM=" + str(mID_NUM), N
        Me.Caption = "INSERT INTO TIM  OK " + str(eis)
        
        If N = 1 Then
            TEL_ID_NUM = GGET_NVALUE("select MAX(ID_NUM) from TIM ")
            Gdb.Execute "INSERT INTO EGGTIM (" + FEGGTIM + ",ID_NUM ) SELECT " + FEGGTIM + "," + str(TEL_ID_NUM) + " FROM " + F_DOUTSIOSREM + ".DOUTSIOSQ.dbo.EGGTIM  WHERE ID_NUM=" + str(mID_NUM)
       End If
  
    End If
    
    R.MoveNext
  Loop
R.Close
'GDBDRAMA.Close

MsgBox "OK " + str(eis)


'INSERT INTO TIM (ATIM,HME,KPE,KLEIDI) SELECT ATIM,HME,KPE,KLEIDI FROM "+F_DOUTSIOSREM+".DOUTSIOSQ.dbo.TIM  WHERE ATIM='v000001'



Exit Sub
SFALMA1:

MsgBox "ƒ≈Õ œÀœ À«—Ÿ»« ≈ Àœ√Ÿ ·‰ıÌ·ÏÈ·Ú ÛıÌ‰ÂÛÁÚ"
Exit Sub


SFALMA:

MsgBox "ƒ≈Õ œÀœ À«—Ÿ»« ≈ Àœ√Ÿ  ¡»’”‘≈—«”«” ƒ… ‘’œ’"

End Sub

Private Sub Command17_Click()

End Sub

Private Sub dtpicker1_LostFocus()
 If Year(DTPicker1.Value) <> Year(gLhjh) Then
      '  MsgBox "À¡»œ” ≈‘œ”"
      '  dtpicker1.SetFocus
   End If
End Sub

Private Sub EXP_SMARTWARE_Click()

    '  a = toascii("C:\MERCVB\FPERP.TXT", "C:\TEMP\PEL.TXT")
    Dim R As New ADODB.Recordset

    Dim PIN(100, 2)

    ' FORTONO TON PINAKA 20 ME TIS ANTISTOIXISEIS TON PEDION
    '  ¡… ≈À≈√◊Ÿ ¡Õ ≈◊≈… ƒ«ÀŸ»≈… œ  Ÿƒ… œ”

    Dim COUNTER As Long: COUNTER = 0

    R.Open "SELECT *FROM PINAKES WHERE TYPOS=101 AND TIMH>0 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        COUNTER = COUNTER + 1
        PIN(COUNTER, 1) = R("PERIGRAFH")    '–≈ƒ…œ ‘œ’  ’—…œ’ ¡—◊≈…œ’ (SQLSERVER)
        PIN(COUNTER, 2) = R("TIMH")    ' ¡—…»Ãœ” ”‘«À«” ‘œ’ EXCEL
        R.MoveNext
    Loop

    R.Close
 
    Dim a As String, k As Integer

    Dim F As Integer

    F = FreeFile
    Open "C:\TEMP\PEL.ASC" For Output As #F
    R.Open "SELECT *FROM PEL WHERE EIDOS='e' order by EPO", Gdb, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        a = ""

        For k = 1 To COUNTER

            If InStr(PIN(k, 1), "@") = 0 Then '  ¡»¡—œ –≈ƒ…œ
                If Len(PIN(k, 1)) > 0 Then

                    ' ·ÒÈËÏÁÙÈÍÔ
                    If R(PIN(k, 1)).Type = 7 Or R(PIN(k, 1)).Type = 5 Or R(PIN(k, 1)).Type = 131 Then     'IsNumeric(
                        a = a + Right(Space(100) + Format(IIf(IsNull(R(PIN(k, 1))), 0, R(PIN(k, 1))), String(PIN(k, 2), "#") + ".00"), PIN(k, 2))
                    Else
                        ' string
                        a = a + Left(IIf(IsNull(R(PIN(k, 1))), Space(100), R(PIN(k, 1))) + Space(100), PIN(k, 2))
                    End If
                End If

            Else
                a = a + mID(PIN(k, 1) + Space(100), 2, PIN(k, 2))
            End If

        Next

        Print #F, a
        'COUNTER = COUNTER + 1
        R.MoveNext
    Loop

    R.Close
 
    Close #F
 
    'EIDH
    COUNTER = 0
    R.Open "SELECT *FROM PINAKES WHERE TYPOS=102 AND TIMH>0 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        COUNTER = COUNTER + 1
        PIN(COUNTER, 1) = R("PERIGRAFH")    '–≈ƒ…œ ‘œ’  ’—…œ’ ¡—◊≈…œ’ (SQLSERVER)
        PIN(COUNTER, 2) = R("TIMH")    ' ¡—…»Ãœ” ”‘«À«” ‘œ’ EXCEL
        R.MoveNext
    Loop

    R.Close
 
    F = FreeFile
    Open "C:\TEMP\APO.ASC" For Output As #F
    R.Open "SELECT *FROM EID  order by ONO", Gdb, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        a = ""

        For k = 1 To COUNTER

            If InStr(PIN(k, 1), "@") = 0 Then '  ¡»¡—œ –≈ƒ…œ
                If Len(PIN(k, 1)) > 0 Then

                    ' ·ÒÈËÏÁÙÈÍÔ
                    If R(PIN(k, 1)).Type = 7 Or R(PIN(k, 1)).Type = 5 Or R(PIN(k, 1)).Type = 131 Then     'IsNumeric(
                        a = a + Right(Space(100) + Format(IIf(IsNull(R(PIN(k, 1))), 0, R(PIN(k, 1))), String(PIN(k, 2), "#") + ".0000"), PIN(k, 2))
                    Else
                        ' string
                        a = a + Left(IIf(IsNull(R(PIN(k, 1))), Space(100), R(PIN(k, 1))) + Space(100), PIN(k, 2))
                    End If
                End If

            Else
                a = a + mID(PIN(k, 1) + Space(100), 2, PIN(k, 2))
            End If

        Next

        Print #F, a
        'COUNTER = COUNTER + 1
        R.MoveNext
    Loop

    R.Close
 
    Close #F
 
    'TIMOKATALOGOS
    COUNTER = 0
    R.Open "SELECT *FROM PINAKES WHERE TYPOS=103 AND TIMH>0 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        COUNTER = COUNTER + 1
        PIN(COUNTER, 1) = R("PERIGRAFH")    '–≈ƒ…œ ‘œ’  ’—…œ’ ¡—◊≈…œ’ (SQLSERVER)
        PIN(COUNTER, 2) = R("TIMH")    ' ¡—…»Ãœ” ”‘«À«” ‘œ’ EXCEL
        R.MoveNext
    Loop

    R.Close
 
    F = FreeFile
    Open "C:\TEMP\TIMOK.ASC" For Output As #F
    R.Open "SELECT *FROM TIMOKAT  order by ONO", Gdb, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        a = ""

        For k = 1 To COUNTER

            If InStr(PIN(k, 1), "@") = 0 Then '  ¡»¡—œ –≈ƒ…œ
                If Len(PIN(k, 1)) > 0 Then

                    ' ·ÒÈËÏÁÙÈÍÔ
                    If R(PIN(k, 1)).Type = 7 Or R(PIN(k, 1)).Type = 5 Or R(PIN(k, 1)).Type = 131 Then     'IsNumeric(
                        a = a + Right(Space(100) + Format(IIf(IsNull(R(PIN(k, 1))), 0, R(PIN(k, 1))), String(PIN(k, 2), "#") + ".0000"), PIN(k, 2))
                    Else
                        ' string
                        a = a + Left(IIf(IsNull(R(PIN(k, 1))), Space(100), R(PIN(k, 1))) + Space(100), PIN(k, 2))
                    End If
                End If

            Else
                a = a + mID(PIN(k, 1) + Space(100), 2, PIN(k, 2))
            End If

        Next

        Print #F, a
        'COUNTER = COUNTER + 1
        R.MoveNext
    Loop

    R.Close
 
    Close #F
 
    MsgBox "OK"
 
End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        'File1.LinePrint Mid(mPEL, 1, 4) + Grid.TextMatrix(k, 6) + "*" + aritmisi.Caption + "=" + d1 + "+" + d2 + "$" + FormatDateTime(Now, vbGeneralDate) + "@" + PISTOSI.Text
        'TIMOL.TXT      (INPUT)
        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>

        ' 0004004*11=4.00+8.02$12/10/06 11:22:58 ??@?ISTOSH
        ' 0004005*11=5.00+9.16$12/10/06 11:22:59 ??@?ISTOSH
        ' 0004008*11=8.00+12.08$12/10/06 11:22:59 ??@?ISTOSH

        ' 1.PNI        (OUTPUT)   ”‘œ◊œ”
        '‘…900000117-11-06 ¡¬.012     1
        '¡Õ¡.23     TI900000117-11-0655     12,00      0,66      7,92
        '¡Õ¡.24     TI900000117-11-0655      1,00     11,70     11,70
        '¡Õ¡.25     TI900000117-11-0655      1,00     11,25     11,25

        '‘…900000217-11-06 ¡¬.012     1
        '¡Õ¡.23     TI900000217-11-0655     12,00      0,66      7,92
        '¡Õ¡.24     TI900000217-11-0655      1,00     11,70     11,70
        '¡Õ¡.25     TI900000217-11-0655      1,00     11,25     11,25

        'PREPEI NA MHN DEXETAI DIPLA
        'KAI NA GRAFEI TON TROPO PLIROMIS

        Dim m_ait

        Dim m_onoma

        Dim k, a

        Dim DB As Database

        Dim R  As New ADODB.Recordset

        Dim c, D, e, F, G, n1, n2

        Dim m_hme  ' = hme ' ‘≈À≈’‘¡…¡ «Ã≈—œÃ«Õ…¡

        Dim ATIM, KODPEL, KODEID, POS, TIMM, hme

        Dim CC

        On Error GoTo LATOS

100     CC = CurDir

110     CC = "c:\mercvb"
120     m_ait = ".."

        'Set db = OpenDatabase(CC, False, False, "dBase III;")

        Dim ArTim, M_PEL

130     ArTim = "ddddd"
140     Me.Caption = "OPENDBASE"

150     Me.Caption = "≈Õ¡—Œ«-¡Õœ…√Ã¡ ¡—◊≈…œ’ C:\TIMOL.TXT"
160     Open CC + "\timol.txt" For Input As #1
170     Me.Caption = 11

        Dim gm_fpa, m_E

        Dim M_POSO(5), syn_m_POSO

180     M_POSO(1) = 0
190     M_POSO(2) = 0
200     M_POSO(3) = 0
210     M_POSO(4) = 0
220     M_POSO(5) = 0
230     syn_m_POSO = 0
240     m_E = "0"

250     Me.Caption = "À«Œ«-¡Õœ…√Ã¡ ¡—◊≈…œ’ C:\TIMOL.TXT"

        Dim nn

260     nn = 0

        '--------------------------------------------------------------------------------------
270     Do While Not EOF(1)
280         nn = nn + 1
290         Line Input #1, a
300         Me.Caption = Format(nn, "####") + " " + a
310         List1.AddItem Format(nn, "####") + " " + a

320         KODPEL = Left(a, 4)    'R("shm1") ' Í˘‰.ÂÎ·ÙÁ
330         R.Open "select * from EID where KOD='" + Trim(mID(a, 5, 3)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

340         gm_fpa = R("fpa")
350         m_onoma = R("ono")

360         KODEID = R("KOD")    ' Í˘‰.ÂÈ‰ÔıÚ
370         R.Close

380         n1 = InStr(1, a, "*")
390         n2 = InStr(1, a, "=")
400         ATIM = mID(a, n1 + 1, n2 - n1 - 1)    'ariumos timologioy

410         Me.Caption = "111"

420         If F_TIM = 1 Then
430             ATIM = "T" + Format(Val(ATIM), "000000")
            Else

440             If InStr(1, Right(Trim(a), 4), "-") Then  ' TIMOLOGIO
450                 ATIM = "T" + Format(Val(ATIM), "000000")
                Else
460                 ATIM = "L" + Format(Val(ATIM), "000000")
                End If
            End If

470         Me.Caption = "22"

480         n1 = InStr(1, a, "=")
490         n2 = InStr(1, a, "+")
500         POS = Val(mID(a, n1 + 1, n2 - n1 - 1))    'ÔÛÔÙÁÙ·

510         n1 = InStr(1, a, "+")
520         n2 = InStr(1, a, "$")
530         TIMM = Val(mID(a, n1 + 1, n2 - n1 - 1))    'ÙÈÏﬁ ÏÔÌ‹‰·Ú

540         Me.Caption = "333"

550         n1 = InStr(1, a, "$")
560         n2 = InStr(n1, a, " ")
570         hme = Right(mID(a, n1 + 1, n2 - n1 - 1), 10)  'ÁÏÂÒÔÏÁÌﬂ·

580         G = CDate(hme)
590         hme = Format(G, "dd/mm/yyyy")

600         n1 = InStr(1, a, "@")
610         e = mID(a, n1 + 1, 1)    ' –…”‘Ÿ”« Ã≈‘—«‘œ…”

620         If e = "–" Then
630             e = "0"  '–…”‘Ÿ”«
            Else
640             e = "1"  ' Ã≈‘—«‘œ…”
            End If

650         Me.Caption = "4444"

660         If ArTim <> ATIM Then
670             syn_m_POSO = M_POSO(1) * 1.09 + M_POSO(2) * 1.19 + M_POSO(3) * 1.13 + M_POSO(5)

                ' Print #2, "       "
                'Print #2, to437(atim + hme + KODPEL) + "     " + e  m_pel
680             If ArTim <> "ddddd" Then
690                 G = "INSERT INTO TIM " & "(EIDOS,HME,ATIM,KPE,AJ1,AJ2,AJ3,AJ4,AJ5,AJI)" & " VALUES ('e'," & "'" + Format(m_hme, "MM/DD/YYYY") + "'," & "'" + ArTim + "'," & "'" + M_PEL + "'," & Replace(Format(M_POSO(1), "0000000.00"), ",", ".") + "," & Replace(Format(M_POSO(2), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(3), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(3), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(5), "###0.00"), ",", ".") + "," & Replace(Format(syn_m_POSO, "0000000.00"), ",", ".") + ");"

700                 Gdb.Execute G

710                 Gdb.Execute "INSERT INTO EGG " & "(HME,ATIM,KOD,EIDOS,APA,XRE,XREOSI,AIT)" & " VALUES (" & "'" + Format(hme, "MM/DD/YYYY") + "'," & "'" + ArTim + "'," & "'" + M_PEL + "'," & "'e'," & "'" + mID(ArTim, 2, 6) + "'," & Replace(Format(syn_m_POSO, "0000000.00"), ",", ".") + "," & Replace(Format(syn_m_POSO, "###0.00"), ",", ".") + ",'" & m_ait + "');"

720                 If m_E <> "0" Then
                        '·Ò˜ÂÈÔ ÏÂÙÒÁÙÔÈÛ 3.pni

730                     Gdb.Execute "INSERT INTO EGG " & "(HME,ATIM,KOD,EIDOS,APA,XRE,PISTOSI,AIT)" & " VALUES (" & "'" + Format(hme, "MM/DD/YYYY") + "'," & "'" + ArTim + "'," & "'" + M_PEL + "'," & "'e'," & "'" + mID(ArTim, 2, 6) + "'," & Replace(Format(syn_m_POSO, "0000000.00"), ",", ".") + "," & Replace(Format(syn_m_POSO, "###0.00"), ",", ".") + "," & "'Ã≈‘—«‘¡');"

                    End If
                End If

740             M_POSO(1) = 0
750             M_POSO(2) = 0
760             M_POSO(3) = 0
770             M_POSO(4) = 0
780             M_POSO(5) = 0
790             ArTim = ATIM
            End If

800         G = "INSERT INTO EGGTIM " & "(EIDOS,KODE,HME,ATIM,PELKOD,POSO,TIMM,PIS,ONOMA)" & " VALUES ('e','" + KODEID + "'," & "'" + Format(hme, "MM/DD/YYYY") + "'," & "'" + ATIM + "'," & "'" + KODPEL + "'," & Replace(Format(POS, "0000000.00"), ",", ".") + "," & Replace(Format(Round(TIMM, 2), "###0.00"), ",", ".") + "," & Replace(Format(POS, "0000000.00"), ",", ".") + ",'" + m_onoma + "');"

810         Gdb.Execute G

820         If gm_fpa = 1 Then
830             M_POSO(1) = M_POSO(1) + POS * TIMM
840         ElseIf gm_fpa = 2 Then
850             M_POSO(2) = M_POSO(2) + POS * TIMM
860         ElseIf gm_fpa = 3 Then
870             M_POSO(3) = M_POSO(3) + POS * TIMM
880         ElseIf gm_fpa = 5 Then
890             M_POSO(5) = M_POSO(5) + POS * TIMM    '      '      0 %
            End If

900         If InStr(1, Right(Trim(a), 4), "-") Or F_TIM = 1 Then    ' TIMOLOGIO
910             m_ait = "‘…ÃœÀœ√…œ –ŸÀ«”«”"

920         ElseIf InStr(1, Right(Trim(a), 4), "IKH") Then  ' TIMOLOGIO
930             m_ait = "¡–œƒ.À…¡Õ… «”"
            Else
940             m_ait = "ƒ≈À‘…œ ¡–œ”‘œÀ«”"
            End If

950         M_PEL = KODPEL
960         m_hme = hme    ' ‘≈À≈’‘¡…¡ «Ã≈—œÃ«Õ…¡
970         m_E = e    ' ÈÛÙ˘ÛÁ ÙÂÎÂıÙ·ﬂ·
        Loop

        '-------------------------------------------------------------------------------------------------
980     Close #1

990     syn_m_POSO = M_POSO(1) * 1.09 + M_POSO(2) * 1.19 + M_POSO(3) * 1.13 + M_POSO(5)
1000    G = "INSERT INTO TIM " & "(EIDOS,HME,ATIM,KPE,AJ1,AJ2,AJ3,AJ4,AJ5,AJI)" & " VALUES ('e'," & "'" + Format(m_hme, "MM/DD/YYYY") + "'," & "'" + ATIM + "'," & "'" + M_PEL + "'," & Replace(Format(M_POSO(1), "0000000.00"), ",", ".") + "," & Replace(Format(M_POSO(2), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(3), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(4), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(5), "###0.00"), ",", ".") + "," & Replace(Format(syn_m_POSO, "0000000.00"), ",", ".") + ");"
1010    Gdb.Execute G

1020    If m_E <> "0" Then
            '·Ò˜ÂÈÔ ÏÂÙÒÁÙÔÈÛ 3.pni

1030        Gdb.Execute "INSERT INTO EGG " & "(HME,ATIM,KOD,EIDOS,APA,XRE,PISTOSI,AIT)" & " VALUES (" & "'" + Format(hme, "MM/DD/YYYY") + "'," & "'" + ATIM + "'," & "'" + M_PEL + "'," & "'e'," & "'" + mID(ATIM, 2, 6) + "'," & Replace(Format(syn_m_POSO, "0000000.00"), ",", ".") + "," & Replace(Format(syn_m_POSO, "###0.00"), ",", ".") + "," & "'Ã≈‘—«‘¡');"

        End If

1040    MsgBox "œ  ÂÌÁÏÂÒ˛ËÁÍÂ"

        Exit Sub

LATOS:
1050    MsgBox "‰È·ÍÔÁ " + Chr(13) + Err.Description

        On Error Resume Next

1060    Close #1

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
'Function to437(string_ As String) As String
'    Dim a$, k As Integer, S As String, t As Integer, s928 As String, s437 As String
'    'metatrepei eggrafo apo 437->928
'    s928 = "¡¬√ƒ≈∆«»… ÀÃÕŒœ–—”‘’÷◊ÿŸ-·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÛÙıˆ˜¯˘-Ú‹›ﬁﬂ¸˝˛"
'    s437 = "ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñó-òôöõúùûü†°¢£§•¶ß®©´¨≠ÆØ‡-™·‚„ÂÊÁÈ"    ' saehioyv
'    's437 = "ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñó-òôöõúùûü†°¢£§•¶ß®©´¨≠ÆØ‡™" '·‚„ÂÊÁÈ"
'    'ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñóòôöõúùûü†°¢£§•¶ß®©´¨≠ÆØ‡
'
'    a$ = string_
'    '                                                        saehioyv
'    'GoTo 11
'    'Open Text2.Text For Output As #2
'    'Open Text1.Text For Input As #1
'    'Do While Not EOF(1)
'    '  Line Input #1, a$
'    For k = 1 To Len(a$)
'        S = mID(a$, k, 1)
'        t = InStr(s928, S)
'        If t > 0 Then
'            Mid$(a$, k, 1) = mID$(s437, t, 1)
'        End If
'    Next
'11
'    to437 = a$
'
'
'
'
'End Function

Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>

        '=========================================
        '‰ÁÏÈÔıÒ„ﬂ· eggtim Í·È ‘…Ã
        '=========================================

        Dim k, L, m, N As Integer

        Dim DB As Database, re As New ADODB.Recordset, rt As New ADODB.Recordset

        Dim R  As New ADODB.Recordset

        Dim FEGG(10)

        Dim m_tim

100     POSO.Text = Replace(POSO.Text, ",", ".")

110     m_tim = Format(Val(MATIM), "000000")

120     re.Open "SELECT  *FROM PEL WHERE EIDOS='r' and AFM='" + Left(AFM.Text, 9) + "'", Gdb, adOpenDynamic, adLockOptimistic

        Dim M_KOD

130     If re.EOF Then
140         MsgBox "ƒ≈Õ ’–¡—◊≈… œ –—œÃ«»≈’‘«”"

            Exit Sub

        End If

150     M_KOD = re("KOD")
160     re.Close

        Dim m_dat

170     m_dat = DTPicker1.Value

        'ans = vbNo
180     re.Open "select * from TIM where  HME>='" + Format(m_dat, "mm/dd/yyyy") + "' AND HME<'" + Format(DateAdd("D", 1, m_dat), "mm/dd/yyyy") + "'  and ATIM='Œ" + m_tim + "'", Gdb, adOpenDynamic, adLockOptimistic

190     If Not re.EOF Then
200         MsgBox "’‹Ò˜ÂÈ ﬁ‰Á  ·Ò·ÛÙ·ÙÈÍ¸ ÏÂ ﬂ‰È· ÁÏÂÒÔÏÁÌﬂ· Í·È ﬂ‰ÈÔ ·ÒÈËÏ¸. "

            Exit Sub

        End If

210     re.Close

220     re.Open "SELECT  *FROM EGGTIM", Gdb, adOpenDynamic, adLockOptimistic
230     rt.Open "SELECT  *FROM TIM", Gdb, adOpenDynamic, adLockOptimistic

240     re.AddNew
250     re("hme") = DTPicker1.Value
260     re("atim") = "Œ" + m_tim

270     re("poso") = 1
280     re("timm") = Val(POSO.Text) / 1.19
290     re("kode") = "9-1"
300     re("onoma") = "Ã≈‘¡÷œ—… ¡ 19%"
310     re("mona") = "‘≈Ã"
320     re("kau_ajia") = Val(POSO.Text) / 1.19
330     re("MIK_AJIA") = Val(POSO.Text)

340     re("fpa") = 2
350     re("ekpt") = 0
360     re("pelkod") = M_KOD
370     re("eidos") = "r"
380     re("APOT") = 1
390     re("XRE") = 1
400     re.Update

410     rt.AddNew
420     rt("hme") = DTPicker1.Value
430     rt("atim") = "Œ" + m_tim

440     rt("aj2") = Val(POSO.Text) / 1.19
450     rt("aj1") = 0

460     rt("aj3") = 0
470     rt("aj4") = 0
480     rt("aj5") = 0

490     rt("fpa2") = Val(POSO.Text) - Val(POSO.Text) / 1.19
500     rt("kpe") = M_KOD
510     rt("aji") = Val(POSO.Text)
520     rt("eidos") = "r"
530     rt("EIDPAR") = ""
540     rt("B_N1") = 1
550     rt.Update

560     R.Open "SELECT  *FROM EGG", Gdb, adOpenDynamic, adLockOptimistic
570     R.AddNew
580     R("hme") = DTPicker1.Value
590     R("apa") = Val(m_tim)
600     R("atim") = "Œ" + m_tim
610     R("ait") = "ÏÂÙ·ˆÔÒÈÍ·"
620     R("XREOSI") = 0: R("PISTOSI") = Val(POSO.Text)    ' ¡–œ÷’√« ÕULL
630     R("eidos") = "r"
640     R("kod") = M_KOD
650     R.Update

660     R.AddNew
670     R("hme") = DTPicker1.Value
680     R("apa") = Val(m_tim)
690     R("atim") = "Œ" + m_tim
700     R("ait") = "ÏÂÙÒÁÙ·"
710     R("XREOSI") = Val(POSO.Text): R("PISTOSI") = 0    ' ¡–œ÷’√« ÕULL
720     R("eidos") = "r"
730     R("kod") = M_KOD
740     R.Update

750     R.Close

760     rt.Close
770     re.Close

780     Me.Caption = m_tim + " Í·ÙÂ˜˘ÒﬁËÁ"
790     DTPicker1.SetFocus

        Exit Sub

MHNYMA:
800     HandleError "Par1:enhm_eggtim"

810     Resume Next

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>

        'kataxorhsh   Z

        '=========================================
        '‰ÁÏÈÔıÒ„ﬂ· eggtim Í·È ‘…Ã
        '=========================================

        Dim k, L, m, N As Integer

        Dim DB As Database, re As New ADODB.Recordset, rt As New ADODB.Recordset


Dim MID_TIM As Long
Dim m_ideggtim As Long



        Dim R  As New ADODB.Recordset

        Dim FEGG(10)

        Dim m_tim

100     POSO.Text = Replace(POSO.Text, ",", ".")

110     m_tim = Format(Val(MATIM), "000000")

        're.Open "SELECT  *FROM PEL WHERE EIDOS='r' and AFM='" + Left(AFM.Text, 9) + "'", _
        'Gdb, adOpenDynamic, adLockOptimistic

        Dim M_KOD

        'If re.EOF Then
        '  MsgBox "ƒ≈Õ ’–¡—◊≈… œ –—œÃ«»≈’‘«”"
        '  Exit Sub
        'End If
120     M_KOD = "1111"  ' re("KOD")
        're.Close

        Dim m_dat

130     m_dat = DTPicker1.Value

        'ans = vbNo
140     re.Open "select * from TIM where  HME>='" + Format(m_dat, "mm/dd/yyyy") + "' AND HME<'" + Format(DateAdd("D", 1, m_dat), "mm/dd/yyyy") + "'  and ATIM='ˆ" + m_tim + "'", Gdb, adOpenDynamic, adLockOptimistic

150     If Not re.EOF Then
160         MsgBox "’‹Ò˜ÂÈ ﬁ‰Á  ·Ò·ÛÙ·ÙÈÍ¸ ÏÂ ﬂ‰È· ÁÏÂÒÔÏÁÌﬂ· Í·È ﬂ‰ÈÔ ·ÒÈËÏ¸. "

            Exit Sub

        End If

170     re.Close

        Dim sql As String

        Dim SQL2


Dim KAU6, FPA6, KAU4, FPA4 As String

KAU4 = gDot(PERIB.Text)
FPA4 = gDot(str(gVal(PERIB.Text) * Val(FPA.Text) / 100))




KAU6 = str((gVal(POSO.Text) - Val(KAU4) - Val(FPA4)) / ((100 + Val(FPA.Text)) / 100))




FPA6 = gDot(gVal(POSO.Text) - Val(KAU6) - Val(KAU4) - Val(FPA4))




      SQL2 = "SELECT ID_NUM FROM TIM WHERE ID_NUM = @@IDENTITY"

        sql = "INSERT INTO TIM(KLEIDI,ATIM,KPE,B_N1,EIDPAR,EIDOS,AJ1,AJ2,AJ3,AJ4,AJ5,FPA1,FPA2,FPA3,FPA4,AJ6,FPA6,AJ7,FPA7) VALUES "
        sql = sql + "('" + "ˆ" + m_tim + "','" + "ˆ" + m_tim + "','" + M_KOD + "',1,'','e',0,0,0,0,0,0,0,0,0,0,0,0,0);"

        Gdb.Execute sql

        re.Open SQL2, Gdb, adOpenDynamic, adLockOptimistic
        MID_TIM = re(0)
        re.Close

        Gdb.Execute "UPDATE TIM SET HME='" + Format(DTPicker1.Value, "mm/DD/YYYY") + "' where ID_NUM=" + str(MID_TIM)
        '         Gdb.Execute "UPDATE TIM SET ATIM='" + "ˆ" + m_tim + "'" + " where ID_NUM=" + str(mID)

        '        Gdb.Execute "UPDATE TIM SET TIMM=" + str(Val(POSO.Text) / ((100 + Val(FPA.Text)) / 100)) + " where ID_NUM=" + str(mID)
        '        Gdb.Execute "UPDATE TIM SET ATIM='" + "ˆ" + m_tim + "'" + " where ID_NUM=" + str(mID)
        Gdb.Execute "UPDATE TIM SET AJ6=" + KAU6 + " where ID_NUM=" + str(MID_TIM)
        Gdb.Execute "UPDATE TIM SET AJI=" + gDot(str(gVal(POSO.Text))) + " where ID_NUM=" + str(MID_TIM)
         Gdb.Execute "UPDATE TIM SET AJ4=" + gDot(gVal(KAU4)) + " where ID_NUM=" + str(MID_TIM)
        
        Gdb.Execute "UPDATE TIM SET FPA4=" + gDot(gVal(FPA4)) + " where ID_NUM=" + str(MID_TIM)
        
        Gdb.Execute "UPDATE TIM SET FPA6=" + FPA6 + " where ID_NUM=" + str(MID_TIM)

       

























'ıÔÎÔÈ· ıÎÈÍ· 24%
        sql = "INSERT INTO EGGTIM(KODE,PIS,POSO,ONOMA,MONA,FPA,EKPT,EIDOS,APOT,HME,ATIM,TIMM,KAU_AJIA,MIK_AJIA,ID_NUM) VALUES ('23122002',0,1,'Z TAMEIAKH”','‘≈Ã',6,0,'e',1,"




        'HME,ATIM,TIMM,KAU_AJIA,MIK_AJIA)
        sql = sql + "'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "',"
        sql = sql + "'" + "ˆ" + m_tim + "',"
        sql = sql + "" + KAU6 + ","
        sql = sql + "" + KAU6 + ","
        sql = sql + gDot(str(Val(KAU6) + Val(FPA6))) + "," + str(MID_TIM) + ")"
        SQL2 = "SELECT ID FROM EGGTIM WHERE ID = @@IDENTITY"
        Dim K2 As Long
        Gdb.Execute sql, K2


'–≈—…¬¡ÀÀœÕ‘… œ ‘≈Àœ”
If Val(Replace(PERIB.Text, ",", ".")) > 0 Then
        sql = "INSERT INTO EGGTIM(KODE,PIS,POSO,ONOMA,MONA,FPA,EKPT,EIDOS,APOT,HME,ATIM,TIMM,KAU_AJIA,MIK_AJIA,ID_NUM) VALUES ('99-24',0,1,'ÂÒÈ‚','‘≈Ã',6,0,'e',1,"
        'HME,ATIM,TIMM,KAU_AJIA,MIK_AJIA)
        sql = sql + "'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "',"
        sql = sql + "'" + "ˆ" + m_tim + "',"
        
        
        sql = sql + "" + KAU4 + ","
        sql = sql + "" + KAU4 + ","
        sql = sql + gDot(str(Val(KAU4) + Val(FPA4))) + "," + str(MID_TIM) + ")"
        
        
        
        SQL2 = "SELECT ID FROM EGGTIM WHERE ID = @@IDENTITY"
        Gdb.Execute sql, K2
End If


















        'Dim R3 As New ADODB.Recordset

        On Error GoTo 0

        'R3.Open Sql2, Gdb, adOpenDynamic, adLockOptimistic

        'R3.Close
        'R3.Close
        re.Open "SELECT @@IDENTITY FROM EGGTIM", Gdb, adOpenDynamic, adLockOptimistic

        'Dim mID

        m_ideggtim = re(0)
        re.Close

        'sql = "INSERT INTO TIM(KPE,B_N1,EIDPAR,EIDOS,AJ1,AJ2,AJ3,AJ4,AJ5,FPA1,FPA2,FPA3,FPA4) VALUES ('" + M_KOD + "',1,'','e',0,0,0,0,0,0,0,0,0);"
        
        
        
        
'        Sql2 = "SELECT ID_NUM FROM TIM WHERE ID_NUM = @@IDENTITY"
'
'        sql = "INSERT INTO TIM(KLEIDI,ATIM,KPE,B_N1,EIDPAR,EIDOS,AJ1,AJ2,AJ3,AJ4,AJ5,FPA1,FPA2,FPA3,FPA4,AJ6,FPA6,AJ7,FPA7) VALUES "
'        sql = sql + "('" + "ˆ" + m_tim + "','" + "ˆ" + m_tim + "','" + M_KOD + "',1,'','e',0,0,0,0,0,0,0,0,0,0,0,0,0);"
'
'        Gdb.Execute sql
'
'        R3.Open Sql2, Gdb, adOpenDynamic, adLockOptimistic
'        mID = R3(0)
'        R3.Close
'
'        Gdb.Execute "UPDATE TIM SET HME='" + Format(dtpicker1.Value, "mm/DD/YYYY") + "' where ID_NUM=" + str(mID)
'        '         Gdb.Execute "UPDATE TIM SET ATIM='" + "ˆ" + m_tim + "'" + " where ID_NUM=" + str(mID)
'
'        '        Gdb.Execute "UPDATE TIM SET TIMM=" + str(Val(POSO.Text) / ((100 + Val(FPA.Text)) / 100)) + " where ID_NUM=" + str(mID)
'        '        Gdb.Execute "UPDATE TIM SET ATIM='" + "ˆ" + m_tim + "'" + " where ID_NUM=" + str(mID)
'        Gdb.Execute "UPDATE TIM SET AJ6=" + KAU6 + " where ID_NUM=" + str(mID)
'        Gdb.Execute "UPDATE TIM SET AJI=" + gDot(str(gVal(poso.Text))) + " where ID_NUM=" + str(mID)
'         Gdb.Execute "UPDATE TIM SET AJ4=" + gDot(gVal(KAU4)) + " where ID_NUM=" + str(mID)
'
'        Gdb.Execute "UPDATE TIM SET FPA4=" + gDot(gVal(FPA4)) + " where ID_NUM=" + str(mID)
'
'        Gdb.Execute "UPDATE TIM SET FPA6=" + FPA6 + " where ID_NUM=" + str(mID)
'
'


        '540     r.Open "SELECT  *FROM EGG", Gdb, adOpenDynamic, adLockOptimistic
        '550     r.AddNew
        '560     r("hme") = DTPicker1.Value
        '570     r("apa") = Val(m_tim)
        '580     r("atim") = "ˆ" + m_tim
        '590     r("ait") = "Ã≈‘—«‘¡"
        '600     r("XREOSI") = 0: r("PISTOSI") = Val(POSO.Text)    ' ¡–œ÷’√« ÕULL
        '610     r("eidos") = "e"
        '620     r("kod") = M_KOD
        '630     r.Update
        '
        '640     r.AddNew
        '650     r("hme") = DTPicker1.Value
        '660     r("apa") = Val(m_tim)
        '670     r("atim") = "ˆ" + m_tim
        '680     r("ait") = "–ŸÀ«”≈…” ‘¡Ã≈…¡ «”"
        '690     r("XREOSI") = Val(POSO.Text): r("PISTOSI") = 0    ' ¡–œ÷’√« ÕULL
        '700     r("eidos") = "e"
        '710     r("kod") = M_KOD
        '720     r.Update

        '----------------------  P. Y . -------------------------------------
        If Val(Text3.Text) > 0 Then
        
            Dim M_IDEGG As Long
        
            sql = "INSERT INTO EGGTIM(KODE,PIS,POSO,ONOMA,MONA,FPA,EKPT,EIDOS,APOT,HME,ATIM,TIMM,KAU_AJIA,MIK_AJIA) VALUES ('23122002',0,1,'Z TAMEIAKH”','‘≈Ã',6,0,'e',1,"
            sql = sql + "'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "',"
            sql = sql + "'" + F_PY + m_tim + "',"
            sql = sql + "" + str(Val(Text3.Text) / ((100 + Val(FPA.Text)) / 100)) + ","
            sql = sql + "" + str(Val(Text3.Text) / ((100 + Val(FPA.Text)) / 100)) + ","
            sql = sql + str(Val(Text3.Text)) + ")"
            SQL2 = "SELECT ID FROM EGGTIM WHERE ID = @@IDENTITY"
            Gdb.Execute sql, K2
   
            On Error GoTo 0
        If K2 > 0 Then
            re.Open "SELECT @@IDENTITY FROM EGGTIM", Gdb, adOpenDynamic, adLockOptimistic
            M_IDEGG = re(0)
            re.Close
        Else
             M_IDEGG = 0
        End If
   
            SQL2 = "SELECT ID_NUM FROM TIM WHERE ID_NUM = @@IDENTITY"
            'sql = "INSERT INTO TIM(KLEIDI,ATIM,KPE,B_N1,EIDPAR,EIDOS,AJ1,AJ2,AJ3,AJ4,AJ5,FPA1,FPA2,FPA3,FPA4) VALUES "
            'sql = sql + "('" + "ˆ" + m_tim + "','" + F_PY + m_tim + "','" + M_KOD + "',1,'','e',0,0,0,0,0,0,0,0,0);"
   
            'Gdb.Execute sql
            F_PY = Trim(F_PY)
            
            sql = "INSERT INTO TIM(KLEIDI,ATIM,KPE,B_N1,EIDPAR,EIDOS,AJ1,AJ2,AJ3,AJ4,AJ5,FPA1,FPA2,FPA3,FPA4,AJ6,FPA6) VALUES "
            sql = sql + "('" + F_PY + m_tim + "','" + F_PY + m_tim + "','" + M_KOD + "',1,'','e',0,0,0,0,0,0,0,0,0,0,0);"
   
           '
   
   
            F_PY = Trim(F_PY)
            
            sql = "INSERT INTO TIM(KLEIDI,ATIM,KPE,B_N1,EIDPAR,EIDOS,AJ1,AJ2,AJ3,AJ4,AJ5,FPA1,FPA2,FPA3,FPA4,AJ6,FPA6) VALUES "
            sql = sql + "('" + F_PY + m_tim + "','" + F_PY + m_tim + "','" + M_KOD + "',1,'','e',0,0,0,0,0,0,0,0,0,0,0);"
   
   Gdb.Execute sql
   
   
   
            re.Open SQL2, Gdb, adOpenDynamic, adLockOptimistic
            MID_TIM = re(0)
            re.Close
            Gdb.Execute "UPDATE TIM SET HME='" + Format(DTPicker1.Value, "mm/DD/YYYY") + "' where ID_NUM=" + str(MID_TIM)
            Gdb.Execute "UPDATE TIM SET AJ6=" + str(Val(Text3.Text) / ((100 + Val(FPA.Text)) / 100)) + " where ID_NUM=" + str(MID_TIM)
            Gdb.Execute "UPDATE TIM SET AJI=" + str(Val(Text3.Text)) + " where ID_NUM=" + str(MID_TIM)
            Gdb.Execute "UPDATE TIM SET FPA6=AJI-AJ6" + " where ID_NUM=" + str(MID_TIM)
            Z.AddItem m_tim + " " + str(Val(POSO.Text)) + "---" + F_PY + m_tim + str(Val(Text3.Text)) + " ÛıÌÔÎÔ: " + str(Val(POSO.Text) + Val(Text3.Text)), 0

             Gdb.Execute "UPDATE EGGTIM SET ID_NUM=" + str(MID_TIM) + " WHERE ID=" + str(M_IDEGG)




        Else
730         Z.AddItem m_tim + " " + str(Val(POSO.Text)), 0

        End If

        '----------------------  P. Y . -------------------------------------

770     Me.Caption = m_tim + " Í·ÙÂ˜˘ÒﬁËÁ"
780     DTPicker1.Value = DateAdd("d", 1, DTPicker1.Value)
790     MATIM = Format(Val(MATIM) + 1, "####0")
800     POSO.Text = "     "

        Text3.Text = ""
        SYNOLIKO.Text = ""
        PERIB.Text = ""

810     Command4.TabIndex = 90
820     DTPicker1.SetFocus

        Exit Sub

MHNYMA:
830     HandleError "Par1:enhm_eggtim"

840     Resume Next

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub DataGrid1_Click()
'
'End Sub

Private Sub Command7_Click()

        '<EhHeader>
        On Error GoTo Command7_Click_Err

        '</EhHeader>

100     bohu53.SHOW

        '<EhFooter>
        Exit Sub

Command7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Command7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Command7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command8_Click()

        '<EhHeader>
        On Error GoTo Command8_Click_Err

        '</EhHeader>

100     bohu51.SHOW

        '<EhFooter>
        Exit Sub

Command8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Command8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Command8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command9_Click()

        '<EhHeader>
        On Error GoTo Command9_Click_Err

        '</EhHeader>

        '100  GoTo 66

        '
        '  'Connect
        '    If RapiConnect Then
        '        Me.Caption = "Status: Connected!"
        '    Else
        '        Me.Caption = "Status: Disconnected!"
        '    End If
        '
        '
        '   If Not RapiIsConnected Then
        '        MsgBox "Device is not connected. Please connect first."
        '        Exit Sub
        '    End If
        '    If FileExists("C:\MERCVB\TIMOL.TXT") Then
        '       On Error Resume Next
        '       Kill "C:\MERCVB\TIMOL.TXT"
        '    End If
        '
        ''ÎÁ¯Á apo PDA
        '    Call RAPICopyCEFileToPC("\EGGTIM.TXT", "C:\MERCVB\TIMOL.TXT")

        '====================================================================================

        'PREPEI NA MHN DEXETAI DIPLA
        'KAI NA GRAFEI TON TROPO PLIROMIS

        Dim m_ait

        Dim m_onoma

        Dim k, a

        Dim DB As Database

        Dim R  As New ADODB.Recordset

        Dim c, D, e, F, G, n1, n2

        Dim m_hme  ' = hme ' ‘≈À≈’‘¡…¡ «Ã≈—œÃ«Õ…¡

        Dim ATIM, KODPEL, KODEID, POS, TIMM, hme

        Dim CC

        On Error GoTo LATOS

110     CC = CurDir

120     CC = "c:\mercvb"
130     m_ait = ".."

        'Set db = OpenDatabase(CC, False, False, "dBase III;")

        Dim ArTim, M_PEL

140     ArTim = "ddddd"
150     Me.Caption = "OPENDBASE"

        Dim Col(15, 2)

        '1,6    AR.TIMOL                         1
        '8,10   HMEPOMHNIA                      2
        '20,14   KODIKOS EIDOYS                 3
        '35,30   ONOMASIA EIDOYS                4
        '66,10   TIMH MONADOS                   5
        '77,10   POSOTHTA                       6
        '88,5   EKPTOSI %                       7
        '96,3   FPA                             8
        '100,6   KODIKOS PELATH                     9
        '110,4   TROPOS PLIROMIS                    10
        '115,10  afairoymeni  EKPTOSI petrelaioy        11

        'TO IDIO YPARXEI KAI STO PDA GIA EXAGOGH
        Dim COMMA As Long

        Dim NK    As Integer

160     NK = 0
170     Open CC + "\EGGTIMCOLS.TXT" For Input As #2

180     Do While Not EOF(2)
190         Line Input #2, a
200         COMMA = InStr(a, ",")
210         NK = NK + 1
220         Col(NK, 1) = Val(Left(a, COMMA))
230         Col(NK, 2) = Val(mID(a, COMMA + 1, 3))
        Loop

240     Close #2

250     Open CC + "\LOG.txt" For Output As #4

260     Me.Caption = "≈Õ¡—Œ«-¡Õœ…√Ã¡ ¡—◊≈…œ’ C:\TIMOL.TXT"
270     Open CC + "\EGGTIM.txt" For Input As #1

280     Line Input #1, a

290     Me.Caption = 11

        Dim gm_fpa, m_E

        Dim M_POSO(5), syn_m_POSO

300     M_POSO(1) = 0
310     M_POSO(2) = 0
320     M_POSO(3) = 0
330     M_POSO(4) = 0
340     M_POSO(5) = 0
350     syn_m_POSO = 0
360     m_E = "0"

370     Me.Caption = "À«Œ«-¡Õœ…√Ã¡ ¡—◊≈…œ’ C:\TIMOL.TXT"

        Dim nn

380     nn = 0

        '--------------------------------------------------------------------------------------
390     Do While Not EOF(1)
400         nn = nn + 1
410         Line Input #1, a
420         Me.Caption = Format(nn, "####") + " " + a
430         List1.AddItem Format(nn, "####") + " " + a

            ' KODPEL = mID(A, 117, 6) 'R("shm1") ' Í˘‰.ÂÎ·ÙÁ
            ' KODPEL = LTrim(mID(A, Col(9, 1), Col(9, 2) + 1)) 'R("shm1") ' Í˘‰.ÂÎ·ÙÁ
440         KODPEL = LTrim(mID(a, Col(9, 1), Col(9, 2) + 1))    'R("shm1") ' Í˘‰.ÂÎ·ÙÁ

450         R.Open "select * from EID where KOD='" + Trim(LTrim(mID(a, Col(3, 1), Col(3, 2)))) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

460         If R.EOF Then
470             MsgBox "ƒ≈Õ ¬—≈»« ≈ œ  Ÿƒ… œ” " + Trim(LTrim(mID(a, Col(3, 1), Col(3, 2))))
480             Close #1
490             Close #4

                Exit Sub

            End If

500         gm_fpa = R("fpa")
510         m_onoma = R("ono")

520         KODEID = R("KOD")    ' Í˘‰.ÂÈ‰ÔıÚ
530         R.Close

540         ATIM = mID(a, Col(1, 1), Col(1, 2))  ' mID(A, 4, 7) 'ariumos timologioy

550         Me.Caption = "111"

            'If F_TIM = 1 Then
            '   atim = "T" + Format(Val(atim), "000000")
            'Else
            '  If InStr(1, Right(Trim(A), 4), "-") Then  ' TIMOLOGIO
            'atim = "T" + Format(Val(atim), "000000")
            '  Else
            '    atim = "L" + Format(Val(atim), "000000")
            '  End If
            'End If
            '
            '

            '1,6    AR.TIMOL
            '8,10   HMEPOMHNIA
            '20,14   KODIKOS EIDOYS
            '35,30   ONOMASIA EIDOYS
            '66,10   TIMH MONADOS
            '77,10   POSOTHTA
            '88,5   EKPTOSI %
            '96,3   FPA
            '100,6   KODIKOS PELATH
            '110,4   TROPOS PLIROMIS
            '115,10  afairoymeni  EKPTOSI petrelaioy

560         Me.Caption = "22"

570         POS = Val(mID(a, Col(6, 1), Col(6, 2)))  'ÔÛÔÙÁÙ·
580         TIMM = Val(mID(a, Col(5, 1), Col(5, 2)))   'ÙÈÏﬁ ÏÔÌ‹‰·Ú

590         Me.Caption = "333"

            '–—œ”œ◊« « «Ã≈—œÃ«Õ…¡ ”‘œ ¡SCII EINAI –.◊.  23/04/2010   ¡… œ◊… 04/23/2010
600         hme = mID(a, Col(2, 1), Col(2, 2))   'ÁÏÂÒÔÏÁÌﬂ·
610         hme = mID(hme, 4, 2) + "/" + Left(hme, 2) + "/" + mID(hme, 7, 4)

            'G = CDate(hme)
            ' hme = Format(G, "dd/mm/yyyy")

620         e = Left(Trim(mID(a, Col(10, 1), Col(10, 2))), 1)  ' –…”‘Ÿ”« Ã≈‘—«‘œ…”

630         If e = "–" Then
640             e = "0"  '–…”‘Ÿ”«
            Else
650             e = "1"  ' Ã≈‘—«‘œ…”
            End If

660         Me.Caption = "4444"

670         If ArTim <> ATIM Then
680             syn_m_POSO = M_POSO(1) * 1.09 + M_POSO(2) * 1.19 + M_POSO(3) * 1.13 + M_POSO(5)

                ' Print #2, "       "
                'Print #2, to437(atim + hme + KODPEL) + "     " + e  m_pel
690             If ArTim <> "ddddd" Then

                    '        R.Open "select COUNT(*) FROM TIM WHERE ATIM='" + ArTim + "' AND HME='" + Format(m_hme, "MM/DD/YYYY") + "'", Gdb, adOpenDynamic, adLockBatchOptimistic
                    '        If R(0) > 0 Then
                    '           MsgBox "‘œ ‘…ÃœÀœ√…œ " + ArTim + " ≈…Õ¡… «ƒ«  ¡‘¡◊Ÿ—«Ã≈ÕO "
                    '           COMMA = MsgBox("NA ƒ…¡√—¡÷≈… ‘œ «ƒ« ’–¡—◊ŸÕ? " + ArTim, vbYesNo)
                    '           If COMMA = vbNo Then
                    '              R.Close
                    '              Exit Sub
                    '           Else
                    '              Gdb.Execute "DELETE FROM TIM WHERE ATIM='" + ArTim + "' AND HME='" + Format(m_hme, "MM/DD/YYYY") + "'"
                    '              Gdb.Execute "DELETE FROM EGGTIM WHERE ATIM='" + ArTim + "' AND HME='" + Format(m_hme, "MM/DD/YYYY") + "'"
                    '              Gdb.Execute "DELETE FROM EGG WHERE ATIM='" + ArTim + "' AND HME='" + Format(m_hme, "MM/DD/YYYY") + "'"
                    '           End If
                    '        End If
                    '        R.Close

700                 G = "INSERT INTO TIM " & "(EIDOS,B_N1,HME,ATIM,KPE,AJ1,AJ2,AJ3,AJ4,AJ5,AJI)" & " VALUES ('e',1," & "'" + m_hme + "'," & "'" + ArTim + "'," & "'" + M_PEL + "'," & Replace(Format(M_POSO(1), "0000000.00"), ",", ".") + "," & Replace(Format(M_POSO(2), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(3), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(3), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(5), "###0.00"), ",", ".") + "," & Replace(Format(syn_m_POSO, "0000000.00"), ",", ".") + ");"

710                 Write #4, G

720                 Gdb.Execute G, k

730                 G = "INSERT INTO EGG " & "(HME,ATIM,KOD,EIDOS,APA,XRE,XREOSI,AIT)" & " VALUES (" & "'" + hme + "'," & "'" + ArTim + "'," & "'" + M_PEL + "'," & "'e'," & "'" + mID(ArTim, 2, 6) + "'," & Replace(Format(syn_m_POSO, "0000000.00"), ",", ".") + "," & Replace(Format(syn_m_POSO, "###0.00"), ",", ".") + ",'" & m_ait + "');"

740                 Write #4, G

750                 Gdb.Execute G, k

760                 If m_E <> "0" Then
                        '·Ò˜ÂÈÔ ÏÂÙÒÁÙÔÈÛ 3.pni

770                     G = "INSERT INTO EGG " & "(HME,ATIM,KOD,EIDOS,APA,XRE,PISTOSI,AIT)" & " VALUES (" & "'" + hme + "'," & "'" + ArTim + "'," & "'" + M_PEL + "'," & "'e'," & "'" + mID(ArTim, 2, 6) + "'," & Replace(Format(syn_m_POSO, "0000000.00"), ",", ".") + "," & Replace(Format(syn_m_POSO, "###0.00"), ",", ".") + "," & "'Ã≈‘—«‘¡');"

780                     Write #4, G
790                     Gdb.Execute G, k

                    End If
                End If

800             M_POSO(1) = 0
810             M_POSO(2) = 0
820             M_POSO(3) = 0
830             M_POSO(4) = 0
840             M_POSO(5) = 0
850             ArTim = ATIM
            End If

            'ÂÎÂ„˜ÔÚ ‰ÈÎÔÂ„„Ò·ˆﬁÚ

860         R.Open "select COUNT(*) FROM TIM WHERE ATIM='" + ArTim + "' AND HME='" + Format(m_hme, "MM/DD/YYYY") + "'", Gdb, adOpenDynamic, adLockBatchOptimistic

870         If R(0) > 0 Then
                ' MsgBox "‘œ ‘…ÃœÀœ√…œ " + ArTim + " ≈…Õ¡… «ƒ«  ¡‘¡◊Ÿ—«Ã≈ÕO "
                ' COMMA = MsgBox("NA ƒ…¡√—¡÷≈… ‘œ «ƒ« ’–¡—◊ŸÕ? " + ArTim, vbYesNo)
880             COMMA = vbYes

890             If COMMA = vbNo Then
900                 R.Close

                    Exit Sub

                Else
910                 Gdb.Execute "DELETE FROM TIM WHERE ATIM='" + ArTim + "' AND HME='" + Format(m_hme, "MM/DD/YYYY") + "'"
920                 Gdb.Execute "DELETE FROM EGGTIM WHERE ATIM='" + ArTim + "' AND HME='" + Format(m_hme, "MM/DD/YYYY") + "'"
930                 Gdb.Execute "DELETE FROM EGG WHERE ATIM='" + ArTim + "' AND HME='" + Format(m_hme, "MM/DD/YYYY") + "'"
                End If
            End If

940         R.Close

950         G = "INSERT INTO EGGTIM " & "(EIDOS,APOT,KODE,HME,ATIM,PELKOD,POSO,TIMM,PIS,ONOMA)" & " VALUES ('e',1,'" + KODEID + "'," & "'" + hme + "'," & "'" + ATIM + "'," & "'" + KODPEL + "'," & Replace(Format(POS, "0000000.00"), ",", ".") + "," & Replace(Format(Round(TIMM, 2), "###0.00"), ",", ".") + "," & Replace(Format(POS, "0000000.00"), ",", ".") + ",'" + m_onoma + "');"
960         Write #4, G

970         Gdb.Execute G, k

980         If gm_fpa = 1 Then
990             M_POSO(1) = M_POSO(1) + POS * TIMM
1000        ElseIf gm_fpa = 2 Then
1010            M_POSO(2) = M_POSO(2) + POS * TIMM
1020        ElseIf gm_fpa = 3 Then
1030            M_POSO(3) = M_POSO(3) + POS * TIMM
1040        ElseIf gm_fpa = 5 Then
1050            M_POSO(5) = M_POSO(5) + POS * TIMM    '      '      0 %
            End If

1060        If Left(ATIM, 1) = "T" Then    ' TIMOLOGIO
1070            m_ait = "ƒ¡-‘…ÃœÀœ√…œ –ŸÀ«”«”"

1080        ElseIf Left(ATIM, 1) = "L" Then   ' TIMOLOGIO
1090            m_ait = "¡–œƒ.À…¡Õ… «”"
            Else
1100            m_ait = "ƒ≈À‘…œ ¡–œ”‘œÀ«”"
            End If

1110        M_PEL = KODPEL
1120        m_hme = hme    ' ‘≈À≈’‘¡…¡ «Ã≈—œÃ«Õ…¡
1130        m_E = e    ' ÈÛÙ˘ÛÁ ÙÂÎÂıÙ·ﬂ·
        Loop

        '-------------------------------------------------------------------------------------------------
1140    Close #1

1150    syn_m_POSO = M_POSO(1) * 1.09 + M_POSO(2) * 1.19 + M_POSO(3) * 1.13 + M_POSO(5)
1160    G = "INSERT INTO TIM " & "(EIDOS,HME,ATIM,KPE,AJ1,AJ2,AJ3,AJ4,AJ5,AJI)" & " VALUES ('e'," & "'" + Format(m_hme, "MM/DD/YYYY") + "'," & "'" + ATIM + "'," & "'" + M_PEL + "'," & Replace(Format(M_POSO(1), "0000000.00"), ",", ".") + "," & Replace(Format(M_POSO(2), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(3), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(4), "###0.00"), ",", ".") + "," & Replace(Format(M_POSO(5), "###0.00"), ",", ".") + "," & Replace(Format(syn_m_POSO, "0000000.00"), ",", ".") + ");"

1170    Write #4, G

1180    Gdb.Execute G

1190    If m_E <> "0" Then
            '·Ò˜ÂÈÔ ÏÂÙÒÁÙÔÈÛ 3.pni

1200        G = "INSERT INTO EGG " & "(HME,ATIM,KOD,EIDOS,APA,XRE,PISTOSI,AIT)" & " VALUES (" & "'" + Format(hme, "MM/DD/YYYY") + "'," & "'" + ATIM + "'," & "'" + M_PEL + "'," & "'e'," & "'" + mID(ATIM, 2, 6) + "'," & Replace(Format(syn_m_POSO, "0000000.00"), ",", ".") + "," & Replace(Format(syn_m_POSO, "###0.00"), ",", ".") + "," & "'Ã≈‘—«‘¡');"

1210        Write #4, G

        End If

1220    Close #4

1230    MsgBox "œ  ÂÌÁÏÂÒ˛ËÁÍÂ"

        Exit Sub

LATOS:
1240    MsgBox "‰È·ÍÔÁ " + Chr(13) + Err.Description

        'Close #1

        '<EhFooter>
        Exit Sub

Command9_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Command9_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Command9_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "BOHU5"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command2_Click

110     If KeyCode = vbKeyReturn Then
120         KeyCode = 0    'suppress the beep
130         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     Me.Picture = LoadPicture(gPicture)
110     Me.KeyPreview = True
120     mForm_Load Me, fh, fw, ft, fl



F_KATHG_METAG = Val(FINDPARAMETROI(1, "BOHU5", "F_KATHG_METAG", "3", " ¡‘«√œ—…¡ –—œÃ«»≈’‘ŸÕ Ã≈ Ã≈‘¡÷œ—… ≈”"))   ' ÃœÕœ ‘…ÃœÀœ√…¡

F_DIR_APOS_KARAM = (FINDPARAMETROI(1, "BOHU5", "F_DIR_APOS_KARAM", "C:\CL\DATA", "÷¡ ≈Àœ” ÷œ—‘Ÿ”«”  ¡—¡Ã≈”…Õ«"))   ' ÃœÕœ ‘…ÃœÀœ√…¡
F_DIR_EPIS_KARAM = (FINDPARAMETROI(1, "BOHU5", "F_DIR_EPIS_KARAM", "C:\CL\DATA", "÷¡ ≈Àœ” ≈–…”‘—œ÷«”  ¡—¡Ã≈”…Õ«"))   ' ÃœÕœ ‘…ÃœÀœ√…¡


 F_DOUTSIOSREM = (FINDPARAMETROI(1, "BOHU5", "F_DOUTSIOSREM", "[2.3.4.5,49705]", "ONOMA LINKED SERVER ‰ÔıÙÛÈÔı –.X.2.3.4.5,49705"))   ' ÃœÕœ ‘…ÃœÀœ√…¡
'                                    F_DOUTSIOSPARAS = (FINDPARAMETROI(1, "BOHU5", "F_DOUTSIOSPARAS", "‰", "·Ò·ÛÙ·ÙÈÍ¸ ˆ¸ÒÙ˘ÛÁÚ ıÔÍ·Ù·ÛÙﬁÏ·ÙÔÚ"))   ' ÃœÕœ ‘…ÃœÀœ√…¡

F_LINKEDSERVER = (FINDPARAMETROI(1, "BOHU5", "F_LINKEDSERVER", "", "ONOMA ¡–œÃ.SERVER ME TO ONOMA LINKEDSERVER –.◊. [192.11.1.1\SQLEXPRESS,49818].[EMP]"))   ' ÃœÕœ ‘…ÃœÀœ√…¡
F_LINKED_PARAST = (FINDPARAMETROI(1, "BOHU5", "F_LINKED_PARAST", "AT", " Ÿƒ… œ… –¡—/ ŸÕ –œ’ »¡ √…Õœ’Õ IMPORT –.◊. AT "))    ' ÃœÕœ ‘…ÃœÀœ√…¡



 If InStr(F_DOUTSIOSREM, "DOUTSIOS") = 0 Then
    IMPORTSERV.Enabled = False
    
    
 End If


Dim R As New ADODB.Recordset


On Error GoTo parakato

1000    R.Open "SELECT AFM+';'+EPO AS PERIGRAFH2 FROM PEL WHERE EIDOS='r' AND KODGAL=" + str(F_KATHG_METAG) + " ORDER BY EPO", Gdb, adOpenDynamic, adLockOptimistic

        

        'R.Index = "pinakes"

        'FPA
        'R.Seek "=", 1, 1
1010    Do While Not R.EOF

1020        'If R("typos") = 1 Then
1030            AFM.AddItem R!PERIGRAFH2  ' str(R("AYJON")) + " -> " + str(R("TIMH"))
            'End If

1040        R.MoveNext
        Loop

        ' mon.metrhshs
1050    R.Close

parakato:


On Error GoTo Form_Load_Err


'130     AFM.AddItem "084137936 NEA ƒ—¡Ã¡"
'140     AFM.AddItem "998876400  ¡——¡”"
'150     AFM.AddItem "010466342 ¡–œ”‘œÀ…ƒ«”"
'160     AFM.AddItem "006447226 √≈Ÿ—√…¡ƒ«”  ’—…¡ œ”"
'170     AFM.AddItem "030545281 …Ÿ¡ÕÕ…ƒ«”"
'180     AFM.AddItem "999208470 ≈–‘¡”"
'190     AFM.AddItem "39923450 ƒœŒ¡"
'200     AFM.AddItem "082498076 ƒ≈À‘¡ Ã≈‘¡÷œ—… «"

210     DTPicker1.Value = Now

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

220     F_TIM = Val(FINDPARAMETROI(1, "BOHU5", "F_TIM", "0", "MONO TIMOÀœ√…¡ IMPORT"))   ' ÃœÕœ ‘…ÃœÀœ√…¡
        

224     F_PY = FINDPARAMETROI(1, "BOHU5", "F_PY", "ƒ", " Ÿƒ… œ” –¡—/ œ’ –/’ ∆")   ' ÃœÕœ ‘…ÃœÀœ√…¡

230     FPA.Text = g_Fpa(6)

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'-------------------------------------------------------------------------------------
Private Sub Command6_Click()

        ' ¡–œ ≈◊CEL
        'C:\MODELO2.XLS
        ' œ…  ’–œ  ONO XTI LTI KOD erg BARCODE fpa mon
        '2-71300 1   1   GIM ¡ÀÃ–œ’Ã S200 MICKEY FACES   2.95     2.95    5204549271302   5204549271302
        '2-71310 1   1   GIM ¡ÀÃ–œ’Ã Ã100 MICKEY FACES   1.50     1.50    9999999999999   5204549713109
        '654554  1   22  ” ’À¡—œ”                        10.00   20.90    5204182081290   5204182081290
        '56434   1   3   À¡À¡œ                            5.00   10.00    5555            5555
        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>

        ' Data1.Connect = "Excel 8.0;"
        ' Data1.DatabaseName = "C:\MODELO2.xls"
        '  Data1.RecordSource = "÷˝ÎÎÔ1$"

        Dim a

        Dim F928 As Integer

100     F928 = 1

110     If Len(Combo1.Text) = 0 Then
120         MsgBox "≈–…À≈Œ‘≈ ¡—◊≈…œ EID / PEL"

            Exit Sub

        End If

130     Adodc2.ConnectionString = gConnect
        ' Adodc2.RecordSource = "select * FROM " + Combo1.Text 'PEL WHERE  EIDOS='e' ORDER BY KOD"
        '   Adodc2.Refresh

        'Set fSCR = CreateObject("MSScriptControl.ScriptControl")
        'fSCR.language = "vbscript"
        'fSCR.AddObject "ADODC2", Adodc2
        'fSCR.AddObject "mactext", MACtEXT

        '  Adodc2.RecordSource = "select * FROM EID ORDER BY KOD"
        '  Adodc2.Refresh

140     a = toascii3("C:\MERCVB\F33.TXT", "C:\EID.TXT")

150     MsgBox "‘›ÎÔÚ Àπÿ«”"

        'Me.Caption = mac("Adodc2.Recordset(3)")

        'Set fSCR = Nothing

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function toascii3(ByVal arxeio As String, moutput As String)

        '<EhHeader>
        On Error GoTo toascii3_Err

        '</EhHeader>

        '·Ô EXCEL

        Dim DUM

        Dim eField(1 To 50), ePict(1 To 50), eProw(1 To 50), ePcol(1 To 50), eFonts(1 To 50)

        Dim mField(1 To 50), mPict(1 To 50), mProw(1 To 50), mPcol(1 To 50), mFonts(1 To 50)

        Dim sField(1 To 50), sPict(1 To 50), sProw(1 To 50), sPcol(1 To 50), sFonts(1 To 50)

        Dim m_No_of_seir, i, k, m, c

        Dim m_seir_synol, m_seir_eid, sf

        Dim a(1 To 250), ar_ped(1 To 250)

        Dim xa(1 To 250), m_m, npic

        Dim m1

        ' dim gm_str(1 To 250)
        Dim DB     As Database

        Dim TIM    As Recordset

        Dim PEL    As Recordset

        Dim EGGTIM As Recordset

        Dim m_entol_ektyp

        Dim xart11, kod_tim

        Dim ejodos

        Dim e, MM, s

        Dim ektypoths

        Dim m_syn_row, syn_row

        Dim CDOK

        Dim h$

100     List1.Clear
110     List1.AddItem "Ã«  ¡‘¡◊Ÿ—«»≈Õ‘¡  Ÿƒ/BARC"

120     List2.Clear
130     List2.AddItem " ¡‘¡◊Ÿ—«»≈Õ‘¡  Ÿƒ/BARC"

        'Data1.Refresh
        'ELEM.MoveFirst

        'If IsNull(ELEM(6)) Then
        ' MsgBox "AƒEIO ¡—◊≈…œ"""
        ' Exit Function
        'End If
        Dim OK As Integer

140     Open "C:\MODELO2.CSV" For Input As #1

        Dim AR(8), SS$

150     Input #1, SS$

160     Do While Not EOF(1)    ' ELEM.EOF
170         Input #1, SS$

180         DUM = FETES_DELIM(SS$, ELEM)

190         If IsNull(ELEM(6)) Then Exit Do

200         Adodc2.RecordSource = "SELECT * FROM EID WHERE KOD='" + LTrim(ELEM(6)) + "'"
210         Adodc2.Refresh

220         If Adodc2.Recordset.EOF Then
230             Adodc2.Recordset.AddNew
            End If

            ' œ…  ’–œ  ONO XTI LTI KOD erg BARCODE fpa mon

            ' On Error Resume Next

240         Adodc2.Recordset("ERG") = ELEM(8)
250         Adodc2.Recordset("AEG") = ELEM(1)
260         Adodc2.Recordset("KODLOG") = ELEM(2)
270         Adodc2.Recordset("ONO") = ELEM(3)

280         If Val(ELEM(9)) = 0 Then
290             Adodc2.Recordset("FPA") = 2
            Else
300             Adodc2.Recordset("FPA") = ELEM(9)
            End If

310         Adodc2.Recordset("MON") = ELEM(10)
320         Adodc2.Recordset("XTI") = ELEM(4)

            'If Not IsNull(Adodc2.Recordset("LTI5")) Then
330         Adodc2.Recordset("LTI5") = ELEM(5)
            ' End If

340         Adodc2.Recordset("KOD") = LTrim(ELEM(6))

350         OK = 1

            On Error GoTo NOTOK

360         Adodc2.Recordset.Update

            On Error Resume Next

370         If OK = 1 Then
380             List2.AddItem ELEM(6) + "/" + ELEM(8)
            End If

            'On Error Resume Next
390         e = "INSERT INTO BARCODES (KOD,ERG) VALUES ('" + LTrim(ELEM(6)) + "','" + LTrim(ELEM(7)) + "')"

            On Error Resume Next

400         Gdb.Execute e

410         DoEvents
420         Me.Caption = h$

            '   Print #1, S2
            '     ELEM.MoveNext
        Loop

        ' Close #1
430     Adodc2.Recordset.Close
440     Close #1

450     If Combo1.Text = "EID" Then
            '    Gdb.Execute "UPDATE EID SET FPA=2 WHERE FPA=19"
            '   Gdb.Execute "UPDATE EID SET FPA=1 WHERE FPA=9"
        End If

        Exit Function

NOTOK:
460     List1.AddItem ELEM(6) + "/" + ELEM(8)
470     OK = 0

480     Resume Next

        '<EhFooter>
        Exit Function

toascii3_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.toascii3 " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.toascii3 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function FETES_DELIM(LINE, ELEM)

        '<EhHeader>
        On Error GoTo FETES_DELIM_Err

        '</EhHeader>
        Dim KL, KE

        '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
100     For KE = 1 To 50
110         ELEM(KE) = ""
        Next

120     KL = 0   ' metraei xaraktires
130     KE = 0   ' metritis toy pinaka  ELEMENT

        Do
140         KE = KE + 1  ' metritis toy pinaka  ELEMENT
150         KL = KL + 1  ' metraei xaraktires

160         Do While mID$(LINE, KL, 1) <> ";"    ' tab
170             ELEM(KE) = ELEM(KE) + mID$(LINE, KL, 1)
180             KL = KL + 1  ' metraei xaraktires

190             If KL > Len(LINE) Then Exit Do
200             If mID$(LINE, KL, 1) = Chr(3) Then Exit Do
            Loop

210         If KL > Len(LINE) Then Exit Do

220     Loop Until mID$(LINE, KL, 1) = Chr(3)    'OLO TO MHKOS THS GRAMMHS

230     FETES_DELIM = 0

        '<EhFooter>
        Exit Function

FETES_DELIM_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu5.FETES_DELIM " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu5.FETES_DELIM " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 15925, 9075, fh, fw, ft, fl
End Sub

Private Sub IMPORT_MHLIOY_Click()
Dim D1, d2 As String

D1 = "01/01/2026"
d2 = "31/01/2026"
D1 = InputBox("¡–œ ", "·¸", D1)

d2 = InputBox("≈Ÿ” ", "›˘Ú", d2)

If Len(Trim(D1)) <> 10 Or Len(Trim(d2)) <> 10 Then
    MsgBox " À¡»œ” «Ã≈—œÃ«Õ…¡"
    Exit Sub
End If

D1 = mID(D1, 4, 3) + Left(D1, 3) + Right(D1, 4)
d2 = mID(d2, 4, 3) + Left(d2, 3) + Right(d2, 4)



Dim FEGGTIM As String
Dim COLTIM As String
Dim TEL_ID_NUM As Long
Dim eis As Integer
Dim N As Integer


 FEGGTIM = "EIDOS,ATIM,POSO,MONA,TIMM,KERDOS,KODE,HME,ERGO,FPA,PROOD,PROOD_AJ,EKPT,KAU_AJIA,MIK_AJIA,ONOMA,MIKTA,KOLA,PELKOD,PROELEYSH,XRE,PIS,APOT,ATIM2,FCURRENCY,EKPT2,MIKTAKILA"
  COLTIM = "KLEIDI,KPE,HME,TRP,ATIM,ART,AJI,EIDOS,METAF,EKPT,EIDPAR,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,TYP,AJ1, AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9,EKPT1,EKPT2,EKPT3,EKPT4,EKPT5,HME_KATAX,KERDOS,SKOPOS,SXETIKO,PARAT,ELGA,SYNPOS,SKOPOS2,FORTOSH,PROOR,AYTOK,B_C1,B_C2,B_N1,B_N2,KR1,KR2,ATIM2"
 
'Dim F_LINKEDSERVER As String '
'Dim F_LINKED_PARAST As String
'INSERT INTO PEL (EIDOS,KOD,EPO) SELECT 'r',KOD,EPO
' FROM [LAGAKIS.DDNS.NET\SQLEXPRESS,49818].[EMP].[dbo].[PEL] WHERE EIDOS='r'
Dim R As New ADODB.Recordset
F_LINKEDSERVER = Trim(F_LINKEDSERVER)
R.Open "SELECT * FROM " + F_LINKEDSERVER + ".dbo.TIM WHERE LEFT(ATIM,1) IN (" + F_LINKED_PARAST + ") AND  HME>='" + D1 + "' AND HME<='" + d2 + "'", Gdb, adOpenDynamic, adLockOptimistic
Dim mmID As Long
Do While Not R.EOF
    mmID = R!id_num
    If GGET_NVALUE("SELECT COUNT(*) FROM TIM WHERE ATIM='" + R!ATIM + "'") > 0 Then
    
    Else
        eis = eis + 1
        Gdb.Execute "INSERT INTO TIM (" + COLTIM + ") SELECT " + COLTIM + " FROM " + F_LINKEDSERVER + ".dbo.TIM  WHERE  ID_NUM=" + str(mmID), N
        If N = 1 Then
            TEL_ID_NUM = GGET_NVALUE("select MAX(ID_NUM) from TIM ")
            Gdb.Execute "INSERT INTO EGGTIM (" + FEGGTIM + ",ID_NUM ) SELECT " + FEGGTIM + "," + str(TEL_ID_NUM) + " FROM " + F_LINKEDSERVER + ".dbo.EGGTIM  WHERE ID_NUM=" + str(mmID)
            Gdb.Execute "update " + F_LINKEDSERVER + ".dbo.TIM set ART='1' WHERE ID_NUM=" + str(mmID)
       
       
       End If
    End If
    DoEvents
    Me.Caption = str(eis) + " " + R!ATIM
    R.MoveNext
    'Exit Do
Loop


MsgBox "OK " + str(eis)









End Sub

Private Sub KARAM_Click()
Dim ROM As New ADODB.Recordset
 '  a = toascii("C:\MERCVB\FPERP.TXT", "C:\TEMP\PEL.TXT")
    Dim R As New ADODB.Recordset

    Dim PIN(100, 2)

    ' FORTONO TON PINAKA 20 ME TIS ANTISTOIXISEIS TON PEDION
    '  ¡… ≈À≈√◊Ÿ ¡Õ ≈◊≈… ƒ«ÀŸ»≈… œ  Ÿƒ… œ”

    Dim COUNTER As Long: COUNTER = 0

    
    
    '------------------------------------------- √…œ—¡Õ…ƒ…” –≈À¡‘≈”--------------------------------
   ' Dim R As New ADODB.Recordset
Dim sql As String


sql = "SELECT * FROM PEL WHERE EIDOS='e' "
'CSTM.HECODE, [CSTM].HENAME,
'ISNULL(CNTC.HETIN,'')  AS AFM
',ISNULL( [HEPROFESSIONS].HENAME,'')  AS [EPAGG], ISNULL([HETAXOFFICES].[HENAME],'')  AS[DOY], ISNULL([TRBR].[HEPHONE1],'') AS THL1,
'ISNULL( [TRBR].[HEPHONE2],'') AS THL2
',ISNULL( [TRBR].HECITY,'') AS [POL],
'ISNULL([TRBR].HESTREET,'') AS DIE,ISNULL(TRBR.HESTREETNUMBER,'') AS ARITMOS ,
'ISNULL(( SELECT SUM( HEEBALANCE) FROM [HETRADERACCUMULATORS] WHERE  ([CSTM].[HEID] =[HECSTMID] and [TRBR].[HEID] = [HETRBRID]
'AND  [CSTM].[HEACTIVE] = 1 AND  YEAR(HEDATE)=2021)),0)  AS DD,HEDISCOUNT AS EKPTOSIS
'from [HECUSTOMERS] [CSTM] WITH(NOLOCK)
'inner join [HETRADERS] [TRDR] WITH(NOLOCK)  on ([CSTM].[HETRDRID] = [TRDR].[HEID])
'Inner Join
'[HECONTACTS] [CNTC] WITH(NOLOCK)  on ([TRDR].[HECNTCID] = [CNTC].[HEID])
'inner join [HETRADERBRANCHES] [TRBR] WITH(NOLOCK)  on ([TRBR].[HETRDRID] = [TRDR].[HEID] )
'left join [HETAXOFFICES] WITH(NOLOCK)  on ([TRDR].[HETFFCID] = [HETAXOFFICES].[HEID])
'left join [HEPROFESSIONS] WITH(NOLOCK)  on ([TRBR].[HEPRFSID] = [HEPROFESSIONS].[HEID])
'ORDER BY CSTM.HECODE


'Dim F_DIR_APOS_KARAM As String '
'Dim F_DIR_EPIS_KARAM As String








R.Open sql, Gdb, adOpenDynamic, adLockOptimistic

Open F_DIR_APOS_KARAM + "\CUST" For Output As #1
Dim s As String

Dim FPA As String
' If IS_MERC = 1 Then
'      Dim R2 As New ADODB.Recordset
'      GMERC.Execute "UPDATE PEL SET AYP=0"
'      Me.Caption = "ok merc=1"
' End If
 
      
Do While Not R.EOF
   s = ""
   s = s + Left(R!kod + Space(15), 15) + " "
   s = s + Left(to437(R!EPO) + Space(35), 31) + " "
   s = s + Left(to437(CNull(R!EPA)) + Space(35), 22) + " "
   
   
'   If IS_MERC = 1 Then
'      Debug.Print "ok 1"
'      R2.Open "SELECT COUNT(*) FROM PEL   WHERE  EIDOS='e' and KOD='" + R!HECODE + "'", GMERC, adOpenDynamic, adLockOptimistic
'      If R2(0) = 0 Then
'       Debug.Print "ok 2"
'          GMERC.Execute "INSERT INTO PEL (EIDOS,KOD) VALUES ('e','" + R!HECODE + "')"
'      End If
'      R2.Close
'
'
'
'   GMERC.Execute "update PEL set EPO='" + Left(R!HENAME, 35) + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
'
'    GMERC.Execute "update PEL set EPA='" + R!EPAGG + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
'
'   GMERC.Execute "update PEL set DIE='" + Left(R!DIE, 35) + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
'
'   GMERC.Execute "update PEL set POL='" + R!pol + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
'
'   GMERC.Execute "update PEL set AFM='" + R!AFM + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
'   GMERC.Execute "update PEL set DOY='" + Replace(R!doy, "'", "`") + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
'   GMERC.Execute "update PEL set AYP=ISNULL(AYP,0)+" + Replace(Format(R!DD, "####0.00"), ",", ".") + "  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
'
'   GMERC.Execute "update PEL set THL='" + Left(R!THL1, 10) + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
'
'   GMERC.Execute "update PEL set PEK=" + Replace(Format(R!EKPTOSIS, "####0.00"), ",", ".") + "  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
' End If
   
   ' Left(to437(R!THL1)
   
   
   
   
   s = s + Left(to437(CNull(R!DIE)) + Space(30), 22) + " "
    s = s + Left(to437(CNull(R!DIE)) + Space(30), 22) + " "
    
    s = s + Left(to437(CNull(R!pol)) + Space(30), 22) + " "
    
    
   
    s = s + Left(to437(CNull(R!AFM)) + Space(13), 11) + " "  'AFM
    
     s = s + Left(to437(CNull(R!doy)) + Space(30), 22) + " " 'DOY
    
    s = s + Right(Space(35) + Format(R!TYP, "####0.00"), 13) + " "
    
    s = s + Left("00.00", 5) + " "  ' EKPT
    
      s = s + Left("0000002000.00", 13) + " "  ' PLAFON
    
     s = s + Left(to437(CNull(R!THL)) + Space(25), 22) + " " 'THL
     
     s = s + Left(to437(CNull(R!ch4)) + Space(30), 22) + " " 'THL2
     
     
     
  
     s = s + Space(5) + " " 'KENA
     s = s + Space(1) + " " 'KENA
     
     
     
     
     s = s + "00000" 'ekpt
     
   
   
  
   Print #1, s
   R.MoveNext
Loop


Close #1
R.Close


MsgBox " –≈À¡‘≈” œ "

        
'------------------------------------√…œ— –≈À¡‘≈” ‘≈Àœ” -------------------------------------------------------------------





'
'    R.Open "SELECT *FROM PINAKES WHERE TYPOS=101 AND TIMH>0 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
'
'    Do While Not R.EOF
'        COUNTER = COUNTER + 1
'        PIN(COUNTER, 1) = R("PERIGRAFH")    '–≈ƒ…œ ‘œ’  ’—…œ’ ¡—◊≈…œ’ (SQLSERVER)
'        PIN(COUNTER, 2) = R("TIMH")    ' ¡—…»Ãœ” ”‘«À«” ‘œ’ EXCEL
'        R.MoveNext
'    Loop
'
'    R.Close
'
'    Dim a As String, k As Integer
'
'    Dim F As Integer
'
'    F = FreeFile
'    Open "C:\TEMP\PEL.ASC" For Output As #F
'    R.Open "SELECT *FROM PEL WHERE EIDOS='e' order by EPO", Gdb, adOpenDynamic, adLockOptimistic
'
'    Do While Not R.EOF
'        a = ""
'
'        For k = 1 To COUNTER
'
'            If InStr(PIN(k, 1), "@") = 0 Then '  ¡»¡—œ –≈ƒ…œ
'                If Len(PIN(k, 1)) > 0 Then
'
'                    ' ·ÒÈËÏÁÙÈÍÔ
'                    If R(PIN(k, 1)).Type = 7 Or R(PIN(k, 1)).Type = 5 Or R(PIN(k, 1)).Type = 131 Then     'IsNumeric(
'                        a = a + Right(Space(100) + Format(IIf(IsNull(R(PIN(k, 1))), 0, R(PIN(k, 1))), String(PIN(k, 2), "#") + ".00"), PIN(k, 2))
'                    Else
'                        ' string
'                        a = a + Left(IIf(IsNull(R(PIN(k, 1))), Space(100), R(PIN(k, 1))) + Space(100), PIN(k, 2))
'                    End If
'                End If
'
'            Else
'                a = a + mID(PIN(k, 1) + Space(100), 2, PIN(k, 2))
'            End If
'
'        Next
'
'        Print #F, a
'        'COUNTER = COUNTER + 1
'        R.MoveNext
'    Loop
'
'    R.Close
'
'    Close #F
'
'    'EIDH
'    COUNTER = 0
'    R.Open "SELECT *FROM PINAKES WHERE TYPOS=102 AND TIMH>0 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
'
'    Do While Not R.EOF
'        COUNTER = COUNTER + 1
'        PIN(COUNTER, 1) = R("PERIGRAFH")    '–≈ƒ…œ ‘œ’  ’—…œ’ ¡—◊≈…œ’ (SQLSERVER)
'        PIN(COUNTER, 2) = R("TIMH")    ' ¡—…»Ãœ” ”‘«À«” ‘œ’ EXCEL
'        R.MoveNext
'    Loop
'
'    R.Close
'
'    F = FreeFile
'    Open "C:\TEMP\APO.ASC" For Output As #F
'    R.Open "SELECT *FROM EID  order by ONO", Gdb, adOpenDynamic, adLockOptimistic
'
'    Do While Not R.EOF
'        a = ""
'
'        For k = 1 To COUNTER
'
'            If InStr(PIN(k, 1), "@") = 0 Then '  ¡»¡—œ –≈ƒ…œ
'                If Len(PIN(k, 1)) > 0 Then
'
'                    ' ·ÒÈËÏÁÙÈÍÔ
'                    If R(PIN(k, 1)).Type = 7 Or R(PIN(k, 1)).Type = 5 Or R(PIN(k, 1)).Type = 131 Then     'IsNumeric(
'                        a = a + Right(Space(100) + Format(IIf(IsNull(R(PIN(k, 1))), 0, R(PIN(k, 1))), String(PIN(k, 2), "#") + ".0000"), PIN(k, 2))
'                    Else
'                        ' string
'                        a = a + Left(IIf(IsNull(R(PIN(k, 1))), Space(100), R(PIN(k, 1))) + Space(100), PIN(k, 2))
'                    End If
'                End If
'
'            Else
'                a = a + mID(PIN(k, 1) + Space(100), 2, PIN(k, 2))
'            End If
'
'        Next
'
'        Print #F, a
'        'COUNTER = COUNTER + 1
'        R.MoveNext
'    Loop
'
'    R.Close
'
'    Close #F
'
'
    
'---------------------------------------- „ÈÔÒ·ÌÈ‰ÈÛ ÂÈ‰Á -------------------------------------------------------
'Dim R As New ADODB.Recordset
'Dim sql As String
sql = "SELECT FPA  as vat,  MON, KOD ,ONO AS NAME, "
sql = sql + "  ISNULL(LTI,0) AS PRICE,SYSKMAX "
sql = sql + " FROM EID where LEN(ISNULL(CH3,''))>0 "
'sql = sql + " LEFT OUTER JOIN HEMEASUREMENTUNITS u WITH (NOLOCK) ON i.HEAMSNTID = u.HEID"
'sql = sql + " inner join [HEVATCLASSES] v  on (I.[HEVTCLID] = v.[HEID])"


   R.Open sql, Gdb, adOpenDynamic, adLockOptimistic

'If IS_MERC = 1 Then
'     Dim R2 As New ADODB.Recordset
'End If

Open F_DIR_APOS_KARAM + "\PROD" For Output As #1
'Dim s As String

'Dim fpa As String

Do While Not R.EOF
   s = ""
   s = s + Left(R!kod + Space(15), 15) + " "
   s = s + Left(to437(R!Name) + Space(35), 35) + " "
   s = s + Left(R!kod + Space(35), 15) + " "
   
   s = s + Left(to437(R!MON) + Space(3), 3) + " "
    s = s + Left(to437(R!MON) + Space(3), 3) + " "
     s = s + Right$("     " + Format(nNull(R!SYSKMAX), "#0"), 5) + " " 'METATR
     
 s = s + Right(Space(35) + Format(R!Price, "####0.00"), 13) + " "
 
 
 
 
' If IS_MERC = 1 Then
'
' R2.Open "SELECT COUNT(*) FROM EID   WHERE   KOD='" + R!kod + "'", GMERC, adOpenDynamic, adLockOptimistic
'   If R2(0) = 0 Then
'          GMERC.Execute "INSERT INTO EID (KOD) VALUES ('" + R!kod + "')"
'   End If
'   R2.Close
'
'
'
'
'
'   GMERC.Execute "update EID set ONO='" + R!Name + "'  WHERE   KOD='" + R!kod + "'"
'
'   GMERC.Execute "update EID set MON='" + R!MON + "'  WHERE   KOD='" + R!kod + "'"
'   GMERC.Execute "update EID set LTI=" + Replace(Format(R!Price, "####0.00"), ",", ".") + "  WHERE   KOD='" + R!kod + "'"
'
'End If

 
 
 
 
 
 
     FPA = "    7"
     If R!VAT = 2 Then
        FPA = "   24" ' 24%
'        If IS_MERC = 1 Then
'            GMERC.Execute "update EID set FPA=2  WHERE   KOD='" + R!kod + "'"
'        End If
     ElseIf R!VAT = 1 Then
        FPA = "   13" '13%
'        If IS_MERC = 1 Then
'            GMERC.Execute "update EID set FPA=1  WHERE   KOD='" + R!kod + "'"
'        End If
'      ElseIf Val(mID(R!VAT, 2, 6)) = 3 Then
'       fpa = "    6"
      ElseIf R!VAT = 5 Then
        FPA = "    0"
     End If
     s = s + FPA + " "
     s = s + Space(13) + " " 'ypol
     s = s + Space(5) + " " 'ekpt
     s = s + Space(4)
     
 
'   S = S + Left(R!HENAME + Space(35), 35) + " "
'   S = S + Left(R!HENAME + Space(35), 35) + " "
'   S = S + Left(R!HENAME + Space(35), 35) + " "
'   S = S + Left(R!HENAME + Space(35), 35) + " "
'
   
      
   
   
   
 '  "  " + R("HECODE")
   Print #1, s
   R.MoveNext
Loop


Close #1

MsgBox "≈…ƒ« œ "




    
    
    
    
    
    
    
 
    'TIMOKATALOGOS
    COUNTER = 0
  R.Close
  R.Open "SELECT *FROM PINAKES WHERE TYPOS=103 AND TIMH>0 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        COUNTER = COUNTER + 1
        PIN(COUNTER, 1) = R("PERIGRAFH")    '–≈ƒ…œ ‘œ’  ’—…œ’ ¡—◊≈…œ’ (SQLSERVER)
        PIN(COUNTER, 2) = R("TIMH")    ' ¡—…»Ãœ” ”‘«À«” ‘œ’ EXCEL
        R.MoveNext
    Loop

    R.Close
 Dim F As Integer
 Dim a As String
    F = FreeFile
    Open F_DIR_APOS_KARAM + "\TIMOK.ASC" For Output As #F
    R.Open "SELECT ISNULL(EKPT,0) AS EKPT,isnull(TIMOKEID,'') as TIMOKEID,isnull(TIMOKPEL,'') as TIMOKPEL FROM TIMOKAT WHERE NOT(EKPT IS NULL) order by TIMOKPEL", Gdb, adOpenDynamic, adLockOptimistic

    Do While Not R.EOF
        a = ""
        ROM.Open "SELECT *FROM EID WHERE CH1='" + R!timokEID + "'  order by KOD", Gdb, adOpenDynamic, adLockOptimistic
        Do While Not ROM.EOF
           If R!EKPT > 0 Then
               Print #F, Left(R!timokPEL + "   ", 3) + Left(ROM!kod + "                  ", 15) + Right(Space(10) + Format(ROM!lti, "##0.00"), 10) + Right$(Space$(10) + Format(R!EKPT, "#####0.00"), 10)
           End If
           
           ROM.MoveNext
           
        Loop
        ROM.Close
        
        
'        For k = 1 To COUNTER
'
'            If InStr(PIN(k, 1), "@") = 0 Then '  ¡»¡—œ –≈ƒ…œ
'                If Len(PIN(k, 1)) > 0 Then
'
'                    ' ·ÒÈËÏÁÙÈÍÔ
'                    If R(PIN(k, 1)).Type = 7 Or R(PIN(k, 1)).Type = 5 Or R(PIN(k, 1)).Type = 131 Then     'IsNumeric(
'                        a = a + Right(Space(100) + Format(IIf(IsNull(R(PIN(k, 1))), 0, R(PIN(k, 1))), String(PIN(k, 2), "#") + ".0000"), PIN(k, 2))
'                    Else
'                        ' string
'                        a = a + Left(IIf(IsNull(R(PIN(k, 1))), Space(100), R(PIN(k, 1))) + Space(100), PIN(k, 2))
'                    End If
'                End If
'
'            Else
'                a = a + mID(PIN(k, 1) + Space(100), 2, PIN(k, 2))
'            End If
'
'        Next

        'Print #F, a
        'COUNTER = COUNTER + 1
        R.MoveNext
    Loop

    R.Close
 
    Close #F
 
    MsgBox "OK"




End Sub

Private Sub karam2_Click()
     
     APOKARAM 2
End Sub

Private Sub Label10_Click()

        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        'pinakas pedion
100     UPDATE2_PINAKES 103, Label10  ', Combo3

        Exit Sub

End Sub

Private Sub Label8_Click()

        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        'pinakas pedion
100     UPDATE2_PINAKES 101, Label8  ', Combo3

        Exit Sub

End Sub

Private Sub Label9_Click()

    Dim F As New DataEntry2, DD

    Dim R As New ADODB.Recordset
        
    'pinakas pedion
    UPDATE2_PINAKES 102, Label9
        
    Exit Sub

End Sub

Private Sub MATIM_LostFocus()
     PERIB.SetFocus
     
End Sub

Private Sub PERIB_LostFocus()

    Command5.SetFocus


End Sub

Private Sub SYNOLO_Click()

    POSO.Text = Val(Replace(SYNOLIKO.Text, ",", ".")) - Val(Replace(Text3.Text, ",", "."))

End Sub

Private Sub Text3_LostFocus()
   POSO.Text = Val(Replace(SYNOLIKO.Text, ",", ".")) - Val(Replace(Text3.Text, ",", "."))
End Sub
