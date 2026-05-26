VERSION 5.00
Object = "{CCA214C0-DFEB-4C91-9F0D-2665F77F6E23}#1.2#0"; "IDAutomationLinear.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form APOT1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   10605
   ClientLeft      =   -315
   ClientTop       =   135
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   ScaleHeight     =   10605
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   7080
      TabIndex        =   53
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "Εξοδος"
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CheckBox epanalhptikh 
      BackColor       =   &H00FF0000&
      Caption         =   "Επαναληπτική εγγραφή"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   6000
      TabIndex        =   51
      Top             =   5970
      Width           =   2355
   End
   Begin VB.TextBox CH1 
      Height          =   375
      Left            =   2040
      TabIndex        =   49
      Top             =   5280
      Width           =   1455
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   1680
      TabIndex        =   47
      Top             =   4920
      Width           =   1800
   End
   Begin VB.TextBox Text14 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   6945
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   2220
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Tελευταιος Κωδικός"
      Height          =   315
      Left            =   4800
      TabIndex        =   45
      Top             =   45
      Width           =   1695
   End
   Begin VB.PictureBox Standard 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5985
      Picture         =   "APOT1.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   44
      Top             =   7845
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox Skin1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5985
      Picture         =   "APOT1.frx":0F42
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   43
      Top             =   8145
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7950
      Picture         =   "APOT1.frx":149F
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   42
      Top             =   6870
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Εκτύπωση ετικετας"
      Height          =   375
      Left            =   6960
      TabIndex        =   40
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   1680
      TabIndex        =   11
      Top             =   3795
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      Left            =   5160
      TabIndex        =   12
      Top             =   3795
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   9135
      TabIndex        =   35
      Top             =   75
      Width           =   2070
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ΝΕΟ BARCODE"
      Height          =   285
      Left            =   9150
      TabIndex        =   34
      Top             =   705
      Width           =   2100
   End
   Begin VB.CommandButton Command11 
      Caption         =   "ΔΙΑΓΡΑΦΗ BARCODE"
      Height          =   285
      Left            =   9150
      TabIndex        =   33
      Top             =   1395
      Width           =   2100
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Height          =   300
      Index           =   13
      Left            =   1680
      TabIndex        =   4
      Top             =   1275
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   10
      Left            =   4680
      TabIndex        =   14
      Top             =   4200
      Width           =   1125
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   8160
      Top             =   7800
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
      Height          =   300
      Index           =   11
      Left            =   1680
      TabIndex        =   26
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   12
      Left            =   1680
      TabIndex        =   28
      Top             =   6400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   9
      Left            =   1680
      TabIndex        =   15
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   7
      Left            =   1680
      TabIndex        =   10
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   6
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Height          =   300
      Index           =   4
      Left            =   1680
      TabIndex        =   3
      Top             =   885
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3735
      TabIndex        =   9
      ToolTipText     =   "Aν γράψω με το χέρι την υποοικογένεια την εισάγει αυτόματα στον πίνακα με τις υποοικογένειες"
      Top             =   2880
      Width           =   3075
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2475
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2085
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Νέα Εγγραφή"
      Height          =   855
      Left            =   5520
      TabIndex        =   16
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   1680
      TabIndex        =   13
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   3815
   End
   Begin VB.TextBox Text1 
      DataField       =   " "
      Height          =   300
      Index           =   0
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   1
      Top             =   75
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8040
      Top             =   7320
      Visible         =   0   'False
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
   Begin VB.Label lblAUTO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "αυτ.αριθμ"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   3720
      TabIndex        =   54
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "......."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   0
      TabIndex        =   50
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κατηγορία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Index           =   26
      Left            =   0
      TabIndex        =   48
      Top             =   4920
      Width           =   1575
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode2 
      Height          =   1065
      Left            =   3090
      TabIndex        =   41
      Top             =   7800
      Visible         =   0   'False
      Width           =   1995
      _cx             =   3519
      _cy             =   1879
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
   Begin VB.Label FLAG 
      Height          =   300
      Left            =   720
      TabIndex        =   39
      Top             =   7995
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Τιμή Αγοράς"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   38
      Top             =   3795
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ποσ.% κέρδους"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   3840
      TabIndex        =   37
      Top             =   3795
      Width           =   1215
   End
   Begin VB.Label barcode 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6960
      TabIndex        =   36
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδ.Προμηθευτή"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   32
      Top             =   1275
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Υποοικογένεια"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   3765
      TabIndex        =   31
      Top             =   2580
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Τιμή Λιανική "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   3600
      TabIndex        =   30
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδ.Λογ.Πωλήσεων"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   19
      Left            =   120
      TabIndex        =   29
      Top             =   6400
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδ.Λογ.Αγορών"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   20
      Left            =   120
      TabIndex        =   27
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ελάχιστο στόκ"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   120
      TabIndex        =   25
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Προμηθευτής"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ράφι-Θέση"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   23
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Βοηθ.Κλειδί"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   885
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Χονδρική τιμή"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Οικογένεια"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κατηγορία ΦΠΑ"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   2475
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Μονάδα μέτρησης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   2085
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Περιγραφή είδους"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδικός"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1455
   End
End
Attribute VB_Name = "APOT1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim f_TIMOKAT    As Integer

Dim F_add_barc   As Integer

Dim f_autoNumber As Integer

Dim PROTH_FORA



' ΑΝ ΘΑ ΕΜΦΑΝΙΖΟΝΤΑΙ
Dim F_BOHU    As Integer

Dim F_KODPROM As Integer

Dim F_RAFI    As Integer

Dim F_KAT_FPA As Integer


Dim last_new As String
Dim f_LastNumber As String




Dim F_OIKO    As Integer

Dim F_YPOOIK  As Integer

Dim F_XTI     As Integer

Dim F_LTI     As Integer

Dim F_PROM    As Integer

Dim F_STOK    As Integer

Dim fe        As New FORMA

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Function get_next_EID(ARX As String) As String

        '<EhHeader>
        On Error GoTo get_next_EID_Err

        '</EhHeader>

        Dim DB As Database

        Dim R  As New ADODB.Recordset

        Dim k  As Long

        Dim x  As String
        Dim r2      As New ADODB.Recordset

100     r2.Open "SELECT KOD FROM EID WHERE KOD LIKE '" + ARX + "%' ORDER BY KOD ", Gdb, adOpenDynamic, adLockOptimistic

110     If r2.EOF Then
120         get_next_EID = ""
        Else
130         r2.MoveLast
140         get_next_EID = r2(0)
        End If

        Exit Function

        '  K = -1
        '
        '  Do While Not R.EOF
        '   x = Mid$(R(0), 7, 3)
        '     K = K + 1
        '     If Val(x) = Format(K, "000") Then
        '        R.MoveNext
        '     Else
        '        If K > 999 Then
        '           MsgBox "Εχει γεμίσει ο κωδικός"
        '           get_next_EID = "   "
        '        Else
        '           get_next_EID = ARX + Format(K, "000")
        '           R.Close
        '        End If
        '
        '        Exit Function
        '     End If
        '  Loop
        'End If

        '<EhFooter>
        Exit Function

get_next_EID_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.get_next_EID " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.get_next_EID " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

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

'Function Update_Next_Pel(meidos As String) As Integer
'
'    Dim db As Database
'    Dim R As Recordset
'
'
'End Function

Sub LAG2_ETIK()

        '========================etiketes ========================================================
        '<EhHeader>
        On Error GoTo LAG2_ETIK_Err

        '</EhHeader>
        Dim D1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        Dim mlabel As String

        Dim MPER As String

        Dim mtimes As Single

        Dim TIMM As Single

        Dim timologio As String

        'On Error Resume NextN  N

        Dim x As Printer

100     For Each x In Printers

110         If InStr(x.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
120             Set Printer = x

                ' Stop looking for a printer.
                Exit For

            End If

        Next

        On Error GoTo LAG2_ETIK_Err

130     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

140     MPER = Text1(1).Text

150     mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)

160     TIMM = Format(Val(Text1(10).Text), "#####.00")

        Dim F_ArPerEtik    '

170     F_ArPerEtik = 5

        '------------------------------------------------------------------------
180     For k = 1 To Val(mtimes)

190         If True Then
200             Printer.FontName = "128"
210             Printer.FontSize = 6
220             Printer.Print Tab(12); mlabel
230             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
240             Printer.FontSize = 8
250             Printer.Print
260             Printer.Print
270             Printer.Print
280             Printer.Print

290             Printer.Print Tab(F_ArPerEtik + 10); mlabel
300             Printer.Print Tab(F_ArPerEtik + 10); MPER
310             Printer.Print Tab(F_ArPerEtik + 10); timologio
320             Printer.FontSize = 14
330             Printer.FontBold = True
340             Printer.Print Tab(F_ArPerEtik + 10); Format(Val(Trim(LTrim(TIMM))), "00.00") + " €"
350             Printer.FontBold = False

360             Printer.EndDoc
            Else
370             Printer.Print Tab(F_ArPerEtik + 10); MPER
380             Printer.Print Tab(F_ArPerEtik + 20); mlabel
390             BarCode2.SymbologyID = 128
400             BarCode2.DataToEncode = Text1(4).Text
410             Printer.PaintPicture BarCode2.Picture, 348, 824
420             Printer.FontSize = 8
430             Printer.Print
440             Printer.Print
450             Printer.Print
460             Printer.Print
470             Printer.EndDoc
            End If

        Next

        Exit Sub

FIND_N_NUMBER:
        '************************
480     rec.MoveFirst
490     rec.Move 10

        'rec.Edit
500     If IsNull(rec("epel")) Then
510         rec("epel") = 1
        Else
520         rec("epel") = rec("epel") + 1
        End If

530     rec.Update
540     mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)
550     Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Text1(0).Text + "'"
560     Return

        '<EhFooter>
        Exit Sub

LAG2_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.LAG2_ETIK " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.LAG2_ETIK " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub LAG_ETIK()

        '========================etiketes ========================================================
        '<EhHeader>
        On Error GoTo LAG_ETIK_Err

        '</EhHeader>
        Dim D1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        Dim mlabel As String

        Dim MPER As String

        Dim mtimes As Single

        Dim TIMM As Single

        Dim timologio As String

        Dim F_ArPerEtik

        'On Error Resume NextN  N

        Dim x As Printer

100     For Each x In Printers

110         If InStr(x.DeviceName, "SATO") > 0 Then
                ' Set printer as system default.
120             Set Printer = x

                ' Stop looking for a printer.
                Exit For

            End If

        Next

        On Error GoTo LAG_ETIK_Err

130     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

        Dim r1 As New ADODB.Recordset

140     r1.Open "SELECT * FROM BARCODES WHERE KOD='" + Text1(0).Text + "'", Gdb, adOpenForwardOnly, adLockReadOnly

150     If r1.EOF Then
160         GoSub FIND_N_NUMBER
        Else

170         If Val(r1("ERG")) = 0 Then
180             GoSub FIND_N_NUMBER
            Else
190             mlabel = r1("erg")
            End If
        End If

200     MPER = Text1(2).Text

210     mtimes = InputBox("Πόσες ετικέτες;", MPER, 1)

220     TIMM = Format(Val(Text1(10).Text), "#####.00")

        'timologio = EGGTIM.Recordset("kode") + left(TIM.Recordset("atim"), 6) + "=" + EGGTIM.Recordset("pelkod") + Format(EGGTIM.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + Format(EGGTIM.Recordset("hme"), "dd/mm/yyyy")

230     For k = 1 To Val(mtimes)

240         If True Then
250             Printer.FontName = "128"
260             Printer.FontSize = 6
270             Printer.Print Tab(12); mlabel
280             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
290             Printer.FontSize = 8
300             Printer.Print
310             Printer.Print
320             Printer.Print
330             Printer.Print

340             Printer.Print Tab(F_ArPerEtik + 10); mlabel
350             Printer.Print Tab(F_ArPerEtik + 10); MPER
360             Printer.Print Tab(F_ArPerEtik + 10); timologio
370             Printer.FontSize = 14
380             Printer.FontBold = True
390             Printer.Print Tab(F_ArPerEtik + 10); Format(Val(Trim(LTrim(TIMM))), "00.00") + " €"
400             Printer.FontBold = False

410             Printer.EndDoc
            Else
420             Printer.Print Tab(F_ArPerEtik + 10); MPER
430             Printer.Print Tab(F_ArPerEtik + 20); mlabel
440             Printer.FontName = "ean13"    ' "128"
450             Printer.FontSize = 6
460             Printer.Print Tab(F_ArPerEtik + 8); Text1(0).Text    ' mlabel
470             Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
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
610     Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + Text1(0).Text + "'"
620     Return

        '<EhFooter>
        Exit Sub

LAG_ETIK_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.LAG_ETIK " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.LAG_ETIK " & "at line " & Erl

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
         "in ADOMERCNEW.APOT1.Combo1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Combo1_KeyPress " & "at line " & Erl

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
         "in ADOMERCNEW.APOT1.Combo2_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Combo2_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo3_GotFocus()
    ' Call SendMessage(Combo3.hWnd, CB_SHOWDROPDOWN, 1, 0)
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
         "in ADOMERCNEW.APOT1.Combo3_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Combo3_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo4_GotFocus()
    '  Call SendMessage(Combo4.hWnd, CB_SHOWDROPDOWN, 1, 0)
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Combo4_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Combo4_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Combo4_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Combo4_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo4_LostFocus()

        ' YPOOIKOGENEIES
        '<EhHeader>
        On Error GoTo Combo4_LostFocus_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

        ' Exit Sub
100     If Len(Combo4.Text) > 0 Then
            ' BALE_TSONTES "3", Left(Combo4.Text, 2)
110         R.Open "SELECT *FROM PINAKES WHERE TYPOS=3 AND TIMH=" + Left(Combo4.Text, 2) + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
120         Combo3.Clear

130         Do While Not R.EOF

140             If R("typos") = 3 Then
150                 Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
                Else

                    Exit Do

                End If

160             R.MoveNext
            Loop

170         R.Close

        End If

        '<EhFooter>
        Exit Sub

Combo4_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Combo4_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Combo4_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command10_Click()

        '<EhHeader>
        On Error GoTo Command10_Click_Err

        '</EhHeader>
        On Error Resume Next

        Dim D


         List1.AddItem Text3.Text




100   '  Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + Text1(0).Text + "','" + Text3.Text + "')", D
      '  Adodc2.ConnectionString = gConnect
      '  Adodc2.RecordSource = "select * from BARCODES WHERE KOD='" + Text1(0).Text + "'"
        
110     ' Adodc2.Refresh

120     'If D = 0 Then
130      '   MsgBox "Δεν καταχωρήθηκε"
       ' End If

       ' DataGrid1.Refresh
        


        '<EhFooter>
        Exit Sub

Command10_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Command10_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Command10_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command11_Click()

        '<EhHeader>
        On Error GoTo Command11_Click_Err



         List1.Clear

        '</EhHeader>

100    ' Gdb.Execute "DELETE  FROM BARCODES WHERE ERG='" + Text3.Text + "' AND KOD='" + Text1(0).Text + "'"
110    ' Adodc2.Refresh

        '<EhFooter>
        Exit Sub

Command11_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Command11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Command11_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim ANS

100     If Command2.Caption = "Νέα Εγγραφή" Then



        Else
110         ANS = MsgBox("Να χαθουν οι αλλαγές", vbYesNo)

120         If ANS = vbNo Then
                   Exit Sub
            Else
            
'                 'If DataGrid1.ROWS.Count > 0 Then
'                 Dim R As New ADODB.Recordset
'                 R.Open "SELECT COUNT(*) FROM EID WHERE KOD='" + Text1(0).Text + "'", Gdb, adOpenDynamic
'                 IF R(0)
'
            
            
            
            
            
            
            End If
            

                   
        End If

130     MDIForm1.StatusBar1.Panels(6).Text = Text1(0).Text

140     Unload Me

150     If gApoMenu = True Then
160         SendKeys "%"
170         SendKeys "{RIGHT}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim k      As Integer, DUM As Integer, DB As Database, R As New ADODB.Recordset
Dim n11 As Long
        Dim x      As String

        Dim mlabel As String

        Dim SQL3   As String

        Dim RECS   As Long

100     If Command2.Caption = "Νέα Εγγραφή" Then

110         Command2.Caption = "Ενημέρωση"
120         Command2.Refresh

' ερχεται απο par1
If APOT1.flag.Caption = "1" Then

Else
           n11 = 1 + gVal(f_LastNumber)
           Text1(0).Text = Format(n11, Left$("0000000000", Len(f_LastNumber)))
           lblAUTO.Visible = True
           
        
End If



130         For k = 1 To 3
140             Text1(k).Enabled = True
150             Text1(k).ForeColor = vbBlack
            Next

160         For k = 6 To 7
170             Text1(k).Enabled = True
180             Text1(k).ForeColor = vbBlack
            Next

190         For k = 11 To 13
200             Text1(k).Enabled = True
210             Text1(k).ForeColor = vbBlack
            Next

220         Text1(9).Enabled = True
230         Text1(5).Enabled = True
240         Text1(5).ForeColor = vbBlack

250         Combo1.Enabled = True
260         Combo2.Enabled = True
270         Combo3.Enabled = True

            If F_KAT_FPA - 1 < 0 Then F_KAT_FPA = 1

280         Combo2.Text = Combo2.List(F_KAT_FPA - 1)
290         Text1(0).SetFocus

        Else

            'ελεγχος μην ξαναυπάρχει ο κωδικός
            'Set db = OpenDatabase(gDir, False, False, gConnect)
300         R.Open "select count(*) as ar from EID where  KOD='" + Text1(0).Text + "';", Gdb, adOpenDynamic, adLockOptimistic


            If Len(Trim$(Text1(0).Text)) = 0 Then
                 MsgBox "Δεν επιτρέπεται μηδενικός κωδικός"
                 Exit Sub
            End If




310         If R("ar") > 0 And Len(Text1(0).Text) > 0 Then
320             MsgBox "Υπάρχει ο κωδικός"

                Exit Sub

            End If

330         If Val(Left(Combo2.Text, 2)) = 0 Then
340             MsgBox "ΦΠΑ ;"

                Exit Sub

            End If



' ερχεται απο par1
 If APOT1.flag.Caption = "1" Then

 Else
           
        


            ' ενημέρωση μετρητή
              n11 = 1 + gVal(f_LastNumber)
       f_LastNumber = Format(n11, Left$("0000000000", Len(f_LastNumber)))
            Gdb.Execute ("UPDATE PARAMETROI SET TIMH='" + f_LastNumber + "' WHERE  FORMA='APOT1' AND VAR='f_LastNumber'")
  End If
            

            '   dum = Update_Next_Pel(x)
            On Error GoTo NoUpdate

350         data1.Recordset.AddNew

            On Error Resume Next

360         data1.Recordset("kod") = UCase(Text1(0))  ', "0000")
370         data1.Recordset("ONO") = Left(Text1(1).Text, data1.Recordset("ONO").DefinedSize)

380         data1.Recordset("FPA") = Val(Left(Combo2.Text, 2))

390         data1.Recordset("mon") = Combo1.Text
400         data1.Recordset("lastupd") = gEnarjh '


            '       DATA2.Recordset("kodlog") = Val(left(Combo3.Text, 3)) ' Combo3.ListIndex  'Right(Combo3.Text, 3)

410         data1.Recordset("aeg") = Val(Left(Combo4.Text, 3))    ' Combo3.ListIndex  'Right(Combo3.Text, 3)
            'Data1.Recordset("kodlog") = Right(Combo3.Text, 3)

            '       'αν εχω προσθέσει υπο οικογένεια την καταχωρεί αυτόματα
            '       If Val(Combo3.Text) = 0 Then
            '           SQL3 = "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (3," + Str(Combo3.ListCount + 1) + ",'" + Combo3.Text + "'," + Left(Combo4.Text, 2) + ")"
            '           Gdb.Execute SQL3
            '       End If

420         data1.Recordset("kodlog") = Val(Left(Combo3.Text, 3))

430         data1.Recordset("xti") = Val(Replace(Text1(2).Text, ",", "."))
440         data1.Recordset("lti") = Val(Replace(Text1(3).Text, ",", "."))    'Val(Text1(3).Text)
450         data1.Recordset("lti5") = Val(Replace(Text1(10).Text, ",", "."))    ' Val(Text1(10).Text)
460         data1.Recordset("pos_kerD") = Val(Replace(Text1(5).Text, ",", "."))    ' Val(Text1(5).Text)

470         data1.Recordset("prom") = Text1(7).Text
480         data1.Recordset("spa") = Val(Text1(9).Text)
            ' DATA1.Recordset("apos") = Val(Text1(8).Text)
490         data1.Recordset("ues") = Text1(6).Text

500         data1.Recordset("kodsynod") = Text1(12).Text
510         data1.Recordset("kodlogag") = Text1(11).Text

520         data1.Recordset("CH1") = CH1.Text

530         data1.Recordset("ERG") = Text1(4).Text    'ΕΡΓ
540         data1.Recordset("KODERG") = Text1(13).Text    ' ΚΩΔ.ΠΡΟΜΗΘΕΥΤΗ
550         data1.Recordset("MEMO") = Text14.Text
560         data1.Recordset("KATHGORIA") = Val(Left(Combo5.Text, 3))    ' Combo3.ListIndex  'Right(Combo3.Text, 3)

            On Error GoTo NoUpdate

570         data1.Recordset.Update

Dim FLAG_BARC As Integer
FLAG_BARC = 0
Dim NCC As Integer
On Error Resume Next
Dim ll As Integer
For ll = 0 To List1.ListCount - 1
     NCC = 0
     Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + Text1(0).Text + "','" + List1.List(ll) + "')", NCC
     If NCC = 0 Then
         
         MsgBox "ΔΕΝ ΚΑΤΕΧΩΡΗΘΗ Ο ΚΩΔΙΚΟΣ " + List1.List(ll)
     
     
     End If
     
     FLAG_BARC = 1
Next


List1.Clear

Text3.Text = ""





            On Error Resume Next
            
            
' ΔΕΝ ΠΕΡΑΣΑ BARCODES OΠΟΤΕ ΔΗΜΙΟΥΡΓΕΙ ΑΠΟ ΜΟΝΟΣ ΤΟΥ BARCODE
If FLAG_BARC = 0 Then
            
            

580         If F_add_barc = 1 Then   ' file("C:\LAGEURO\MPOYG.TXT") Or
                'ΤΟ ΒΑRCODE ΕΙΝΑΙ Ο ΚΩΔΙΚΟΣ
590             Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + UCase(Text1(0).Text) + "','" + UCase(Text1(0).Text) + "')"
600         ElseIf F_add_barc = 2 Then

                'ΤΟ ΒΑRCODE ΕΙΝΑΙ TO BOHΘ.ΚΛΕΙΔΙ
610             If Val(Text1(4).Text) = 0 Then
620                 GoSub FIND_N_NUMBER
                End If

630             RECS = 0

640             Do While RECS = 0

650                 Gdb.Execute "insert into BARCODES (KOD,ERG) VALUES ('" + UCase(Text1(0).Text) + "','" + UCase(Text1(4).Text) + "')", RECS

660                 If RECS = 0 Then GoSub FIND_N_NUMBER
                Loop

            End If


End If





670         f_TIMOKAT = Val(FINDPARAMETROI(1, "PELAT2", "F_TIMOKAT", "0", "Aριθμός Τιμοκαταλόγων"))

680         If f_TIMOKAT > 0 Then

690             For k = 1 To f_TIMOKAT
700                 Gdb.Execute "INSERT INTO TIMOKAT SELECT KOD,0 AS EKPT, " + str(k) + " AS TIMOK,ONO FROM EID WHERE KOD='" + UCase(Text1(0)) + "'"
                Next

            End If

710         If flag.Caption = "1" And epanalhptikh.Value = vbUnchecked Then    ' ΕΡΧΕΤΑΙ ΑΠΟ ΤΑ ΤΙΜΟΛΟΓΙΑ
720             MDIForm1.StatusBar1.Panels(6).Text = Text1(0).Text
730             Unload Me
            End If

740         MHNYMA2.Timer1.Interval = 1000
750         MHNYMA2.Label2.Caption = "Η εγγραφή " + Text1(1).Text + " αποθηκεύτηκε με επιτυχία"
760         MHNYMA2.SHOW 1

770         Command2.Caption = "Νέα Εγγραφή"
        End If

        ' Command2.SetFocus

        Exit Sub

NoUpdate:
780     MsgBox "Δεν έγινε η αποθήκευση" + Chr(13) + Err.Description


790     Resume Next

        Exit Sub

FIND_N_NUMBER:

        '************************
        Dim rec  As ADODB.Recordset

        'rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
        Dim REC2 As New ADODB.Recordset

800     REC2.Open "SELECT * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
810     REC2.MoveFirst
820     REC2.Move 10

        'rec.Edit
830     If IsNull(REC2("epel")) Then
840         REC2("epel") = 1
        Else
850         REC2("epel") = REC2("epel") + 1
        End If

860     REC2.Update

870     Text1(4).Text = "5555" + Right("00000000" + LTrim(str(REC2("EPEL"))), 8) + ChDigEAN13("5555" + Right("00000000" + LTrim(str(REC2("EPEL"))), 8))
880     Gdb.Execute "UPDATE EID SET ENERGO=1,ERG='" + Text1(4).Text + "' WHERE KOD='" + Text1(0).Text + "'"
890     Return

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        'APOT2.Command7_Click
        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>

100     If Len(Dir("C:\LAGEURO\LAGAKIS.TXT", vbNormal)) > 0 Then
110         LAG_ETIK

            Exit Sub

        End If

120     If Len(Dir("C:\LAGEURO\LAGAKIS2.TXT", vbNormal)) > 0 Then
130         LAG2_ETIK

            Exit Sub

        End If

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()






        'AYTOMATH AΡΙΘΜΟΔΟΤΗΣΗ
        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>
        Dim x As String

        '   If Len(Text1(0).Text) = 6 And Mid$(Text1(0).Text, 3, 1) = "-" And Mid$(Text1(0).Text, 6, 1) = "-" Then
100     x = Text1(0).Text
110     Text1(0).Text = get_next_EID(x)
        '   Else
        '     If Len(Text1(0).Text) = 0 Then
        '      Else
        '        Text1(0).Text = "" 'get_next_EID(X)
        '        Text1(0).SetFocus
        '     End If

        ' End If

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Activate()

        '<EhHeader>
        On Error GoTo Form_Activate_Err

        '</EhHeader>

        Dim DB, R As New ADODB.Recordset

100     If PROTH_FORA = False Then Exit Sub
110     PROTH_FORA = False

        'BALE_TSONTES "11", "0"

        'Set db = OpenDatabase(gDir, False, False, gConnect)
120     R.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

        'R.Index = "pinakes"

        'FPA
        'R.Seek "=", 1, 1
130     Do While Not R.EOF

140         If R("typos") = 1 Then
150             Combo2.AddItem str(R("AYJON")) + " -> " + str(R("TIMH"))
            End If

160         R.MoveNext
        Loop

        ' mon.metrhshs
        'R.Seek "=", 2, 1

170     R.Close
180     R.Open "SELECT *FROM PINAKES WHERE TYPOS=2 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

190     Do While Not R.EOF

200         If R("typos") = 2 Then
210             Combo1.AddItem R("PERIGRAFH")
            Else

                Exit Do

            End If

220         R.MoveNext
        Loop

230     Combo1.ListIndex = 0

        ' kathgories
        ' R.Seek "=", 3, 1
240     R.Close
250     R.Open "SELECT TOP 5 *FROM PINAKES WHERE TYPOS=3 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

260     Do While Not R.EOF

270         If R("typos") = 3 Then
280             Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
                'Combo3.AddItem left(R("PERIGRAFH") + Space(20), 20) + Format(R("ayjon"), _
                 "000")
            Else

                Exit Do

            End If

290         R.MoveNext
        Loop

300     R.Close

        ' ΚΑΤΗΓΟΡΙΕΣ

310     R.Open "SELECT *FROM PINAKES WHERE TYPOS=15 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
        'R.MoveFirst

320     If R.EOF Then
330         Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH) VALUES (15,1,'ΕΜΠΟΡΕΥΜΑ')"
        End If

340     Do While Not R.EOF

350         If R("typos") = 15 Then
360             Combo5.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
            Else

                Exit Do

            End If

370         R.MoveNext
        Loop

380     R.Close

        On Error Resume Next

        ' OIKOGENEIES

390     R.Open "SELECT *FROM PINAKES WHERE TYPOS=11 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
400     R.MoveFirst

410     Do While Not R.EOF
420         Combo4.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
430         R.MoveNext
        Loop

440     R.Close

        On Error GoTo Form_Activate_Err

450     F_add_barc = Val(FINDPARAMETROI(1, "PAR1", "add_barc", "1", "Γ.16 Barcode έιναι 1.Κωδ 2.Βοηθ.Κλειδί"))

        'Dim f_autoNumber As Integer

460     f_autoNumber = Val(FINDPARAMETROI(1, "PELAT2", "F_autoNumber", "0", "Αρίθμηση αυτόματη 00-00-000  =1 Οχι=0"))
        f_LastNumber = (FINDPARAMETROI(1, "APOT1", "f_LastNumber", "00000", "ΑΥΤΟΜΑΤΗ ΑΡΙΘΜΗΣΗ ΑΡΧΙΚΟ ΚΑΙ ΜΟΡΦΗ Π.Χ. 00000 "))
        
        f_LastNumber = Trim(f_LastNumber)
        
        ' ΑΝ ΘΑ ΕΜΦΑΝΙΖΟΝΤΑΙ
470     F_BOHU = Val(FINDPARAMETROI(1, "APOT1", "F_BOHU", "1", "Εμφανίζεται το βοηθητικο κλειδί=1"))
480     F_KODPROM = Val(FINDPARAMETROI(1, "APOT1", "F_KODPROM", "1", "Εμφανίζεται ο ΚΩΔ.ΠΡΟΜΗΘΕΥΤΉ=1"))
490     F_RAFI = Val(FINDPARAMETROI(1, "APOT1", "F_RAFI", "1", "Εμφανίζεται το ΡΑΦΙ=1"))
500     F_OIKO = Val(FINDPARAMETROI(1, "APOT1", "F_OIKO", "1", "Εμφανίζεται ΟΙΚΟΓΕΝΕΙΑ=1"))
510     F_YPOOIK = Val(FINDPARAMETROI(1, "APOT1", "F_YPOOIK", "1", "Εμφανίζεται ΥΠΟΟΙΚΟΓΕΝΕΙΑ=1"))
520     F_XTI = Val(FINDPARAMETROI(1, "APOT1", "F_XTI", "1", "Εμφανίζεται Η ΧΟΝΔΡΙΚΗ ΤΙΜΗ=1"))
530     F_LTI = Val(FINDPARAMETROI(1, "APOT1", "F_LTI", "1", "Εμφανίζεται Η ΛΙΑΝΙΚΗ ΤΙΜΗ=1"))
540     F_PROM = Val(FINDPARAMETROI(1, "APOT1", "F_PROM", "1", "Εμφανίζεται Ο ΠΡΟΜΗΘΕΥΤΗΣ=1"))
550     F_STOK = Val(FINDPARAMETROI(1, "APOT1", "F_STOK", "1", "Εμφανίζεται ΤΟ ΕΛΑΧΙΣΤΟ ΣΤΟΚ=1"))

        
        F_KAT_FPA = Val(FINDPARAMETROI(1, "APOT1", "F_KAT_FPA", "6", "Προεπιλεγμένη κατηγ.ΦΠΑ"))


560     If F_BOHU = 0 Then Text1(4).Visible = False
570     If F_KODPROM = 0 Then Text1(13).Visible = False
580     If F_RAFI = 0 Then Text1(6).Visible = False
590     If F_XTI = 0 Then Text1(3).Visible = False
600     If F_LTI = 0 Then Text1(10).Visible = False
610     If F_PROM = 0 Then Text1(7).Visible = False
620     If F_STOK = 0 Then Text1(9).Visible = False

630     If F_OIKO = 0 Then Combo4.Visible = False
640     If F_YPOOIK = 0 Then Combo3.Visible = False

650     data1.ConnectionString = gConnect   ' "dBase IV;"

660     data1.RecordSource = "SELECT TOP 10 *FROM EID"
670     data1.Refresh

680     Command2_Click
        ' Me.Refresh

        'Me.width = 15000
        'Me.height = 15000

        '  Command2.Caption = "Ενημέρωση"
        '  Text1(0).SetFocus

        ' Command2.SetFocus

        '<EhFooter>
        Exit Sub

Form_Activate_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Form_Activate " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Form_Activate " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "APOT1"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Form_DblClick " & "at line " & Erl

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
            'If Command1. Then   KeyPreview = True
120         Command1_Click
            'Else
            '   Command1.SetFocus
            'End If

        End If

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '   Text1(0).Text = get_next_pel("e")
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     Me.Picture = LoadPicture(gPicture)
110     KeyPreview = True
120     PROTH_FORA = True
        '    'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        '
        '    'SkinFramework1.ApplyWindow Me.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
        '
        '

        ' mForm_Load Me, fh, fw, ft, fl
        'fe.mForm_Load Me
130     fe.mForm_Load APOT1
        last_new = ""
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    ' ''''''''''''''''''''''''''''' mForm_Resize Me, 8, 15240, 10605, fh, fw, ft, fl
    '  fe.mForm_Resize Me, 15240, 10605   ', fh, fw, ft, fl
End Sub

Private Sub Label1_Click(index As Integer)

        'OIKOGENEIES,YPOOIKOGENEIES,MON.METRHSHS
        '<EhHeader>
        On Error GoTo Label1_Click_Err

        '</EhHeader>
        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        '-----------------------------------------------------------------------------------------
100     If index = 2 Then    'ΜΟΝΑΔΕΣ ΜΕΤΡΗΣΗΣ
            'sql οπου θα γίνεται το update ή το addnew
            '  f.SHOW
110         Load F

120         Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=2"
130         F.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=2 "    '    sql
            'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
140         F.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " & " FROM PINAKES WHERE TYPOS=2 ORDER BY AYJON "    'ParamGrid   sql2

            'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
150         F.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=2 AND "    'ParamGrid              DelSQL

160         F.ParamGrid.TextMatrix(0, 6) = "TYPOS"    'EPIPLEON FIELD
170         F.ParamGrid.TextMatrix(0, 7) = 2    ' TIMH TOY EPIPLEON FIELD

            'ποιό update query θα δημιουργείται (προαιρετικά )
            'αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
            '                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
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

180         F.ParamGrid.TextMatrix(0, 4) = ""  'DD
190         F.ParamGrid.TextMatrix(1, 1) = "A/A"
200         F.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "
            'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
210         F.Labels(0).Width = F.Labels(0).Width + 1000    '* TextWidth("A")
220         F.Textboxes(0).Left = F.Textboxes(0).Left + 1000    ' * TextWidth("A")
230         F.Left = Label1(index).Left
240         F.Top = Label1(index).Top + MDIForm1.Top

250         F.SHOW 1
260         Combo1.Clear
270         R.Open "SELECT *FROM PINAKES WHERE TYPOS=2 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

280         Do While Not R.EOF

290             If R("typos") = 2 Then
300                 Combo1.AddItem R("PERIGRAFH")
                Else

                    Exit Do

                End If

310             R.MoveNext
            Loop

320         R.Close

        End If

        '-----------------------------------------------------------------------------------------
330     If index = 4 Then    'OIKOGENEIES
            'sql οπου θα γίνεται το update ή το addnew
            '  f.SHOW
340         Load F

350         Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=11"
360         F.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=11 "    '    sql
            'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
370         F.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " & " FROM PINAKES WHERE TYPOS=11 ORDER BY AYJON "    'ParamGrid   sql2

            'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
380         F.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=11 AND "    'ParamGrid              DelSQL

            'ποιό update query θα δημιουργείται (προαιρετικά )
            'αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
            '                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL

390         F.ParamGrid.TextMatrix(0, 6) = "TYPOS"    'EPIPLEON FIELD
400         F.ParamGrid.TextMatrix(0, 7) = 11    ' TIMH TOY EPIPLEON FIELD

            'DD = " UPDATE PINAKES SET TYPOS=11 where TYPOS IS NULL AND AYJON="" + t1   "
            ' AddNewScript
            ' DD = "Sub Main()" & vbCrLf & _
            '       " DIM mCon  " & vbCrLf & _
            '      " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
            '     " mCon.Open """ + gConnect + " "" " & vbCrLf & _
            '    " mCon.execute """ + DD + " " & vbCrLf & _
            '  "Set r=Nothing" & vbCrLf & _
            ' "Set mCon=Nothing" & vbCrLf & _
            '     "End Sub"
410         F.ParamGrid.TextMatrix(0, 4) = ""   'DD
420         F.ParamGrid.TextMatrix(1, 1) = "A/A"
430         F.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "
            'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
440         F.Labels(0).Width = F.Labels(0).Width + 1000    '* TextWidth("A")
450         F.Textboxes(0).Left = F.Textboxes(0).Left + 1000    ' * TextWidth("A")
            'f.Left = 0
            'f.Top = 0
            'f.height = 18000

460         F.Left = Label1(index).Left
470         F.Top = Label1(index).Top + MDIForm1.Top

480         F.SHOW 1
490         Combo4.Clear
500         R.Open "SELECT *FROM PINAKES WHERE TYPOS=11 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

510         Do While Not R.EOF
520             Combo4.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
530             R.MoveNext
            Loop

540         R.Close

            ' Me.Top = -200
            ' Me.Left = 0
            ' Me.width = 17000
            ' Me.height = 17000
            '
            '
            ' Me.SHOW

        End If

        '----------------------------------------------------------------------------------------
550     If index = 14 Then    'YPO - OIKOGENEIES
560         If Val(Left(Combo4.Text, 2)) = 0 Then

                Exit Sub

            End If

            'sql οπου θα γίνεται το update ή το addnew
570         F.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=3 AND TIMH= " + Left(Combo4.Text, 2)    '    sql
            'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
580         F.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " & " FROM PINAKES WHERE TYPOS=3 AND TIMH=" + Left(Combo4.Text, 2) + " ORDER BY AYJON "    'ParamGrid   sql2

            'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
590         F.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=3  AND TIMH=" + Left(Combo4.Text, 2) + " AND "    'ParamGrid              DelSQL

600         F.ParamGrid.TextMatrix(0, 6) = "TYPOS"    'EPIPLEON FIELD
610         F.ParamGrid.TextMatrix(0, 7) = 3    ' TIMH TOY EPIPLEON FIELD

620         F.ParamGrid.TextMatrix(0, 8) = "TIMH"    '2 EPIPLEON FIELD
630         F.ParamGrid.TextMatrix(0, 9) = Left(Combo4.Text, 2)  ' TIMH TOY 2 EPIPLEON FIELD

            'ποιό update query θα δημιουργείται (προαιρετικά )
            'αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
            '                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
            ' DD = " UPDATE PINAKES SET TYPOS=3,TIMH=" + Left(Combo4.Text, 2) + " where TYPOS IS NULL AND AYJON="" + t1   "
            '' AddNewScript
            '  DD = "Sub Main()" & vbCrLf & _
            '         " DIM mCon  " & vbCrLf & _
            '         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
            '         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
            '         " MSGBOX   t1  " & vbCrLf & _
            '         " mCon.execute """ + DD + " " & vbCrLf & _
            '        "Set r=Nothing" & vbCrLf & _
            '        "Set mCon=Nothing" & vbCrLf & _
            '        "End Sub"

640         F.ParamGrid.TextMatrix(0, 4) = ""  ' DD
650         F.ParamGrid.TextMatrix(1, 1) = "A/A"
660         F.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "

            'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
670         F.Labels(0).Width = F.Labels(0).Width + 1000    '* TextWidth("A")
680         F.Textboxes(0).Left = F.Textboxes(0).Left + 1000    ' * TextWidth("A")
            'f.Left = 0
            'f.Top = 0
            'f.height = 18000

690         F.Left = Label1(index).Left
700         F.Top = Label1(index).Top + MDIForm1.Top

710         F.SHOW 1

720         Combo3.Clear
730         R.Open "SELECT * FROM PINAKES WHERE TYPOS=3 AND TIMH= " + Left(Combo4.Text, 2) + " ORDER BY AYJON ", Gdb, adOpenDynamic, adLockOptimistic

740         Do While Not R.EOF
750             Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
760             R.MoveNext
            Loop

770         R.Close

        End If

        '  Me.WindowState = 2

        '<EhFooter>
        Exit Sub

Label1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Label1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Label1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Option1_Click(Index As Integer)
'
'End Sub

Private Sub Text1_Change(index As Integer)

        '<EhHeader>
        On Error GoTo Text1_Change_Err

        '</EhHeader>
100     If index = 5 Then
110         Text1(3).Text = gVal(Text1(2).Text) * (100 + gVal(Text1(5).Text)) / 100
120         Text1(10).Text = gVal(Text1(3).Text) * (1 + g_Fpa(Val(Left(Combo2.Text, 2))) / 100)
          
        End If

130     If index = 3 Then
140         Text1(10).Text = gVal(Text1(3).Text) * (1 + g_Fpa(Val(Left(Combo2.Text, 2))) / 100)
        End If

        '<EhFooter>
        Exit Sub

Text1_Change_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Text1_Change " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Text1_Change " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_GotFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text1_GotFocus_Err

        '</EhHeader>

100     Text1(index).BackColor = vbYellow

110     Text1(index).SelStart = 0
120     Text1(index).SelLength = Len(Text1(index).Text)

        '<EhFooter>
        Exit Sub

Text1_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Text1_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Text1_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Text1_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then

110         KeyAscii = 0    'suppress the beep

120         If index = 9 Then
130             Command2.SetFocus
            Else
140             keybd_event VK_TAB, 0, 0, 0    'send a tab
            End If
        End If

        '<EhFooter>
        Exit Sub

Text1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Text1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Text1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_LostFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text1_LostFocus_Err

        '</EhHeader>
        Dim x As String

100     Text1(index).BackColor = vbWhite

110     If F_add_barc = 1 Then   ' file("C:\LAGEURO\MPOYG.TXT") Or
120         barcode.Caption = Text1(0).Text
        End If

        Dim R As New ADODB.Recordset

130     If index = 0 Then

            If InStr(Text1(0).Text, ",") > 0 Then
             
             ' MsgBox "Δεν επιτρέπεται το κομμα (,) στον κωδικό.Εγινε αλλαγή σε τελεία(.)"
              'Text1(0).Text = Replace(Text1(0).Text, ",", ".")
              
            End If
            


140         R.Open "select count(*) as ar from EID where  KOD='" + Text1(0).Text + "';", Gdb, adOpenDynamic, adLockOptimistic

150         If R("ar") > 0 And Len(Trim(Text1(0).Text)) > 0 Then
160             MsgBox "Υπάρχει ο κωδικός"
                If last_new = Text1(0).Text Then
                   Text1(0).Text = ""
                   ' Exit Sub
                End If

170             If Len(Trim(Text1(0).Text)) > 0 Then

                    On Error Resume Next

180                 Text1(0).SetFocus
                End If
                last_new = Text1(0).Text
                Exit Sub

            End If
        End If

        '<EhFooter>
        Exit Sub

Text1_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT1.Text1_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT1.Text1_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

