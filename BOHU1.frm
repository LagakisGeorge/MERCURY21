VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BOHU1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Πίνακες"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10470
   ScaleWidth      =   13170
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "ΔΙΑΓΡΑΦΗ ΧΡΗΣΗΣ"
      Height          =   375
      Left            =   7228
      TabIndex        =   15
      Top             =   6786
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Εξοδος"
      Height          =   375
      Left            =   7254
      TabIndex        =   12
      Top             =   7371
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   375
      Left            =   6840
      Top             =   9120
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSDataGridLib.DataGrid DBGRid2 
      Bindings        =   "BOHU1.frx":0000
      Height          =   3105
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   5477
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "AYJON"
         Caption         =   "Α/Α"
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
         DataField       =   "PERIGRAFH"
         Caption         =   "ΠΕΡΙΓΡΑΦΗ"
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
      BeginProperty Column02 
         DataField       =   "TIMH"
         Caption         =   "TIMH"
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
      BeginProperty Column03 
         DataField       =   "N1"
         Caption         =   "MYDATA"
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
      BeginProperty Column04 
         DataField       =   "C1"
         Caption         =   "C1-ΕΣΟΔΑ CATEGORY"
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
      BeginProperty Column05 
         DataField       =   "LOG1"
         Caption         =   "ΕΞΟΔΑ"
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
      BeginProperty Column06 
         DataField       =   "ID"
         Caption         =   "ID"
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
      BeginProperty Column07 
         DataField       =   "C2"
         Caption         =   "UBL-C2"
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
      BeginProperty Column08 
         DataField       =   "C3"
         Caption         =   "Ε3_ΕΣΟΔΑ-C3"
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
      BeginProperty Column09 
         DataField       =   "N2"
         Caption         =   "N2"
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
      BeginProperty Column10 
         DataField       =   "N3"
         Caption         =   "N3"
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
      BeginProperty Column11 
         DataField       =   "C4"
         Caption         =   "C4"
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
      BeginProperty Column12 
         DataField       =   "N4"
         Caption         =   "N4"
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
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   375
      Left            =   1200
      Top             =   9120
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   375
      Left            =   -120
      Top             =   9120
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      DataSourceName  =   "emp3"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Δημιουργία νέας χρήσης"
      Height          =   375
      Left            =   7228
      TabIndex        =   5
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Παράμετροι Εκτυπώσεων"
      Height          =   375
      Left            =   7228
      TabIndex        =   2
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Κλείνουν ανοιχτές φόρμες-Χρήση νέας εταιρείας"
      Height          =   255
      Left            =   7397
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   4199
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "BOHU1.frx":0014
      Left            =   345
      List            =   "BOHU1.frx":0016
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Πίνακες ΦΠΑ,Κατηγοριών"
      Height          =   8295
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6690
      Begin VB.TextBox tsearch 
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Top             =   5400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Μον.Μετρ. UBL"
         Height          =   360
         Left            =   5040
         TabIndex        =   19
         Top             =   7920
         Width           =   990
      End
      Begin VB.ListBox MYDATA 
         BackColor       =   &H0080FFFF&
         Height          =   3960
         Left            =   4920
         TabIndex        =   18
         Top             =   3855
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Εκτύπωση"
         Height          =   270
         Left            =   2535
         TabIndex        =   14
         Top             =   360
         Width           =   2190
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "BOHU1.frx":0018
         Height          =   2310
         Left            =   150
         TabIndex        =   13
         Top             =   5685
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   4075
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         BackColor       =   33023
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
      Begin VB.CommandButton Command6 
         Caption         =   "Διαγραφή"
         Height          =   360
         Left            =   1521
         TabIndex        =   11
         Top             =   3861
         Width           =   1035
      End
      Begin VB.TextBox TIMH 
         Height          =   330
         Left            =   1495
         TabIndex        =   10
         Top             =   4836
         Width           =   945
      End
      Begin VB.TextBox PERIGRAFH 
         Height          =   330
         Left            =   1482
         TabIndex        =   9
         Top             =   4407
         Width           =   3345
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Προσθήκη"
         Height          =   360
         Left            =   3744
         TabIndex        =   8
         Top             =   3861
         Width           =   1035
      End
      Begin VB.Label lsearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Αναζήτηση"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   5400
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Τιμή"
         ForeColor       =   &H8000000E&
         Height          =   260
         Left            =   351
         TabIndex        =   17
         Top             =   4901
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Περιγραφή"
         ForeColor       =   &H8000000E&
         Height          =   260
         Left            =   234
         TabIndex        =   16
         Top             =   4420
         Width           =   1209
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Επιλογή επιχείρησης"
      Height          =   5175
      Left            =   7020
      TabIndex        =   4
      Top             =   240
      Width           =   5161
      Begin MSDataGridLib.DataGrid DBGrid3 
         Bindings        =   "BOHU1.frx":002C
         Height          =   3861
         Left            =   273
         TabIndex        =   7
         Top             =   364
         Width           =   4537
         _ExtentX        =   8017
         _ExtentY        =   6826
         _Version        =   393216
         BackColor       =   16777152
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
            ScrollBars      =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc KATHG 
      Height          =   375
      Left            =   4320
      Top             =   9120
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      DataSourceName  =   "emp3"
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
End
Attribute VB_Name = "BOHU1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub Combo1_Click()

        '<EhHeader>
        On Error GoTo Combo1_Click_Err

        '</EhHeader>

100     DbGrid2.Enabled = True



LSEARCH.Visible = False
tsearch.Visible = False




        'If Combo1.ListIndex = 5 Then
        '   Data2.RecordSource = "SELECT *FROM PINAKES"
        'Else
        Dim sql As String

110     If Val(Left(Combo1.Text, 2)) = 3 Then
120         sql = "SELECT *FROM PINAKES WHERE TYPOS=3 order BY TIMH,AYJON "
130         KATHG.RecordSource = "SELECT AYJON,PERIGRAFH FROM PINAKES WHERE TYPOS=11 order BY TIMH,AYJON "
140         KATHG.ConnectionString = gConnect
150         KATHG.Refresh

        Else
160         sql = "SELECT *FROM PINAKES WHERE TYPOS=" + Left(Combo1.Text, 2) + " order BY AYJON "
        End If

mydata.Visible = False


'
''1 ΦΠΑ συντελεστής 24% 24%
'2 ΦΠΑ συντελεστής 13% 13%
'3 ΦΠΑ συντελεστής 6% 6%
'4 ΦΠΑ συντελεστής 17% 17%
'5 ΦΠΑ συντελεστής 9% 9%
'6 ΦΠΑ συντελεστής 4% 4%
'7 ’νευ Φ.Π.Α. 0%
'8 Εγγραφές χωρίς ΦΠΑ (πχ Μισθοδοσία, Αποσβέσεις) -

' monades metrisis
'Κωδικός Περιγραφή
If Val(Left(Combo1.Text, 2)) = 2 Then
    mydata.Visible = True
     mydata.Clear
    mydata.AddItem "1 Τεμάχια  H87"
    mydata.AddItem "2 Κιλά     KGM"
    mydata.AddItem "3 Λίτρα    LTR"
     mydata.AddItem "4 Μέτρα    MTR"
      mydata.AddItem "5 τετ.μετ  MTK"
       mydata.AddItem "6 Κυβ.Μετ  MTQ"
        mydata.AddItem "7 Λοιπά"
         mydata.AddItem " TNE TONNOI  "
    mydata.AddItem " XBX ΚΙΒΩΤΙΟ  "
    mydata.AddItem " DAY,HUR,MIN HMER/ΩΡΑ/ΛΕΠΤ  "
     mydata.AddItem " XPK,XPX,XSA ΠΑΚΕΤΟ/ΠΑΛΕΤΑ/ΣΑΚΚΟΣ  "
     
     
LSEARCH.Visible = True
tsearch.Visible = True


     
         KATHG.RecordSource = "SELECT [Code value],[Name],[Description] from MONADES "
       KATHG.ConnectionString = gConnect
         KATHG.Refresh
     
     
     
     
    




End If

If Val(Left(Combo1.Text, 2)) = 13 Then
    mydata.Visible = True
     mydata.Clear
    mydata.AddItem "UBL=>ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ"
   
   
    
End If

'----------- kathg fpa ------------------
If Val(Left(Combo1.Text, 2)) = 1 Then
'
    mydata.Visible = True
     mydata.Clear
    mydata.AddItem "1 ΦΠΑ συντελεστής 24%"
    mydata.AddItem "2 ΦΠΑ συντελεστής 13% "
    mydata.AddItem "3 ΦΠΑ συντελεστής 6%"
    mydata.AddItem "4 ΦΠΑ συντελεστής 17%"
    mydata.AddItem "5 ΦΠΑ συντελεστής 9%"
    mydata.AddItem "6 ΦΠΑ συντελεστής 4%"
    mydata.AddItem "7 ’νευ Φ.Π.Α. 0%"
    mydata.AddItem "8 Εγγραφές χωρίς ΦΠΑ (πχ Μισθοδοσία, Αποσβέσεις)"
End If


























   ' MsgBox "Η ΤΙΜΗ αν είναι 1 τότε πιστώνει ισόποσα τον πελάτη /(χρεώνει προμηθευτή).Αν είναι 0 δεν κάνει τίποτα"
'End If










' ΤΡΟΠΟΙ ΠΛΗΡΩΜΗΣ
If Val(Left(Combo1.Text, 2)) = 12 Then
'
    mydata.Visible = True
     mydata.Clear
    mydata.AddItem "1 Επαγ.Λογαριασμός Πληρωμών Ελλάδος"
    mydata.AddItem "2 Επαγ. Λογαριασμός Πληρωμών Εξωτερικού"
    mydata.AddItem "3 Μετρητά"
    mydata.AddItem "4 Επιταγή"
    mydata.AddItem "5 Επί Πιστώσει"
    mydata.AddItem "6 Web Banking"
    mydata.AddItem "7 POS/e-POS"
     mydata.AddItem "UBL=>ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ"

    MsgBox "Η ΤΙΜΗ αν είναι 1 τότε πιστώνει ισόποσα τον πελάτη /(χρεώνει προμηθευτή).Αν είναι 0 δεν κάνει τίποτα"
End If


'------------  SKOPOS ------------------------------------------
If Val(Left(Combo1.Text, 2)) = 7 Then
    mydata.Visible = True
    mydata.Clear
    mydata.AddItem "1 Πώληση"
    mydata.AddItem "2 Πώληση για Λογαριασμό Τρίτων"
    mydata.AddItem "3 Δειγματισμός"
    mydata.AddItem "4 Έκθεση"
    mydata.AddItem "5 Επιστροφή"
    mydata.AddItem "6 Φύλαξη"
    mydata.AddItem "7 Επεξεργασία Συναρμολόγηση"
    mydata.AddItem "8 Μεταξύ Εγκαταστάσεων Οντότητας"

End If



170     data2.RecordSource = sql
180     DbGrid2.columns(0).Width = 500
190     DbGrid2.columns(1).Width = 2500
200     DbGrid2.columns(2).Width = 500

210     data2.Refresh

220     If Check1 Then
            'typos2
230         print3_xar UCase(sql), "00000", Combo1.Text, 0

        End If

        '<EhFooter>
        Exit Sub

Combo1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU1.Combo1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU1.Combo1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command1_Click()

        '   uSEL_PARASTAT.Show
        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim mgdir As String, a, B

        Dim DB As Database

        Dim mtit

100     mtit = InputBox("Δωσε τον TITΛΟ της νέας χρήσης ")

110     mgdir = InputBox("Δωσε την βάση της νέας χρήσης ")

        'A = Dir(mgdir, vbDirectory)
        '
        'If Len(A) > 0 Then
        '  MsgBox "Υπάρχει ο φάκελος. Ακύρωση δημιουργίας"
        '  Exit Sub
        'End If

        On Error Resume Next

        'Gdb.Close
        'Gdb.Open gConnect

        Dim N

120     Gdb.Execute "INSERT INTO ETAIREIES (TITLOS,DSN)  VALUES ('" + mtit + "','" + mgdir + "')", N

130     MDIForm1.MakeBackup "C:\MERCVB"

        Dim DD As String

140     Gdb.Execute "create database " + mgdir

        Exit Sub

        '---------- RESTORE ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ ΣΤΗΝ DATABASE ΤΗΣ ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ -------------------------------------

150     Set DB = OpenDatabase("", False, False, gConnect + ";" + "DATABASE=" + mgdir + ";")
160     Me.MousePointer = vbHourglass

        Dim LISTA(20)

170     LISTA(1) = "PEL"
180     LISTA(2) = "EID"
190     LISTA(3) = "EGG"
200     LISTA(4) = "EGGTIM"
210     LISTA(5) = "TIM"
220     LISTA(6) = "GRA"
230     LISTA(7) = "MEM"
240     LISTA(8) = "PINAKES"
250     LISTA(9) = "XREOPIS"
260     LISTA(10) = "ARITMISI"
270     LISTA(11) = "PARASTAT"
280     LISTA(12) = "BARCODES"
290     LISTA(13) = "HMEROL"
300     LISTA(14) = "HMEROL2"
310     LISTA(15) = "PARAMETROI"
320     LISTA(16) = "CalendarEvents"
330     LISTA(17) = "CalendarRecurrencePatterns"

340     Me.Caption = "2.ΑΝΤΙΓΡΑΦΗ ΣΕ ΠΑΛΙΑ ΧΡΟΝΙΑ"

        Dim dbnea As New ADODB.Connection

350     dbnea.Open gConnect + ";database=" + mgdir

        'Gdb.Open gConnect + ";DATABASE=" + cpalia
        Dim k

        Dim LATOS

360     For k = 1 To 17
370         LATOS = 0

            '   On Error GoTo latos
380         DoEvents
390         Me.Caption = LISTA(k)

            '   dbnea.Execute "DROP TABLE " + LISTA(k)
400         If LISTA(k) = "EGG" Or LISTA(k) = "EGGTIM" Or LISTA(k) = "TIM" Then
410             DB.Execute "SELECT top 1 * INTO " + LISTA(k) + "  FROM  " + LISTA(k) + "  IN '" + DD + "'"
            Else
420             DB.Execute "SELECT top 1 * INTO " + LISTA(k) + "  FROM  " + LISTA(k) + "  IN '" + DD + "'"
            End If

        Next

430     Me.MousePointer = vbNormal

440     dbnea.Close

        Exit Sub

        Exit Sub

450     data1.RecordSource = ""
460     data2.RecordSource = ""
470     DATA3.RecordSource = ""
480     data1.Refresh
490     data2.Refresh
500     DATA3.Refresh
        ' On Error Resume Next

510     B = Array("eid.dbf", "tim.dbf", "eggtim.dbf", "gra.dbf", "egg.dbf", "parastat.dbf", "mem.dbf")

520     For Each a In B

            On Error Resume Next

530         FileCopy gDir + "\" + a, mgdir + "\" + a
        Next

        Dim DUM

540     DUM = OpenD(DB)

        On Error GoTo Command1_Click_Err

550     DB.Execute "select pinakes.* into pinakes in '" + mgdir + "' '" + gConnect + "' from pinakes;"

560     If gConnect = "dBase III;" Then
570         Set DB = OpenDatabase(mgdir, False, False, gConnect)
        Else
580         Set DB = OpenDatabase(mgdir, False, False)
        End If

590     DB.Execute "delete *from egg"
600     DB.Execute "delete *from eggtim"
610     DB.Execute "delete *from tim"
620     DB.Execute "update pel set ayp=typ"
630     DB.Execute "update eid set apos=pos"

640     DATA3.RecordSource = "etaireie.dbf"
650     DATA3.Refresh

660     DATA3.Recordset.AddNew
670     DATA3.Recordset("dir") = mgdir
680     DATA3.Recordset("titlos") = mtit
690     DATA3.Recordset("nomisma") = "E"
700     DATA3.Recordset.Update

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU1.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU1.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

      Dim R As Long
   On Error Resume Next
   R = ShellExecute(0, "open", "https://docs.peppol.eu/poacc/billing/3.0/codelist/UNECERec20/", 0, 0, 1)
' WebBrowser1.Navigate2 "https://el.wikipedia.org/wiki/ISO_3166-1"

'https://docs.peppol.eu/poacc/billing/3.0/codelist/UNECERec20/

End Sub

Private Sub Command3_Click()


Dim arxgConnect As String

arxgConnect = gConnect

'Private Sub MDIForm_DblClick()
Dim frm As Form

    For Each frm In Forms
        If Not frm.Name = "MDIForm1" Then 'avoid unloading mdi form
            If frm.MDIChild = True Then 'unload only mdi children
               If Not frm.Name = "BOHU1" Then
                  Unload frm
               End If
            End If
        End If
    Next frm

'End Sub





'frmLogin.SHOW 1




        'On Error Resume Next
        '  Gdb.Close
        ' gConnect = Trim(DATA3.Recordset("DSN"))+";DATABASE="+
        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>
        On Error GoTo SHOW

        Dim D

100     D = Trim(DATA3.Recordset("DSN"))
        ' Gdb.DefaultDatabase = D

        '   Gdb.Close
Dim DS

Dim mfile As String
mfile = App.Path + "\MERCPATH.TXT"
If Len(Dir(mfile, vbNormal)) > 2 Then
   mfile = App.Path + "\MERCPATH.TXT"
Else
   mfile = "C:\MERCPATH.TXT"
End If






     Open mfile For Input As #1
       Line Input #1, DS
       Line Input #1, gConnect
     Close #1







110     If InStr(gConnect, "SQLOLE") > 0 Then

            Dim c As Long

120         c = InStr(gConnect, "Catalog=")

130         gConnect = Left(gConnect, c + 7) + D

140     ElseIf InStr(UCase(gConnect), "DSN") > 0 Then

150         gConnect = gConnect + ";DATABASE=" + D

        Else

160         Open "C:\MERCPATH.TXT" For Input As #1
170         Line Input #1, gDir
180         Line Input #1, gConnect
190         Close #1

200         gConnect = gConnect + ";DATABASE=" + D
        End If

210     If InStr(UCase(D), "DSN") > 0 Then
220         gConnect = D
        End If

        On Error Resume Next

230     Gdb.Close

        On Error GoTo outout

240     Gdb.Open gConnect

        On Error Resume Next

250     If Left(D, 1) = "D" Then
260         gPicture = "C:\MERCVB\PALIA.BMP"
        Else
270         gPicture = "C:\MERCVB\BACKGROUND.BMP"
        End If

        On Error GoTo OUTSIDE
        
        
    Dim P_VER As String ' VERSION PROGRAMMATOS
'P_VER = Format(App.Major, "00") + "." + Format(App.Minor, "00") + "." + Format(App.Revision, "00")
    
   Dim F_VER As String
        
'F_VER = FINDPARAMETROI(1, "MDIFORM1", "F_VER", "..", "MERCURY VERSION ")


'Dim F_VER As String ' VERSION DATABASE
'Dim P_VER As String ' VERSION PROGRAMMATOS
P_VER = Format(App.Major, "00") + "." + Format(App.Minor, "00") + "." + Format(App.Revision, "00")


F_VER = FINDPARAMETROI(1, "MDIFORM1", "F_VER", "..", "MERCURY VERSION ")




If Trim(F_VER) = Trim(P_VER) Then
        'OK
Else
340     Update_Structures
        Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + P_VER + "' WHERE  FORMA='MDIFORM1' AND VAR='F_VER' "
End If





'If Trim(F_VER) = Trim(P_VER) Then
        'OK
'Else
'340     Update_Structures
'        Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + P_VER + "' WHERE  FORMA='MDIFORM1' AND VAR='F_VER' "
'End If


'Gdb.Close
'Gdb.Open gConnect

PARAMETROI_TO_ARRAY 'par1


280     'Update_Structures
        Dim checkok As Integer
        checkok = 2
290     LOADPARAMETERS checkok
        If checkok = -1 Then GoTo OUTSIDE
        

300     MDIForm1.Picture = LoadPicture(gPicture)

310     MDIForm1.StatusBar1.Panels(1).Text = DATA3.Recordset("TITLOS")
320     Unload Me

        Exit Sub



outout:


MsgBox "αδυνατη η συνδεση"

gConnect = arxgConnect
Gdb.Open gConnect

Exit Sub






OUTSIDE:

     Update_Structures
       'On Error Resume Next
'        Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + P_VER + "' WHERE  FORMA='MDIFORM1' AND VAR='F_VER' "


'330     MsgBox "Αδύνατη η σύνδεση στην εταιρεία " + Trim(DATA3.Recordset("TITLOS"))
          Resume Next
        Exit Sub

SHOW:
     MsgBox Err.Description

         Resume Next
        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        ''MsgBox Err.Description & vbCrLf & _
          "in ADOMERCNEW.BOHU1.Command3_Click " & _
          "at line " & Erl, _
          vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU1.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()
    ' bohu11.Show
End Sub

Private Sub Command5_Click()

        ' ΠΡΟΣΘΕΤΩ ΚΑΤΗΓΟΡΙΑ
        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>

100     If Len(Combo1.Text) = 0 Then
110         MsgBox "ΔΙΑΛΕΞΤΕ ΚΑΤΗΓΟΡΙΑ"

            Exit Sub

        End If

        Dim a As String, RECS As Long

        Dim R As New ADODB.Recordset

120     If 1 + Combo1.ListIndex = 3 Then    ' υποοικογενεια
130         If Val(TIMH.Text) = 0 Then
140             MsgBox "ΔΙΑΛΕΞΤΕ ΟΙΚΟΓΕΝΕΙΑ"

                Exit Sub

            Else
150             R.Open "SELECT MAX(AYJON) FROM PINAKES WHERE TYPOS=3 and TIMH=" + TIMH.Text, Gdb, adOpenForwardOnly, adLockReadOnly
            End If

        Else
160         R.Open "SELECT MAX(AYJON) FROM PINAKES WHERE TYPOS=" + str(1 + Combo1.ListIndex), Gdb, adOpenForwardOnly, adLockReadOnly
        End If

        Dim MMAX As Integer

170     MMAX = IIf(IsNull(R(0)), 0, R(0))

180     If Val(TIMH.Text) = 0 Then TIMH.Text = 0

190     a = "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES ("
200     a = a + str(1 + Combo1.ListIndex) + ","    ' TYPOS
210     a = a + str(1 + MMAX) + ","    ' AYJON
220     a = a + "'" + Left(PERIGRAFH.Text, 20) + "'" + ","
230     a = a + TIMH.Text + ")"

        On Error GoTo LATOS

240     Gdb.Execute a, RECS

250     data2.Refresh

        Exit Sub

LATOS:
260     MsgBox Err.Description

270     Resume Next

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU1.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU1.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>
        Dim N As Long, T As Long

        On Error Resume Next

100     N = data2.Recordset("ayjon")
110     T = data2.Recordset("typos")

120     Gdb.Execute "delete FROM PINAKES WHERE AYJON=" + str(N) + " AND TYPOS=" + str(T)

130     data2.Refresh

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU1.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU1.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command7_Click()

        '<EhHeader>
        On Error GoTo Command7_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU1.Command7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU1.Command7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command8_Click()

        '<EhHeader>
        On Error GoTo Command8_Click_Err

        '</EhHeader>
        
        Gdb.Open gConnect
        

100     Gdb.Execute "DELETE FROM  ETAIREIES WHERE TITLOS='" + Trim(DBGrid3.columns(1).Text) + "'"
110     DATA3.Refresh

120     DBGrid3.Refresh

        '<EhFooter>
        Exit Sub

Command8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU1.Command8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU1.Command8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        '</EhHeader>

100     data2.RecordSource = "SELECT * FROM PINAKES WHERE TYPOS=3 AND TIMH=" + str(KATHG.Recordset("AYJON")) + " ORDER BY AYJON"
110     data2.Refresh

        '<EhFooter>
        Exit Sub

DataGrid1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU1.DataGrid1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU1.DataGrid1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command7_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU1.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU1.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl
110     Me.KeyPreview = True

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

120     Me.Picture = LoadPicture(gPicture)
130     data1.RecordSource = "SELECT *FROM PARASTAT"
140     data1.ConnectionString = gConnect

        Dim mConnect, mdir
Dim mfile As String
mfile = App.Path + "\MERCPATH.TXT"
If Len(Dir(mfile, vbNormal)) < 2 Then
   mfile = App.Path + "\MERCPATH.TXT"
Else
   mfile = "C:\MERCPATH.TXT"
End If









        On Error Resume Next

       ' On Error GoTo CREATE_ETAIR

        ' DATA3.Refresh
        On Error Resume Next

220     data2.RecordSource = "SELECT *FROM PINAKES"
230     data2.ConnectionString = gConnect

240     Combo1.AddItem " 1.ΚΑΤΗΓΟΡΙΕΣ ΦΠΑ"
250     Combo1.AddItem " 2.ΜΟΝΑΔΕΣ ΜΕΤΡΗΣΗΣ"
260     Combo1.AddItem " 3.ΥΠΟΟΙΚΟΓΕΝΕΙΕΣ/ΚΑΤΗΓΟΡΙΕΣ"
270     Combo1.AddItem " 4.ΑΠΟΘΗΚΕΣ"
280     Combo1.AddItem " 5.ΠΩΛΗΤΕΣ"
290     Combo1.AddItem " 6.ΤΡΑΠΕΖΕΣ"
300     Combo1.AddItem " 7.ΣΚΟΠΟΣ"
310     Combo1.AddItem " 8.ΤΟΠΟΣ ΦΟΡΤΩΣΗΣ"
320     Combo1.AddItem " 9.ΤΟΠΟΣ ΠΡΟΟΡΙΣΜΟΥ"
330     Combo1.AddItem "10.ΑΥΤΟΚΙΝΗΤΑ"
340     Combo1.AddItem "11.ΟΙΚΟΓΕΝΕΙΣ"
350     Combo1.AddItem "12.ΤΡΟΠΟΙ ΠΛΗΡΩΜΗΣ"
360     Combo1.AddItem "13.ΚΑΤΗΓΟΡΙΕΣ ΠΕΛΑΤΩΝ"
370     Combo1.AddItem "14.ΥΠΑΛΛΗΛΟΙ"
380     Combo1.AddItem "15.ΚΑΤΗΓΟΡΙΕΣ ΕΙΔΩΝ"
390     Combo1.AddItem "16.Λογαριασμοί Οψεως"
400     Combo1.AddItem "17.Λογαριασμοί Ταμιευτηρίου"
410     Combo1.AddItem "18.Λογαριασμοί Δανείων"
420     Combo1.AddItem "19.Αιτιολογίες Εργασιών"
430     Combo1.AddItem "20.Import ειδών από Excel"

440     Combo1.AddItem "21.Import Πελατών/Προμηθευτών από Excel"
450     Combo1.AddItem "22.Προελεύσεις ειδών"
460     Combo1.AddItem "23.X23"
470     Combo1.AddItem "24.X24"
480     Combo1.AddItem "25.X25"
490     Combo1.AddItem "26.X26"
500     Combo1.AddItem "27.X27"
510     Combo1.AddItem "28.X28"
520     Combo1.AddItem "29.X29"

530     Combo1.AddItem "30.XΡΩΜΑΤΑ ΠΑΙΔΙΚΑ"
540     Combo1.AddItem "31.ΜΕΓΕΘΗ ΠΑΙΔΙΚΑ"

550     Combo1.AddItem "32.ΧΡΩΜΑΤΑ ΕΦΗΒΙΚΑ"
560     Combo1.AddItem "33.ΜΕΓΕΘΗ ΕΦΗΒΙΚΑ"

570     Combo1.AddItem "34.ΧΡΩΜΑΤΑ ΓΥΝΑΙΚΕΙΑ"
580     Combo1.AddItem "35.ΜΕΓΕΘΗ ΓΥΝΑΙΚΕΙΑ"

590     Combo1.AddItem "36.ΧΡΩΜΑΤΑ ΑΝΔΡΙΚΑ"
600     Combo1.AddItem "37.ΜΕΓΕΘΗ ΑΝΔΡΙΚΑ"

610     Combo1.AddItem "38.X38"
620     Combo1.AddItem "39.X39"
630     Combo1.AddItem "40.BOHΘΗΤΙΚΟΣ"
640     Combo1.AddItem "41.ΕΙΔΗ ΕΠΙΣΚΕΥΩΝ"
         Combo1.AddItem "42.ΚΩΔ.ΠΑΡ/ΚΟΥ ELINE"
         Combo1.AddItem "43.ΕΙΔΗ ΠΑΡ/ΚΟΥ MYDATA"
          Combo1.AddItem "44.ΑΠΑΛΛΑΓΕΣ ΦΠΑ"
          Combo1.AddItem "45.ΠΑΡΑΚΡΑΤΗΣΕΙΣ"
           Combo1.AddItem "46."
           Combo1.AddItem "47."
           Combo1.AddItem "48."
           Combo1.AddItem "49."
           Combo1.AddItem "50.  ΤΥΠΟΣ ΟΧΗΜΑΤΟΣ"
           Combo1.AddItem "51.ΜΑΡΚΑ ΟΧΗΜΑΤΟΣ"
           Combo1.AddItem "52.ΣΥΝΕΡΓΕΙΑ"
           Combo1.AddItem "53.ΠΛΥΝΤΗΡΙΑ"
           Combo1.AddItem "54.ΤΕΛΗ"
           Combo1.AddItem "55.ΦΟΡΟΙ"
           Combo1.AddItem "56.ΧΑΡΤΟΣΗΜΟ"
          
' Combo1.AddItem "42.ΜΕΤΑΦΟΡΙΚΕΣ"

650     Combo1.Text = " "



      If Len(Dir(App.Path + "\MERCPATH.TXT")) > 0 Then

         Open App.Path + "\MERCPATH.TXT" For Input As #1
         Line Input #1, gDir
         Line Input #1, mConnect
         Close #1
        Else

         Open "C:\MERCPATH.TXT" For Input As #1
            Line Input #1, gDir
            Line Input #1, mConnect
         Close #1

        End If




'150     Open mfile For Input As #1
'160     Line Input #1, mdir
'170     Line Input #1, mConnect
'180     Close #1

        '  RichTextBox1.Text = mdir + Chr(13) + mConnect

190     DATA3.ConnectionString = mConnect

        On Error GoTo sub_exit

200     DATA3.RecordSource = "SELECT * FROM ETAIREIES ORDER BY TITLOS"
210     DATA3.Refresh












        Exit Sub

CREATE_ETAIR:
660     MsgBox Err.Description

670     Gdb.Execute "CREATE TABLE ETAIREIES ( DSN  char(60),TITLOS  CHAR(30) )"
680     Gdb.Execute "INSERT INTO ETAIREIES (DSN,TITLOS) VALUES ('DSN=COMPAQ','ΤΡΕΧΟΥΣΑ')"
690     Gdb.Execute "INSERT INTO ETAIREIES (DSN,TITLOS) VALUES ('DSN=2005','2005')"

700     Resume Next

CREATE_ETAIREIES:

710     On Error Resume Next
Gdb.Execute "CREATE TABLE ETAIREIES ( DSN  char(60),TITLOS  CHAR(30) )"

        'On Error Resume Next
720     Gdb.Execute "SELECT *  INTO ETAIREIES FROM ETAIREIES"

730     DATA3.RecordSource = "SELECT * FROM ETAIREIES"
740     DATA3.Refresh

750     Resume Next

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.BOHU1.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.BOHU1.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>
        
        
sub_exit:
        Exit Sub
        
        
        
        

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 7, 10620, 9915, fh, fw, ft, fl 'mForm_Resize Me,8, 10620, 9915, fh, fw, ft, fl
    DBGrid3.columns(1).Width = DBGrid3.Width - DBGrid3.columns(1).Width
    
    'ResizeFormFor Me
End Sub

Private Sub tsearch_Change()
Dim Q As String
Q = UCase(Trim(tsearch.Text))
On Error Resume Next

       KATHG.RecordSource = "SELECT [Code value],[Name],[Description] from MONADES where upper([Description]) like '%" + Q + "%' or upper([Code value]) like '%" + Q + "%' or upper([Name]) like '%" + Q + "%'"
       KATHG.ConnectionString = gConnect
         KATHG.Refresh
End Sub
