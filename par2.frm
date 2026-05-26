VERSION 5.00
Object = "{CCA214C0-DFEB-4C91-9F0D-2665F77F6E23}#1.2#0"; "IDAutomationLinear.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form par2 
   BackColor       =   &H00FF0000&
   Caption         =   "Εττικέτες-Διαγ.Παραστ."
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   12225
   WindowState     =   2  'Maximized
   Begin VB.CheckBox kostologhsh 
      Caption         =   "Κοστολόγηση"
      Height          =   255
      Left            =   5040
      TabIndex        =   33
      Top             =   8640
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc EID 
      Height          =   375
      Left            =   3000
      Top             =   9000
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "EID"
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
   Begin VB.CommandButton cmdKOSTOLOGHSH 
      Caption         =   "Κοστολόγηση"
      Height          =   420
      Left            =   4935
      TabIndex        =   32
      Top             =   7155
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000080FF&
      Caption         =   "Αλλαγή σε άλλο Παραστατικό με τον ίδιο αριθμό"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox ekkrem 
      BackColor       =   &H00FF0000&
      Caption         =   "Μόνο τα εκκρεμή"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   8280
      TabIndex        =   27
      Top             =   7050
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "par2.frx":0000
      Left            =   120
      List            =   "par2.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   1200
      Width           =   2415
   End
   Begin MSComctlLib.ImageCombo Combo2 
      Height          =   330
      Left            =   75
      TabIndex        =   24
      Top             =   105
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
      ImageList       =   "ImageList1"
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3240
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   8220
      Visible         =   0   'False
      Width           =   3270
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Επαναφέρω σε εκκρεμότητα"
      Height          =   270
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Το μετατρέπω σε ατιμολόγητο"
      Top             =   7920
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid timGrid 
      Height          =   1575
      Left            =   105
      TabIndex        =   21
      Top             =   1695
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   8
      FixedRows       =   7
      BackColorFixed  =   15581332
      TextStyle       =   1
      TextStyleFixed  =   1
      GridLinesFixed  =   1
      MergeCells      =   1
   End
   Begin VB.CheckBox noask 
      Caption         =   "Ολα ετικέτες"
      Height          =   210
      Left            =   3240
      TabIndex        =   20
      Top             =   7905
      Width           =   3270
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Βγάζω από εκκρεμότητα"
      Height          =   270
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Το μετατρέπω σε τιμολογημένο"
      Top             =   7635
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   9795
      Top             =   2730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2370
      Left            =   8010
      ScaleHeight     =   115.5
      ScaleMode       =   2  'Point
      ScaleWidth      =   153.75
      TabIndex        =   17
      Top             =   8595
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ΜΕΤΟΝΟΜΑΣΙΑ ΦΩΤΟΓΡΑΦΙΩΝ"
      Height          =   450
      Left            =   5805
      TabIndex        =   16
      Top             =   9120
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1485
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.CheckBox Check5 
      Caption         =   "CODE 128"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   7635
      Value           =   1  'Checked
      Width           =   3270
   End
   Begin MSAdodcLib.Adodc EGGTIM 
      Height          =   375
      Left            =   360
      Top             =   9660
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSDataGridLib.DataGrid DBGrid1 
      Bindings        =   "par2.frx":002E
      Height          =   1545
      Left            =   4560
      TabIndex        =   8
      Top             =   165
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   2725
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Bindings        =   "par2.frx":0040
      Height          =   3615
      Left            =   90
      TabIndex        =   7
      Top             =   3300
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   16761024
      Cols            =   10
      FixedCols       =   0
      BackColorBkg    =   16761024
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSAdodcLib.Adodc TIM 
      Height          =   375
      Left            =   2940
      Top             =   9630
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "tim"
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
   Begin VB.CommandButton Command5 
      Caption         =   "Ετικέτες"
      Height          =   420
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "το ""C:\MERCVB\MELIS.TXT"" =>κολα αντι ποσοτητα"
      Top             =   7155
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "¨Εξοδος"
      Height          =   420
      Left            =   10650
      TabIndex        =   5
      Top             =   7710
      Width           =   1365
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Διαγραφή"
      Height          =   420
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7155
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Εκτύπωση"
      Height          =   420
      Left            =   8190
      TabIndex        =   3
      Top             =   7740
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Υπολογισμός"
      Height          =   360
      Left            =   2820
      TabIndex        =   1
      Top             =   1200
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1485
      TabIndex        =   0
      Top             =   495
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   2835
      TabIndex        =   14
      Top             =   450
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   155254785
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   300
      Left            =   2835
      TabIndex        =   15
      Top             =   825
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   155254785
      CurrentDate     =   38814
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6975
      Top             =   8115
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":0055
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":05A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":0AE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":102B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":18BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":1E0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":1F21
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":2473
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":29C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":2F17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":34A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par2.frx":3A3B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox bohuitiko 
      Height          =   285
      Left            =   9375
      TabIndex        =   29
      Top             =   6555
      Width           =   135
   End
   Begin VB.Label exodos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   6720
      TabIndex        =   34
      Top             =   6960
      Width           =   45
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode3 
      Height          =   900
      Left            =   4560
      TabIndex        =   30
      Top             =   10320
      Visible         =   0   'False
      Width           =   1995
      _cx             =   3519
      _cy             =   1587
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
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   30
      Top             =   6930
      Width           =   12015
   End
   Begin VB.Label CallFromOther 
      BackStyle       =   0  'Transparent
      Caption         =   "CallFromOther"
      Height          =   240
      Left            =   9615
      TabIndex        =   28
      Top             =   705
      Visible         =   0   'False
      Width           =   2160
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode1 
      Height          =   1065
      Left            =   120
      TabIndex        =   25
      Top             =   8640
      Visible         =   0   'False
      Width           =   1500
      _cx             =   2646
      _cy             =   1879
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
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
      NarrowBarWidth  =   0,03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   0
      Code128CharSet  =   2
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   1
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0,2
      TopMarginCM     =   0,2
      SupplementToEncode=   ""
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode2 
      Height          =   1065
      Left            =   8250
      TabIndex        =   19
      Top             =   7635
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
      SymbologyId     =   13
      DataToEncode    =   "123456789012"
      Orientation     =   0
      BarHeight       =   1
      NarrowBarWidth  =   0,03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   0
      Code128CharSet  =   2
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   0
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0,2
      TopMarginCM     =   0,2
      SupplementToEncode=   ""
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Επωνυμία"
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
      Left            =   60
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Εως"
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
      Left            =   2385
      TabIndex        =   11
      Top             =   825
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Από"
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
      Left            =   2385
      TabIndex        =   10
      Top             =   465
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Αρ.Παρ/κού"
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
      Left            =   45
      TabIndex        =   2
      Top             =   495
      Width           =   1335
   End
End
Attribute VB_Name = "par2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f_printer As String

Dim F_ET_EPIX_FONTSIZE As Integer
Dim F_ET_EPIX_X As Long
Dim F_ET_EPIX_Y As Long

            Dim F_ET_PER_X, F_ET_PER_Y, F_ET_BAR_H
  
   Dim F_ET_KOD_X, F_ET_KOD_Y
            Dim F_ET_BAR_X, F_ET_BAR_Y
            Dim F_ET_TIM_X, F_ET_TIM_Y








Dim F_STADIO    ' STADIO PARASTATIKOY

'Dim F_PSIFIA_ATIM ' = Val(FindParametroi(1,"PAR1", "F_PSIFIA_ATIM", "5", "ΨΗΦΙΑ ΤΙΜΟΛΟΓΙΟΥ (5)"))
Dim F_1ST_CHOICE    ' ΠΟΙΟ ΠΑΡΑΣΤΑΤΙΚΟ ΘΑ ΔΕΙΧΝΕΙ ΠΡΏΤΟ

Dim f_psifia_atim

Dim F_ArPerEtik    ' αριστερο περιθωριο ετικεττας

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
Dim f_IS_PARTIDES As Integer  ' = Val(FindParametroi(1,"APOT2", "f_IS_PARTIDES", "0", "Αν Χρησιμοποιώ παρτίδες=1"))
Private Sub bohuitiko_LostFocus()

        '<EhHeader>
        On Error GoTo bohuitiko_LostFocus_Err

        '</EhHeader>

100     Combo2.Text = "*"
110     Command1_Click

        '<EhFooter>
        Exit Sub

bohuitiko_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.bohuitiko_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.bohuitiko_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdKOSTOLOGHSH_Click()





If TIM.Recordset("EIDOS") = "e" Then
   MsgBox " κοστολογηση μονο σε αγορές"
   Exit Sub
End If








            EID.ConnectionString = gConnect
            






        '<EhHeader>
        On Error GoTo cmdKOSTOLOGHSH_Click_Err

        '</EhHeader>
        Dim D1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        'On Error Resume Next
        Dim ar_per

100     ar_per = F_ArPerEtik    ' ρυθμιζομενο

        Dim x As Printer



        'eggtim.DatabaseName = gDir
        On Error GoTo LATOS

140     D1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
150     d2 = DateAdd("d", 1, D1)

160     TIM.ConnectionString = gConnect

170     eggtim.ConnectionString = gConnect

180     eggtim.RecordSource = "select KODE,EID.ONO,POSO,TIMM,EKPT,EKPT2,FPA,STR(round(EID.LTI5,2),10,2) as [ΛΙΑΝ.ΤΙΜ],STR(ROUND(EID.LTI,2),10,2) AS [XON.TIM],STR(EID.XTI,10,2) AS [TIM.ΑΓ],EID.ERG,EID.CH4,PROELEYSH,KOLA,MON,EID.FPA,PELKOD,ATIM,HME from EGGTIM inner join EID on EGGTIM.KODE=EID.KOD where left(ATIM," + str(1 + f_psifia_atim) + ")='" + Left(TIM.Recordset("atim"), 1 + f_psifia_atim) + "' and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'  order by PROOD_AJ ;"

190     eggtim.Refresh
        'Label5.Caption = TIM.Recordset("epo")
        'Label2.Caption = TIM.Recordset("die")
        'Label3.Caption = TIM.Recordset("epa")
        'Label4.Caption = TIM.Recordset("afm")

        'DUM = OpenD(db)
        'Set rec = db.OpenRecordset("select *from mem")

      

200     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

210     If eggtim.Recordset.EOF Then
220         MsgBox "ΔΕΝ ΕΥΡΕΘΗΣΑΝ ΕΓΓΡΑΦΕΣ "

            Exit Sub

        End If

        Dim nkod

230     eggtim.Recordset.MoveFirst

        Dim mlabel, MPER, mtimes, TIMM, timologio
        
                kostologhsh.Value = vbChecked
                

240     Do While Not eggtim.Recordset.EOF


                 EID.RecordSource = "SELECT *FROM EID WHERE KOD='" + eggtim.Recordset("kode") + "'"    ' , DB, adOpenKeyset, adLockOptimistic
                 EID.Refresh


                            PAR112.Label10 = EID.Recordset("XTI")
500
                              
                            PAR112.FPAEID.Caption = eggtim.Recordset!FPA
                            PAR112.ltiproekptosis = eggtim.Recordset("timm") * (100 - eggtim.Recordset("EKPT")) * (100 - eggtim.Recordset("EKPT2")) / 100 / 100 '************** Grid1.TextMatrix(k, F_T)
                            PAR112.Label11 = PAR112.ltiproekptosis ' ************* NeaTim

510                         '**********If Val(Synolo.Caption) <> 0 Then
520                          '***************   par112.Label4 = 100 * Val(metaf) / Val(Synolo)
                           ' *********  End If

530                         PAR112.POSO.Caption = eggtim.Recordset("poso") ' "0" '************* p
540                         PAR112.Text1(2) = Round(EID.Recordset("POS_KERD"), 3)

550                         If IsNull(EID.Recordset("LTI5")) Then
560                             PAR112.Text1(4) = 0
                            Else
                                PAR112.Text1(4) = EID.Recordset("LTI5")
                            End If

570
580                         PAR112.Text1(7) = EID.Recordset("LTI")

590                         If IsNull(EID.Recordset("pos_kerd2")) Then
600                              PAR112.Text1(6) = 0
                            Else
                                 PAR112.Text1(6) = EID.Recordset("pos_kerd2")
                            End If

610

620                         If IsNull(EID.Recordset("lti2")) Then
630                             PAR112.LTI2 = 0
                            Else
                                PAR112.LTI2 = EID.Recordset("LTI2")
                            End If
                            
                            
                            If IsNull(EID.Recordset("lti3")) Then
                                 PAR112.LTI3 = 0
                            Else
                                 PAR112.LTI3 = EID.Recordset("LTI3")
                            End If
                           
                           
                            
                            
                            
                         



640                         PAR112.Label1(1).Caption = EID.Recordset("kod") + "-" + EID.Recordset("ono")
650                         PAR112.SHOW 1

660                         If True Then '*****************kostologhsh Then  ' μην τυχον και ακυρωσε την κοστολογηση εν τω μεταξύ

'670                       ************      If Val(Synolo.Caption) <> 0 Then
'680                                 EID.Recordset("POS_KERD3") = 100 * Val(metaf) / Val(Synolo)
'                                End If

690
'    **************                            If IsNull(EID.Recordset("lastupd")) Then
'                                     EID.Recordset("lastupd") = f_dat
'                                Else
'                                   If EID.Recordset("lastupd") < f_dat Then
'                                         EID.Recordset("lastupd") = f_dat
'                                   End If
'
'                                End If
                            
                            
                                Dim m_dat As String: m_dat = Format(eggtim.Recordset("HME"), "mm/dd/yyyy")
                                
                                If IsNull(EID.Recordset("lastupd")) Then
                                     
                                     Gdb.Execute "update EID SET LASTUPD='" + m_dat + "' WHERE KOD='" + eggtim.Recordset("kode") + "'"
                                Else
                                   If EID.Recordset("lastupd") < eggtim.Recordset("HME") Then
                                         Gdb.Execute "update EID SET LASTUPD='" + m_dat + "' WHERE KOD='" + eggtim.Recordset("kode") + "'"
                                   End If
                                
                                End If
                            
                            
                               
                               Gdb.Execute "update EID SET XTI=" + Replace(PAR112.ltiproekptosis, ",", ".") + " WHERE KOD='" + eggtim.Recordset("kode") + "'"
                                 Gdb.Execute "update EID SET LTI5=" + Replace(PAR112.Text1(4), ",", ".") + " WHERE KOD='" + eggtim.Recordset("kode") + "'"
                                   Gdb.Execute "update EID SET LTI=" + Replace(PAR112.Text1(7), ",", ".") + " WHERE KOD='" + eggtim.Recordset("kode") + "'"
                                Dim POSK As Double
                                POSK = (100 * gVal(PAR112.Text1(4))) / ((100 + g_Fpa(EID.Recordset("FPA"))) / 100 * gVal(PAR112.ltiproekptosis)) - 100
                                   Gdb.Execute "update EID SET POS_KERD=" + Replace(str(POSK), ",", ".") + " WHERE KOD='" + eggtim.Recordset("kode") + "'"
        
                                   POSK = (100 * gVal(PAR112.Text1(7))) / gVal(PAR112.ltiproekptosis) - 100               '  EID.Recordset("XTI") - 100
                                    Gdb.Execute "update EID SET POS_KERD2=" + Replace(str(POSK), ",", ".") + " WHERE KOD='" + eggtim.Recordset("kode") + "'"
        
                                   Gdb.Execute "update EID SET LTI2=" + Replace(PAR112.LTI2.Text, ",", ".") + " WHERE KOD='" + eggtim.Recordset("kode") + "'"
                                       Gdb.Execute "update EID SET LTI3=" + Replace(PAR112.LTI3.Text, ",", ".") + " WHERE KOD='" + eggtim.Recordset("kode") + "'"
        
                            
700                             'EID.Recordset("XTI") = Replace(PAR112.ltiproekptosis, ",", ".") '   ************NeaTim


710                            ' EID.Recordset("LTI5") = Val(Replace(PAR112.Text1(4), ",", "."))  ' par112.Text1(4)

720                            ' EID.Recordset("LTI") = Val(Replace(PAR112.Text1(7), ",", "."))

730                           '  EID.Recordset("POS_KERD") = (100 * EID.Recordset("LTI5")) / ((100 + g_Fpa(EID.Recordset("FPA"))) / 100 * EID.Recordset("XTI")) - 100

740                            ' EID.Recordset("POS_KERD2") = (100 * EID.Recordset("LTI")) / EID.Recordset("XTI") - 100   '  EID.Recordset("XTI") - 100
                               ' If F_XONDRIKH = 3 Then
                                  ' EID.Recordset("LTI2") = Val(Replace(PAR112.LTI2.Text, ",", "."))
                                  ' EID.Recordset("LTI3") = Val(Replace(PAR112.LTI3.Text, ",", "."))
                               ' End If

                                'EID.Recordset.Update
                            End If
      
                        DoEvents
                        If PAR2.EXODOS.Caption = "1" Then
                           Exit Do
                        End If
1040        eggtim.Recordset.MoveNext

        Loop



SHOW_TIMOL
        Exit Sub

FIND_N_NUMBER:
        '************************
1050    rec.MoveFirst
1060    rec.Move 10

        'rec.Edit
1070    If IsNull(rec("epel")) Then
1080        rec("epel") = 1
        Else
1090        rec("epel") = rec("epel") + 1
        End If

1100    rec.Update
1110    mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)
1120    Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + eggtim.Recordset("KODE") + "'"
1130    Return

LATOS:

        '<EhFooter>
Resume Next
Exit Sub

cmdKOSTOLOGHSH_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command9_Click()
Dim a As String
a = ""
a = InputBox("δωσε το χαρακτηριστικό γράμμα του νέου τύπου", "..", Left(TIM.Recordset("atim"), 1))

If Len(a) = 0 Then
    MsgBox "ΛΑΘΟΣ ΠΑΡΑΣΤΑΤΙΚΟ"
    Exit Sub
End If



a = Left(a, 1)




Dim R As New ADODB.Recordset
R.Open "SELECT COUNT(*) FROM PARASTAT WHERE EIDOS='" + a + "'", Gdb, adOpenDynamic, adLockOptimistic
If R(0) = 0 Then
   MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ ΤΕΤΟΙΟ ΠΑΡΑΣΤΑΤΙΚΟ ΠΟΥ ΖΗΤΑΤΕ"
   Exit Sub
End If
R.Close



Gdb.Execute "update TIM SET ATIM='" + Left(a, 1) + mID(TIM.Recordset("atim"), 2, 7) + "' WHERE ID_NUM=" + str(TIM.Recordset("ID_NUM"))
Gdb.Execute "update EGGTIM SET ATIM='" + Left(a, 1) + mID(TIM.Recordset("atim"), 2, 7) + "' WHERE ID_NUM=" + str(TIM.Recordset("ID_NUM"))
Gdb.Execute "update EGG SET ATIM='" + Left(a, 1) + mID(TIM.Recordset("atim"), 2, 7) + "' WHERE IDTIM=" + str(TIM.Recordset("ID_NUM"))



End Sub

Private Sub Form_Paint()

    '  If CallFromOther.Caption = "1" Then
    '    Command1_Click
    '  End If
    '

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 12465, 10365, fh, fw, ft, fl
End Sub

'Private Sub Form_Load()
'End Sub

Private Sub Form_Unload(Cancel As Integer)

End Sub

 Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim D1 As Date, d2 As Date, synt As String

        'eggtim.DatabaseName = gDir---

100     D1 = DTPicker1.Value

110     d2 = DateAdd("d", 1, DTPicker2.Value)

120     TIM.ConnectionString = gConnect
        ' TIM.DatabaseName = gDir

130     If Val(Text1.Text) = 0 Then    ' ΒΓΆΖΕΙ ΟΛΑ ΤΑ ΠΑΡΑΣΤΑΤΙΚΑ ΚΑΙ ΔΙΑΛΕΓΕΙΣ
140         If Right(Combo2.Text, 1) = "*" Then
150             synt = " left(ATIM,1)>' ' "

160             If Len(Text1.Text) > 0 Then
170                 synt = " ATIM='" + Text1.Text + "'"
                End If

            Else
180             synt = " left(ATIM,1)='" + Right(Combo2.Text, 1) + "'"
            End If

        Else
190         synt = " ATIM LIKE '%" + Text1.Text + "%'"
            'If f_psifia_atim = 5 Then
            '    synt = " left(ATIM,6)='" + Right(Combo2.Text, 1) + Right("00000" + Text1.Text, 5) + "'"
            'Else
            '    synt = " left(ATIM,7)='" + Right(Combo2.Text, 1) + Right("000000" + Text1.Text, 6) + "'"
            'End If
        End If

200     If Len(Text2.Text) > 0 Then
210         synt = synt + " AND PEL.EPO LIKE '" + SameLetters(Text2.Text) + "%'"
        End If

220     If ekkrem.Value = vbChecked Then

230         synt = synt + " AND ART IS NULL "

        End If

        Dim XX As String

        On Error GoTo Command1_Click_Err

240     If Combo3.Text = Combo3.List(1) Then
250         XX = "select '-'+PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.KOD,PEL.EIDOS,TRP,SKOPOS,FORTOSH,ART,PROOR,PARAT,ART,B_N1,B_N2,SXETIKO,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,ID_NUM,ENTITYMARK,ORA   " & " from TIM LEFT join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' order by PEL.EPO,HME;"
        Else
260         XX = "select '-'+PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.KOD,PEL.EIDOS,TRP,SKOPOS,FORTOSH,ART,PROOR,PARAT,ART,B_N1,B_N2,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,ID_NUM,ENTITYMARK,ORA   " & " from TIM LEFT join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' order by HME,PEL.EPO;"
        End If

270     TIM.RecordSource = XX
280     TIM.Refresh

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '------------------------------------------------------------------------
        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim sql  As String, fDB As Database, D1 As Date, d2 As Date

        Dim synt As String, DB As Database

        Dim DUM

        Dim x As String

        On Error GoTo LATOS
        
        
        Dim MAPOT As Integer
        MAPOT = 1
        
        If IsNull(TIM.Recordset("B_N1")) Then
           MAPOT = 1
        End If
        
        

100     D1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
110     d2 = DateAdd("d", 1, D1)

120     x = "select KODE,ATIM,left(ONOMA,35) as [Περιγραφή],POSO as [Ποσότητα],TIMM as [Tιμ_Μον]," & " EKPT as [Εκπτωση],str(KAU_AJIA,10,2) as [K_Αξία],FPA AS [Κατ.ΦΠΑ],str(MIK_AJIA,10,2) as [ΜΕΦΠΑ_Αξία],PROELEYSH AS [ΜΕΓ] from EGGTIM where" & " left(ATIM," + str(f_psifia_atim + 1) + ")='" + Left(TIM.Recordset("ATIM"), f_psifia_atim + 1) + "'"
130     x = x + " and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'"

140     x = x + " AND APOT=" + Format(MAPOT, "0") + " order by PROOD_AJ;"

        Dim TIT As String

150     TIT = TIM.Recordset(0) + " Αριθ." + TIM.Recordset(1) + "  " + Format(TIM.Recordset(2), "dd/mm/yyyy")

160     DUM = print3_xar(x, "111111111", TIT, 0)

        Exit Sub

        ' προσοχη δεν πρέπει να σβήνω τους πίνακες tim,eggtim

        '-----------  next lines are OK --------------------------

        On Error Resume Next

170     D1 = D1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
180     d2 = DateAdd("d", 1, D1)

190     Set DB = OpenDatabase("c:\mercvb\reports\reports.mdb")

200     If gConnect = "Access" Then
210         Set fDB = OpenDatabase(gDir, False, False)

        Else
220         Set fDB = OpenDatabase(gDir, False, False, gConnect)
        End If

        'arn_enhm_apot
        'db.Execute "drop table eggtim"

230     sql = "select fpa,atim,format(hme,'dd/mm/yy') as [Ημερομ],kode as [Kωδ_Είδ],onoma as [Περιγραφή],poso as [Ποσότητα],timm as [Tιμ_Μον],timm*poso as [Αξία]  into eggtim in 'c:\mercvb\reports\reports.mdb' "
240     sql = sql + " from eggtim where left(atim,7)='" + TIM.Recordset("atim") + "' and hme>=#" + Format(D1, "mm/dd/yyyy") + "# and hme<#" + Format(d2, "mm/dd/yyyy") + "#;"

250     fDB.Execute sql

        On Error GoTo Command2_Click_Err

260     synt = " left(atim,7)='" + Left(eggtim.Recordset("atim"), 7) + "'"

270     sql = "select pel.epo,atim,format(hme,'dd/mm/yy') as shme,aji as ajia,pel.epa,pel.die,pel.afm into tim in 'c:\mercvb\reports\reports.mdb' "
280     sql = sql + "from tim inner join pel on tim.eidos+tim.kpe=pel.eidos+pel.kod where " + synt + " and hme>=#" + Format(D1, "mm/dd/yyyy") + "# and hme<#" + Format(d2, "mm/dd/yyyy") + "#  order by prood_aj;"

        On Error Resume Next

        'db.Execute "drop table tim"

290     fDB.Execute sql

        'CrystalReport1.Action = 1

        Exit Sub

LATOS:

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '-----------------------διαγραφή παραστατικού ---------------
        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>
        Dim D1     As Date, d2 As Date

        Dim fDB    As Database

        Dim id_num As Long

        Dim MATIM  As String

        Dim ANS    As Integer, sql As String

        On Error GoTo LATOS

100     D1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
110     d2 = DateAdd("d", 1, D1)
        id_num = TIM.Recordset("id_num")
        
        If id_num = 0 Then
            MsgBox "Προσοχη!! Αδύνατη η διαγραφή λόγω έλλειψης ID_NUM"

            Exit Sub

        End If
        
        Dim mmark As String

        mmark = CNull(TIM.Recordset("ENTITYMARK"))
        '         If Len(MARK) > 2 Then
        '            MsgBox "Προσοχη!! Αδύνατη η διαγραφή εχει ήδη πάρει ΜΑΡΚ(MYDATA)"
        '            Exit Sub
        '         End If
         
        If Len(Trim(mmark)) > 2 Then
            If Left$(mmark, 2) = "ER" Then
      
            Else
                MsgBox "Aδύνατη η διαγραφή γιατί έχει πάρει MARK"

                Exit Sub

            End If
      
        End If
        
        'Set db = OpenDatabase("c:\mercvb\reports\reports.mdb")
        'If gConnect = "Access" Then
        '  Set fDB = OpenDatabase(gDir, False, False)

        'Else
        '  Set fDB = OpenDatabase(gDir, False, False, gConnect)
        'End If
        '------------------------
        Dim g_hme, g_atim, DUM

120     g_atim = TIM.Recordset("ATIM")

130     g_hme = TIM.Recordset("sHME")

140     ANS = MsgBox("Είσαι σίγουρος για την διαγραφή του παραστατικού " + g_atim, vbYesNo)

150     If ANS = vbNo Then Exit Sub
        'On Error Resume Next

160     If Len(gdirlog) > 0 Then
170         DUM = del_Gefyroma(g_hme, g_atim, 1)

180         If TIM.Recordset("trp") = "ΜΕ" Then
190             DUM = del_Gefyroma(g_hme, g_atim, 2)
            End If
        End If

200     MATIM = TIM.Recordset("atim")

        ' DEL GEFYROSI
        ' ENHM_APOT
        ' ENHM_PEL
210     DoEvents

        On Error Resume Next
    
220     arn_enhm_pel
        'Exit Sub

230     DoEvents
        'arn_enhm_apot

        'Exit Sub
        Dim N As Long

240     sql = "delete  "
250     sql = sql + " from EGGTIM where  ID_NUM=" + str(id_num) ' left(ATIM,7)='" + Left(MATIM, 7) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"
260     Gdb.Execute sql, N
        
        If N = 0 Then  'ΠΑΛΙΟΤΕΡΗ ΕΚΔΟΣΗ ΤΙΜΟΛΟΓΙΟΥ ΠΟΥ ΔΕΝ ΕΒΑΖΕ ID_NUM STO EGGTIM
            sql = "delete  "
            sql = sql + " from EGGTIM where  left(ATIM,7)='" + Left(MATIM, 7) + "' and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"
            Gdb.Execute sql, N
        End If

270     DoEvents

280     sql = "delete  "
290     sql = sql + " from TIM where ID_NUM=" + str(id_num)
        'Left(ATIM,7)='" + Left(MATIM, 7) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"

300     Gdb.Execute sql

310     DoEvents

        If f_IS_PARTIDES = 1 Then

            Dim rpart    As New ADODB.Recordset

            Dim DEL_PART As String, COUNT_DEL As Integer: DEL_PART = "": COUNT_DEL = 0

            Dim AFAIR    As Single

            Dim NC3      As Long

            rpart.Open "select *  FROM PARTIDES WHERE ID_NUM=" + str(id_num), Gdb, adOpenDynamic, adLockOptimistic

            Do While Not rpart.EOF
                DEL_PART = DEL_PART + CNull(rpart!LOTNUMBER) + Chr(13)
                AFAIR = rpart!POSOAG
                COUNT_DEL = COUNT_DEL + 1
          
                ' ΑΝ Η ΠΑΡΤΙΔΑ ΕΙΝΑΙ ΑΠΟ ΤΙΣ ΠΩΛΗΣΕΙΣ ΝΑ ΞΑΝΑΠΡΟΣΘΕΤΕΙ ΤΟ ΥΠΟΛΟΙΠΟ
                If AFAIR < 0 Then
                    Gdb.Execute "UPDATE PARTIDES SET YPOLOIPO=YPOLOIPO+" + Replace(str(Abs(AFAIR)), ",", ".") + " WHERE ID=(SELECT TOP 1 ID FROM PARTIDES WHERE LOTNUMBER='" + CNull(rpart!LOTNUMBER) + "' AND POSOAG>0 AND YPOLOIPO>" + Replace(str(Abs(AFAIR)), ",", ".") + ")", NC3
                End If

                rpart.MoveNext
            Loop

            If COUNT_DEL > 0 Then
                MsgBox "Προσοχη !!! θα διαγραφουν οι ΑΝΤΙΣΤΟΙΧΕΣ ΕΓΓΡΑΦΕΣ παρτιδων :" + Chr(13) + DEL_PART
                Gdb.Execute "delete FROM PARTIDES WHERE ID_NUM=" + str(id_num)
            End If

            rpart.Close
        End If

        'ΕΝΗΜΕΡΩΣΗ ΜΕΤΡΗΤΗ ΕΦΟΣΟΝ ΕΙΝΑΙ ΤΟ ΤΕΛΕΥΤΑΙΟ ΠΟΥ ΚΟΠΗΚΕ
        Dim r1    As New ADODB.Recordset

        Dim AR    As Long

        Dim TELAR As Long

        r1.Open "select *  FROM PARASTAT WHERE EIDOS='" + Left(MATIM, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic

        If Not r1.EOF Then
            AR = nNull(r1("ARITMISI"))
            r1.Close

            If AR > 0 Then
                r1.Open "select *  FROM ARITMISI WHERE ID=" + str(AR), Gdb, adOpenDynamic, adLockOptimistic
                TELAR = nNull(r1("ARITMISI"))
                r1.Close

                If TELAR = Val(mID(MATIM, 2, 7)) Then
                    Gdb.Execute "UPDATE ARITMISI SET ARITMISI=ARITMISI-1 WHERE ID=" + str(AR)
                    MsgBox "ΕΝΗΜΕΡΩΘΗΚΕ ΚΑΙ Ο ΤΕΛΕΥΤΑΙΟΣ ΑΡΙΘΜΟΣ ΣΤΟ " + str(TELAR - 1)
            
                End If ' TELAR
       
            End If  ' AR>0
        End If ' R1 EOF

320     TIM.Refresh

        Exit Sub

LATOS:

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

        'F_1ST_CHOICE = Val(FindParametroi(1,"PAR2", "F_1ST_CHOICE", "2", "Πρoεπιλεγμένο παραστατικό")) 'posa psifia tha exei h kathe seira
        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>

        Dim DUM

100     DUM = UpdateParametroi("PAR2", "F_1ST_CHOICE", Format(1 + Val(Left(Combo2.Text, 2)), "##"), "Πρoεπιλεγμένο παραστατικό")    'posa psifia tha exei h kathe seira

        'αν το καλουν απο αλλού αστο
110     If CallFromOther.Caption = "1" Then
120         PAR1.LASTTIMOL.Refresh
130         Unload Me
        Else
140         Unload Me
150         SendKeys "%"
160         SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}"
        End If

        'Combo2.ListIndex
        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        '========================etiketes ========================================================
        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>
        Dim D1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        Dim DB As Database, DUM, rec As New ADODB.Recordset

        'On Error Resume Next
        Dim ar_per

100     ar_per = F_ArPerEtik    ' ρυθμιζομενο

        Dim x As Printer

110     For Each x In Printers

120         If InStr(x.DeviceName, f_printer) > 0 And InStr(x.DeviceName, "408") = 0 Then
                ' Set printer as system default.
130             Set Printer = x

                ' Stop looking for a printer.
                Exit For

            End If

        Next

        'eggtim.DatabaseName = gDir
        On Error GoTo LATOS

140     D1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
150     d2 = DateAdd("d", 1, D1)

160     TIM.ConnectionString = gConnect

170     eggtim.ConnectionString = gConnect

180     eggtim.RecordSource = "select KODE,POSO,TIMM,EID.ONO,ROUND(ISNULL(EID.LTI5,0),2) AS LTI5,ROUND(EID.LTI,2) AS LTI,ROUND(EID.XTI,2) AS XTI,EID.ERG,EID.CH4,PROELEYSH,KOLA,ISNULL(MON,'') AS MON,EGGTIM.FPA,PELKOD,ATIM,HME from EGGTIM inner join EID on EGGTIM.KODE=EID.KOD where left(ATIM," + str(1 + f_psifia_atim) + ")='" + Left(TIM.Recordset("atim"), 1 + f_psifia_atim) + "' and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "'  order by PROOD_AJ ;"

190     eggtim.Refresh
        'Label5.Caption = TIM.Recordset("epo")
        'Label2.Caption = TIM.Recordset("die")
        'Label3.Caption = TIM.Recordset("epa")
        'Label4.Caption = TIM.Recordset("afm")

        'DUM = OpenD(db)
        'Set rec = db.OpenRecordset("select *from mem")

        On Error GoTo Command5_Click_Err

200     rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic

210     If eggtim.Recordset.EOF Then
220         MsgBox "ΔΕΝ ΕΥΡΕΘΗΣΑΝ ΕΓΓΡΑΦΕΣ "

            Exit Sub

        End If

        Dim nkod

230     eggtim.Recordset.MoveFirst

        Dim mlabel, MPER, mtimes, TIMM, timologio

240     Do While Not eggtim.Recordset.EOF

250         If IsNull(eggtim.Recordset("ERG")) Then
260             GoSub FIND_N_NUMBER
            Else

270             If Val(eggtim.Recordset("ERG")) = 0 Then
280                 GoSub FIND_N_NUMBER
                Else
290                 mlabel = eggtim.Recordset("erg")
                End If
            End If

300         MPER = eggtim.Recordset("ono")
310         nkod = eggtim.Recordset("kode")

320         If noask.Value = vbChecked Then
330             mtimes = eggtim.Recordset("poso")
            Else
            
                If Len(Dir("C:\MERCVB\MELIS.TXT", vbNormal)) = 0 Then
                   mtimes = InputBox("Πόσες ετικέτες;", MPER, eggtim.Recordset("poso"))
                Else
                      mtimes = InputBox("Πόσες ετικέτες;", MPER, eggtim.Recordset("KOLA"))
                End If
            
'340             mtimes = InputBox("Πόσες ετικέτες;", MPER, EGGTIM.Recordset("poso"))
            End If

350         TIMM = Format(eggtim.Recordset("LTI5"), "####0.00")
360         timologio = eggtim.Recordset("kode") + Left(TIM.Recordset("atim"), 6) + "=" + eggtim.Recordset("pelkod") + Format(eggtim.Recordset("xti") * (100 + g_Fpa(2)) / 100, "####.00") + "-" + mID(Format(eggtim.Recordset("hme"), "yyyyMMDD"), 3, 6) + "-" + Format(eggtim.Recordset("POSO"), "##0")

            Dim EPIPLEON As Long

370         EPIPLEON = 0

380         For k = 1 To Val(mtimes)
                 

390             If Check5 Then  ' lag

400                 If Len(Dir("C:\LAGEURO\LAGAKIS2.TXT")) Then
410                    '''''''''''''''''''''''''''''''''' ar_per = 5
                       ''''''''''''''''''''''''''''''' ' BarCode2.SymbologyID = CODE128
420                  '   BarCode2.DataToEncode = mlabel
430                   '  Printer.PaintPicture BarCode2.Picture, 1800, 1050, 2000, 400      ' ORIZ,KATHETA

                     
                     
                     
                           If Len(mlabel) = 12 Then
                              BarCode3.SymbologyID = UPCA
                           ElseIf Len(mlabel) = 13 Then
                              BarCode3.SymbologyID = EAN13
                           Else
                              BarCode3.SymbologyID = CODE128
                           End If
                     
                     
                     
                     
                     
                     BarCode3.DataToEncode = mlabel
                     Printer.PaintPicture BarCode3.Picture, 1800, 1050, 2000, 400      ' ORIZ,KATHETA





440                     EPIPLEON = 3

                    Else
450                     Printer.FontName = "128"
460                     Printer.FontSize = 5
470                     Printer.CurrentY = 0
480                     Printer.Print Tab(10); mlabel
                    End If

490                 Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName 12305019

500                 Printer.FontSize = 7
                    '  Printer.Print
                    '  Printer.Print
                    ' Printer.Print
510                 Printer.CurrentY = 110
520                 Printer.Print

530                 Printer.Print Tab(ar_per + EPIPLEON); mlabel
540                 Printer.Print Tab(ar_per + EPIPLEON); MPER
550                 Printer.Print Tab(ar_per + EPIPLEON); timologio
560                 Printer.FontSize = 16
570                 Printer.FontBold = True
580                 Printer.Print Tab(ar_per); Trim(LTrim(TIMM)) + " €  /" + eggtim.Recordset("MON")

590                 Printer.EndDoc

600             ElseIf Len(Dir("C:\MERCVB\MELIS.TXT", vbNormal)) > 0 Then

610                 Printer.FontSize = 9
620                 Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
630                 Printer.FontBold = True
640                 Printer.CurrentY = 6  '12
                    Printer.Print Tab(1); "ΜΕΛΙΣΣΑΡΙΔΗΣ Ν&ΣΙΑ ΟΕ"
                    Printer.Print Tab(1); "ΛΑΧΑΝΑΓΟΡΑ ΔΡΑΜΑΣ"
                    Printer.Print Tab(1); "ΤΗΛ.2521031987"
650                 Printer.Print Tab(1); "ΕΙΔΟΣ: " + MPER
                    Printer.Print Tab(1); "ΠΡΟΕΛΕΥΣΗ    ΠΑΡΤIΔA :"
                    Printer.Print Tab(1); eggtim.Recordset("PROELEYSH")
                    Printer.Print Tab(1); "ΠΟΙΟΤΗΤΑ: Ι "
                    Printer.Print Tab(1); "GR " + eggtim.Recordset("CH4")
                 '   Printer.Print Tab(0); "GR" + EGGTIM.Recordset("CH4")

'660                 BarCode1.ShowText = NO
'670                 BarCode1.NarrowBarWidth = 0.025  '0.03
'680                 BarCode1.DataToEncode = mlabel
'690                 Printer.PaintPicture BarCode1.Picture, -18, 124   '-88,124
'700                 Printer.Print ; " "; mlabel
'710                 Printer.FontSize = 8
'720                 Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
'730                 Printer.CurrentY = 900
'740                 Printer.Print Tab(5); Format(Val(TIMM), "###0.00") + " ΕΥΡΩ "
750                 Printer.EndDoc



                   ElseIf Len(Dir("C:\MERCVB\EFRAIM.TXT")) Then

 
4000                 If Len(Dir("C:\MERCVB\EFRAIM2.TXT")) Then
                           If Len(mlabel) = 12 Then
                              BarCode3.SymbologyID = UPCA
                           ElseIf Len(mlabel) = 13 Then
                              BarCode3.SymbologyID = EAN13
                           Else
                              BarCode3.SymbologyID = CODE128
                           End If

                           BarCode3.DataToEncode = mlabel
                           Printer.PaintPicture BarCode3.Picture, F_ET_BAR_X, F_ET_BAR_Y, 1050, F_ET_BAR_H, 800   ' ORIZ,KATHETA
4400                       EPIPLEON = 3
                    Else
4500                     Printer.FontName = "128"
4600                     Printer.FontSize = 5
4700                     Printer.CurrentY = 0
4800                     Printer.Print Tab(10); mlabel
                    End If
4900                 Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName 12305019
5000                 Printer.FontSize = 7
5100                 Printer.CurrentY = 110
5200                 Printer.Print
5300                 Printer.Print Tab(ar_per + EPIPLEON); mlabel
5400                 Printer.Print Tab(ar_per + EPIPLEON); MPER
5500                 Printer.Print Tab(ar_per + EPIPLEON); timologio
5600                 Printer.FontSize = 16
5700                 Printer.FontBold = True
5800                 Printer.Print Tab(ar_per); Trim(LTrim(TIMM)) + " €  /" + eggtim.Recordset("MON")
5900                 Printer.EndDoc

760             ElseIf Len(Dir("C:\LAGEURO\LOISIR.TXT")) Then

                    ' 12345
770                 Printer.FontSize = 6   '10000011  10000012

780                 Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName00000002
790                 Printer.FontBold = True
                    ' Printer.CurrentY = 120
800                 BarCode1.ShowText = NO
810                 BarCode1.NarrowBarWidth = 0.025  '0.03
820                 BarCode1.DataToEncode = mlabel    ' "12345678" '  EGGTIM.Recordset("kode")
                    '  Printer.PaintPicture BarCode1.Picture, -100, 30, , 470 '-18,124
830                 Printer.PaintPicture BarCode1.Picture, -130, 30, , 470    '-18,124

840                 Printer.FontSize = 6    ' 10000020

850                 Printer.CurrentY = 400    '546  ΠΑΝΟ ΣΤΗΝ ΕΤΙΚΕΤΑ   '
860                 Printer.Print Tab(7); eggtim.Recordset("kode")

870                 Printer.FontSize = 7
880                 Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
890                 Printer.CurrentY = 480    '420 OK ΠΑΝΟ ΣΤΗΝ ΕΤΙΚΕΤΑ
                    ' Printer.Print

900                 Printer.Print Tab(6); Format(Val(TIMM), "###0.00") + "€ "

910                 Printer.EndDoc

                    '  Printer.FONTSIZE = 6
                    '  Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName00000002
                    '  Printer.FontBold = True
                    '  Printer.CurrentY = 120
                    '  BarCode1.ShowText = YES
                    '  BarCode1.NarrowBarWidth = 0.025  '0.03
                    '  BarCode1.DataToEncode = "12345678" '  EGGTIM.Recordset("kode")
                    '  Printer.PaintPicture BarCode1.Picture, -100, 0 '-18,124
                    ''  Printer.Print ; " "; mlabel
                    '  Printer.FONTSIZE = 8
                    '  Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
                    '  Printer.CurrentY = 100
                    '  Printer.Print Tab(5); Format(Val(TIMM), "###0.00") + " ΕΥΡΩ "
                    '  Printer.EndDoc

                Else

920                 Printer.Print Tab(20); MPER
930                 Printer.Print Tab(30); mlabel
940                 Printer.FontName = "ean13"    ' "128"
950                 Printer.FontSize = 6
960                 Printer.Print Tab(18); eggtim.Recordset("kode")    ' mlabel
970                 Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
980                 Printer.FontSize = 8
990                 Printer.Print
1000                Printer.Print
1010                Printer.Print
1020                Printer.Print
1030                Printer.EndDoc
                End If

            Next

1040        eggtim.Recordset.MoveNext

        Loop

        Exit Sub

FIND_N_NUMBER:
        '************************
1050    rec.MoveFirst
1060    rec.Move 10

        'rec.Edit
1070    If IsNull(rec("epel")) Then
1080        rec("epel") = 1
        Else
1090        rec("epel") = rec("epel") + 1
        End If

1100    rec.Update
1110    mlabel = Right("00000000" + LTrim(str(rec("EPEL"))), 8)
1120    Gdb.Execute "UPDATE EID SET ERG='" + mlabel + "' WHERE KOD='" + eggtim.Recordset("KODE") + "'"
1130    Return

LATOS:

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>

100     eggtim.Recordset.MoveFirst

        Dim a, B

110     a = InputBox("φακελος με τις φωτογραφιες π.χ. c:\mercvb ")
120     B = InputBox("Αρχικός αριθμός")

        '
        '
        Dim PI

        'PI = "C:\MERCVB\IMAGES\" + Text1(0).Text
        'If Len(Dir(PI)) > 0 Then
        '   Picture1.Picture = LoadPicture(PI)
        'Else
        '   Picture1.Picture = LoadPicture()
        'End If

        Dim ANS, F

130     Do While Not eggtim.Recordset.EOF

140         PI = a + "\" + Format(B, "0000") + ".jpg"

150         If Len(Dir(PI)) > 0 Then
160             Picture1.Picture = LoadPicture(PI)
            Else

170             Me.Caption = "Αναζήτηση φωτογραφίας του είδους :" + eggtim.Recordset(3) + eggtim.Recordset(0)
180             CD1.InitDir = a
190             CD1.ShowSave
200             Picture1.Picture = LoadPicture(CD1.FILENAME)
210             PI = CD1.FILENAME
            End If

220         ANS = MsgBox(Left(eggtim.Recordset(3), 25) + Chr(13) + mID$(eggtim.Recordset(3), 26, 15), vbYesNo, "Να αποθηκευτεί η φωτογραφία")

230         If Len(Dir("C:\MERCVB\IMAGES\" + eggtim.Recordset("kode"))) > 0 Then
240             Kill "C:\MERCVB\IMAGES\" + eggtim.Recordset("kode")
            End If

250         If ANS = vbYes Then
                ' A + "\" + Format(B, "0000") + ".jpg"
260             Name PI As "C:\MERCVB\IMAGES\" + eggtim.Recordset("kode")

            Else
270             CD1.InitDir = a
280             CD1.ShowSave
                '       Picture1.Picture = LoadPicture(CD1.FILENAME)
            End If

290         B = B + 1
300         eggtim.Recordset.MoveNext
        Loop

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command7_Click()

        '----------------------βγαζο απο ekkremothta ---------------
        '<EhHeader>
        On Error GoTo Command7_Click_Err

        '</EhHeader>
        Dim D1    As Date, d2 As Date

        Dim fDB   As Database

        Dim MATIM As String

        Dim ANS   As Integer, sql As String

        On Error GoTo LATOS

100     D1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
110     d2 = DateAdd("d", 1, D1)

        Dim g_hme, g_atim, DUM

120     g_atim = TIM.Recordset("ATIM")

130     g_hme = TIM.Recordset("sHME")

140     MATIM = TIM.Recordset("atim")

150     sql = "UPDATE TIM SET ART='1' where left(ATIM," + Format(f_psifia_atim + 1, "0") + ")='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"

        Dim k As Integer

160     Gdb.Execute sql, k

170     If k = 1 Then TIMGrid.TextMatrix(5, 0) = "ΤΙΜΟΛΟΓΗΜΕΝΟ"

180     DoEvents

        Dim x

190     x = TIM.Recordset.Bookmark
200     TIM.Refresh
210     TIM.Recordset.Bookmark = x

        Exit Sub

LATOS:

        '<EhFooter>
        Exit Sub

Command7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Command7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Command7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command8_Click()

        '<EhHeader>
        On Error GoTo Command8_Click_Err

        '</EhHeader>

        '----------------------epanafero se ekkremothta ---------------
        Dim D1    As Date, d2 As Date

        Dim fDB   As Database

        Dim MATIM As String

        Dim ANS   As Integer, sql As String

        On Error GoTo LATOS

100     D1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))
110     d2 = DateAdd("d", 1, D1)

        Dim g_hme, g_atim, DUM

120     g_atim = TIM.Recordset("ATIM")

130     g_hme = TIM.Recordset("sHME")

140     MATIM = TIM.Recordset("atim")

150     sql = "UPDATE TIM SET ART=NULL where left(ATIM," + Format(f_psifia_atim + 1, "0") + ")='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(D1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"

        Dim k As Integer

160     Gdb.Execute sql, k

170     If k = 1 Then TIMGrid.TextMatrix(5, 0) = "ΜΗ ΤΙΜΟΛΟΓΗΜΕΝΟ"

180     DoEvents

        Dim x

190     x = TIM.Recordset.Bookmark
200     TIM.Refresh
210     TIM.Recordset.Bookmark = x

        Exit Sub

LATOS:

        '<EhFooter>
        Exit Sub

Command8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Command8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Command8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DBGrid1_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = 13 Then
110         DBGrid1_RowColChange 0, 0
        End If

        '<EhFooter>
        Exit Sub

DBGrid1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.DBGrid1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.DBGrid1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     SHOW_TIMOL
     End Sub
     
     
Sub SHOW_TIMOL()

        '<EhHeader>
        On Error GoTo DBGrid1_RowColChange_Err

        '</EhHeader>
        Dim D1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        On Error Resume Next

        'eggtim.DatabaseName = gDir
100     D1 = CDate(Format(TIM.Recordset("shme"), "dd/mm/yyyy"))  ' DTPicker1.Value
110     d2 = DateAdd("d", 1, D1)
        'TIM.DatabaseName = gDir
120     TIM.ConnectionString = gConnect
130     eggtim.ConnectionString = gConnect

        On Error GoTo DBGrid1_RowColChange_Err

        Dim sql As String

        Dim RDB As New ADODB.Recordset

140     RDB.Open "select * from PARASTAT where EIDOS='" + Left(TIM.Recordset("atim"), 1) + "';", Gdb, adOpenForwardOnly, adLockReadOnly

150     If TIM.Recordset.EOF Then Exit Sub

160     sql = "select KODE,ONOMA as [Περιγραφή],str(POSO,10,2) as [Ποσότητα]," & " STR(ROUND(TIMM,2),10,2) as [Tιμ_Μον],str(TIMM*POSO,10,2) as [Αξία], STR(EKPT,10,2)  as [Εκπτωση],EKPT2 AS [ΕΚΠ.2η],(SELECT TOP 1 TIMH FROM PINAKES WHERE TYPOS=1 AND AYJON=EGGTIM.FPA) AS [%ΦΠΑ]," & "str(TIMM*POSO*(100-EKPT)/100*(CASE WHEN EKPT2 IS NULL THEN 1 ELSE (100-EKPT2)/100 END),10,2) as [Αξία μετά ΕΚΠ],STR(ROUND(EID.XTI,2),10,2) AS [ΤΙΜ.ΑΓ],STR(ROUND(LTI,2),10,2) AS [ΧΟΝΔ.ΤΙΜΗ],STR(ROUND(LTI5,2),10,2) AS [ΛΙΑΝ.ΤΙΜΗ],PROELEYSH AS [ΜΕΓ],ATIM,HME as [Ημερομ] from EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD " & "    "                    ' left(ATIM," + str(f_psifia_atim + 1) + " )='" + Left(TIM.Recordset("ATIM"), 1 + f_psifia_atim)
170     sql = sql + " where  ID_NUM=" + str(TIM.Recordset("ID_NUM"))                    ' "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy")

180     If IsNull(TIM.Recordset("B_N1")) Then
190         sql = sql + "  AND APOT=1 ORDER BY PROOD_AJ;"
        Else
200         sql = sql + "  AND APOT=" + Format(TIM.Recordset("B_N1"), "0") + "  order by PROOD_AJ;"
        End If

210     eggtim.RecordSource = sql
220     eggtim.Refresh


'Exit Sub

        On Error Resume Next

        'Label13.Caption = TIM.Recordset("B_N1")
        'Label14.Caption = TIM.Recordset("B_N2")
        'Label5.Caption = TIM.Recordset("epo")
        'Label2.Caption = TIM.Recordset("die")
        'Label3.Caption = TIM.Recordset("epa")
        'Label4.Caption = TIM.Recordset("afm")
        'If IsNull(TIM.Recordset("SKOPOS")) Then Label10.Caption = "" Else Label10.Caption = TIM.Recordset("SKOPOS")
        'If IsNull(TIM.Recordset("PROOR")) Then label11.Caption = "" Else label11.Caption = TIM.Recordset("PROOR")
        'If IsNull(TIM.Recordset("ART")) Then Label12.Caption = "- " Else Label12.Caption = "ΤΙΜΟΛΟΓΗΘΗΚΕ"
        'Label11.Caption = TIM.Recordset("PROOR")

        'If IsNull(TIM.Recordset("trp")) Then Text3.Text = "" Else Text3.Text = TIM.Recordset("trp")
        'If IsNull(TIM.Recordset("SKOPOS")) Then Text4.Text = "" Else Text4.Text = TIM.Recordset("SKOPOS")
        'If IsNull(TIM.Recordset("fortosh")) Then Text5.Text = "" Else Text5.Text = TIM.Recordset("fortosh")
        'If IsNull(TIM.Recordset("PROOR")) Then Text6.Text = "" Else Text6.Text = TIM.Recordset("PROOR")

230     TIMGrid.ColWidth(0) = 4000
240     TIMGrid.ColAlignment(0) = 1

250     TIMGrid.TextMatrix(0, 0) = TIM.Recordset("epo")
260     TIMGrid.TextMatrix(1, 0) = TIM.Recordset("die")
270     TIMGrid.TextMatrix(2, 0) = TIM.Recordset("epa")
280     TIMGrid.TextMatrix(3, 0) = TIM.Recordset("afm")
290     TIMGrid.TextMatrix(4, 0) = Format(TIM.Recordset("SHME"), "DD/MM/YYYY") + "    " + TIM.Recordset("ora") + "  " + TIM.Recordset("ATIM")

300     If RDB("STADIO") <= 3 Then
310         Command7.Enabled = True
320         Command8.Enabled = True

330         If Left(TIM.Recordset("ART"), 1) = "1" Then
340             TIMGrid.TextMatrix(5, 0) = "ΤΙΜΟΛΟΓΗΜΕΝΟ "

350             If IsNull(TIM.Recordset("sxetiko")) Then
                Else
360                 TIMGrid.TextMatrix(5, 0) = "ΤΙΜΟΛΟΓΗΜΕΝΟ " + TIM.Recordset("sxetiko")
                End If

            Else
370             TIMGrid.TextMatrix(5, 0) = "ΜΗ ΤΙΜΟΛΟΓΗΜΕΝΟ"
            End If

        Else
380         Command7.Enabled = False
390         Command8.Enabled = False
400         TIMGrid.TextMatrix(5, 0) = ""
        End If

410     TIMGrid.ColWidth(1) = 16000
420     TIMGrid.ColAlignment(1) = 1

430     TIMGrid.TextMatrix(0, 1) = "Πληρωμή : " + TIM.Recordset("trp")
440     TIMGrid.TextMatrix(1, 1) = TIM.Recordset("SKOPOS")
450     TIMGrid.TextMatrix(2, 1) = TIM.Recordset("fortosh")
460     TIMGrid.TextMatrix(3, 1) = TIM.Recordset("PROOR")
470     TIMGrid.TextMatrix(4, 1) = "Από Αποθήκη " + Format(TIM.Recordset("B_N1"), "##") + " " + "Σε  Αποθήκη " + Format(TIM.Recordset("B_N2"), "##")
480     TIMGrid.TextMatrix(5, 1) = TIM.Recordset("parat")

        ' Label8.Caption = TIM.Recordset("ATIM") + " " + Format(TIM.Recordset("SHME"), "DD/MM/YYYY")

490     MSFlexGrid1.ColWidth(0) = 1500
500     MSFlexGrid1.ColWidth(1) = 4000
510     MSFlexGrid1.ColAlignment(3) = 1
520     synola

        '<EhFooter>
        Exit Sub

DBGrid1_RowColChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.DBGrid1_RowColChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.DBGrid1_RowColChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub synola()

        ' μορφοποίηση του MSFLEXGRID1
        '<EhHeader>
        On Error GoTo SYNOLA_Err

        '</EhHeader>
        Dim S2

        Dim D1 As Date, d2 As Date, R As Integer, c As Integer, k As Integer, s As Single

        'On Error Resume Next
100     MSFlexGrid1.TextMatrix(0, 1) = "Αρ.Τιμ"
110     MSFlexGrid1.AddItem ("" & Chr(9))    ' & "Σύνολο" & Chr(9) & eis & Chr(9) & pl & Chr(9) & Okeis & Chr(9) + Okpl)

        On Error GoTo SYNOLA_Err

120     R = MSFlexGrid1.rows
130     c = MSFlexGrid1.ColS
140     s = 0
        S2 = 0
        Dim SMETA As Single
150     For k = 1 To R - 1
160         MSFlexGrid1.TextMatrix(k, 12) = Format(MSFlexGrid1.TextMatrix(k, 12), "dd/mm/yyyy")
            ' MSFlexGrid1.TextMatrix(k, 5) = Format(MSFlexGrid1.TextMatrix(k, 5), "######.00")
170        ' MSFlexGrid1.TextMatrix(k, 6) = Format(Val(MSFlexGrid1.TextMatrix(k, 6)), "#####.00")
180         s = s + Val(Replace(MSFlexGrid1.TextMatrix(k, 6), ",", "."))
            S2 = S2 + Val(MSFlexGrid1.TextMatrix(k, 4))
            SMETA = SMETA + Val(MSFlexGrid1.TextMatrix(k, 8))
            
            
            
        Next

190     MSFlexGrid1.TextMatrix(R - 1, 6) = Format(s, "#####.00")
        MSFlexGrid1.TextMatrix(R - 1, 8) = Format(SMETA, "#####.00")


        MSFlexGrid1.row = R - 1
        MSFlexGrid1.Col = 8
        MSFlexGrid1.CellBackColor = vbGreen
        MSFlexGrid1.TextMatrix(R - 1, 4) = Format(S2, "#####.00")
200     MSFlexGrid1.ColAlignment = 6
210     MSFlexGrid1.AddItem ("" & Chr(9))
        'MSFlexGrid1.TextMatrix(R, c - 1) = Format(TIM.Recordset("ajia") - s, "#####.00")
        ' MSFlexGrid1.TextMatrix(R, c - 2) = "ΦΠΑ"
220     MSFlexGrid1.AddItem ("" & Chr(9))

        MSFlexGrid1.row = R + 1
        MSFlexGrid1.Col = 8
        MSFlexGrid1.CellBackColor = vbGreen

230     MSFlexGrid1.TextMatrix(R + 1, 8) = Format(TIM.Recordset("ajia"), "#####.00")
        MSFlexGrid1.TextMatrix(R + 1, 3) = "Συν.Αξία"
     
        MSFlexGrid1.TextMatrix(R, 8) = Format(TIM.Recordset("FPA6") + TIM.Recordset("FPA7") + TIM.Recordset("FPA1") + TIM.Recordset("FPA2") + TIM.Recordset("FPA3") + TIM.Recordset("FPA4"), "#####.00")
        MSFlexGrid1.TextMatrix(R, 3) = "Αξία ΦΠΑ"
        'FPA1 , FPA2, FPA3, FPA4, FPA6, FPA7, FPA8, FPA9
        
240     MSFlexGrid1.ColAlignment(3) = 1

        '<EhFooter>
        Exit Sub

SYNOLA_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.SYNOLA " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.SYNOLA " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub dtpicker1_LostFocus()

        '<EhHeader>
        On Error GoTo DTPicker1_LostFocus_Err

        '</EhHeader>
100     If DTPicker2.Value < DTPicker1.Value Then
110         DTPicker2.Value = DTPicker1.Value
        End If

        '<EhFooter>
        Exit Sub

DTPicker1_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.DTPicker1_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.DTPicker1_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Activate()

        '<EhHeader>
        On Error GoTo Form_Activate_Err

        '</EhHeader>

        Dim x  As String, k As Integer

        Dim DB As Database

        Dim R  As New ADODB.Recordset
        
        
        f_IS_PARTIDES = Val(FINDPARAMETROI(1, "APOT2", "f_IS_PARTIDES", "0", "Αν Χρησιμοποιώ παρτίδες=1"))

100     f_psifia_atim = Val(FINDPARAMETROI(1, "PAR1", "F_PSIFIA_ATIM", "5", "ΨΗΦΙΑ ΤΙΜΟΛΟΓΙΟΥ (5)"))
        ' db.Open gDir

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

110     TIMGrid.ColWidth(0) = 4000
120     TIMGrid.ColWidth(1) = 16000

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If
130     F_1ST_CHOICE = Val(FINDPARAMETROI(1, "PAR2", "F_1ST_CHOICE", "2", "Πρoεπιλεγμένο παραστατικό"))   'posa psifia tha exei h kathe seira
140     F_ArPerEtik = Val(FINDPARAMETROI(1, "PAR2", "F_ArPerEtik", "2", "Αριστερό περιθώριο ετικέττας")) 'posa psifia tha exei h kathe seira


        f_printer = FINDPARAMETROI(1, "APOT2", "F_PRINTER", "SATO", "ΜΕΡΟΣ ΤΗΣ ΟΝΟΜΑΣΙΑΣ ΤΟΥ ΕΚΤΥΠ.ΕΤΙΚΕΤΤΩΝ")
  

150     R.Open "select POL,TITLOS,EIDOS,STADIO from PARASTAT ORDER BY POL,TITLOS;", Gdb, adOpenDynamic, adLockOptimistic

        'Combo2.Clear
160     If R.RecordCount = 0 Then

            Exit Sub

        End If

        Dim L As Integer, FirstTime As Boolean

        Dim o As ComboItem

170     Combo3.Text = Combo3.List(1)

180     L = 0
190     FirstTime = False

200     Set o = Combo2.ComboItems.Add(1, "", "ΟΛΑ ΤΑ ΠΑΡΑΣΤΑΤΙΚΑ     *")
        ' Combo2.ComboItems(K).Image = 2

210     R.MoveFirst
220     k = 1

230     Do While Not R.EOF

240         If Not IsNull(R("TITLOS")) And Left(R("TITLOS"), 1) <> "-" Then
250             If R("POL") = 2 Then  ' αγορες
260                 If FirstTime = False Then L = k: FirstTime = True

                    ' Combo2.AddItem Format(K, "##") + "---" + " " + R("TITLOS") + Space(1) + R("eidos")
                    ' Combo2.ItemBackColor(K) = vbRed

270                 Set o = Combo2.ComboItems.Add(k + 1, "", Format(k, "##") + "---" + " " + R("TITLOS") + Space(1) + R("eidos"))
280                 Combo2.ComboItems(k).Image = 2
                    ' Combo2. vbRed
                Else

                    ' Combo2.AddItem Format(K, "##") + " " + R("TITLOS") + Space(1) + R("eidos")

290                 Set o = Combo2.ComboItems.Add(k + 1, "", Format(k, "##") + "---" + " " + R("TITLOS") + Space(1) + R("eidos"))
300                 Combo2.ComboItems(k + 1).Image = 3

                End If

310             k = k + 1
                'mEidos_Par(K) = R("eidos")
            End If

320         R.MoveNext
        Loop

330     Combo2.Text = Combo2.ComboItems(F_1ST_CHOICE)
        ' Combo2.ListIndex = F_1ST_CHOICE
        ' Combo2.Text = Combo2.List(F_1ST_CHOICE)

        'DTPicker1.value = Now
        'DTPicker2.value = Now
        'If CallFromOther.Caption = "1" Then
        '    Command1_Click
        'End If

        '<EhFooter>
        Exit Sub

Form_Activate_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Form_Activate " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Form_Activate " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'ενημέρωση αποθήκης αρνητικα
'Sub arn_enhm_apot()
'
'    Dim G, m, GA, MA, P
'    Dim R As New ADODB.Recordset, k, L As Integer
'    Dim r2
'    Dim e As New ADODB.Recordset
'
'    Dim gm_f_gfkod As String
'    Dim m2, gm_f_gfFPA
'    Dim m_m As Boolean
'    Dim f_k, f_p
'    Dim f_dat
'    Dim F_POS_APOU, F_AJIA_APOU
'    Dim meid As New ADODB.Recordset
'
'
'    f_k = 0    ' kodikoi
'    f_p = 3    ' posotites
'    'If gConnect = "Access" Then
'    'Set db = OpenDatabase(gDir, False, False)
'    'Else
'    'Set db = OpenDatabase(gDir, False, False, gConnect)
'    'End If
'    f_dat = EGGTIM.Recordset("Ημερομ")
'    G = "G" + Format(Month(f_dat), "00")
'    m = "M" + Format(Month(f_dat), "00")
'    GA = "GA" + Format(Month(f_dat), "00")
'    MA = "MA" + Format(Month(f_dat), "00")
'
'    For k = 1 To MSFlexGrid1.ROWS - 1    'Do While Not R.EOF
'        If Len(MSFlexGrid1.TextMatrix(k, f_k)) > 0 Then
'
'            meid.Open "SELECT *FROM EID WHERE KOD='" + MSFlexGrid1.TextMatrix(k, f_k) + "'", Gdb, adOpenKeyset, adLockOptimistic
'
'            If meid.RecordCount = 1 Then
'                P = Val(MSFlexGrid1.TextMatrix(k, f_p))    '("POSO")
'
'                ' a = P * Val(MSFlexGrid1.TextMatrix(K, f_t)) * (100 - Val(MSFlexGrid1.TextMatrix(K, f_e))) / 100
'
'                R.Open "select * from PARASTAT where EIDOS='" + Left(TIM.Recordset("atim"), 1) + "';", Gdb, adOpenForwardOnly, adLockReadOnly
'
'                'f_PEL = R("pel")
'                Dim f_matim
'                F_POS_APOU = R("POS_APOU")
'                F_AJIA_APOU = R("ajia_APOU")
'                R.Close
'
'
'
'                Dim mYP As Single, mG As Single, MM As Single
'                mYP = meid!pos
'
'
'                Select Case F_POS_APOU
'                Case "1"    ' EISAGOGES=1
'                    mYP = meid("POS") + P
'                    mG = meid(G) + P
'                Case "2"    ' -EISAGOGES
'                    mYP = meid("POS") - P
'                    '         mG = meid(G) - P
'                Case "3"    ' EXAGOGES=1
'                    mYP = meid("POS") - P
'                    MM = meid(m) + P
'                Case "4"    ' -EXAGOGES
'                    mYP = meid("POS") + P
'                    MM = meid(m) - P
'                End Select
'
'                Select Case F_AJIA_APOU
'                Case "1"    ' EISAGOGES=1
'                    ' meid(GA) = meid(GA) + a
'                Case "2"    ' -EISAGOGES
'                    '         meid(GA) = meid(GA) - a
'                Case "3"    ' EXAGOGES=1
'                    '         meid(MA) = meid(MA) + a
'                Case "4"    ' -EXAGOGES
'                    '  meid(MA) = meid(MA) - a
'                End Select
'                Gdb.Execute "UPDATE EID SET POS=" + Str(mYP) + " ," + G + "=" + Str(mG) + "," + m + "=" + Str(MM)
'            End If  ' If meid.RecordCount = 1 Then
'
'
'            meid.Close
'        End If    'len( msflexgrid1.TextMatrix(K, f_k) )>0
'    Next
'
'
'
'End Sub

Function del_Gefyroma(g_hme, g_atim, pist)

        '-------- pist=1 pistosei   pist=2 metrhtois
        '*********************** γεφυρωμα ****************************
        '<EhHeader>
        On Error GoTo del_Gefyroma_Err

        '</EhHeader>
        Dim xre_pis(20), xre_Pis_rec(20), xreosis(20), GetList(20)

        Dim DUM

        Dim jj, JJK, m_ait As String

        Dim kodiklog(9), TYP(9), XP(9), kau(9), a As String, mnq As String, eid_par As String

        Dim mgre As String

        Dim k    As Integer

        Dim atr_xre, atr_pis, lw, DIAFORETIKA

100     atr_xre = 0: atr_pis = 0

        Dim ok_gefyr

        Dim DB       As Database, dbL As Database

        Dim abat     As Recordset, temp As Recordset, r2 As Recordset, gefyres As Recordset

        Dim bbs      As String

        Dim anoigma  As Integer

        Dim kod_pel  As String

        Dim mbat_pel As String

        Dim abat2    As Recordset

        Dim lo       As Integer

        Dim JJ_ROBLEM

        Dim mmhkos, MHKOS(5) As Integer

        Dim MHNAS

        Dim sql, K2 As Integer, logar(5) As Integer

110     If gConnect = "Access" Then
120         Set DB = OpenDatabase(gDir, False, False)
130         Set dbL = OpenDatabase(gdirlog, False, False)
        Else

            On Error Resume Next

140         Set DB = OpenDatabase(gDir, False, False, gConnect)
150         Set dbL = OpenDatabase(gdirlog, False, False, gConnect)
        End If

        Dim par As Recordset

        'ΒΡΙΣΚΩ ΑΠΟ ΤΙΣ GEFYRES.DBF ΤΗΝ ΚΑΤΑΛΛΗΛΗ ΕΓΓΡΑΦΗ
160     Set par = DB.OpenRecordset("SELECT *FROM PARASTAT WHERE EIDOS='" + Left(g_atim, 1) + "';")
170     Set gefyres = DB.OpenRecordset("GEFYRES")

180     If pist = 1 Then
190         gefyres.Move par("GEF_P") - 1  ' RECORD ΓΕΦΥΡΏΣΕΩΝ ΕΠΙ ΠΙΣΤΩΣΕΙ
        Else
200         gefyres.Move par("GEF_m") - 1  ' RECORD ΓΕΦΥΡΏΣΕΩΝ metritois
        End If

        ' Set gTim = db.OpenRecordset("select *from tim where atim='" + g_atim + "' and hme>=#" + Format(g_hme, "mm/dd/yyyy") + "#;")
210     ok_gefyr = 1

        '
220     If g_hme <= ghme_orist Then

230         MsgBox " ΔΕΝ ΕΠΙΤΡΕΠΟΝΤΑΙ ΚΙΝΗΣΕΙΣ  ΣΕ ΟΡΙΣΤΙΚΟΠΟΙΗΜΕΝΕΣ ΕΓΓΡΑΦΕΣ"

            Exit Function

        End If

240     If InStr("oOοΟ", gefyres("gefyrono")) Then
250         ok_gefyr = 0

            Exit Function

        End If

        Dim EIDPAR

260     eid_par = "  "

270     If gefyres("mEID_PAR") <> "  " Then
280         eid_par = Left(gefyres("mEID_PAR"), 5) + "-" + g_atim    ' Τ.πωλ-Τ00012Γ
        End If

290     If gefyres("mmGRE") <> " " Then
300         mgre = gefyres("mmGRE")    ' Σ
        End If

310     m_ait = gefyres("parathrhs")    'π.χ.  ΠΩΛΗΣΕΙΣ
        '
        ''
        ''
        '''
        '''
        '''
        '  For K = 1 To 9
        '      a = LTrim(Str(K))
        '      kodiklog(K) = Gefyres("kod" + a)  ' 3000****
        '      typ(K) = Gefyres("typos" + a)
        '      xp(K) = Gefyres("xpi" + a)
        '  Next
        '
        '  m_ait = Gefyres("parathrhs") 'π.χ.  ΠΩΛΗΣΕΙΣ

        '*** ψαχνω για διπλοεγγραφές

320     gvar = "Select *from temp where eidhme='1' and hme=#"
330     gvar = gvar + Format(g_hme, "mm/dd/yyyy") + "#  and mid(apa,7,7)='"
340     gvar = gvar + mID(eid_par, 7, 7) + "'"

350     Set temp = dbL.OpenRecordset(gvar)

360     If temp.RecordCount > 0 Then
            ' OK
        End If

370     temp.MoveFirst

        ' ενημερώνω τα αρχείο λογαριασμών abat
380     Do While Not temp.EOF    ' and  eidhme+dtos(hme)+left(gre,1)+apa =  '1' + dtos(g_hme) + mgre  + EID_PAR
390         mmhkos = Len(Trim(temp("kod")))

400         If temp("eidHME") = "3" Then    ' απογραφή
410             MHNAS = "00"
            Else
420             MHNAS = find_mhnas(temp("hme"))
            End If

430         sql = "select *from abat WHERE "

440         If mmhkos = gDIG_4 Then sql = sql + "(bat='4' and left(kod," + str(gDIG_4) + ")='" + temp("kod") + "') or"
450         If mmhkos >= gDIG_3 Then sql = sql + "(bat='3' and left(kod," + str(gDIG_3) + ")='" + Left(temp("kod"), gDIG_3) + "') or "
460         If mmhkos >= 4 Then sql = sql + "(bat='2' and left(kod,4)='" + Left(temp("kod"), 4) + "') or "
470         If mmhkos >= 2 Then sql = sql + "(bat='1' and left(kod,2)='" + Left(temp("kod"), 2) + "') "

            'συγκεντρώνω τους λογαριασμους που με ενδιαφέρουν
480         Set abat2 = dbL.OpenRecordset(sql)
            ' εαν λειπουν τους ανοιγω
490         MHKOS(1) = 2: MHKOS(2) = 4: MHKOS(3) = gDIG_3: MHKOS(4) = gDIG_4

500         If mmhkos = 2 Then lo = 1
510         If mmhkos = 4 Then lo = 2
520         If mmhkos = gDIG_3 Then lo = 3
530         If mmhkos = gDIG_4 Then lo = 4
540         If abat2.RecordCount < lo Then

550             For K2 = 1 To lo: logar(K2) = 0: Next

                On Error Resume Next

560             abat2.MoveFirst

570             Do While Not abat2.EOF
580                 logar(Val(abat2("bat"))) = 1
590                 abat2.MoveNext
                Loop

600             For K2 = 1 To lo

610                 If logar(K2) = 0 Then
620                     abat2.AddNew
630                     abat2("bat") = LTrim(str(K2))
640                     abat2("kod") = Left(temp("kod"), MHKOS(K2))
650                     abat2.Update
                    End If

                Next

            End If

660         abat2.MoveFirst

670         Do While Not abat2.EOF
680             abat2.Edit

690             If temp("xpi") = "Π" Then
700                 If IsNull(abat2("ppros" + MHNAS)) Then abat2("ppros" + MHNAS) = temp("xre") Else abat2("ppros" + MHNAS) = abat2("ppros" + MHNAS) + temp("xre")
                Else

710                 If IsNull(abat2("ppros" + MHNAS)) Then abat2("xpros" + MHNAS) = temp("xre") Else abat2("xpros" + MHNAS) = abat2("xpros" + MHNAS) + temp("xre")
                End If

720             abat2.Update
730             abat2.MoveNext
            Loop

740         abat2.Close
750         temp.MoveNext
        Loop

760     temp.Close
770     abat.Close

        ' ΔΙΑΓΡΑΦΩ ΤΙς ΕΓΓΡΑΦΕΣ
780     gvar = "DELETE *from temp where eidhme='1' and hme=#"
790     gvar = gvar + Format(g_hme, "mm/dd/yyyy") + "# and left(apa,13)='"
800     gvar = gvar + eid_par + "'"
810     dbL.Execute gvar
        'DELETE *from temp where eidhme='1' and hme=#12/27/2004# and left(apa,13)='.ΑΠ.Λ-L00145Α'
820     dbL.Close

        '
        '
        '
        '
        '
        '<EhFooter>
        Exit Function

del_Gefyroma_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.del_Gefyroma " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.del_Gefyroma " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Sub arn_enhm_pel()

        '<EhHeader>
        On Error GoTo arn_enhm_pel_Err

        '</EhHeader>
        On Error GoTo arn_enhm_pel_Err

        '----------------------------------------------------------------------

        Dim x As String

        'Dim DB As New ADODB.Connection
        'Dim E As New ADODB.Recordset
        'Dim R As New ADODB.Recordset, K
        'DB.Open gDir

        Dim e As New ADODB.Recordset, R As New ADODB.Recordset, k

        Dim PEL As New ADODB.Recordset

        Dim f_pel

        'If gConnect = "Access" Then
        '  Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

        '

100     R.Open "select * from PARASTAT where EIDOS='" + Left(TIM.Recordset("atim"), 1) + "';", Gdb, adOpenForwardOnly, adLockReadOnly

110     f_pel = R("pel")

        Dim f_matim

120     f_matim = TIM.Recordset("atim")

130     PEL.Open "SELECT *FROM PEL WHERE EIDOS='" + TIM.Recordset("EIDOS") + "' AND KOD='" + TIM.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic

        'Set R = db.OpenRecordset("egg") ' , DB, adOpenKeyset, adLockOptimistic

        Dim f_suma    '= Val(SynoloG.Caption)

140     f_suma = TIM.Recordset("ajia")

        'PEL.Edit

        '**********************************************************************

        Dim TELYP As Single

150     If TIM.Recordset("TRP") = "ΜΕ" Then
        Else

160         If PEL("eidos") = "e" Then
170             If f_pel = "1" Then
180                 TELYP = PEL("typ") - f_suma
                End If

190             If f_pel = "2" Then
200                 TELYP = PEL("typ") + f_suma
                End If
            End If

210         If PEL("eidos") = "r" Then
220             If f_pel = "1" Then
230                 TELYP = PEL("typ") - f_suma
                End If

240             If f_pel = "2" Then
250                 TELYP = PEL("typ") + f_suma
                End If
            End If

260         Gdb.Execute "UPDATE PEL SET TYP=" + str(TELYP) + " WHERE  EIDOS='" + TIM.Recordset("EIDOS") + "' AND KOD='" + TIM.Recordset("KOD") + "'"
        End If

        Dim sql

270     sql = "delete from EGG where EIDOS='" + TIM.Recordset("EIDOS") + "' AND KOD='" + TIM.Recordset("KOD") + "' and ATIM='" + TIM.Recordset("atim") + "'"
280     Gdb.Execute sql

290     R.Close

        '<EhFooter>
        Exit Sub

arn_enhm_pel_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.arn_enhm_pel " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.arn_enhm_pel " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "PAR2"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command4_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl

110     Me.Picture = LoadPicture(gPicture)
120     Me.KeyPreview = True
130     DTPicker1.Value = Now
140     DTPicker2.Value = Now











 
150         F_ET_PER_X = Val(FINDPARAMETROI(1, "PAR2", "F_ET_PER_X", "0", "12. 1.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΠΕΡΙΓΡΑΦΗΣ"))
160         F_ET_PER_Y = Val(FINDPARAMETROI(1, "PAR2", "F_ET_PER_Y", "0", "12. 2.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y ΠΕΡΙΓΡΑΦΗΣ"))
      
           

170         F_ET_KOD_X = Val(FINDPARAMETROI(1, "PAR2", "F_ET_KOD_X", "0", "12. 3.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ ΚΩΔΙΚΟΥ"))
180         F_ET_KOD_Y = Val(FINDPARAMETROI(1, "PAR2", "F_ET_KOD_Y", "200", "12. 4.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Υ ΚΩΔΙΚΟΥ"))


190         F_ET_BAR_X = Val(FINDPARAMETROI(1, "PAR2", "F_ET_BAR_X", "0", "12. 5.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ BARCODE"))
200         F_ET_BAR_Y = Val(FINDPARAMETROI(1, "PAR2", "F_ET_BAR_Y", "600", "12. 6.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y BARCODE"))
            F_ET_BAR_H = Val(FINDPARAMETROI(1, "PAR2", "F_ET_BAR_H", "500", "12. 7.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  ΥΨΟΣ BARCODE"))


210         F_ET_TIM_X = Val(FINDPARAMETROI(1, "PAR2", "F_ET_TIM_X", "200", "12. 8.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Χ TIMΗΣ"))
220         F_ET_TIM_Y = Val(FINDPARAMETROI(1, "PAR2", "F_ET_TIM_Y", "800", "12. 9.ΕΤΙΚΕΤΤΑ ΡΑΦΙ  Y TIMΗΣ"))






























        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Label4_Click()
'
'End Sub

Private Sub MSFlexGrid1_Click()

    '<EhHeader>
    On Error GoTo MSFlexGrid1_Click_Err

    '</EhHeader>
    Dim F

    '  f = MSFlexGrid1.Text
    '
    '
    '
    '' kartella
    '  apot21.Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)
    '
    '  apot21.Text2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)
    '
    '  apot21.Label1.Caption = "1" ' Left(APOTH1.Text, 1)
    '
    '  apot21.Label2.Caption = "0"
    '
    '
    '
    '  apot21.SHOW
    '

    '<EhFooter>
    Exit Sub

MSFlexGrid1_Click_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.par2.MSFlexGrid1_Click " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.MSFlexGrid1_Click " & "at line " & Erl

    Resume Next

    '</EhFooter>

End Sub
'Sub SOKRATIS_etik()
'
'End Sub

Private Sub MSFlexGrid1_DblClick()

        '<EhHeader>
        On Error GoTo MSFlexGrid1_DblClick_Err

        '</EhHeader>
        Dim F

100     F = MSFlexGrid1.Text

        ' kartella
110     apot21.Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)

120     apot21.Text2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)

130     apot21.Label1.Caption = "1"    ' Left(APOTH1.Text, 1)

140     apot21.Label2.Caption = "0"

150     apot21.SHOW

        '<EhFooter>
        Exit Sub

MSFlexGrid1_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.MSFlexGrid1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.MSFlexGrid1_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                Y As Single)

        '<EhHeader>
        On Error GoTo MSFlexGrid1_MouseUp_Err

        '</EhHeader>
100     If Button = 2 Then

110         MSFlexGrid1.Col = 0
            ' APOT2.kod.Text = MSFlexGrid1.Text   ' "10"

120         APOT2.kod.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)

130         APOT2.SHOW
140         APOT2.kod.SetFocus
150         SendKeys "{ENTER}"
        End If

        '<EhFooter>
        Exit Sub

MSFlexGrid1_MouseUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par2.MSFlexGrid1_MouseUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par2.MSFlexGrid1_MouseUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

