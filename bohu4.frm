VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form bohu4 
   BackColor       =   &H00FF0000&
   Caption         =   "Κινήσεις Πελ.&Προμ"
   ClientHeight    =   12405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12405
   ScaleWidth      =   19905
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7440
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox FORMA2 
      DataField       =   "FORMA2"
      DataSource      =   "data1"
      Height          =   285
      Left            =   2400
      TabIndex        =   32
      Top             =   5160
      Width           =   3375
   End
   Begin VB.CommandButton CMD2 
      Caption         =   "?"
      Height          =   300
      Left            =   6000
      TabIndex        =   31
      Top             =   5160
      Width           =   255
   End
   Begin VB.CommandButton cmd 
      Caption         =   "?"
      Height          =   300
      Left            =   6000
      TabIndex        =   30
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox FORMA1 
      DataField       =   "FORMA1"
      DataSource      =   "data1"
      Height          =   285
      Left            =   2400
      TabIndex        =   28
      Top             =   4800
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   11640
      TabIndex        =   27
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "eidos"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   15
      Top             =   0
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "pel"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   14
      Top             =   320
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "tameiakh"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   13
      Top             =   640
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "titlos"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   12
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "xreopis"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   2400
      TabIndex        =   11
      Top             =   1280
      Width           =   3375
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   19905
      TabIndex        =   5
      Top             =   11505
      Width           =   19905
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Νέα Εγγραφή"
         Height          =   300
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Ενημέρωση"
         Height          =   300
         Left            =   1320
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Διαγραφή"
         Height          =   300
         Left            =   2520
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Ανανέωση"
         Height          =   300
         Left            =   3720
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Ε&ξοδος"
         Height          =   300
         Left            =   4920
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ARITMISI"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   4
      Top             =   1620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TAYTPEL"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   2400
      TabIndex        =   3
      Top             =   2820
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TAYTXPI"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   2
      Top             =   3180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EPIT"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   2400
      TabIndex        =   1
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "KODLOG"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   2400
      TabIndex        =   0
      Top             =   3540
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   360
      Top             =   5880
      Width           =   7800
      _ExtentX        =   13758
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
      Caption         =   "Πινακας κινήσεων"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   6720
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
            Picture         =   "bohu4.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bohu4.frx":040A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLABELFORMAS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ΦΟΡΜΑ ΕΚΤΥΠΩΣΗΣ"
      ForeColor       =   &H8000000B&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   33
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblLABELFORMAS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ΦΟΡΜΑ ΕΚΤΥΠΩΣΗΣ"
      ForeColor       =   &H8000000B&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   29
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "A/A"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Πελάτης=e Προμ=r"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Ταμειακή"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Τίτλος"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Χρέωση/Πίστωση"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   1280
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Aρίθμηση"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   21
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Είδος (r/e) Κωδικός "
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   2820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Χρέωση/Πίστωση"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   3180
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ταυτόχρονη χρεωπίστωση του λογαριασμού"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   2520
      Width           =   5295
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Επιτ=1/Μετ=0/Συμψ=2 /KAPT=2 /ΚΤΘ=5"
      ForeColor       =   &H8000000E&
      Height          =   390
      Index           =   8
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   1755
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδ.Λογ.Ταμειακού Λογ"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   16
      Top             =   3540
      Width           =   1815
   End
End
Attribute VB_Name = "bohu4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

' Declare the .NET HelloWorldForm.
 
 'Dim WithEvents hello As HelloWorldForm
'Dim lag As HelloWorldForm

'Dim lag2 As georgeLibrary1.InteropForm1  'george2 '.


'Private Sub Command1_Click()

'Set lag2 = New InteropForm1

'lag2.SHOW


    ' Instantiate the form.
'    Set hello = New HelloWorldForm
'    hello.Initialize "Hello from VB6"
'Set lag2 = New HelloWorldForm

'lag2.Initialize "asasa"


'lag2.SHOW vbModeless

    ' Show the form.
    'hello.SHOW vbModeless

    ' Other standard Form methods and properties are available...
    'hello.Move 0, 0, 800, 800
    'hello.Caption = "Other Caption"

    ' Other custom methods and properties are available, as well...
    'hello.HelloText = "Other Hello Text"
    'hello.ReverseBackgroundColors


'End Sub



' Handle the custom application event.
Private Sub eventMessenger_ApplicationEventRaised(ByVal eventName As String, ByVal eventArgs As Variant)
    Select Case eventName
        Case "CRITICAL_ERROR"
            ' Show the error code which is passed in the eventArgs for this type of event.
            MsgBox "Error Code: " & eventArgs, vbExclamation, "Critical Error"
            ' Shut down the application.
            Unload Me
    End Select
End Sub



Private Sub cmd_Click()

   cd1.ShowOpen
   FORMA1.Text = cd1.FILENAME
   
End Sub

Private Sub CMD2_Click()

   cd1.ShowOpen
   FORMA2.Text = cd1.FILENAME
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then cmdClose_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu4.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu4.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     '  mForm_Load Me, fh, fw, ft, fl
110     Me.Picture = LoadPicture(gPicture)

120     data1.ConnectionString = gConnect
130     data1.RecordSource = "select *from XREOPIS ORDER BY EIDOS"
140     data1.Refresh
150     Me.KeyPreview = True
        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu4.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu4.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    ResizeFormFor Me
    
    '</EhHeader>

    'mForm_Resize Me,8, 7035, 5050, fh, fw, ft, fl
End Sub

Private Sub Form_Unload(Cancel As Integer)

        '<EhHeader>
        On Error GoTo Form_Unload_Err

        '</EhHeader>

100     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu4.Form_Unload " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu4.Form_Unload " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdAdd_Click()

        '<EhHeader>
        On Error GoTo cmdAdd_Click_Err

        '</EhHeader>
        On Error GoTo AddErr

100     data1.Recordset.AddNew

        Exit Sub

AddErr:
110     MsgBox Err.Description

        '<EhFooter>
        Exit Sub

cmdAdd_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu4.cmdAdd_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu4.cmdAdd_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdDelete_Click()

        '<EhHeader>
        On Error GoTo cmdDelete_Click_Err

        '</EhHeader>
        On Error GoTo DeleteErr

100     With data1.Recordset
110         .Delete
120         .MoveNext

130         If .EOF Then .MoveLast
        End With

        Exit Sub

DeleteErr:
140     MsgBox Err.Description

        '<EhFooter>
        Exit Sub

cmdDelete_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu4.cmdDelete_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu4.cmdDelete_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdRefresh_Click()

        'This is only needed for multi user apps
        '<EhHeader>
        On Error GoTo cmdRefresh_Click_Err

        '</EhHeader>
        On Error GoTo RefreshErr

100     data1.Refresh

        Exit Sub

RefreshErr:
110     MsgBox Err.Description

        '<EhFooter>
        Exit Sub

cmdRefresh_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu4.cmdRefresh_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu4.cmdRefresh_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdUpdate_Click()

        '<EhHeader>
        On Error GoTo cmdUpdate_Click_Err

        '</EhHeader>
        On Error Resume Next

100     data1.Recordset.UpdateBatch

        '<EhFooter>
        Exit Sub

cmdUpdate_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu4.cmdUpdate_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu4.cmdUpdate_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdClose_Click()

        '<EhHeader>
        On Error GoTo cmdClose_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

cmdClose_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu4.cmdClose_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu4.cmdClose_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Image1_Click()
 'Shell "EXPLORER.EXE " + Me.Controls("FORMA1.Text, vbMaximizedFocus)
     Shell "EXPLORER.EXE " + Me.Controls("FORMA1").Text, vbMaximizedFocus
End Sub

Private Sub IMAGE2_Click()
  Shell "EXPLORER.EXE " + Me.Controls("FORMA2").Text, vbMaximizedFocus
End Sub
