VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form bohu4 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   7188
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   14892
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7188
   ScaleWidth      =   14892
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6480
      TabIndex        =   27
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "eidos"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   15
      Top             =   0
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "pel"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   14
      Top             =   320
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "tameiakh"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   13
      Top             =   640
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "titlos"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   12
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "xreopis"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   2040
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
      ScaleWidth      =   14892
      TabIndex        =   5
      Top             =   6285
      Width           =   14895
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
      Left            =   2040
      TabIndex        =   4
      Top             =   1620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TAYTPEL"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   3
      Top             =   2820
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TAYTXPI"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   2
      Top             =   3180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EPIT"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   1
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "KODLOG"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   0
      Top             =   3540
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2040
      Top             =   3900
      Width           =   3360
      _ExtentX        =   5927
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
      Caption         =   "Πινακας κινήσεων"
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
      Left            =   120
      TabIndex        =   18
      Top             =   2340
      Width           =   5295
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Επιτ=1/Μετρ=0/2=Συμψ"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   1980
      Width           =   1815
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



Private Sub Command1_Click()

'Set lag2 = New InteropForm1
'
'lag2.SHOW
'
'
'Set lag = New HelloWorldForm
'
'lag.Initialize "asasa"
'
'
'lag.SHOW vbModeless

   

End Sub



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

120     DATA1.ConnectionString = gConnect
130     DATA1.RecordSource = "select *from XREOPIS ORDER BY EIDOS"
140     DATA1.Refresh
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

100     DATA1.Recordset.AddNew

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

100     With DATA1.Recordset
110         .delete
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

100     DATA1.Refresh

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

100     DATA1.Recordset.UpdateBatch

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

