VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form DataEntry2 
   BackColor       =   &H00FF0000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   7095
   ClientLeft      =   3090
   ClientTop       =   5355
   ClientWidth     =   8685
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "dataentry.frx":0000
      Height          =   5550
      Left            =   135
      TabIndex        =   8
      ToolTipText     =   "Mε διπλό κλικ διορθώνω την εγγραφή"
      Top             =   570
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   9790
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
      ColumnCount     =   10
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
         MarqueeStyle    =   4
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
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox mactext 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9930
      TabIndex        =   7
      Top             =   45
      Visible         =   0   'False
      Width           =   720
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   165
      Top             =   5640
      Visible         =   0   'False
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   635
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
   Begin VB.CommandButton BUT 
      Caption         =   "Εξοδος"
      Height          =   260
      Index           =   3
      Left            =   6405
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   2145
   End
   Begin VB.CommandButton BUT 
      Caption         =   "Διαγραφή"
      Enabled         =   0   'False
      Height          =   260
      Index           =   2
      Left            =   4275
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   2145
   End
   Begin VB.CommandButton BUT 
      Caption         =   "Αποθήκευση"
      Enabled         =   0   'False
      Height          =   260
      Index           =   1
      Left            =   2130
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   2145
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid ParamGrid 
      Height          =   1905
      Left            =   6420
      TabIndex        =   3
      Top             =   6375
      Visible         =   0   'False
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   3360
      _Version        =   393216
      Rows            =   20
      Cols            =   10
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.TextBox Textboxes 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Index           =   0
      Left            =   2460
      TabIndex        =   2
      Top             =   270
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CommandButton BUT 
      Caption         =   "Νέα εγγραφή"
      Height          =   260
      Index           =   0
      Left            =   -15
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   2145
   End
   Begin VB.Label Labels 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   270
      Visible         =   0   'False
      Width           =   2235
   End
End
Attribute VB_Name = "DataEntry2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const VK_TAB = &H9

Dim i As Integer

Dim Captions(50) As String

Dim mfields(50) As String

Dim Sizes(50) As Integer

Dim sql As String

Dim Sql2 As String

Dim delSql As String

Dim UpdSQL As String

Dim F_NEW As Integer

Dim defSQL As String    'default string που προσθέτω

Dim N As Integer    ' αριθμός πεδίων

Dim fSCR As Object    ' SCRIPT CONTROL

Dim f_Typos1Field As Long  ' τι τυπος είναι το 1ο πεδίο

Dim fAddNewScript

Dim F_FIELD_EPIPLEON As String, F_TIMH
Attribute F_TIMH.VB_VarUserMemId = 1073938446

Dim F_2FIELD_EPIPLEON As String, F_2TIMH
Attribute F_2FIELD_EPIPLEON.VB_VarUserMemId = 1073938448
Attribute F_2TIMH.VB_VarUserMemId = 1073938448

Dim flagProthFora
Attribute flagProthFora.VB_VarUserMemId = 1073938450

Sub PaintForm()

        '<EhHeader>
        On Error GoTo PaintForm_Err

        '</EhHeader>

        Dim i As Integer

        ' n = Val(ParamGrid.TextMatrix(0, 0)) 'ο αριθμός των σειρών

100     sql = ParamGrid.TextMatrix(0, 1)
110     Sql2 = ParamGrid.TextMatrix(0, 2)    ' string adodc1
120     delSql = ParamGrid.TextMatrix(0, 3)    ' string adodc1

130     UpdSQL = ParamGrid.TextMatrix(0, 4)    ' string adodc1

140     fAddNewScript = ParamGrid.TextMatrix(0, 5)    ' string meta to addnew button

150     F_FIELD_EPIPLEON = ParamGrid.TextMatrix(0, 6)    'EPIPLEON FIELD
160     F_TIMH = ParamGrid.TextMatrix(0, 7)  ' TIMH TOY EPIPLEON FIELD

170     F_2FIELD_EPIPLEON = ParamGrid.TextMatrix(0, 8)    'EPIPLEON FIELD
180     F_2TIMH = ParamGrid.TextMatrix(0, 9)  ' TIMH TOY EPIPLEON FIELD

        'f.ParamGrid.TextMatrix(0, 8) = "TIMH" '2 EPIPLEON FIELD
        'f.ParamGrid.TextMatrix(0, 9) = Left(Combo4.Text, 2)  ' TIMH TOY 2 EPIPLEON FIELD

        Dim R As New ADODB.Recordset

190     R.Open Sql2, Gdb, adOpenKeyset, adLockOptimistic

200     N = R.FIELDS.Count

210     For i = 1 To N
            ' Captions(i + 1) = r.Fields(i).Name 'ονομα ετικέτας
220         Captions(i) = ParamGrid.TextMatrix(i, 1)  'ονομα ετικέτας

230         If R.FIELDS(i - 1).DefinedSize > 60 Then
                 Sizes(i) = R.FIELDS(i - 1).DefinedSize
240            '  Sizes(i) = 30    '  R.FIELDS(I - 1).ActualSize 'definedsize  ' φάρδος textbox
            Else
250             Sizes(i) = R.FIELDS(i - 1).DefinedSize  ' ActualSize 'definedsize  ' φάρδος textbox
            End If

260         mfields(i) = R.FIELDS(i - 1).Name    ' field textbox

        Next

        'το χρειαζομαι για να συντάξω το sql αναζητησης
270     f_Typos1Field = R.FIELDS(0).Type  '3 integer  135=date  202=string

280     For i = 1 To N
290         Load Labels(i)
300         Load Textboxes(i)
310         Labels(i).Top = Labels(i - 1).Top + 1.2 * Labels(0).Height
320         Labels(i).Left = Labels(0).Left
330         Textboxes(i).Top = Textboxes(i - 1).Top + 1.2 * Textboxes(i).Height
340         Textboxes(i).Left = Textboxes(0).Left
350         Textboxes(i).Width = (Sizes(i) + 1) * TextWidth("A")
360         Textboxes(i).MaxLength = Sizes(i)
370         Labels(i).Caption = Captions(i)
            ' Labels(I).Visible = True
            ' Textboxes(I).Visible = True
        Next

        ' ParamGrid.HEIGHT =7905  MAXIMUM

        '<EhFooter>
        Exit Sub

PaintForm_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.PaintForm " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.PaintForm " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub paintgrid()

        '<EhHeader>
        On Error GoTo paintgrid_Err

        '</EhHeader>

        Dim i As Integer

        Dim l As Long

100     l = 0

110     Adodc1.ConnectionString = gConnect
120     Adodc1.RecordSource = Sql2
130     Adodc1.Refresh

140     For i = 1 To Adodc1.Recordset.FIELDS.Count

            'grid.ColWidth(i - 1) = (Sizes(i) + 1) * TextWidth("A")    'flexgrid
            'grid.TextMatrix(0, i - 1) = ParamGrid.TextMatrix(i, 1)    'flexgrid
            ' L = L + grid.ColWidth(i - 1)                             'flexgrid
            ' ΓΙΑ ΝΑ ΧΩΡΑΕΙ Ο ΤΙΤΛΟΣ
150         If Sizes(i) < Len(Captions(i)) Then
160             grid.columns(i - 1).Width = Len(Captions(i)) * TextWidth("A")
            Else
170             grid.columns(i - 1).Width = (Sizes(i) + 1) * TextWidth("A")
            End If

180         grid.columns(i - 1).Caption = ParamGrid.TextMatrix(i, 1)
190         l = l + grid.columns(i - 1).Width
        Next

200     If l + 5 * TextWidth("A") > Me.Width Then
210         grid.Width = Me.Width - 3 * TextWidth("A")
        Else
220         grid.Width = l + 5 * TextWidth("A")
        End If

        '<EhFooter>
        Exit Sub

paintgrid_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.paintgrid " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.paintgrid " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub BUT_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo BUT_Click_Err

        '</EhHeader>
        Dim ANS    As Integer

        Dim R      As New ADODB.Recordset

        Dim i      As Integer

        Dim CurRow As Long

        Dim mkod   As String

        '-----------------------------------------------------------------------------------------------------------------------
100     If Index = 0 Then    ' προσθετω εγγραφη
110         F_NEW = 1

120         For i = 1 To N
130             Textboxes(i).Text = ""
            Next

140         BUT(0).Enabled = False
150         BUT(1).Enabled = True

160         If Len(fAddNewScript) > 0 Then
170             trexe fAddNewScript
            End If

180         For i = 1 To N

                ' Textboxes(I).Text = "              "
190             Labels(i).Visible = True
200             Textboxes(i).Visible = True
            Next

210         grid.Top = Labels(N).Top + 1.2 * Labels(N).Height
220         Textboxes(1).SetFocus
        End If

        '----------------------------------------------------------------------------------------------------------------
230     If Index = 1 Then    'ΑΠΟΘΗΚΕΥΣΗ
240         mkod = Textboxes(1).Text  ' grid.Columns(0).Text

250         If f_Typos1Field = adInteger Or f_Typos1Field = adSingle Or f_Typos1Field = adDouble Then   ' integer
260             R.Open sql + " and " + mfields(1) + "= " + mkod + " ", Gdb, adOpenDynamic, adLockOptimistic
            Else
270             R.Open sql + " and " + mfields(1) + "='" + mkod + "'", Gdb, adOpenKeyset, adLockOptimistic
            End If

280         If F_NEW = 1 Then
290             R.AddNew
300             F_NEW = 0
            End If

            Dim m_ID As String

            m_ID = str(R!ID)

            On Error GoTo LatosTypoi

310         For i = 1 To N
320             R(mfields(i)) = Textboxes(i).Text
330             Labels(i).Visible = False
340             Textboxes(i).Visible = False
            Next

            'ΠΟΙΟ ΠΕΔΙΟ ΕΝΗΜΕΡΩΝΩ ΧΩΡΙΣ ΝΑ ΒΛΕΠΕΙ Ο ΧΡΗΣΤΗΣ
350         If Len(F_FIELD_EPIPLEON) > 0 Then
360             R(F_FIELD_EPIPLEON) = F_TIMH
            End If

            'ΠΟΙΟ ΠΕΔΙΟ ΕΝΗΜΕΡΩΝΩ ΧΩΡΙΣ ΝΑ ΒΛΕΠΕΙ Ο ΧΡΗΣΤΗΣ
370         If Len(F_2FIELD_EPIPLEON) > 0 Then
380             R(F_2FIELD_EPIPLEON) = F_2TIMH
            End If

            On Error GoTo NOSAVE

390         R.Update

            Dim x As String

400         If Len(UpdSQL) > 0 Then
                'X = mac(UpdSQL)
                'gdb.Execute UpdSQL
410             trexe UpdSQL
            End If

420         grid.Top = BUT(0).Top + BUT(0).Height + TextHeight("A")
430         grid.Visible = True
440         Adodc1.Refresh
450         paintgrid
460         BUT(0).Enabled = True    ' add
470         BUT(1).Enabled = False    'save
480         BUT(2).Enabled = False    'DELETE
490         BUT(0).SetFocus

            Exit Sub

        End If

        '-------------------------------------------------------------------------------------------------------------------
500     If Index = 2 Then    'διαγραφη
510         ANS = MsgBox("Είστε σίγουροι για την διαγραφή;", vbYesNo)

520         If ANS = vbYes Then
530             CurRow = grid.row    '+ 1
                ' MKOD = grid.TextMatrix(CURROW, 0) ' kodikos   'flexgrid
540             mkod = grid.columns(0).Text
550             Gdb.Execute delSql + "  " + mfields(1) + "='" + mkod + "'"
560             grid.Visible = True
570             Adodc1.Refresh
580             BUT(0).Enabled = True    ' add
590             BUT(1).Enabled = False    'save
            End If

600         For i = 1 To N
610             Labels(i).Visible = False
620             Textboxes(i).Visible = False
            Next

        End If

630     If Index = 3 Then    ' ΕΞΟΔΟσ
640         Unload Me

        End If

        Exit Sub

LatosTypoi:
650     MsgBox "λάθος στο πεδίο " + Labels(i)
660     Textboxes(i).SetFocus

NOSAVE:

        '<EhFooter>
        Exit Sub

BUT_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.BUT_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.BUT_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Command1_Click()
'    trexe "AA"
'End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     flagProthFora = 1
110     F_NEW = 0

120     Me.KeyPreview = True
        '  Me.Picture = LoadPicture(gPicture)

        '    'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        '    'SkinFramework1.ApplyWindow Me.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Paint()

        '<EhHeader>
        On Error GoTo Form_Paint_Err

        '</EhHeader>
        Dim k As Integer

100     If flagProthFora = 1 Then

110         PaintForm
120         paintgrid

130         Set fSCR = CreateObject("MSScriptControl.ScriptControl")
140         fSCR.language = "vbscript"
150         fSCR.AddObject "mactext", MACtEXT

160         For k = 1 To N
170             fSCR.AddObject "t" + Format(k, "##"), Textboxes(k)
            Next

180         flagProthFora = 0
        End If

        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.Form_Paint " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.Form_Paint " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub grid_Click()
    '   grid_DblClick
End Sub

Private Sub grid_DblClick()

        '<EhHeader>
        On Error GoTo grid_DblClick_Err

        '</EhHeader>
        Dim CurRow As Long

        Dim mkod   As String

100     CurRow = grid.row    '+ 1
        'MKOD = grid.TextMatrix(CURROW, 0) ' kodikos  'flexgrid

        'ΑΝ ΔΕΝ ΕΧΕΙ ΕΓΓΡΑΦΕΣ ΝΑ ΜΗΝ ΠΡΟΧΩΡΑΕΙ
110     If Adodc1.Recordset.EOF Then Exit Sub

120     mkod = grid.columns(0).Text

        Dim R As New ADODB.Recordset

        Dim i As Integer

        ' "SELECT KOD,EPO,AYP FROM PEL0" 'ParamGrid

        ' r.Open Sql + " and " + mFields(1) + "='" + MKOD + "'", gdb, adOpenKeyset, adLockOptimistic

130     If f_Typos1Field = 3 Then    ' integer
140         If mkod = "" Then
150             R.Open sql + " and " + mfields(1) + " is null ", Gdb, adOpenKeyset, adLockOptimistic
            Else
160             R.Open sql + " and " + mfields(1) + "= " + mkod + " ", Gdb, adOpenKeyset, adLockOptimistic
            End If

        Else

170         If mkod = "" Then
180             R.Open sql + " and " + mfields(1) + " is null ", Gdb, adOpenKeyset, adLockOptimistic
            Else
190             R.Open sql + " and " + mfields(1) + "='" + mkod + "'", Gdb, adOpenDynamic, adLockOptimistic
            End If

            ' R.Open sql + " and " + mfields(1) + "='" + mkod + "'", Gdb, adOpenKeyset, adLockOptimistic
        End If

200     For i = 1 To N

210         If IsNull(R(mfields(i))) Then
220             Textboxes(i).Text = ""
            Else
230             Textboxes(i).Text = R(mfields(i))
            End If

240         Labels(i).Visible = True
250         Textboxes(i).Visible = True
        Next

260     grid.Visible = False

270     BUT(0).Enabled = False    ' add
280     BUT(1).Enabled = True    'save
290     BUT(2).Enabled = True    'DELETE

300     Textboxes(1).SetFocus

        '<EhFooter>
        Exit Sub

grid_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.grid_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.grid_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub grid_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo grid_KeyUp_Err

        '</EhHeader>
100     If KeyCode = 13 Then
110         grid_DblClick
        End If

        '<EhFooter>
        Exit Sub

grid_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.grid_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.grid_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' grid_DblClick
End Sub

Private Sub Textboxes_GotFocus(Index As Integer)

        '<EhHeader>
        On Error GoTo Textboxes_GotFocus_Err

        '</EhHeader>

100     Textboxes(Index).SelStart = 0
110     Textboxes(Index).SelLength = Len(Textboxes(Index).Text)

120     Textboxes(Index).BackColor = vbYellow

        '<EhFooter>
        Exit Sub

Textboxes_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.Textboxes_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.Textboxes_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Textboxes_KeyPress(Index As Integer, KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Textboxes_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

Textboxes_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.Textboxes_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.Textboxes_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Textboxes_LostFocus(Index As Integer)

        '<EhHeader>
        On Error GoTo Textboxes_LostFocus_Err

        '</EhHeader>
        Dim x

100     x = ParamGrid.TextMatrix(Index, 2)

110     If Len(x) > 0 Then
120         trexe x
        End If

130     Textboxes(Index).BackColor = Textboxes(0).BackColor

        '<EhFooter>
        Exit Sub

Textboxes_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.Textboxes_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.Textboxes_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Function mac(ByVal S As String) As String
''δινω ενα στρινγκ και μου επιστρέφει μία τιμή
'    Dim X
'    On Error GoTo 0
'    mactext.Text = "   "
'    On Error GoTo err2
'    X = "MACtEXT.Text = " + S
'    fSCR.ExecuteStatement X
'
'    'If F928 = 1 Then
'    mac = mactext.Text
'    'Else
'    '    mac = to437(mactext.Text)
'    'End If
'
'    Exit Function
'
'err2:
'    Me.Caption = X
'    mac = "  "
'    Resume Next
'End Function

Private Sub trexe(ByVal Scode As String)
        ' gconnect = "DSN=MERC2"
        'scode = "Sub Main()" & vbCrLf & _
         "Msgbox 15 " & vbCrLf & _
         " DIM mCon  " & vbCrLf & _
         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
         " mCon.Open """ + gconnect + """ " & vbCrLf & _
         " mCon.execute ""UPDATE ARITMISI2 SET ARITMISI=457  WHERE ID= 7 "" " & vbCrLf & _
         "DIM R " & vbCrLf & _
         " Set R=CreateObject(""ADODB.Recordset"")" & vbCrLf & _
         " r.Open "" SELECT ARITMISI FROM ARITMISI2 WHERE ID=7 "",mCon " & vbCrLf & _
         "Msgbox r(0) " & vbCrLf & _
         "Set r=Nothing" & vbCrLf & _
         "Set mCon=Nothing" & vbCrLf & _
         "End Sub"

        '<EhHeader>
        On Error GoTo trexe_Err

        '</EhHeader>

100     Call fSCR.AddCode(Scode)
110     Call fSCR.Run("Main") ' AN UELV NA DHMIOYRHSV MIA PROCEDURE

        '<EhFooter>
        Exit Sub

trexe_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.DataEntry2.trexe " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.DataEntry2.trexe " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

