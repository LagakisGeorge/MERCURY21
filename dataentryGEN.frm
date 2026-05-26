VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form DataEntry3 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   6360
   ClientLeft      =   3090
   ClientTop       =   5355
   ClientWidth     =   8580
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   240
      Left            =   4920
      TabIndex        =   9
      Top             =   300
      Width           =   3255
   End
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "dataentryGEN.frx":0000
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
Attribute VB_Name = "DataEntry3"
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

Dim n As Integer    ' αριθμός πεδίων

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

' to PARAMGRID ΦΕΡΝΕΙ ΟΛΕΣ ΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ
' Η ΣΕΙΡΑ 0 ΧΡΗΣΙΜΟΠΟΙΕΙΤΑΙ ΓΙΑ ΤΑ QUERIES
'Η ΣΤΗΛΗ 1 ( ΞΕΚΙΝΩΝΤΑΣ ΑΠΟ 1Η ΣΕΙΡΑ ) ΕΧΕΙ ΤΟΥ ΤΙΤΛΟΥΣ ΤΩΝ ΠΕΔΙΩΝ
'Η ΣΤΗΛΗ 2 ( ΞΕΚΙΝΩΝΤΑΣ ΑΠΟ 1Η ΣΕΙΡΑ ) ΕΧΕΙ ΤΟΝ ΤΥΠΟ ΤΩΝ ΠΕΔΙΩΝ (1=ΑΡΙΘΜΟΣ 2=STRING 3=DATE )





Sub PaintForm()
        '<EhHeader>
        On Error GoTo PaintForm_Err
        '</EhHeader>


        Dim i As Integer

        ' n = Val(ParamGrid.TextMatrix(0, 0)) 'ο αριθμός των σειρών

100     sql = ParamGrid.TextMatrix(0, 1)
102     Sql2 = ParamGrid.TextMatrix(0, 2)    ' string adodc1
104     delSql = ParamGrid.TextMatrix(0, 3)    ' string adodc1

106     UpdSQL = ParamGrid.TextMatrix(0, 4)    ' string adodc1

108     fAddNewScript = ParamGrid.TextMatrix(0, 5)    ' string meta to addnew button

110     F_FIELD_EPIPLEON = ParamGrid.TextMatrix(0, 6)    'EPIPLEON FIELD
112     F_TIMH = ParamGrid.TextMatrix(0, 7)  ' TIMH TOY EPIPLEON FIELD

114     F_2FIELD_EPIPLEON = ParamGrid.TextMatrix(0, 8)    'EPIPLEON FIELD
116     F_2TIMH = ParamGrid.TextMatrix(0, 9)  ' TIMH TOY EPIPLEON FIELD

        'f.ParamGrid.TextMatrix(0, 8) = "TIMH" '2 EPIPLEON FIELD
        'f.ParamGrid.TextMatrix(0, 9) = Left(Combo4.Text, 2)  ' TIMH TOY 2 EPIPLEON FIELD

        Dim R As New ADODB.Recordset

118     R.Open Sql2, Gdb, adOpenKeyset, adLockOptimistic

120     n = R.FIELDS.Count

122     For i = 1 To n
            ' Captions(i + 1) = r.Fields(i).Name 'ονομα ετικέτας
124         Captions(i) = ParamGrid.TextMatrix(i, 1)  'ονομα ετικέτας

126         If R.FIELDS(i - 1).DefinedSize > 60 Then
128             Sizes(i) = 30    '  R.FIELDS(I - 1).ActualSize 'definedsize  ' φάρδος textbox
            Else
130             Sizes(i) = R.FIELDS(i - 1).DefinedSize  ' ActualSize 'definedsize  ' φάρδος textbox
            End If

132         mfields(i) = R.FIELDS(i - 1).Name    ' field textbox

        Next

        'το χρειαζομαι για να συντάξω το sql αναζητησης
134     f_Typos1Field = R.FIELDS(0).Type  '3 integer  135=date  202=string

136     For i = 1 To n
138         Load Labels(i)
140         Load Textboxes(i)
142         Labels(i).Top = Labels(i - 1).Top + 1.2 * Labels(0).Height
144         Labels(i).Left = Labels(0).Left
146         Textboxes(i).Top = Textboxes(i - 1).Top + 1.2 * Textboxes(i).Height
148         Textboxes(i).Left = Textboxes(0).Left
150         Textboxes(i).Width = (Sizes(i) + 1) * TextWidth("A")
152         Textboxes(i).MaxLength = Sizes(i)
154         Labels(i).Caption = Captions(i)
            ' Labels(I).Visible = True
            ' Textboxes(I).Visible = True
        Next

        ' ParamGrid.HEIGHT =7905  MAXIMUM


        '<EhFooter>
        Exit Sub

PaintForm_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.PaintForm " & _
               "at line " & Erl
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

102     Adodc1.ConnectionString = gConnect
104     Adodc1.RecordSource = Sql2
106     Adodc1.Refresh

108     For i = 1 To Adodc1.Recordset.FIELDS.Count

            'grid.ColWidth(i - 1) = (Sizes(i) + 1) * TextWidth("A")    'flexgrid
            'grid.TextMatrix(0, i - 1) = ParamGrid.TextMatrix(i, 1)    'flexgrid
            ' L = L + grid.ColWidth(i - 1)                             'flexgrid
            ' ΓΙΑ ΝΑ ΧΩΡΑΕΙ Ο ΤΙΤΛΟΣ
110         If Sizes(i) < Len(Captions(i)) Then
112             grid.Columns(i - 1).Width = Len(Captions(i)) * TextWidth("A")
            Else
114             grid.Columns(i - 1).Width = (Sizes(i) + 1) * TextWidth("A")
            End If

116         grid.Columns(i - 1).Caption = ParamGrid.TextMatrix(i, 1)
118         l = l + grid.Columns(i - 1).Width
        Next

120     If l + 5 * TextWidth("A") > Me.Width Then
122         grid.Width = Me.Width - 3 * TextWidth("A")
        Else
124         grid.Width = l + 5 * TextWidth("A")
        End If


        '<EhFooter>
        Exit Sub

paintgrid_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.paintgrid " & _
               "at line " & Erl
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
102         F_NEW = 1

104         For i = 1 To n
106             Textboxes(i).Text = ""
            Next

108         BUT(0).Enabled = False
110         BUT(1).Enabled = True

112         If Len(fAddNewScript) > 0 Then
114             trexe fAddNewScript
            End If

116         For i = 1 To n

                ' Textboxes(I).Text = "              "
118             Labels(i).Visible = True
120             Textboxes(i).Visible = True
            Next

122         grid.Top = Labels(n).Top + 1.2 * Labels(n).Height
124         Textboxes(1).Enabled = False
126         Textboxes(2).SetFocus
        End If

        '----------------------------------------------------------------------------------------------------------------
128     If Index = 1 Then    'ΑΠΟΘΗΚΕΥΣΗ
130         save_record
        End If

        '-------------------------------------------------------------------------------------------------------------------
132     If Index = 2 Then    'διαγραφη
134         ANS = MsgBox("Είστε σίγουροι για την διαγραφή;", vbYesNo)

136         If ANS = vbYes Then
138             CurRow = grid.row    '+ 1
                ' MKOD = grid.TextMatrix(CURROW, 0) ' kodikos   'flexgrid
140             mkod = grid.Columns(0).Text
142             Gdb.Execute delSql + "  " + mfields(1) + "=" + mkod
144             grid.Visible = True
146             Adodc1.Refresh
148             BUT(0).Enabled = True    ' add
150             BUT(1).Enabled = False    'save
            End If

152         For i = 1 To n
154             Labels(i).Visible = False
156             Textboxes(i).Visible = False
            Next

        End If

158     If Index = 3 Then    ' ΕΞΟΔΟσ
160         Unload Me

        End If

        Exit Sub

LatosTypoi:
162     MsgBox "λάθος στο πεδίο " + Labels(i)
164     Textboxes(i).SetFocus

NOSAVE:


        '<EhFooter>
        Exit Sub

BUT_Click_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.BUT_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub save_record()
    '====================================================================
        '<EhHeader>
        On Error GoTo save_record_Err
        '</EhHeader>
Dim mkod
Dim R As New ADODB.Recordset
100         mkod = Textboxes(1).Text  ' id

               Dim m_ID As String
102         If F_NEW = 1 Then
                'R.AddNew
104             F_NEW = 0
                ' Gdb.Execute "INSERT INTO PARTIDES (KOD) VALUES ('')   "
106             Gdb.Execute "INSERT INTO PARTIDES (KOD) VALUES ('')   "
108             R.Open "SELECT MAX(ID) FROM PARTIDES", Gdb, adOpenKeyset, adLockOptimistic
110             mkod = str(R(0))
112             Textboxes(1) = mkod
114             R.Close
                
            End If
            
116          R.Open sql + " where ID=" + mkod, Gdb, adOpenDynamic, adLockOptimistic
118          m_ID = str(R!ID)
            

            

            

            On Error Resume Next ' GoTo LatosTypoi

120         For i = 2 To n

122              If ParamGrid.TextMatrix(i, 2) = 1 Then ' aritmos
124                  R(mfields(i)) = Textboxes(i).Text
126              ElseIf ParamGrid.TextMatrix(i, 2) = 2 Then 'string
128                   R(mfields(i)) = Textboxes(i).Text
                    ' R(mfields(i)) = """" + Textboxes(i).Text + """"
130              ElseIf ParamGrid.TextMatrix(i, 2) = 3 Then  'date
132                  R(mfields(i)) = CDate(Textboxes(i).Text)
                 End If
            Next
134          For i = 1 To n
136             Labels(i).Visible = False
138             Textboxes(i).Visible = False
            Next
            
            
            

            'ΠΟΙΟ ΠΕΔΙΟ ΕΝΗΜΕΡΩΝΩ ΧΩΡΙΣ ΝΑ ΒΛΕΠΕΙ Ο ΧΡΗΣΤΗΣ
140         If Len(F_FIELD_EPIPLEON) > 0 Then
142             R(F_FIELD_EPIPLEON) = F_TIMH
            End If

            'ΠΟΙΟ ΠΕΔΙΟ ΕΝΗΜΕΡΩΝΩ ΧΩΡΙΣ ΝΑ ΒΛΕΠΕΙ Ο ΧΡΗΣΤΗΣ
144         If Len(F_2FIELD_EPIPLEON) > 0 Then
146             R(F_2FIELD_EPIPLEON) = F_2TIMH
            End If

            On Error Resume Next

148         R.Update

            Dim X As String

150         If Len(UpdSQL) > 0 Then
                'X = mac(UpdSQL)
                'gdb.Execute UpdSQL
152             trexe UpdSQL
            End If

154         grid.Top = BUT(0).Top + BUT(0).Height + TextHeight("A")
156         grid.Visible = True
158         Adodc1.Refresh
160         paintgrid
162         BUT(0).Enabled = True    ' add
164         BUT(1).Enabled = False    'save
166         BUT(2).Enabled = False    'DELETE
168         BUT(0).SetFocus

            Exit Sub

        '<EhFooter>
        Exit Sub

save_record_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.save_record " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub






Private Sub Command1_Click()
 runscript
End Sub

'Private Sub Command1_Click()
'    trexe "AA"
'End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>


100     flagProthFora = 1
102     F_NEW = 0

104     Me.KeyPreview = True
        '  Me.Picture = LoadPicture(gPicture)

        '    'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        '    'SkinFramework1.ApplyWindow Me.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics


        '<EhFooter>
        Exit Sub

Form_Load_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Paint()
        '<EhHeader>
        On Error GoTo Form_Paint_Err
        '</EhHeader>

        Dim K As Integer

100     If flagProthFora = 1 Then

102         PaintForm
104         paintgrid

106         Set fSCR = CreateObject("MSScriptControl.ScriptControl")
108         fSCR.language = "vbscript"
110         fSCR.AddObject "mactext", mactext

112         For K = 1 To n
114             fSCR.AddObject "t" + Format(K, "00"), Textboxes(K)
            Next

116         flagProthFora = 0
        End If


        '<EhFooter>
        Exit Sub

Form_Paint_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.Form_Paint " & _
               "at line " & Erl
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
102     If Adodc1.Recordset.EOF Then Exit Sub

104     mkod = grid.Columns(0).Text

        Dim R As New ADODB.Recordset

        Dim i As Integer

106             R.Open sql + " where " + mfields(1) + "= " + mkod + " ", Gdb, adOpenKeyset, adLockOptimistic

108     For i = 1 To n

110         If IsNull(R(mfields(i))) Then
112             Textboxes(i).Text = ""
            Else
114             Textboxes(i).Text = R(mfields(i))
            End If

116         Labels(i).Visible = True
118         Textboxes(i).Visible = True
        Next

120     grid.Visible = False

122     BUT(0).Enabled = False    ' add
124     BUT(1).Enabled = True    'save
126     BUT(2).Enabled = True    'DELETE

128     Textboxes(1).Enabled = False
130     Textboxes(2).SetFocus


        '<EhFooter>
        Exit Sub

grid_DblClick_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.grid_DblClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub grid_KeyUp(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo grid_KeyUp_Err
        '</EhHeader>

100     If KeyCode = 13 Then
102         grid_DblClick
        End If


        '<EhFooter>
        Exit Sub

grid_KeyUp_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.grid_KeyUp " & _
               "at line " & Erl
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
102     Textboxes(Index).SelLength = Len(Textboxes(Index).Text)

104     Textboxes(Index).BackColor = vbYellow


        '<EhFooter>
        Exit Sub

Textboxes_GotFocus_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.Textboxes_GotFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Textboxes_KeyPress(Index As Integer, KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo Textboxes_KeyPress_Err
        '</EhHeader>

100     If KeyAscii = vbKeyReturn Then
102         KeyAscii = 0    'suppress the beep
104         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If


        '<EhFooter>
        Exit Sub

Textboxes_KeyPress_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.Textboxes_KeyPress " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Textboxes_LostFocus(Index As Integer)
        '<EhHeader>
        On Error GoTo Textboxes_LostFocus_Err
        '</EhHeader>

        Dim X

100     X = ParamGrid.TextMatrix(Index, 2)

102     If Len(X) > 0 Then
104         trexe X
        End If

106     Textboxes(Index).BackColor = Textboxes(0).BackColor


        '<EhFooter>
        Exit Sub

Textboxes_LostFocus_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.Textboxes_LostFocus " & _
               "at line " & Erl
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
102     ' Call fSCR.Run("Main") ' καλω ολοκληρη procedure


        '<EhFooter>
        Exit Sub

trexe_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.trexe " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub runscript() ' (ByVal Scode As String)
Dim Scode As String
         mactext.Text = gConnect  '= "DSN=MERC2"
         Scode = "Sub Main()" & vbCrLf & _
         "Msgbox t02 " & vbCrLf & _
         " DIM mCon  " & vbCrLf & _
         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
         " mCon.Open """ + mactext.Text + """ " & vbCrLf & _
         "DIM R " & vbCrLf & _
         " Set R=CreateObject(""ADODB.Recordset"")" & vbCrLf & _
         " r.Open "" SELECT ONO FROM EID WHERE KOD='002' "",mCon " & vbCrLf & _
         "Msgbox r(0) " & vbCrLf & _
         "T03=r(0) " & vbCrLf & _
         "Set mCon=Nothing" & vbCrLf & _
         "End Sub"
        
'

'         "Set r=Nothing" & vbCrLf & _

        
        
        
        
        
        
'          "  mCon.execute ""UPDATE ARITMISI2 SET ARITMISI=457  WHERE ID= 7 "" " & vbCrLf & _
        '<EhHeader>
        On Error GoTo trexe_Err
        '</EhHeader>


100     Call fSCR.AddCode(Scode)
102     Call fSCR.Run("Main") ' καλω ολοκληρη procedure


        '<EhFooter>
        Exit Sub

trexe_Err:
      SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.DataEntry3.trexe " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub



'Private Sub trexe(ByVal Scode As String)
'        ' gconnect = "DSN=MERC2"
'        'scode = "Sub Main()" & vbCrLf & _
'         "Msgbox 15 " & vbCrLf & _
'         " DIM mCon  " & vbCrLf & _
'         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
'         " mCon.Open """ + gconnect + """ " & vbCrLf & _
'         " mCon.execute ""UPDATE ARITMISI2 SET ARITMISI=457  WHERE ID= 7 "" " & vbCrLf & _
'         "DIM R " & vbCrLf & _
'         " Set R=CreateObject(""ADODB.Recordset"")" & vbCrLf & _
'         " r.Open "" SELECT ARITMISI FROM ARITMISI2 WHERE ID=7 "",mCon " & vbCrLf & _
'         "Msgbox r(0) " & vbCrLf & _
'         "Set r=Nothing" & vbCrLf & _
'         "Set mCon=Nothing" & vbCrLf & _
'         "End Sub"
'
'        '<EhHeader>
'        On Error GoTo trexe_Err
'
'        '</EhHeader>
'
'100     Call fSCR.AddCode(Scode)
'110     Call fSCR.Run("Main") ' AN UELV NA DHMIOYRHSV MIA PROCEDURE
'
'        '<EhFooter>
'        Exit Sub
'
'
