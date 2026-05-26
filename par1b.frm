VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form par1b 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6765
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPalia 
      Caption         =   "–·ÎÈ‹ ◊ÒÔÌÈ‹"
      Height          =   300
      Left            =   480
      TabIndex        =   7
      Top             =   4680
      Width           =   2655
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      ItemData        =   "par1b.frx":0000
      Left            =   240
      List            =   "par1b.frx":0002
      MultiSelect     =   1  'Simple
      TabIndex        =   6
      Top             =   0
      Width           =   9210
   End
   Begin VB.CommandButton Command2 
      Caption         =   "¡Ì·Ì›˘ÛÁ"
      Height          =   360
      Left            =   7440
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.TextBox parast 
      Height          =   372
      Left            =   1560
      TabIndex        =   3
      Text            =   "'t','T','Y'"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox PEL 
      Height          =   288
      Left            =   480
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈ÈÎÔ„ﬁ ”˜ÂÙÈÊ¸ÏÂÌÔı –·Ò·ÛÙ·ÙÈÍÔ˝"
      Height          =   360
      Left            =   480
      TabIndex        =   1
      Top             =   5040
      Width           =   6732
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "par1b.frx":0004
      Height          =   4455
      Left            =   12240
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7858
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.5
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   2160
      Top             =   6000
      Visible         =   0   'False
      Width           =   3612
      _ExtentX        =   6376
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
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "par1b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkPalia_Click()
  List1.Clear
  If chkPalia.Value = vbChecked Then    '   <> gDir Then ' palia xronia
        f_PALIAXRONIA = Trim(FINDPARAMETROI(1, "PAR1", "F_PALIAXRONIA", "2005", "√.18 DSN –¡À…¡” ◊—œÕ…¡”"))
        Dim GTGDB As New ADODB.Connection
        GTGDB.Open gConnect + ";DATABASE=" + f_PALIAXRONIA
        Adodc1.ConnectionString = gConnect + ";DATABASE=" + f_PALIAXRONIA
        
     Else
        Adodc1.ConnectionString = gConnect
        'GTGDB.Open gConnect
     End If


    Adodc1.RecordSource = "select ATIM,CONVERT(CHAR(10),HME,103) AS HMEP , AJI AS [”’Õ.¡Œ…¡],ENTITYMARK FROM TIM WHERE KPE='" + PEL.Text + "' AND LEFT(ATIM,1) IN (select EIDOS from PARASTAT WHERE ISMYDATA=1) order by HME DESC"
   ' Adodc1.ConnectionString = gConnect
   Adodc1.Refresh
   
   Do While Not Adodc1.Recordset.EOF
        List1.AddItem Adodc1.Recordset("atim") + " ; " + Adodc1.Recordset("HMEP") + " ;" + CNull(Adodc1.Recordset("ENTITYMARK")) + "; " + Format(Adodc1.Recordset("”’Õ.¡Œ…¡"), "#####0.00")
   
   
   
        Adodc1.Recordset.MoveNext
   
   Loop
   
   



End Sub

Private Sub Command1_Click()

'  PAR1.par11Value.Caption = "* " + Trim(Adodc1.Recordset("atim")) + " " + Format(Adodc1.Recordset("hmep"), "DD/mm/yyyy") + " " + Trim$(ID_NUM)
   
   Dim C As String
   C = CNull(Adodc1.Recordset("ENTITYMARK"))
   If Left(C, 3) = "ƒœ " Then
       C = mID(C, 4, 15)
   End If
   
  On Error Resume Next
  
  
   Dim N As Integer
      Dim C0 As String
      C0 = ""
      For N = 0 To List1.ListCount - 1
        'Adodc1.Recordset("atim") + " ; " + Adodc1.Recordset("HMEP") + " ;" + Adodc1.Recordset("ENTITYMARK") + " ;" + Format(Adodc1.Recordset("”’Õ.¡Œ…¡"), "#####0.00")
        If List1.Selected(N) Then
            C = Split(List1.List(N), ";")(2)
            If Left(C, 3) = "ƒœ " Then
               C = mID(C, 4, 15)
            End If
            C0 = C0 + C + ";"
        End If
      Next
  
  
  
  
 If Label1.Caption = "0" Then
     PAR1.par11Value.Caption = C0 '   C + "* " + Trim(Adodc1.Recordset("atim")) + " " + Format(Adodc1.Recordset("hmep"), "DD/mm/yyyy") + " " + Trim$(ID_NUM)
  Else
     
     ' KALESTHKE APO PAR7MYDATA
     Par7MyData.SXETMART.Caption = C0
  End If
   
   
   Unload Me
   
'   Adodc1.RecordSource = "select ATIM,CONVERT(CHAR(10),HME,103) AS HMEP , AJI AS [”’Õ.¡Œ…¡] FROM TIM WHERE KPE='" + PEL.Text + "' AND LEFT(ATIM,1) IN (" + parast + ") order by HME DESC"
'   Adodc1.ConnectionString = gConnect
'   Adodc1.Refresh
'
'
   

End Sub

Private Sub Command2_Click()



'   Adodc1.RecordSource = "select ATIM,CONVERT(CHAR(10),HME,103) AS HMEP , AJI AS [”’Õ.¡Œ…¡],ID_NUM FROM TIM WHERE KPE='" + PEL.Text + "' AND LEFT(ATIM,1) IN (" + parast + ") order by HME DESC"
'   Adodc1.ConnectionString = gConnect
'   Adodc1.Refresh

End Sub

Private Sub Form_Paint()
   
   
   
   


   



   

   Adodc1.RecordSource = "select ATIM,CONVERT(CHAR(10),HME,103) AS HMEP , AJI AS [”’Õ.¡Œ…¡],ENTITYMARK FROM TIM WHERE KPE='" + PEL.Text + "' AND LEFT(ATIM,1) IN (select EIDOS from PARASTAT WHERE ISMYDATA=1) order by HME DESC"
   Adodc1.ConnectionString = gConnect
   Adodc1.Refresh
   
   Do While Not Adodc1.Recordset.EOF
        List1.AddItem Adodc1.Recordset("atim") + " ; " + Adodc1.Recordset("HMEP") + " ;" + CNull(Adodc1.Recordset("ENTITYMARK")) + "; " + Format(Adodc1.Recordset("”’Õ.¡Œ…¡"), "#####0.00")
   
   
   
        Adodc1.Recordset.MoveNext
   
   Loop
   
   
   
   
   
   
   
   
End Sub

'<IncomeClassificationsDoc xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:icls="https://www.aade.gr/myDATA/incomeClassificaton/v1.0" xmlns="https://www.aade.gr/myDATA/incomeClassificaton/v1.0">
'<incomeInvoiceClassification>
'<invoiceMark>400000017540203</invoiceMark>
'<invoicesIncomeClassificationDetails>
'<lineNumber>1</lineNumber>
'<incomeClassificationDetailData>
'<classificationType>E3_562</classificationType>
'<classificationCategory>category1_5</classificationCategory>
'<amount>2000.00</amount>
'<id>1</id>
'</incomeClassificationDetailData>
'</invoicesIncomeClassificationDetails>
'</incomeInvoiceClassification>
'</IncomeClassificationsDoc>
