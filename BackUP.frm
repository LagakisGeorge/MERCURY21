VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form BOHU82 
   AutoRedraw      =   -1  'True
   Caption         =   "Backup / Restore with Events"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8385
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   6840
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Data File Name:"
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   495
      Left            =   6960
      TabIndex        =   20
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox cmbDatabaseName 
      Height          =   315
      Left            =   4560
      TabIndex        =   18
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtDataFileName 
      Height          =   288
      Left            =   4560
      TabIndex        =   16
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Frame frmConnectionInfo 
      Caption         =   "Connection:"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtServerName 
         Height          =   288
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblServer 
         Caption         =   "Server:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frmAuthorization 
      Caption         =   "Authorization:"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
      Begin VB.TextBox txtUserName 
         Height          =   288
         Left            =   1320
         TabIndex        =   9
         Top             =   1080
         Width           =   2292
      End
      Begin VB.TextBox txtPassword 
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1440
         Width           =   2292
      End
      Begin VB.OptionButton optWinNTAuth 
         Caption         =   "Use Windows NT authentication"
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2892
      End
      Begin VB.OptionButton optSSAuth 
         Caption         =   "Use SQL Server authentication"
         Height          =   252
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   3252
      End
      Begin VB.Label lblUserName 
         Caption         =   "Login name:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.TextBox txtStatus 
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   3960
      Width           =   7815
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Backup"
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Backup/Restore File Name:"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Database To Backup/Restore:"
      Height          =   255
      Left            =   4560
      TabIndex        =   17
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lblQueryResults 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   615
   End
End
Attribute VB_Name = "BOHU82"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Backup Restore with Events Sample Application
' Microsoft Copyright 2000
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Dim gSQLServer As SQLDMO.SQLServer
Dim WithEvents oBackupEvent As SQLDMO.Backup
Attribute oBackupEvent.VB_VarHelpID = -1
Dim WithEvents oRestoreEvent As SQLDMO.Restore
Attribute oRestoreEvent.VB_VarHelpID = -1

Dim gbConnected As Boolean
Dim gDatabaseName As String
Dim gBkupRstrFileName As String
Dim gBkupRstrFilePath As String
Dim FFILE  ' FILENAME TXT WITH SETTINGS
Const gTitle = "Server Connection"



Private Sub Command1_Click()
  
'USE master
'GO
'
'-- the original database (use 'SET @DB = NULL' to disable backup)
'DECLARE @DB varchar(200)
'SET @DB = 'MERCURY'
'
'-- the backup filename
'DECLARE @BackupFile varchar(2000)
'SET @BackupFile = 'c:\BACC\backup.dat'
'
'-- the new database name
'DECLARE @TestDB  varchar(200)
'SET @TestDB = 'TestDB2'
'
'-- the new database files without .mdf/.ldf
'DECLARE @RestoreFile varchar(2000)
'SET @RestoreFile = 'c:\BACC\TESTDB2'
'
'
'-- ****************************************************************
'--                    no change below this line
'-- ****************************************************************
'
'
'DECLARE @query varchar(2000)
'
'DECLARE @DataFile varchar(2000)
'SET @DataFile = @RestoreFile + '.mdf'
'
'DECLARE @LogFile varchar(2000)
'SET @LogFile = @RestoreFile + '.ldf'
'
'IF @DB IS NOT NULL
'BEGIN
'    SET @query = 'BACKUP DATABASE ' + @DB + ' TO DISK = ' +
'QUOTENAME(@BackupFile, '''')
'    EXEC (@query)
'End
'
'-- RESTORE FILELISTONLY FROM DISK = 'C:\temp\backup.dat'
'-- RESTORE HEADERONLY FROM DISK = 'C:\temp\backup.dat'
'-- RESTORE LABELONLY FROM DISK = 'C:\temp\backup.dat'
'-- RESTORE VERIFYONLY FROM DISK = 'C:\temp\backup.dat'
'
'IF EXISTS(SELECT * FROM sysdatabases WHERE name = @TestDB)
'BEGIN
'    SET @query = 'DROP DATABASE ' + @TestDB
'    EXEC (@query)
'End
'
'RESTORE HEADERONLY FROM DISK = @BackupFile
'DECLARE @File int
'SET @File = @@ROWCOUNT
'
'DECLARE @Data varchar(500)
'DECLARE @Log varchar(500)
'
'SET @query = 'RESTORE FILELISTONLY FROM DISK = ' + QUOTENAME(@BackupFile, '''')
'
'CREATE TABLE #restoretemp
'(
'LogicalName varchar(500),
'PhysicalName varchar(500),
'type varchar(10),
'FilegroupName varchar(200),
'size int,
'MaxSize bigint
')
'INSERT #restoretemp EXEC (@query)
'
'SELECT @Data = LogicalName FROM #restoretemp WHERE type = 'D'
'SELECT @Log = LogicalName FROM #restoretemp WHERE type = 'L'
'
'PRINT @Data
'PRINT @Log
'
'TRUNCATE TABLE #restoretemp
'DROP TABLE #restoretemp
'
'IF @File > 0
'BEGIN
'    SET @query = 'RESTORE DATABASE ' + @TestDB + ' FROM DISK = ' +
'QUOTENAME(@BackupFile, '''') +
'        ' WITH MOVE ' + QUOTENAME(@Data, '''') + ' TO ' +
'QUOTENAME(@DataFile, '''') + ', MOVE ' +
'        QUOTENAME(@Log, '''') + ' TO ' + QUOTENAME(@LogFile, '''') + ',
'FILE = ' + CONVERT(varchar, @File)
'    EXEC (@query)
'End
'GO
'
'
'
'-----Inline Attachment Follows-----

End Sub

Private Sub Form_Load()
' Me.Picture = LoadPicture(gPicture)

    Set gSQLServer = Nothing
    optWinNTAuth.Value = True
    gbConnected = False
    WinNTAuthOptionsOn
    buttonsConnectClosed
    FFILE = "C:\MERCVB\BACKUP.TXT"

   If Len(Dir(FFILE, vbNormal)) > 0 Then


Dim A
10 Open FFILE For Input As #1

    Line Input #1, A: txtServerName.Text = A
    Line Input #1, A: txtUserName = A
    Line Input #1, A: txtPassword = A
    Line Input #1, A: cmbDatabaseName = A
    Line Input #1, A: txtDataFileName = A
   End If
40 Close #1



End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbConnected = True Then
        Call gSQLServer.Disconnect
    End If
    If Not gSQLServer Is Nothing Then
        Set gSQLServer = Nothing
    End If


10 Open FFILE For Output As #1

    Print #1, txtServerName
    Print #1, txtUserName
    Print #1, txtPassword
    Print #1, cmbDatabaseName
    Print #1, txtDataFileName

40 Close #1



End Sub



Private Sub cmdConnect_Click()
    Dim ServerName As String
    Dim UserName As String
    Dim Password As String
    Dim A As String, b As String

    On Error GoTo ErrHandler:

    If gSQLServer Is Nothing Then
        Set gSQLServer = New SQLDMO.SQLServer
    End If

    ' Put text box values into connection variables.
    ServerName = txtServerName.Text
    UserName = txtUserName.Text
    Password = txtPassword.Text

    ' Set the login timeout.
    gSQLServer.LoginTimeout = 15

    ' Decision code for login authorization type: WinNT or SQL Server.
    If optWinNTAuth.Value = True Then
        gSQLServer.LoginSecure = True
    End If

    ' Change mousepointer while trying to connect.
    Screen.MousePointer = vbHourglass

    gSQLServer.Connect ServerName, UserName, Password
    gbConnected = True
    A = cmbDatabaseName
    b = txtDataFileName
    ' List all of the database names.
    FillDatabaseList

    cmbDatabaseName.Text = A
    txtDataFileName = b
    ' Change mousepointer back to the default after connect.
    Screen.MousePointer = vbDefault

    ' Notify user that connection was successful.
    MsgBox "Connection to server successful.", vbOKOnly, gTitle

    buttonsConnectOpen

    ' Clear up the status text in the "result field".
    txtStatus.Text = ""

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Description

    ' Change mousepointer back if it's hourglass.
    If Screen.MousePointer = vbHourglass Then
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub cmdDisconnect_Click()
    On Error GoTo ErrHandler:

    Dim Msg As String
    Dim Response As String

    ' Disconnect from a connected server.
    If gbConnected = True Then
        Msg = "Disconnect from Server?"
        Response = MsgBox(Msg, vbOKCancel, gTitle)
        If Response = vbOK Then
            Call gSQLServer.Disconnect
            Set gSQLServer = Nothing
            cmbDatabaseName.Clear
            txtDataFileName.Text = ""
            txtStatus.Text = ""
            gbConnected = False
            buttonsConnectClosed
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Description
    Resume Next
End Sub

Private Sub cmdBackup_Click()
    On Error GoTo ErrHandler:

    Dim oBackup As SQLDMO.Backup

    gDatabaseName = cmbDatabaseName.Text
    Set oBackup = New SQLDMO.Backup
    Set oBackupEvent = oBackup ' enable events

    oBackup.Database = gDatabaseName
    gBkupRstrFileName = txtDataFileName.Text
    oBackup.Files = gBkupRstrFileName

    ' Delete the datafile to allow the application to create a brand new file.
    ' This will prevent attaching the new backup data to the old data if there
    ' is any.
    If Len(Dir(gBkupRstrFileName)) > 0 Then
        Kill (gBkupRstrFileName)
    End If

    ' Change mousepointer while trying to connect.
    Screen.MousePointer = vbHourglass

    ' Backup the database.
    oBackup.SQLBackup gSQLServer

    ' Change mousepointer back to the default after connect.
    Screen.MousePointer = vbDefault

    Set oBackupEvent = Nothing ' disable events
    Set oBackup = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Description
    Resume Next
End Sub

Private Sub cmdRestore_Click()
    On Error GoTo ErrHandler:

    Dim oRestore As SQLDMO.Restore

    Dim Msg As String
    Dim Response As String




'gConnect = gConnect + ";DATABASE=master"
'MDIForm1.Timer1.Enabled = False
 '  Gdb.Close
    '  Gdb.Open gConnect
   
   
 'On Error Resume Next
  '  Gdb.Close
    
 'On Error GoTo OUTSIDE
   'Gdb.Open gConnect













'    Msg = "You must choose the right database name according to the data file name selected. Do you want to continue?"
'    Response = MsgBox(Msg, vbYesNo, gTitle)
'    If Response = vbNo Then
'        Exit Sub
'    End If

    gDatabaseName = cmbDatabaseName.Text
    Set oRestore = New SQLDMO.Restore
    Set oRestoreEvent = oRestore        ' enable events

    oRestore.Database = gDatabaseName
    gBkupRstrFileName = txtDataFileName.Text
    oRestore.Files = gBkupRstrFileName
    oRestore.ReplaceDatabase = True
      
    ' Change mousepointer while trying to connect.
    Screen.MousePointer = vbHourglass

    ' Restore the database.
    oRestore.SQLRestore gSQLServer

    ' Change mousepointer back to the default after connect.
    Screen.MousePointer = vbDefault

    Set oRestoreEvent = Nothing         ' disable events
    Set oRestore = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Description
    Resume Next
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo ErrHandler:

    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "All Files (*.*)|*.*|Backup Files (*.bak)|*.bak"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.InitDir = gBkupRstrFilePath
    CommonDialog1.DefaultExt = "bak"
    CommonDialog1.DialogTitle = "Data File Name:"
    CommonDialog1.Action = 1
    txtDataFileName.Text = CommonDialog1.FileName
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub



' VB will create the right prototypes for you, if you select the oBackupEvent in
' the drop down listbox of your editor
Private Sub oBackupEvent_Complete(ByVal Message As String)
    PrintStat "oBackupEvent_Complete -- " & Message
End Sub

Private Sub oBackupEvent_NextMedia(ByVal Message As String)
    PrintStat "oBackupEvent_NextMedia -- " & Message
End Sub

Private Sub oBackupEvent_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    PrintStat "oBackupEvent_PercentComplete -- " & Message & " " & Percent
End Sub

Private Sub oRestoreEvent_Complete(ByVal Message As String)
    PrintStat "oRestoreEvent_Complete -- " & Message
End Sub

Private Sub oRestoreEvent_NextMedia(ByVal Message As String)
    PrintStat "oRestoreEvent_NextMedia -- " & Message
End Sub

Private Sub oRestoreEvent_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    PrintStat "oRestoreEvent_PercentComplete -- " & Message & " " & Percent
End Sub

Private Sub PrintStat(ByRef Message As String)
    txtStatus.Text = txtStatus.Text + Message + vbCrLf
End Sub



Private Sub optSSAuth_Click()
    If optSSAuth.Value = True Then
        SSAuthOptionsOn
    End If
End Sub

Private Sub optWinNTAuth_Click()
    optWinNTAuth.Value = True
    WinNTAuthOptionsOn
    txtUserName.Text = ""
    txtPassword.Text = ""
End Sub

Private Sub buttonsConnectClosed()
    cmdConnect.Default = True

    cmdConnect.Enabled = True
    cmdBackup.Enabled = False
    cmdRestore.Enabled = False
    cmdDisconnect.Enabled = False

    cmdBrowse.Enabled = False
    cmbDatabaseName.Enabled = False
    txtDataFileName.Enabled = False

    ' Enable the Authorization stuff.
    optWinNTAuth.Enabled = True
    optSSAuth.Enabled = True
    txtServerName.Enabled = True
    lblServer.Enabled = True
    If optWinNTAuth = True Then
        WinNTAuthOptionsOn
    Else
        SSAuthOptionsOn
    End If
End Sub

Private Sub buttonsConnectOpen()
    cmdConnect.Enabled = False
    cmdBackup.Enabled = True
    cmdRestore.Enabled = True
    cmdDisconnect.Enabled = True

    cmdBrowse.Enabled = True
    cmbDatabaseName.Enabled = True
    txtDataFileName.Enabled = True

    ' Disable the Authorization stuff.
    optWinNTAuth.Enabled = False
    optSSAuth.Enabled = False
    txtServerName.Enabled = False
    lblServer.Enabled = False
    lblUserName.Enabled = False
    lblPassword.Enabled = False
    txtUserName.Enabled = False
    txtPassword.Enabled = False
End Sub

Private Sub WinNTAuthOptionsOn()
    lblUserName.Enabled = False
    lblPassword.Enabled = False
    txtUserName.Enabled = False
    txtPassword.Enabled = False
End Sub

Private Sub SSAuthOptionsOn()
    lblUserName.Enabled = True
    lblPassword.Enabled = True
    txtUserName.Enabled = True
    txtPassword.Enabled = True
End Sub




Private Sub FillDatabaseList()
    cmbDatabaseName.Clear

    ' Enumerate all of the databases and add the names to the list box.
    Dim oDB As SQLDMO.Database
    For Each oDB In gSQLServer.Databases
        If oDB.SystemObject = False Then
            cmbDatabaseName.AddItem oDB.Name
        End If
    Next oDB

    ' Take care of the assignment of gBkupRstrFilePath.
    Dim MyPos As Integer
    gBkupRstrFilePath = CurDir
    MyPos = InStr(1, CurDir, "DevTools", 1)
    If MyPos > 0 Then
        gBkupRstrFilePath = Left(gBkupRstrFilePath, MyPos - 1)
        If Len(Dir(gBkupRstrFilePath + "backup", vbDirectory)) Then
            gBkupRstrFilePath = gBkupRstrFilePath + "backup\"
        Else
            gBkupRstrFilePath = "c:\"
        End If
    Else
        gBkupRstrFilePath = "c:\"
    End If

    ' Select the first database in the list.
    If cmbDatabaseName.ListCount > 0 Then
        cmbDatabaseName.ListIndex = 0
        ' Assign the default backup/restore file name.
        If Len(cmbDatabaseName.Text) > 0 Then
            txtDataFileName.Text = gBkupRstrFilePath + cmbDatabaseName.Text + ".bak"
        End If
    End If

End Sub


Private Sub cmbDatabaseName_Click()
    ' Assign the default backup/restore file name.
    If Len(cmbDatabaseName.Text) > 0 Then
        txtDataFileName.Text = gBkupRstrFilePath + cmbDatabaseName.Text + ".bak"
    End If
End Sub
'USE master
'GO
'
'-- the original database (use 'SET @DB = NULL' to disable backup)
'DECLARE @DB varchar(200)
'SET @DB = 'PcTopp'
'
'-- the backup filename
'DECLARE @BackupFile varchar(2000)
'SET @BackupFile = 'c:\pctopp\sqlserver\backup.dat'
'
'-- the new database name
'DECLARE @TestDB varchar(200)
'SET @TestDB = 'TestDB'
'
'-- the new database files without .mdf/.ldf
'DECLARE @RestoreFile varchar(2000)
'SET @RestoreFile = 'c:\pctopp\sqlserver\backup'
'
'
'-- ****************************************************************
'--                    no change below this line
'-- ****************************************************************
'
'
'DECLARE @query varchar(2000)
'
'DECLARE @DataFile varchar(2000)
'SET @DataFile = @RestoreFile + '.mdf'
'
'DECLARE @LogFile varchar(2000)
'SET @LogFile = @RestoreFile + '.ldf'
'
'IF @DB IS NOT NULL
'BEGIN
'    SET @query = 'BACKUP DATABASE ' + @DB + ' TO DISK = ' +
'QUOTENAME(@BackupFile, '''')
'    EXEC (@query)
'End
'
'-- RESTORE FILELISTONLY FROM DISK = 'C:\temp\backup.dat'
'-- RESTORE HEADERONLY FROM DISK = 'C:\temp\backup.dat'
'-- RESTORE LABELONLY FROM DISK = 'C:\temp\backup.dat'
'-- RESTORE VERIFYONLY FROM DISK = 'C:\temp\backup.dat'
'
'IF EXISTS(SELECT * FROM sysdatabases WHERE name = @TestDB)
'BEGIN
'    SET @query = 'DROP DATABASE ' + @TestDB
'    EXEC (@query)
'End
'
'RESTORE HEADERONLY FROM DISK = @BackupFile
'DECLARE @File int
'SET @File = @@ROWCOUNT
'
'DECLARE @Data varchar(500)
'DECLARE @Log varchar(500)
'
'SET @query = 'RESTORE FILELISTONLY FROM DISK = ' + QUOTENAME(@BackupFile
', '''')
'
'CREATE TABLE #restoretemp
'(
' LogicalName varchar(500),
' PhysicalName varchar(500),
' type varchar(10),
' FilegroupName varchar(200),
' size int,
' MaxSize bigint
')
'INSERT #restoretemp EXEC (@query)
'
'SELECT @Data = LogicalName FROM #restoretemp WHERE type = 'D'
'SELECT @Log = LogicalName FROM #restoretemp WHERE type = 'L'
'
'PRINT @Data
'PRINT @Log
'
'TRUNCATE TABLE #restoretemp
'DROP TABLE #restoretemp
'
'IF @File > 0
'BEGIN
'    SET @query = 'RESTORE DATABASE ' + @TestDB + ' FROM DISK = ' +
'QUOTENAME(@BackupFile, '''') +
'        ' WITH MOVE ' + QUOTENAME(@Data, '''') + ' TO ' +
'QUOTENAME(@DataFile, '''') + ', MOVE ' +
'        QUOTENAME(@Log, '''') + ' TO ' + QUOTENAME(@LogFile, '''') + ',
'FILE = ' + CONVERT(varchar, @File)
'    EXEC (@query)
'End
'GO
'
'Plain Text Attachment [ Download File | Save to Yahoo! Briefcase ]
'
'USE master
'GO
'
'-- the original database (use 'SET @DB = NULL' to disable backup)
'DECLARE @DB varchar(200)
'SET @DB = 'PcTopp'
'
'-- the backup filename
'DECLARE @BackupFile varchar(2000)
'SET @BackupFile = 'c:\pctopp\sqlserver\backup.dat'
'
'-- the new database name
'DECLARE @TestDB varchar(200)
'SET @TestDB = 'TestDB'
'
'-- the new database files without .mdf/.ldf
'DECLARE @RestoreFile varchar(2000)
'SET @RestoreFile = 'c:\pctopp\sqlserver\backup'
'
'
'-- ****************************************************************
'--                    no change below this line
'-- ****************************************************************
'
'
'DECLARE @query varchar(2000)
'
'DECLARE @DataFile varchar(2000)
'SET @DataFile = @RestoreFile + '.mdf'
'
'DECLARE @LogFile varchar(2000)
'SET @LogFile = @RestoreFile + '.ldf'
'
'IF @DB IS NOT NULL
'BEGIN
'    SET @query = 'BACKUP DATABASE ' + @DB + ' TO DISK = ' +
'QUOTENAME(@BackupFile, '''')
'    EXEC (@query)
'End
'
'-- RESTORE FILELISTONLY FROM DISK = 'C:\temp\backup.dat'
'-- RESTORE HEADERONLY FROM DISK = 'C:\temp\backup.dat'
'-- RESTORE LABELONLY FROM DISK = 'C:\temp\backup.dat'
'-- RESTORE VERIFYONLY FROM DISK = 'C:\temp\backup.dat'
'
'IF EXISTS(SELECT * FROM sysdatabases WHERE name = @TestDB)
'BEGIN
'    SET @query = 'DROP DATABASE ' + @TestDB
'    EXEC (@query)
'End
'
'RESTORE HEADERONLY FROM DISK = @BackupFile
'DECLARE @File int
'SET @File = @@ROWCOUNT
'
'DECLARE @Data varchar(500)
'DECLARE @Log varchar(500)
'
'SET @query = 'RESTORE FILELISTONLY FROM DISK = ' + QUOTENAME(@BackupFile ,
''''')
'
'CREATE TABLE #restoretemp
'(
' LogicalName varchar(500),
' PhysicalName varchar(500),
' type varchar(10),
' FilegroupName varchar(200),
' size int,
' MaxSize bigint
')
'INSERT #restoretemp EXEC (@query)
'
'SELECT @Data = LogicalName FROM #restoretemp WHERE type = 'D'
'SELECT @Log = LogicalName FROM #restoretemp WHERE type = 'L'
'
'PRINT @Data
'PRINT @Log
'
'TRUNCATE TABLE #restoretemp
'DROP TABLE #restoretemp
'
'IF @File > 0
'BEGIN
'    SET @query = 'RESTORE DATABASE ' + @TestDB + ' FROM DISK = ' +
'QUOTENAME(@BackupFile, '''') +
'        ' WITH MOVE ' + QUOTENAME(@Data, '''') + ' TO ' + QUOTENAME(@DataFile,
''''') + ', MOVE ' +
'        QUOTENAME(@Log, '''') + ' TO ' + QUOTENAME(@LogFile, '''') + ', FILE =
'' + CONVERT(varchar, @File)
'    EXEC (@query)
'End
'GO
'
'

