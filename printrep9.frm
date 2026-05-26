VERSION 5.00
Begin VB.Form Ucr9print 
   Caption         =   "Εκτύπωση"
   ClientHeight    =   570
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2652
   End
End
Attribute VB_Name = "Ucr9print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Sub Command1_Click()
'
'Dim crxApp As CRAXDRT.Application
'Dim crxRpt As CRAXDRT.Report
'Set crxApp = New CRAXDRT.Application
'Set crxRpt = crxApp.OpenReport("C:\mercvb\reports\timsql.rpt")
'crxRpt.PrintOut False, NoCopies
'
'
'End Sub

Public Sub printingCR9(ByVal ektypoths As String, ByVal mfile As String)
'Dim mdsn As String
'Dim mdatabase As String
'Dim musr As String
'Dim mpwd As String
'
'
'Open "C:\mercvb\reportconn.TXT" For Input As #1
'    Line Input #1, mdsn
'    Line Input #1, mdatabase
'    Line Input #1, musr
'    Line Input #1, mpwd
'
'Close #1

On Error Resume Next


Dim crxApp As CRAXdDRT.Application




Dim crxRpt As CRAXdDRT.Report
Set crxApp = New CRAXdDRT.Application
If Len(Dir(mfile)) = 0 Or Len(mfile) = 0 Then
    MsgBox "501.Δεν υπάρχει το αρχείο " + mfile
    Exit Sub
End If

Set crxRpt = crxApp.OpenReport(mfile, 1)

'Set crxRpt = crxApp.OpenReport("C:\mercvb\reports\timsql.rpt")
If Left(ektypoths, 2) = ".." Then
Else
   crxRpt.SelectPrinter "", ektypoths, ""
End If

crxRpt.PrintOut False, 1 ' NoCopies
Exit Sub



''Declare a Connection Info Object
'Dim ConnectionInfo As CRAXdDRT.ConnectionProperties
''Set the Connection Info to Connection Properties of ‘the table object
'Set ConnectionInfo = crxRpt.Database.Tables(1).ConnectionProperties
''Set the ODBC DSN
'ConnectionInfo.Item("DSN") = mdsn '"MERCHP"
''Set the database name
'ConnectionInfo.Item("Database") = mdatabase ' "MERCURY"
''Set the user name
'ConnectionInfo.Item("User ID") = musr '"sa"
''Set the password
'ConnectionInfo.Item("Password") = mpwd '"sa"
''Set the fully qualified table name if different from ‘the
''original data source
''Report.Database.Tables(1).location = “database.Owner.tablename”
'crxRpt.PrintOut False, 1 'NoCopies
'
'

''Dim crxApp As CRAXDRT.Application
''Dim crxRpt As CRAXDRT.Report
'crxApp.LogOnServer "p2sodbc.DLL", mLogonServer, mLogonDB, mlogonUserID, mpwd
'Set crxApp = New CRAXDRT.Application
'Set crxRpt = crxApp.OpenReport(mfile)
'crxRpt.ReadRecords
''crxRpt.PrinterName = "Bullzip PDF Printer"
'crxRpt.PrintOut False, NoCopies
'
'crxApp.LogOnServer "P2SSQL.DLL", "HPSRV\SQLEXPRESS", "mercury", "sa", "sa"
'crxRpt.Database = "mercury"
'Dim datcmd1 As ADODB.Command
'    ' Create a new instance of an ADO command object
'    Set datcmd1 = New ADODB.Command
'    Set datcmd1.ActiveConnection = Gdb
'    datcmd1.CommandText = "EGGTIMREP"
'    datcmd1.CommandType = adCmdText
'
'crxRpt.Database.AddADOCommand Gdb, datcmd1
'
''crxRpt.DriverName = "mercury"
''.ReadRecords
'crxApp.LogOnServer "P2SSQL.DLL", "MERCHP", "mercury", "sa", "sa"
End Sub

