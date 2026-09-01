Attribute VB_Name = "kentriko"

Option Explicit


'Global Const crptToWindow = 4
Global gf_Bearer As String
Global gfURL1 As String

Global g_Logot As Integer ' logotherapia

Global Gdb      As New ADODB.Connection

Global gDir     As String    'directoy με αρχεία d:\lageuro

Global gConnect As String    'dbase IV;  Access;

'Γενικές
Global g_Fpa(1 To 10)    ' κατηγορία ΦΠΑ ΓΙΑ ΤΗΝ ΚΑΘΕ ΜΙΑ ΚΑΤΗΓΟΡΙΑ Π.Χ. 2->18%

Global gxron_enar, gmhn_enar, gEnarjh, gLhjh, gWorkDay As Date    'gLHJH1=GLHJH+1 DAY

Global ghme_orist   ' , gsmatim
Attribute ghme_orist.VB_VarUserMemId = 1073741834

Global gpel5
Attribute gpel5.VB_VarUserMemId = 1073741835

Global gXEIRISTHS, gUserId    'ΔΙΑΚΑΙΟΜΑ,ΑΡΙΘΜΟΣ ΧΡΗΣΤΗ
Attribute gXEIRISTHS.VB_VarUserMemId = 1073741836
Attribute gUserId.VB_VarUserMemId = 1073741836
Global gParam(600, 2)
Global gshmer   'ghme_tel,
Attribute gshmer.VB_VarUserMemId = 1073741838

Global gPicture
Attribute gPicture.VB_VarUserMemId = 1073741839

Global gApoMenu  As Boolean    ' ερχομαι απο μενου

Global g_Stop    As Integer    ' 0=adiaforo 1=se loop  2=bges apo loop να μπορω να σταματώ τα μακρινά loop

Global gCapitals As Integer

'************************Τιμολόγησης
Global gBuff
Attribute gBuff.VB_VarUserMemId = 1073741842
Global gAsterisk As Integer '1= * οπως %  0=* σαν απλος χαρακτηρας

Global gisOnlineMydata As Integer


Global geidos_paras As Integer    ' ποιο παραστατικό επέλεξα

'*************************Λογιστικής
Global gdirlog
Attribute gdirlog.VB_VarUserMemId = 1073741844

Global gupdating_kin
Attribute gupdating_kin.VB_VarUserMemId = 1073741845

Global gDIG_3, gDIG_4   ', gok_gefyr  ...
Attribute gDIG_3.VB_VarUserMemId = 1073741846
Attribute gDIG_4.VB_VarUserMemId = 1073741846

'*****************************Αποθήκης
Global gUEOR_APOU
Attribute gUEOR_APOU.VB_VarUserMemId = 1073741848

Global gvar
Attribute gvar.VB_VarUserMemId = 1073741849


Global gFormaPel  As String



Global Const VK_ENTER = &H13

Global Const VK_TAB = &H9

Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As Long, _
                              ByVal wMsg As Long, _
                              ByVal wParam As Long, _
                              lParam As Any) As Long

Global Const CB_SHOWDROPDOWN = 335




Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long


'         Call SendMessage(Combo1.hWnd, CB_SHOWDROPDOWN, 1, 0)

'Here is how I use ShellExecute to call the default browser and open a specific web page. In a module declare the API,
Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

'Private Declare Sub keybd_event Lib "user32" (ByVal bScan As Byte
', ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Declare Sub keybd_event _
        Lib "user32" (ByVal bVk As Byte, _
                      ByVal bScan As Byte, _
                      ByVal dwFlags As Long, _
                      ByVal dwExtraInfo As Long)

' χρειάζεται για τις μακροεντολές δες executeLine(),mac()
'Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Foo1 As Long, ByVal Foo2 As Long, ByVal fcheckonly As Long) As Long

'32 BIT
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function PEPrintReport _
        Lib "CRPE.DLL" (ByVal reportFilePath As String, _
                        ByVal toPrinter As Integer, _
                        ByVal toWindow As Integer, _
                        ByVal Title As String, _
                        ByVal Left As Integer, _
                        ByVal Top As Integer, _
                        ByVal Width As Integer, _
                        ByVal Height As Integer, _
                        ByVal Style As Long, _
                        ByVal parentWindow As Integer) As Integer
Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowsDirectory _
               Lib "kernel32" _
               Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
                                             ByVal nSize As Long) As Long

'
'τοπικές μεταβλητές για τα load_form() και ektyp_form()
'formes   load_form()      ektyp_form()
Global gm_str(250) As String

Global gm_f(250)   As String

Global gpic(250)   As String

Global gm_r(250)   As Integer

Global gm_c(250)   As Integer

'
Public Sub PutGGPS(ByVal u As String)
  Dim ANS As Integer: ANS = vbNo
      Dim p, AFM As String
      
      
      If Left$(u, 9) = "LAGAKIS23" Or Len(Trim(u)) = 0 Then
           ANS = MsgBox("Ζητήστε τους κωδικούς ευρεσης μεσω ΑΦΜ απο τον λογιστή σας" + Chr(13) + "Θα λήξει σύντομα ο βοηθητικος κωδικός που εχετε" + Chr(13) + "Θέλετε να τους δωσετε τώρα;", vbYesNo)
      End If
 
      If ANS = vbYes Then
         u = InputBox("user (χρήστης)", "Χρήστης")
         p = InputBox("ΚΩΔΙΚΌΣ(PASSWORD)", "ΚΩΔΙΚΟΣ")
         AFM = InputBox("ΑΦΜ ΕΠΙΧΕΙΡΗΣΗΣ", "TO ΑΦΜ ΣΑΣ")
         If Len(Trim(u)) > 1 Then
           Gdb.Execute "update PARAMETROI SET TIMH='" + u + "' where FORMA='PELAT1' and  VAR='f_EYRESHAFM_USER'"
           Gdb.Execute "update PARAMETROI SET TIMH='" + p + "' where FORMA='PELAT1' and  VAR='f_EYRESHAFM_PWD'"
           Gdb.Execute "update PARAMETROI SET TIMH='" + AFM + "' where FORMA='PELAT1' and  VAR='f_EYRESHAFM_AFM'"
         End If
     
 
      End If
      
End Sub
      

Public Sub Update_Structures()


        '<EhHeader>
        On Error GoTo Update_Structures_Err

        '</EhHeader>
        Dim dum2

        'Dim R As New ADODB.Recordset
        Dim k

        Dim sql As String

        '       On Error Resume Next

        Dim R   As New ADODB.Recordset

        R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'SN'", Gdb, adOpenDynamic, adLockOptimistic
        On Error Resume Next

        If R(0) = 0 Then
            Gdb.Execute "CREATE TABLE SN (" & " [ARIU] [nvarchar](18) NULL, " & " [HME] [smalldatetime] NULL," & "[KODPEL] [nvarchar](5) NULL," & " [ATIM] [nvarchar](7) NULL, " & " [ID] [int] IDENTITY(1,1) NOT NULL, " & " [KOD] [char](20) NULL )"
        End If

        R.Close

        R.Open "SELECT COUNT(*) AS N  FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME='XRHSEIS';", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

        If R(0) = 0 Then
            sql = "CREATE TABLE XRHSEIS([ENARXH] [datetime] NULL,[LHXH] [datetime] NULL,[ETOS] [int] NOT NULL,CONSTRAINT [PK_XRHSEIS] PRIMARY KEY CLUSTERED([ETOS] Asc) ON [PRIMARY]) ON [PRIMARY]"
   
            Gdb.Execute sql, k  ' "CREATE TABLE XRHSEIS ( ENARXH DATETIME, LHXH DATETIME,ETOS INT )", K
        End If

        R.Close
        '5



        R.Open "SELECT COUNT(*) AS N  FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME='TRAINPERIODS';", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

        If R(0) = 0 Then
            sql = "CREATE TABLE TRAINPERIODS(KOD VARCHAR(12),[ENARXH] [date] NULL,[LHXH] [date] NULL,MHNAS VARCHAR(20),[SYNEDRIES] [int],ID INT IDENTITY(1,1) ) "
   
            Gdb.Execute sql, k  ' "CREATE TABLE XRHSEIS ( ENARXH DATETIME, LHXH DATETIME,ETOS INT )", K
        End If

        R.Close


        dum2 = ADD_FIELD("TRAINPERIODS", "CENARXH", "VARCHAR(10)")
        dum2 = ADD_FIELD("TRAINPERIODS", "CLHXH", "VARCHAR(10)")



 R.Open "SELECT COUNT(*) AS N  FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME='PARTIDES';", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

        If R(0) = 0 Then
            sql = "CREATE TABLE PARTIDES(KOD VARCHAR(16),[LOTNUMBER] [varchar](16),[HMEL] [datetime],[POSOAG] [numeric](10, 2),[POSOPOL] [numeric](10, 2) NULL,ID INT IDENTITY(1,1) , "
            sql = sql + "[ATIM] [varchar](8),[HME] [datetime],[FLAG] [int],[YPOLOIPO] [numeric](10, 3),SEIRA INT,[ID_NUM] [int],[EIDOSKPE] [varchar](12) )"
            Gdb.Execute sql, k  ' "CREATE TABLE XRHSEIS ( ENARXH DATETIME, LHXH DATETIME,ETOS INT )", K
        End If

        R.Close



'CREATE TABLE [dbo].[PARTIDES](
'    [KOD] [varchar](16) NULL,
'    [LOTNUMBER] [varchar](16) NULL,
'    [HMEL] [datetime] NULL,
'    [POSOAG] [numeric](10, 2) NULL,
'    [POSOPOL] [numeric](10, 2) NULL,


'    [ATIM] [varchar](8) NULL,
'    [HME] [datetime] NULL,
'    [FLAG] [int] NULL,
'    [ID] [int] IDENTITY(1,1) NOT NULL,
'    [YPOLOIPO] [numeric](10, 3) NULL,
'    [ID_NUM] [int] NULL,
'    [EIDOSKPE] [varchar](12) NULL
') ON [PRIMARY]
'
'











        R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'ETAIREIES'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

        If R(0) = 0 Then
            Gdb.Execute "CREATE TABLE ETAIREIES ( DSN  char(60),TITLOS  CHAR(30) )"
        End If

        R.Close

100     dum2 = ADD_FIELD("XRHSEIS", "ID", "INT IDENTITY(1,1)")
110     dum2 = ADD_FIELD("PEL", "ID", "INT IDENTITY(1,1)")
120     dum2 = ADD_FIELD("EID", "ID", "INT IDENTITY(1,1)")

        dum2 = ADD_FIELD("EID", "ENERGO", "INT DEFAULT (1)")
        
        If dum2 = 1 Then
            Gdb.Execute "update EID SET ENERGO=1 "
        End If
        
130     dum2 = ADD_FIELD("GRA", "ID", "INT IDENTITY(1,1)")
140     dum2 = ADD_FIELD("MEM", "ID", "INT IDENTITY(1,1)")
150     dum2 = ADD_FIELD("SN", "ID", "INT IDENTITY(1,1)")

152     dum2 = ADD_FIELD("PINAKES", "ID", "INT IDENTITY(1,1)")
        
        dum2 = ADD_FIELD("PINAKES", "LOG1", "CHAR(20)")
        dum2 = ADD_FIELD("PINAKES", "N1", "INT")
        dum2 = ADD_FIELD("PINAKES", "C1", "CHAR(20)")
160     dum2 = ADD_FIELD("PARASTAT", "ID", "INT IDENTITY(1,1)")
        dum2 = ADD_FIELD("PARASTAT", "DEFSKOPOS", "INT DEFAULT(1)")
        
        dum2 = ADD_FIELD("PINAKES", "N2", "INT")
        dum2 = ADD_FIELD("PINAKES", "C2", "CHAR(20)")
        dum2 = ADD_FIELD("PINAKES", "N3", "INT")
        dum2 = ADD_FIELD("PINAKES", "C3", "CHAR(20)")
        dum2 = ADD_FIELD("PINAKES", "N4", "INT")
        dum2 = ADD_FIELD("PINAKES", "C4", "CHAR(20)")



dum2 = ADD_FIELD("PEL", "CONTRACTINGAUTHORITYID", "NVARCHAR(30)")
        'contractIdentifier
        dum2 = ADD_FIELD("PEL", "CONTRACTIDENTIFIER", "NVARCHAR(30)")

        dum2 = ADD_FIELD("PEL", UCase("organizationalUnitName"), "NVARCHAR(30)")
        'contractIdentifier
        dum2 = ADD_FIELD("PEL", UCase("organizationalUnitCode"), "NVARCHAR(30)")

        dum2 = ADD_FIELD("PEL", UCase("budtype"), "NVARCHAR(30)")
        'contractIdentifier
        dum2 = ADD_FIELD("PEL", UCase("budidentifier"), "NVARCHAR(30)")





170     dum2 = ADD_FIELD("PARASTAT", "YP1", "INT")
180     dum2 = ADD_FIELD("PARASTAT", "YP2", "INT")
190     dum2 = ADD_FIELD("PARASTAT", "YP3", "INT")
200     dum2 = ADD_FIELD("PARASTAT", "YP4", "INT")
210     dum2 = ADD_FIELD("PARASTAT", "YP5", "INT")
220     dum2 = ADD_FIELD("PARASTAT", "C1", "CHAR(20)")
230     dum2 = ADD_FIELD("PARASTAT", "C2", "CHAR(20)")
240     dum2 = ADD_FIELD("PARASTAT", "LAJ1", "CHAR(30)")
250     dum2 = ADD_FIELD("PARASTAT", "LAJ2", "CHAR(30)")
260     dum2 = ADD_FIELD("PARASTAT", "LAJ3", "CHAR(30)")
270     dum2 = ADD_FIELD("PARASTAT", "LAJ4", "CHAR(30)")
280     dum2 = ADD_FIELD("PARASTAT", "LAJ5", "CHAR(30)")
290     dum2 = ADD_FIELD("PARASTAT", "LAJ6", "CHAR(30)")
300     dum2 = ADD_FIELD("PARASTAT", "LAJ7", "CHAR(30)")
'proepilFPA
 dum2 = ADD_FIELD("PARASTAT", "PROEPILFPA", "INT")

dum2 = ADD_FIELD("PARASTAT", "LPARAKRAT", "CHAR(30)")

dum2 = ADD_FIELD("PARASTAT", "MYEID", "nvarCHAR(10)")
dum2 = ADD_FIELD("PARASTAT", "MYTYPESOD", "nvarCHAR(20)")

dum2 = ADD_FIELD("PARASTAT", "MYCATEG", "nvarCHAR(20)")
dum2 = ADD_FIELD("PARASTAT", "MYISAYTOTIM", "INT")

dum2 = ADD_FIELD("PARASTAT", "ISMYDATA", "INT")
dum2 = ADD_FIELD("PARASTAT", "ISDIAKIN", "INT")


dum2 = ADD_FIELD("PARASTAT", "MYTYPEXOD", "nvarCHAR(20)")


dum2 = ADD_FIELD("PARASTAT", "MYCATEXOD", "nvarCHAR(20)")
dum2 = ADD_FIELD("PARASTAT", "DEFAULTAPALL", "INT")


dum2 = ADD_FIELD("PARASTAT", "B2G", "INT")

        dum2 = ADD_FIELD("PARASTAT", "UBL", "VARCHAR(30)")
        
        
dum2 = ADD_FIELD("PARASTAT", "ISANASTROFODA", "int DEFAULT (0)")
dum2 = ADD_FIELD("PARASTAT", "DEFANASTROFO", "int DEFAULT (0)")
        
        
        
        dum2 = ADD_FIELD("MEM", "PELARIT", "NVARCHAR(80)")
        dum2 = ADD_FIELD("MEM", "PELPOL", "NVARCHAR(80)")
        dum2 = ADD_FIELD("MEM", "PELTK", "NVARCHAR(80)")


dum2 = ADD_FIELD("MEM", "MYDATAUSER", "NVARCHAR(80)")
dum2 = ADD_FIELD("MEM", "MYDATAPASW", "NVARCHAR(80)")



dum2 = ADD_FIELD("MEM", "MYDATAISDEMO", "int CONSTRAINT AA DEFAULT 0 WITH VALUES")
dum2 = ADD_FIELD("MEM", "MYDATAFROMEID1", "int CONSTRAINT BB DEFAULT 0 WITH VALUES")


        dum2 = ADD_FIELD("MEM", "PAROXUSER", "NVARCHAR(80)")
        dum2 = ADD_FIELD("MEM", "PAROXPASS", "NVARCHAR(80)")
        dum2 = ADD_FIELD("MEM", "PAROXKEY", "NVARCHAR(80)")

dum2 = ADD_FIELD("PARASTAT", "FORMA1", "CHAR(80)")
dum2 = ADD_FIELD("PARASTAT", "FORMA2", "CHAR(80)")
dum2 = ADD_FIELD("PARASTAT", "FORMA3", "CHAR(80)")

dum2 = ADD_FIELD("PARASTAT", "AFMGRANDPEL", "CHAR(15)")



310     dum2 = ADD_FIELD("PARASTAT", "LFPA1", "CHAR(30)")
320     dum2 = ADD_FIELD("PARASTAT", "LFPA2", "CHAR(30)")
330     dum2 = ADD_FIELD("PARASTAT", "LFPA3", "CHAR(30)")
340     dum2 = ADD_FIELD("PARASTAT", "LFPA4", "CHAR(30)")
350     dum2 = ADD_FIELD("PARASTAT", "LFPA5", "CHAR(30)")
360     dum2 = ADD_FIELD("PARASTAT", "LFPA6", "CHAR(30)")
370     dum2 = ADD_FIELD("PARASTAT", "LFPA7", "CHAR(30)")
380     dum2 = ADD_FIELD("PARASTAT", "LFPA16", "CHAR(30)")
390     dum2 = ADD_FIELD("PARASTAT", "LFPA17", "CHAR(30)")
400     dum2 = ADD_FIELD("PARASTAT", "XREPEL", "CHAR(30)")
410     dum2 = ADD_FIELD("PARASTAT", "PISPEL", "CHAR(30)")
420     dum2 = ADD_FIELD("PARASTAT", "AKYR", "CHAR(1)")
430     dum2 = ADD_FIELD("PARASTAT", "AYTOK", "INT")
432     dum2 = ADD_FIELD("PARASTAT", "SET_EIDON", "INT")
434     dum2 = ADD_FIELD("PARASTAT", "POLITES", "CHAR(30)")
436     dum2 = ADD_FIELD("PARASTAT", "PARAKRATISI", "NUMERIC(10,5)")
438     dum2 = ADD_FIELD("PARASTAT", "MYF", "INT")
439     dum2 = ADD_FIELD("PARASTAT", "TAMEIAKI", "CHAR(30)")
        
        Gdb.Execute "UPDATE PARASTAT SET PARAKRATISI=0 WHERE PARAKRATISI IS NULL"

440     If dum2 = 1 Then
            
450         Gdb.Execute "UPDATE PARASTAT SET AKYR='κ'"
460         MsgBox "ελέγξτε τα ακυρωτικά γιατί σαν προεπιλεγμένο έχει ορισθεί το κ"

        End If

470     dum2 = ADD_FIELD("EID", "MESTIMPOL", "DECIMAL")

480     dum2 = ADD_FIELD("EGGTIM", "FCURRENCY", "DECIMAL")
        dum2 = ADD_FIELD("EGGTIM", "EKPT2", "DECIMAL")
        dum2 = ADD_FIELD("EGGTIM", "ID_NUM", "INT")
        
        dum2 = ADD_FIELD("EGGTIM", "ID", "INT IDENTITY(1,1)")
        
        dum2 = ADD_FIELD("EGGTIM", "KOLA", "DECIMAL")
        dum2 = ADD_FIELD("EGGTIM", "MIKTAKILA", "DECIMAL")
        
        dum2 = ADD_FIELD("EGGTIM", "XVRA", "CHAR(20)")
         
        dum2 = ADD_FIELD("EGGTIM", "LITRA", "NUMERIC(6,3)")
        dum2 = ADD_FIELD("EGGTIM", "EFK", "NUMERIC(6,3)")
        
        dum2 = ADD_FIELD("EGGTIM", "AJAGOR", "NUMERIC(16,3)")
        dum2 = ADD_FIELD("EGGTIM", "AJPOL", "NUMERIC(16,3)")
        
        'OTHERMEASUREMENTUNITQUANTITY , OTHERMEASUREMENTUNITTITLE
        
        dum2 = ADD_FIELD("EGGTIM", "OTHERMEASUREMENTUNITQUANTITY", "INT")
        dum2 = ADD_FIELD("EGGTIM", "OTHERMEASUREMENTUNITTITLE", "NVARCHAR(30)")
        
        dum2 = ADD_FIELD("EGGTIM", "DISCOUNTOPTION", "INT")
      '  classificationCategory
        dum2 = ADD_FIELD("EGGTIM", "CLASSIFICATIONCATEGORY", "NVARCHAR(15)")
        dum2 = ADD_FIELD("EGGTIM", "CLASSIFICATIONTYPE", "NVARCHAR(15)")
        ' KOLA , MIKTAKILA
        
        
        ' dum2 = ADD_FIELD("EGGTIM", "ID_NUM", "INT IDENTITY(1,1)")

490     dum2 = ADD_FIELD("EID", "EPIUYP", "INT")
500     dum2 = ADD_FIELD("EID", "SYSKMAX", "INT")
510     dum2 = ADD_FIELD("EID", "SYSKMIN", "INT")
520     dum2 = ADD_FIELD("EID", "NUM1", "DECIMAL")
530     dum2 = ADD_FIELD("EID", "NUM2", "DECIMAL")
540     dum2 = ADD_FIELD("EID", "NUM3", "DECIMAL")
550     dum2 = ADD_FIELD("EID", "CH1", "CHAR(20)")
560     dum2 = ADD_FIELD("EID", "CH2", "CHAR(20)")
570     dum2 = ADD_FIELD("EID", "CH3", "CHAR(20)")
580     dum2 = ADD_FIELD("EID", "CH4", "CHAR(20)")
590     dum2 = ADD_FIELD("EID", "CH5", "CHAR(20)")
600     dum2 = ADD_FIELD("EID", "CH6", "CHAR(20)")
610     dum2 = ADD_FIELD("EID", "HM1", "DATETIME")
620     dum2 = ADD_FIELD("EID", "HM2", "DATETIME")
630     dum2 = ADD_FIELD("EID", "HM3", "DATETIME")
640     dum2 = ADD_FIELD("EID", "LASTUPD", "DATETIME")
642     dum2 = ADD_FIELD("EID", "PROMPOL", "NUMERIC(6,2)")

        dum2 = ADD_FIELD("TIM", "APALAGIFPA", "INT DEFAULT (0)")

        dum2 = ADD_FIELD("EID", "EFK", "NUMERIC(6,3)")

        dum2 = ADD_FIELD("EID", "APALLFPA", "INT")


    dum2 = ADD_FIELD("EID", "CPV", "NVARCHAR(30)")

    dum2 = ADD_FIELD("EID", "DISCOUNTOPTION", "INT")

        dum2 = ADD_FIELD("TIMOKAT", "TIMOKPEL", "VARCHAR(5)")
        dum2 = ADD_FIELD("TIMOKAT", "TIMOKEID", "VARCHAR(5)")
        dum2 = ADD_FIELD("TIMOKAT", "TIMH", "DECIMAL")
        dum2 = ADD_FIELD("TIMOKAT", "TIMOKID", "INT IDENTITY(1,1)")

650     dum2 = ADD_FIELD("PEL", "NUM1", "DECIMAL")
660     dum2 = ADD_FIELD("PEL", "NUM2", "DECIMAL")
670     dum2 = ADD_FIELD("PEL", "NUM3", "DECIMAL")
680     dum2 = ADD_FIELD("PEL", "NUM4", "DECIMAL")
682     dum2 = ADD_FIELD("PEL", "NUM5", "DECIMAL")
684     dum2 = ADD_FIELD("PEL", "NUM6", "DECIMAL")

        
         dum2 = ADD_FIELD("PEL", "HM7", "DATETIME")
         dum2 = ADD_FIELD("PEL", "HM8", "DATETIME")
         dum2 = ADD_FIELD("PEL", "HM9", "DATETIME")
         dum2 = ADD_FIELD("PEL", "HM10", "DATETIME")
         dum2 = ADD_FIELD("PEL", "HM11", "DATETIME")

         dum2 = ADD_FIELD("PEL", "CH7", "VARCHAR(50)")

        dum2 = ADD_FIELD("PEL", "COMB1", "VARCHAR(50)")
        dum2 = ADD_FIELD("PEL", "COMB2", "VARCHAR(50)")
        dum2 = ADD_FIELD("PEL", "COMB3", "VARCHAR(50)")
        dum2 = ADD_FIELD("PEL", "COMB4", "VARCHAR(50)")
        dum2 = ADD_FIELD("PEL", "COMB5", "VARCHAR(50)")
        dum2 = ADD_FIELD("PEL", "COUNTRY", "CHAR(2) CONSTRAINT PELCOUNTRY DEFAULT 'GR' WITH VALUES")

        
        dum2 = ADD_FIELD("PEL", "ENERGOS", "INT DEFAULT (1)")
        
        
        dum2 = ADD_FIELD("PEL", "ARKYKL", "NVARCHAR(18)")
        dum2 = ADD_FIELD("PEL", "VCAT", "NVARCHAR(18)")
        dum2 = ADD_FIELD("PEL", "VFACT", "NVARCHAR(18)")
        dum2 = ADD_FIELD("PEL", "VID", "NVARCHAR(18)")
        
         dum2 = ADD_FIELD("PEL", "STREETNUMBER", "NVARCHAR(4)")
         dum2 = ADD_FIELD("PEL", "BRANCH", "NVARCHAR(2) DEFAULT '0' ")
        
        
        
        
        If dum2 = 1 Then
            Gdb.Execute "update PEL SET ENERGOS=1 "
            Dim fff As String
            fff = ""
        End If

        dum2 = ADD_FIELD("PEL", "ADT", "VARCHAR(20)")

690     dum2 = ADD_FIELD("PEL", "EMAIL", "CHAR(35)")
696     dum2 = ADD_FIELD("PEL", "KINHTO", "CHAR(35)")
700     dum2 = ADD_FIELD("PEL", "PVLHTHS", "CHAR(20)")
710     dum2 = ADD_FIELD("XREOPIS", "ARITMISI", "INT")
720     dum2 = ADD_FIELD("XREOPIS", "EPIT", "INT")
730     dum2 = ADD_FIELD("XREOPIS", "AA", "INT")
740     dum2 = ADD_FIELD("XREOPIS", "TAYTPEL", "CHAR(30)")
750     dum2 = ADD_FIELD("XREOPIS", "TAYTXPI", "CHAR(5)")
760     dum2 = ADD_FIELD("XREOPIS", "TAMEIAKH", "BIT")
770     dum2 = ADD_FIELD("XREOPIS", "KODLOG", "CHAR(30)")
        dum2 = ADD_FIELD("XREOPIS", "FORMA1", "NVARCHAR(50)")
        dum2 = ADD_FIELD("XREOPIS", "FORMA2", "NVARCHAR(50)")
        
        dum2 = ADD_FIELD("XREOPIS", "C1", "NVARCHAR(50)")
        dum2 = ADD_FIELD("XREOPIS", "C2", "NVARCHAR(50)")
        dum2 = ADD_FIELD("XREOPIS", "N1", "NUMERIC(8,3)")
        dum2 = ADD_FIELD("XREOPIS", "N2", "NUMERIC(8,3)")


        ' On Error GoTo Update_Structures_Err
780     dum2 = ADD_FIELD("EGG", "USERID", "INT")
790     dum2 = ADD_FIELD("EGG", "ID", "INT IDENTITY(1,1)")
800     dum2 = ADD_FIELD("EGG", "IDGRA", "INT")
810     dum2 = ADD_FIELD("EGG", "IDTIM", "INT")
820     dum2 = ADD_FIELD("EGG", "AAXREOPIS", "INT")
830     dum2 = ADD_FIELD("EGG", "ATIM2", "CHAR(15)")
'EIDXPI
        dum2 = ADD_FIELD("EGG", "USERID", "INT")

         dum2 = ADD_FIELD("EGG", "EIDXPI", "INT")
         dum2 = ADD_FIELD("EGG", "B_C1", "CHAR(5)")


840     dum2 = ADD_FIELD("PARAMETROI", "ID", "INT IDENTITY(1,1)")

         dum2 = ADD_FIELD("PARAMETROI", "KATEG", "INT")

850     dum2 = ADD_FIELD("GRA", "ID", "INT IDENTITY(1,1)")
        dum2 = ADD_FIELD("GRA", "HME_METAB", "DATETIME")
        
        
860     dum2 = ADD_FIELD("TIM", "KERDOS", "NUMERIC(12,3)")
        dum2 = ADD_FIELD("TIM", "PARAKRATISI", "NUMERIC(12,3)")
        dum2 = ADD_FIELD("TIM", "ID_NUM", "INT IDENTITY(1,1)")

870     dum2 = ADD_FIELD("TIM", "KLEIDI", "CHAR(15)")

        dum2 = ADD_FIELD("TIM", "LITRA", "NUMERIC(8,3)")
        dum2 = ADD_FIELD("TIM", "EFK", "NUMERIC(8,3)")
        dum2 = ADD_FIELD("TIM", "ORA", "CHAR(8)")
        dum2 = ADD_FIELD("TIM", "IDG", "uniqueidentifier NOT NULL DEFAULT NEWID()")
        
         dum2 = ADD_FIELD("TIM", "HMEPARAD", "DATETIME")
         dum2 = ADD_FIELD("TIM", "ORAPARAD", "CHAR(6)")
          dum2 = ADD_FIELD("TIM", "MARKPLHR", "VARCHAR(16)")
        ' MARKPLHR
         
         dum2 = ADD_FIELD("TIM", "ISANAPODO", "INT")
         dum2 = ADD_FIELD("TIM", "ANAPODOAITIA", "INT")
         'AFMTRITOU
          dum2 = ADD_FIELD("TIM", "AFMTRITOU", "VARCHAR(15)")
        
dum2 = ADD_FIELD("TIM", "ENTITYUID", "nvarchar(40)")
dum2 = ADD_FIELD("TIM", "QRURL", "nvarchar(200)")
dum2 = ADD_FIELD("TIM", "ENTITYMARK", "nvarchar(40)")
dum2 = ADD_FIELD("TIM", "ENTITY", "int")
dum2 = ADD_FIELD("TIM", "AADEKAU", "float")
dum2 = ADD_FIELD("TIM", "AADEFPA", "float")
dum2 = ADD_FIELD("TIM", "ENTLINEN", "int")
dum2 = ADD_FIELD("TIM", "INCMARK", "nvarchar(43)")
dum2 = ADD_FIELD("TIM", "APALAGIFPA", "INT DEFAULT (0)")
dum2 = ADD_FIELD("TIM", "SXETMARK", "NVARCHAR(80)")

dum2 = ADD_FIELD("TIM", "SXETMARKS", "NTEXT")

dum2 = ADD_FIELD("TIM", "ELINE", "NVARCHAR(120)")
        
dum2 = ADD_FIELD("TIM", "ANASTOLHFPA", "INT DEFAULT (0)")
dum2 = ADD_FIELD("TIM", "TYPOMENO", "INT DEFAULT (0)")
 dum2 = ADD_FIELD("TIM", "AKYROMENO", "INT DEFAULT (0)")
        
dum2 = ADD_FIELD("TIM", "C1", "nvarchar(120)")
dum2 = ADD_FIELD("TIM", "C2", "nvarchar(45)")
dum2 = ADD_FIELD("TIM", "C3", "nvarchar(45)")
dum2 = ADD_FIELD("TIM", "C4", "nvarchar(45)")

dum2 = ADD_FIELD("TIM", "C5", "nvarchar(45)")
dum2 = ADD_FIELD("TIM", "C6", "nvarchar(45)")
dum2 = ADD_FIELD("TIM", "C7", "nvarchar(45)")
'toWeigh

dum2 = ADD_FIELD("TIM", "TOWEIGH", "INT DEFAULT (0)")

dum2 = ADD_FIELD("TIM", "OTHERMOVEPURPOSETITLE", "nvarchar(45)")

dum2 = ADD_FIELD("TIM", "C12", "nvarchar(45)")
dum2 = ADD_FIELD("TIM", "C13", "nvarchar(45)")
dum2 = ADD_FIELD("TIM", "C14", "nvarchar(45)")

dum2 = ADD_FIELD("TIM", "KR1", "NUMERIC(12,3)")
dum2 = ADD_FIELD("TIM", "KR2", "NUMERIC(12,3)")


dum2 = ADD_FIELD("TIM", "NUM1", "NUMERIC(12,3)")
dum2 = ADD_FIELD("TIM", "NUM2", "NUMERIC(12,3)")
dum2 = ADD_FIELD("TIM", "NUM3", "NUMERIC(12,3)")

dum2 = ADD_FIELD("TIM", "NUM11", "NUMERIC(12,3)")
        
 dum2 = ADD_FIELD("TIM", "XMLTEXT", "NTEXT")
        
        
 dum2 = ADD_FIELD("TIM", "EXCHANGERATE", "NUMERIC(12, 5)")
 dum2 = ADD_FIELD("TIM", "CURRENCY", "nvarchar(3)")
 ',CURRENCY
 dum2 = ADD_FIELD("TIM", "JWT", "nvarchar(800)")
 dum2 = ADD_FIELD("TIM", "AUTHENTICATIONCODE", "NVARCHAR(42)") 'authenticationCode
       
  dum2 = ADD_FIELD("TIM", "STREET", "NVARCHAR(80)")
  dum2 = ADD_FIELD("TIM", "STREETNUMBER", "NVARCHAR(4)")
  dum2 = ADD_FIELD("TIM", "POSTALCODE", "NVARCHAR(6)")
  dum2 = ADD_FIELD("TIM", "CITY", "NVARCHAR(30)")
   dum2 = ADD_FIELD("TIM", "BRANCH", "NVARCHAR(2)")
        
        
        
        
        ' KOLA , MIKTAKILA
'------------  AN GIA KAPOIO LOGO ΔΕΝ ΕΙΝΑΙ IDENTITY TO ID_NUM TO ΣΒΗΝΩ ΚΑΙ ΤΟ ΞΑΝΑΔΗΜΙΟΥΡΓΩ
 
 
Dim r2 As New ADODB.Recordset
 r2.Open "SELECT OBJECTPROPERTY(object_id('TIM'), 'TableHasIdentity')", Gdb, adOpenForwardOnly, adLockReadOnly
' /*   SQL Server returns 1 if the table has an identity column, 0 if it doesn't.   */
  If r2(0) = 0 Then
     Gdb.Execute "ALTER TABLE TIM DROP COLUMN ID_NUM"
     dum2 = ADD_FIELD("TIM", "ID_NUM", "INT IDENTITY(1,1)")
     ' R2.Close
  End If
r2.Close


880     If dum2 = 1 Then
890         Gdb.Execute "UPDATE TIM SET KLEIDI=CAST(ID_NUM AS CHAR(10))"

            ' Gdb.Execute "ALTER TABLE TIM ALTER COLUMN  KLEIDI  CHAR(15) NOT NULL "

            ' Gdb.Execute "ALTER TABLE TIM ADD PRIMARY KEY (KLEIDI); "

        End If

        ' Gdb.Execute "UPDATE TIM SET KEY=CAST(ID_NUM AS CHAR(10))", dum2
        '  Gdb.Execute "ALTER TABLE TIM ADD PRIMARY KEY (KEY); ", dum2

        Dim ll As Long, LL2 As Long


900     R.Open "SELECT * FROM PARAMETROI", Gdb, adOpenForwardOnly, adLockReadOnly
910     ll = R.FIELDS("TIMH").DefinedSize
920     LL2 = R.FIELDS("SXOLIA").DefinedSize
930     R.Close

940     If ll < 12 Then
950         Gdb.Execute "ALTER TABLE PARAMETROI ALTER COLUMN TIMH CHAR(42)"
            '{ new_data_type [ ( precision [ , scale
            '     On Error GoTo Update_Structures_Err
            '    DUM2 = ADD_FIELD("PARAMETROI", "TIMH2", "CHAR(30)")
            '    Dim N
            '    Gdb.Execute "UPDATE PARASTAT SET TIMH2=TIMH", N
            '    Gdb.Execute "ALTER TABLE PARASTAT DROP COLUMN TIMH"
            '    DUM2 = ADD_FIELD("PARAMETROI", "TIMH", "CHAR(30)")
            '    Gdb.Execute "UPDATE PARASTAT SET TIMH=TIMH2", N
        End If

960     If LL2 < 80 Then
970         Gdb.Execute "ALTER TABLE PARAMETROI ALTER COLUMN SXOLIA CHAR(80)", ll
        End If

980     dum2 = ADD_FIELD("PARASTAT", "ID", "INT IDENTITY(1,1)")
990     dum2 = ADD_FIELD("PARASTAT", "YP1", "INT")
1000    dum2 = ADD_FIELD("PARASTAT", "YP2", "INT")
1010    dum2 = ADD_FIELD("PARASTAT", "YP3", "INT")
1020    dum2 = ADD_FIELD("PARASTAT", "YP4", "INT")
1030    dum2 = ADD_FIELD("PARASTAT", "YP5", "INT")
1040    dum2 = ADD_FIELD("PARASTAT", "C1", "CHAR(20)")
1050    dum2 = ADD_FIELD("PARASTAT", "C2", "CHAR(20)")

1060    R.Open "select  COUNT(*) from PARASTAT WHERE YP1=1 ", Gdb, adOpenDynamic, adLockOptimistic

1070    If R(0) = 0 Then
1080        Gdb.Execute "UPDATE PARASTAT SET YP1=1"
        End If

1090    R.Close

1100    R.Open "select  COUNT(*) from PARASTAT WHERE YP2=1 ", Gdb, adOpenDynamic, adLockOptimistic

1110    If R(0) = 0 Then
1120        Gdb.Execute "UPDATE PARASTAT SET YP2=1"
        End If

1130    R.Close









Dim SQ As String
 r2.Open "select COUNT(*) FROM sys.views where name = 'TIMOLOGIA'", Gdb, adOpenForwardOnly, adLockReadOnly
' /*   SQL Server returns 1 if the table has an identity column, 0 if it doesn't.   */
  If r2(0) = 0 Then
  
 SQ = " CREATE VIEW [TIMOLOGIA] AS SELECT         PEL.EIDOS,  PEL.EPO,  PEL.ONO,  PEL.DIE,  PEL.POL,  PEL.THL,  PEL.EPA,  PEL.AFM,  PEL.PEK,  PEL.AEG,  PEL.AYP,  PEL.TYP,  PEL.XRESYN,  PEL.KOD,  PEL.SHM1, "
 SQ = SQ + " PEL.SHM2,  PEL.KART,  PEL.XREMHN,  PEL.PISMHN,  PEL.XRE,  PEL.PIS,  PEL.PMXRE,  PEL.PMPIS,  PEL.LASTUPDT,  PEL.PISSYN,  PEL.ARTIM,  PEL.SUMTIM,  PEL.KODGAL,"
 SQ = SQ + " PEL.PLAISIO,  PEL.ARPARAG,  PEL.HMELHJ,  PEL.HME_LHJ,  PEL.TYPOS,  PEL.XRVMA,  PEL.DOY,  PEL.PLAFON,  PEL.HMERESAPOP,  PEL.MEMO,  PEL.HM1,  PEL.HM2,"
 SQ = SQ + " PEL.HM3,  PEL.HM4,  PEL.HM5,  PEL.HM6,  PEL.CH1,  PEL.CH2,  PEL.CH3,  PEL.CH4,  PEL.CH5,  PEL.CH6,  PEL.ENERGOS,  PEL.NUM1,  PEL.NUM2,  PEL.NUM3,"
 SQ = SQ + " PEL.NUM4,  PEL.NUM5,  PEL.NUM6,  PEL.NUM7,  PEL.NUM8,  PEL.PVLHTHS,  PEL.EMAIL,  PEL.ID,  PEL.KINHTO,  PEL.COMB1,  PEL.COMB2,  PEL.COMB3,  PEL.COMB4,"
 SQ = SQ + " PEL.COMB5,  PEL.ADT,  PEL.HM7,  PEL.HM8,  PEL.HM9,  PEL.HM10,  PEL.HM11,  PEL.CH7,  PEL.NUMER3,  PEL.NUMER4,  PEL.NUMER5,"
 SQ = SQ + " PEL.NUMER6,  PEL.NUMER7,  PEL.NUMER8,  PEL.COUNTRY,  PARASTAT.TITLOS,"
                         
 SQ = SQ + " TIM.HME AS TIMHME,  TIM.TRP,  TIM.ATIM AS TIMATIM,  TIM.AJI,  TIM.METAF,  TIM.EIDPAR,"
                         
                         
 SQ = SQ + " TIM.FPA1,  TIM.FPA2,  TIM.FPA3,  TIM.FPA4,  TIM.FPA6,  TIM.FPA7,  TIM.FPA8,  TIM.FPA9,  TIM.TYP AS TIMTYP,  TIM.AJ1,  TIM.AJ2,  TIM.AJ3,  TIM.AJ4,  TIM.AJ5,  TIM.AJ6,"
 SQ = SQ + " TIM.AJ8,  TIM.AJ7,  TIM.AJ9,  TIM.EKPT1,  TIM.EKPT2 AS TIMEKPT2,  TIM.EKPT3,  TIM.EKPT4,  TIM.EKPT5,  TIM.SKOPOS,  TIM.SXETIKO,  TIM.PARAT,  TIM.FORTOSH,  TIM.SYNPOS,"
 SQ = SQ + " TIM.SKOPOS2,  TIM.PROOR,  TIM.AYTOK,  TIM.KR1,  TIM.KR2,  TIM.PARAKRATISI,  TIM.ORA,  TIM.C1,  TIM.C2,  TIM.NUM1 AS TIMNUM1,  TIM.C3,  TIM.NUM3 AS TIMNUM3,"
 SQ = SQ + " TIM.NUM2 AS TIMNUM2,  TIM.C12,  TIM.C13,  TIM.NUM11,  PEL.EPO AS PELEPO,  PEL.DIE AS PELDIE,  PEL.POL AS PELPOL,  PEL.THL AS PELTHL,  PEL.EPA AS PELEPA,  PEL.AFM AS PELAFM,"
 SQ = SQ + " PEL.PEK AS PELPEK,  PEL.AEG AS PELAEG,"
                         
 SQ = SQ + " PEL.TYP AS PELTYP,  EGGTIM.POSO,  EGGTIM.MONA,  EGGTIM.TIMM,  EGGTIM.KODE,  EGGTIM.FPA AS EGGTIMFPA,  EGGTIM.EKPT,"
 SQ = SQ + " EGGTIM.KAU_AJIA,  EGGTIM.MIK_AJIA,  EGGTIM.ONOMA,  EGGTIM.KOLA,  EGGTIM.MIKTA,  EGGTIM.PELKOD,  EGGTIM.PROELEYSH,  EID.ONO AS EIDONO,  EID.ERG,  EID.XTI,  EID.LTI,EID.LTI5 "
   SQ = SQ + "FROM             EGGTIM INNER JOIN"
 SQ = SQ + "  TIM ON  EGGTIM.ID_NUM =  TIM.ID_NUM INNER JOIN"
 SQ = SQ + "  PARASTAT ON LEFT( TIM.ATIM, 1) =  PARASTAT.EIDOS INNER JOIN"
 SQ = SQ + "  PEL ON  TIM.EIDOS =  PEL.EIDOS AND  TIM.KPE =  PEL.KOD INNER JOIN"
 SQ = SQ + "  EID ON  EGGTIM.KODE =  EID.KOD"


  
  
     Gdb.Execute SQ
     
    
  End If
r2.Close




'   QQ = "{""externalSystemId"":""" + mID_NUM + """ " + ","
'    QQ = QQ + " ""issuerVatNumber"":""" + afmCompany + """" + ","
'    QQ = QQ + " ""invoiceIssueDate"":""" + Format(SQLDThme, "yyyy-MM-dd") + ""","
'
'    QQ = QQ + " ""companyBranch"":""" + SQLDTBRANCH + ""","
'    QQ = QQ + " ""invoiceType"":""" + Split(Mctypos, ";")(0) + ""","


'    QQ = QQ + " ""invoiceSeries"":""" + C_SEIRA + ""","
'    QQ = QQ + " ""invoiceAA"":""" + mID(SQLDTATIM, 2, 6) + ""","
'    QQ = QQ + " ""netValue"":" + Replace(Format(MKAU, "#####0.00"), ",", ".") + ","
'    QQ = QQ + " ""VatAmount"":" + Replace(Format(MFPA, "#####0.00"), ",", ".") + ","
'    QQ = QQ + " ""totalValue"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
'
'    QQ = QQ + " ""paymentAmount"":" + Replace(Format(MFPA + MKAU, "#####0.00"), ",", ".") + ","
'    QQ = QQ + " ""NspCode"":""2"" , "
'    QQ = QQ + " ""terminalId"":""" + TID + """ }"     '"99999069"
'
'    Open "C:\TXTFILES\REQUES111TAPOPAROXO" + extID + ".TXT" For Output As #1
'    Write #1, QQ
'
'    Close #1
'
'    ANSW = REQWINHTTP(QQ, uripos, LTrim(mID_NUM))
'
'    Dim ser As Object
'
'    Set ser = JSON.parse(ANSW)
'
'    F_PAROX_SIGNATURE = ser.Item("paymentToken")("signature")
'    posTimestamp = ser.Item("paymentToken")("timestamp")
'    posuid = ser.Item("uid")






 sql = " CREATE TABLE [dbo].[REQPAYMENTS](" _
       & "[EXTERNALSYSTEMID] [nchar](20) NULL," _
       & "[ISSUERVATNUMBER] [nchar](10) NULL," _
       & "[INVOICEISSUEDATE] DATE NULL," _
       & "[COMPANYBRANCH] [nchar](2) NULL," _
       & "[INVOICETYPE] [nchar](10) NULL," _
       & "[INVOICESERIES] [nchar](10) NULL," _
       & "[INVOICEAA] [nchar](10) NULL," _
       & "[NETVALUE] FLOAT NULL," _
       & "[VATAMOUNT] [float] NULL," _
       & "[TOTALVALUE] [float] NULL," _
       & "[PAYMENTATAMOUNT] [float] NULL," _
       & "[NSPCODE] [NCHAR](10) NULL, " _
       & "[TERMONALID] [nchar](10) NULL," _
       & "[SIGNATURE] [nchar](180) NULL," _
       & "[TIMESTAMP] [nchar](20) NULL," _
       & "[UID] [nchar](50) NULL," _
       & "[MARK] [nchar](20) NULL," _
       & "[ID] [int] IDENTITY(1,1) NOT NULL" _
    & ") ON [PRIMARY]"


         
       
       
       
       
       
       
       
       R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'REQPAYMENTS'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

        If R(0) = 0 Then
           Gdb.Execute sql
        End If
        
        
 dum2 = ADD_FIELD("REQPAYMENTS", "TRANSACTIONID", "NVARCHAR(40) NULL ")  'TransactionId NVARCHAR(40)
       







  sql = " CREATE TABLE [dbo].[APESTALMENA](" _
       & "[AFM] [nchar](10) NULL," _
       & "[ISSUEDATE] [nchar](10) NULL," _
       & "[AA] [nchar](10) NULL," _
       & "[TYPOS] [nchar](10) NULL," _
       & "[PAYTYPE] [nchar](10) NULL," _
       & "[PAYAJI] [float] NULL," _
       & "[TOTALNETVALUE] [float] NULL," _
       & "[TOTALVATAMOUNT] [float] NULL," _
       & "[TOTALGROSSVALUE] [float] NULL, " _
       & "[CLASSIFICATIONTYPE] [nchar](10) NULL," _
       & "[MARK] [nchar](20) NULL," _
       & "[ID] [int] IDENTITY(1,1) NOT NULL" _
    & ") ON [PRIMARY]"


         
       R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'APESTALMENA'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

        If R(0) = 0 Then
           Gdb.Execute sql
        End If
        
        
  
   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE11", "[nchar](10) NULL")
   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE12", "[nchar](10) NULL")
   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT11", "[nchar](10) NULL")
   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT12", "[nchar](10) NULL")
 
   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE21", "[nchar](10) NULL")
   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE22", "[nchar](10) NULL")
   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT21", "[nchar](10) NULL")
   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT22", "[nchar](10) NULL")

   dum2 = ADD_FIELD("APESTALMENA", "MARK2", "[nchar](20) NULL")

    dum2 = ADD_FIELD("APESTALMENA", "AMOUNT1", "float NULL")
    dum2 = ADD_FIELD("APESTALMENA", "AMOUNT2", "float NULL")

    dum2 = ADD_FIELD("APESTALMENA", "ISANTIST", "INT NULL")
    dum2 = ADD_FIELD("APESTALMENA", "ATIM", "NVARCHAR(12) NULL ")

    dum2 = ADD_FIELD("APESTALMENA", "DEMO", "INT NULL ")


 R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'APESTALMENA2'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

        If R(0) = 0 Then
           Gdb.Execute "select top 0 * INTO APESTALMENA2 from APESTALMENA"
        End If






        '<EhFooter>
        Exit Sub

Update_Structures_Err:
        ''MsgBox Err.Description & vbCrLf & _
          "in ADOMERCNEW.kentriko.Update_Structures " & _
          "at line " & Erl, _
          vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.Update_Structures " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Function Max(n1, n2)

        '<EhHeader>
        On Error GoTo max_Err

        '</EhHeader>
100     If n1 > n2 Then Max = n1 Else Max = n2

        '<EhFooter>
        Exit Function

max_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.max " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.max " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function Min(n1, n2)

        '<EhHeader>
        On Error GoTo min_Err

        '</EhHeader>
100     If n1 < n2 Then Min = n1 Else Min = n2

        '<EhFooter>
        Exit Function

min_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.min " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.min " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function CNull(c) As String

        'ΑΝ ΕΙΝΑΙ NULL ΕΠΙΣΤΡΕΦΕΙ " "
        '<EhHeader>
        On Error GoTo CNull_Err

        '</EhHeader>
        On Error Resume Next

100     If IsNull(c) Then
110         CNull = " "
        Else

120         CNull = c
        End If

        '<EhFooter>
        Exit Function

CNull_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.CNull " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.CNull " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function nNull(c) As Single

        'αν ειναι null επιστρεφει 0
        '<EhHeader>
        On Error GoTo nNull_Err

        '</EhHeader>

100     If IsNull(c) Then
110         nNull = 0
        Else
120         nNull = c
        End If

        '<EhFooter>
        Exit Function

nNull_Err:
        '  'MsgBox Err.Description & vbCrLf & _
            "in ADOMERCNEW.kentriko.nNull " & _
            "at line " & Erl, _
            vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.nNull " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function OpenD(ByRef DB As Database)

        '<EhHeader>
        On Error GoTo OpenD_Err

        '</EhHeader>

        'Dim db As Database

100     If gConnect = "dBase III;" Then
            '  Set db = OpenDatabase(gDir, False, False, gConnect)
        Else
            '   Set db = OpenDatabase(gDir, False, False)
        End If

        'db = OpenD
        '<EhFooter>
        Exit Function

OpenD_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.OpenD " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.OpenD " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Sub update_tables()

        '<EhHeader>
        On Error GoTo update_tables_Err

        '</EhHeader>
        Dim DB As Database, R, a, DUM

        ' αν υπάρχει ο πίνακας PINAKES
        '===============================

100     DUM = OpenD(DB)    '  Set DB = OpenDatabase(gDir, False, False, gConnect)

        ' αν δεν υπάρχει δημιουργεί τον πίνακα PINAKES
110     If Len(Dir(gDir + "\pinakes.dbf")) < 2 Then
120         DB.Execute "create table pinakes (typos integer,ayjon integer ,perigrafh text, timh single);"
        End If

        ' ΔΗΜΙΟΥΡΓΕΊ ΤΟ ΙΝΔΕΧ
        On Error GoTo update_tables_Err

        ' αν δεν υπάρχει δημιουργεί τον πίνακα PINAKES
130     If Len(Dir(gDir + "\pinakes.NDX")) < 2 Then
140         DB.Execute "CREATE INDEX PINAKES ON PINAKES (TYPOS,AYJON)"
        End If

150     Set R = DB.OpenRecordset("pinakes")
160     R.Close
170     DB.Close

        'r.Index = "pinakes"

        '<EhFooter>
        Exit Sub

update_tables_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.update_tables " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.update_tables " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Public Function ExecuteLine(Scode As String, Optional fcheckonly As Boolean) As Boolean
'    gvar = ""    ' gia na midenizo tin makroentoli
''ExecuteLine = EbExecuteLine(StrPtr(Scode), 0&, 0&, Abs(fcheckonly)) = 0
'
''  παράδειγμα χρησιμοποίησης μακροεντολής
''  --------------------------------------
'
''   Private Sub Command1_Click()
''   Dim res As Boolean, var As Single
'' αν χρησιμοποιούντα recordset πρεπει να είναι GLOBAL όπως GEID,GPEL,κ.λ.π.
''   res = ExecuteLine("var=2+3*(5+6):a$=var")
''
''   End Sub
'End Function

'

'Sub NOTES()
'' IMPORT APO MEMOTEXT.DBF SE ACCESS
'' select * into Memotext from [dBase IV;database=D:\lageuro].Memotext
'End Sub
Function load_forma(arxeio As String, spacing)

        'spacing=1 klassiko poy ta ypologizei me +1 thn stili toy pedioy (x1)
        'spacing=0 kanonika
        '*********************************************************
        '-----------    GLOBAL ΜΕΤΑΒΛΗΤΕΣ  ----------------
        'gm_str() οι σειρές του f99.txt
        'gpic(i) το στρινγκ του πεδίου
        'gm_r(i),gm_c(i) σειρά και στήλη του πεδίου
        'gm_f(i) το όνομα του πεδίου
        '<EhHeader>
        On Error GoTo load_forma_Err

        '</EhHeader>

        'gm_str(6)="     ^XXXXXXXXXXXXXXXXXXXXXXXXXXX               ^xxxxxxxx    @XXXXXXXX  ~ono_par  ~SUBS(tim->atim,2,5)  ~TIM->HME
        '        gpic(2)='XXXXXXXXXXXXXXXXXXXXXXXXXXX
        'gm_r(2)=5  gm_c(2)=6    gm_f(2)=ono_par
        '========================================================
        'Dim gm_str(1 To 250) As String
        Dim ar_ped(1 To 250) As Integer

        Dim k                As Integer, I As Integer, lastseir As Integer

        Dim npic             As Integer, m_npic As Integer, L1 As Integer, N As Integer

        Dim a(1 To 250)      As Integer

        Dim u                As Integer, mhk_seir As Integer

        Dim X1               As Integer, X2 As Integer

        Dim xa(1 To 30)      As Integer

        '-----------------------------------
        On Error GoTo load_forma_Err

        'Close #1

        'Close #1
100     Open arxeio For Input As #1

110     For k = 1 To 250
120         gm_str(k) = ""
130         gm_f(k) = ""
140         gpic(k) = " "
150         gm_r(k) = 0
160         gm_c(k) = 0
170         ar_ped(k) = 0
        Next

180     I = 1

190     Do While Not EOF(1)
200         Line Input #1, gm_str(I)
210         I = I + 1
        Loop

220     Close #1
        '----------------------------------

230     lastseir = 0
240     npic = 0
        ' ΔΙΑΒΑΖΩ ΤΑ PICTURES ΚΑΙ ΤΗΝ ΘΕΣΗ ΤΟΥΣ

        '*-------------------------------------------------------------------
250     For k = 1 To 250
260         L1 = InStr(gm_str(k), "^")

270         For N = 1 To 30
280             a(N) = 0
            Next

            ' * δεν έχει καθόλου ^
290         If L1 = 0 Then
300             ar_ped(k) = 0
310         ElseIf Left(gm_str(k), 3) = "^^^" Then   'SXOLIA
320             npic = npic + 1
330             ar_ped(k) = 0
            Else
                '* βρέθηκε σειρά με ^ , ψάχνω μήπως έχει και άλλα
340             a(1) = L1    '&& a[]  θέσεις όπου βρέθηκαν τα ^

350             For u = 2 To 30
360                 a(u) = InStr(mID$(gm_str(k), a(u - 1) + 1, Len(gm_str(k))), "^")

370                 If a(u) = 0 Then

                        Exit For

                    End If

380                 a(u) = a(u) + a(u - 1)
                Next

                ' μηκος σειράς
390             mhk_seir = InStr(gm_str(k), "~")

400             If mhk_seir = 0 Then
410                 MsgBox "δεν έχω το σημάδι ~ στην σειρά" + str(k)

                    Exit Function

                End If

420             ar_ped(k) = u - 1

                'ΣΕ ΑΥΤΗΝ ΤΗΝ ΣΕΙΡΑ ΒΡΕΘΗΚΑΝ U-1 ΠΕΔΙΑ

430             For I = 1 To u - 1
440                 npic = npic + 1  'ΑΥΞΑΝΕΙ Ο ΑΡΙΘΜΟΣ ΤΩΝ ΠΕΔΙΩΝ

450                 If spacing = 1 Then  ' bgazei +1  (λανθασμένο κρατείται για συμβατότητα)
460                     X1 = a(I) + 1  '&&  if ( i=1,1,a(i))
470                     X2 = IIf(a(I + 1) = 0, (mhk_seir - 1) - X1 + 1, a(I + 1) - 1 - X1)
                    Else
480                     X1 = a(I) + 1    '&&  if ( i=1,1,a(i))
490                     X2 = IIf(a(I + 1) = 0, (mhk_seir - 1) - X1 + 1, a(I + 1) - X1)
                    End If

                    'το PICTURE του πεδίου
500                 gpic(npic) = mID$(gm_str(k), X1, X2)

                    'η σειρά του πεδίου
510                 gm_r(npic) = IIf(I = 1, k - lastseir, 0)    'αφου είναι στην ίδια σειρά να μην προσθέτει σειρές

                    'η στήλη του πεδίου
520                 gm_c(npic) = a(I)
                Next

530             lastseir = k
            End If

        Next

540     m_npic = npic

550     npic = 0

560     For k = 1 To 250
570         L1 = InStr(gm_str(k), "~")

580         For N = 1 To 30
590             xa(N) = 0
            Next

600         If Left(gm_str(k), 3) = "^^^" Then
610             npic = npic + 1
620         ElseIf L1 = 0 Then        ' &&  ›¤ β®  ΅ζΆ¦¬ ~

            Else
                'βρέθηκε σειρά με ~ , ψάχνω μήπως έχει και άλλα
630             xa(1) = L1    'a[]  θέσεις όπου βρέθηκαν τα ~

640             For u = 2 To 30
650                 xa(u) = InStr(mID$(gm_str(k), xa(u - 1) + 1, Len(gm_str(k))), "~")

660                 If xa(u) = 0 Then

                        Exit For

                    End If

670                 xa(u) = xa(u) + xa(u - 1)
                Next

                'μηκος σειράς
680             mhk_seir = Len(gm_str(k))

                'ΣΕ ΑΥΤΗΝ ΤΗΝ ΣΕΙΡΑ ΒΡΕΘΗΚΑΝ U-1 ΠΕΔΙΑ

690             For I = 1 To u - 1
700                 npic = npic + 1  'ΑΥΞΑΝΕΙ Ο ΑΡΙΘΜΟΣ ΤΩΝ ΠΕΔΙΩΝ
710                 X1 = xa(I) + 1  '&&  if ( i=1,1,a(i))
720                 X2 = IIf(xa(I + 1) = 0, (mhk_seir) - X1 + 1, xa(I + 1) - 1 - X1)

                    '* o titlow του πεδίου
730                 gm_f(npic) = mID$(gm_str(k), X1, X2)

                Next

740             If ar_ped(k) <> u - 1 Then
750                 If ar_ped(k) > u - 1 Then
760                     MsgBox "στην σειρά " + Format(k, "##") + Left(gm_str(k), 40) + ".... εχω παραπάνω  " + str(ar_ped(k) - (u - 1)) + "^  από  ~"
                    Else
770                     MsgBox " στην σειρά " + Format(k, "##") + " " + Left(gm_str(k), 40) + ".... έχω παραπάνω " + str(-ar_ped(k) + (u - 1)) + " ~ από ^ "
                    End If
                End If

780             lastseir = k
            End If

        Next

        '<EhFooter>
        Exit Function

load_forma_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.load_forma " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.load_forma " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Function DTOS(S)
'
'End Function
Function Get_AJ(ByRef pol As String, _
                polepis As String, _
                ago As String, _
                AGOEPIS As String) As Boolean

        '<EhHeader>
        On Error GoTo Get_AJ_Err

        '</EhHeader>

        Dim DB As Database

        Dim R  As New ADODB.Recordset, X As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

100     R.Open UCase("select pol,eidos,ajia_apou from parastat"), Gdb, adOpenDynamic, adLockOptimistic

110     pol = ""
120     R.MoveFirst

130     Do While Not R.EOF

140         If Not IsNull(R("eidos") + R("pol") + R("ajia_apou")) Then
150             If R("pol") = "1" And R("ajia_apou") = "3" Then
160                 pol = pol + "'" + R("eidos") + "',"
                End If

170             If R("pol") = "1" And R("ajia_apou") = "4" Then
180                 polepis = polepis + "'" + R("eidos") + "',"
                End If

190             If R("pol") = "2" And R("ajia_apou") = "1" Then
200                 ago = ago + "'" + R("eidos") + "',"
                End If

210             If R("pol") = "2" And R("ajia_apou") = "2" Then
220                 AGOEPIS = AGOEPIS + "'" + R("eidos") + "',"
                End If

            End If

230         R.MoveNext
        Loop

240     pol = Left(pol, Len(pol) - 1)

250     If Len(polepis) > 0 Then
260         polepis = Left(polepis, Len(polepis) - 1)
        Else
270         polepis = ""
        End If

280     AGOEPIS = Left(AGOEPIS, Len(AGOEPIS) - 1)
290     ago = Left(ago, Len(ago) - 1)
300     Get_AJ = True

        '<EhFooter>
        Exit Function

Get_AJ_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.Get_AJ " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.Get_AJ " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function Get_AJ_ASCII(ByRef pol As String, _
                      polepis As String, _
                      ago As String, _
                      AGOEPIS As String) As Boolean

        '<EhHeader>
        On Error GoTo Get_AJ_ASCII_Err

        '</EhHeader>

        Dim DB As Database

        Dim R  As New ADODB.Recordset, X As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

100     R.Open ("select ISNULL(POL,'1') AS POL ,ISNULL(EIDOS,'_') as EIDOS ,ISNULL(AJIA_APOU,0) AS AJIA_APOU from PARASTAT"), Gdb, adOpenDynamic, adLockOptimistic

110     pol = ""
120     R.MoveFirst

130     Do While Not R.EOF

140         If Not IsNull(R("eidos") + R("pol") + R("ajia_apou")) Then
150             If R("pol") = "1" And R("ajia_apou") = "3" Then
160                 pol = pol + "" + str(Asc(R("eidos"))) + ","
                End If

170             If R("pol") = "1" And R("ajia_apou") = "4" Then
180                 polepis = polepis + "" + str(Asc(R("eidos"))) + ","
                End If

190             If R("pol") = "2" And R("ajia_apou") = "1" Then
200                 ago = ago + "" + str(Asc(R("eidos"))) + ","
                End If

210             If R("pol") = "2" And R("ajia_apou") = "2" Then
220                 AGOEPIS = AGOEPIS + "" + str(Asc(R("eidos"))) + ","
                End If

            End If

230         R.MoveNext
        Loop

240     pol = Left(pol, Len(pol) - 1)

250     If Len(polepis) > 0 Then
260         polepis = Left(polepis, Len(polepis) - 1)
        Else
270         polepis = "999"  'ME KENO DHMIOYRGEI PROBLHMA
        End If

280
290     ago = Left(ago, Len(ago) - 1)
300     Get_AJ_ASCII = True

350     If Len(AGOEPIS) > 0 Then
360         AGOEPIS = Left(AGOEPIS, Len(AGOEPIS) - 1)
        Else
370         AGOEPIS = "999" 'ME KENO DHMIOYRGEI PROBLHMA
        End If

        '<EhFooter>
        Exit Function

Get_AJ_ASCII_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.Get_AJ_ASCII " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.Get_AJ_ASCII " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function Get_EIDPAR(ByRef PolAJ As String, _
                    ByRef ePolAj As String, _
                    ByRef PolPos As String, _
                    ByRef ePolPos As String, _
                    ByRef AgoAJ As String, _
                    ByRef eAgoAJ As String, _
                    ByRef AgoPos As String, _
                    ByRef eAgoPos As String) As Boolean

        '<EhHeader>
        On Error GoTo Get_EIDPAR_Err

        '</EhHeader>

        'Βρίσκει από το PARASTAT τα παραστατικα που σχετίζονται με :
        ' PolAJ , ePolAj    αξια πωλήσεων  επιστροφές πωλήσεν (αξία)
        ' PolPos, ePolPos   ποσότητα  πωλήσεων , επιστροφές πωλήσεν (ποσότητες)
        ' AgoAJ , eAgoAJ    αξία αγορών , επιστροφές αγορών  (αξία)
        ' AgoPos ,eAgoPos   ποσότητα αγορών , επιστροφές αγορών  (ποσότητες)
        'κλήση
        'dim  dummy ,PolAJ As String,  ePolAj As String, _
         PolPos As String,  ePolPos As String, _
         AgoAJ As String,  eAgoAJ As String, _
         AgoPos As String, eAgoPos As String

        'dummy=Get_EIDPAR(PolAJ,ePolAj,PolPos,ePolPos,AgoAJ,eAgoAJ,AgoPos,eAgoPos)

        Dim DB As Database

        Dim R  As New ADODB.Recordset, X As String

100     R.Open UCase("select POL,EIDOS,POS_APOU,AJIA_APOU from PARASTAT"), Gdb, adOpenDynamic, adLockOptimistic

        '  pol = ""
110     R.MoveFirst

120     Do While Not R.EOF

130         If Not IsNull(R("eidos") + R("pol") + R("POS_APOU") + R("ajia_apou")) Then

                '============ΑΞΙΕΣ ==========================
140             If R("pol") = "1" And R("ajia_apou") = "3" Then
150                 PolAJ = PolAJ + "'" + R("eidos") + "',"
                End If

160             If R("pol") = "1" And R("ajia_apou") = "4" Then
170                 ePolAj = ePolAj + "'" + R("eidos") + "',"
                End If

180             If R("pol") = "2" And R("ajia_apou") = "1" Then
190                 AgoAJ = AgoAJ + "'" + R("eidos") + "',"
                End If

200             If R("pol") = "2" And R("ajia_apou") = "2" Then
210                 eAgoAJ = eAgoAJ + "'" + R("eidos") + "',"
                End If

                '-------------- ΠΟΣΟΤΗΤΕΣ --------------------
220             If R("pol") = "1" And R("POS_apou") = "3" Then
230                 PolPos = PolPos + "'" + R("eidos") + "',"
                End If

240             If R("pol") = "1" And R("POS_apou") = "4" Then
250                 ePolPos = ePolPos + "'" + R("eidos") + "',"
                End If

260             If R("pol") = "2" And R("POS_apou") = "1" Then
270                 AgoPos = AgoPos + "'" + R("eidos") + "',"
                End If

280             If R("pol") = "2" And R("POS_apou") = "2" Then
290                 eAgoPos = eAgoPos + "'" + R("eidos") + "',"
                End If

            End If

300         R.MoveNext
        Loop

        Dim Eidh As New Collection

310     If Len(PolAJ) > 0 Then PolAJ = Left(PolAJ, Len(PolAJ) - 1) Else PolAJ = ""
320     If Len(ePolAj) > 0 Then ePolAj = Left(ePolAj, Len(ePolAj) - 1) Else ePolAj = ""

330     If Len(PolPos) > 0 Then PolPos = Left(PolPos, Len(PolPos) - 1) Else PolPos = ""
340     If Len(ePolPos) > 0 Then ePolPos = Left(ePolPos, Len(ePolPos) - 1) Else ePolPos = ""

350     If Len(AgoAJ) > 0 Then AgoAJ = Left(AgoAJ, Len(AgoAJ) - 1) Else AgoAJ = ""
360     If Len(eAgoAJ) > 0 Then eAgoAJ = Left(eAgoAJ, Len(eAgoAJ) - 1) Else eAgoAJ = ""

370     If Len(AgoPos) > 0 Then AgoPos = Left(AgoPos, Len(AgoPos) - 1) Else AgoPos = ""
380     If Len(eAgoPos) > 0 Then eAgoPos = Left(eAgoPos, Len(eAgoPos) - 1) Else eAgoPos = ""

        '<EhFooter>
        Exit Function

Get_EIDPAR_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.Get_EIDPAR " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.Get_EIDPAR " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function FindPar(s As String) As String

        '<EhHeader>
        On Error GoTo FindPar_Err

        '</EhHeader>
        Dim DB, R As New ADODB.Recordset

        'Set db = OpenDatabase(gDir, False, False, gConnect)
100     R.Open "select TITLOS,EIDOS from PARASTAT;", Gdb, adOpenDynamic, adLockOptimistic

110     FindPar = " "
120     R.MoveFirst

130     Do While Not R.EOF

140         If Left(s, 1) = R("EIDOS") Then
150             FindPar = R("TITLOS")

                Exit Do

            End If

160         R.MoveNext
        Loop

170     R.Close

        '<EhFooter>
        Exit Function

FindPar_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.FindPar " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.FindPar " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function GFORMAT(s, F)

        '<EhHeader>
        On Error GoTo GFORMAT_Err

        '</EhHeader>
        Dim G

100     G = Format(s, F)

110     If Left(F, 1) = "#" Then
120         GFORMAT = Right(Space(Len(F)) + G, Len(F))
        Else
130         GFORMAT = G
        End If

        '<EhFooter>
        Exit Function

GFORMAT_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.GFORMAT " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.GFORMAT " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function find_mhnas(hmeromhnia)

        '*********************** βρισκει τον μηνα που θα χρεωθεί *****
        '<EhHeader>
        On Error GoTo find_mhnas_Err

        '</EhHeader>

100     find_mhnas = Right(str((Year(hmeromhnia) - gxron_enar) * 12 + (Month(hmeromhnia) - gmhn_enar + 1)), 2)

        '<EhFooter>
        Exit Function

find_mhnas_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.find_mhnas " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.find_mhnas " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function ChDigEAN13(XX As String) As String

        '--- υπολογισμός  ψηφίου  EAN13
        '<EhHeader>
        On Error GoTo ChDigEAN13_Err

        '</EhHeader>
        Dim num As Long, m As Integer

        Dim k   As Integer

100     num = 0

110     For k = 2 To 12 Step 2
120         num = num + Val(mID$(XX, k, 1))
        Next

130     num = num * 3

140     For k = 1 To 11 Step 2
150         num = num + Val(mID$(XX, k, 1))
        Next

160     k = num Mod 10

170     If k = 0 Then
180         m = 0
        Else
190         m = 10 - k
        End If

200     ChDigEAN13 = Format(m, "0")

        '<EhFooter>
        Exit Function

ChDigEAN13_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.ChDigEAN13 " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.ChDigEAN13 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function check_afm(M_AFM) As Integer

        '<EhHeader>
        On Error GoTo check_afm_Err

        '</EhHeader>
        Dim SUMA, k

100     SUMA = 0
110     check_afm = 1
120     k = 1

130     For k = 1 To 8
140         SUMA = SUMA + Val(mID(M_AFM, k, 1)) * 2 ^ (9 - k)
        Next

150     If SUMA Mod 11 <> Val(Right(Trim(M_AFM), 1)) Then
160         If SUMA Mod 11 = 10 And Val(Right(Trim(M_AFM), 1)) = 0 Then
            Else
170             MsgBox "Λάθος στο ΑΦΜ"
180             check_afm = 0
            End If
        End If

        '<EhFooter>
        Exit Function

check_afm_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.check_afm " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.check_afm " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function
'***      printing_etik(barcode,mtem,mono,mlti,mlti2,mxti,val(mfpa),pel->epo,67 )  ' mpoyg
'printing_etik("354454534534",2,"tem",1.31,1.34,1.23,9,"lagakis",67 )  ' mpoyg
'Function printing_etik(barcode, fores, mperig, mlti, mlti2, mxti, FF, mprom, BAR_MODE)    ' ' mpoyg
''****************************************************************
'    Dim DUM, YCOS, k
'
'    barcode = Left(barcode, 12)
'    YCOS = 50
'    '  *  Wait until at least 10 characters have been received:
'    '*  comOK = COM_OPEN(1, 2000)   // Open COM2
'
'    '*  if .not. comOK
'    '*     wait 'δεν βρήκα την σειριακή ανοικτή '
'    '*  endif
'
'
'
'    For k = 1 To fores
'        Open "COM1" For Output As #1
'
'        Print #1, Chr(27) + "@"
'
'        Print #1, Chr(27) + Chr(51) + Chr(53)
'        '
'        '
'        Print #1, Chr(29) + Chr(33) + Chr(0) + Left(mperig, 25)
'        'Print #1, Chr(10)
'        '
'        '
'        Print #1, Chr(29) + Chr(33) + Chr(0) + "euro :"
'        'Print #1, Chr(29) + Chr(33) + Chr(32) + Format(mlti, "###.##")
'        'Print #1, Chr(10)
'        '
'        '
'        '
'        '    Print #1, Chr(29) + Chr(33) + Chr(0) + "   "
'        '    Print #1, Chr(29) + Chr(33) + Chr(32) + " "
'        '
'        '
'        '
'        '
'        '
'        '   Print #1, Chr(10)
'        '
'        Print #1, Chr(27) + "a" + Chr(1)   ' ALIGN  KENTRAREI
'        Print #1, Chr(29) + "w" + Chr(2)     ' PLATOS
'        Print #1, Chr(29) + Chr(72) + Chr(0)    ' ΠΡΙΝΤΙΝΓ ΠΟΣΙΤΙΟΝ
'        Print #1, Chr(29) + Chr(104) + Chr(38)    ' ΥΨΟΣ
'        '
'        Print #1, Chr(29) + "H" + Chr(2)                            ' GIA NA TYPVNEI AYTOMATA TO N KVDIKO
'        Print #1, Chr(27) + "$" + Chr(40) + Chr(0)
'        Print #1, Chr(29) + "k" + Chr$(BAR_MODE) + Chr$(Len(Trim(barcode))) + Trim(barcode)
'
'
'        If k = fores Then
'            Print #1, Chr(29) + Chr(12)
'        End If
'
'        Close #1
'        ' dum = Shell("C:\PRINT.EXE C:\temp.txt /D:COM1", vbHide)
'
'
'        'set devi to scre
'
'
'    Next
'
'
'
'End Function

Sub HandleError(PROC As String)

        '<EhHeader>
        On Error GoTo HandleError_Err

        '</EhHeader>
        On Error Resume Next

100     Open "c:\MERCVB\errors.txt" For Append As #112

110     Print #112, PROC + "-" + Format(Now, "dd/mm/yyyy  hh:mm") + "-" + Format(Erl, "0000") + "-" + Err.Description
120     Close #112

        '<EhFooter>
        Exit Sub

HandleError_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.HandleError " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.HandleError " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub sqls(ByVal PROC As String)

        '<EhHeader>
        On Error GoTo sqls_Err

        '</EhHeader>

        On Error Resume Next

100     Open "c:\MERCVB\SQLS.txt" For Append As #112

110     Print #112, PROC + "-" + Format(Now, "dd/mm/yyyy  hh:mm") + "-" + Format(Erl, "0000") + "-" + Err.Description
120     Close #112

        '<EhFooter>
        Exit Sub

sqls_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.sqls " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.sqls " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


Sub PARAMETROI_TO_ARRAY()
On Error GoTo sqls_Err
  Dim SQL3S As String:  SQL3S = "select * from PARAMETROI" ' where FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
    
      Dim RR3 As New ADODB.Recordset
      Dim RRN As Integer: RRN = 0
      RR3.Open SQL3S, Gdb, adOpenForwardOnly, adLockReadOnly
      Do While Not RR3.EOF
        RRN = RRN + 1
        gParam(RRN, 0) = RR3!FORMA
        gParam(RRN, 1) = RR3!Var
        gParam(RRN, 2) = RR3!TIMH
         RR3.MoveNext
       Loop
       RR3.Close
       Exit Sub
sqls_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.sqls " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in PARAMETROI_TO_ARRAY " & "at line " & Erl

        Resume Next



End Sub

Function FINDPARAMETROI(KATEG As Integer, FORMA As String, _
                        parametros As String, _
                        DEFAULT As String, _
                        SXOLIA As String)

        '<EhHeader>
        On Error GoTo FINDPARAMETROI_Err
Dim FF As Integer  'FOUND=1
FF = 0
      FINDPARAMETROI = ""
      Dim nn As Long
      
      
      
      'ΟΤΑΝ ΑΛΛΑΖΩ ΤΟ PARAMETROI KAI TO ΞΑΝΑΔΙΑΒΑΖΩ (ΑΥΤ.ΑΡΙΘΜΗΣΗ ΕΙΔΩΝ) ΔΕΝ ΑΝΑΝΕΩΝΕΤΑΙ
      ' ΓΙΑΥΤΟ ΤΟ ΕΚΑΝΑ ΡΕΜ (ΝΕΧΤ 20 ΛΙΝΕΣ)
  If FORMA = "PAR1" And KATEG > 0 Then
      For nn = 1 To UBound(gParam)
       If gParam(nn, 0) = FORMA And gParam(nn, 1) = parametros Then
        FINDPARAMETROI = gParam(nn, 2)
        FF = 1
        Exit For

       End If

      Next
      If FF = 0 Then
         'ΘΑ ΠΑΕΙ ΝΑ ΚΑΝΕΙ ΙΝΣΕΡΤ
      Else
        Exit Function

      End If

End If






        '</EhHeader>
        ' f_autoNumber = Val(FINDPARAMETROI(1, "PELAT2", "F_autoNumber", "0", "Αρίθμηση αυτόματη 00-00-000  =1 Οχι=0"))
        'F_1ST_CHOICE = Val(FindParametroi(1,"PAR2", "F_1ST_CHOICE", "2", "Πρoεπιλεγμένο παραστατικό")) 'posa psifia tha exei h kathe seira
        Dim R   As New ADODB.Recordset

        Dim sql As String

100     If DEFAULT = "DELETE" Then
110         Gdb.Execute "DELETE FROM PARAMETROI WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
120         FINDPARAMETROI = 0

            Exit Function

        End If

130     If DEFAULT = "UPDATE" Then
140         Gdb.Execute "UPDATE PARAMETROI SET VAR='" + parametros + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
            'Gdb.Execute "UPDATE PARAMETROI SET SXOLIA='" + SXOLIA + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
150         FINDPARAMETROI = 0

            Exit Function

        End If

        Dim N As Long

        'SEIR_SELID1

160     sql = "select * from PARAMETROI where FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
    
170     R.Open sql, Gdb, adOpenDynamic, adLockBatchOptimistic

        Dim ll As Long

        ll = R("SXOLIA").DefinedSize

180     If R.EOF Then
190         sql = "insert into PARAMETROI (FORMA,VAR,TIMH,SXOLIA) VALUES ('" + FORMA + "','" + parametros + "','" + DEFAULT + "','" + Left(SXOLIA, ll) + "')"

200         Gdb.Execute sql, N

210         FINDPARAMETROI = DEFAULT
212     Else

214         If IsNull(R("TIMH")) Then
216             Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + DEFAULT + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
218             FINDPARAMETROI = DEFAULT
219         Else
220             FINDPARAMETROI = Trim(R("TIMH"))
221         End If

222     End If

        On Error Resume Next

        'Dim n As Long

230     If Left(SXOLIA, 20) <> Left(R("SXOLIA"), 20) Then
240         Gdb.Execute "UPDATE PARAMETROI SET SXOLIA='" + Left(SXOLIA, ll) + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'", N
        End If

250     If IsNull(R("SXOLIA")) Then
260         Gdb.Execute "UPDATE PARAMETROI SET SXOLIA='" + SXOLIA + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'", N
        End If
        
        
        If IsNull(R("KATEG")) Then
            Gdb.Execute "UPDATE PARAMETROI SET KATEG=1 WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'", N
        End If



If KATEG = -1 Then
       If Trim(R("SXOLIA")) <> Trim(SXOLIA) Then
          Gdb.Execute "UPDATE PARAMETROI SET SXOLIA='" + Replace(SXOLIA, "'", "`") + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
       End If
Else

        If KATEG <> R("KATEG") Then
            Gdb.Execute "UPDATE PARAMETROI SET KATEG=" + str(KATEG) + " WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'", N
        End If
End If
        






270     R.MoveNext

280     If Not R.EOF Then
290         If parametros = Trim(R("VAR")) Then
       
300             Gdb.Execute "DELETE FROM PARAMETROI  WHERE  ID=" + str(R("ID")), N
            End If
        End If

        '<EhFooter>
        Exit Function

FINDPARAMETROI_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.FINDPARAMETROI " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.FINDPARAMETROI " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Function FIND2PARAMETROI(FORMA As String, Var As String, DEFAULT As String, SXOLIA As String)
'   ' me STORED PROCEDURE
'    If DEFAULT = "DELETE" Then
'        Gdb.Execute "DELETE FROM PARAMETROI WHERE  FORMA='" + FORMA + "' AND VAR='" + Var + "'"
'        FIND2PARAMETROI = 0
'        Exit Function
'    End If
'
'    If DEFAULT = "UPDATE" Then
'        Gdb.Execute "UPDATE PARAMETROI SET VAR='" + Var + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + Var + "'"
'        FIND2PARAMETROI = 0
'        Exit Function
'    End If
'
''--------------------------------------------------------
'
'    Dim OB As ADODB.Command
'    Dim N As Integer
'    N = 1
'    Dim PR() As Variant
'    ReDim PR(N)
'
'    On Error Resume Next
'
'    Set OB = New ADODB.Command
'    OB.ActiveConnection = Gdb
'    OB.CommandType = adCmdStoredProc
'    OB.CommandText = "SP_PARAMETROI"  'STORED.Text
'
'    PR(0) = Var    ' 01/01/2006
'
'    PR(1) = FORMA    ' 12/31/2006
'
'
'    Dim R As ADODB.Recordset
'
'    'ΑΝ ΘΕΛΩ ΝΑ ΤΡΑΒΑΩ ΣΕ RECORDSET
'    Set R = OB.Execute(, PR)
'
'Dim sql As String
'
'
'    If R.EOF Then
'        sql = "insert into PARAMETROI (FORMA,VAR,TIMH,SXOLIA) VALUES ('" + FORMA + "','" + Var + "','" + DEFAULT + "','" + SXOLIA + "')"
'        Gdb.Execute sql, N
'        FIND2PARAMETROI = DEFAULT
'    Else
'        FIND2PARAMETROI = R("TIMH")
'    End If
'    Set OB = Nothing
'
'
'
'    If Left(SXOLIA, 20) <> Left(R("SXOLIA"), 20) Then
'        Gdb.Execute "UPDATE PARAMETROI SET SXOLIA='" + SXOLIA + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + Var + "'", N
'    End If
'
'    If IsNull(R("SXOLIA")) Then
'        Gdb.Execute "UPDATE PARAMETROI SET SXOLIA='" + SXOLIA + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + Var + "'", N
'    End If
'End Function

Function ADD_FIELD(ByVal Table, ByVal FIELD, ByVal FIELDTYPE)

        '-----------------------------------------------------------------------------------
        '<EhHeader>
        On Error GoTo ADD_FIELD_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

        Dim sql

100     sql = "SELECT TOP 0 * FROM " + Table + ";"
110     MDIForm1.Caption = Table + FIELD
120     R.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

        Dim k As Long

        Dim OK


130     OK = 0

140     For k = 0 To R.FIELDS.Count - 1

150         If R.FIELDS(k).Name = FIELD Then
160             OK = 1
            End If

        Next

170     ADD_FIELD = 0

180     If OK = 0 Then
190         sql = "alter table " + Table + " ADD " + FIELD + " " + FIELDTYPE
200         ADD_FIELD = 1
            On Error GoTo ADD_FIELD_Err
210         Gdb.Execute sql
        End If

        '<EhFooter>
        Exit Function

ADD_FIELD_Err:
        ' 'MsgBox Err.Description & vbCrLf & _
           "in ADOMERCNEW.kentriko.ADD_FIELD " & _
           "at line " & Erl, _
           vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Table + " " + FIELD + "  " + Err.Description & " in ADOMERCNEW.kentriko.ADD_FIELD " & "at line " & Erl

        Exit Function

        ' Resume Next
        '</EhFooter>

End Function

Function ADD_FIELDDAO(ByVal Table, ByVal FIELD, ByVal FIELDTYPE, DAODB As Database)

        '<EhHeader>
        On Error GoTo ADD_FIELDDAO_Err

        '</EhHeader>
        Dim R As Recordset

        'Dim DAODB As Database
        Dim sql

        'Set DAODB = OpenDatabase(DB)

100     sql = "SELECT TOP 1 * FROM " + Table + ";"

110     Set R = DAODB.OpenRecordset(sql)

        Dim k As Long

        Dim OK

120     OK = 0

130     For k = 0 To R.FIELDS.Count - 1

140         If R.FIELDS(k).Name = FIELD Then
150             OK = 1
            End If

        Next

160     If OK = 0 Then
170         sql = "alter table " + Table + " ADD " + FIELD + " " + FIELDTYPE
180         R.Close

190         DAODB.Execute sql
        End If

        '<EhFooter>
        Exit Function

ADD_FIELDDAO_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.ADD_FIELDDAO " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.ADD_FIELDDAO " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function UpdateParametroi(FORMA As String, _
                          parametros As String, _
                          DEFAULT As String, _
                          SXOLIA As String)

        '<EhHeader>
        On Error GoTo UpdateParametroi_Err

        '</EhHeader>

        'F_1ST_CHOICE = Val(FindParametroi(1,"PAR2", "F_1ST_CHOICE", "2", "Πρoεπιλεγμένο παραστατικό")) 'posa psifia tha exei h kathe seira
        Dim R   As New ADODB.Recordset

        Dim sql As String

        On Error Resume Next

100     Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + DEFAULT + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
110     UpdateParametroi = 0

        '<EhFooter>
        Exit Function

UpdateParametroi_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.UpdateParametroi " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.UpdateParametroi " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Sub UPDATE_YPOLOIPA2()
        'into DOKEGGT1
        'B = "CREATE VIEW dbo.[EIDT3]" _

        '<EhHeader>
        On Error GoTo UPDATE_YPOLOIPA2_Err

        '</EhHeader>

        Dim a As String

        Dim B As String

        Dim D As Long

        'Dim krit As String

        ' ΒΡΙΣΚΩ ΜΙΑ ΦΟΡΑ ΤΑ ΥΠΟΛΟΙΠΑ ΑΝΑ ΑΠΟΘΗΚΗ

        On Error Resume Next

100     Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

110     a = "SELECT KODE," & " SUM(CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X," & " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P," & " SUM(CASE APOT  WHEN 2  THEN XRE ELSE 0  END ) AS S2X," & " SUM(CASE APOT  WHEN 2  THEN PIS ELSE 0  END ) AS S2P," & " SUM(CASE APOT  WHEN 3  THEN XRE ELSE 0  END ) AS S3X," & " SUM(CASE APOT  WHEN 3  THEN PIS ELSE 0  END ) AS S3P," & " SUM(CASE APOT  WHEN 4  THEN XRE ELSE 0  END ) AS S4X," & " SUM(CASE APOT  WHEN 4  THEN PIS ELSE 0  END ) AS S4P  "
120     a = a + " INTO DOKEGGT1 FROM EGGTIM  "
130     a = a + " GROUP BY KODE"

        Dim TT As Long

140     TT = GetCurrentTime()
150     Gdb.Execute a, D

        'MsgBox GetCurrentTime() - TT

160     If D = 0 Then
170         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"

            Exit Sub

        End If

180     Gdb.Execute "UPDATE EID SET POS01=0,POS02=0,POS03=0,POS04=0,POS=0"

        'b = "CREATE VIEW dbo.[EIDT3]" _
        '& "AS " _
        '& " SELECT     dbo.EID.*, (dbo.DOKEGGT1.S1X-dbo.DOKEGGT1.S1P) AS EXPR1, (dbo.DOKEGGT1.S2X-dbo.DOKEGGT1.S2P) AS EXPR2 " _
        '& " FROM         dbo.DOKEGGT1 INNER JOIN " _
        '& " dbo.EID ON dbo.DOKEGGT1.KODE = dbo.EID.KOD"
        'Gdb.Execute b, d

190     Gdb.Execute "update EID SET" & " POS01=(CASE WHEN S1X IS NULL THEN 0 ELSE S1X END) - (CASE WHEN S1P IS NULL THEN 0 ELSE S1P END)," & " POS02=(CASE WHEN S2X IS NULL THEN 0 ELSE S2X END) - (CASE WHEN S2P IS NULL THEN 0 ELSE S2P END)," & " POS03=(CASE WHEN S3X IS NULL THEN 0 ELSE S3X END) - (CASE WHEN S3P IS NULL THEN 0 ELSE S3P END)," & " POS04=(CASE WHEN S4X IS NULL THEN 0 ELSE S4X END) - (CASE WHEN S4P IS NULL THEN 0 ELSE S4P END)" & " FROM EID INNER JOIN DOKEGGT1 ON EID.KOD=DOKEGGT1.KODE ", D

200     Gdb.Execute "UPDATE EID SET  POS = POS01 + POS02+POS03+POS04", D

210     D = 0
        'Gdb.Execute "update EIDT3 SET POS01=EXPR1,POS02=EXPR2,POS=EXPR1+EXPR2 ", d

220     D = D + 1

        '--------------------------------------------------------------------
        Exit Sub

        '
        ''into DOKEGGT1
        ''B = "CREATE VIEW dbo.[EIDT3]" _
        '
        '
        'Dim A As String
        'Dim b As String
        'Dim d As Long
        ''Dim krit As String
        '
        '' ΒΡΙΣΚΩ ΜΙΑ ΦΟΡΑ ΤΑ ΥΠΟΛΟΙΠΑ ΑΝΑ ΑΠΟΘΗΚΗ
        '
        'On Error Resume Next
        '
        'Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1
        '
        '
        '
        'A = "SELECT KODE,SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X," _
        '& " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P," _
        '& "SUM( CASE APOT  WHEN 2 THEN XRE ELSE 0  END ) AS S2X," _
        '& "SUM( CASE APOT WHEN 2 THEN PIS ELSE 0 END ) AS S2P INTO DOKEGGT1 "
        'A = A + " FROM EGGTIM  "
        'A = A + " GROUP BY KODE"
        'Gdb.Execute A, d
        '
        'If d = 0 Then
        '   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
        '   Exit Sub
        'End If
        '
        'Gdb.Execute "UPDATE EID SET POS01=0,POS02=0,POS=0"
        '
        ''b = "CREATE VIEW dbo.[EIDT3]" _
        ''& "AS " _
        ''& " SELECT     dbo.EID.*, (dbo.DOKEGGT1.S1X-dbo.DOKEGGT1.S1P) AS EXPR1, (dbo.DOKEGGT1.S2X-dbo.DOKEGGT1.S2P) AS EXPR2 " _
        ''& " FROM         dbo.DOKEGGT1 INNER JOIN " _
        ''& " dbo.EID ON dbo.DOKEGGT1.KODE = dbo.EID.KOD"
        ''Gdb.Execute b, d
        '
        'Gdb.Execute "update EID SET POS01=(CASE WHEN S1X IS NULL THEN 0 ELSE S1X END) - (CASE WHEN S1P IS NULL THEN 0 ELSE S1P END)," _
        '& " POS02=(CASE WHEN S2X IS NULL THEN 0 ELSE S2X END) - (CASE WHEN S2P IS NULL THEN 0 ELSE S2P END)" _
        '& " FROM EID INNER JOIN DOKEGGT1 ON EID.KOD=DOKEGGT1.KODE "
        '
        '
        'Gdb.Execute "UPDATE EID SET  POS = POS01 + POS02"
        '
        '
        'd = 0
        ''Gdb.Execute "update EIDT3 SET POS01=EXPR1,POS02=EXPR2,POS=EXPR1+EXPR2 ", d
        '
        '
        'd = d + 1
        '<EhFooter>
        Exit Sub

UPDATE_YPOLOIPA2_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.UPDATE_YPOLOIPA2 " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.UPDATE_YPOLOIPA2 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


Public Sub TelHmeAgoras(ByVal APO As Date, ByVal EOS As Date)
MDIForm1.Caption = " "
        '===================================================================================
        '<EhHeader>
        On Error GoTo TelHmeAgoras_Err

        '</EhHeader>
        Dim arx_arit, tel_arit, buff, kod_tim, ono_par, p1, progr, k

        Dim m_rec, m_synt, synt1, synt2

        Dim SUMAEID As New ADODB.Recordset

100     arx_arit = 0: tel_arit = 0

        ' sele TIM
        'SET index to ARTHME

        Dim polhs As String

        Dim PolXParoxh As String

        Dim Paroxh As String

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

        Dim a

110     a = Get2_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh)

        Dim eggtim As New ADODB.Recordset

        ''E','α',
120     eggtim.Open "SELECT * FROM EGGTIM WHERE LEFT(ATIM,1) IN ('E','α','λ'," + ago + ") AND HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "' ORDER BY HME", Gdb, adOpenDynamic, adLockOptimistic

        Dim R As New ADODB.Recordset

        Dim EID As New ADODB.Recordset

130     If eggtim.EOF Then

            Exit Sub

        End If

140     eggtim.MoveFirst

        Dim B, D, ATIM

150     k = 0

        Dim N
 MDIForm1.Caption = Format(k, "####")
160     Do While Not eggtim.EOF
            'ΒΡΙΣΚΩ ΤΗΝ ΤΕΛΕΥΤΑΙΑ ΤΙΜΗ ΑΓΟΡΑΣ
170         a = eggtim("KODE")
180         D = eggtim("HME")
190         ATIM = eggtim("atim")
200         k = k + 1

210         If Not IsNull(eggtim("kode")) Then
220             If IsNull(eggtim("TIMM")) Then
                Else

230                 If IsNull(eggtim("EKPT")) Then
                    Else

240                     If eggtim("TIMM") > 0 Then
250                         Gdb.Execute "UPDATE EID SET XTI=" + str(nNull(eggtim("TIMM") * (1 - nNull(eggtim("EKPT")) / 100) * (1 - nNull(eggtim("EKPT2")) / 100))) + ",CH1='" + ATIM + "',HPAR='" + Format(D, "MM/DD/YYYY") + "',PAR=" + str(eggtim("POSO")) + "  where KOD='" + a + "'", N
                        Else
260                         Gdb.Execute "UPDATE EID SET CH1='" + ATIM + "',HPAR='" + Format(D, "MM/DD/YYYY") + "',PAR=" + str(eggtim("POSO")) + "  where KOD='" + a + "'", N
                        End If
                    End If
              
                End If
            End If

270         If N > 0 Then
280             MDIForm1.Caption = Format(k, "####") + " - " + Format(D, "dd/mm/yyyy hh:mm") + " "
            End If

290         If (k Mod 10 = 0) Then DoEvents

300         DoEvents

310         eggtim.MoveNext
        Loop

320     eggtim.Close

        '<EhFooter>
        Exit Sub

TelHmeAgoras_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.TelHmeAgoras " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.TelHmeAgoras " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub






Public Function Get2_AJ(ByRef pol As String, _
                 polepis As String, _
                 ago As String, _
                 AGOEPIS As String, _
                 ByRef PolXParoxh As String, _
                 Paroxh As String) As Boolean

        '<EhHeader>
        On Error GoTo Get2_AJ_Err

        '</EhHeader>

        Dim DB As Database

        Dim R  As New ADODB.Recordset, X As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

100     R.Open "select POL,EIDOS,AJIA_APOU,SHM1 from PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

110     pol = ""
120     PolXParoxh = ""

130     Paroxh = ""

140     R.MoveFirst

150     Do While Not R.EOF

160         If Not IsNull(R("eidos") + R("pol") + R("ajia_apou")) Then
170             If R("pol") = "1" And R("ajia_apou") = "3" Then
180                 pol = pol + "'" + R("eidos") + "',"
                End If

190             If R("pol") = "1" And R("ajia_apou") = "3" Then
200                 If Not IsNull(R("shm1")) And Val(Left(R("shm1"), 2)) <> 2 Then
210                     PolXParoxh = PolXParoxh + "'" + R("eidos") + "',"
                    End If
                End If

220             If R("pol") = "1" And R("ajia_apou") = "3" Then
230                 If Not IsNull(R("shm1")) And Val(Left(R("shm1"), 2)) = 2 Then
240                     Paroxh = Paroxh + "'" + R("eidos") + "',"
                    End If
                End If

250             If R("pol") = "1" And R("ajia_apou") = "4" Then
260                 polepis = polepis + "'" + R("eidos") + "',"
                End If

270             If R("pol") = "2" And R("ajia_apou") = "1" Then
280                 ago = ago + "'" + R("eidos") + "',"
                End If

290             If R("pol") = "2" And R("ajia_apou") = "2" Then
300                 AGOEPIS = AGOEPIS + "'" + R("eidos") + "',"
                End If
            End If

310         R.MoveNext
        Loop

320     pol = Left(pol, Len(pol) - 1)
330     PolXParoxh = Left(PolXParoxh, Len(PolXParoxh) - 1)

        On Error Resume Next

340     Paroxh = Left(Paroxh, Len(Paroxh) - 1)

350     If Len(polepis) > 0 Then
360         polepis = Left(polepis, Len(polepis) - 1)
        Else
370         polepis = ""
        End If

380     AGOEPIS = Left(AGOEPIS, Len(AGOEPIS) - 1)
390     ago = Left(ago, Len(ago) - 1)
400     Get2_AJ = True

        '<EhFooter>
        Exit Function

Get2_AJ_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Get2_AJ " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Get2_AJ " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function


















'Sub UPDATE_YPOLOIPA()
''=================================================================================
'    Dim a As String
'    Dim b As String
'    Dim D As Long
'    'Dim krit As String
'
'    ' ΒΡΙΣΚΩ ΜΙΑ ΦΟΡΑ ΤΑ ΥΠΟΛΟΙΠΑ ΑΝΑ ΑΠΟΘΗΚΗ
'
'    On Error Resume Next
'    Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1
'    Gdb.Execute "DROP TABLE DOKEGGT2"  'ΣΟΥΜΕΣ ΑΠΟΘΗΚΗΣ 2
'    'DOKEGGT12 VIEW =  ΕΝΩΝΕΙ ΤΑ DOKEGGT1 KAI DOKEGGT2
'    ' VIEW [EID-DOKEGGT3] ΓΙΑ ΤΟ UPDATE TOY EID   POS01,POS02
'
'    Gdb.Execute "DROP TABLE DOKEGGT3"    ' TA APOTELESMATA TOY DOKEGGT12 ΣΕ ΝΕΟ ΑΡΧΕΙΟ
'
'
'    a = "SELECT KODE,SUM(XRE) AS S1X,SUM(PIS) AS S1P,0 AS S31,0 AS S32 INTO DOKEGGT1 "
'    a = a + " FROM EGGTIM WHERE APOT=1 OR APOT IS NULL  GROUP BY KODE"
'
'    b = "SELECT KODE,SUM(XRE) AS S2X,SUM(PIS) AS S2P INTO DOKEGGT2 "
'    b = b + " FROM EGGTIM WHERE APOT=2  GROUP BY KODE "    '  AND HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY")
'    Gdb.Execute a, D
'
'    If D = 0 Then
'        MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
'        Exit Sub
'    End If
'
'
'    Gdb.Execute "UPDATE DOKEGGT1 SET S1X=0 WHERE S1X IS NULL"
'
'    Gdb.Execute "UPDATE DOKEGGT1 SET S1P=0 WHERE S1P IS NULL"
'
'
'
'
'    D = 0
'
'
'
'    Gdb.Execute b, D
'    If D = 0 Then
'        MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT2"
'        ' Exit Sub
'    End If
'
'    If D <> 0 Then    ' ΕΧΕΙ ΠΑΡΑΠΑΝΩ ΑΠΟ ΜΙΑ ΑΠΟΘΗΚΗ
'        Gdb.Execute "UPDATE DOKEGGT2 SET S2X=0 WHERE S2X IS NULL"
'        Gdb.Execute "UPDATE DOKEGGT2 SET S2P=0 WHERE S2P IS NULL"
'
'        ' =  ΕΝΩΝΕΙ ΤΑ DOKEGGT1 KAI DOKEGGT2
'        Gdb.Execute "CREATE VIEW DOKEGGT12   AS SELECT DOKEGGT1.*,S2X,S2P FROM DOKEGGT1 LEFT JOIN DOKEGGT2 ON DOKEGGT1.KODE=DOKEGGT2.KODE "
'        a = "SELECT KODE AS KOD,S1X-S1P AS YP1,EXPR1-EXPR2 AS YP2 "
'        a = a + " INTO DOKEGGT3 FROM DOKEGGT12 "
'
'    Else  ' ΕΧΕΙ ΜΟΝΟ ΜΙΑ ΑΠΟΘΗΚΗ
'
'
'        ' Gdb.Execute "CREATE VIEW DOKEGGT12 AS SELECT DOKEGGT1.*,S2X,S2P FROM DOKEGGT1 ", d
'        a = "SELECT KOD,S1X-S1P AS YP1,0 AS YP2 "
'        a = a + " INTO DOKEGGT3 FROM DOKEGGT1 INNER JOIN EID ON DOKEGGT1.KODE=EID.KOD "
'
'    End If
'
'    Gdb.Execute a, D
'
'
'    Gdb.Execute "CREATE VIEW dbo.[EID-DOKEGGT3]" _
'              & "AS " _
'              & " SELECT     dbo.EID.*, dbo.DOKEGGT3.YP1 AS EXPR1, dbo.DOKEGGT3.YP2 AS EXPR2 " _
'              & " FROM         dbo.DOKEGGT3 INNER JOIN " _
'              & " dbo.EID ON dbo.DOKEGGT3.KOD = dbo.EID.KOD"
'
'
'
'    D = 0
'    Gdb.Execute "update [EID-DOKEGGT3] SET POS01=EXPR1,POS02=EXPR2,POS=EXPR1+EXPR2 ", D
'
'
'End Sub

Function gAlignFormat(MM, PIC)

        ' RIGHT ALIGNMENT
        '<EhHeader>
        On Error GoTo gAlignFormat_Err

        '</EhHeader>
100     If InStr(PIC, "#") > 0 Then
110         gAlignFormat = Right(Space(Len(PIC)) + MM, Len(Trim(PIC)))
        Else
            ' LEFT ALIGNMENT
120         gAlignFormat = Left(MM + Space(Len(PIC)), Len(PIC))
        End If

        '<EhFooter>
        Exit Function

gAlignFormat_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.gAlignFormat " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.gAlignFormat " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function SameLetters(ByVal st As String) As String

        'ΓΙΑ ΝΑ ΨΑΧΝΕΙ ΚΑΙ ΜΕ ΤΟ Α ΤΟ ΛΑΤΙΝΙΚΟ ΚΑΙ ΜΕ ΤΟ Α ΤΟ ΕΛΛΗΝΙΚΟ
        'ΔΙΝΩ ΕΝΑ ΣΤΡΙΝΚ Π.Χ. 'ΜΟ' ΚΑΙ ΜΟΥ ΕΠΙΣΤΡΕΦΕΙ '[ΜΜ][ΟΟ]' ΓΙΑ
        ' ΝΑ ΜΟΥ ΒΡΙΣΚΕΙ ΚΑΙ ΤΟ "MO BITEL"   KAI TO  "ΜΟ ΧΩΡΙΔΗΣ"
        '<EhHeader>
        On Error GoTo SameLetters_Err


         
         If gAsterisk = 1 Then
              st = Replace(st, "*", "%")
         End If




        '</EhHeader>
        Dim L As Integer, k As Integer

        Dim s As String

        Dim N As Integer

        Dim c As String

100     If gCapitals = 1 Then
110         st = UCase(st)
        End If
    
120     st = Replace(st, "'", "~")
        
        If Len(st) > 1 Then
122     '   ST = Replace(ST, "*", "%")
        End If
        
130     L = Len(st)
140     s = ""

150     For k = 1 To L
160         c = mID$(st, k, 1)
170         N = InStr("αΑaAβΒbBγΓgGδΔdDεΕeEζΖzZηΗhHθΘuUιΙiIκΚkKλΛlLμΜmMνΝnNJjΞξοΟoOπΠpPρΡrRσςΣsSτΤtTυΥyYφΦfFχΧxXψΨcCωΩvV", c)

180         If N = 0 Then
190             s = s + c
200         ElseIf N > 0 Then
                    
210
                    If c = "α" Or c = "Α" Or c = "a" Or c = "A" Then
                        s = s + "[AΑaαάΆ]"
                    ElseIf c = "β" Or c = "Β" Or c = "b" Or c = "B" Then
                        s = s + "[BΒbβ]"
                    
                        
                       
310                 ElseIf c = "γ" Or c = "Γ" Or c = "g" Or c = "G" Then
320                     s = s + "[GΓgγ]"

270                 ElseIf c = "δ" Or c = "Δ" Or c = "d" Or c = "D" Then
280                     s = s + "[DΔdδ]"

430                 ElseIf c = "ε" Or c = "Ε" Or c = "e" Or c = "E" Then
440                     s = s + "[EΕεeέΈ]"


790                 ElseIf c = "ζ" Or c = "Ζ" Or c = "z" Or c = "Z" Then
800                     s = s + "[ZΖzζ]"

690                 ElseIf c = "η" Or c = "Η" Or c = "h" Or c = "H" Then
700                     s = s + "[HΗηhήΉ]"


230                 ElseIf c = "θ" Or c = "Θ" Or c = "u" Or c = "U" Then
240                     s = s + "[UΘθu]"


                    ElseIf c = "ι" Or c = "Ι" Or c = "i" Or c = "I" Then
                        s = s + "[IΙιiίΊ]"
                        
                        
730                 ElseIf c = "κ" Or c = "Κ" Or c = "k" Or c = "K" Then
740                     s = s + "[KΚκk]"



350                 ElseIf c = "λ" Or c = "Λ" Or c = "l" Or c = "L" Then
360                     s = s + "[LΛlλ]"

930                 ElseIf c = "μ" Or c = "Μ" Or c = "m" Or c = "M" Then
940                     s = s + "[MΜμm]"


890                 ElseIf c = "ν" Or c = "Ν" Or c = "n" Or c = "N" Then
900                     s = s + "[NΝνn]"

330                 ElseIf c = "ξ" Or c = "Ξ" Or c = "j" Or c = "J" Then
340                     s = s + "[JΞξj]"

                    ElseIf c = "ο" Or c = "Ο" Or c = "o" Or c = "O" Then
                        s = s + "[OΟοoόΌ]"

610                 ElseIf c = "π" Or c = "Π" Or c = "p" Or c = "P" Then
620                     s = s + "[pPπΠ]"


                    ElseIf c = "ρ" Or c = "Ρ" Or c = "r" Or c = "R" Then
220                     s = s + "[RΡρr]"



250                 ElseIf c = "σ" Or c = "Σ" Or c = "s" Or c = "S" Or c = "ς" Then
260                     s = s + "[SΣσςs]"

470                 ElseIf c = "τ" Or c = "Τ" Or c = "t" Or c = "T" Then
480                     s = s + "[TΤτt]"


                    ElseIf c = "υ" Or c = "Υ" Or c = "y" Or c = "Y" Then
                        s = s + "[YΥυy]"


290                 ElseIf c = "φ" Or c = "Φ" Or c = "f" Or c = "F" Then
300                     s = s + "[FΦφf]"


830                 ElseIf c = "χ" Or c = "Χ" Or c = "x" Or c = "X" Then
840                     s = s + "[XΧxχ]"




370                 ElseIf c = "ψ" Or c = "Ψ" Or c = "c" Or c = "C" Then
380                     s = s + "[CΨcψ]"

390                 ElseIf c = "ω" Or c = "Ω" Or c = "v" Or c = "V" Then
400                     s = s + "[VΩωv]"
                    'ElseIf C = "*" Then
                     '   s = s + "%"

                End If
                

            End If

        Next

970     SameLetters = s

        '<EhFooter>
        Exit Function

SameLetters_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.SameLetters " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.SameLetters " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Sub MILSEC(ByVal MILISECONDS As Long)

        '<EhHeader>
        On Error GoTo MILSEC_Err

        '</EhHeader>
        Dim Start As Long

100     Start = GetCurrentTime()

        Do
110     Loop Until GetCurrentTime() - Start > MILISECONDS

        '<EhFooter>
        Exit Sub

MILSEC_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.MILSEC " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.MILSEC " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function ypoloipa_pel(ByVal mBUFF As String, _
                      ByVal APO As Date, _
                      ByVal EOS As Date, ByVal ENERGOS As Integer) As Single

        'ypologismos ypoloipon
        '<EhHeader>
        On Error GoTo ypoloipa_pel_Err

        '</EhHeader>
        Dim X As String, order As String

        Dim k As Integer, R As New ADODB.Recordset    ', sum As Single

        Dim execX, file

        ' F_RUN = 1

        On Error Resume Next

100     Gdb.Execute "DROP TABLE EGG1"

        Dim L As Long

        Dim Q As String: Q = ""
        If ENERGOS = 1 Then
          Q = " ENERGOS=1 AND "
        End If


        Gdb.Execute "UPDATE PEL  SET TYP=isnull(AYP)+ (SELECT SUM(isnull(XREOSI,0))-SUM(isnull(PISTOSI,0)) FROM EGG WHERE EGG.KOD=PEL.KOD AND EGG.EIDOS=PEL.EIDOS)  FROM" & " PEL WHERE " + Q + " EIDOS='" + mBUFF + "'"

'110     Gdb.Execute "UPDATE PEL  SET TYP=0,XRESYN=0,PISSYN=0 FROM" & " PEL WHERE " + Q + " EIDOS='" + mBUFF + "'", L
'
'120     If gXEIRISTHS >= 8 Then
'
'130         Gdb.Execute "SELECT UPPER(EIDOS) as MEIDOS,KOD,SUM(isnull(XREOSI,0)) AS SX," & " SUM(isnull(PISTOSI,0)) AS SP INTO EGG1" & " FROM EGG where  EIDOS='" + mBUFF + "' and  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "' GROUP BY UPPER(EIDOS),KOD "
'
'        Else
'
'140         Gdb.Execute "SELECT EIDOS as MEIDOS,KOD,SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END ) AS SX," & " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END ) AS SP INTO EGG1" & " FROM EGG where  EIDOS='" + mBUFF + "' and HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "' GROUP BY EIDOS,KOD "
'
'        End If
'
'150     Gdb.Execute "UPDATE PEL  SET TYP=AYP WHERE  " + Q + " EIDOS='" + mBUFF + "'"
'
'160     If mBUFF = "e" Then    '  "e",
'170         If gXEIRISTHS >= 8 Then
'180             Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" & " PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND UPPER(PEL.EIDOS)=EGG1.MEIDOS WHERE  " + Q + " EIDOS='" + mBUFF + "'", L
'            Else
'190             Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" & " PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.EIDOS WHERE  " + Q + " EIDOS='" + mBUFF + "'"", L"
'            End If
'
'        Else
'
'            On Error GoTo ypoloipa_pel_Err
'
'200         If gXEIRISTHS = 8 Or gXEIRISTHS = 2 Then
'210             Gdb.Execute "UPDATE PEL  SET TYP=AYP-(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )+" & " (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" & " PEL INNER JOIN EGG1 ON PEL.KOD=EGG1.KOD AND UPPER(PEL.EIDOS)=EGG1.MEIDOS  WHERE  " + Q + " EIDOS='" + mBUFF + "'"
'            Else
'
'                'On Error Resume Next
'220             Gdb.Execute "UPDATE PEL  SET TYP=AYP-( ISNULL(EGG1.SX,0) )+" & " ( isnull(EGG1.SP,0) ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" & " PEL INNER JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.MEIDOS WHERE  " + Q + " EIDOS='" + mBUFF + "'"
'            End If
'
'        End If

230     R.Open "select SUM(TYP) FROM PEL WHERE  " + Q + " EIDOS='" + mBUFF + "'", Gdb, adOpenForwardOnly, adLockReadOnly
240     ypoloipa_pel = R(0)

        '<EhFooter>
        Exit Function

ypoloipa_pel_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.ypoloipa_pel " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.ypoloipa_pel " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function Ypoloipo_Pel(ByVal mBUFF As String, ByVal MKOD As String, ByVal ENERGOS As Integer) As Single

        '<EhHeader>
        On Error GoTo Ypoloipo_Pel_Err

        '</EhHeader>
        Dim R   As New ADODB.Recordset

        Dim ayp As Single

        Dim sql As String
        Dim Q As String: Q = ""
        If ENERGOS = 1 Then
          Q = " ENERGOS=1 AND "
        End If
        
        

100     R.Open "select AYP FROM PEL WHERE " + Q + " EIDOS='" + mBUFF + "' AND KOD='" + MKOD + "'", gConnect, adOpenForwardOnly, adLockReadOnly

110     If Not R.EOF Then
120         If IsNull(R(0)) Then
130             ayp = 0
            Else
140             ayp = R(0)
            End If

        Else
150         ayp = 0
        End If

160     R.Close

170     sql = "select sum(XREOSI),sum(PISTOSI) from EGG WHERE EIDOS='" + mBUFF + "' and KOD='" + MKOD + "'"
180     R.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

190     If R.EOF Then
200         Ypoloipo_Pel = 0
        Else

210         If IsNull(R(0)) Then
220             Ypoloipo_Pel = ayp
            Else
230             Ypoloipo_Pel = ayp + R(0) - R(1)
            End If
        End If

240     R.Close

        '<EhFooter>
        Exit Function

Ypoloipo_Pel_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.Ypoloipo_Pel " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.Ypoloipo_Pel " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function Ypoloipo_today(ByVal mBUFF As String, _
                        ByVal MKOD As String, _
                        mnow As Date) As Single

        '<EhHeader>
        On Error GoTo Ypoloipo_today_Err

        '</EhHeader>
        Dim R   As New ADODB.Recordset

        Dim ayp As Single

        Dim sql As String

        'R.Open "select AYP FROM PEL WHERE EIDOS='" + mBUFF + "' AND KOD='" + mkod + "'", gConnect, adOpenForwardOnly, adLockReadOnly
        ' If Not R.EOF Then
        '   If IsNull(R(0)) Then
        '      ayp = 0
        '   Else
        '      ayp = R(0)
        '   End If
        ' Else
        '    ayp = 0
        ' End If

        'R.Close

100     sql = "select sum(XREOSI),sum(PISTOSI) from EGG WHERE EIDOS='" + mBUFF + "' and KOD='" + MKOD + "' and HME>='" + Format(mnow, "mm/dd/yyyy") + "'  and HME<'" + Format(DateAdd("d", 1, mnow), "mm/dd/yyyy") + "'"
110     R.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

120     If R.EOF Then
130         Ypoloipo_today = 0
        Else

140         If IsNull(R(0)) Then
150             Ypoloipo_today = 0
            Else
160             Ypoloipo_today = R(0) - R(1)
            End If
        End If

170     R.Close

        '<EhFooter>
        Exit Function

Ypoloipo_today_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.Ypoloipo_today " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.Ypoloipo_today " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function add_schedule(DATE1, Aitiologia, Topos)

        '<EhHeader>
        On Error GoTo add_schedule_Err

        '</EhHeader>
        Dim mXML

100     mXML = "<Calendar CompactMode="

110     mXML = mXML + """"
120     mXML = mXML + "1"
130     mXML = mXML + """"
140     mXML = mXML + "/>"

        On Error Resume Next

        Dim strSQL As String

150     strSQL = "INSERT INTO CalendarEvents ("
160     strSQL = strSQL + "Subject, "
170     strSQL = strSQL + "Location, "
180     strSQL = strSQL + "Body, "
190     strSQL = strSQL + "StartDateTime,EndDateTime,"
200     strSQL = strSQL + "RecurrenceState,RemainderSoundFile,BusyStatus,IsAllDayEvent,ImportanceLevel,"
210     strSQL = strSQL + "IsPrivate,IsMeeting,LabelID,RExceptionEndTimeOrig,RExceptionStartTimeOrig,RecurrencePatternID,CustomPropertiesXMLData,IsReminder,ReminderMinutesBeforeStart)  "
220     strSQL = strSQL + "VALUES('" + Aitiologia + "','" + Topos + "','','"

230     strSQL = strSQL + Format(DATE1, "mm/dd/yyyy hh:mm") + "','" + Format(DateAdd("h", 1, DATE1), "mm/dd/yyyy hh:mm") + "',"
240     strSQL = strSQL + "0,'',2,0,0,"
250     strSQL = strSQL + "0,0,3,'','',0,'" + mXML + "',0,15)"

        '   strSQL = strSQL & MakeSQLstr(pEvent.CustomProperties.SaveToString) & "); " & vbCrLf
260     Gdb.Execute strSQL

        '<EhFooter>
        Exit Function

add_schedule_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.add_schedule " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.add_schedule " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function RightAlign(ByVal s As String, Size As Long) As String

        '<EhHeader>
        On Error GoTo RightAlign_Err

        '</EhHeader>

100     RightAlign = Right(Space(Size) + s, Size)

        '<EhFooter>
        Exit Function

RightAlign_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.RightAlign " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.RightAlign " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function CREATE_DATABASE(DB)

        '<EhHeader>
        On Error GoTo CREATE_DATABASE_Err

        '</EhHeader>

        Dim DefaultWorkspace As Workspace

        Dim CurrentDatabase  As Database, MyDatabase As Database

100     Set DefaultWorkspace = DBEngine.Workspaces(0)
        'Set CurrentDatabase = DefaultWorkspace.Databases(0)
        ' Create new, encrypted database.
110     Set MyDatabase = DefaultWorkspace.CREATEDATABASE(DB, dbLangGreek)
        ' Fill in new database.
        ' Set property to that of current database.
        'MyDatabase.QueryTimeout = CurrentDatabase.QueryTimeout
        '...
        ' Complete structure of MyDatabase, possibly using structure of
        ' CurrentDatabase.
        '...
120     MyDatabase.Close

        '<EhFooter>
        Exit Function

CREATE_DATABASE_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.CREATE_DATABASE " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.CREATE_DATABASE " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Sub BALE_TSONTES(TYPOS As String, TIMH As String)

        'ΔΗΜΙΟΥΡΓΕΙ ΣΕΙΡΕΣ ΜΕ ΠΕΡΙΓΡΑΦΗ ".." ΓΙΑ ΝΑ ΜΗΝ ΕΧΕΙ ΚΕΝΑ ΣΤΗΝ ΑΡΙΘΜΗΣΗ
        '  ΚΑΙ ΜΑΣ ΧΑΛΑΕΙ ΤΙΣ ΚΑΤΗΓΟΡΙΕΣ
        '<EhHeader>
        On Error GoTo BALE_TSONTES_Err

        '</EhHeader>
        Exit Sub

        Dim R As New ADODB.Recordset, N As Long, SQL3 As String

        Dim k As Long

        ' δεν μας ενδιαφέρει η ΤΙΜΗ
100     If TIMH = "0" Then
110         R.Open "SELECT MAX(AYJON) FROM PINAKES WHERE TYPOS=" + TYPOS, Gdb, adOpenDynamic, adLockOptimistic
        Else
120         R.Open "SELECT MAX(AYJON) FROM PINAKES WHERE TYPOS=" + TYPOS + " AND TIMH=" + str(Val(TIMH)), Gdb, adOpenDynamic, adLockOptimistic
        End If

130     If IsNull(R(0)) Or R(0) = 0 Then Exit Sub
140     N = R(0)
150     R.Close

160     If N > 50 Then Exit Sub

170     For k = 1 To N

180         If TIMH = "0" Then
190             R.Open "SELECT * FROM PINAKES WHERE TYPOS=" + TYPOS + " AND AYJON=" + str(k), Gdb, adOpenDynamic, adLockOptimistic
            Else
200             R.Open "SELECT * FROM PINAKES WHERE TYPOS=" + TYPOS + " AND TIMH=" + TIMH + "  AND AYJON=" + str(k), Gdb, adOpenDynamic, adLockOptimistic
            End If

210         If R.EOF Then
220             SQL3 = "INSERT INTO PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (" + TYPOS + "," + str(k) + ",'..'," + TIMH + ")"
230             Gdb.Execute SQL3
            End If

240         R.Close
        Next

        '<EhFooter>
        Exit Sub

BALE_TSONTES_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.BALE_TSONTES " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.BALE_TSONTES " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub UPDATE_PINAKES(MTYPOS As Integer, OO As Object, MCOMBO As Object)

        'mtypos  αριθμος πίνακα
        'οο: to object που θέλω για το κεντράρισμα της οθόνης DataEntry
        'mcombo : το combo που θα γεμίσει με τις τιμές
        '<EhHeader>
        On Error GoTo UPDATE_PINAKES_Err

        '</EhHeader>
        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        'ΜΟΝΑΔΕΣ ΜΕΤΡΗΣΗΣ
        'sql οπου θα γίνεται το update ή το addnew
        '  f.SHOW
100     Load F

110     Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=" + str(MTYPOS)
120     F.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=" + str(MTYPOS)    '    sql

        '------------------------------------------------------------------------------------
        'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
130     F.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " & " FROM PINAKES WHERE TYPOS=" + str(MTYPOS) + " ORDER BY AYJON "    'ParamGrid   sql2

        '------------------------------------------------------------------------------------
        'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
140     F.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=" + str(MTYPOS) + " AND "    'ParamGrid              DelSQL

        '------------------------------------------------------------------------------------
        'ποιό update query θα δημιουργείται (προαιρετικά )
        'αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
        '                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
        ' DD = " UPDATE PINAKES SET TYPOS=" + Str(MTYPOS) + " where TYPOS IS NULL AND AYJON="" + t1   "
        ' AddNewScript
        ' ' DD = "Sub Main()" & vbCrLf & _
        '         " DIM mCon  " & vbCrLf & _
        '         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
        '         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
        '         " mCon.execute """ + DD + " " & vbCrLf & _
        '        "Set r=Nothing" & vbCrLf & _
        '        "Set mCon=Nothing" & vbCrLf & _
        '        "End Sub"

150     F.ParamGrid.TextMatrix(0, 6) = "TYPOS"    'EPIPLEON FIELD
160     F.ParamGrid.TextMatrix(0, 7) = MTYPOS    ' TIMH TOY EPIPLEON FIELD

        '------------------------------------------------------------------------------------
170     F.ParamGrid.TextMatrix(0, 4) = ""  'DD
180     F.ParamGrid.TextMatrix(1, 1) = "A/A"
190     F.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "
        'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
200     F.Labels(0).Width = F.Labels(0).Width + 1000    '* TextWidth("A")
210     F.Textboxes(0).Left = F.Textboxes(0).Left + 1000    ' * TextWidth("A")
        '------------------------------------------------------------------------------------

        'αν παει η φόρμα να φύγει από την οθόνη τότε να κεντραρεται δεξια με την mdiform1
220     If OO.Left + F.Width > MDIForm1.Left + MDIForm1.Width Then
230         F.Left = MDIForm1.Left + MDIForm1.Width - F.Width
        Else
240         F.Left = OO.Left
        End If

250     F.Top = OO.Top + MDIForm1.Top

260     F.SHOW 1
270     MCOMBO.Clear
280     R.Open "SELECT *FROM PINAKES WHERE TYPOS=" + str(MTYPOS) + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

290     Do While Not R.EOF
300        ' MCOMBO.AddItem Format(R!AYJON, "00") + R("PERIGRAFH")
            MCOMBO.AddItem R("PERIGRAFH")  ' na mhn prosuetei ton aritmo
310         R.MoveNext
        Loop

320     R.Close

330     MCOMBO.Refresh

        '<EhFooter>
        Exit Sub

UPDATE_PINAKES_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.UPDATE_PINAKES " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.UPDATE_PINAKES " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function IfNull(Value As Variant, Optional NullValue As Variant = "") As Variant

        '<EhHeader>
        On Error GoTo IfNull_Err

        '</EhHeader>
100     If IsNull(Value) Then
110         IfNull = NullValue
        Else

120         IfNull = Value
        End If

        '<EhFooter>
        Exit Function

IfNull_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.IfNull " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.IfNull " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Sub UPDATE2_PINAKES(MTYPOS As Integer, OO As Object)    ', MCOMBO As Object)

        'mtypos  αριθμος πίνακα
        'οο: to object που θέλω για το κεντράρισμα της οθόνης DataEntry
        'mcombo : το combo που θα γεμίσει με τις τιμές
        '<EhHeader>
        On Error GoTo UPDATE2_PINAKES_Err

        '</EhHeader>
        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        'ΜΟΝΑΔΕΣ ΜΕΤΡΗΣΗΣ
        'sql οπου θα γίνεται το update ή το addnew
        '  f.SHOW
100     Load F

110     Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=" + str(MTYPOS)
120     F.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=" + str(MTYPOS)    '    sql

        '------------------------------------------------------------------------------------
        'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
130     F.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH,TIMH " & " FROM PINAKES WHERE TYPOS=" + str(MTYPOS) + " ORDER BY AYJON "    'ParamGrid   sql2

        '------------------------------------------------------------------------------------
        'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
140     F.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=" + str(MTYPOS) + " AND "    'ParamGrid              DelSQL

        '------------------------------------------------------------------------------------
        'ποιό update query θα δημιουργείται (προαιρετικά )
        'αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
        '                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
        ' DD = " UPDATE PINAKES SET TYPOS=" + Str(MTYPOS) + " where TYPOS IS NULL AND AYJON="" + t1   "
        ' AddNewScript
        ' ' DD = "Sub Main()" & vbCrLf & _
        '         " DIM mCon  " & vbCrLf & _
        '         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
        '         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
        '         " mCon.execute """ + DD + " " & vbCrLf & _
        '        "Set r=Nothing" & vbCrLf & _
        '        "Set mCon=Nothing" & vbCrLf & _
        '        "End Sub"

150     F.ParamGrid.TextMatrix(0, 6) = "TYPOS"    'EPIPLEON FIELD
160     F.ParamGrid.TextMatrix(0, 7) = MTYPOS    ' TIMH TOY EPIPLEON FIELD

        '------------------------------------------------------------------------------------
170     F.ParamGrid.TextMatrix(0, 4) = ""  'DD
180     F.ParamGrid.TextMatrix(1, 1) = "A/A"
190     F.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "
200     F.ParamGrid.TextMatrix(3, 1) = "ΤΙMH "
        'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
210     F.Labels(0).Width = F.Labels(0).Width + 1000    '* TextWidth("A")
220     F.Textboxes(0).Left = F.Textboxes(0).Left + 1000    ' * TextWidth("A")
        '------------------------------------------------------------------------------------

        'αν παει η φόρμα να φύγει από την οθόνη τότε να κεντραρεται δεξια με την mdiform1
230     If OO.Left + F.Width > MDIForm1.Left + MDIForm1.Width Then
240         F.Left = MDIForm1.Left + MDIForm1.Width - F.Width
        Else
250         F.Left = OO.Left
        End If

260     F.Top = OO.Top + MDIForm1.Top
270     F.SHOW 1
        'MCOMBO.Clear
        ' R.Open "SELECT *FROM PINAKES WHERE TYPOS=" + Str(MTYPOS) + " ORDER BY AYJON", Gdb, adOpenDynamic, _
        '     adLockOptimistic
        'Do While Not R.EOF
        '  MCOMBO.AddItem Format(R!ayjon, "00") + R("PERIGRAFH")
        '  R.MoveNext
        'Loop
        'R.Close
        '
        'MCOMBO.Refresh

        '<EhFooter>
        Exit Sub

UPDATE2_PINAKES_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.UPDATE2_PINAKES " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.UPDATE2_PINAKES " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Function Olografos(j)

        '*******************************************************************
        '* φέρνει αριθμό και τον κάνω ολογράφως
        '<EhHeader>
        On Error GoTo Olografos_Err

        '</EhHeader>
        Dim ekat, xil, MON, s, nn

100     s = ""

110     ekat = Int(j / 1000000)

120     If ekat > 0 Then
130         If ekat = 1 Then
140             s = "ΕΝΑ EKATOMYΡIΟ "
            Else
150             s = Tria_Olografos(ekat) + "EKATOMYΡIA "
            End If
        End If

160     nn = j - ekat * 1000000
170     xil = Int(nn / 1000)

180     If xil > 0 Then
190         If xil = 1 Then
200             s = s + "ΧΙΛΙA "
            ElseIf xil Mod 10 = 3 Then
                s = s + Replace(Tria_Olografos(xil), "ΤΡIA", "ΤΡΕΙΣ") + "ΧΙΛΙΑΔΕΣ "
            ElseIf xil Mod 10 = 4 Then
                s = s + Replace(Tria_Olografos(xil), "ΤΕΣΣEΡA", "ΤΕΣΣΕΡΕΙΣ") + "ΧΙΛΙΑΔΕΣ "
            Else
210             s = s + Tria_Olografos(xil) + "ΧΙΛΙΑΔΕΣ "
            End If
        End If

220     nn = nn - xil * 1000
230     MON = Int(nn)

240     nn = nn - MON

250     If nn = 0 Then
260         s = s + Tria_Olografos(MON) + " ΕΥΡΩ"
        Else
270         s = s + Tria_Olografos(MON) + " ΕΥΡΩ & " + Tria_Olografos(Round(nn * 100, 0)) + " ΛΕΠΤΑ" ' LTrim(str(Round(nn * 100, 0))) + " ΛΕΠΤΑ"
        End If

280     Olografos = s

        '<EhFooter>
        Exit Function

Olografos_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.Olografos " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.Olografos " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function  'return s

Public Function Tria_Olografos(N)

        '"*******************************************************************
        '"* φέρνει τριψήφιο και το κάνω ολογράφως
        '<EhHeader>
        On Error GoTo Tria_Olografos_Err

        '</EhHeader>
        Dim ek, dek, MON, strOL, nn

100     ek = Int(N / 100)
110     strOL = ""

120     Select Case ek

            Case 1
130             strOL = "EKATO "

140         Case 2
150             strOL = "ΔΙΑΚΟΣΙA "

160         Case 3
170             strOL = "ΤΡΙΑΚΟΣΙA "

180         Case 4
190             strOL = "ΤΕΤΡΑΚΟΣΙA "

200         Case 5
210             strOL = "ΠΕΝΤΑΚΟΣΙA "

220         Case 6
230             strOL = "EΞΑΚΟΣΙA "

240         Case 7
250             strOL = "ΕΠΤΑΚΟΣΙA "

260         Case 8
270             strOL = "ΟΚΤΑΚΟΣΙA "

280         Case 9
290             strOL = "ΕΝΝΙΑΚΟΣΙA "
        End Select

300     nn = N - ek * 100   '&& px  578 - 500 = 78
310     dek = Int(nn / 10)

        Dim ENTEKA As Long

        Dim m11_12 As Long

320     m11_12 = 0

330     Select Case dek

            Case 1

340             If nn = 11 Then
350                 m11_12 = 1
360                 strOL = strOL + "ΕΝΤΕΚΑ "
370                 Tria_Olografos = strOL
380             ElseIf nn = 12 Then
390                 m11_12 = 1
400                 strOL = strOL + "ΔΩΔΕΚΑ "
410                 Tria_Olografos = strOL
                Else
420                 strOL = strOL + "ΔΕΚΑ "
                End If

430         Case 2
440             strOL = strOL + "ΕΙΚΟΣΙ "

450         Case 3
460             strOL = strOL + "ΤΡΙΑΝΤΑ "

470         Case 4
480             strOL = strOL + "ΣΑΡΑΝΤΑ "

490         Case 5
500             strOL = strOL + "ΠΕΝΗΝΤΑ "

510         Case 6
520             strOL = strOL + "ΕΞΗΝΤΑ "

530         Case 7
540             strOL = strOL + "ΕΒΔΟΜΗΝΤΑ "

550         Case 8
560             strOL = strOL + "ΟΓΔΟΝΤΑ "

570         Case 9
580             strOL = strOL + "ΕΝΕΝΗΝΤΑ "
        End Select

590     nn = nn - dek * 10  ' && px  78 - 70 = 8
600     MON = Int(nn)

610     Select Case MON

            Case 1

620             If m11_12 = 0 Then strOL = strOL + "ENA "

630         Case 2

640             If m11_12 = 0 Then strOL = strOL + "ΔΥΟ "

650         Case 3
660             strOL = strOL + "ΤΡIA "

670         Case 4
680             strOL = strOL + "ΤΕΣΣEΡA "

690         Case 5
700             strOL = strOL + "ΠΕΝΤΕ "

710         Case 6
720             strOL = strOL + "ΕΞΙ "

730         Case 7
740             strOL = strOL + "ΕΠΤΑ "

750         Case 8
760             strOL = strOL + "ΟΚΤΩ "

770         Case 9
780             strOL = strOL + "ΕΝΝΕΑ "
        End Select

790     Tria_Olografos = strOL

        '<EhFooter>
        Exit Function

Tria_Olografos_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.Tria_Olografos " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.Tria_Olografos " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function  '"return strOL

'Sub openXLS(ByVal file As String, ByRef xlsheet As Worksheet, ByRef xl As Application)
''Dim xl As New excel.Application
''Dim xlsheet As excel.Worksheet
'
'    Dim xlwbook As excel.workbook
'
'
'
'
'
'
'    'Dim xl As New excel.Application
'    'Dim xlwbook As New xl.workbook
'    'Dim xlsheet As excel.Worksheet
'
'    Set xlwbook = xl.Workbooks.Open(file)
'    Set xlsheet = xlwbook.Sheets.Item(1)
'
'    Dim ko As String
'
'    ko = Trim(xlsheet.cells(1, 1))    ' row,col
'
'    'xl.Quit
'    'Set xlwbook = Nothing
'    'Set xl = Nothing
'
'
'End Sub
'Dim xl As New excel.Application
'Dim xlsheet As excel.Worksheet
'Dim xlwbook As excel.workbook
'

'Set xlwbook = xl.Workbooks.Open(Text1.Text)
'Set xlsheet = xlwbook.Sheets.Item(1)
'ko = Trim(xlsheet.cells(mRow, KOD_COLUMN)) '  ' ΚΩΔΙΚΟς ΤΟΥ ΕΙΔΟΥΣ  EXCEL

'xl.Quit
'Set xlwbook = Nothing
'Set xl = Nothing

Function gfind_parastat(par As String, ReadWrite As Integer) As Long

        'ReadWrite=0   Βρίσκει τον αριθμο παραστατικού          -------------------------
        'ReadWrite=1   Aποθηκεύει τον αριθμο παραστατικού + 1   -------------------------
        'par ειναι το είδος του παραστατικου  π.χ. par="T"
        '<EhHeader>
        On Error GoTo gfind_parastat_Err

        '</EhHeader>

        Dim k As Integer

        Dim X, Bohuhtiko As String

        On Error GoTo MHNYMA

        'Dim db As Database
        Dim R As New ADODB.Recordset

100     R.Open "SELECT * FROM PARASTAT WHERE EIDOS='" + par + "'", Gdb, adOpenDynamic, adLockOptimistic

110     If IsNull(R("ARITMISI")) Then
120         X = 0
        Else
130         X = R("ARITMISI")
        End If

140     R.Close

        'MEM..MoveFirst
        'For k = 1 To 9: f_mfpa(k) = R("fpa" + LTrim(Str(k))): Next

150     If X = 0 Then
160         gfind_parastat = 0

            Exit Function

        End If

170     If ReadWrite = 0 Then
180         R.Open "SELECT * FROM ARITMISI WHERE ID=" + str(X), Gdb, adOpenDynamic, adLockOptimistic

190         If IsNull(R("ARITMISI")) Then
200             gfind_parastat = 0
            Else
210             gfind_parastat = R("ARITMISI") + 1
            End If

220         R.Close

        Else
230         Gdb.Execute "UPDATE ARITMISI SET ARITMISI=ARITMISI+1  WHERE  ID=" + str(X)

        End If

        Exit Function

MHNYMA:
240     HandleError "Par1:FIND_Parastat"

250     Resume Next

        '<EhFooter>
        Exit Function

gfind_parastat_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.gfind_parastat " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.gfind_parastat " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function DOT(ByVal SS As String)

        ' ΜΕΤΑΤΡΕΠΕΙ ΤΟΝ ΑΡΙΘΜΟ ΑΠΟ ΚΟΜΑ ΣΕ ΤΕΛΕΙΑ  Π.Χ.   "12,34"  ΣΕ "12.34"
        '<EhHeader>
        On Error GoTo DOT_Err

        '</EhHeader>
100     DOT = Replace(SS, ",", ".")

        '<EhFooter>
        Exit Function

DOT_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.DOT " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.DOT " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function BRES_YPOLPEL(ByVal meidos As String, ByVal MKOD As String)

        'ENHMERVSH ME PROHGOYMENO YPOLOIPO
        '----------------------------------
        '<EhHeader>
        On Error GoTo BRES_YPOLPEL_Err

        '</EhHeader>
        Dim TYP As Single, SQLT As String, rt As New ADODB.Recordset
    
100     SQLT = "SELECT KOD,"
110     SQLT = SQLT + "SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END )  -"
120     SQLT = SQLT + " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END )   AS YP"
130     SQLT = SQLT + " FROM EGG  where HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND EIDOS='" + meidos + "' AND KOD='" + MKOD + "' GROUP BY KOD"
    
140     rt.Open SQLT, Gdb, adOpenDynamic, adLockOptimistic
    
150     If rt.EOF Or IsNull(rt("YP")) Then
160         TYP = 0
        Else
170         TYP = rt("YP")
        End If
    
180     rt.Close
    
190     rt.Open "select AYP FROM PEL WHERE EIDOS='" + meidos + "' AND KOD='" + MKOD + "'", Gdb, adOpenDynamic, adLockOptimistic
    
200     If IsNull(rt("AYP")) Then
210         TYP = TYP
        Else
220         TYP = TYP + rt("AYP")
        End If
    
230     Gdb.Execute "UPDATE PEL SET TYP=" + DOT(str(TYP)) + " WHERE  EIDOS='" + meidos + "' AND KOD='" + MKOD + "'"
    
240     rt.Close
    
250     BRES_YPOLPEL = TYP
  
        '<EhFooter>
        Exit Function

BRES_YPOLPEL_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.BRES_YPOLPEL " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.BRES_YPOLPEL " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function GDBEXE(ByVal PROC As String, ByVal Entolh As String)

  On Error GoTo GDBEXE_Err

    Dim k As Long

10    Gdb.Execute Entolh, k

    'On Error Resume Next

12    Gdb.Execute "INSERT INTO LOGGING (PROCED,ENTOLH,HME) VALUES ('" + PROC + "','" + Replace(Entolh, "'", "''") + "',GETDATE())"
    GDBEXE = k
    Exit Function
    
GDBEXE_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.BRES_YPOLPEL " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
         On Error Resume Next
        SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.GDBEXE_Err " & "at line " & Erl
        SAVE_ERROR Entolh

        Resume Next
    
    

End Function

Sub SAVE_ERROR(COMMENT)

    On Error Resume Next

    'SAVE_ERROR Err.Description & " in Project1.Form1.cmdCommand2_Click " & " at line " & Erl
    Dim F As Integer

    F = FreeFile
    Open "C:\MERCVB\ERR.TXT" For Append As #F
    Write #F, Format(Now, "DD/MM/YYYY HH:MM") + COMMENT

    Close #F

End Sub

Function TABLE_EXIST(ByVal TABLE_NAME As String) As Boolean

    'ΕΛΕΓΧΕΙ ΣΤΗΝ ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ ΑΝ ΥΠΑΡΧΕΙ ΕΝΑ TABLE
    Dim R As New ADODB.Recordset

    R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = '" + TABLE_NAME + "'", Gdb, adOpenDynamic, adLockOptimistic

    If R(0) = 0 Then
        TABLE_EXIST = False
    Else
        TABLE_EXIST = True
    End If
     
End Function

Sub enter_to_tab(ByRef KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0    'suppress the beep
        keybd_event VK_TAB, 0, 0, 0    'send a tab
    End If

End Sub






Function gDot(ByVal a As String) As String
   'ΜΕΤΑΤΡΕΠΕΙ ΤΟ "12,13" se  "12.13"
   
   
       a = Replace(a, ",", ".")
   
   
   gDot = a
   
   
   
   
   End Function






Function gVal(ByVal a As String) As Double
'ΜΕΤΑΤΡΕΠΕΙ ΤΟ "12,13" ή το "12.13"   ΣΕ 12.13
        '<EhHeader>
        On Error GoTo gVal_Err

        '</EhHeader>
        Dim KOMA
Dim xil
Dim teleia
Dim a2 As String
KOMA = InStr(a, ",")
teleia = InStr(a, ".")

a = Replace(a, "*", "")

If KOMA = 0 Then
    gVal = Val(a)
Else
  
  
   If teleia = 0 Then
      gVal = Val(Replace(a, ",", "."))
   Else
      
      'εχω και κομμα και τελεια  π.χ.   1.011.000,12  ή 1,011,000.12
      
      If teleia > KOMA Then  'aggliko systhma p.x. 1,011,000.12
         gVal = Val(Replace(a, ",", ""))
      Else 'ellhniko  1.011.000,12
         a2 = Replace(a, ".", "")
         gVal = Val(Replace(a2, ",", "."))
      End If
   
   End If


End If

        ' KOMA = InStr(A, ","): If KOMA > 0 Then Mid$(A, KOMA, 1) = "."
100    ' a = Replace(a, "*", "")
110    ' a = Replace(a, ",", ".")
120   '  gVal = Val(a) ' τελικο ποσο που ζητάω

        '<EhFooter>
        Exit Function

gVal_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.gVal " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in kentriko.gVal " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Sub load_Combo(MCOMBO As ComboBox, ByVal query As String)
'ΠΡΩΤΟ ΠΕΔΙΟ ΤΟ ΟΝΟΜΑ ΚΑΙ ΔΕΥΤΕΡΟ ΤΟ ID ΠΟΥ ΑΠΟΘΗΚΕΥΕΤΑΙ
' ΣΤΟ ITEMDATA
 Dim R As New ADODB.Recordset
 R.Open query, Gdb, adOpenDynamic, adLockOptimistic
 Dim k As Integer
 k = 0
 Do While Not R.EOF
    MCOMBO.AddItem R(0)
    MCOMBO.ItemData(k) = R(1)
    k = k + 1
    R.MoveNext
 Loop
R.Close
End Sub


Public Sub UPDATE_YPOLOIPA3(mTable As String, List11 As ListBox, APO As DTPicker, EOS As DTPicker)
        'into DOKEGGT1
        'B = "CREATE VIEW dbo.[EIDT3]" _

        '<EhHeader>
        On Error GoTo UPDATE_YPOLOIPA3_Err

        '</EhHeader>

        Dim a As String

        Dim B As String

        Dim D As Long

        'Dim krit As String

        ' ΒΡΙΣΚΩ ΜΙΑ ΦΟΡΑ ΤΑ ΥΠΟΛΟΙΠΑ ΑΝΑ ΑΠΟΘΗΚΗ

        On Error Resume Next

        ' On Error GoTo UPDATE_YPOLOIPA3_Err

        '----------------------------------------------------------------------------------------
        'ΠΑΙΡΝΕΙ ΜΟΝΟ ΑΠΟ ΤΟ ΤΟ ΛΙΣΤΒΟΧ ΤΙΣ ΚΙΝΗΣΕΙΣ
        Dim pol

        Dim k

100     pol = ""

110     For k = 0 To List11.ListCount - 1

120         If List11.Selected(k) = True Then
130             pol = pol + LTrim(str(Asc(Right(List11.List(k), 1)))) + ","
            End If

        Next

140     If Len(pol) - 1 > 0 Then
150         pol = Left(pol, Len(pol) - 1)
        Else
160         pol = "' '"
        End If

        '----------------------------------------------------------------------------------------

170     Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

180     a = "SELECT KODE," & " SUM(CASE APOT  WHEN 1  THEN isnull(XRE,0) ELSE 0  END ) AS S1X,"
        a = a + " SUM(CASE APOT  WHEN 1  THEN isnull(PIS,0) ELSE 0  END ) AS S1P,"
        a = a + " SUM(CASE APOT  WHEN 2  THEN isnull(XRE,0) ELSE 0  END ) AS S2X," & " SUM(CASE APOT  WHEN 2  THEN isnull(PIS,0) ELSE 0  END ) AS S2P,"
        a = a + " SUM(CASE APOT  WHEN 3  THEN isnull(XRE,0) ELSE 0  END ) AS S3X," & " SUM(CASE APOT  WHEN 3  THEN isnull(PIS,0) ELSE 0  END ) AS S3P,"
        a = a + " SUM(CASE APOT  WHEN 4  THEN isnull(XRE,0) ELSE 0  END ) AS S4X," & " SUM(CASE APOT  WHEN 4  THEN isnull(PIS,0) ELSE 0  END ) AS S4P  "
190     a = a + " INTO DOKEGGT1 FROM EGGTIM  "
200     a = a + "where HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "'  GROUP BY KODE"   'AND ascii(left(ATIM,1)) in (" + pol + ")

        Dim TT As Long

210     TT = GetCurrentTime()
220     Gdb.Execute a, D

        'MsgBox GetCurrentTime() - TT

230     If D = 0 Then
240         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"

            
            Exit Sub

        End If

250     Gdb.Execute "UPDATE " + mTable + " SET POS01=0,POS02=0,POS03=0,POS04=0,POS=0", D

        'b = "CREATE VIEW dbo.[EIDT3]" _
        '& "AS " _
        '& " SELECT     dbo.EID.*, (dbo.DOKEGGT1.S1X-dbo.DOKEGGT1.S1P) AS EXPR1, (dbo.DOKEGGT1.S2X-dbo.DOKEGGT1.S2P) AS EXPR2 " _
        '& " FROM         dbo.DOKEGGT1 INNER JOIN " _
        '& " dbo.EID ON dbo.DOKEGGT1.KODE = dbo.EID.KOD"
        'Gdb.Execute b, d

260     Gdb.Execute "update " + mTable + " SET" & " POS01=(CASE WHEN S1X IS NULL THEN 0 ELSE S1X END) - (CASE WHEN S1P IS NULL THEN 0 ELSE S1P END)," & " POS02=(CASE WHEN S2X IS NULL THEN 0 ELSE S2X END) - (CASE WHEN S2P IS NULL THEN 0 ELSE S2P END)," & " POS03=(CASE WHEN S3X IS NULL THEN 0 ELSE S3X END) - (CASE WHEN S3P IS NULL THEN 0 ELSE S3P END)," & " POS04=(CASE WHEN S4X IS NULL THEN 0 ELSE S4X END) - (CASE WHEN S4P IS NULL THEN 0 ELSE S4P END)" & " FROM EID INNER JOIN DOKEGGT1 ON EID.KOD=DOKEGGT1.KODE ", D

270     Gdb.Execute "UPDATE " + mTable + " SET  POS = POS01 + POS02+POS03+POS04", D

280     D = 0
        'Gdb.Execute "update EIDT3 SET POS01=EXPR1,POS02=EXPR2,POS=EXPR1+EXPR2 ", d

290     D = D + 1

        '<EhFooter>
        Exit Sub

UPDATE_YPOLOIPA3_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.UPDATE_YPOLOIPA3 " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.UPDATE_YPOLOIPA3 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Function FolderExists(sFullPath As String) As Boolean
Dim myFSO As Object
Set myFSO = CreateObject("Scripting.FileSystemObject")
FolderExists = myFSO.FolderExists(sFullPath)
End Function


Public Function GGET_CVALUE(query As String) As String
   Dim R As New ADODB.Recordset
   
   On Error Resume Next
   R.Open query, Gdb, adOpenDynamic, adLockOptimistic
   
    GGET_CVALUE = ""
   If Not R.EOF Then
      GGET_CVALUE = R(0)
   End If
   R.Close
   
End Function

Public Function GGET_DVALUE(query As String) As Date
   Dim R As New ADODB.Recordset
   
   On Error Resume Next
   R.Open query, Gdb, adOpenDynamic, adLockOptimistic
   
    GGET_DVALUE = Null
   If Not R.EOF Then
      GGET_DVALUE = R(0)
   End If
   R.Close
   
End Function

Public Function GGET_NVALUE(query As String) As Single
   Dim R As New ADODB.Recordset
   
   On Error Resume Next
   R.Open query, Gdb, adOpenDynamic, adLockOptimistic
   
    GGET_NVALUE = 0
   If Not R.EOF Then
      GGET_NVALUE = R(0)
   End If
   
   
   
   R.Close
   
End Function


Public Function HttpGetRequest(URL As String) As DOMDocument
   ' Dim req As XMLHTTP60
   ' Set req = New XMLHTTP60
    Dim REQ As XMLHTTP40
    Set REQ = New XMLHTTP40
    
    REQ.Open "GET", URL, False
    REQ.send ""

    Dim resp As DOMDocument
    If REQ.responseText <> vbNullString Then
        Set resp = New DOMDocument40
        resp.loadXML REQ.responseText
    Else
        Set resp = REQ.responseXML
    End If
    Set HttpGetRequest = resp
End Function




Public Sub guploadAzure(ByVal id_num As Long, FILENAME As String, Size As Long)





    Dim URL As String

        URL = gget_url(id_num, FILENAME, Size)

128     Dim upfilename As String: upfilename = "C:\pdf\" + FILENAME
            
            
           ' URL = get_url(id_num, FILENAME, Size)
            
            
        On Error Resume Next
           
130     gUploadAfIle URL, upfilename

132     '  VerifyUpload (id_num)

        Dim k As Integer

        For k = 1 To 4
           Dim MCC As String
            MCC = True ' VerifyUpload(id_num)
            
           ' gUploadAfIle URL, upfilename
            
            If Len(MCC) = "True" Then
                'continue
                Exit For
            Else

                k = k ' continue

            End If
            
            
           ' MCC = VerifyUpload(id_num)
            

            'url = get_url(id_num, FILENAME, Size)
            
           ' url = get_url(id_num, FILENAME, Size)
            
            gUploadAfIle URL, upfilename
         
        Next

134    ' UploadAfIle URL, upfilename


'download (id_num)




















End Sub


Function gget_url(ByVal id_num As Long, FILENAME As String, Size As Long) As String
'======================================================================================
  
        Dim ResultJSON  As String
            
        Dim jsonContent As String
100         jsonContent = ResultJSON ' " {""externa

        Dim REQ As WinHttp.WinHttpRequest
102     Set REQ = New WinHttp.WinHttpRequest
            
104     Dim URI As String: URI = gfURL1 + "api/uploadFileRequest" '?Bearer=" + gf_Bearer /api/
        Dim Q   As String
            
      
106     Dim URL As String: URL = ""
    On Error Resume Next
    
108     ResultJSON = "{  ""externalSystemId"":""" + LTrim(Trim(str(id_num))) + """ , ""filename"":""" + "C:\\PDF\\" + FILENAME + """, ""fileSize"": " + " " + LTrim(str(Size)) + "}   "  'debug
      
110     With REQ
112         .Open "POST", URI, async:=False
114         .setRequestHeader "Content-Type", "application/hal+json"
116         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
118         .setRequestHeader "Authorization", "Bearer " + gf_Bearer
120         .send (ResultJSON)
        End With
     
            Dim p

122         Dim sInputJson As String: ResultJSON = REQ.responseText
           
 
124        ' Set p = JSON.parse(sInputJson)
126        ' url = p.Item("uploadUrl")  ' Split(Split(sInputJson, ":")(1), ",")(0)

   'get_url = url
   
'   Set p = JSON.parse(sInputJson)
'   url = p.Item("uploadUrl")  ' Split(Split(sInputJson, ":")(1), ",")(0)
'   get_url = url
   
   
   Set p = JSON.parse(ResultJSON)
   URL = p.Item("uploadUrl")  ' Split(Split(sInputJson, ":")(1), ",")(0)
  ' ResultJSON = url
   
   
   gget_url = URL
   
End Function






Public Sub gUploadAfIle(surl As String, sFileName As String)
        Dim adoStream As Object
100     Set adoStream = CreateObject("ADODB.Stream")
102     adoStream.Mode = 3          ' read write
104     adoStream.Type = 1          ' adTypeBinary
106     adoStream.Open
108     adoStream.LoadFromFile (sFileName)
110     With CreateObject("Microsoft.XMLHTTP")
112         adoStream.Position = 0
114         .Open "PUT", surl, False
116         .setRequestHeader "Content-Length", "0" 'this is not a must
118         .setRequestHeader "x-ms-blob-type", "BlockBlob"
120         .send adoStream.Read(adoStream.Size)
        End With
122     Set adoStream = Nothing
End Sub
       
Public Sub gVerifyUpload(mexternalSystemId As Long)
100     Dim ResultJSON As String: ResultJSON = "{    ""externalSystemId"": """ + " " + LTrim(str(mexternalSystemId)) + """ }" ' JsonConvert.SerializeObject(SOU).toString

102     Dim jsonContent As String: jsonContent = ResultJSON '

        Dim REQ         As WinHttp.WinHttpRequest
104     Set REQ = New WinHttp.WinHttpRequest
106     Dim URI: URI = gfURL1 + "api/Finalize" '?Bearer=" + gf_Bearer /api/

108     Dim URL As String: URL = ""

110     With REQ
112         .Open "POST", URI, async:=False
114         .setRequestHeader "Content-Type", "application/hal+json"
116         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
118         .setRequestHeader "Authorization", "Bearer " + gf_Bearer

120         .send (jsonContent) '"{    ""externalSystemId"": """ + mexternalSystemId + """ }")
        End With
        
122     Dim sInputJson As String: sInputJson = REQ.responseText
        
End Sub

Public Sub gdownload(mexternalSystemId As Long)

100     Dim ResultJSON  As String: ResultJSON = "{    ""externalSystemId"": """ + " " + LTrim(str(mexternalSystemId)) + """ }"  ' JsonConvert.SerializeObject(SOU).toString
102     Dim jsonContent As String: jsonContent = ResultJSON '

        Dim REQ         As WinHttp.WinHttpRequest
104     Set REQ = New WinHttp.WinHttpRequest
106     Dim URI: URI = gfURL1 + "api/getFiles" '?Bearer=" + gf_Bearer /api/
108     Dim URL As String: URL = ""
110     With REQ
112         .Open "POST", URI, async:=False
114         .setRequestHeader "Content-Type", "application/hal+json"
116         .setRequestHeader "Accept", "text/*, application/hal+json, application/json"
118         .setRequestHeader "Authorization", "Bearer " + gf_Bearer
120         .send (jsonContent) '"{    ""externalSystemId"": """ + mexternalSystemId + """ }")
                
        End With
            
122     Dim sInputJson As String: sInputJson = REQ.responseText
            
End Sub


    ' ExportReportToPDF(ByVal FILENAME As String, ByVal ReportTitle As String, pdf_filename As String) As Long
Public Function gExport2ReportToPDF(ByVal FILENAMErpt As String, _
                            ByVal ReportTitle As String, _
                            pdf_filename As String, F_EKTYPOTHS, CrystalReport1 As CrystalReport) As Long

        On Error GoTo lab999

        ' Dim XXX2 As Printer
        '   Dim pdfExists As Integer
        '   pdfExists = 0
        '   Dim printerName As String
        '                  For Each XXX2 In Printers
        '
        '                     If InStr(UCase(XXX2.DeviceName), F_PDFNAME) > 0 Then
        '                            ' Set printer as system default.
        '                            printerName = XXX2.DeviceName
        '                            Set Printer = XXX2
        '                             pdfExists = 1
        '                            ' Stop looking for a printer.
        '                            Exit For
        '
        '                        End If
        '
        '                    Next
                    
        ' ΚΡΑΤΑΩ ΤΗΝ ΡΥΘΜΙΣΗ ΓΙΑ ΝΑ ΜΗΝ ΤΗΝ ΧΑΛΑΣΩ
100     Dim M_F_EKTYPOTHS As String: M_F_EKTYPOTHS = F_EKTYPOTHS
102     F_EKTYPOTHS = "MERCURYPDF"
104     CrystalReport1.ReportFileName = FILENAMErpt 'F_FORMA1
106     Kill "c:\pdf\pdf.pdf"
108     Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
110     MILSEC 1000
112     FileCopy "c:\pdf\pdf.pdf", "c:\pdf\" + pdf_filename

114     gExport2ReportToPDF = FileLen("c:\pdf\" + pdf_filename)

        'ΕΠΑΝΑΦΕΡΩ ΤΗΝ ΡΥΘΜΙΣΗ
116     F_EKTYPOTHS = M_F_EKTYPOTHS

        '  'pdf_filename = pdf_filename
        '        Dim crxApp As CRAXdDRT.Application
        '        Dim ReportObject As CRAXdDRT.Report
        '       Set crxApp = New CRAXdDRT.Application
        '        'ReportObject = crxApp.OpenReport(FILENAME)
        '
        '        Set ReportObject = crxApp.OpenReport(FILENAME)
        '        Dim objExportOptions As CRAXdDRT.ExportOptions
        '        ReportObject.ReportTitle = ReportTitle
        '        With ReportObject
        '            .EnableParameterPrompting = False
        '            .MorePrintEngineErrorMessages = True
        '        End With
        '        ' objExportOptions = ReportObject.ExportOptions
        '        Set objExportOptions = ReportObject.ExportOptions
        '        With objExportOptions
        '            .DestinationType = CRAXdDRT.CRExportDestinationType.crEDTDiskFile '   CRAXDDRT.crEDTDiskFile
        '            .DiskFileName = "c:\pdf\" + pdf_filename ' FILENAME
        '            .FormatType = CRAXdDRT.CRExportFormatType.crEFTPortableDocFormat 'crEFTPortableDocFormat
        '            .PDFExportAllPages = True
        '        End With
        '        ReportObject.Export False ' True
        '        'Dim myFile As New FileInfo(pdf_filename)
        '        ' Dim sizeInBytes As Long =
        '        Export2ReportToPDF = 1 ' myFile.Length
        Exit Function
        ' MsgBox "δεν αποθηκευτηκε"
lab999:
        'ReportObject.Export True ' χειροκινητα
118     Resume Next

End Function





