VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker eos 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   154271745
      CurrentDate     =   46198
   End
   Begin MSComCtl2.DTPicker apo 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   154271745
      CurrentDate     =   46198
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      Top             =   2880
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCommand1_Click()
cmdKOYKAKH
End Sub





Sub cmdKOYKAKH()

Dim gdb As New ADODB.Connection
Dim gConnect As String

10 Open "C:\MERCVB\MERCPATH.TXT" For Input As #1
20  Line Input #1, gDir
30   Line Input #1, gConnect
40 Close #1


gdb.Open Trim(gConnect)





'gdb.Open "DSN=MERCSQL"
        

 f_koyk177 = 177





    Dim sql As String

    sql = "SELECT        G.ID_NUM AS DocID, R.TITLOS AS DocDescr, R.SYNT_TITL AS DocCode, G.ATIM AS DocNum, SUBSTRING(G.ATIM, 2, 6) AS DocAA, FORMAT(G.HME, 'yyyy-MM-ddTHH:mm:ss.fffZ') AS DocDate, G.PROOD_AJ AS LineNum, P.KOD AS CustCode,"
    sql = sql + " P.EPO AS CustName, P.AFM AS CustAFM, P.EPA AS CustProfession, P.DOY AS CustDOY, P.XRVMA AS CustZipCode, P.DIE AS CustAddress, P.POL AS CustCity, '1' AS VatStatus, LEFT(T.TRP, 1) AS Payment, "
    sql = sql + "  T.TRP AS PaymentDesc, G.KODE AS EidosCode, E.ONO AS EidosDescription, G.MONA AS UNIT, G.POSO AS Qty, G.TIMM AS Price, G.POSO * G.TIMM AS NetValue, (CASE WHEN G.FPA = 1 THEN 13 WHEN G.FPA = 2 THEN 24 END) "
    sql = sql + "   AS FPAPRC, G.MIK_AJIA - G.KAU_AJIA AS VatValue, G.POSO * G.TIMM - G.KAU_AJIA AS DiscVal, 0 AS DiscVal2, CONVERT(CHAR(10), G.HME, 103) AS InsertDate, CONVERT(varCHAR(10), G.HME, 103) "
    sql = sql + "  + ' ' + T.ORA AS LastDate "
    sql = sql + " FROM            dbo.EGGTIM AS G INNER JOIN "
    sql = sql + "  dbo.PEL AS P ON G.EIDOS = P.EIDOS AND G.PELKOD = P.KOD INNER JOIN "
    sql = sql + "  dbo.TIM AS T ON G.ID_NUM = T.ID_NUM INNER JOIN "
    sql = sql + "  dbo.EID AS E ON G.KODE = E.KOD INNER JOIN "
    sql = sql + "  dbo.PARASTAT AS R ON LEFT(G.ATIM, 1) = R.EIDOS "
    sql = sql + "where R.MYEID='1.1' AND E.PROM LIKE '%ΚΟΥΚΑ%' AND G.HME>='" + Format(apo.Value, "MM/dd/yyyy") + "' AND G.HME<='" + Format(eos.Value, "MM/dd/yyyy") + "'"

    Dim Reqpos As WinHttp.WinHttpRequest

    Set Reqpos = New WinHttp.WinHttpRequest
    '

    Dim Q As String

    Dim R As New ADODB.Recordset
 ' Gdb.Execute "DROP TABLE IF EXISTS TEMP1010;"
 
 'Gdb.Execute "IF OBJECT_ID('dbo.TEMP1010', 'U') IS NOT NULL  DROP TABLE dbo.TEMP1010;"
' Gdb.Execute sql
  
    R.Open sql, gdb, adOpenDynamic, adLockOptimistic  '"SELECT * FROM TEMP1010"

    Q = "["

    Do While Not R.EOF
        If Len(Q) > 30 Then
            Q = Q + ",{"
         Else
            Q = Q + "{"
        End If
        Q = Q + " ""DocID"": """ + Str(R!DocID) + """,  "
        Q = Q + " ""DocDescr"": """ + R!DocDescr + """,  "
        '=====
        Q = Q + " ""DocCode"": """ + R!DocCode + """,  "
        Q = Q + " ""DocNum"": """ + R!Docnum + """,  "

        Q = Q + " ""DocAA"": """ + R!Docaa + """,  "
        Q = Q + " ""DocDate"": """ + Left(R!Docdate, 19) + """,  "

        Q = Q + " ""AlterDoc"": """ + R!DocCode + """,  " 'ebala to idio me doccode
        Q = Q + " ""LineNum"": " + Str(R!LINENUM) + ",  "

        Q = Q + " ""CustCode"": """ + R!custCode + """,  "
        Q = Q + " ""CustName"": """ + R!CustName + """,  "
        Q = Q + " ""CustAFM"": """ + R!custafm + """,  "
        Q = Q + " ""CustProfession"": """ + R!custprofession + """,  "

        Q = Q + " ""CustDOY"": """ + R!custdoy + """,  "
        Q = Q + " ""CustZipCode"": """ + R!custzipCode + """,  "
        Q = Q + " ""CustAddress"": """ + R!CustAddress + """,  "
        Q = Q + " ""CustCity"": """ + R!custCity + """,  "
        '-----------------------------------------------------------

        Q = Q + " ""VatStatus"": 1,  "
        Q = Q + " ""Payment"": """ + R!Payment + """,  "
    
        Q = Q + " ""PaymentDescr"": """ + R!PaymentDesc + """,  "
        Q = Q + " ""EidosCode"": """ + R!EidosCode + """,  "

        Q = Q + " ""EidosDescription"": """ + R!eidosdescription + """,  "
        Q = Q + " ""Unit"": """ + R!Unit + """,  "
        Q = Q + " ""Qty"":" + Str(R!QTY) + ",  "
        Q = Q + " ""Price"": " + Replace(Format(R!Price, "###0.00"), ",", ".") + ",  "

        Q = Q + " ""NetValue"":" + Replace(Format(R!NetValue, "###0.00"), ",", ".") + ",  "
        Q = Q + " ""FPAPrc"": """ + Str(nNull(R!FPAPRC)) + """,  "
        Q = Q + " ""VATValue"":" + Replace(Format(R!VATValue, "###0.00"), ",", ".") + ",  "
        Q = Q + " ""Discval"": " + Replace(Format(R!DISCVal, "###0.00"), ",", ".") + ",  "

        Q = Q + " ""DiscVal2"":" + Replace(Format(R!DISCVal2, "###0.00"), ",", ".") + ",  "
        Q = Q + " ""InsertDate"": """ + Left(R!Docdate, 19) + """,  "
        Q = Q + " ""LastUpdate"": """ + Left(R!Docdate, 19) + """,  "
        Q = Q + " ""IsActive"": 1,  "

        Q = Q + " ""OrderID"": null,  "
        Q = Q + " ""debugging"": ""p""  "

        Q = Q + " }  "
    
        DoEvents
        Me.Caption = R!Docnum
        R.MoveNext
    
    Loop

    Q = Q + " ] "
   ' Gdb.Execute "DROP TABLE TEMP1010"
'    Open "C:\MERCVB\KOYKAKIS.TXT" For Output As #3
 '   Print #3, Q
  '  Close #3

    If Len(Q) < 20 Then
        Me.Caption = "δεν υπαρχουν εγγραφες"
        Open "C:\MERCVB\KOYKAKIS.TXT" For Append As #3
        Print #3, Format(apo.Value, "dd/MM/yyyy") + Chr(13)
        Print #3, "ΔΕΝ ΥΠΑΡΧΟΥΝ"
        Close #3
        Exit Sub

    End If

    Dim uripos As String: uripos = "https://apps.koukakisfarm.gr/dc/dc.php"

    '
    With Reqpos
        .Open "POST", uripos, async:=False
      
        .setRequestHeader "accept", "application/json"
        .setRequestHeader "X-Client-Code", "ANTHOPOULOS"     ', "Bearer " + gf_Bearer
        .setRequestHeader "X-Client-Password", "Y7uaGuNqjWkGcQH5R2f5"
        .setRequestHeader "Content-Type", "application/json"
        ' qpos = " {  ""Bearer"":""" + gf_Bearer + """, " & ResultJSONpos & ",     ""externalSystemId"": """ + LTrim(str(SQLDT("ID_NUM"))) + """ }"
        .send (Q)
                        
        Me.Caption = "απαντηση : " + Chr(13) + Reqpos.responseText
                         
    End With
                        
    Dim posjson As String: posjson = Reqpos.responseText
    
         Open "C:\MERCVB\KOYKAKIS.TXT" For Append As #3
        Print #3, Format(apo.Value, "dd/MM/yyyy") + Chr(13)
        Print #3, posjson
        Close #3
        
    
    
    
    
                        
    Dim ser     As Object
                        
    Set ser = JSON.parse(posjson)

End Sub

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
       
        Resume Next

        '</EhFooter>

End Function


Private Sub Form_Load()
     apo = Now
     eos = Now + 1
     cmdKOYKAKH
     End
End Sub
