Attribute VB_Name = "Module2"
Option Explicit

'Sub all_year()
''**************** ·ÎÎ·„ﬁ ˜ÒÔÌÈ‹Ú ****************
'    Dim directory
'    '   * ·ÌÙÈ„Ò‹ˆ˘ Ù· ·Ò˜Âﬂ· ÛÙÁÌ ·ÎÈ‹ ˜ÒÔÌÈ‹
'    '
'    '        * ÂÛ˘ÙÂÒÈÍ¸ backup
'
'    'dum = Shell("c:\mercvb\pkzip.exe  " + gDir + "dbf *.db*", vbMinimizedFocus)
'    '
'    'directory = Right(Str(Year(Now), 4), 2)
'    'rmdirdel '+ directory , 0 , '' ,'' )
'    'mkdir
'    'FileCopy "*.dbf"   ' + directory , 0 , '' ,'' )
'    '        swpruncmd ( 'copy  *.ntx  ' + directory , 0 , '' ,'' )
'    '
'    '
'    '        swpruncmd ( 'cd ' + directory , 0 , '' ,'' )
'    '
'    '        if right(  Trim( curdir() )  , 2 )=directory  // OK
'    '        Else
'    '           wait 'Aƒ’Õ¡‘« « ¡ÀÀ¡√« ◊—œÕ…¡”'
'    '           quit
'    '        End If
'    '
'    '
'    '
'    '   do open_data with .t.,1,"KINHSEIS"
'    '    ZAP
'    '
'    '
'    '   do open_data with .t.,1,"PEL"
'    '    go top
'    '    repl all artim with 0,sumtim with 0,xresyn with 0,pissyn with 0,;
'    '             ayp with typ,xremhn with 0,pismhn with 0,aeg with 0
'    '   do open_data with .t.,1,"EGG"
'    '    ZAP
'    '
'    '    Close
'    '
'    '   do open_data with .t.,1,'TIM'
'    '     ZAP
'    '     Close
'    '   do open_data with .t.,1,'EGGTIM'
'    '     ZAP
'    '     Close
'    '
'    '    *** ----------- ÏÁ‰ÂÌﬂÊÂÈ ÙÁÌ ·ÔËﬁÍÁ ------------
'    '    do open_data with .t.,1,'EID'
'    '    repl all m01 with 0,m02 with 0,m03 with 0,m04 with 0,m05 with 0,m06 with 0,;
'    '             m07 with 0,m08 with 0,m09 with 0,m10 with 0,m11 with 0,m12 with 0,;
'    '             eissyn with 0,exsyn with 0,apos with pos,aeg with 0,ajiago with 0,;
'    '             ajipol with 0,ajiapo with if(abs(pos*xti)>99999999,0,pos*xti),;
'    '             g01 with 0,g02 with 0,g03 with 0,g04 with 0,g05 with 0,g06 with 0,;
'    '             g07 with 0,g08 with 0,g09 with 0,g10 with 0,g11 with 0,g12 with 0
'    '    Close
'    '
'    '   do open_data with .t.,1,'GRA'
'    '   SET DELE OFF
'    '   PACK
'    '   SET DELE ON
'    '   Close
'    'Return
'
'
'End Sub

Sub LOADPARAMETERS(ByRef FromBEGINING As Integer)

        '<EhHeader>
        On Error GoTo LOADPARAMETERS_Err

        '</EhHeader>
        Dim k As Integer

        Dim R As New ADODB.Recordset

        Dim mess

        On Error Resume Next

        'Exit Sub
        On Error GoTo LATOS

100     mess = "SELECT PINAKES"
110     R.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
        'Dim MESS

        'FPA
        
        
        For k = 1 To 9: g_Fpa(k) = 0: Next
        
        
120     k = 0

130     Do While Not R.EOF

140         If R("typos") = 1 Then
150             k = k + 1
160             g_Fpa(k) = R("TIMH")
            End If

170         R.MoveNext
            'Exit Do

        Loop

180     R.Close

190     mess = "SELECT MEM"
200     R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic

210     R.MoveFirst
220     gdirlog = R("C3")    ' = DIRLOG.Text

230     If Len(Dir(gdirlog, vbDirectory)) > 2 Then
        Else
240         gdirlog = ""
        End If

250     R.MoveNext

260     ' gEnarjh = R("HME")
270     R.MoveNext
280     '  gLhjh = R("HME")
290     '   gLHJH1 = DateAdd("D", 1, gLhjh)

300     R.Close

        Dim L, MEM As New ADODB.Recordset

310     mess = "SELECT PARASTAT"
320     MEM.Open "SELECT *FROM PARASTAT ", Gdb, adOpenDynamic, adLockOptimistic

330     MEM.MoveFirst

340     Do While Not MEM.EOF

'350         For k = 0 To MEM.FIELDS.Count - 1
'
'360             If IsNull(MEM.FIELDS(k)) Then
'
'370                 If MEM.FIELDS(k).Type = 8 Then    ' DATE
'380                     MEM.FIELDS(k) = Now
'390                 ElseIf MEM.FIELDS(k).Type = 3 Then    ' INTEGER
'400                     MEM.FIELDS(k) = 0
'410                 ElseIf MEM.FIELDS(k).Type = 10 Then    ' STRING
'420                     MEM(MEM.FIELDS(k).Name) = " "
'
'                    Else
'                        ' Mem.Fields(k) = Mem.Fields(k).DefaultValue
'                    End If
'
'430                 MEM.Update
'                End If
'
'            Next

440         MEM.MoveNext

            Exit Do  ' TO EBALA GIATIARGOYSE STON ARISTOGLOY  31/10/2009

        Loop

450     MEM.Close

        Dim fld As ADODB.FIELD, Flds, OK

460     OK = 0
470     mess = "SELECT EID"
480     R.Open "SELECT TOP 1 * FROM EID", Gdb, adOpenForwardOnly, adLockReadOnly

        ' flds = R.FIELDS
490     For Each fld In R.FIELDS

500         If fld.Name = "KODERG" Then
510             OK = 1
            End If

        Next
    
520     R.Close

530     If OK = 0 Then Gdb.Execute "ALTER TABLE EID ADD KODERG CHAR(14)"

540     OK = 0
550     mess = "SELECT TIM"
560     R.Open "SELECT TOP 1 * FROM TIM", Gdb, adOpenForwardOnly, adLockReadOnly

        'flds = R.FIELDS
570     For Each fld In R.FIELDS

580         If fld.Name = "ID_NUM" Then
590             OK = 1
            End If

        Next

600     R.Close

610     If OK = 0 Then Gdb.Execute "ALTER TABLE TIM ADD  ID_NUM int IDENTITY(1,1)"

        'On Error GoTo ADD
        'R.Open "SELECT TOP 1  KODERG FROM EID", Gdb, adOpenForwardOnly, adLockReadOnly
        'On Error Resume Next
        'R.Close`

        'On Error GoTo ALTER
        'R.Open "SELECT TOP 1  ID_NUM FROM TIM", Gdb, adOpenForwardOnly, adLockReadOnly
        'On Error Resume Next
        'R.Close

        Exit Sub

        'ADD:

        'Gdb.Execute "ALTER TABLE EID ADD KODERG CHAR(14)"
        'Resume Next'

        'ALTER:
        'Gdb.Execute "ALTER TABLE TIM ADD  ID_NUM int IDENTITY(1,1)"
        'Resume Next

LATOS:
620     MsgBox "¡ƒ’Õ¡Ã…¡ ÷œ—‘Ÿ”«” –¡—¡Ã≈‘—ŸÕ " + mess
          If FromBEGINING = 1 Then
             Resume Next
          Else
             FromBEGINING = -1
             Exit Sub
          End If
          
          
630     End

        '<EhFooter>
        Exit Sub

LOADPARAMETERS_Err:
        '   'MsgBox Err.Description & vbCrLf & _
             "in ADOMERCNEW.Module2.LOADPARAMETERS " & _
             "at line " & Erl, _
             vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.Module2.LOADPARAMETERS " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function StrZero(ByVal L As Long, ByVal N As Integer) As String

        'dino   45,6=> 000045
        '<EhHeader>
        On Error GoTo StrZero_Err

        '</EhHeader>

100     StrZero = Right("00000000000000000000" + LTrim(str(L)), N)

        '<EhFooter>
        Exit Function

StrZero_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.Module2.StrZero " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.Module2.StrZero " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function
'Sub preload()
''gegg_xpi = 0
''gDRX_TO_EURO = 0
''gtimh_synal = 0
'    gxron_enar = Year(Date)
'    gmhn_enar = 1
'    ghme_orist = boy(Date) - 1
'    ' gsmatim = Space(7)
'    gBuff = "e"
'    'gCentury = 0
'    ' ginform = 1
'    'gprint_fonts = "space(1)"
'    ' g_selis = 1
'    gDIG_3 = 7
'    gDIG_4 = 8
'    '  geidos_paras = 2
'    'gSOMA = "”ŸÃ¡"
'    'gPROSOPO = "–—œ”Ÿ–œ"
'
'    'gs_Ueis = "": gs_Aeis = "": gs_Uejag = "": gs_Aejag = "": gs_Uaj_ejag = "": gs_Aaj_ejag = ""
'    'gs_Uxre = "": gs_Axre = "": gs_Upis = "": gs_Apis = ""
'
'    'gs_Ulian = ""
'    ' gjj_problem = 0
'    'gm_apografh = 0: gm_PROMHU = 0: gTO_BLEPO_PRIN_TYPOUEI = 1
'
'End Sub
'Function recn(ByVal R As Recordset) As Long
'    recn = R.AbsolutePosition
'End Function
'Function file(ByVal S As String) As Boolean
'    Dim S2
'    S2 = Dir(S)
'    If S2 = "" Then file = False Else file = True
'
'End Function
'Function boy(ByVal S As Date)
'    boy = CDate(Format("01/01/" + Str(Year(S))))
'End Function
'Function MaxD(ByVal d1 As Date, ByVal d2 As Date) As Date
'    If DateDiff("d", d1, d2) > 0 Then MaxD = d1 Else MaxD = d2
'End Function
'Function MaxN(ByVal n1 As Long, ByVal n2 As Long) As Long
'    If n1 > n2 Then MaxN = n1 Else MaxN = n2
'End Function
'Function MinD(ByVal d1 As Date, ByVal d2 As Date) As Date
'    If DateDiff("d", d1, d2) > 0 Then MinD = d1 Else MinD = d2
'End Function
'Function MinN(ByVal n1 As Long, ByVal n2 As Long) As Long
'    If n1 > n2 Then MinN = n2 Else MinN = n1
'End Function

