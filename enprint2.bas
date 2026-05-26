Attribute VB_Name = "Module5"
Option Explicit

Dim f_kodik As Recordset, f_tab(40)

Dim f_db As Database
Attribute f_db.VB_VarUserMemId = 1073741827

Dim F_437 As Integer

Dim fENARJH_APO_SEIRA
Attribute fENARJH_APO_SEIRA.VB_VarUserMemId = 1073741828

Dim F_TYPEHEADER As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SETTINGS_PROGID = "biopdf.PDFSettings"  'https://www.biopdf.com/download.php
Const UTIL_PROGID = "biopdf.PDFUtil"

Private Function PrinterIndex(ByVal printerName As String) As Integer  'https://www.biopdf.com/download.php
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If LCase(Printers(i).DeviceName) Like LCase(printerName) Then
            PrinterIndex = i
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function


  Public Function ExportReportToPDF(ByVal FILENAME As String, ByVal ReportTitle As String, pdf_filename As String) As Long

        pdf_filename = "C:\PDF\" + pdf_filename
        Dim crxApp As CRAXdDRT.Application
        Dim ReportObject As CRAXdDRT.Report
      Set crxApp = New CRAXdDRT.Application
       Set ReportObject = crxApp.OpenReport(FILENAME)




        Dim objExportOptions As CRAXdDRT.ExportOptions

        ReportObject.ReportTitle = ReportTitle

        With ReportObject
            .EnableParameterPrompting = False
            .MorePrintEngineErrorMessages = True
        End With

       Set objExportOptions = ReportObject.ExportOptions

        With objExportOptions
            .DestinationType = CRAXdDRT.CRExportDestinationType.crEDTDiskFile '   CRAXDDRT.crEDTDiskFile
            .DiskFileName = pdf_filename ' FILENAME
            .FormatType = CRAXdDRT.CRExportFormatType.crEFTPortableDocFormat 'crEFTPortableDocFormat
            .PDFExportAllPages = True
        End With

        ReportObject.Export (False)
       ' Dim myFile As New FileInfo(pdf_filename)
        ' Dim sizeInBytes As Long =
       ' ExportReportToPDF = myFile.Length

    End Function























'Public Sub ExportReportToPDF(ByVal ReportObject As String, ByVal FILENAME As String, ByVal ReportTitle As String)


'   Dim crxApp
'Dim crxRpt As CRAXdDRT.Report
'Set crxApp = New CRAXdDRT.Application
'
'Set crxRpt = crxApp.OpenReport(ReportObject, 1)
'
'    Dim objExportOptions As CRAXdDRT.ExportOptions
'
'    crxRpt.ReportTitle = ReportTitle
'
'    With crxRpt
'        .EnableParameterPrompting = False
'        .MorePrintEngineErrorMessages = True
'    End With
'
'    Set objExportOptions = crxRpt.ExportOptions
'
'    With objExportOptions
'        .DestinationType = crEDTDiskFile
'        .DiskFileName = FILENAME
'        .FormatType = crEFTPortableDocFormat
'        .PDFExportAllPages = True
'    End With
 
   ' crxRpt.Export
    
 
'End Function

Function print3_xar(ByVal sql As String, _
                    SUgm_str, _
                    ByVal EPIKEF As String, _
                    ByVal GROUPN As Integer)
        '<EhHeader>
        On Error GoTo print3_xar_Err
        '</EhHeader>

        '"*********************** P R I N T _ X A R **************-------------**********
        '"**************      κάνει παρουσίαση αρχείου στο χαρτί
        '"*** n=αριθμός fields που παρουσιάζονται
        '"*** synuhkh2 η συνθήκη για το  IF, sugm_str βλέπει που θα κάνει σούμες ,π.χ. "00100" κάνει σούμες στο 3ο field
        '"*** sum_pic  το picture γιά τις σούμες
        '"*** Ei,Fi,Pi  :Επικεφαλίδα παρουσίασης,Fields που παρουσιάζονται,Picture παρουσίασης
        '"** synt_eject:=0  αλλάζει σελίδα όταν μεταβάλλεται η στήλη
        Dim MFIELDS(120), mSYN

        Dim synt_eject
        Dim F_FILESAVE As String
           F_FILESAVE = FINDPARAMETROI(1, "EKTYPOTES", "F_FILESAVE", "C:\MERCVB\PRINT.CSV", "ΑΡΧΕΙΟ ΕΚΤΥΠΩΣΗΣ CSV ΑΡΧΕΙΟΥ ")
100     F_TYPEHEADER = Val(FINDPARAMETROI(1, "EKTYPOTES", "F_TYPEHEADER", "1", "1=ΤΥΠΩΝΟΝΤΑΙ ΕΠΙΚΕΦΑΛΙΔΕΣ ΑΡΧΕΙΟ "))

        Dim sumes, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), F(120)

        Dim mBSEIRA

        Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE

        Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4

        Dim aaP, aaP2, ektypoths

        Dim EPIK, xeirisths, PPF, epik3

        Dim Sthles, kw, PPD, SELIS

        Dim LSYN, AA, epik2, i As Integer

        Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject

        Dim m_sumes(120), SUMES0(120)

        Dim SS

        Dim m_ekt As Integer

        Dim DUM
        
        Dim gr As Integer
        
        
        
        
        
        
        
        
        

        'm_ekt = MsgBox("Ναι: Τυπώνω " + Chr(13) + "Οχι:Βλέπω αρχείο εκτύπωσης", vbYesNo)

102     utDialog.Left = Screen.Width / 2 - utDialog.Width / 2
104     utDialog.Top = Screen.Height / 2 - utDialog.Height / 2

106     utDialog.SHOW 1

108     DoEvents

110     m_ekt = MDIForm1.StatusBar1.Panels(6).Text

112     If m_ekt = 1 Then    'EKTYPOTHS
114         DUM = print4_xar(sql, SUgm_str, EPIKEF, GROUPN)

            Exit Function

        End If


116     If m_ekt = 2 Then
118          PRINT_TEXT sql, SUgm_str, EPIKEF, GROUPN
            Exit Function
        End If




120     If m_ekt = 5 Then
122          PRINT_CSV sql, SUgm_str, EPIKEF, GROUPN, 1, F_FILESAVE  ' "c:\mercvb\print.csv"
            Exit Function
        End If


        If m_ekt = 8 Then
        
             gr = MsgBox("Να γίνεται ομαδοποίηση στο 1ο πεδίο;", vbYesNo)

        If gr = vbYes Then
            GROUPN = 1
        Else
            GROUPN = 0
        End If
        
            print_CALC sql, SUgm_str, EPIKEF, GROUPN
            Exit Function
        End If








124     If m_ekt = 11 Then    'EKTYPOTHS
126      DUM = print4_xar(sql, SUgm_str, EPIKEF, 1)

            Exit Function

        End If
        
        
   
        If m_ekt = 12 Then    'pdf
            DUM = print4_pdf(sql, SUgm_str, EPIKEF, 1)
            Exit Function
        End If
             
        
        
        
        
        

128     If m_ekt = 0 Then

            Exit Function

        End If

        'If m_ekt = 2 Then
        'αυτο το sub (σημειωματάριο)
        '
        'End If

130     If m_ekt = 13 Then

          '  On Error GoTo print3_xar_Err
           gr = MsgBox("Να γίνεται ομαδοποίηση στο 1ο πεδίο;", vbYesNo)
          
          If gr = vbYes Then
            GROUPN = 1
          Else
            GROUPN = 0
          End If
  
132         DUM = print7_excel(sql, SUgm_str, EPIKEF, GROUPN)

            Exit Function

        End If

134     If m_ekt = 3 Then

           ' On Error GoTo print3_xar_Err
           
            gr = MsgBox("Να γίνεται ομαδοποίηση στο 1ο πεδίο;", vbYesNo)
          
          If gr = vbYes Then
            GROUPN = 1
          Else
            GROUPN = 0
          End If
           

136         DUM = print7_excel(sql, SUgm_str, EPIKEF, GROUPN)

            Exit Function

        End If

        '--------------------------------------------

138     If m_ekt = 4 Then
140         DUM = print_WORD(sql, SUgm_str, EPIKEF, GROUPN)

            Exit Function

        End If
        
        
        
        
        

142     If m_ekt = 0 Then

            'DUM = print_WORD(sql, SUgm_str, EPIKEF, GROUPN)
            Exit Function

        End If

144     PRINT_TEXT sql, SUgm_str, EPIKEF, GROUPN

        '<EhFooter>
        Exit Function

print3_xar_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.print3_xar " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function

Function print_CALC(sql As String, _
                      SUgm_str, _
                      ByVal EPIKEF As String, _
                      ByVal GROUPN As Integer)
        '<EhHeader>
        On Error GoTo print_calc_Err
        '</EhHeader>

        '"*********************** P R I N T _ X A R ************************
        '"**************      κάνει παρουσίαση αρχείου στο χαρτί
        '"*** n=αριθμός fields που παρουσιάζονται
        '"*** synuhkh2 η συνθήκη για το  IF, sugm_str βλέπει που θα κάνει σούμες ,π.χ. "00100" κάνει σούμες στο 3ο field
        '"*** sum_pic  το picture γιά τις σούμες
        '"*** Ei,Fi,Pi  :Επικεφαλίδα παρουσίασης,Fields που παρουσιάζονται,Picture παρουσίασης
        '"** synt_eject:=0  αλλάζει σελίδα όταν μεταβάλλεται η στήλη
        Dim MFIELDS(120), mSYN

        Dim synt_eject

        Dim sumes, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), F(120)

        Dim mBSEIRA

        Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE

        Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4

        Dim aaP, aaP2, ektypoths

        Dim EPIK, xeirisths, PPF, epik3

        Dim Sthles, kw, PPD, SELIS

        Dim LSYN, AA, epik2, i As Integer

        Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject

        Dim m_sumes(120), SUMES0(120)

        Dim SS

        Dim m_ekt As Integer

        Dim DUM

        Dim returnValue

        Dim f_excelPath
        
        
        
'-------------------------------------------------  calc ----------------------------------------
Dim oSM
Dim oDesk, oDoc As Object
Dim arg()
Dim ocell As Object
Dim oSheet As Object
Dim importDesc(0 To 2) As Object
Dim oImportrange As Object
Dim oSQL As Object
Dim oContext As Object
Set oSM = CreateObject("com.sun.star.ServiceManager")
Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
Set oDoc = oDesk.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg())
' Set oSheet = oDoc.Sheets.getByName("Φύλλο1")
Set oSheet = oDoc.getSheets().getByIndex(0)
oSM.createInstance ("com.sun.star.sheet.DatabaseImportDescriptor ")
oSM.createInstance ("com.sun.star.util.Ximportable")


' Call oSheet.getCellByPosition(στηλη , σειρα ).setformula(" δεδομενα ")
'            Call oSheet.getCellByPosition(J, i).setString("Row:" & i + 1 & ",Col:" & J + 1)


'For i = 0 To Val(txtRows.Text) - 1
'        For J = 0 To Val(txtCols.Text) - 1
'            Call oSheet.getCellByPosition(J, i).setString("Row:" & i + 1 & ",Col:" & J + 1)
'        Next J
'    Next i
'
'    oSheet.getCellRangeByPosition(0, 0, Val(txtCols.Text) - 1, 0).CharWeight = 150 'Bold
'    oSheet.getCellRangeByPosition(0, 1, Val(txtCols.Text) - 1, Val(txtRows.Text) - 1).CharColor = 16711680 ' color
'    oSheet.getCellRangeByPosition(0, 0, Val(txtCols.Text) - 1, 0).CellBackColor = RGB(0, 220, 240) 'cell back color
'    oSheet.getCellRangeByPosition(0, 0, Val(txtCols.Text) - 1, 0).CharHeight = 15 'Font size
'
'    oSheet.getCellRangeByPosition(0, 0, Val(txtCols.Text) - 1, Val(txtRows.Text) - 1).CharFontName = "Calibri" 'font name
'    oSheet.getCellRangeByPosition(0, 0, Val(txtCols.Text) - 1, Val(txtRows.Text) - 1).CharPosture = 2 'Italic
'    oSheet.getCellRangeByPosition(0, 0, Val(txtCols.Text) - 1, 0).CharUnderLine = 1 'UnderLine
'
'    '// Set Fixed width to all columns
'    Dim range As Object
'    Set range = oSheet.getCellRangeByPosition(0, 0, Val(txtCols.Text) - 1, 0)
'        Dim columns As Object
'        Set columns = range.getColumns()
'        For i = 0 To Val(txtCols.Text) - 1
'            columns.getByIndex(i).optimalWidth = True
'        Next
'        '//
'

 '// Set Fixed width to all columns
'    Dim range As Object
'    Set range = oSheet.getCellRangeByPosition(0, 0, Val(txtCols.Text) - 1, 0)
'        Dim columns As Object
'        Set columns = range.getColumns()
'        For i = 0 To Val(txtCols.Text) - 1
'            columns.getByIndex(i).optimalWidth = True
'        Next
'
'Dim savepath As String
'    savepath = App.Path & "\SampleCalcfile.ods"
'    savepath = ConvertToUrl(savepath)

'Call oDoc.storeToURL(SavePath, arg())
'        oDoc.Close (True)
'        Set oDoc = Nothing
'




'-------------------------------------------------  calc ----------------------------------------







        On Error Resume Next

        Dim F_XROMATA As Integer

104     F_XROMATA = Val(FINDPARAMETROI(1, "MDIFORM1", "F_XROMATA", "0", "ΧΡΩΜΑΤΑ ΣΤΟ EXCEL=1 ΟΧΙ=0"))

        '"MDIFORM1"

        '  If (MenuShow.Caption = "&Show") Then
        '    MenuShow.Caption = "&Hide"
106  '   workbook.Activate

108   '  Set myXL = workbook.ActiveSheet

        Dim FF As New UDialog

110     FF.SHOW
112     FF.OKButton.Visible = False
114     FF.List1.Visible = False

116     FF.CancelButton.Caption = "ΔΙΑΚΟΠΗ"
118     FF.CancelButton.Top = 120
120     FF.CancelButton.Left = 120
122     FF.CancelButton.Width = 2895
124     FF.CancelButton.Height = 495

        'FF.Top = 3000
        'FF.Left = 3000

126     FF.Width = 3210
128     FF.Height = 810

130     FF.Left = Screen.Width / 2 - FF.Width / 2
132     FF.Top = Screen.Height / 2 - FF.Height / 2

        'FF.Top = MDIForm1.Top + (MDIForm1.height) / 2 ' Command2.Top
        'FF.Left = MDIForm1.Left + (MDIForm1.width) / 2 ' Command2.Left

134     FF.Caption = "ΔΙΑΔΙΚΑΣΙΑ ΥΠΟΛΟΓΙΣΜΟΥ"

        ''------------------------ΠΕΤΑΕΙ ΟΛΟΚΛΗΡΟ ΤΟ ΑΡΧΕΙΟ ΑΛΛΑ ΚΟΛΑΕΙ ΣΤΗΝ [ΧΟΝΔ.ΤΙΜΗ]------------------------------
        'Dim db1 As Database
        'Dim FROM As Long
        'On Error Resume Next
        '  Kill "C:\EKTYP.XLS"
        '  DoEvents
        '  myXL.SaveAs "C:\EKTYP.XLS"
        '
        '  Call workbook.Close(False)
        '  excel.Quit
        '  Set excel = Nothing
        '
        'FROM = InStr(1, UCase(sql), "FROM")
        'Set db1 = OpenDatabase("", False, False, gConnect)
        'db1.Execute Left(sql, FROM - 1) + " " + " into 32 in 'c:\EKTYP.xls' 'Excel 8.0;' " + Trim(Mid(sql, FROM, 500))
        ''------------------------------------------------------

136     MDIForm1.MousePointer = vbHourglass
        'excel.Visible = True

        'f_excelPath = FindParametroi(1,"MDIFORM1", "f_excelPath", "C:\Program Files\Microsoft Office\OFFICE11", "Φάκελος Excel")
        ''C:\Program Files\Microsoft Office\OFFICE11
        'returnValue = Shell(f_excelPath + "\EXCEL.EXE", vbMaximizedFocus) ' vbMinimizedNoFocus) ' Run Microsoft Excel.
        '
        'Set myXL = GetObject("", "Excel.Sheet")

138     For k = 1 To 120: m_sumes(k) = 0: SUMES0(k) = 0: Next

        Dim N

        Dim F_T As New ADODB.Recordset

        Dim TT  As Long

140     TT = GetCurrentTime()

        'Dim recs As Long, fp As Long
        'fp = InStr(UCase(sql), "FROM")

        'F_T.Open "select count(*) " + Trim(Mid(sql, fp, 100)), Gdb, adOpenForwardOnly, adLockReadOnly
        'recs = F_T(0)
        'F_T.Close

142     F_T.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

144     N = F_T.FIELDS.Count

146     For k = 1 To N

148         If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then    'DATE
150             f_tab(k) = 2 + f_tab(k - 1) + 8
            Else

152             If F_T.FIELDS(k - 1).DefinedSize > 200 Then
154                 f_tab(k) = 2 + f_tab(k - 1) + 30
                Else
156                 f_tab(k) = 3 + f_tab(k - 1) + F_T.FIELDS(k - 1).DefinedSize
                End If
            End If

        Next

158     LSYN = f_tab(k - 1)

160     mPal = "    "
162     mPAL22 = "     "
164     MOLIS_ALLAJE = 0

        'If IsNull(EPIK) Then
        '   EPIK = Format(Date, "dd/mm/yyyy")
        'End If

166     marxeio2 = "EKT" + xeirisths + ".TXT"
        'On Error Resume Next

168     m_sthl_ektyp(1) = 0  ' int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )

170     For k = 1 To N
172         AA = LTrim(str(k))
174         m_sthl_ektyp(k + 1) = m_sthl_ektyp(k) + Len(macro("p", AA)) + 1
        Next

        Dim R As New ADODB.Recordset

        'R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
        '        k = 1
        '        With myXL
        '            .cells(1, 3) = R("pelono")
        '            .cells(2, 3) = R("pelepa")
        '            .cells(3, 3) = Now
        '            .cells(4, 3) = EPIKEF
        '        End With
        '
        On Error GoTo print_calc_Err

        '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------

176 If F_TYPEHEADER = 1 Then '===============================================
178     AYJ = 5

180     For k = 0 To N - 1
182         Call oSheet.getCellByPosition(k + 1, 5).setformula(F_T(k).Name)
184       '  call oSheet.getCellByPosition(K + 1,5).Interior.Color = RGB(180, 180, 180)
        Next
        
        
        
    oSheet.getCellRangeByPosition(1, 5, k + 1, 5).CharWeight = 150 'Bold
    oSheet.getCellRangeByPosition(1, 5, k + 1, 5).CharColor = 16711680 ' color
    oSheet.getCellRangeByPosition(1, 5, k + 1, 5).CellBackColor = RGB(0, 220, 240) 'cell back color
    oSheet.getCellRangeByPosition(1, 5, k + 1, 5).CharHeight = 15 'Font size
        
        
        
        
        
186  '   excel.Windows(1).SplitRow = 5
188   '  excel.Windows(1).FreezePanes = True

190    ' myXL.ROWS(5).Font.Size = 14
192     'myXL.ROWS(5).Font.fontStyle = 12


    Else
194     AYJ = 0
    End If  '=================================================================




        'excel.Windows(1).SplitColumn = 1

        '  ROWS("6:6").Select
        '    ActiveWindow.FreezePanes = True
        '    Range("B14").Select
        '    ActiveWindow.SmallScroll Down:=0
        '    ROWS("8:8").Select
        '    With selection.Interior
        '        .Pattern = xlSolid
        '        .PatternColorIndex = xlAutomatic
        '        .ThemeColor = xlThemeColorAccent3
        '        .TintAndShade = 0.599993896298105
        '        .PatternTintAndShade = 0
        '    End With

        Dim LAST_TIMH

        Dim synt1

196     synt1 = IIf(IsNull("SYNT1"), "true", synt1)   ' όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
       ' AYJ = 5

198     Typose = 0

        Dim row, Pedio_epan As Integer, timh_epan

200     Pedio_epan = GROUPN - 1
        '   If Pedio_epan < 0 Then Pedio_epan = 0
202     timh_epan = "ωωωωωωω"
        
        '     F_T (k)
        Dim synola_SELIDOS

204     synola_SELIDOS = False

        Dim M_SYN(20) As Single

        Dim n2

206     For n2 = 0 To 20: M_SYN(n2) = 0: Next

        '--------------------------------------------------------
208     g_Stop = 1    'entos loop

210     Do While Not F_T.EOF
212         AYJ = AYJ + 1

214         If FF.CancelButton.Enabled = False Then

                Exit Do

            End If

            'If g_Stop = 2 Then
            '   Exit Do
            'End If

            'MDIForm1.Caption = AYJ
            '----------------- ΤΥΠΩΣΕ ΟΛΑ ΤΑ ΠΕΔΙΑ ---------------------
216         For k = 0 To N - 1

                'If k = Pedio_epan And timh_epan = F_T(k) Then
218             If k = Pedio_epan Then
220                 If timh_epan = F_T(k) Then
                        ' MHN TYPVNEIS TO PEDIO XANA
                    Else ' ΑΛΛΑΓΗ GROUP
222                     AYJ = AYJ + 1
224                     GoSub PRINTGROUP
                        '                          For n2 = 0 To n - 1
                        '
                        '                               If m_sumes(n2) > 0 Then
                        '                                  If mID$(SUgm_str, n2 + 1, 1) = "1" Then    ' SUMA
                        '                                    call oSheet.getCellByPosition(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2), "######,##0.00"), F_T(n2).DefinedSize + 2)
                        '                                     call oSheet.getCellByPosition(AYJ - 1, n2).Interior.Color = RGB(100, 100, 500)
                        '                                  Else
                        '                                    call oSheet.getCellByPosition(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2) / (AYJ - 7), "######,##0.00"), F_T(n2).DefinedSize + 2)
                        '                                  End If
                        '
                        '                               Else
                        '                                  call oSheet.getCellByPosition(AYJ - 1, n2 + 1) = ""  ' Right(Space(50), F_T(K).DefinedSize);
                        '                               End If
                        '
                        '                               M_SYN(n2) = m_sumes(n2)
                        '
                        '                           Next

                    End If
                End If
              
226             If k = Pedio_epan And timh_epan = F_T(k) Then
                Else

228                 If Left(F_T(k), 3) = "@@@" Then    '  A/A
230                     Call oSheet.getCellByPosition(k + 1, AYJ).setformula(Right(Space(30) + Format(AYJ - 5, "######"), F_T(k).DefinedSize))
232                 ElseIf F_T(k).Type = 7 Or F_T(k).Type = 5 Or F_T(k).Type = 3 Or F_T(k).Type = 131 Then   'IsNumeric(f_t(K)) πραγματικο

234                     If IsNull(F_T(k)) Then
236                         Call oSheet.getCellByPosition(k + 1, AYJ).setformula(Right(Space(30) + Format(0, "######,##0.00"), 13))   '  F_T(K).DefinedSize))
                        Else
238                         Call oSheet.getCellByPosition(k + 1, AYJ).setformula(Right(Space(30) + Format(F_T(k), "######,##0.00"), 13))    '  F_T(K).DefinedSize)
                        End If

240                 ElseIf F_T(k).Type = 8 Or F_T(k).Type = 135 Then   'DATE
242                     Call oSheet.getCellByPosition(k + 1, AYJ).setformula(Right(Space(30) + Format(F_T(k), "DD/MM/YYYY"), 10))
                    Else    ' 10 STRING

244                     If IsNull(F_T(k)) Then

                        Else

                            ' Print #1, Tab(f_tab(K)); to928(F_T(K));
246                         If k < N - 1 Then
                                ' για να μην παταει στην επόμενη στήλη
248                             Call oSheet.getCellByPosition(k + 1, AYJ).setformula("'" + Left(F_T(k), f_tab(k + 1) - f_tab(k) - 1))
                            Else
250                             Call oSheet.getCellByPosition(k + 1, AYJ).setformula("'" + F_T(k))     'to928(F_T(K))
                            End If
                        End If

252                     If m_sthl_ektyp(k) > 2 Then
                            ' If K = 1 Then Print #1,
                        End If  'm_sthl_ektyp(K) > 2

                    End If
                End If
              
                ' soymes---------------------
254             If mID$(SUgm_str, k + 1, 1) = "1" Or mID$(SUgm_str, k + 1, 1) = "2" Then
256                 If IsNull(F_T(k)) Then

                    Else
258                     m_sumes(k) = m_sumes(k) + Val(Replace(F_T(k), ",", "."))
                    End If
                End If    ' mid$(sugm_str,k,1)

260             If AYJ Mod 2 = 0 And F_XROMATA = 1 Then
262                 'call oSheet.getCellByPosition( K + 1,ayj).Interior.Color = RGB(200, 200, 200)
                End If

264             If Pedio_epan >= 0 Then
266                 timh_epan = F_T(Pedio_epan)
                End If

            Next

            ' Print #1,
            '   If GROUPN > 0 Then
            '       LAST_TIMH = F_T(GROUPN)
            '   End If
268         F_T.MoveNext
            '   If Not F_T.EOF Then
            '      If GROUPN > 0 Then
            '          If LAST_TIMH <> F_T(GROUPN) Then
            '              AYJ = AYJ + 1
            '          End If
            '      End If
            '  End If
                   
            ' αν τελειωσε να κανει μερικη σουμα στο τελευταιο group
            '                   If F_T.EOF Then
            '                          For n2 = 0 To n - 1
            '                               If m_sumes(n2) > 0 Then
            '                                  If mID$(SUgm_str, n2 + 1, 1) = "1" Then    ' SUMA
            '                                    call oSheet.getCellByPosition(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2), "######,##0.00"), F_T(n2).DefinedSize + 2)
            '                                     call oSheet.getCellByPosition(AYJ - 1, n2).Interior.Color = RGB(100, 100, 500)
            '                                  Else
            '                                    call oSheet.getCellByPosition(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2) / (AYJ - 7), "######,##0.00"), F_T(n2).DefinedSize + 2)
            '                                  End If
            '
            '                               Else
            '                                  call oSheet.getCellByPosition(AYJ - 1, n2 + 1) = ""  ' Right(Space(50), F_T(K).DefinedSize);
            '                               End If
            '
            '                               M_SYN(n2) = m_sumes(n2)
            '
            '                           Next
            '                     End If

270         If AYJ Mod 100 = 0 Then

272             DoEvents
274             FF.Caption = "Εγγραφή " + Format(AYJ, "######")    ' + "/" + Format(F_T.RecordCount, "######")
            End If

        Loop

276     Unload FF
                          
278     If GROUPN > 0 Then
280         AYJ = AYJ + 2
282         GoSub PRINTGROUP
                          
            '                          For n2 = 0 To n - 1
            '                               If m_sumes(n2) > 0 Then
            '                                  If mID$(SUgm_str, n2 + 1, 1) = "1" Then    ' SUMA
            '                                    call oSheet.getCellByPosition(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2), "######,##0.00"), F_T(n2).DefinedSize + 2)
            '                                     call oSheet.getCellByPosition(AYJ - 1, n2).Interior.Color = RGB(100, 100, 500)
            '                                  Else
            '                                    call oSheet.getCellByPosition(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2) / (AYJ - 7), "######,##0.00"), F_T(n2).DefinedSize + 2)
            '                                  End If
            '                               Else
            '                                  call oSheet.getCellByPosition(AYJ - 1, n2 + 1) = ""  ' Right(Space(50), F_T(K).DefinedSize);
            '                               End If
            '                               M_SYN(n2) = m_sumes(n2)
            '                           Next
        End If

284     g_Stop = 0    'adiaforo
        '"
286     AYJ = AYJ + 1

        'call oSheet.getCellByPosition(AYJ, 1) = String$(LSYN - 2, "-")
        '"
        ' Print #1, Chr(13)  ' 6/12/2007
        '"
        '   aa = f_kodik("sum_seltxt")
        Dim PARAM    '

288     PARAM = IIf(AA = " ", "  ", macro(AA, 0))
        'pr_SUMselidas param
        'PRINT2_Xsumes "ΓΕΝΙΚΟ ΣΥΝΟΛΟ"
290     AYJ = AYJ + 1
292     GoSub printSUM

        'AppActivate returnValue

        ' On Error Resume Next

294   '  myXL.Columns("A:K").Select
296   '  myXL.Columns.AutoFit

        'Dim R As New ADODB.Recordset
298    R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
300    k = 1

302 If F_TYPEHEADER = 1 Then

304    'With myXL
306        ' .cells(1, 1) = R("pelono")
           Call oSheet.getCellByPosition(1, 1).setformula(R("pelono"))
           Call oSheet.getCellByPosition(1, 2).setformula(R("pelepa"))
           Call oSheet.getCellByPosition(1, 3).setformula(Now)
           Call oSheet.getCellByPosition(1, 4).setformula(EPIKEF)
308     '   .cells(2, 1) = R("pelepa")
310      '  .cells(3, 1) = Now
312       ' .cells(4, 1) = EPIKEF
        'End With

    End If


  Dim range As Object
    Set range = oSheet.getCellRangeByPosition(0, 0, 50, 50)
        Dim columns As Object
        Set columns = range.getColumns()
        For i = 0 To 50 - 1
            columns.getByIndex(i).optimalWidth = True
        Next







314    MDIForm1.MousePointer = vbNormal
        'excel.Visible = True

316   ' excel.Visible = True

        On Error Resume Next

318   ' Kill "C:\EKTYP.XLS"

320    DoEvents

322    'myXL.SaveAs "C:\EKTYP.XLS"

        Dim ANS3 As Long

324   ' ANS3 = MsgBox("Κλείνω το EXCEL", vbYesNo)

326    If ANS3 = vbYes Then
328       ' Call workbook.Close(False)
330       ' excel.Quit
332       ' Set excel = Nothing
        End If
        
        
        
        
   ' Call oDoc.storeToURL(savepath, arg())
  '  oDoc.Close (True)
  '  Set oDoc = Nothing


        
        
        
        
        
        
        
        
        
        
        
        
        

        Exit Function

PRINTGROUP:

334     For n2 = 0 To N - 1

336         If m_sumes(n2) > 0 Then
338             If mID$(SUgm_str, n2 + 1, 1) = "1" Then    ' SUMA
340                 Call oSheet.getCellByPosition(n2 + 1, AYJ - 1).setformula(Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2), "######,##0.00"), F_T(n2).DefinedSize + 2))
342                 'call oSheet.getCellByPosition(n2 + 1,ayj-1).Interior.Color = RGB(100, 100, 500)
                Else
344                 Call oSheet.getCellByPosition(n2 + 1, AYJ - 1).setformula(Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2) / (AYJ - 7), "######,##0.00"), F_T(n2).DefinedSize + 2))
                End If

            Else
346             Call oSheet.getCellByPosition(n2 + 1, AYJ - 1).setformula("")   ' Right(Space(50), F_T(K).DefinedSize);
            End If

348         M_SYN(n2) = m_sumes(n2)
        Next

350     Return

printSUM:

352    For k = 0 To N - 1

354        If m_sumes(k) > 0 Then
356            If mID$(SUgm_str, k + 1, 1) = "1" Then    ' SUMA
358                Call oSheet.getCellByPosition(k + 1, AYJ).setformula(Right(Space(30) + Format(m_sumes(k), "######,##0.00"), F_T(k).DefinedSize + 2))
                Else
360                Call oSheet.getCellByPosition(k + 1, AYJ).setformula(Right(Space(30) + Format(m_sumes(k) / (AYJ - 7), "######,##0.00"), F_T(k).DefinedSize + 2))
                End If

            Else
362             Call oSheet.getCellByPosition(k + 1, AYJ).setformula("")    ' Right(Space(50), F_T(K).DefinedSize);
            End If

        Next

364    Return

printEpik:

366    For k = 0 To N - 1

368        If m_sumes(k + 1) > 0 Then
370            Call oSheet.getCellByPosition(k + 1, AYJ).setformula(Right(Space(30) + Format(m_sumes(k + 1), "######,##0.00"), F_T(k).DefinedSize + 2))
            Else
372            Call oSheet.getCellByPosition(k + 1, AYJ).setformula(Right(Space(50), F_T(k).DefinedSize))
            End If

        Next

374    Return


        '<EhFooter>
        Exit Function

print_calc_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.print_calc " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function

'openoffice ---------------------calc
'Dim oSM
'Dim oDesk, oDoc As Object
'Dim arg()
'Dim ocell As Object
'Dim oSheet As Object
'Dim importDesc(0 To 2) As Object
'Dim oImportrange As Object
'Dim oSQL As Object
'Dim oContext As Object
'
'
'Set oSM = CreateObject("com.sun.star.ServiceManager")
'Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
'
'Set oDoc = oDesk.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg())
'Set oSheet = oDoc.Sheets.getByName("Φύλλο1")
'
'oSM.createInstance ("com.sun.star.sheet.DatabaseImportDescriptor ")
'oSM.createInstance ("com.sun.star.util.Ximportable")
'
''importDesc(0).Name = "DatabaseName"
''importDesc(0).Value = "provider=SQLOLEDB;Server={ServerName};Uid={UserName};Pwd={Password};Database={DatabaseName}"
''
''importDesc(1).Name = "SourceType"
''importDesc(1).Value = "SQL"
''
''importDesc(2).Name = "SourceObject"
''importDesc(2).Value = "Select * from {tablename}"
'For K = 1 To 100
'   Call oSheet.getCellByPosition(1, K).setformula(str(K))
'Next



'------------------------------------------------------------------------------------------------------
Public Sub PRINT_CSV(sql, SUgm_str, EPIKEF, GROUPN, TITLOS_PEDION, mfile)
        '<EhHeader>
        On Error GoTo PRINT_CSV_Err
        '</EhHeader>
        Dim MFIELDS(120), mSYN

        Dim synt_eject
Dim F_FILESAVE As String

100     F_TYPEHEADER = Val(FINDPARAMETROI(1, "EKTYPOTES", "F_TYPEHEADER", "1", "1=ΤΥΠΩΝΟΝΤΑΙ ΕΠΙΚΕΦΑΛΙΔΕΣ ΑΡΧΕΙΟ "))

F_TYPEHEADER = 0
' F_TYPEHEADER = Val(FINDPARAMETROI(1, "EKTYPOTES", "F_TYPEHEADER", "1", "1=ΤΥΠΩΝΟΝΤΑΙ ΕΠΙΚΕΦΑΛΙΔΕΣ ΑΡΧΕΙΟ "))
       ' F_FILESAVE = FINDPARAMETROI(1, "EKTYPOTES", "F_FILESAVE", "C:\MERCVB\PRINT.CSV", "ΑΡΧΕΙΟ ΕΚΤΥΠΩΣΗΣ CSV ΑΡΧΕΙΟΥ ")

        Dim sumes, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), F(120)

        Dim mBSEIRA

        Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE

        Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4

        Dim aaP, aaP2, ektypoths

        Dim EPIK, xeirisths, PPF, epik3

        Dim Sthles, kw, PPD, SELIS

        Dim LSYN, AA, epik2, i As Integer

        Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject

        Dim m_sumes(120), SUMES0(120)

        Dim SS

        Dim m_ekt As Integer

        Dim DUM


102     For k = 1 To 120: m_sumes(k) = 0: SUMES0(k) = 0: Next

        Dim N

        Dim F_T As New ADODB.Recordset

        Dim TT  As Long

104     TT = GetCurrentTime()

         On Error GoTo LATOSSQL
106     F_T.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly


On Error GoTo PRINT_CSV_Err

108     N = F_T.FIELDS.Count

110     For k = 1 To N

112         If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then    'DATE
114             f_tab(k) = 1 + f_tab(k - 1) + 8
            Else

116             If F_T.FIELDS(k - 1).DefinedSize > 200 Then
118                 f_tab(k) = 1 + f_tab(k - 1) + 30
                Else
120                 f_tab(k) = 2 + f_tab(k - 1) + F_T.FIELDS(k - 1).DefinedSize + 3
                End If
            End If

        Next

   

122     LSYN = f_tab(k - 1)

124     mPal = "    "
126     mPAL22 = "     "
128     MOLIS_ALLAJE = 0








        'If IsNull(EPIK) Then
        '   EPIK = Format(Date, "dd/mm/yyyy")
        'End If

130    ' marxeio2 = F_FILESAVE '  "EKT" + xeirisths + ".TXT"
     ' mfile = F_FILESAVE
        On Error Resume Next

132     Close #1
134     Open mfile For Output As #1

136     If LSYN > 128 Then
138         Sthles = LSYN + 10
140     ElseIf LSYN > 80 Then
142         Sthles = LSYN + 10    ' 136
        End If

144     m_sthl_ektyp(1) = 0  ' int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )

146     For k = 1 To N
148         AA = LTrim(str(k))

150         m_sthl_ektyp(k + 1) = m_sthl_ektyp(k) + 6    ' Len(macro("p", aa)) + 1
        Next

        ' PRINT2_Xepik

        Dim R As New ADODB.Recordset

152     R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
        'R("


154 If F_TYPEHEADER = 1 Then '  = Val(FindParametroi(1,"EKTYPOTES", "F_TYPEHEADER", "1", "1=ΤΥΠΩΝΟΝΤΑΙ ΕΠΙΚΕΦΑΛΙΔΕΣ ΑΡΧΕΙΟ "))

156    ' Print #1, R("pelono")
158    ' Print #1, R("pelepa")

160    ' Print #1, Now
162    ' Print #1, Tab(LSYN / 2 - Len(EPIKEF) / 2); EPIKEF

        '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------
      If TITLOS_PEDION = 1 Then
164      For k = 0 To N - 1

166         If f_tab(k) + F_T(k).DefinedSize > f_tab(k + 1) Then
168             Print #1, Tab(f_tab(k)); Left(F_T(k).Name, f_tab(k + 1) - f_tab(k));
            Else

170             If F_T(k).Type = 7 Or F_T(k).Type = 5 Or F_T(k).Type = 131 Then   'IsNumeric(f_t(K)) πραγματικο
172                 Print #1, Tab(f_tab(k)); RightAlign(F_T(k).Name, F_T(k).DefinedSize);    ' Left(F_T(k).Name, F_T(k).DefinedSize);
                Else
174                 Print #1, Tab(f_tab(k)); Left(F_T(k).Name, F_T(k).DefinedSize);
                End If
            End If
176         Print #1, ";";
        Next
      End If

        'δευτερη σειρα επικεφαλίδας
178     For k = 0 To N - 1
            '        If f_tab(K) + F_T(K).DefinedSize > f_tab(K + 1) Then
            '             Print #1, Tab(f_tab(K)); Left(F_T(K).Name, f_tab(K + 1) - f_tab(K));
            '        Else
            '             Print #1, Tab(f_tab(K)); Left(F_T(K).Name, F_T(K).DefinedSize);
            '        End If
            '
            '    Print #1, Tab(f_tab(K)); Mid$(F_T(K).Name, F_T(K).DefinedSize + 1, Len(F_T(K).Name) - F_T(K).DefinedSize + 1);
        Next

180   ' Print #1, Tab(m_sthl_ektyp(1)); String$(LSYN - 2, "-")


    End If  ' F_TYPEHEADER = Val(FindParametroi(1,"EKTYPOTES", "F_TYPEHEADER", "1", "1=ΤΥΠΩΝΟΝΤΑΙ ΕΠΙΚΕΦΑΛΙΔΕΣ ΑΡΧΕΙΟ "))




182     For i = 1 To N
184         m_sumes(i) = 0
186         SUMES0(i) = 0
        Next

        Dim LAST_TIMH

        Dim synt1

188     synt1 = IIf(IsNull("SYNT1"), "true", synt1)   ' όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
190     AYJ = 0
192     Typose = 0

        Dim row

        Dim synola_SELIDOS

194     synola_SELIDOS = False

        '--------------------------------------------------------
196     Do While Not F_T.EOF
198         AYJ = AYJ + 1

200         If end_page = 2 Then    ' ΤΥΠΩΝΕΙ ΓΡΑΜΜΟΥΛΕΣ --------------------------------

202          If F_TYPEHEADER = 1 Then   '  Val(FindParametroi(1,"EKTYPOTES", "F_TYPEHEADER", "1", "1=ΤΥΠΩΝΟΝΤΑΙ ΕΠΙΚΕΦΑΛΙΔΕΣ ΑΡΧΕΙΟ "))
204            ' Print #1, Tab(m_sthl_ektyp(1) + 1); String$(LSYN, "-")
             End If
206             end_page = 0

208             MOLIS_ALLAJE = 1
            End If    ' end_page=2

'210         If synola_SELIDOS Then
'212             If AYJ Mod 80 = 0 And F_TYPEHEADER = 1 Then    '
'                    'syn_epik
'214                ' GoSub printSUM
'216                ' Print #1,
'218                ' Print #1,
'
'220                ' Print #1,
'
'222                 end_page = 0
'
'                    '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------
'
'224                ' Print #1,
'226                ' Print #1,
'228                ' Print #1,
'
'230                 For k = 0 To N - 1
'
'232                     If F_T(k).Type = 7 Or F_T(k).Type = 5 Or F_T(k).Type = 131 Then   'IsNumeric(f_t(K)) πραγματικο
'234                         Print #1, Tab(f_tab(k)); RightAlign(F_T(k).Name, F_T(k).DefinedSize)     ' Left(F_T(k).Name, F_T(k).DefinedSize);
'                        Else
'236                         Print #1, Tab(f_tab(k)); Left(F_T(k).Name, F_T(k).DefinedSize);
'                        End If
'
'                    Next
'
'                    'δευτερη σειρα επικεφαλίδας
'238                 For k = 0 To N - 1
'240                     Print #1, Tab(f_tab(k)); mID$(F_T(k).Name, F_T(k).DefinedSize + 1, Len(F_T(k).Name) - F_T(k).DefinedSize + 1); ";";
'                    Next
'
'242                 'Print #1, Tab(m_sthl_ektyp(1)); String$(LSYN - 2, "-")
'                End If  'end_page = 1 Then
'            End If

            '----------------- ΤΥΠΩΣΕ ΟΛΑ ΤΑ ΠΕΔΙΑ ---------------------
244         For k = 0 To N - 1

            '  Print #1, F_T(k);


246             If Left(F_T(k), 3) = "@@@" Then    '  A/A
248                ' Print #1, Tab(f_tab(k)); Right(Space(30) + Format(AYJ, "######"), F_T(k).DefinedSize);
250             ElseIf F_T(k).Type = 7 Or F_T(k).Type = 5 Or F_T(k).Type = 131 Then   'IsNumeric(f_t(K)) πραγματικο
252                 Print #1, Replace(Right(Space(30) + Format(F_T(k), "#########0.00"), F_T(k).DefinedSize), ",", ".");   ' F_T(k).DefinedSize);
254             ElseIf F_T(k).Type = 8 Then   'DATE
256                Print #1, Right(Space(30) + Format(F_T(k), "DD/MM/YYYY"), F_T(k).DefinedSize);
                Else    ' 10 STRING

258                If IsNull(F_T(k)) Then
260                    Print #1, Tab(f_tab(k));
                    Else

                        ' Print #1, Tab(f_tab(K)); to928(F_T(K));
262                    If k < N - 1 Then
                            ' για να μην παταει στην επόμενη στήλη
264                        Print #1, to928(F_T(k));
                        Else
266                        Print #1, to928(F_T(k));
                        End If
                    End If

268                If m_sthl_ektyp(k) > 2 Then
270                    If k = 1 Then Print #1,
                    End If  'm_sthl_ektyp(K) > 2

                End If

                ' soymes---------------------
272            If mID$(SUgm_str, k + 1, 1) = "1" Or mID$(SUgm_str, k + 1, 1) = "2" Then
274                If IsNull(F_T(k)) Then

                    Else
276                    m_sumes(k) = m_sumes(k) + F_T(k)
                    End If
                End If    ' mid$(sugm_str,k,1)
278            Print #1, ";";
            Next

280        Print #1,

282        If GROUPN > 0 Then
284            LAST_TIMH = F_T(GROUPN - 1)
            End If

286        F_T.MoveNext

288        If Not F_T.EOF Then
290            If GROUPN > 0 Then
292                If LAST_TIMH <> F_T(GROUPN - 1) Then
294                    Print #1,
                    End If
                End If
            End If

        Loop

        Dim YPOSEL(10)
        Dim PARAM


        '"



296 If F_TYPEHEADER = 1 Then


298  '  Print #1, Tab(m_sthl_ektyp(1) + 1); String$(LSYN - 2, "-")
300    PARAM = IIf(AA = " ", "  ", AA)  ' macro(aa, 0))
302  '  GoSub printSUM
304   ' Print #1,
306   ' Print #1, Tab(m_sthl_ektyp(1) + 1); String$(LSYN - 2, "-")
308   ' For k = 1 To 4
310    '    Print #1,    '  1, 0; YPOSEL(k)
       ' Next

    End If


312    If eject = "y" Then
            'Print #1, Chr(12)
        End If

        ' αποθηκεύω τις σούμες στα e1,e2,.. για να τις χρησιμοποιώ
        '  For K = 1 To N
        '   aa = "e" + LTrim(Str(K))
        '    macro(aa, 0) = m_sumes(K)
        ' Next

314    Close #1

       

        On Error Resume Next

316  ' DUM = Shell("notepad.exe " + mfile, vbMaximizedFocus)
          
         ' MsgBox "Αποθηκεύθηκε στο αρχείο " + mfile



         MHNYMA2.Timer1.Interval = 1000
         MHNYMA2.Label2.Caption = "Αποθηκεύθηκε στο αρχείο " + mfile
         MHNYMA2.SHOW 1



        
318    SELIS = 1

        Exit Sub

printSUM:

320    For k = 0 To N - 1

322        If k = 0 Then
324            If Not IsNumeric(m_sumes(0)) Then
326                m_sumes(0) = "ΣΎΝΟΛΑ"
                End If
            End If

328        If m_sumes(k) > 0 Then
330            If mID$(SUgm_str, k + 1, 1) = "1" Then    ' αθροισμα
332                Print #1, Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k), "#######,##0.00"), F_T(k).DefinedSize);
                    'Print #1, Tab(f_tab(k)-2); Right(Space(30) + Format(m_sumes(k), "#######,##0.00"), 2 + F_T(k).DefinedSize);
                Else  'mesos oros
334                Print #1, Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k) / AYJ, "######,##0.00"), F_T(k).DefinedSize);
                End If

            Else

336            If k = 0 Then
338                Print #1, Tab(f_tab(k) - 2); "Σύνολα";    ' Right(Space(50), F_T(K).DefinedSize);
                Else
                    '  Print #1, Tab(f_tab(k) - 2); ""; ' Right(Space(50), F_T(K).DefinedSize);
                End If
            End If
340        Print #1, ";";
        Next

342    Print #1,
344    Return

printEpik:

346    For k = 0 To N - 1

348        If m_sumes(k + 1) > 0 Then
350            Print #1, Tab(f_tab(k) - 2); Right(Space(30) + Format(m_sumes(k + 1), "######,##0.00"), F_T(k).DefinedSize + 2);
            Else
352            Print #1, Tab(f_tab(k)); Right(Space(50), F_T(k).DefinedSize);
            End If
354          Print #1, ";";
        Next

356    Print #1,
358    Return


        '<EhFooter>
        Exit Sub

PRINT_CSV_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.PRINT_CSV " & _
               "at line " & Erl


        Resume Next
        
        
LATOSSQL:
         MHNYMA2.Timer1.Interval = 2000
         MHNYMA2.Label2.Caption = "ΛΑΘΟΣ SQL SCRIPT " + sql
         MHNYMA2.SHOW 1
        
        '</EhFooter>
End Sub



'------------------------------------------------------------------------------------------------------
Sub PRINT_TEXT(sql, SUgm_str, EPIKEF, GROUPN)
        '<EhHeader>
        On Error GoTo PRINT_TEXT_Err
        '</EhHeader>
        Dim MFIELDS(120), mSYN

        Dim synt_eject

100     F_TYPEHEADER = Val(FINDPARAMETROI(1, "EKTYPOTES", "F_TYPEHEADER", "1", "1=ΤΥΠΩΝΟΝΤΑΙ ΕΠΙΚΕΦΑΛΙΔΕΣ ΑΡΧΕΙΟ "))

        Dim sumes, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), F(120)

        Dim mBSEIRA

        Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE

        Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4

        Dim aaP, aaP2, ektypoths

        Dim EPIK, xeirisths, PPF, epik3

        Dim Sthles, kw, PPD, SELIS

        Dim LSYN, AA, epik2, i As Integer

        Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject

        Dim m_sumes(120), SUMES0(120)

        Dim SS

        Dim m_ekt As Integer

        Dim DUM


102     For k = 1 To 120: m_sumes(k) = 0: SUMES0(k) = 0: Next

        Dim N

        Dim F_T As New ADODB.Recordset

        Dim TT  As Long

104     TT = GetCurrentTime()
    On Error GoTo PRINT_TEXT_Err
        ' On Error GoTo print3_xar_Err
106     F_T.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

108     N = F_T.FIELDS.Count

110     For k = 1 To N

112         If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then    'DATE
114             f_tab(k) = 1 + f_tab(k - 1) + 8
            Else

116             If F_T.FIELDS(k - 1).DefinedSize > 200 Then
118                 f_tab(k) = 1 + f_tab(k - 1) + 30
                Else
120                 f_tab(k) = 2 + f_tab(k - 1) + F_T.FIELDS(k - 1).DefinedSize + 3
                End If
            End If

        Next

        '    For k = 1 To N
        '        If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then    'DATE
        '            f_tab(k) = 2 + f_tab(k - 1) + 8
        '        Else
        '            If F_T.FIELDS(k - 1).DefinedSize > 200 Then
        '                f_tab(k) = 2 + f_tab(k - 1) + 30
        '            Else
        '                f_tab(k) = 3 + f_tab(k - 1) + F_T.FIELDS(k - 1).DefinedSize + 3
        '            End If
        '        End If
        '    Next

122     LSYN = f_tab(k - 1)

124     mPal = "    "
126     mPAL22 = "     "
128     MOLIS_ALLAJE = 0








        'If IsNull(EPIK) Then
        '   EPIK = Format(Date, "dd/mm/yyyy")
        'End If

130     marxeio2 = "EKT" + xeirisths + ".TXT"

        On Error Resume Next

132     Close #1
134     Open "c:\print" For Output As #1

136     If LSYN > 128 Then
138         Sthles = LSYN + 10
140     ElseIf LSYN > 80 Then
142         Sthles = LSYN + 10    ' 136
        End If

144     m_sthl_ektyp(1) = 0  ' int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )

146     For k = 1 To N
148         AA = LTrim(str(k))

150         m_sthl_ektyp(k + 1) = m_sthl_ektyp(k) + 6    ' Len(macro("p", aa)) + 1
        Next

        ' PRINT2_Xepik

        Dim R As New ADODB.Recordset

152     R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
        'R("


154 If F_TYPEHEADER = 1 Then '  = Val(FindParametroi(1,"EKTYPOTES", "F_TYPEHEADER", "1", "1=ΤΥΠΩΝΟΝΤΑΙ ΕΠΙΚΕΦΑΛΙΔΕΣ ΑΡΧΕΙΟ "))

156     Print #1, R("pelono")
158     Print #1, R("pelepa")

160     Print #1, Now
162     Print #1, Tab(LSYN / 2 - Len(EPIKEF) / 2); EPIKEF

        '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------

164     For k = 0 To N - 1

166         If f_tab(k) + F_T(k).DefinedSize > f_tab(k + 1) Then
168             Print #1, Tab(f_tab(k)); Left(F_T(k).Name, f_tab(k + 1) - f_tab(k));
            Else

170             If F_T(k).Type = 7 Or F_T(k).Type = 5 Or F_T(k).Type = 131 Then   'IsNumeric(f_t(K)) πραγματικο
172                 Print #1, Tab(f_tab(k)); RightAlign(F_T(k).Name, F_T(k).DefinedSize);    ' Left(F_T(k).Name, F_T(k).DefinedSize);
                Else
174                 Print #1, Tab(f_tab(k)); Left(F_T(k).Name, F_T(k).DefinedSize);
                End If

                '         Print #1, Tab(f_tab(k)); Left(F_T(k).Name, F_T(k).DefinedSize);
            End If

        Next

        'δευτερη σειρα επικεφαλίδας
176     For k = 0 To N - 1
            '        If f_tab(K) + F_T(K).DefinedSize > f_tab(K + 1) Then
            '             Print #1, Tab(f_tab(K)); Left(F_T(K).Name, f_tab(K + 1) - f_tab(K));
            '        Else
            '             Print #1, Tab(f_tab(K)); Left(F_T(K).Name, F_T(K).DefinedSize);
            '        End If
            '
            '    Print #1, Tab(f_tab(K)); Mid$(F_T(K).Name, F_T(K).DefinedSize + 1, Len(F_T(K).Name) - F_T(K).DefinedSize + 1);
        Next

178     Print #1, Tab(m_sthl_ektyp(1)); String$(LSYN - 2, "-")


    End If  ' F_TYPEHEADER = Val(FindParametroi(1,"EKTYPOTES", "F_TYPEHEADER", "1", "1=ΤΥΠΩΝΟΝΤΑΙ ΕΠΙΚΕΦΑΛΙΔΕΣ ΑΡΧΕΙΟ "))




180     For i = 1 To N
182         m_sumes(i) = 0
184         SUMES0(i) = 0
        Next

        Dim LAST_TIMH

        Dim synt1

186     synt1 = IIf(IsNull("SYNT1"), "true", synt1)   ' όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
188     AYJ = 0
190     Typose = 0

        Dim row

        Dim synola_SELIDOS

192     synola_SELIDOS = False

        '--------------------------------------------------------
194     Do While Not F_T.EOF
196         AYJ = AYJ + 1

198         If end_page = 2 Then    ' ΤΥΠΩΝΕΙ ΓΡΑΜΜΟΥΛΕΣ --------------------------------

200          If F_TYPEHEADER = 1 Then   '  Val(FindParametroi(1,"EKTYPOTES", "F_TYPEHEADER", "1", "1=ΤΥΠΩΝΟΝΤΑΙ ΕΠΙΚΕΦΑΛΙΔΕΣ ΑΡΧΕΙΟ "))
202             Print #1, Tab(m_sthl_ektyp(1) + 1); String$(LSYN, "-")
             End If
204             end_page = 0

206             MOLIS_ALLAJE = 1
            End If    ' end_page=2

208         If synola_SELIDOS Then
210             If AYJ Mod 80 = 0 And F_TYPEHEADER = 1 Then    '
                    'syn_epik
212                 GoSub printSUM
214                 Print #1,
216                 Print #1,

218                 Print #1,

220                 end_page = 0

                    '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------

222                 Print #1,
224                 Print #1,
226                 Print #1,

228                 For k = 0 To N - 1

230                     If F_T(k).Type = 7 Or F_T(k).Type = 5 Or F_T(k).Type = 131 Then   'IsNumeric(f_t(K)) πραγματικο
232                         Print #1, Tab(f_tab(k)); RightAlign(F_T(k).Name, F_T(k).DefinedSize + 3)    ' Left(F_T(k).Name, F_T(k).DefinedSize);
                        Else
234                         Print #1, Tab(f_tab(k)); Left(F_T(k).Name, F_T(k).DefinedSize);
                        End If

                    Next

                    'δευτερη σειρα επικεφαλίδας
236                 For k = 0 To N - 1
238                     Print #1, Tab(f_tab(k)); mID$(F_T(k).Name, F_T(k).DefinedSize + 1, Len(F_T(k).Name) - F_T(k).DefinedSize + 1);
                    Next

240                 Print #1, Tab(m_sthl_ektyp(1)); String$(LSYN - 2, "-")
                End If  'end_page = 1 Then
            End If

            '----------------- ΤΥΠΩΣΕ ΟΛΑ ΤΑ ΠΕΔΙΑ ---------------------
242         For k = 0 To N - 1

244             If Left(F_T(k), 3) = "@@@" Then    '  A/A
246                 Print #1, Tab(f_tab(k)); Right(Space(30) + Format(AYJ, "######"), F_T(k).DefinedSize);
248             ElseIf F_T(k).Type = 7 Or F_T(k).Type = 5 Or F_T(k).Type = 131 Then   'IsNumeric(f_t(K)) πραγματικο
250                 Print #1, Tab(f_tab(k)); Right(Space(30) + Format(F_T(k), "#######,##0.00"), 14);    ' F_T(k).DefinedSize);
252            ElseIf F_T(k).Type = 8 Then   'DATE
254                Print #1, Tab(f_tab(k)); Right(Space(30) + Format(F_T(k), "DD/MM/YYYY"), 10);
                Else    ' 10 STRING

256                If IsNull(F_T(k)) Then
258                    Print #1, Tab(f_tab(k));
                    Else

                        ' Print #1, Tab(f_tab(K)); to928(F_T(K));
260                    If k < N - 1 Then
                            ' για να μην παταει στην επόμενη στήλη
262                        Print #1, Tab(f_tab(k)); Left(to928(F_T(k)), f_tab(k + 1) - f_tab(k) - 1);
                        Else
264                        Print #1, Tab(f_tab(k)); to928(F_T(k));
                        End If
                    End If

266                If m_sthl_ektyp(k) > 2 Then
268                    If k = 1 Then Print #1,
                    End If  'm_sthl_ektyp(K) > 2

                End If

                ' soymes---------------------
270            If mID$(SUgm_str, k + 1, 1) = "1" Or mID$(SUgm_str, k + 1, 1) = "2" Then
272                If IsNull(F_T(k)) Then

                    Else
274                    m_sumes(k) = m_sumes(k) + F_T(k)
                    End If
                End If    ' mid$(sugm_str,k,1)

            Next

276        Print #1,

278        If GROUPN > 0 Then
280            LAST_TIMH = F_T(GROUPN - 1)
            End If

282        F_T.MoveNext

284        If Not F_T.EOF Then
286            If GROUPN > 0 Then
288                If LAST_TIMH <> F_T(GROUPN - 1) Then
290                    Print #1,
                    End If
                End If
            End If

        Loop

        Dim YPOSEL(10)
        Dim PARAM


        '"



292 If F_TYPEHEADER = 1 Then


294    Print #1, Tab(m_sthl_ektyp(1) + 1); String$(LSYN - 2, "-")
296    PARAM = IIf(AA = " ", "  ", AA)  ' macro(aa, 0))
298    GoSub printSUM
300    Print #1,
302    Print #1, Tab(m_sthl_ektyp(1) + 1); String$(LSYN - 2, "-")
304    For k = 1 To 4
306        Print #1,    '  1, 0; YPOSEL(k)
        Next

    End If


308    If eject = "y" Then
            'Print #1, Chr(12)
        End If

        ' αποθηκεύω τις σούμες στα e1,e2,.. για να τις χρησιμοποιώ
        '  For K = 1 To N
        '   aa = "e" + LTrim(Str(K))
        '    macro(aa, 0) = m_sumes(K)
        ' Next

310    Close #1

        ' pp = 1
        'If pp = 1 Then
        ' If ektypoths = -1 Then
        ' Else

        '  PPD = MsgBox("ΤΟ ΒΛΕΠΩ ΠΡΙΝ ΤΥΠΩΘΕΙ", vbYesNo)
        '  If PPD = vbYes Then

        On Error Resume Next

312    DUM = Shell("c:\mercvb\notepad.exe c:\print", vbMaximizedFocus)

        '     PPF = MsgBox("Προχωρώ στην εκτύπωση", vbYesNo)
        '    If PPF = vbYes Then
        '       DUM = Shell("c:\mercvb\notepad.exe /p c:\print", vbMaximizedFocus)
        '        End If
        '    Else
        '    End If
        '  End If
        ' End If
314    SELIS = 1

        Exit Sub

printSUM:

316    For k = 0 To N - 1

318        If k = 0 Then
320            If Not IsNumeric(m_sumes(0)) Then
322                m_sumes(0) = "ΣΎΝΟΛΑ"
                End If
            End If

324        If m_sumes(k) > 0 Then
326            If mID$(SUgm_str, k + 1, 1) = "1" Then    ' αθροισμα
328                Print #1, Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k), "#######,##0.00"), 14);
                    'Print #1, Tab(f_tab(k)-2); Right(Space(30) + Format(m_sumes(k), "#######,##0.00"), 2 + F_T(k).DefinedSize);
                Else  'mesos oros
330                Print #1, Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k) / AYJ, "######,##0.00"), F_T(k).DefinedSize + 2);
                End If

            Else

332            If k = 0 Then
334                Print #1, Tab(f_tab(k) - 2); "Σύνολα";    ' Right(Space(50), F_T(K).DefinedSize);
                Else
                    '  Print #1, Tab(f_tab(k) - 2); ""; ' Right(Space(50), F_T(K).DefinedSize);
                End If
            End If

        Next

336    Print #1,
338    Return

printEpik:

340    For k = 0 To N - 1

342        If m_sumes(k + 1) > 0 Then
344            Print #1, Tab(f_tab(k) - 2); Right(Space(30) + Format(m_sumes(k + 1), "######,##0.00"), F_T(k).DefinedSize + 2);
            Else
346            Print #1, Tab(f_tab(k)); Right(Space(50), F_T(k).DefinedSize);
            End If

        Next

348    Print #1,
350    Return


        '<EhFooter>
        Exit Sub

PRINT_TEXT_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.PRINT_TEXT " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Sub

Function print7_excel(sql As String, _
                      SUgm_str, _
                      ByVal EPIKEF As String, _
                      ByVal GROUPN As Integer)
'10/12/2014
'οι παρακατω σειρές εγιναν γιατι αλλαξε ο τυπος
'apo varchar(30) se nvarchar(max) για να δουλευει και κυριλλικά
'το προβλημα ηταν ότι στα κυριλλικα εβγαζε κενα
'==================================
'F_T.Close
'FF.Adodc1.ConnectionString = gConnect
'FF.Adodc1.RecordSource = sql
'FF.Adodc1.Refresh
'Dim RR As Object
'Set RR = FF.Adodc1.Recordset
'==================================
        
        
        
        
        
        
        
        '<EhHeader>
        On Error GoTo print7_excel_Err
        '</EhHeader>

        '"*********************** P R I N T _ X A R ************************
        '"**************      κάνει παρουσίαση αρχείου στο χαρτί
        '"*** n=αριθμός fields που παρουσιάζονται
        '"*** synuhkh2 η συνθήκη για το  IF, sugm_str βλέπει που θα κάνει σούμες ,π.χ. "00100" κάνει σούμες στο 3ο field
        '"*** sum_pic  το picture γιά τις σούμες
        '"*** Ei,Fi,Pi  :Επικεφαλίδα παρουσίασης,Fields που παρουσιάζονται,Picture παρουσίασης
        '"** synt_eject:=0  αλλάζει σελίδα όταν μεταβάλλεται η στήλη
        Dim MFIELDS(120), mSYN

        Dim synt_eject

        Dim sumes, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), F(120)

        Dim mBSEIRA

        Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE

        Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4

        Dim aaP, aaP2, ektypoths

        Dim EPIK, xeirisths, PPF, epik3

        Dim Sthles, kw, PPD, SELIS

        Dim LSYN, AA, epik2, i As Integer

        Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject

        Dim m_sumes(120), SUMES0(120)

        Dim SS

        Dim m_ekt As Integer

        Dim DUM

        Dim returnValue

        Dim f_excelPath

        'GROUPN = GROUPN - 1

        'ETSI DEN DOYLEYEI ME OLA TA EXCEL ALLA SE BOHTHAEI STON KODIKA

        '    Dim excel As excel.Application
        '   Dim workbook As excel.workbook
        '  Dim myXL As excel.Worksheet

        'ETSI DOYLEYEI ME OLA TA EXCEL
        Dim excel    As Object ' excel.Application

        Dim workbook As Object ' excel.workbook

        Dim myXL     As Object ' excel.Worksheet

100     Set excel = CreateObject("excel.Application")
    
        ' Excel.Visible = True
102     Set workbook = excel.Workbooks.Add

        On Error Resume Next

        Dim F_XROMATA As Integer

104     F_XROMATA = Val(FINDPARAMETROI(1, "MDIFORM1", "F_XROMATA", "0", "ΧΡΩΜΑΤΑ ΣΤΟ EXCEL=1 ΟΧΙ=0"))

        '"MDIFORM1"

        '  If (MenuShow.Caption = "&Show") Then
        '    MenuShow.Caption = "&Hide"
106     workbook.Activate

108     Set myXL = workbook.ActiveSheet

        Dim FF As New UDialog

110     FF.SHOW
112     FF.OKButton.Visible = False
114     FF.List1.Visible = False

116     FF.CancelButton.Caption = "ΔΙΑΚΟΠΗ"
118     FF.CancelButton.Top = 120
120     FF.CancelButton.Left = 120
122     FF.CancelButton.Width = 2895
124     FF.CancelButton.Height = 495

        'FF.Top = 3000
        'FF.Left = 3000

126     FF.Width = 3210
128     FF.Height = 810

130     FF.Left = Screen.Width / 2 - FF.Width / 2
132     FF.Top = Screen.Height / 2 - FF.Height / 2

        'FF.Top = MDIForm1.Top + (MDIForm1.height) / 2 ' Command2.Top
        'FF.Left = MDIForm1.Left + (MDIForm1.width) / 2 ' Command2.Left

134     FF.Caption = "ΔΙΑΔΙΚΑΣΙΑ ΥΠΟΛΟΓΙΣΜΟΥ"

        ''------------------------ΠΕΤΑΕΙ ΟΛΟΚΛΗΡΟ ΤΟ ΑΡΧΕΙΟ ΑΛΛΑ ΚΟΛΑΕΙ ΣΤΗΝ [ΧΟΝΔ.ΤΙΜΗ]------------------------------
        'Dim db1 As Database
        'Dim FROM As Long
        'On Error Resume Next
        '  Kill "C:\EKTYP.XLS"
        '  DoEvents
        '  myXL.SaveAs "C:\EKTYP.XLS"
        '
        '  Call workbook.Close(False)
        '  excel.Quit
        '  Set excel = Nothing
        '
        'FROM = InStr(1, UCase(sql), "FROM")
        'Set db1 = OpenDatabase("", False, False, gConnect)
        'db1.Execute Left(sql, FROM - 1) + " " + " into 32 in 'c:\EKTYP.xls' 'Excel 8.0;' " + Trim(Mid(sql, FROM, 500))
        ''------------------------------------------------------

136     MDIForm1.MousePointer = vbHourglass
        'excel.Visible = True

        'f_excelPath = FindParametroi(1,"MDIFORM1", "f_excelPath", "C:\Program Files\Microsoft Office\OFFICE11", "Φάκελος Excel")
        ''C:\Program Files\Microsoft Office\OFFICE11
        'returnValue = Shell(f_excelPath + "\EXCEL.EXE", vbMaximizedFocus) ' vbMinimizedNoFocus) ' Run Microsoft Excel.
        '
        'Set myXL = GetObject("", "Excel.Sheet")

138     For k = 1 To 120: m_sumes(k) = 0: SUMES0(k) = 0: Next

        Dim N

        Dim F_T As New ADODB.Recordset

        Dim TT  As Long

140     TT = GetCurrentTime()

        'Dim recs As Long, fp As Long
        'fp = InStr(UCase(sql), "FROM")

        'F_T.Open "select count(*) " + Trim(Mid(sql, fp, 100)), Gdb, adOpenForwardOnly, adLockReadOnly
        'recs = F_T(0)
        'F_T.Close

142     F_T.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

144     N = F_T.FIELDS.Count

146     For k = 1 To N

148         If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then    'DATE
150             f_tab(k) = 2 + f_tab(k - 1) + 8
            Else

152             If F_T.FIELDS(k - 1).DefinedSize > 200 Then
154                 f_tab(k) = 2 + f_tab(k - 1) + 90
                Else
156                 f_tab(k) = 3 + f_tab(k - 1) + F_T.FIELDS(k - 1).DefinedSize
                End If
            End If

        Next

158     LSYN = f_tab(k - 1)

160     mPal = "    "
162     mPAL22 = "     "
164     MOLIS_ALLAJE = 0
166     marxeio2 = "EKT" + xeirisths + ".TXT"

168     m_sthl_ektyp(1) = 0  ' int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )

170     For k = 1 To N
172         AA = LTrim(str(k))
174         m_sthl_ektyp(k + 1) = m_sthl_ektyp(k) + Len(macro("p", AA)) + 1
        Next

        Dim R As New ADODB.Recordset

        On Error GoTo print7_excel_Err

        '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------

176 If F_TYPEHEADER = 1 Then '===============================================
178     AYJ = 5

180     For k = 0 To N - 1
182         myXL.cells(5, k + 1) = F_T(k).Name
184         myXL.cells(5, k + 1).Interior.Color = RGB(180, 280, 180)
        Next
        
186     excel.Windows(1).SplitRow = 5
188     excel.Windows(1).FreezePanes = True

190     myXL.rows(5).Font.Size = 14
192     myXL.rows(5).Font.fontStyle = 12


    Else
194     AYJ = 0
    End If  '=================================================================

        Dim LAST_TIMH

        Dim synt1

196     synt1 = IIf(IsNull("SYNT1"), "true", synt1)   ' όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
       ' AYJ = 5

198     Typose = 0

        Dim row, Pedio_epan As Integer, timh_epan

200     Pedio_epan = GROUPN - 1
        '   If Pedio_epan < 0 Then Pedio_epan = 0
202     timh_epan = "ωωωωωωω"
        
        '     F_T (k)
        Dim synola_SELIDOS

204     synola_SELIDOS = False

        Dim M_SYN(20) As Single

        Dim n2

206     For n2 = 0 To 20: M_SYN(n2) = 0: Next

        '--------------------------------------------------------
208     g_Stop = 1    'entos loop

'οι παρακατω σειρές εγιναν γιατι αλλαξε ο τυπος
'apo varchar(30) se nvarchar(max) για να δουλευει και κυριλλικά
'το προβλημα ηταν ότι στα κυριλλικα εβγαζε κενα
'==================================
F_T.Close
FF.Adodc1.ConnectionString = gConnect
FF.Adodc1.RecordSource = sql
FF.Adodc1.Refresh
Dim RR As Object
Set RR = FF.Adodc1.Recordset
'==================================
210     Do While Not RR.EOF
212         AYJ = AYJ + 1

214         If FF.CancelButton.Enabled = False Then

                Exit Do

            End If
            '----------------- ΤΥΠΩΣΕ ΟΛΑ ΤΑ ΠΕΔΙΑ ---------------------
216         For k = 0 To N - 1
218             If k = Pedio_epan Then
220                 If timh_epan = RR(k) Then
                        ' MHN TYPVNEIS TO PEDIO XANA
                    Else ' ΑΛΛΑΓΗ GROUP
222                     AYJ = AYJ + 1
224                     GoSub PRINTGROUP
                    End If
                End If
              
226             If k = Pedio_epan And timh_epan = RR(k) Then
                Else

228                 If Left(RR(k), 3) = "@@@" Then    '  A/A
230                     myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(AYJ - 5, "######"), RR(k).DefinedSize)
232                 ElseIf RR(k).Type = 7 Or RR(k).Type = 5 Or RR(k).Type = 3 Or RR(k).Type = 131 Then   'IsNumeric(RR(K)) πραγματικο

234                     If IsNull(RR(k)) Then
236                         'myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(0, "######,##0.00"), 13)  '  RR(K).DefinedSize)
                        Else
238                         myXL.cells(AYJ, k + 1) = RR(k) '  Format(RR(k), "######,##0.00") ', 13) ' RR(k) ' Right(Space(30) + Format(RR(k), "######,##0.00"), 13)    '  RR(K).DefinedSize)
                        End If

240                 ElseIf RR(k).Type = 8 Or RR(k).Type = 135 Then   'DATE
242                     myXL.cells(AYJ, k + 1) = RR(k) 'Right(Space(30) + Format(RR(k), "DD/MM/YYYY"), 10)
                    Else    ' 10 STRING

244                     If IsNull(RR(k)) Then

                        Else

                            ' Print #1, Tab(f_tab(K)); to928(F_T(K));
246                         If k < N - 1 Then
                                ' για να μην παταει στην επόμενη στήλη
248                             myXL.cells(AYJ, k + 1) = "'" + RR(k) ' Left(RR(k), f_tab(k + 1) - f_tab(k) - 1)
                            Else
250                             myXL.cells(AYJ, k + 1) = "'" + RR(k)    'to928(RR(K))
                            End If
                        End If

252                     If m_sthl_ektyp(k) > 2 Then
                            ' If K = 1 Then Print #1,
                        End If  'm_sthl_ektyp(K) > 2

                    End If
                End If
              
                ' soymes---------------------
254             If mID$(SUgm_str, k + 1, 1) = "1" Or mID$(SUgm_str, k + 1, 1) = "2" Then
256                 If IsNull(RR(k)) Then

                    Else
258                     m_sumes(k) = m_sumes(k) + Val(Replace(RR(k), ",", "."))
                    End If
                End If    ' mid$(sugm_str,k,1)

260             If AYJ Mod 2 = 0 And F_XROMATA = 1 Then
262                 myXL.cells(AYJ, k + 1).Interior.Color = RGB(200, 300, 200)
                End If

264             If Pedio_epan >= 0 Then
266                 timh_epan = RR(Pedio_epan)
                End If

            Next

268         RR.MoveNext

270         If AYJ Mod 100 = 0 Then

272             DoEvents
274             FF.Caption = "Εγγραφή " + Format(AYJ, "######")    ' + "/" + Format(RR.RecordCount, "######")
            End If

        Loop

276     Unload FF
                          
278     If GROUPN > 0 Then
280         AYJ = AYJ + 2
282         GoSub PRINTGROUP
                          
            '                          For n2 = 0 To n - 1
            '                               If m_sumes(n2) > 0 Then
            '                                  If mID$(SUgm_str, n2 + 1, 1) = "1" Then    ' SUMA
            '                                    myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2), "######,##0.00"), RR(n2).DefinedSize + 2)
            '                                     myXL.cells(AYJ - 1, n2).Interior.Color = RGB(100, 100, 500)
            '                                  Else
            '                                    myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2) / (AYJ - 7), "######,##0.00"), RR(n2).DefinedSize + 2)
            '                                  End If
            '                               Else
            '                                  myXL.cells(AYJ - 1, n2 + 1) = ""  ' Right(Space(50), RR(K).DefinedSize);
            '                               End If
            '                               M_SYN(n2) = m_sumes(n2)
            '                           Next
        End If

284     g_Stop = 0    'adiaforo
        '"
286     AYJ = AYJ + 1

        'myXL.cells(AYJ, 1) = String$(LSYN - 2, "-")
        '"
        ' Print #1, Chr(13)  ' 6/12/2007
        '"
        '   aa = f_kodik("sum_seltxt")
        Dim PARAM    '

288     PARAM = IIf(AA = " ", "  ", macro(AA, 0))
        'pr_SUMselidas param
        'PRINT2_Xsumes "ΓΕΝΙΚΟ ΣΥΝΟΛΟ"
290     AYJ = AYJ + 1
292     GoSub printSUM

        'AppActivate returnValue

        ' On Error Resume Next

294     myXL.columns("A:K").Select
296     myXL.columns.AutoFit

        myXL.columns("A:K").Select
        myXL.columns("a:k").NumberFormat = "0.00"

        'Dim R As New ADODB.Recordset
298    R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
300    k = 1

302 If F_TYPEHEADER = 1 Then

304    With myXL
306        .cells(1, 1) = R("pelono")
308        .cells(2, 1) = R("pelepa")
310        .cells(3, 1) = Now
312        .cells(4, 1) = EPIKEF
        End With

    End If



314    MDIForm1.MousePointer = vbNormal
        'excel.Visible = True

316    excel.Visible = True

        On Error Resume Next

318    Kill "C:\EKTYP.XLS"

320    DoEvents

322    myXL.SaveAs "C:\EKTYP.XLS"

        Dim ANS3 As Long

324    ANS3 = MsgBox("Κλείνω το EXCEL", vbYesNo)

326    If ANS3 = vbYes Then
328        Call workbook.Close(False)
330        excel.Quit
332        Set excel = Nothing
        End If

        Exit Function

PRINTGROUP:

334     For n2 = 0 To N - 1

336         If m_sumes(n2) > 0 Then
338             If mID$(SUgm_str, n2 + 1, 1) = "1" Then    ' SUMA
340                ' myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2), "######,##0.00"), RR(n2).DefinedSize + 2)
                    ' αντικατασταθηκε με αριθμητικό
                    myXL.cells(AYJ - 1, n2 + 1) = Round(m_sumes(n2) - M_SYN(n2), 2)

342                 myXL.cells(AYJ - 1, n2 + 1).Interior.Color = RGB(100, 300, 500)
                    
                Else
344                 myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2) / (AYJ - 7), "######,##0.00"), RR(n2).DefinedSize + 2)
                End If

            Else
346             myXL.cells(AYJ - 1, n2 + 1) = ""  ' Right(Space(50), RR(K).DefinedSize);
            End If

348         M_SYN(n2) = m_sumes(n2)
        Next

350     Return

printSUM:

352    For k = 0 To N - 1

354        If m_sumes(k) > 0 Then
356            If mID$(SUgm_str, k + 1, 1) = "1" Then    ' SUMA
358                myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(m_sumes(k), "######,##0.00"), RR(k).DefinedSize + 2)
                Else
360                myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(m_sumes(k) / (AYJ - 7), "######,##0.00"), RR(k).DefinedSize + 2)
                End If

            Else
362            myXL.cells(AYJ, k + 1) = ""    ' Right(Space(50), RR(K).DefinedSize);
            End If

        Next

364    Return

printEpik:

366    For k = 0 To N - 1

368        If m_sumes(k + 1) > 0 Then
370            myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(m_sumes(k + 1), "######,##0.00"), RR(k).DefinedSize + 2)
            Else
372            myXL.cells(AYJ, k + 1) = Right(Space(50), RR(k).DefinedSize)
            End If

        Next

374    Return


        '<EhFooter>
        Exit Function

print7_excel_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.print7_excel " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function




Function oldprint7_excel(sql As String, _
                      SUgm_str, _
                      ByVal EPIKEF As String, _
                      ByVal GROUPN As Integer)
'10/12/2014
'οι παρακατω σειρές εγιναν γιατι αλλαξε ο τυπος
'apo varchar(30) se nvarchar(max) για να δουλευει και κυριλλικά
'το προβλημα ηταν ότι στα κυριλλικα εβγαζε κενα
'==================================
'F_T.Close
'FF.Adodc1.ConnectionString = gConnect
'FF.Adodc1.RecordSource = sql
'FF.Adodc1.Refresh
'Dim RR As Object
'Set RR = FF.Adodc1.Recordset
'==================================
        
        
        
        
        
        
        
        '<EhHeader>
        On Error GoTo print7_excel_Err
        '</EhHeader>

        '"*********************** P R I N T _ X A R ************************
        '"**************      κάνει παρουσίαση αρχείου στο χαρτί
        '"*** n=αριθμός fields που παρουσιάζονται
        '"*** synuhkh2 η συνθήκη για το  IF, sugm_str βλέπει που θα κάνει σούμες ,π.χ. "00100" κάνει σούμες στο 3ο field
        '"*** sum_pic  το picture γιά τις σούμες
        '"*** Ei,Fi,Pi  :Επικεφαλίδα παρουσίασης,Fields που παρουσιάζονται,Picture παρουσίασης
        '"** synt_eject:=0  αλλάζει σελίδα όταν μεταβάλλεται η στήλη
        Dim MFIELDS(120), mSYN

        Dim synt_eject

        Dim sumes, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), F(120)

        Dim mBSEIRA

        Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE

        Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4

        Dim aaP, aaP2, ektypoths

        Dim EPIK, xeirisths, PPF, epik3

        Dim Sthles, kw, PPD, SELIS

        Dim LSYN, AA, epik2, i As Integer

        Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject

        Dim m_sumes(120), SUMES0(120)

        Dim SS

        Dim m_ekt As Integer

        Dim DUM

        Dim returnValue

        Dim f_excelPath

        'GROUPN = GROUPN - 1

        'ETSI DEN DOYLEYEI ME OLA TA EXCEL ALLA SE BOHTHAEI STON KODIKA

        '    Dim excel As excel.Application
        '   Dim workbook As excel.workbook
        '  Dim myXL As excel.Worksheet

        'ETSI DOYLEYEI ME OLA TA EXCEL
        Dim excel    As excel.Application

        Dim workbook As excel.workbook

        Dim myXL     As excel.Worksheet

100     Set excel = New excel.Application   'CreateObject("excel.Application")
    
        ' Excel.Visible = True
102     Set workbook = excel.Workbooks.Add

        On Error Resume Next

        Dim F_XROMATA As Integer

104     F_XROMATA = Val(FINDPARAMETROI(1, "MDIFORM1", "F_XROMATA", "0", "ΧΡΩΜΑΤΑ ΣΤΟ EXCEL=1 ΟΧΙ=0"))

        '"MDIFORM1"

        '  If (MenuShow.Caption = "&Show") Then
        '    MenuShow.Caption = "&Hide"
106     workbook.Activate

108     Set myXL = workbook.ActiveSheet

        Dim FF As New UDialog

110     FF.SHOW
112     FF.OKButton.Visible = False
114     FF.List1.Visible = False

116     FF.CancelButton.Caption = "ΔΙΑΚΟΠΗ"
118     FF.CancelButton.Top = 120
120     FF.CancelButton.Left = 120
122     FF.CancelButton.Width = 2895
124     FF.CancelButton.Height = 495

        'FF.Top = 3000
        'FF.Left = 3000

126     FF.Width = 3210
128     FF.Height = 810

130     FF.Left = Screen.Width / 2 - FF.Width / 2
132     FF.Top = Screen.Height / 2 - FF.Height / 2

        'FF.Top = MDIForm1.Top + (MDIForm1.height) / 2 ' Command2.Top
        'FF.Left = MDIForm1.Left + (MDIForm1.width) / 2 ' Command2.Left

134     FF.Caption = "ΔΙΑΔΙΚΑΣΙΑ ΥΠΟΛΟΓΙΣΜΟΥ"

        ''------------------------ΠΕΤΑΕΙ ΟΛΟΚΛΗΡΟ ΤΟ ΑΡΧΕΙΟ ΑΛΛΑ ΚΟΛΑΕΙ ΣΤΗΝ [ΧΟΝΔ.ΤΙΜΗ]------------------------------
        'Dim db1 As Database
        'Dim FROM As Long
        'On Error Resume Next
        '  Kill "C:\EKTYP.XLS"
        '  DoEvents
        '  myXL.SaveAs "C:\EKTYP.XLS"
        '
        '  Call workbook.Close(False)
        '  excel.Quit
        '  Set excel = Nothing
        '
        'FROM = InStr(1, UCase(sql), "FROM")
        'Set db1 = OpenDatabase("", False, False, gConnect)
        'db1.Execute Left(sql, FROM - 1) + " " + " into 32 in 'c:\EKTYP.xls' 'Excel 8.0;' " + Trim(Mid(sql, FROM, 500))
        ''------------------------------------------------------

136     MDIForm1.MousePointer = vbHourglass
        'excel.Visible = True

        'f_excelPath = FindParametroi(1,"MDIFORM1", "f_excelPath", "C:\Program Files\Microsoft Office\OFFICE11", "Φάκελος Excel")
        ''C:\Program Files\Microsoft Office\OFFICE11
        'returnValue = Shell(f_excelPath + "\EXCEL.EXE", vbMaximizedFocus) ' vbMinimizedNoFocus) ' Run Microsoft Excel.
        '
        'Set myXL = GetObject("", "Excel.Sheet")

138     For k = 1 To 120: m_sumes(k) = 0: SUMES0(k) = 0: Next

        Dim N

        Dim F_T As New ADODB.Recordset

        Dim TT  As Long

140     TT = GetCurrentTime()

        'Dim recs As Long, fp As Long
        'fp = InStr(UCase(sql), "FROM")

        'F_T.Open "select count(*) " + Trim(Mid(sql, fp, 100)), Gdb, adOpenForwardOnly, adLockReadOnly
        'recs = F_T(0)
        'F_T.Close

142     F_T.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

144     N = F_T.FIELDS.Count

146     For k = 1 To N

148         If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then    'DATE
150             f_tab(k) = 2 + f_tab(k - 1) + 8
            Else

152             If F_T.FIELDS(k - 1).DefinedSize > 200 Then
154                 f_tab(k) = 2 + f_tab(k - 1) + 90
                Else
156                 f_tab(k) = 3 + f_tab(k - 1) + F_T.FIELDS(k - 1).DefinedSize
                End If
            End If

        Next

158     LSYN = f_tab(k - 1)

160     mPal = "    "
162     mPAL22 = "     "
164     MOLIS_ALLAJE = 0

        'If IsNull(EPIK) Then
        '   EPIK = Format(Date, "dd/mm/yyyy")
        'End If

166     marxeio2 = "EKT" + xeirisths + ".TXT"
        'On Error Resume Next

168     m_sthl_ektyp(1) = 0  ' int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )

170     For k = 1 To N
172         AA = LTrim(str(k))
174         m_sthl_ektyp(k + 1) = m_sthl_ektyp(k) + Len(macro("p", AA)) + 1
        Next

        Dim R As New ADODB.Recordset

        'R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
        '        k = 1
        '        With myXL
        '            .cells(1, 3) = R("pelono")
        '            .cells(2, 3) = R("pelepa")
        '            .cells(3, 3) = Now
        '            .cells(4, 3) = EPIKEF
        '        End With
        '
        On Error GoTo print7_excel_Err

        '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------

176 If F_TYPEHEADER = 1 Then '===============================================
178     AYJ = 5

180     For k = 0 To N - 1
182         myXL.cells(5, k + 1) = F_T(k).Name
184         myXL.cells(5, k + 1).Interior.Color = RGB(180, 180, 180)
        Next
        
186     excel.Windows(1).SplitRow = 5
188     excel.Windows(1).FreezePanes = True

190     myXL.rows(5).Font.Size = 14
192     myXL.rows(5).Font.fontStyle = 12


    Else
194     AYJ = 0
    End If  '=================================================================




        'excel.Windows(1).SplitColumn = 1

        '  ROWS("6:6").Select
        '    ActiveWindow.FreezePanes = True
        '    Range("B14").Select
        '    ActiveWindow.SmallScroll Down:=0
        '    ROWS("8:8").Select
        '    With selection.Interior
        '        .Pattern = xlSolid
        '        .PatternColorIndex = xlAutomatic
        '        .ThemeColor = xlThemeColorAccent3
        '        .TintAndShade = 0.599993896298105
        '        .PatternTintAndShade = 0
        '    End With

        Dim LAST_TIMH

        Dim synt1

196     synt1 = IIf(IsNull("SYNT1"), "true", synt1)   ' όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
       ' AYJ = 5

198     Typose = 0

        Dim row, Pedio_epan As Integer, timh_epan

200     Pedio_epan = GROUPN - 1
        '   If Pedio_epan < 0 Then Pedio_epan = 0
202     timh_epan = "ωωωωωωω"
        
        '     F_T (k)
        Dim synola_SELIDOS

204     synola_SELIDOS = False

        Dim M_SYN(20) As Single

        Dim n2

206     For n2 = 0 To 20: M_SYN(n2) = 0: Next

        '--------------------------------------------------------
208     g_Stop = 1    'entos loop

'οι παρακατω σειρές εγιναν γιατι αλλαξε ο τυπος
'apo varchar(30) se nvarchar(max) για να δουλευει και κυριλλικά
'το προβλημα ηταν ότι στα κυριλλικα εβγαζε κενα
'==================================
F_T.Close
FF.Adodc1.ConnectionString = gConnect
FF.Adodc1.RecordSource = sql
FF.Adodc1.Refresh
Dim RR As Object
Set RR = FF.Adodc1.Recordset
'==================================



'Dim RR As New ADODB.Recordset
'RR.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

210     Do While Not RR.EOF
212         AYJ = AYJ + 1

214         If FF.CancelButton.Enabled = False Then

                Exit Do

            End If

            'If g_Stop = 2 Then
            '   Exit Do
            'End If

            'MDIForm1.Caption = AYJ
            '----------------- ΤΥΠΩΣΕ ΟΛΑ ΤΑ ΠΕΔΙΑ ---------------------
216         For k = 0 To N - 1
               ' If IsNull(F_T(k)) Then
              '     MsgBox "null"
               ' End If
                'If k = Pedio_epan And timh_epan = F_T(k) Then
218             If k = Pedio_epan Then
220                 If timh_epan = RR(k) Then
                        ' MHN TYPVNEIS TO PEDIO XANA
                    Else ' ΑΛΛΑΓΗ GROUP
222                     AYJ = AYJ + 1
224                     GoSub PRINTGROUP
                        '                          For n2 = 0 To n - 1
                        '
                        '                               If m_sumes(n2) > 0 Then
                        '                                  If mID$(SUgm_str, n2 + 1, 1) = "1" Then    ' SUMA
                        '                                    myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2), "######,##0.00"), F_T(n2).DefinedSize + 2)
                        '                                     myXL.cells(AYJ - 1, n2).Interior.Color = RGB(100, 100, 500)
                        '                                  Else
                        '                                    myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2) / (AYJ - 7), "######,##0.00"), F_T(n2).DefinedSize + 2)
                        '                                  End If
                        '
                        '                               Else
                        '                                  myXL.cells(AYJ - 1, n2 + 1) = ""  ' Right(Space(50), F_T(K).DefinedSize);
                        '                               End If
                        '
                        '                               M_SYN(n2) = m_sumes(n2)
                        '
                        '                           Next

                    End If
                End If
              
226             If k = Pedio_epan And timh_epan = RR(k) Then
                Else

228                 If Left(RR(k), 3) = "@@@" Then    '  A/A
230                     myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(AYJ - 5, "######"), RR(k).DefinedSize)
232                 ElseIf RR(k).Type = 7 Or RR(k).Type = 5 Or RR(k).Type = 3 Or RR(k).Type = 131 Then   'IsNumeric(RR(K)) πραγματικο

234                     If IsNull(RR(k)) Then
236                         myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(0, "######,##0.00"), 13)  '  RR(K).DefinedSize)
                        Else
238                         myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(RR(k), "######,##0.00"), 13)    '  RR(K).DefinedSize)
                        End If

240                 ElseIf RR(k).Type = 8 Or RR(k).Type = 135 Then   'DATE
242                     myXL.cells(AYJ, k + 1) = RR(k) 'Right(Space(30) + Format(RR(k), "DD/MM/YYYY"), 10)
                    Else    ' 10 STRING

244                     If IsNull(RR(k)) Then

                        Else

                            ' Print #1, Tab(f_tab(K)); to928(F_T(K));
246                         If k < N - 1 Then
                                ' για να μην παταει στην επόμενη στήλη
248                             myXL.cells(AYJ, k + 1) = "'" + Left(RR(k), f_tab(k + 1) - f_tab(k) - 1)
                            Else
250                             myXL.cells(AYJ, k + 1) = "'" + RR(k)    'to928(RR(K))
                            End If
                        End If

252                     If m_sthl_ektyp(k) > 2 Then
                            ' If K = 1 Then Print #1,
                        End If  'm_sthl_ektyp(K) > 2

                    End If
                End If
              
                ' soymes---------------------
254             If mID$(SUgm_str, k + 1, 1) = "1" Or mID$(SUgm_str, k + 1, 1) = "2" Then
256                 If IsNull(RR(k)) Then

                    Else
258                     m_sumes(k) = m_sumes(k) + Val(Replace(RR(k), ",", "."))
                    End If
                End If    ' mid$(sugm_str,k,1)

260             If AYJ Mod 2 = 0 And F_XROMATA = 1 Then
262                 myXL.cells(AYJ, k + 1).Interior.Color = RGB(200, 200, 200)
                End If

264             If Pedio_epan >= 0 Then
266                 timh_epan = RR(Pedio_epan)
                End If

            Next

            ' Print #1,
            '   If GROUPN > 0 Then
            '       LAST_TIMH = RR(GROUPN)
            '   End If
268         RR.MoveNext
            '   If Not RR.EOF Then
            '      If GROUPN > 0 Then
            '          If LAST_TIMH <> RR(GROUPN) Then
            '              AYJ = AYJ + 1
            '          End If
            '      End If
            '  End If
                   
            ' αν τελειωσε να κανει μερικη σουμα στο τελευταιο group
            '                   If RR.EOF Then
            '                          For n2 = 0 To n - 1
            '                               If m_sumes(n2) > 0 Then
            '                                  If mID$(SUgm_str, n2 + 1, 1) = "1" Then    ' SUMA
            '                                    myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2), "######,##0.00"), RR(n2).DefinedSize + 2)
            '                                     myXL.cells(AYJ - 1, n2).Interior.Color = RGB(100, 100, 500)
            '                                  Else
            '                                    myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2) / (AYJ - 7), "######,##0.00"), RR(n2).DefinedSize + 2)
            '                                  End If
            '
            '                               Else
            '                                  myXL.cells(AYJ - 1, n2 + 1) = ""  ' Right(Space(50), RR(K).DefinedSize);
            '                               End If
            '
            '                               M_SYN(n2) = m_sumes(n2)
            '
            '                           Next
            '                     End If

270         If AYJ Mod 100 = 0 Then

272             DoEvents
274             FF.Caption = "Εγγραφή " + Format(AYJ, "######")    ' + "/" + Format(RR.RecordCount, "######")
            End If

        Loop

276     Unload FF
                          
278     If GROUPN > 0 Then
280         AYJ = AYJ + 2
282         GoSub PRINTGROUP
                          
            '                          For n2 = 0 To n - 1
            '                               If m_sumes(n2) > 0 Then
            '                                  If mID$(SUgm_str, n2 + 1, 1) = "1" Then    ' SUMA
            '                                    myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2), "######,##0.00"), RR(n2).DefinedSize + 2)
            '                                     myXL.cells(AYJ - 1, n2).Interior.Color = RGB(100, 100, 500)
            '                                  Else
            '                                    myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2) / (AYJ - 7), "######,##0.00"), RR(n2).DefinedSize + 2)
            '                                  End If
            '                               Else
            '                                  myXL.cells(AYJ - 1, n2 + 1) = ""  ' Right(Space(50), RR(K).DefinedSize);
            '                               End If
            '                               M_SYN(n2) = m_sumes(n2)
            '                           Next
        End If

284     g_Stop = 0    'adiaforo
        '"
286     AYJ = AYJ + 1

        'myXL.cells(AYJ, 1) = String$(LSYN - 2, "-")
        '"
        ' Print #1, Chr(13)  ' 6/12/2007
        '"
        '   aa = f_kodik("sum_seltxt")
        Dim PARAM    '

288     PARAM = IIf(AA = " ", "  ", macro(AA, 0))
        'pr_SUMselidas param
        'PRINT2_Xsumes "ΓΕΝΙΚΟ ΣΥΝΟΛΟ"
290     AYJ = AYJ + 1
292     GoSub printSUM

        'AppActivate returnValue

        ' On Error Resume Next

294     myXL.columns("A:K").Select
296     myXL.columns.AutoFit






        'Dim R As New ADODB.Recordset
298    R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
300    k = 1

302 If F_TYPEHEADER = 1 Then

304    With myXL
306        .cells(1, 1) = R("pelono")
308        .cells(2, 1) = R("pelepa")
310        .cells(3, 1) = Now
312        .cells(4, 1) = EPIKEF
        End With

    End If



314    MDIForm1.MousePointer = vbNormal
        'excel.Visible = True

316    excel.Visible = True

        On Error Resume Next

318    Kill "C:\EKTYP.XLS"

320    DoEvents

322    myXL.SaveAs "C:\EKTYP.XLS"

        Dim ANS3 As Long

324    ANS3 = MsgBox("Κλείνω το EXCEL", vbYesNo)

326    If ANS3 = vbYes Then
328        Call workbook.Close(False)
330        excel.Quit
332        Set excel = Nothing
        End If

        Exit Function

PRINTGROUP:

334     For n2 = 0 To N - 1

336         If m_sumes(n2) > 0 Then
338             If mID$(SUgm_str, n2 + 1, 1) = "1" Then    ' SUMA
340                 myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2), "######,##0.00"), RR(n2).DefinedSize + 2)
342                 myXL.cells(AYJ - 1, n2 + 1).Interior.Color = RGB(100, 100, 500)
                Else
344                 myXL.cells(AYJ - 1, n2 + 1) = Right(Space(30) + Format(m_sumes(n2) - M_SYN(n2) / (AYJ - 7), "######,##0.00"), RR(n2).DefinedSize + 2)
                End If

            Else
346             myXL.cells(AYJ - 1, n2 + 1) = ""  ' Right(Space(50), RR(K).DefinedSize);
            End If

348         M_SYN(n2) = m_sumes(n2)
        Next

350     Return

printSUM:

352    For k = 0 To N - 1

354        If m_sumes(k) > 0 Then
356            If mID$(SUgm_str, k + 1, 1) = "1" Then    ' SUMA
358                myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(m_sumes(k), "######,##0.00"), RR(k).DefinedSize + 2)
                Else
360                myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(m_sumes(k) / (AYJ - 7), "######,##0.00"), RR(k).DefinedSize + 2)
                End If

            Else
362            myXL.cells(AYJ, k + 1) = ""    ' Right(Space(50), RR(K).DefinedSize);
            End If

        Next

364    Return

printEpik:

366    For k = 0 To N - 1

368        If m_sumes(k + 1) > 0 Then
370            myXL.cells(AYJ, k + 1) = Right(Space(30) + Format(m_sumes(k + 1), "######,##0.00"), RR(k).DefinedSize + 2)
            Else
372            myXL.cells(AYJ, k + 1) = Right(Space(50), RR(k).DefinedSize)
            End If

        Next

374    Return


        '<EhFooter>
        Exit Function

print7_excel_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.print7_excel " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function




Function print_WORD(sql As String, SUgm_str, ByVal EPIKEF As String, GROUPN As Integer)
        '<EhHeader>
        On Error GoTo print_WORD_Err
        '</EhHeader>

        '"*********************** P R I N T _ X A R ************************
        '"**************      κάνει παρουσίαση αρχείου στο χαρτί
        '"*** n=αριθμός fields που παρουσιάζονται
        '"*** synuhkh2 η συνθήκη για το  IF, sugm_str βλέπει που θα κάνει σούμες ,π.χ. "00100" κάνει σούμες στο 3ο field
        '"*** sum_pic  το picture γιά τις σούμες
        '"*** Ei,Fi,Pi  :Επικεφαλίδα παρουσίασης,Fields που παρουσιάζονται,Picture παρουσίασης
        '"** synt_eject:=0  αλλάζει σελίδα όταν μεταβάλλεται η στήλη
        Dim MFIELDS(120), mSYN

        Dim synt_eject

        Dim sumes, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), F(120)

        Dim mBSEIRA

        Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE

        Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4

        Dim aaP, aaP2, ektypoths

        Dim EPIK, xeirisths, PPF, epik3

        Dim Sthles, kw, PPD, SELIS

        Dim LSYN, AA, epik2, i As Integer

        Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject

        Dim m_sumes(120), SUMES0(120)

        Dim SS

        Dim m_ekt As Integer

        Dim DUM

        Dim returnValue

        'Dim myXL As Object
        Dim f_excelPath

        Dim strInsertText

        Dim strColumnWidths(10)

        Dim intIdx

        Dim mWordA   As Word.Application

        Dim mobjWORD As Word.Document

        'Start a new document in Word
100     Set mWordA = CreateObject("Word.Application")
102     Set mobjWORD = mWordA.Documents.Add
        'DOK1.Activate
104     mWordA.Visible = True

        Dim oPara1 As Word.Paragraph

        'Insert a paragraph at the beginning of the document.
106     Set oPara1 = mobjWORD.content.Paragraphs.Add
108     oPara1.range.Text = EPIKEF
110     oPara1.range.Font.Bold = True
        ' mWordA.Selection.ParagraphFo

112     oPara1.Format.SpaceAfter = 24    '24 pt spacing after paragraph.
114     mWordA.selection.ParagraphFormat.alignment = wdAlignParagraphCenter

116     oPara1.range.InsertParagraphAfter
118     oPara1.range.InsertParagraphAfter
120     oPara1.range.InsertParagraphAfter
122     oPara1.range.InsertParagraphAfter

124     mWordA.selection.MoveDown Unit:=wdLine, Count:=2    ' METAKINHSH 2  PARAGRAFOYS PARAKATO

        'ok
        'Dim mobjWord As Object  ' Word.Document
        'Set mobjWord = GetObject("", "Word.Basic")
        'mobjWord.AppMinimize ("Microsoft Word")
        '    mobjWord.FileNew
        'ok

        'Start printing the report
        ' Status "Adding data to report"
        Dim L As Integer

        'Set DOK1 = Word.Documents.Add
        'DOK1.Activate

        'Dim Excel As Excel.Application
        'Dim workbook As Excel.workbook
        'Dim myXL As Excel.Worksheet
        '
        '  Set Excel = New Excel.Application
        '  Set workbook = Excel.Workbooks.Add
        '

        '    workbook.Activate

        'Set myXL = workbook.ActiveSheet

126     For k = 1 To 120: m_sumes(k) = 0: SUMES0(k) = 0: Next

        Dim MYRANGE

        Dim N

        Dim F_T As New ADODB.Recordset

        Dim TT  As Long

128     TT = GetCurrentTime()

130     F_T.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

132     N = F_T.FIELDS.Count

        'Set MYRANGE = mobjWORD.Range(start:=0, End:=0)
134     Set MYRANGE = mobjWORD.Paragraphs(3).range
136     mobjWORD.Tables.Add range:=MYRANGE, NumRows:=3, NumColumns:=N

        'mobjWORD.TableInsertTable NumColumns:=N, _
        '   NumRows:=2, _
        '  InitialColWidth:="2 cm"

138     For k = 1 To N

140         If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then    'DATE
142             f_tab(k) = 2 + f_tab(k - 1) + 8
            Else

144             If F_T.FIELDS(k - 1).DefinedSize > 200 Then
146                 f_tab(k) = 2 + f_tab(k - 1) + 13
                Else
148                 f_tab(k) = 3 + f_tab(k - 1) + F_T.FIELDS(k - 1).DefinedSize
                End If
            End If

        Next

150     LSYN = f_tab(k - 1)

152     mPal = "    "
154     mPAL22 = "     "
156     MOLIS_ALLAJE = 0

        'If IsNull(EPIK) Then
        '   EPIK = Format(Date, "dd/mm/yyyy")
        'End If

158     marxeio2 = "EKT" + xeirisths + ".TXT"

        On Error Resume Next

160     m_sthl_ektyp(1) = 0  ' int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )

162     For k = 1 To N
164         AA = LTrim(str(k))
166         m_sthl_ektyp(k + 1) = m_sthl_ektyp(k) + 7    ' Len(macro("p", aa)) + 1
        Next

        Dim R As New ADODB.Recordset

168     R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
170     k = 1

        'Format the paragraph height
172     mobjWORD.TableSelectTable
174     mobjWORD.FormatParagraph Before:="6 pt"

        '
        '    For K = 0 To N - 1
        '        With mobjWord
        '            .TableSelectColumn
        '            .TableColumnWidth ColumnWidth:=(F_T.FIELDS(K).DefinedSize) / 3
        '            .NextCell
        '            .NextCell
        '        End With
        '    Next
        '

        On Error GoTo print_WORD_Err

        '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------

176     With mobjWORD
178         .Tables(1).Select
180         .Tables(1).rows.Add

            On Error Resume Next

182         For k = 0 To N - 1
                '.Insert F_T(K).Name
                '.NextCell

184             .Tables(1).Cell(1, k + 1).range.Text = F_T(k).Name
            Next

        End With

        ' myXL.rows(5).Font.Size = 14
        ' myXL.rows(5).Font.fontStyle = 12

        '  With myXL
        '     On Error GoTo print_WORD_Err
        '       rows("5:5").Select
        '       With selection.Font
        '         .Name = "Arial"
        '         .fontStyle = "Έντονα"
        '         .Size = 12
        '       End With
        '  End With

        Dim LAST_TIMH

        Dim synt1

186     synt1 = IIf(IsNull("SYNT1"), "true", synt1)   ' όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
188     AYJ = 1

190     Typose = 0

        Dim row

        Dim synola_SELIDOS

192     synola_SELIDOS = False

        '--------------------------------------------------------
194     Do While Not F_T.EOF
196         AYJ = AYJ + 1
198         mobjWORD.Tables(1).rows.Add

            '----------------- ΤΥΠΩΣΕ ΟΛΑ ΤΑ ΠΕΔΙΑ ---------------------
200         For k = 0 To N - 1

202             If Left(F_T(k), 3) = "@@@" Then    '  A/A
204                 mobjWORD.Tables(1).Cell(AYJ, k + 1).range.Text = Right(Space(30) + Format(AYJ - 1, "######"), F_T(k).DefinedSize)

                    ' myXL.cells(AYJ, K + 1) = Right(Space(30) + Format(AYJ - 5, "######"), F_T(K).DefinedSize)
206             ElseIf F_T(k).Type = 7 Or F_T(k).Type = 5 Then   'IsNumeric(f_t(K)) πραγματικο
208                 mobjWORD.Tables(1).Cell(AYJ, k + 1).range.Text = Right(Space(30) + Format(F_T(k), "######,##0.00"), F_T(k).DefinedSize)
210             ElseIf F_T(k).Type = 8 Then   'DATE
212                 mobjWORD.Tables(1).Cell(AYJ, k + 1).range.Text = Right(Space(30) + Format(F_T(k), "DD/MM/YYYY"), 10)
                Else    ' 10 STRING

214                 If IsNull(F_T(k)) Then

                    Else

                        ' Print #1, Tab(f_tab(K)); to928(F_T(K));
216                     If k < N - 1 Then
                            ' για να μην παταει στην επόμενη στήλη
218                         mobjWORD.Tables(1).Cell(AYJ, k + 1).range.Text = Left(to928(F_T(k)), f_tab(k + 1) - f_tab(k) - 1)
                        Else
220                         mobjWORD.Tables(1).Cell(AYJ, k + 1).range.Text = to928(F_T(k))
                        End If
                    End If

222                 If m_sthl_ektyp(k) > 2 Then
                        ' If K = 1 Then Print #1,
                    End If  'm_sthl_ektyp(K) > 2

                End If

                ' soymes---------------------
                '   If Mid$(SUgm_str, k + 1, 1) = "1" Then
224             If mID$(SUgm_str, k + 1, 1) = "1" Or mID$(SUgm_str, k + 1, 1) = "2" Then
226                 If IsNull(F_T(k)) Then

                    Else
228                     m_sumes(k) = m_sumes(k) + F_T(k)
                    End If
                End If    ' mid$(sugm_str,k,1)

            Next

            ' Print #1,
230         If GROUPN > 0 Then
232             LAST_TIMH = F_T(GROUPN - 1)
            End If

234         F_T.MoveNext

236         If Not F_T.EOF Then
238             If GROUPN > 0 Then
240                 If LAST_TIMH <> F_T(GROUPN - 1) Then
242                     AYJ = AYJ + 1
                    End If
                End If
            End If

        Loop

        '"
244     AYJ = AYJ + 1

        'myXL.cells(AYJ, 1) = String$(LSYN - 2, "-")
        '"
        ' Print #1, Chr(13)  ' 6/12/2007
        '"
        '   aa = f_kodik("sum_seltxt")
        Dim PARAM    '

246     PARAM = IIf(AA = " ", "  ", AA)
        'pr_SUMselidas param
        'PRINT2_Xsumes "ΓΕΝΙΚΟ ΣΥΝΟΛΟ"
248     AYJ = AYJ + 1
250     GoSub printSUM

        'AppActivate returnValue

        On Error Resume Next

        ' myXL.Columns("A:K").Select
        ' myXL.Columns.AutoFit

        'Excel.Visible = True
252     mobjWORD.Tables(1).AllowAutoFit = True
254     mobjWORD.Tables(1).AutoFormat

256     MsgBox "Κλείνω το WORD"

        ' Call workbook.Close(False)
        'Word.Quit
258     Set mobjWORD = Nothing

        Exit Function

printSUM:

260     For k = 0 To N - 1

262         If m_sumes(k) > 0 Then
264             If mID$(SUgm_str, k + 1, 1) = "1" Then   'ΣΟΥΜΕΣ
266                 mobjWORD.Tables(1).Cell(AYJ, k + 1).range.Text = Right(Space(30) + Format(m_sumes(k), "######,##0.00"), F_T(k).DefinedSize + 2)
                Else  'ΜΕΣ.ΟΡΟΣ
268                 mobjWORD.Tables(1).Cell(AYJ, k + 1).range.Text = Right(Space(30) + Format(m_sumes(k) / (AYJ - 3), "######,##0.00"), F_T(k).DefinedSize + 2)
                End If

            Else
270             mobjWORD.Tables(1).Cell(AYJ, k + 1).range.Text = ""    ' Right(Space(50), F_T(K).DefinedSize);
            End If

272         mobjWORD.NextCell
        Next

274     Return

printEpik:

276     For k = 0 To N - 1

278         If m_sumes(k + 1) > 0 Then
                '  myXL.cells(AYJ, K + 1) = Right(Space(30) + Format(m_sumes(K + 1), "######,##0.00"), F_T(K).DefinedSize + 2)
            Else
                '  myXL.cells(AYJ, K + 1) = Right(Space(50), F_T(K).DefinedSize)
            End If

        Next

280    Return


        '<EhFooter>
        Exit Function

print_WORD_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.print_WORD " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function

Function print4_pdf(sql As String, SUgm_str, EPIKEF As String, GROUPN As Integer)
'https://www.biopdf.com/download.php
        '<EhHeader>
        On Error GoTo print4_xar_Err
        '</EhHeader>

 Dim prtidx As Integer
    Dim sPrinterName As String
    Dim settings As Object
    Dim util As Object
    
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    
    Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.printerName = sPrinterName
    settings.SetValue "Output", "<desktop>\myfile.pdf"
    settings.SetValue "ConfirmOverwrite", "no"
    settings.SetValue "ShowSaveAS", "never"
    settings.SetValue "ShowSettings", "never"
    settings.SetValue "ShowPDF", "yes"
    settings.SetValue "RememberLastFileName", "no"
    settings.SetValue "RememberLastFolderName", "no"
    settings.WriteSettings True
   
    
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then Err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
        
    Rem -- Print something
'    If optOrientation(0).Value Then
'        Printer.Orientation = PrinterObjectConstants.vbPRORPortrait
'    Else
'        Printer.Orientation = PrinterObjectConstants.vbPRORLandscape
'    End If
    
    Rem -- Set paper size
    Rem -- http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.powerpacks.printing.compatibility.vb6.printer.papersize.aspx
    Rem -- Note: Custom paper size is not supported by VB6 after Windows 98.
    Printer.PaperSize = vbPRPSA3


































        '"*********************** P R I N T _ X A R ************************
        '"**************      κάνει παρουσίαση αρχείου στο χαρτί
        '"*** n=αριθμός fields που παρουσιάζονται
        '"*** synuhkh2 η συνθήκη για το  IF, sugm_str βλέπει που θα κάνει σούμες ,π.χ. "00100" κάνει σούμες στο 3ο field
        '"*** sum_pic  το picture γιά τις σούμες
        '"*** Ei,Fi,Pi  :Επικεφαλίδα παρουσίασης,Fields που παρουσιάζονται,Picture παρουσίασης
        '"** synt_eject:=0  αλλάζει σελίδα όταν μεταβάλλεται η στήλη
        Dim MFIELDS(120), mSYN

        Dim synt_eject

        Dim sumes, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), F(120)

        Dim mBSEIRA

        Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE

        Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4

        Dim aaP, aaP2, ektypoths

        Dim EPIK, xeirisths, PPF, epik3

        Dim Sthles, kw, PPD, SELIS

        Dim LSYN, AA, epik2, i As Integer

        Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject

        Dim m_sumes(120), SUMES0(120)

        Dim SS, SEIRES_selidas As Integer

        Dim MEGISTO_PLATOS As Integer

        Dim SELIDA         As Long

100     SELIDA = 0

        Dim metaf As String

102     MEGISTO_PLATOS = 120

104     For k = 1 To 120: m_sumes(k) = 0: SUMES0(k) = 0: Next    'SEIR_SELID1

106     SEIRES_selidas = Val(FINDPARAMETROI(1, "EKTYPOTES", "SEIR_SELID1", "70", "Σειρές ανά σελίδα(κάθετη)"))
108     F_437 = Val(FINDPARAMETROI(1, "EKTYPOTES", "F_437", "0", "1=437 0=928"))
    
        ' f_psifiaAjias = Val(FindParametroi(1,"PAR1", "f_psifiaAjias", "2", "Δεκαδικά Ψηφία Αξίας Σειρών τιμολογιου")) 'posa psifia tha exei h kathe seira

110     fENARJH_APO_SEIRA = Val(FINDPARAMETROI(1, "EKTYPOTES", "ENARJH_APO_SEIRA", "0", "Από ποιό Α/Α να αρχίσω να τυπώνω; 0=Αρχή"))

        Dim DUM, N

        Dim F_T As New ADODB.Recordset

        Dim TT  As Long

112     TT = GetCurrentTime()

114     F_T.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

116     N = F_T.FIELDS.Count

118     For k = 1 To N

120         If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then    'DATE
122             f_tab(k) = 2 + f_tab(k - 1) + 8
            Else

124             If F_T.FIELDS(k - 1).DefinedSize > 200 Then
126                 f_tab(k) = 2 + f_tab(k - 1) + 40
                Else
128                 f_tab(k) = 3 + f_tab(k - 1) + F_T.FIELDS(k - 1).DefinedSize
                End If
            End If

        Next

130     LSYN = f_tab(k - 1)

132     mPal = "    "
134     mPAL22 = "     "
136     MOLIS_ALLAJE = 0

138     If LSYN > MEGISTO_PLATOS Then
140         SEIRES_selidas = 35
142         Printer.Orientation = vbPRORLandscape    '  2   Documents are printed with the top at the wide side of the paper.
        Else
            'TO PAIRNEI APO TO TABLE PARAMETROI
144         Printer.Orientation = vbPRORPortrait  ' 1   Documents are printed with the top at the narrow side of the paper.

        End If

        'If IsNull(EPIK) Then
        '   EPIK = Format(Date, "dd/mm/yyyy")
        'End If

        'marxeio2 = "EKT" + xeirisths + ".TXT"
        On Error Resume Next

        'Close #1
        'Open "c:\print" For Output As #1
        '  LSYN = 1
        '  For K = 1 To n
        '       LSYN = LSYN + 2 + F_T(K - 1).DefinedSize  'Len(macro("p", aa)) + 1
        '  Next

        '  If LSYN > 128 Then
        '      sthles = LSYN + 10
        '  ElseIf LSYN > 80 Then
        '      sthles = LSYN + 10 ' 136
        '  End If

146     m_sthl_ektyp(1) = 0  ' int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )

148     For k = 1 To N
150         AA = LTrim(str(k))

152         m_sthl_ektyp(k + 1) = m_sthl_ektyp(k) + 7    ' Len(macro("p", aa)) + 1
        Next

154     Printer.FontSize = 8
156     Printer.FontName = "Courier New"
158     Printer.Font.Charset = 161
        'Printer.FontBold = True

        ' PRINT2_Xepik

        Dim R As New ADODB.Recordset

160     R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
        'R("pelono") = pel_ono.Text
        'R("pelepa") = pel_epa.Text
        'R("peldie") = pel_die.Text

162     If fENARJH_APO_SEIRA > 1 Then
            ' να τυπώνει ξεκινώντας από συγκεκριμένη γραμμή χωρίς να χάσει τις σούμες
            ' να μην τυπώνει την επικεφαλίδα
164         MsgBox "Προσοχή θα αρχίσει η εκτύπωση από την σειρά " + str(fENARJH_APO_SEIRA) + " και παρακάτω"

        Else
            ' νορμαλ  τυπώνει την επικεφαλίδα
        
166         SELIDA = SELIDA + 1
        
168         Printer.Print F_(R("pelono"))
170         Printer.Print F_(R("pelDIE"))
172         Printer.Print F_(R("pelepa")) + Space(40) + " Σελίς " + str(SELIDA)
            'Printer.Print Now
174         Printer.FontSize = 12
176         Printer.FontBold = True
178         Printer.Print Tab(9 / 12 * (LSYN / 2 - Len(EPIKEF) / 2)); F_(EPIKEF)
180         Printer.FontSize = 9

            '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------
182         For k = 0 To N - 1
184             Printer.Print Tab(f_tab(k)); F_(Left(F_T(k).Name, F_T(k).DefinedSize));
            Next

            'δευτερη σειρα επικεφαλίδας
186         For k = 0 To N - 1
188             Printer.Print Tab(f_tab(k)); F_(mID$(F_T(k).Name, F_T(k).DefinedSize + 1, Len(F_T(k).Name) - F_T(k).DefinedSize + 1));
            Next

190         Printer.FontBold = False
192         Printer.Print Tab(m_sthl_ektyp(1)); String$(LSYN - 2, "-")
        End If

194     For i = 1 To N
196         m_sumes(i) = 0
198         SUMES0(i) = 0
        Next

        Dim LAST_TIMH

        Dim synt1

        Dim maxWidth As Long

200     synt1 = IIf(IsNull("SYNT1"), "true", synt1)   ' όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
202     AYJ = 0
204     Typose = 0

        Dim row

        Dim synola_SELIDOS

206     synola_SELIDOS = True    ' False

        '--------------------------------------------------------
208     Do While Not F_T.EOF
210         AYJ = AYJ + 1

212         If fENARJH_APO_SEIRA > 1 Then

                ' να τυπώνει ξεκινώντας από συγκεκριμένη γραμμή χωρίς να χάσει τις σούμες
                ' χρειάζεται για τις απογραφές
214             Do While AYJ < fENARJH_APO_SEIRA

216                 For k = 0 To N - 1

218                     If mID$(SUgm_str, k + 1, 1) = "1" Or mID$(SUgm_str, k + 1, 1) = "2" Then
220                         If IsNull(F_T(k)) Then
                            Else
222                             m_sumes(k) = m_sumes(k) + F_T(k)
                            End If
                        End If    '

                    Next

224                 F_T.MoveNext
226                 AYJ = AYJ + 1

228                 If F_T.EOF Then
230                     MsgBox "υπέρβαση τέλους αρχείου"

                        Exit Function

                    End If

                Loop

232             fENARJH_APO_SEIRA = 0    ' για να μην μπαινει ξανά στο ιδιο  If fENARJH_APO_SEIRA > 1 Then
            End If

234         If end_page = 2 Then    ' ΤΥΠΩΝΕΙ ΓΡΑΜΜΟΥΛΕΣ --------------------------------
236             Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN, "-")

238             end_page = 0

240             MOLIS_ALLAJE = 1
            End If    ' end_page=2

242         If synola_SELIDOS Then
244             If (AYJ - 1) Mod SEIRES_selidas = 0 And AYJ > 1 Then    ' .AND. PROTH_FORA<>1 )
                    'syn_epik
246                 Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN, "-")
248                 metaf = "Σε μεταφορά "
250                 GoSub printSUM

252                 Printer.NewPage

                    'ξαναδίνω το orientation γιατί το χάνει μετά το endDoc
254                 If LSYN > MEGISTO_PLATOS Then
256                     Printer.Orientation = vbPRORLandscape    '  2   Documents are printed with the top at the wide side of the paper.
                    Else
258                     Printer.Orientation = vbPRORPortrait  ' 1   Documents are printed with the top at the narrow side of the paper.
                    End If

260                 end_page = 0

                    '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------

                    '890                 Printer.Print
                    '900                 Printer.Print
                    '910                 Printer.Print
         
262                 SELIDA = SELIDA + 1
264                 Printer.Print F_(R("pelono"))
266                 Printer.Print F_(R("pelDIE"))
268                 Printer.Print F_(R("pelepa")) + Space(40) + " Σελίς " + str(SELIDA)
                    'Printer.Print Now
270                 Printer.FontSize = 12
272                 Printer.FontBold = True
274                 Printer.Print Tab(9 / 12 * (LSYN / 2 - Len(EPIKEF) / 2)); F_(EPIKEF)
276                 Printer.FontSize = 9

278                 For k = 0 To N - 1
280                     Printer.Print Tab(f_tab(k)); F_(Left(F_T(k).Name, F_T(k).DefinedSize));
                    Next

                    'δευτερη σειρα επικεφαλίδας
282                 For k = 0 To N - 1
284                     Printer.Print Tab(f_tab(k)); F_(mID$(F_T(k).Name, F_T(k).DefinedSize + 1, Len(F_T(k).Name) - F_T(k).DefinedSize + 1));
                    Next

286                 Printer.Print Tab(m_sthl_ektyp(1)); String$(LSYN - 2, "-")

288                 metaf = "Εκ μεταφοράς "
290                 GoSub printSUM

                End If  'end_page = 1 Then
            End If

            '----------------- ΤΥΠΩΣΕ ΟΛΑ ΤΑ ΠΕΔΙΑ ---------------------
292         For k = 0 To N - 1

294            If k < N - 1 Then
                    ' για να μην παταει στην επόμενη στήλη
296                maxWidth = f_tab(k + 1) - f_tab(k) - 1
                End If

298            If Left(F_T(k), 3) = "@@@" Then    '  A/A
300                Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(AYJ, "######"), Min(F_T(k).DefinedSize, maxWidth));
302            ElseIf F_T(k).Type = 7 Or F_T(k).Type = 5 Then   'IsNumeric(f_t(K)) πραγματικο
304                Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(F_T(k), "######,##0.00"), Min(maxWidth, 14));    '  F_T(k).DefinedSize);
306            ElseIf F_T(k).Type = 8 Then   'DATE
308                Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(F_T(k), "DD/MM/YYYY"), Min(maxWidth, 10));
                Else    ' 10 STRING

310                If IsNull(F_T(k)) Then
312                    Printer.Print Tab(f_tab(k));
                    Else

                        ' PRINTER.PRINT Tab(f_tab(K)); to928(F_T(K));
314                    If k < N - 1 Then
                            ' για να μην παταει στην επόμενη στήλη
316                        Printer.Print Tab(f_tab(k)); F_(Left((F_T(k)), maxWidth));
                            'Printer.Print Tab(f_tab(k)); F_(Left(to928(F_T(k)), maxWidth));
                        Else
318                        Printer.Print Tab(f_tab(k)); F_((F_T(k)));
                            'Printer.Print Tab(f_tab(k)); F_(to928(F_T(k)));
                        End If
                    End If

320                If m_sthl_ektyp(k) > 2 Then
322                    If k = 1 Then Printer.Print
                    End If  'm_sthl_ektyp(K) > 2

                End If

                ' soymes---------------------
324            If mID$(SUgm_str, k + 1, 1) = "1" Or mID$(SUgm_str, k + 1, 1) = "2" Then
326                If IsNull(F_T(k)) Then

                    Else
328                    m_sumes(k) = m_sumes(k) + F_T(k)
                    End If
                End If    ' mid$(sugm_str,k,1)

            Next

330        Printer.Print

332        If GROUPN > 0 Then
334            LAST_TIMH = F_T(GROUPN - 1)
            End If

336        F_T.MoveNext

338        If Not F_T.EOF Then
340            If GROUPN > 0 Then
342                If LAST_TIMH <> F_T(GROUPN - 1) Then
344                    Printer.Print
                    End If
                End If
            End If

        Loop

        '"
346    Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN - 2, "-")

        '"
        ' PRINTER.PRINT Chr(13)  ' 6/12/2007
        '"
        '   aa = f_kodik("sum_seltxt")
        Dim PARAM    '

348    PARAM = IIf(AA = " ", "  ", AA)
        'pr_SUMselidas param
        'PRINT2_Xsumes "ΓΕΝΙΚΟ ΣΥΝΟΛΟ"

350    metaf = "ΣΥΝΟΛΑ "

352    GoSub printSUM
354    Printer.Print
356    Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN - 2, "-")

        '   PRINTER.PRINT Chr(12)
        '"
        Dim YPOSEL(10)

358    For k = 1 To 4
360        Printer.Print    '  1, 0; YPOSEL(k)
        Next

        'PRINTER.PRINT Chr(13)

362    If eject = "y" Then
            'PRINTER.PRINT Chr(12)
        End If

        ' αποθηκεύω τις σούμες στα e1,e2,.. για να τις χρησιμοποιώ
        '  For K = 1 To N
        '   aa = "e" + LTrim(Str(K))
        '    macro(aa, 0) = m_sumes(K)
        ' Next

        'Close #1

364    pp = 1
        '   If pp = 1 Then
        '     If ektypoths = -1 Then
        '     Else
        '
        '
        '       PPD = MsgBox("ΤΟ ΒΛΕΠΩ ΠΡΙΝ ΤΥΠΩΘΕΙ", vbYesNo)
        '       If PPD = vbYes Then
        '
        '          On Error GoTo print4_xar_Err
        '           DUM = Shell("c:\mercvb\notepad.exe c:\print", vbMaximizedFocus)
        '
        '           PPF = MsgBox("Προχωρώ στην εκτύπωση", vbYesNo)
        '           If PPF = vbYes Then
        '               DUM = Shell("c:\mercvb\notepad.exe /p c:\print", vbMaximizedFocus)
        '           End If
        '       Else
        '       End If
        '     End If
        '   End If
366    SELIS = 1
368    Printer.EndDoc


  Rem -- Wait for runonce settings file to disappear
    Dim runonce As String
    runonce = settings.GetSettingsFilePath(True)
    While Dir(runonce, vbNormal) <> ""
        Sleep 100
    Wend
    
    MsgBox "myfile.pdf Αποθηκεύτηκε στην Επιφάνεια Εργαίας", vbInformation, "PDF Created"


















        Exit Function

printSUM:

370    Printer.Print Tab(0); F_(Trim(metaf));

372    For k = 1 To N - 1

374        If k < N - 1 Then
                ' για να μην παταει στην επόμενη στήλη
376            maxWidth = f_tab(k + 1) - f_tab(k) - 1
            End If

378        If m_sumes(k) > 0 Then
380            If mID$(SUgm_str, k + 1, 1) = "1" Then  'ΑΘΡΟΙΣΜΑ
382                Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k), "######,##0.00"), Min(maxWidth, 14));   ' F_T(K).DefinedSize));
                Else  ' ΜΕΣΟΣ ΟΡΟΣ
384                Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k) / AYJ, "######,##0.00"), Min(maxWidth, F_T(k).DefinedSize));
                End If

            Else
                '   Printer.Print Tab(f_tab(K)); ""; ' Right(Space(50), F_T(K).DefinedSize);
            End If

        Next

386    Printer.Print
388    Return

printEpik:

390    For k = 0 To N - 1

392        If k < N - 1 Then
                ' για να μην παταει στην επόμενη στήλη
394            maxWidth = f_tab(k + 1) - f_tab(k) - 1
            End If

396        If m_sumes(k + 1) > 0 Then
398            Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k + 1), "######,##0.00"), Min(maxWidth, F_T(k).DefinedSize));
            Else
400            Printer.Print Tab(f_tab(k)); Right(Space(50), Min(maxWidth, F_T(k).DefinedSize));
            End If

        Next

402    Printer.Print
404    Return


        '<EhFooter>
        Exit Function

print4_xar_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.print4_xar " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function



Function print4_xar(sql As String, SUgm_str, EPIKEF As String, GROUPN As Integer)
        '<EhHeader>
        On Error GoTo print4_xar_Err
        '</EhHeader>

        '"*********************** P R I N T _ X A R ************************
        '"**************      κάνει παρουσίαση αρχείου στο χαρτί
        '"*** n=αριθμός fields που παρουσιάζονται
        '"*** synuhkh2 η συνθήκη για το  IF, sugm_str βλέπει που θα κάνει σούμες ,π.χ. "00100" κάνει σούμες στο 3ο field
        '"*** sum_pic  το picture γιά τις σούμες
        '"*** Ei,Fi,Pi  :Επικεφαλίδα παρουσίασης,Fields που παρουσιάζονται,Picture παρουσίασης
        '"** synt_eject:=0  αλλάζει σελίδα όταν μεταβάλλεται η στήλη
        Dim MFIELDS(120), mSYN

        Dim synt_eject

        Dim sumes, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), F(120)

        Dim mBSEIRA

        Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE

        Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4

        Dim aaP, aaP2, ektypoths

        Dim EPIK, xeirisths, PPF, epik3

        Dim Sthles, kw, PPD, SELIS

        Dim LSYN, AA, epik2, i As Integer

        Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject

        Dim m_sumes(120), SUMES0(120)

        Dim SS, SEIRES_selidas As Integer

        Dim MEGISTO_PLATOS As Integer

        Dim SELIDA         As Long

100     SELIDA = 0

        Dim metaf As String

102     MEGISTO_PLATOS = 120

104     For k = 1 To 120: m_sumes(k) = 0: SUMES0(k) = 0: Next    'SEIR_SELID1

106     SEIRES_selidas = Val(FINDPARAMETROI(1, "EKTYPOTES", "SEIR_SELID1", "70", "Σειρές ανά σελίδα(κάθετη)"))
108     F_437 = Val(FINDPARAMETROI(1, "EKTYPOTES", "F_437", "0", "1=437 0=928"))
    
        ' f_psifiaAjias = Val(FindParametroi(1,"PAR1", "f_psifiaAjias", "2", "Δεκαδικά Ψηφία Αξίας Σειρών τιμολογιου")) 'posa psifia tha exei h kathe seira

110     fENARJH_APO_SEIRA = Val(FINDPARAMETROI(1, "EKTYPOTES", "ENARJH_APO_SEIRA", "0", "Από ποιό Α/Α να αρχίσω να τυπώνω; 0=Αρχή"))

        Dim DUM, N

        Dim F_T As New ADODB.Recordset

        Dim TT  As Long

112     TT = GetCurrentTime()

114     F_T.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

116     N = F_T.FIELDS.Count

118     For k = 1 To N

120         If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then    'DATE
122             f_tab(k) = 2 + f_tab(k - 1) + 8
            
             ElseIf F_T(k - 1).Type = 5 Then    '  NUMBER
123             f_tab(k) = 2 + f_tab(k - 1) + 14
            
            
            
            
            Else

124             If F_T.FIELDS(k - 1).DefinedSize > 200 Then
126                 f_tab(k) = 2 + f_tab(k - 1) + 40
                Else
128                 f_tab(k) = 1 + f_tab(k - 1) + F_T.FIELDS(k - 1).DefinedSize
                End If
            End If

        Next

130     LSYN = f_tab(k - 1)

132     mPal = "    "
134     mPAL22 = "     "
136     MOLIS_ALLAJE = 0

138     If LSYN > MEGISTO_PLATOS Then
140         SEIRES_selidas = 45
142         Printer.Orientation = vbPRORLandscape    '  2   Documents are printed with the top at the wide side of the paper.
        Else
            'TO PAIRNEI APO TO TABLE PARAMETROI
144         Printer.Orientation = vbPRORPortrait  ' 1   Documents are printed with the top at the narrow side of the paper.

        End If

        'If IsNull(EPIK) Then
        '   EPIK = Format(Date, "dd/mm/yyyy")
        'End If

        'marxeio2 = "EKT" + xeirisths + ".TXT"
        On Error Resume Next

        'Close #1
        'Open "c:\print" For Output As #1
        '  LSYN = 1
        '  For K = 1 To n
        '       LSYN = LSYN + 2 + F_T(K - 1).DefinedSize  'Len(macro("p", aa)) + 1
        '  Next

        '  If LSYN > 128 Then
        '      sthles = LSYN + 10
        '  ElseIf LSYN > 80 Then
        '      sthles = LSYN + 10 ' 136
        '  End If

146     m_sthl_ektyp(1) = 0  ' int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )

148     For k = 1 To N
150         AA = LTrim(str(k))

152         m_sthl_ektyp(k + 1) = m_sthl_ektyp(k) + 7    ' Len(macro("p", aa)) + 1
        Next

154     Printer.FontSize = 7
156     Printer.FontName = "Courier New"
158     Printer.Font.Charset = 161
        'Printer.FontBold = True

        ' PRINT2_Xepik

        Dim R As New ADODB.Recordset

160     R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
        'R("pelono") = pel_ono.Text
        'R("pelepa") = pel_epa.Text
        'R("peldie") = pel_die.Text

162     If fENARJH_APO_SEIRA > 1 Then
            ' να τυπώνει ξεκινώντας από συγκεκριμένη γραμμή χωρίς να χάσει τις σούμες
            ' να μην τυπώνει την επικεφαλίδα
164         MsgBox "Προσοχή θα αρχίσει η εκτύπωση από την σειρά " + str(fENARJH_APO_SEIRA) + " και παρακάτω"

        Else
            ' νορμαλ  τυπώνει την επικεφαλίδα
        
166         SELIDA = SELIDA + 1
        
168         Printer.Print F_(R("pelono"))
170         Printer.Print F_(R("pelDIE"))
172         Printer.Print F_(R("pelepa")) + Space(40) + " Σελίς " + str(SELIDA)
            'Printer.Print Now
174         Printer.FontSize = 12
176         Printer.FontBold = True
178         Printer.Print Tab(9 / 12 * (LSYN / 2 - Len(EPIKEF) / 2)); F_(EPIKEF)
180         Printer.FontSize = 9

            '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------
182         For k = 0 To N - 1
               If F_T(k).Type = 5 Then
                  Printer.Print Tab(f_tab(k)); F_(Right("            " + F_T(k).Name, 12)); ' F_T(k).DefinedSize));
               Else
184               Printer.Print Tab(f_tab(k)); F_(Left(F_T(k).Name, F_T(k).DefinedSize));
               End If
            Next

            'δευτερη σειρα επικεφαλίδας
186         For k = 0 To N - 1
188             Printer.Print Tab(f_tab(k)); F_(mID$(F_T(k).Name, F_T(k).DefinedSize + 1, Len(F_T(k).Name) - F_T(k).DefinedSize + 1));
            Next

190         Printer.FontBold = False
192         Printer.Print Tab(m_sthl_ektyp(1)); String$(LSYN - 2, "-")
        End If
        
        
        
        

194     For i = 1 To N
196         m_sumes(i) = 0
198         SUMES0(i) = 0
        Next





        Dim LAST_TIMH

        Dim synt1

        Dim maxWidth As Long

200     synt1 = IIf(IsNull("SYNT1"), "true", synt1)   ' όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
202     AYJ = 0
204     Typose = 0

        Dim row

        Dim synola_SELIDOS

206     synola_SELIDOS = True    ' False

        '--------------------------------------------------------
208     Do While Not F_T.EOF
210         AYJ = AYJ + 1

212         If fENARJH_APO_SEIRA > 1 Then

                ' να τυπώνει ξεκινώντας από συγκεκριμένη γραμμή χωρίς να χάσει τις σούμες
                ' χρειάζεται για τις απογραφές
214             Do While AYJ < fENARJH_APO_SEIRA

216                 For k = 0 To N - 1

218                     If mID$(SUgm_str, k + 1, 1) = "1" Or mID$(SUgm_str, k + 1, 1) = "2" Then
220                         If IsNull(F_T(k)) Then
                            Else
222                             m_sumes(k) = m_sumes(k) + F_T(k)
                            End If
                        End If    '

                    Next

224                 F_T.MoveNext
226                 AYJ = AYJ + 1

228                 If F_T.EOF Then
230                     MsgBox "υπέρβαση τέλους αρχείου"

                        Exit Function

                    End If

                Loop

232             fENARJH_APO_SEIRA = 0    ' για να μην μπαινει ξανά στο ιδιο  If fENARJH_APO_SEIRA > 1 Then
            End If

234         If end_page = 2 Then    ' ΤΥΠΩΝΕΙ ΓΡΑΜΜΟΥΛΕΣ --------------------------------
236             Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN, "-")

238             end_page = 0

240             MOLIS_ALLAJE = 1
            End If    ' end_page=2

242         If synola_SELIDOS Then
244             If (AYJ - 1) Mod SEIRES_selidas = 0 And AYJ > 1 Then    ' .AND. PROTH_FORA<>1 )
                    'syn_epik
246                 Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN, "-")
248                 metaf = "Σε μεταφορά "
250                 GoSub printSUM

252                 Printer.NewPage

                    'ξαναδίνω το orientation γιατί το χάνει μετά το endDoc
254                 If LSYN > MEGISTO_PLATOS Then
256                     Printer.Orientation = vbPRORLandscape    '  2   Documents are printed with the top at the wide side of the paper.
                    Else
258                     Printer.Orientation = vbPRORPortrait  ' 1   Documents are printed with the top at the narrow side of the paper.
                    End If

260                 end_page = 0

                    '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------

                    '890                 Printer.Print
                    '900                 Printer.Print
                    '910                 Printer.Print
         
262                 SELIDA = SELIDA + 1
264                 Printer.Print F_(R("pelono"))
266                 Printer.Print F_(R("pelDIE"))
268                 Printer.Print F_(R("pelepa")) + Space(40) + " Σελίς " + str(SELIDA)
                    'Printer.Print Now
270                 Printer.FontSize = 12
272                 Printer.FontBold = True
274                 Printer.Print Tab(9 / 12 * (LSYN / 2 - Len(EPIKEF) / 2)); F_(EPIKEF)
276                 Printer.FontSize = 9


           '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------
            For k = 0 To N - 1
               If F_T(k).Type = 5 Then
                  Printer.Print Tab(f_tab(k)); F_(Right("            " + F_T(k).Name, 12)); ' F_T(k).DefinedSize));
               Else
                  Printer.Print Tab(f_tab(k)); F_(Left(F_T(k).Name, F_T(k).DefinedSize));
               End If
            Next



278      '           For k = 0 To N - 1
280       '              Printer.Print Tab(f_tab(k)); F_(Left(F_T(k).Name, F_T(k).DefinedSize));
          '          Next

                    'δευτερη σειρα επικεφαλίδας
282                 For k = 0 To N - 1
284                     Printer.Print Tab(f_tab(k)); F_(mID$(F_T(k).Name, F_T(k).DefinedSize + 1, Len(F_T(k).Name) - F_T(k).DefinedSize + 1));
                    Next

286                 Printer.Print Tab(m_sthl_ektyp(1)); String$(LSYN - 2, "-")

288                 metaf = "Εκ μεταφοράς "
290                 GoSub printSUM

                End If  'end_page = 1 Then
            End If

            '----------------- ΤΥΠΩΣΕ ΟΛΑ ΤΑ ΠΕΔΙΑ ---------------------
292         For k = 0 To N - 1

294            If k < N - 1 Then
                    ' για να μην παταει στην επόμενη στήλη
296                maxWidth = f_tab(k + 1) - f_tab(k) - 1
                End If

298            If Left(F_T(k), 3) = "@@@" Then    '  A/A
300                Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(AYJ, "######"), Min(F_T(k).DefinedSize, maxWidth));
302            ElseIf F_T(k).Type = 7 Or F_T(k).Type = 5 Then   'IsNumeric(f_t(K)) πραγματικο
304                Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(F_T(k), "######,##0.00"), Min(maxWidth, 12));    '  F_T(k).DefinedSize);
306            ElseIf F_T(k).Type = 8 Then   'DATE
308                Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(F_T(k), "DD/MM/YYYY"), Min(maxWidth, 10));
                Else    ' 10 STRING

310                If IsNull(F_T(k)) Then
312                    Printer.Print Tab(f_tab(k));
                    Else

                        ' PRINTER.PRINT Tab(f_tab(K)); to928(F_T(K));
314                    If k < N - 1 Then
                            ' για να μην παταει στην επόμενη στήλη
316                        Printer.Print Tab(f_tab(k)); F_(Left((F_T(k)), maxWidth));
                            'Printer.Print Tab(f_tab(k)); F_(Left(to928(F_T(k)), maxWidth));
                        Else
318                        Printer.Print Tab(f_tab(k)); F_((F_T(k)));
                            'Printer.Print Tab(f_tab(k)); F_(to928(F_T(k)));
                        End If
                    End If

320                If m_sthl_ektyp(k) > 2 Then
322                    If k = 1 Then Printer.Print
                    End If  'm_sthl_ektyp(K) > 2

                End If

                ' soymes---------------------
324            If mID$(SUgm_str, k + 1, 1) = "1" Or mID$(SUgm_str, k + 1, 1) = "2" Then
326                If IsNull(F_T(k)) Then

                    Else
328                    m_sumes(k) = m_sumes(k) + F_T(k)
                    End If
                End If    ' mid$(sugm_str,k,1)

            Next

330        Printer.Print

332        If GROUPN > 0 Then
334            LAST_TIMH = F_T(GROUPN - 1)
            End If

336        F_T.MoveNext

338        If Not F_T.EOF Then
340            If GROUPN > 0 Then
342                If LAST_TIMH <> F_T(GROUPN - 1) Then
344                    Printer.Print
                    End If
                End If
            End If

        Loop

        '"
346    Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN - 2, "-")

        '"
        ' PRINTER.PRINT Chr(13)  ' 6/12/2007
        '"
        '   aa = f_kodik("sum_seltxt")
        Dim PARAM    '

348    PARAM = IIf(AA = " ", "  ", AA)
        'pr_SUMselidas param
        'PRINT2_Xsumes "ΓΕΝΙΚΟ ΣΥΝΟΛΟ"

350    metaf = "ΣΥΝΟΛΑ "

352    GoSub printSUM
354    Printer.Print
356    Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN - 2, "-")

        '   PRINTER.PRINT Chr(12)
        '"
        Dim YPOSEL(10)

358    For k = 1 To 4
360        Printer.Print    '  1, 0; YPOSEL(k)
        Next

        'PRINTER.PRINT Chr(13)

362    If eject = "y" Then
            'PRINTER.PRINT Chr(12)
        End If

        ' αποθηκεύω τις σούμες στα e1,e2,.. για να τις χρησιμοποιώ
        '  For K = 1 To N
        '   aa = "e" + LTrim(Str(K))
        '    macro(aa, 0) = m_sumes(K)
        ' Next

        'Close #1

364    pp = 1
        '   If pp = 1 Then
        '     If ektypoths = -1 Then
        '     Else
        '
        '
        '       PPD = MsgBox("ΤΟ ΒΛΕΠΩ ΠΡΙΝ ΤΥΠΩΘΕΙ", vbYesNo)
        '       If PPD = vbYes Then
        '
        '          On Error GoTo print4_xar_Err
        '           DUM = Shell("c:\mercvb\notepad.exe c:\print", vbMaximizedFocus)
        '
        '           PPF = MsgBox("Προχωρώ στην εκτύπωση", vbYesNo)
        '           If PPF = vbYes Then
        '               DUM = Shell("c:\mercvb\notepad.exe /p c:\print", vbMaximizedFocus)
        '           End If
        '       Else
        '       End If
        '     End If
        '   End If
366    SELIS = 1
368    Printer.EndDoc

        Exit Function

printSUM:

370    Printer.Print Tab(0); F_(Trim(metaf));

372    For k = 0 To N - 1

374        If k < N - 1 Then
                ' για να μην παταει στην επόμενη στήλη
376            maxWidth = f_tab(k + 1) - f_tab(k) - 1
            End If

378        If m_sumes(k) > 0 Then
380            If mID$(SUgm_str, k + 1, 1) = "1" Then  'ΑΘΡΟΙΣΜΑ
382                Printer.Print Tab(f_tab(k) - 2); Right(Space(30) + Format(m_sumes(k), "######,##0.00"), Min(maxWidth, 14)); ' F_T(K).DefinedSize));
                Else  ' ΜΕΣΟΣ ΟΡΟΣ
384                Printer.Print Tab(f_tab(k) - 2); Right(Space(30) + Format(m_sumes(k) / AYJ, "######,##0.00"), Min(maxWidth, F_T(k).DefinedSize));
                End If

            Else
                '   Printer.Print Tab(f_tab(K)); ""; ' Right(Space(50), F_T(K).DefinedSize);
            End If

        Next

386    Printer.Print
388    Return

printEpik:

390    For k = 0 To N - 1

392        If k < N - 1 Then
                ' για να μην παταει στην επόμενη στήλη
394            maxWidth = f_tab(k + 1) - f_tab(k) - 1
            End If

396        If m_sumes(k + 1) > 0 Then
398            Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k + 1), "######,##0.00"), Min(maxWidth, F_T(k).DefinedSize));
            Else
400            Printer.Print Tab(f_tab(k)); Right(Space(50), Min(maxWidth, F_T(k).DefinedSize));
            End If

        Next

402    Printer.Print
404    Return


        '<EhFooter>
        Exit Function

print4_xar_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.print4_xar " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function























Function F_(ByVal SS)
        '<EhHeader>
        On Error GoTo F__Err
        '</EhHeader>

100     If F_437 = 1 Then
102         F_ = to437(SS)
        Else
104         F_ = SS
        End If


        '<EhFooter>
        Exit Function

F__Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.F_ " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function

Function macro(string_, AA)
        '<EhHeader>
        On Error GoTo macro_Err
        '</EhHeader>

        Dim fcheckonly As Boolean

100     gvar = ""    ' gia na midenizo tin makroentoli
        ' macro = EbExecuteLine(StrPtr(string_), 0&, 0&, Abs(fcheckonly)) = 0

        '  παράδειγμα χρησιμοποίησης μακροεντολής
        '  --------------------------------------

        '   Private Sub Command1_Click()
        '   Dim res As Boolean, var As Single
        ' αν χρησιμοποιούντα recordset πρεπει να είναι GLOBAL όπως GEID,GPEL,κ.λ.π.
        '   res = ExecuteLine("var=2+3*(5+6):a$=var")
        '
        '   End Sub


        '<EhFooter>
        Exit Function

macro_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.macro " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function

'
'proc CHECK_FIELDS
'*********************
'priv aa, BB
'              aa = StrZero(K, 2)
'              BB=trim(kodik("f&aa)
'
'            IF TYPE(BB)="U"
'                 wait 'προβλημα στο πεδίο '+  aa
'            Else
'              @ 24,22 say &BB
'              IF TYPE(BB)<>"N" .AND. synola(k)='1'
'                 synola(k)='0'
'              End If
'            End If
'Return

'
'#include "\clipper\include\error.ch"
'#include "\clipper\include\Fileio.ch"
'#include "\clipper\include\inkey.ch"
'
'Function ektyp_part(m__rec, mFile)
'
'
''*****************************************************************************
''LOCAL bOldHandler,mRow,mm_Head1,mm_head2,mm_head3,mm_head4
''PARA m__rec,mFile    && 0 ta dino ola edo
''
''PRIV i,Val_Field(150),arr1(50),a,zp,sfalma:=0,apo,eos,sapo,sevs,sure,mTypos(25),mString(25)
''PRIV EPIKEF_OU,SELIS:=0,synola(30),mBSEIRA := space(25),mdiax:=0
''priv YPOSEL(10),eject:='y',f(30),dok
''priv m_ord(10),_a,files:=1,b,arr2(30),head(30),mP(30)
''priv m1,m2,m3,m0,m4,_mTot(6),_kt,_St
''
''
''
''  For K = 1 To 4
''     yposel(k)=''
''  Next
''
''
''  For K = 1 To 30
''     f (K) = 0
''  Next
''
''  set colo to  &wn,&nw,,,gr+/b  &&  w/bg+
''  Clear
''   _a=1
''
''do while .t.
''   Clear
''
''   sure='Ο'
''   store date() to apo,evs
''   xvrio = Space(20)
''    kau = Space(20)
''   mseir = Space(20)
''   ms4 = Space(20)
''   ms5 = Space(20)
''
''   select 40
''     NET_USE('FONTS',.F.,0)
''     K = 0
''     Do While .not.EOF()
''        mTypos(++k)=FONTS("Typos
''        mString(k)  =FONTS("String
''        Skip
''     enddo
''
''   SELECT 40
''
''   if pcount()<2
''      mFile='kodik'
''   End If
''
''     if .not. file(mFile+'.DBF')
''         do CHECK_KODIK with mFile
''     End If
''
''      NET_USE(mfile ,.F.,0,'kodik')
''      IF LASTREC()=0
''         For K = 1 To 20
''            ADD_REC (0)
''         Next
''      End If
''
''      IF FCOUNT()<132
''         do CHECK_KODIK with mFile
''         NET_USE(mFile , .F. , 0 ,'kodik' )
''      End If
''
''   mm_Head1=kodik("m_Head1
''   mm_Head2=kodik("m_Head2
''   mm_Head3=kodik("m_Head3
''   mm_Head4=kodik("m_Head4
''
''   index on epikef2 to dokkod&xeirisths
''
''   go top
''   n = 0
''   Do While .not.EOF()
''      K = recn()
''      n = n + 1
''      arr1(n)= left(kodik("epikef2,55)+'.. '+trans(K,'99')
''      Skip
''   enddo
''
''
''
''  if m__rec = 0
''      _a=achoice ( 3,7,23,68,arr1,,,_a)
''  Else
''      _a=m__rec
''  End If
''
''
''  Clear
''
''   EPIKEF_OU = if( _a=0,'',arr1(_a) )
''   @ 0,20 say EPIKEF_OU
''
''
''
''
''
''
''   @ 4,0 say 'Nα διορθώσω το report' get sure
''   Read
''   if lastkey()=27
''       Close Data
''       Return
''   End If
''
''
''********* ----------------------------------------------------------------
''
''
''
''
''   go Val ( Right(arr1 ( if(_a<=0,1,_a) )  , 3 ) )
''
''
''   if sure $ 'nNνΝ'
''      if arx1='   '  && αν είναι νέο record να ρωτά από πού θα αντιγραφεί
''           mReck = recn()
''           mR = 1
''           @ 6,0 say 'Να αντιγραφεί από το Report No ' get mR
''           Read
''        * διαβάζω το παλιό record
''        go mR
''        For i = 1 To fcount()
''            f = Field(i)
''            Val_Field(i) = &f
''        Next
''
''        go mReck
''        rec_lock (0)
''        * αντιγράφω στο καινούριο record
''        For i = 1 To fcount()
''            f = Field(i)
''            repl &f with Val_Field(i)
''        Next
''        unlock
''
''      End If
''      rec_lock (0)
''      Clear
''      @ 0,20 say EPIKEF_OU
''      set message to 24
''      * @ 2,0 say 'αριθμός αρχείων που θα ανοίξουν ' get Number_arx
''      @ 3,0 say 'Αρχείο κύριο(με / κατάλογο)' get arx1 pict '@S10'   valid find_arxeio(arx1)=0
''
''      @ 4,0 say 'Αρχείο 2ο  ' get arx2
''      @ 5,0 say 'Σχέση που τα συνδέει με το κύριο  ' get relation2
''
''      @ 7,0 say 'Αρχείο 3ο  ' get arx3
''      @ 8,0 say 'Σχέση που τα συνδέει με το κύριο  ' get relation3
''
''      @10,0 say 'Αρχείο 4ο  ' get arx4
''      @11,0 say 'Σχέση που τα συνδέει με το κύριο  ' get relation4
''
''
''      @15,0 say 'Συνθήκη άθροισης on..  ' get total pict '@S30'
''      Read
''
''
''      if total<>'   '
''         For K = 1 To 6
''             _mTot(k) = subs( pedia_tot , 01+(k-1)*10 , 10 )
''         Next
''
''         @ 15,0 say 'Πεδία άθροισης '
''         For K = 1 To 6
''             @ 16,(k-1)*10+k-1  get _mTot(k)
''         Next
''
''         @ 17,0 say 'συνθήκη άθροισης for..' get for_tot
''         Read
''
''         m_tot2=''
''         For K = 1 To 6
''             m_tot2 = m_tot2 + _mTot(k)
''         Next
''         repl pedia_tot with m_tot2
''
''      End If
''
''      repl Number_arx with if(arx2=' ',1,if(arx3=' ',2,if(arx4=' ',3,4)))
''      unlock
''
''
''
''
''
''
''   End If
''
''
''      mNumber_arx=kodik("Number_arx
''      mARX1=trim(kodik("arx1)
''      mARX2=trim(kodik("arx2)
''      mARX3=trim(kodik("arx3)
''      mARX4=trim(kodik("arx4)
''      mRelation4 = Trim(relation4)
''      mRelation2 = Trim(relation2)
''      mRelation3 = Trim(relation3)
''      mBSEIRA = KODIK("bSEIRA
''
''
''
''      do OPEN_DATA WITH .F.,1,mARX1
''      For K = 1 To 4
''          m_ord (K) = indexkey(K)
''      Next
''
''
''      *        net_use(mArx1,.f.,0)
''   if mNumber_arx-1 > 0
''      do OPEN_DATA WITH .F.,mNumber_arx-1,mARX2,mARX3,mARX4
''   End If
''
''    sele &mARX1
''    make_relations()
''
''   if sure $ 'nNνΝ'
''        Clear
''        @ 0,20 say EPIKEF_OU
''
''
''     *do while .t.
''        sele kodik
''        rec_lock (0)
''
''        if kodik("m_head1=space(20) .and. kodik("m_head2=space(20)
''           repl m_Head1 with mm_Head1
''           repl m_Head2 with mm_Head2
''           repl m_Head3 with mm_Head3
''           repl m_Head4 with mm_Head4
''        End If
''
''
''
''        @ 0,0 say 'Τίτλος Εκτύπωσης(οθόνη-μενού)' get kodik("epikef2 PICT '@S30'
''
''
''
''
''
''        mRow = 2
''        @ mRow+0,0 say 'α Τίτλος  (επάνω-αριστ)'  get kodik("m_Head1  PICT '@S30'
''        @ mRow+1,0 say 'β Τίτλος  (επάνω-αριστ)'  get kodik("m_Head2  PICT '@S30'
''        @ mRow+2,0 say 'γ Τίτλος  (επάνω-αριστ)'  get kodik("m_Head3  PICT '@S30'
''        @ mRow+3,0 say 'δ Τίτλος  (επάνω-αριστ)'  get kodik("m_Head4  PICT '@S30'
''
''
''
''
''        mRow = 6
''        @ mRow+0,0 say 'α Επικεφαλίδα Εκτύπωσης(κέντρο)'  get kodik("epikef  PICT '@S30'
''        @ mRow+1,0 say 'β Επικεφαλίδα Εκτύπωσης(κέντρο)'  get kodik("epik_2   PICT '@S30'
''        @ mRow+2,0 say 'γ Επικεφαλίδα Εκτύπωσης(κέντρο)'  get kodik("epik_3  PICT '@S30'
''
''
''
''        @ mRow+4,0 say 'Aριθμός πεδίων' get kodik("nr range 1,25
''
''        For K = 1 To 4
''          @ 15+k,0 say str(k)+'='+m_ord(k)
''        Next
''
''        @ mRow+5,35 say 'ταξινόμηση κατά '    get kodik("tag PICT '@S20'
''
''        @ mRow+6,0 say 'Συνθήκη << όσο συμβαινει τύπωνε >>  ' get kodik("syntwh
''        @ mRow+7,0 say 'Φίλτρο ' get kodik("syntIF PICT '@S70'
''        Read
''
''
''
''
''
''
''        @ 15,0 clear to 24,79
''
''
''        a=1  && βρισκω την προηγούμενη ρύθμιση των fonts
''        For K = 1 To Len(mString)
''            if mString(k)=trim ( KODIK("FONTS )
''               a = K
''               exit
''            End If
''        Next
''
''        @ 8,0 say 'Δώσε τον τύπο της εκτύπωσης'
''        * a=achoice ( 8,30,12,58,mTypos,,,a )
''        repl kodik("fonts with if( a=0,'',mString(a))
''
''
''        if type(kodik("syntwh)='UI' .or. type(kodik("syntwh)='UE'
''           wait 'προβλημα στην συνθήκη << όσο συμβαινει τύπωνε >> '
''           BREAK
''        End If
''        if type(kodik("syntIF)='UI' .or. type(kodik("syntIF)='UE'
''           wait 'προβλημα στο φίλτρο '
''           BREAK
''        End If
''
''        Clear
''        @ 0,20 say EPIKEF_OU
''
''        @ 1, 0  say 'Επικ.Πεδίου'
''        @ 1,22 say 'Τίτλος Πεδίου'
''        @ 1,44 say 'format γραφής'
''        @ 1,70 say 'Σύνολο=1'
''
''      SELECT 41
''         NET_USE("FIELDS",.F.,0)  && ΟΝΟΜΑΤΑ ΠΕΔΙΩΝ
''            set filter to trim(arxeio) $ mARX1+mARX2+mARX3+mARX4
''
''
''        set key 14 to insert_line  && Ctrl + N
''        set key 25 to delete_line  && Ctrl + Y
''        for k=1 to kodik("nr
''            synola(k)=subs(kodik("synolo,k,1)
''        Next
''        set key -9 to opos_allo
''        for k=1 to min(22,kodik("nr)
''            aa = StrZero(K, 2)
''            @ 1+k, 0 say k pict '99'
''            @ 1+k, 3 get kodik("ep&aa pict '@S20' valid find_field(kodik("ep&aa,MARX1,MARX2,MARX3)=0
''            @ 1+k,25 get kodik("f&aa  pict '@S20'
''            @ 1+k,47 get kodik("p&aa  pict '@S20'
''            @ 1+k,70 get synola(k)    pict '9'
''        Next
''        Read
''        set key 14 to  && Ctrl + N
''        set key 25 to  && Ctrl + Y
''
''
''        Clear
''        if kodik("nr > 22
''          for k=23 to kodik("nr
''            aa = StrZero(K, 2)
''            @ k-21, 0 say k pict '99'
''            @ k-21, 3 get kodik("ep&aa pict '@S20' valid find_field(kodik("ep&aa)=0
''            @ k-21,25 get kodik("f&aa  pict '@S20'
''            @ k-21,47 get kodik("p&aa  pict '@S20'
''            @ k-21,70 get synola(k)    pict '9'
''          Next
''          Read
''        End If
''
''        set key -9 to
''
''        sele &mARX1
''        boldhandler:=errorblock({|e| MyERRORHANDLER(e,boldHandler)})
''        begin sequence
''          * σε περίπτωση που χρησιμοποιώ μία άλλη στήλη π.χ. f6 να μην βγάζει λάθος
''          ssss = 1
''          for k=1 to kodik("nr
''             aa = alltrim(Str(K))
''             f&aa='ssss'
''          Next
''          for k=1 to kodik("nr
''            DO CHECK_FIELDS
''          Next
''        recover
''            aa = StrZero(K, 2)
''            IF TYPE("KODIK("F"+AA)="U"
''                 wait 'προβλημα στο πεδίο '+  aa
''            Else
''                 aa=trim(kodik("f&aa)
''                 wait 'προβλημα στο πεδίο '+ STR(K) + '  '+ aa
''            End If
''          ERRORBLOCK (boldhandler)
''        end sequence
''        ERRORBLOCK (boldhandler)
''
''        @ 22,0 say '2η σειρά:' get kodik("bSEIRA PICT '@S40'
''        @ 23,0 say 'Aλλάζει σελίδα και όταν αλλάζει το string πεδίο Νο π.χ. 2  0=δεν αλλάζει' get kodik("s_eject
''        @ 24,0 say 'Θά αλλάζει σελίδα=0 διαχωρισμός με παύλες=1 ' get kodik("diax
''        Read
''        mBSEIRA = KODIK("bSEIRA
''        Clear
''
''        if kodik("s_eject > 0
''            IF kodik("synt_alag='  '
''               REPL  kodik("synt_alag WITH 'mpal=mpal2'
''            End If
''            @ 2,0 say 'Συνθήκη για αλλαγή σελίδας (mpal=mpal2)' get kodik("synt_alag pict '@S20'
''            @ 4,0 say 'Κείμενο που θά γράφεται στά αθροίσματα' get kodik("sum_seltxt
''        End If
''        @ 6, 0 say 'Νά τυπώνει αθροίσματα σε μεταφορά=1   οχι=0 ' get kodik("se_metaf
''        @ 8, 0 say 'Νά τυπώνει αθροίσματα εκ μεταφοράς=1  όχι=0 ' get kodik("ek_metaf
''        Read
''
''        set key 14 to
''        set key 25 to
''        mSYNOLO=''
''        for k=1 to kodik("nr
''            mSYNOLO=mSYNOLO+synola(k)
''        Next
''        REPL kodik("SYNOLO WITH MSYNOLO
''        Clear
''        @ 2,0 say 'Κείμενο στην 1η ημερομηνία(apo)  ' get  kodik("k_hm1
''        @ 3,0 say 'Κείμενο στην 2η ημερομηνία(evs)  ' get  kodik("k_hm2
''        @ 5,0 say 'Κείμενο στo  1o όνομα     (xvrio)' get  kodik("k_ch1
''        @ 6,0 say 'Κείμενο στo  2ο όνομα     (kau)  ' get  kodik("k_ch2
''        @ 7,0 say 'Κείμενο στo  3ο όνομα     (mseir)' get  kodik("k_ch3
''        @ 8,0 say 'Κείμενο στo  4ο όνομα     (ms4)  ' get  kodik("k_ch4
''        @ 9,0 say 'Κείμενο στo  5ο όνομα     (ms5)  ' get  kodik("k_ch5
''        Read
''*        unlock
''
''
''        *EXIT
''     *enddo
''
''
''   End If
''
''
''  Clear
''  if kodik("k_hm1<>'  '
''     @ 1, 0 say trim(kodik("k_hm1) get apo
''  End If
''  if kodik("k_hm2<>'  '
''     @ 2, 0 say trim(kodik("k_hm2) get evs
''  End If
''  if kodik("k_ch1<>'  '
''     @ 4, 0 say trim(kodik("k_ch1) get xvrio
''  End If
''
''  if kodik("k_ch2<>'  '
''     @ 5, 0 say trim(kodik("k_ch2) get kau
''  End If
''
''  if kodik("k_ch3<>'  '
''     @ 6, 0 say trim(kodik("k_ch3) get mseir
''  End If
''
''  if kodik("k_ch4<>'  '
''     @ 7, 0 say trim(kodik("k_ch4) get ms4
''  End If
''
''  if kodik("k_ch5<>'  '
''     @ 8, 0 say trim(kodik("k_ch5) get ms5
''  End If
''
''  Read
''
''     if lastkey()=27
''        Loop
''     End If
''
''     Sapo = dtoc(apo)
''     Sevs = dtoc(evs)
''
''     synt_eject=kodik("s_eject
''
''
''   sele &mARX1
''
''
''    Err = 0
''
''    if err = 1
''       Close Data
''       Loop
''    End If
''
''
''
''    ZP = 1
''    do epilogh2 with 23,zp,3,'εκτύπωση σε οθόνη','εκτύπωση σε εκτυπωτή','edit'
''    IF LASTKEY()=27
''       Close Data
''       Loop
''    End If
''
''       b=trim(kodik("tag)
''
''
''       sele &marx1
''       if mARX1='MIS'
''          do make_MIS_totals
''       Else
''          do make_totals
''       End If
''
''
''
''
''
''
''
''
''
''
''       go top
''
''    progr=if(zp=1,"printbox4","print3_xar")
''
''
''*/*
''    IF ZP=3
''        xxx=kodik("syntIF
''        set filter to &xxx
''        go top
''        arr2 (1) = "lines"
''        for k=1 to kodik("nr
''            aa = StrZero(K, 2)
''            arr2(k)=kodik("f&aa
''            head(k)=kodik("ep&aa
''            mP(k)  =kodik("p&aa
''        Next
''        dbedit ( 1,0,24,79,arr2,"db17_udf",,head )
''        Close Data
''        Return
''    End If
''*/
''
''
''    IF ZP=2
''       @ 24,0 CLEAR TO 24,79
''       @ 24,0 SAY 'ΑΠΟ ΣΕΛΙΔΑ ' GET SELIS
''       Read
''    End If
''
''Clear
''  sthles = 200
''  epik=trim(kodik("epikef)
''  if len(trim(kodik("epikef))<>0
''     epik=&epik
''  End If
''  print_fonts=kodik("fonts
''
''  epik2 = trim(kodik("epik_2)
''  if len(trim(epik2))<>0
''      epik2 = &epik2
''  End If
''
''  epik3 = trim(kodik("epik_3)
''  if len(trim(epik3))<>0
''      epik3 = &epik3
''  End If
''
''  do &PROGR  with KODIK("NR,TRIM(KODIK("SYNTWH),TRIM(KODIK("SYNTIF),TRIM(KODIK("SYNOLO),;
''       TRIM(kodik("EP01),TRIM(kodik("F01),TRIM(kodik("P01),;
''       TRIM(kodik("EP02),TRIM(kodik("F02),TRIM(kodik("P02),;
''       TRIM(kodik("EP03),TRIM(kodik("F03),TRIM(kodik("P03),;
''       TRIM(kodik("EP04),TRIM(kodik("F04),TRIM(kodik("P04),;
''       TRIM(kodik("EP05),TRIM(kodik("F05),TRIM(kodik("P05),;
''       TRIM(kodik("EP06),TRIM(kodik("F06),TRIM(kodik("P06),;
''       TRIM(kodik("EP07),TRIM(kodik("F07),TRIM(kodik("P07),;
''       TRIM(kodik("EP08),TRIM(kodik("F08),TRIM(kodik("P08),;
''       TRIM(kodik("EP09),TRIM(kodik("F09),TRIM(kodik("P09),;
''       TRIM(kodik("EP10),TRIM(kodik("F10),TRIM(kodik("P10),;
''       TRIM(kodik("EP11),TRIM(kodik("F11),TRIM(kodik("P11),;
''       TRIM(kodik("EP12),TRIM(kodik("F12),TRIM(kodik("P12),;
''       TRIM(kodik("EP13),TRIM(kodik("F13),TRIM(kodik("P13),;
''       TRIM(kodik("EP14),TRIM(kodik("F14),TRIM(kodik("P14),;
''       TRIM(kodik("EP15),TRIM(kodik("F15),TRIM(kodik("P15),;
''       TRIM(kodik("EP16),TRIM(kodik("F16),TRIM(kodik("P16),;
''       TRIM(kodik("EP17),TRIM(kodik("F17),TRIM(kodik("P17),;
''       TRIM(kodik("EP18),TRIM(kodik("F18),TRIM(kodik("P18),;
''       TRIM(kodik("EP19),TRIM(kodik("F19),TRIM(kodik("P19),;
''       TRIM(kodik("EP20),TRIM(kodik("F20),TRIM(kodik("P20),;
''       TRIM(kodik("EP21),TRIM(kodik("F21),TRIM(kodik("P21),;
''       TRIM(kodik("EP22),TRIM(kodik("F22),TRIM(kodik("P22),;
''       TRIM(kodik("EP23),TRIM(kodik("F23),TRIM(kodik("P23),;
''       TRIM(kodik("EP24),TRIM(kodik("F24),TRIM(kodik("P24),;
''       TRIM(kodik("EP25),TRIM(kodik("F25),TRIM(kodik("P25),;
''       TRIM(kodik("EP26),TRIM(kodik("F26),TRIM(kodik("P26),;
''       TRIM(kodik("EP27),TRIM(kodik("F27),TRIM(kodik("P27),;
''       TRIM(kodik("EP28),TRIM(kodik("F28),TRIM(kodik("P28),;
''       TRIM(kodik("EP29),TRIM(kodik("F29),TRIM(kodik("P29),;
''       TRIM(kodik("EP30),TRIM(kodik("F30),TRIM(kodik("P30)
''Close Data
''       SELIS = 1
''
''dok='dok'+xeirisths+'.mdx'
''erase &dok
''
''
''
''enddo
''
''
''Return
'End Function
'
'
'Function find_arxeio(dum)
'*********************************************
'priv arxeia(9),Marxeia(9),k,l,sc
'
'if dum<>'/'
'   return 0
'End If
'
'  sc = savescreen(4, 40, 7, 79)
'      arxeia(1)='ERGAZ'
'
'    /*
'       arxeia(2)='TIM'
'       arxeia(3)='GRA'
'       arxeia(4)='EID'
'       arxeia(5)='EGG'
'       arxeia(6)='EGGTIM'
'    */
'
'      Marxeia(1)='Αρχείο Εργαζόμενου'
'    /*
'      Marxeia(2)='Αρχείο Παραστατικών'
'      Marxeia(3)='Αρχείο Επιταγών / Γραμματίων'
'      Marxeia(4)='Αρχείο ειδών '
'      Marxeia(5)='Αρχείο κινήσεων πελατών/προμηθευτών '
'      Marxeia(6)='Αρχείο κινήσεων ειδών'
'     */
'
'For K = 1 To 6
'    @ k+3,40 prompt arxeia(k) message Marxeia(k)
'Next
'menu to l
'restscreen (4,40,7,79,sc)
' if l<>0
'    repl arx1 with arxeia(l)
' End If
'return 0
'
'
'
'
'
'proc PRINTBOX4
'*********************** P R I N T B O X ************************
'**************      κάνει παρουσίαση αρχείου
'*** n=αριθμός fields που παρουσιάζονται,synuhkh συνθήκη για το DO WHILE
'*** synuhkh2 η συνθήκη για το  IF, sugm_str βλέπει που θα κάνει σούμες ,π.χ. "00100" κάνει σούμες στο 3ο field
'*** sum_pic  το picture γιά τις σούμες
'*** Ei,Fi,Pi  :Επικεφαλίδα παρουσίασης,Fields που παρουσιάζονται,Picture παρουσίασης
'
'para n,synuhkh,synuhkh2,sugm_str, _e1,f1,p1, _e2,f2,p2, _e3,f3,p3, _e4,f4,p4, _e5,f5,p5,  _e6,f6,p6, _e7,f7,p7, _e8,f8,p8, _e9,f9,p9 ,_e10,f10,p10,_e11,f11,p11, _e12,f12,p12,_e13,f13,p13,_e14,f14,p14,_e15,f15,p15,_e16,f16,p16, _e17,f17,p17,_e18,f18,p18,_e19,f19,p19,_e20,f20,p20,;
'          _e21,f21,p21, _e22,f22,p22, _e23,f23,p23, _e24,f24,p24, _e25,f25,p25,_e26,f26,p26, _e27,f27,p27,_e28,f28,p28,_e29,f29,p29,_e30,f30,p30
'priv sumes, synuhkh, synuhkh2, sugm_str, cc, pp, AYJ, P
'
'
'declare S(30),sumes(30),OUON(5)
'For K = 1 To 30
'   sumes (K) = 0
'Next
'For K = 1 To 5
'   ouon(k)=''
'Next
'   LSYN = 1
'   epikef0 = "Ι"
'   epikef = "Ί"
'   epikef2 = "Μ"
'   patos = "Θ"
'   GEMIS = "Ί"
'For K = 1 To n
'   aa = alltrim(Str(K))
'   lsyn = lsyn + len(p&aa) + 1
'   epikef0 = epikef0 + repl("Ν",len(p&aa))+"Λ"
'   epikef  =epikef +left(_e&aa+space(50),len(p&aa))+"Ί"
'   epikef2 = epikef2 + repl("Ν",len(p&aa))+"Ξ"
'   PATOS = patos + repl("Ν",len(p&aa))+"Κ"
'   GEMIS = gemis + left(space(50),len(p&aa)) + "Ί"
'Next
'
' do printbox21
'Return
'
'
'
'static proc printbox21
'****************** συνέχεια απο παραπάνω ***********************************
'priv aa,pict(30),f(30)
'synt1 = if ( type("SYNT1")="U",'.t.',synt1)   && όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
'
'm_SEIRA = 1
'
'
'
'
'Clear
'
'arxh0 = 1 + Len(p1 + p2) + 2
'arxh = arxh0 + 1
'nouon = 0
'
'ARXSTHL = IF ( LSYN>77 , 0 , 40 - LSYN/2 )
'
'     @ 0,ARXSTHL SAY left(EPIKEF0,ARXH0) + SUBS(EPIKEF0,ARXH,79-arxh0)
'     @ 1,ARXSTHL SAY left(EPIKEF ,ARXH0) + SUBS(EPIKEF ,ARXH,79-arxh0)
'     @ 2,ARXSTHL SAY left(EPIKEF2,ARXH0) + SUBS(EPIKEF2,ARXH,79-arxh0)
'
'AYJ = 0
'do while &synuhkh
'     set cursor off
'
'
' IF &synuhkh2   &&  η δεύτερη συνθήκη / αν δεν υπάρχει τότε .Τ.
'
'   if &synt1   && η συνθήκη της αποθήκης αλλιώς .Τ.
'      AYJ = AYJ + 1
'
'
'          For K = 1 To n
'            aa = 'f'+ alltrim( str ( k ) )
'            aa = &aa
'            f(k)= &aa
'            aa = 'p'+ alltrim( str ( k ) )
'            aa = &aa
'            pict (K) = aa
'
'            ww = alltrim(Str(K))
'            aas = f&ww
'            sumes(k) = sumes(k) +  if ( subs(sugm_str,k,1)="1" , &aas  , 0 )
'          Next
'
'          u = "Ί"
'          For K = 1 To n
'             kena = len(pict(k))  - len( trans(f(k),pict(k))  )
'             kena = if ( kena=0 ,'' , space(kena) )
'             u = u + trans(f(k),pict(k))+kena+"Ί"
'          Next
'          s (m_SEIRA) = u
'
'      @ m_SEIRA+2,ARXSTHL SAY left(s(m_SEIRA),arxh0) + subs ( S(m_SEIRA) , arxh,79-arxh0 )
'      m_SEIRA = m_SEIRA + 1
'
'
'    endif  && synt1
' ENDIF  &&  η δεύτερη συνθήκη synuhkh2
'
'if type('TELOS') <> 'U'  &&   αν ορίστηκε το telos
'      if recn()=telos  && αν έφτασε στην τελευταία εγγραφή , να το κάνει eof()
'           seek 'ωωωωωωωωωωωωωωω'
'      Else
'           Skip
'      End If
'Else
'      Skip
'End If
'
'if ROW() >= 22
' @ ROW()+1,ARXSTHL SAY left(patos,arxh0)+SUBS(PATOS,ARXH,79-arxh0)
' set cursor on
' sumarisma = ''
' For K = 1 To n
'        ww = alltrim(Str(K))
'        aas = p&ww
'        sumarisma = sumarisma + " " +  if ( subs(sugm_str,k,1)="1" , trans(sumes(k),"&aas"),space( len(aas) )  )
' Next
' sumarisma = sumarisma + ' '
' @ 24,ARXSTHL say left(sumarisma,arxh0)+subs ( sumarisma , arxh , 79-arxh0 )
'
' nouon = nouon + 1
' if nouon > 5
'    For K = 1 To 4
'        ouon(k) = ouon(k+1)
'    Next
'    save screen to ouon(5)
'    nouon = 5
' Else
'    save screen to ouon(nouon)
' End If
'
' @ 24,ARXSTHL say "< Esc > επιστρέφω"
' DO Nwait
' plhk = lastkey()
' piso = 0
' do while (plhk=4 .or. plhk=19  .or. plhk=18 .or. plhk=3) .AND. lsyn>80 && LSYN-n>79 && (" ή <- ή PgDn ή PgUp
'   if plhk=18 && PgUp
'      piso = Min(piso + 1, nouon - 1)
'      restore screen from ouon(nouon-piso)
'   End If
'   if plhk=3 && PgDn
'      piso = Max(piso - 1, 0)
'      restore screen from ouon(nouon-piso)
'   End If
'   if plhk=4 && (" belos
'     arxh = Min(arxh + 2, LSYN - 78 + arxh0)
'     DO HOR_SCROLL
'   End If
'   if plhk=19 && <- belos
'     arxh = Max(arxh - 2, arxh0 + 1)
'     DO HOR_SCROLL
'   End If
'   do nwait
'   plhk = lastkey()
' enddo
'   Clear
'   m_SEIRA = 1
'   if lastkey()=27
'      Return
'   End If
'endif      &&  row() = 22
'
'     @ 0,ARXSTHL SAY left(EPIKEF0,ARXH0) + SUBS(EPIKEF0,ARXH,79-arxh0)
'     @ 1,ARXSTHL SAY left(EPIKEF ,ARXH0) + SUBS(EPIKEF ,ARXH,79-arxh0)
'     @ 2,ARXSTHL SAY left(EPIKEF2,ARXH0) + SUBS(EPIKEF2,ARXH,79-arxh0)
'
'
'
'enddo
'
' Do While Row() < 22
'     s (m_SEIRA) = GEMIS
'     @ m_SEIRA+2,ARXSTHL SAY left(s(m_SEIRA),arxh0) + subs ( S(m_SEIRA) , arxh,79-arxh0 )
'     m_SEIRA = m_SEIRA + 1
' enddo
'
' @ ROW()+1,ARXSTHL SAY left(patos,arxh0)+SUBS(PATOS,ARXH,79-arxh0)
' set cursor on
' sumarisma = ''
' For K = 1 To n
'        ww = alltrim(Str(K))
'        aas = p&ww
'        sumarisma = sumarisma + " " +  if ( subs(sugm_str,k,1)="1" , trans(sumes(k),"&aas"),space( len(aas) )  )
' Next
' sumarisma = sumarisma + ' '
' @ 24,ARXSTHL say left(sumarisma,arxh0)+subs ( sumarisma , arxh , 79-arxh0 )
'
' nouon = nouon + 1
' if nouon > 5
'    For K = 1 To 4
'        ouon(k) = ouon(k+1)
'    Next
'    save screen to ouon(5)
'    nouon = 5
' Else
'    save screen to ouon(nouon)
' End If
'
' @ 24,ARXSTHL say "< Esc > επιστρέφω"
' DO Nwait
' plhk = lastkey()
' piso = 0
' do while plhk=4 .or. plhk=19  .or. plhk=18 .or. plhk=3 && (" ή <- ή PgDn ή PgUp
'   if plhk=18 && PgUp
'      piso = Min(piso + 1, nouon - 1)
'      restore screen from ouon(nouon-piso)
'   End If
'   if plhk=3 && PgDn
'      piso = Max(piso - 1, 0)
'      restore screen from ouon(nouon-piso)
'   End If
'   if plhk=4 && (" belos
'     arxh = Min(arxh + 2, LSYN - 78 + arxh0)
'     DO HOR_SCROLL
'   End If
'   if plhk=19 && <- belos
'     arxh = Max(arxh - 2, arxh0 + 1)
'     DO HOR_SCROLL
'   End If
'   do nwait
'   plhk = lastkey()
' enddo
'   Clear
'   m_SEIRA = 1
'   if lastkey()=27
'      Return
'   End If
'Return
'
'
'
'
'
'
'
'
'static PROC HOR_SCROLL
'****************** ΟΡΙΖΟΝΤΙΟ SCROLLING *************************
'SET CURSOR OFF
'priv K
'     @ 0,ARXSTHL SAY left(EPIKEF0,ARXH0) + SUBS(EPIKEF0,ARXH,79-arxh0)
'     @ 1,ARXSTHL SAY left(EPIKEF ,ARXH0) + SUBS(EPIKEF ,ARXH,79-arxh0)
'     @ 2,ARXSTHL SAY left(EPIKEF2,ARXH0) + SUBS(EPIKEF2,ARXH,79-arxh0)
'     For K = 3 To 22
'        @ k,ARXSTHL SAY left(s(k-2),arxh0) + subs ( S(k-2) , arxh,79-arxh0 )
'     Next
'
'     @ 23,ARXSTHL SAY left(patos,arxh0) + SUBS(PATOS,ARXH,79-arxh0)
'     @ 24,ARXSTHL say left(sumarisma , arxh0) + subs ( sumarisma , arxh , 79-arxh0 )
'SET CURSOR ON
'Return
'
'
'static proc nwait
'**********  waiting untill key pressed *************
'a = inkey()
'Do While a = 0
'  *@ 23,78 SAY ""
'  a = inkey()
'enddo
'Return
'
'
'Sub syn_epik()
'
'
''********** εκτυπωση επικεφαλίδων & συνόλων *************************
''priv M
''     m=kodik("sum_seltxt
''     printer.print+2,1 say ''
''
''       do pr_SUMselidas with if(len(trim(m))=0,' ',&m)
''     if kodik("se_metaf=1
''        do print2_xSUMES  with 'ΣΥΝΟΛΑ ΣΕ ΜΕΤΑΦΟΡΑ'
''     End If
''   ** να μηδενίζει τους αθροιστές μόνο όταν αλλάζει το εν λόγω πεδίο
''   if synt_eject > 0
''      if end_Page = 1
''         for i=1 to n  && αποθηκεύω τα σύνολα για να μπορώ να βγάλω μετά τα συνολα σελίδας
''           sumes0(i) = sumes(i)
''         Next
''      End If
''   Else
''        for i=1 to n  && αποθηκεύω τα σύνολα για να μπορώ να βγάλω μετά τα συνολα σελίδας
''           sumes0(i) = sumes(i)
''        Next
''   End If
''
''
''
''   For K = 1 To 4
''     printer.print+1,0 say YPOSEL(k)
''   Next
''
''   eject
''   For kw = 1 To 4
''       if len ( ar_Print(kw) ) > 0
''          printer.print+1 , 0 say ar_Print(kw)
''       End If
''   Next
''
'' IF ! type("EPIK")="U"
''   printer.print+1,IF(type("STHLES")="U",40,LSYN/2) - LEN(EPIK)/2  SAY EPIK
''   SELIS = SELIS + 1
''   printer.print,lsyn say SELIS
''
''   if epik2='  '
''      printer.print+1,IF(type("STHLES")="U",40,LSYN/2) - LEN(EPIK)/2  SAY  repl('-',len(epik) )
''   Else
''      printer.print+1,IF(type("STHLES")="U",40,LSYN/2) - LEN(TRIM(EPIK2))/2  SAY EPIK2
''      if epik3 <> '  '
''         printer.print+1,IF(type("STHLES")="U",40,LSYN/2) - LEN(TRIM(EPIK3))/2  SAY  epik3
''      End If
''   End If
'' End If
''
''   do print2_xEPIK
''   printer.print+1,_sthl_ektyp(1)+1 SAY repl('-',lsyn-2)
''   * printer.print+1,0 SAY ''
''   ?? CHR(13)
''   if kodik("ek_metaf=1
''      do print2_xSUMES  with 'ΣΥΝΟΛΑ ΑΠΟ ΜΕΤΑΦΟΡΑ'
''   End If
'End Sub
'
'
'
'
'
'
'Sub pr_SUMselidas(TITLOS)
'
''   ******************************** τυπώνει τις σούμες SELIDOS για το print_xar ********
''   para TITLOS
''   printer.print+1,1+_sthl_ektyp(1) SAY TITLOS
''      For K = 1 To n
''       aa = alltrim(Str(K))
''           aaP='P'+aa
''           aaP2=&aaP
''       if subs(sugm_str,k,1)="1"              && τυπώνει τις σούμες
''            printer.print,_sthl_ektyp(k)+1 say sumes(k)-SUMES0(k) pict "&aap2"
''       End If
''      Next
'
'End Sub

'Sub PRINT2_Xsumes(txt)
'
'
''   ******************************** τυπώνει τις σούμες για το print_xar ********
''*  para txt
''   printer.print+1,1+_sthl_ektyp(1) SAY ''   &&  txt
''   For K = 1 To n
''       aa = alltrim(Str(K))
''           aaP='P'+aa
''           aaP2=&aaP
''
''
''       if subs(sugm_str,k,1)="1"              && τυπώνει τις σούμες
''           if _sthl_ektyp(k) > 2
''              printer.print+if(k=1,1,0) , _sthl_ektyp(k) say "³" && διαχωριστικο
''           End If
''
''
''          @  PROW(),_sthl_ektyp(k)+1 say sumes(k) pict "&aap2"
''       End If
''   Next
'End Sub
'
'
'static proc OLDPRINT2_Xepik
'   ***************** τυπώνει επικεφαλίδες απο το print_xar
'  Local xor
'   printer.print+1,0 say ''
'   For K = 1 To n
'       aa = alltrim(Str(K))
'       xor=at(";",_E&aa)-1 && ψηφία από πρώτο κομμάτι που θα τυπωθεί
'       xor = if ( xor=0 , len(trim(_E&aa)) , xor ) && αν δεν έχει χώρισμα το παίρνει όλο
'       xor = if ( xor > len( P&aa ) , len( P&aa ) , xor )
'       @  prow(),_sthl_ektyp(k)+1 say  subs(_E&aa,1,xor)
'   Next
'   printer.print+1,0 say ''
'   For K = 1 To n
'       aa = alltrim(Str(K))
'       xor = at (";",_E&aa) && se ποιό σημείο υπάρχει το ;
'       mhk=len(trim(_E&aa))-xor && μήκος εκτύπωσης
'       mhk= if ( mhk > len( P&aa ) , len( P&aa ) , mhk )
'       @  prow(),_sthl_ektyp(k)+1 say  subs(_E&aa,xor+1,mhk)
'   Next
'Return
'
'Sub PRINT2_Xepik()
'
'
''***************** τυπώνει επικεφαλίδες απο το print_xar   3 SEIRES EPIKEFALIDA
'' Local xor,xor2, _EPI(40,3),k,seires_ep,l,arxh
'' priv aa
''
''* ΧΩΡΙΖΩ ΤΗΝ ΕΠΙΚΕΦΑΛΙΔΑ ΣΕ 3 ΜΕΡΗ
''   seires_ep = 1
''   For K = 1 To n
''       aa = alltrim(Str(K))
''       xor=at(";",_E&aa) && το 1ο ;
''       if xor=0  && μια σειρά μόνο
''          _epi(k,1)=""
''          _epi(k,2)=""
''          _epi(k,3)=_E&aa
''       Else
''          xor2 = Rat(";",_E&aa) && το 2ο ;
''
''          if xor2 > xor   && εχω 2 ερωτηματικά ( 3 σειρές)
''              _epi(k,1)=subs( _E&aa , 1 , xor-1 )
''              _epi(k,2)=subs( _E&aa , xor+1  , xor2- xor -1 )
''              _epi(k,3)=subs( _E&aa , xor2+1 ,len(_E&aa) - xor2 -1 )
''              seires_ep = 3
''          else   && ολο και όλο 1  ; ( 2 σειρές )    xor2=xor1
''              _epi(k,1)=""
''              _epi(k,2)=subs( _E&aa , 1  , xor -1 )
''              _epi(k,3)=subs( _E&aa , xor+1 ,len(_E&aa) - xor -1 )
''              seires_ep = IIf(seires_ep = 1, 2, seires_ep)
''
''          End If
''       End If
''   Next
''
''
''
''
''arxh  = if(seires_ep=2,2,if(seires_ep=1,3,1))
''
''For l = arxh To 3
''
''    printer.print+1,0 say ''
''   For K = 1 To n
''
''       aa = alltrim(Str(K))
''
''           if _sthl_ektyp(k) > 2
''              printer.print+if(k=1,1,0) , _sthl_ektyp(k) say "³" && διαχωριστικο
''           End If
''
''           if len( _epi(k,l) )  > len( P&aa )
''              @  prow(),_sthl_ektyp(k)+1 say  Left ( _epi(k,l) , len( P&aa ) )
''           Else
''              @  prow(),_sthl_ektyp(k)+1 say  _epi(k,l)
''           End If
''
''   Next
''Next
''
'End Sub
'
'
'
'Function Find_Field(dum, MARX1, MARX2, MARX3)
' ****************************************************
' LOCAL ppw,scr2
' declare arr1(50)
'
'   if left( dum , 1 ) $ '*/.'
'
'       save screen to scr2
'
'
'
'
'     IF upper(MARX1) $ 'ERGAZ-MIS'
'       ppw = 1
'       do epilogh3 with 18,40,ppw,2,'ΚΥΡΙΑ ΣΤΟΙΧΕΙΑ ΕΡΓΑΖΟΜΕΝΟΥ','ΣΤΟΙΧΕΙΑ ΜΙΣΘΟΔΟΣΙΑΣ'
'       if ppw=2
'          find_orisma()
'          restore screen from scr2
'          return 0
'       End If
'     End If
'
'       sele Fields
'       SET FILTER TO TRIM(ARXEIO) $ MARX1+MARX2+MARX3
'       GO TOP
'       For i = 1 To Min(fcount(), 50)
'         arr1 (i) = Field(i)
'       Next
'
'       dbedit ( 5,0,23,79,arr1)
'
'       restore screen from scr2
'
'       pa = Right(Trim(readvar()), 2)
'
'       sele kodik
'
'       repl ep&pa with Fields("epikef
'       repl f&pa with  trim(Fields("arxeio)+'("'+Fields("Field_Name
'       repl p&pa with  Fields("picture
'
'   End If
'
'return 0
'
'
'Function find_orisma()
'****************************************
'
'RETURN NIL
'
'
'
'
'
'
'proc insert_line
'*************************************************
'Local GetList:={}
'priv no_line:=0,k,aa,prev
'priv LAST_COLOR:=SETCOLOR()
'       no_line = Val(Right(Trim(readvar()), 2))
'       no_line = Max(1, no_line)
'       no_line = Min(15, no_line)
'       For K = 15 To no_line + 1 Step -1
'           aa = StrZero(K, 2)
'           prev = StrZero(K - 1, 2)
'           repl kodik("ep&aa with kodik("ep&prev
'           repl kodik("f&aa with kodik("f&prev
'           repl kodik("p&aa with kodik("p&prev
'       Next
'
'           aa = StrZero(no_line, 2)
'           repl kodik("ep&aa with ' '
'           repl kodik("f&aa with ' '
'           repl kodik("p&aa with '  '
'
'
'        ** ζωγραφίζω τα get
'        set color to &nw
'        for k=1 to kodik("nr
'            aa = StrZero(K, 2)
'            @ 1+k, 0 say k pict '99'
'            @ 1+k, 3 get kodik("ep&aa pict '@S20' valid find_field(kodik("ep&aa)=0
'            @ 1+k,25 get kodik("f&aa  pict '@S20'
'            @ 1+k,47 get kodik("p&aa  pict '@S20'
'            @ 1+k,70 get synola(k)    pict '9'
'        Next
'        SETCOLOR (LAST_COLOR)
'Return
'
'
'
'proc delete_line
'*************************************************
'Local GetList:={}
'priv no_line:=0,k,aa,epomeno
'priv LAST_COLOR:=SETCOLOR()
'       no_line = Val(Right(Trim(readvar()), 2))
'       no_line = Min(15, no_line)
'       no_line = Max(1, no_line)
'
'       For K = no_line To 22
'           aa = StrZero(K, 2)
'           epomeno = StrZero(K + 1, 2)
'           repl kodik("ep&aa with kodik("ep&epomeno
'           repl kodik("f&aa  with kodik("f&epomeno
'           repl kodik("p&aa  with kodik("p&epomeno
'       Next
'
'           aa = StrZero(15, 2)
'           repl kodik("ep&aa with ' '
'           repl kodik("f&aa with ' '
'           repl kodik("p&aa with '  '
'
'
'        ** ζωγραφίζω τα get
'        set color to &nw
'        for k=1 to kodik("nr
'            aa = StrZero(K, 2)
'            @ 1+k, 0 say k pict '99'
'            @ 1+k, 3 get kodik("ep&aa pict '@S20' valid find_field(kodik("ep&aa)=0
'            @ 1+k,25 get kodik("f&aa  pict '@S20'
'            @ 1+k,47 get kodik("p&aa  pict '@S20'
'            @ 1+k,70 get synola(k)    pict '9'
'        Next
'        SETCOLOR (LAST_COLOR)
'
'Return
'
'
'
'
'   Function MyErrorHandler(e, old)
'  *************************************************************
'      //
'   lOCAL cMessage:='',al(12)
'        // display message and traceback
'        if ( !Empty(e:osCode) )
'                cMessage += " (DOS Error " + LTRIM(e:osCode) + ") "
'        End
'
'         // build error message
'        cMessage := ErrorMessage(e)
'
'
'
'        // build options array
'        // aOptions := {"Break(Διακοπή)", "Quit"}
'        aOptions := {"Quit(Διακοπή)"}
'
'        if (e:canRetry)
'                AAdd(aOptions, "Retry(Ξαναπροσπαθώ)")
'        End
'
'        if (e:canDefault)
'                AAdd(aOptions, "Default(αγνοω-συνεχίζω)")
'        End
'
'
'        // put up alert box
'        nChoice := 0
'        while ( nChoice == 0 )
'
'                if ( Empty(e:osCode) )
'                        nChoice := Alert( cMessage, aOptions )
'
'                Else
'                        nChoice := Alert( cMessage + ;
'                                                        ";(DOS Error " + LTRIM(e:osCode) + ")", ;
'                                                        aOptions )
'                End
'
'
'                if ( nChoice == NIL )
'                End
'
'        End
'
'
'        if ( !Empty(nChoice) )
'
'                // do as instructed
'                if ( aOptions(nChoice) == "Break(Διακοπή)" )
'
'                elseif ( aOptions(nChoice) == "Quit(Διακοπή)" )
'                elseif ( aOptions(nChoice) == "Retry(Ξαναπροσπαθώ)" )
'                        return (.t.)
'
'                elseif ( aOptions(nChoice) == "Default(αγνοω-συνεχίζω)" )
'                        return (.f.)
'
'                End
'
'        End
'
'
'
'
'
'
'
'
'
'
'
'
'        Print cMessage
'        i := 2
'
'
'       * while ( !Empty(ProcName(i)) )
'       Print e: Description
'       For i = 1 To 12
'             if ( !Empty(ProcName(i)) )
'                al (i) = Trim(ProcName(i)) + "(" + Str(ProcLine(i)) + ")  "
'                Print al; (i)
'             Else
'                al(i)='....'
'             End If
'       Next
'
'
'     if .not. file('err.txt')
'       IF (nHandle := FCREATE('err.txt' , FC_NORMAL)) == -1
'         Print "File cannot be created:", FERROR()
'         QUIT
'       End If
'     Else
'       nHandle := FOPEN("err.txt", FO_READWRITE + FO_SHARED)
'       IF FERROR() != 0
'         Print "Cannot open file, DOS error ", FERROR()
'         QUIT
'       End If
'     End If
'
'
'       nLength := FSEEK(nHandle, 0, FS_END)
'
'      FWRITE(nHandle,dtoc(date())+' '+time()+' '+CMESSAGE +chr(10) )
'      For i = 1 To 12
'          if al(i)<>'....'
'             FWRITE(nHandle,'   '+al(i)+chr(10) )
'          End If
'      Next
'      FCLOSE (nHandle)
'       wait ''
'       BREAK      && objError          // Return error object to RECOVER
'     RETURN NIL
'
'
'
'
'
'
'
'
'
'
'
'Procedure opos_allo
'********************************
'Local GetList:={},sc
'priv m:=1,r
'R = readvar()
'sc = savescreen(24, 0, 24, 79)
'@ 24,0 say 'Οπως το πεδίο ' get m
'Read
'M = StrZero(M, 2)
'repl &r with kodik("f&m
'restscreen(24,0,24,79,sc)
'Return
'
'Function make_relations()
'**********************************************
'    if  len(mRelation2)>0
'       SELE &mARX2
'       if upper(indexkey(0)) = upper ( mrelation2 )
'          *
'       Else
'          * index on &mrelation2 to dok2&xeirisths
'       End If
'       sele &mARX1
'       set relation to &mrelation2 into &mARX2 ADDIT
'    End If
'
'    if  len(mRelation3)>0
'        set relation to &mrelation3 into &mARX3 ADDIT
'    End If
'
'    if len(mRelation4)>0
'        set relation to &mrelation4 into &mARX4 ADDIT
'    End If
'return nil
'
'
'proc check_kodik(mFile)
'****************************************************************
'priv strupin:={}
'AADD( STRUPIN , {"NUMBER_ARX","N",  1,0} )
'AADD( STRUPIN , {"ARX1      ","C", 40,0} )
'AADD( STRUPIN , {"ARX2      ","C", 40,0} )
'AADD( STRUPIN , {"ARX3      ","C", 40,0} )
'AADD( STRUPIN , {"ARX4      ","C", 40,0} )
'AADD( STRUPIN , {"RELATION4 ","C", 40,0} )
'AADD( STRUPIN , {"RELATION2 ","C", 40,0} )
'AADD( STRUPIN , {"RELATION3 ","C", 40,0} )
'AADD( STRUPIN , {"ART       ","C", 10,0} )
'AADD( STRUPIN , {"SYNOLO    ","C", 30,0} )
'AADD( STRUPIN , {"NO        ","C", 15,0} )
'AADD( STRUPIN , {"TAG       ","C", 60,0} )
'AADD( STRUPIN , {"EP01      ","C", 30,0} )
'AADD( STRUPIN , {"EP02      ","C", 30,0} )
'AADD( STRUPIN , {"EP03      ","C", 30,0} )
'AADD( STRUPIN , {"EP04      ","C", 30,0} )
'AADD( STRUPIN , {"EP05      ","C", 30,0} )
'AADD( STRUPIN , {"EP06      ","C", 35,0} )
'AADD( STRUPIN , {"EP07      ","C", 35,0} )
'AADD( STRUPIN , {"EP08      ","C", 30,0} )
'AADD( STRUPIN , {"EP09      ","C", 30,0} )
'AADD( STRUPIN , {"EP10      ","C", 30,0} )
'AADD( STRUPIN , {"EP11      ","C", 30,0} )
'AADD( STRUPIN , {"EP12      ","C", 30,0} )
'AADD( STRUPIN , {"EP13      ","C", 30,0} )
'AADD( STRUPIN , {"EP14      ","C", 30,0} )
'AADD( STRUPIN , {"EP15      ","C", 30,0} )
'AADD( STRUPIN , {"F01       ","C",180,0} )
'AADD( STRUPIN , {"F02       ","C",180,0} )
'AADD( STRUPIN , {"F03       ","C",180,0} )
'AADD( STRUPIN , {"F04       ","C",180,0} )
'AADD( STRUPIN , {"F05       ","C",180,0} )
'AADD( STRUPIN , {"F06       ","C",180,0} )
'AADD( STRUPIN , {"F07       ","C",180,0} )
'AADD( STRUPIN , {"F08       ","C",180,0} )
'AADD( STRUPIN , {"F09       ","C",180,0} )
'AADD( STRUPIN , {"F10       ","C",180,0} )
'AADD( STRUPIN , {"F11       ","C",180,0} )
'AADD( STRUPIN , {"F12       ","C",180,0} )
'AADD( STRUPIN , {"F13       ","C",180,0} )
'AADD( STRUPIN , {"F14       ","C",180,0} )
'AADD( STRUPIN , {"F15       ","C",180,0} )
'AADD( STRUPIN , {"P01       ","C", 30,0} )
'AADD( STRUPIN , {"P02       ","C", 30,0} )
'AADD( STRUPIN , {"P03       ","C", 30,0} )
'AADD( STRUPIN , {"P04       ","C", 30,0} )
'AADD( STRUPIN , {"P05       ","C", 30,0} )
'AADD( STRUPIN , {"P06       ","C", 30,0} )
'AADD( STRUPIN , {"P07       ","C", 30,0} )
'AADD( STRUPIN , {"P08       ","C", 30,0} )
'AADD( STRUPIN , {"P09       ","C", 30,0} )
'AADD( STRUPIN , {"P10       ","C", 30,0} )
'AADD( STRUPIN , {"P11       ","C", 30,0} )
'AADD( STRUPIN , {"P12       ","C", 60,0} )
'AADD( STRUPIN , {"P13       ","C", 30,0} )
'AADD( STRUPIN , {"P14       ","C", 30,0} )
'AADD( STRUPIN , {"P15       ","C", 30,0} )
'AADD( STRUPIN , {"SYNTWH    ","C", 80,0} )
'AADD( STRUPIN , {"SYNTIF    ","C",140,0} )
'AADD( STRUPIN , {"ARXEIO    ","C", 25,0} )
'AADD( STRUPIN , {"NR        ","N",  4,0} )
'AADD( STRUPIN , {"EPIKEF    ","C",150,0} )
'AADD( STRUPIN , {"S_EJECT   ","N",  2,0} )
'AADD( STRUPIN , {"FONTS     ","C", 30,0} )
'AADD( STRUPIN , {"K_HM1     ","C", 30,0} )
'AADD( STRUPIN , {"K_HM2     ","C", 30,0} )
'AADD( STRUPIN , {"K_CH1     ","C", 30,0} )
'AADD( STRUPIN , {"K_CH2     ","C", 30,0} )
'AADD( STRUPIN , {"K_CH3     ","C", 30,0} )
'AADD( STRUPIN , {"EP16      ","C", 50,0} )
'AADD( STRUPIN , {"EP17      ","C", 50,0} )
'AADD( STRUPIN , {"EP18      ","C", 50,0} )
'AADD( STRUPIN , {"EP19      ","C", 30,0} )
'AADD( STRUPIN , {"EP20      ","C", 30,0} )
'AADD( STRUPIN , {"F16       ","C",180,0} )
'AADD( STRUPIN , {"F17       ","C",180,0} )
'AADD( STRUPIN , {"F18       ","C",180,0} )
'AADD( STRUPIN , {"F19       ","C",180,0} )
'AADD( STRUPIN , {"F20       ","C",180,0} )
'AADD( STRUPIN , {"P16       ","C", 30,0} )
'AADD( STRUPIN , {"P17       ","C", 30,0} )
'AADD( STRUPIN , {"P18       ","C", 30,0} )
'AADD( STRUPIN , {"P19       ","C", 30,0} )
'AADD( STRUPIN , {"P20       ","C", 30,0} )
'AADD( STRUPIN , {"SYNT_ALAG ","C", 30,0} )
'AADD( STRUPIN , {"BSEIRA    ","C",160,0} )
'AADD( STRUPIN , {"SUM_SELTXT","C", 30,0} )
'AADD( STRUPIN , {"EK_METAF  ","N",  1,0} )
'AADD( STRUPIN , {"SE_METAF  ","N",  1,0} )
'AADD( STRUPIN , {"DIAX      ","C",  1,0} )
'AADD( STRUPIN , {"TOTAL     ","C",120,0} )
'AADD( STRUPIN , {"EP21      ","C", 30,0} )
'AADD( STRUPIN , {"EP22      ","C", 30,0} )
'AADD( STRUPIN , {"EP23      ","C", 30,0} )
'AADD( STRUPIN , {"EP24      ","C", 30,0} )
'AADD( STRUPIN , {"EP25      ","C", 30,0} )
'AADD( STRUPIN , {"P21       ","C", 30,0} )
'AADD( STRUPIN , {"P22       ","C", 30,0} )
'AADD( STRUPIN , {"P23       ","C", 30,0} )
'AADD( STRUPIN , {"P24       ","C", 30,0} )
'AADD( STRUPIN , {"P25       ","C", 30,0} )
'AADD( STRUPIN , {"F21       ","C",180,0} )
'AADD( STRUPIN , {"F22       ","C",180,0} )
'AADD( STRUPIN , {"F23       ","C",180,0} )
'AADD( STRUPIN , {"F24       ","C",180,0} )
'AADD( STRUPIN , {"F25       ","C",180,0} )
'AADD( STRUPIN , {"F26       ","C",180,0} )
'AADD( STRUPIN , {"F27       ","C",180,0} )
'AADD( STRUPIN , {"F28       ","C",180,0} )
'AADD( STRUPIN , {"F29       ","C",180,0} )
'AADD( STRUPIN , {"F30       ","C",180,0} )
'AADD( STRUPIN , {"EP26      ","C", 30,0} )
'AADD( STRUPIN , {"EP27      ","C", 30,0} )
'AADD( STRUPIN , {"EP28      ","C", 30,0} )
'AADD( STRUPIN , {"EP29      ","C", 30,0} )
'AADD( STRUPIN , {"EP30      ","C", 30,0} )
'AADD( STRUPIN , {"P26       ","C", 30,0} )
'AADD( STRUPIN , {"P27       ","C", 30,0} )
'AADD( STRUPIN , {"P28       ","C", 30,0} )
'AADD( STRUPIN , {"P29       ","C", 30,0} )
'AADD( STRUPIN , {"P30       ","C", 30,0} )
'AADD( STRUPIN , {"EPIKEF2   ","C", 70,0} )
'AADD( STRUPIN , {"EPIK_2    ","C", 70,0} )
'AADD( STRUPIN , {"EPIK_3    ","C", 70,0} )
'AADD( STRUPIN , {"K_CH4     ","C", 30,0} )
'AADD( STRUPIN , {"K_CH5     ","C", 30,0} )
'AADD( STRUPIN , {"PEDIA_TOT ","C", 80,0} )
'AADD( STRUPIN , {"FOR_TOT   ","C", 80,0} )
'AADD( STRUPIN , {"M_HEAD1   ","C", 40,0} )
'AADD( STRUPIN , {"M_HEAD2   ","C", 40,0} )
'AADD( STRUPIN , {"M_HEAD3   ","C", 40,0} )
'AADD( STRUPIN , {"M_HEAD4   ","C", 40,0} )
'
'lk_check (mFile)
'Return
'
'Procedure make_totals
'***********************************
'       if len(b)>1
'
'                  **************** κανω τις σούμες  **********************
'           if kodik("total<>'     '
'
'                 mtotal   =trim(kodik("total   )
'                 mfor_tot =trim(kodik("for_tot )
'
'                 for _kt=1 to 6
'                     _mTot(_kt) = subs( kodik("pedia_tot , 01+(_kt-1)*10 , 10 )
'                 Next
'
'
'                 _ST=0
'                 for _kt =1 to 6
'                     if _mTot(_kt)<>' '
'                         aa = alltrim ( str ( _kt ) )
'                        _st = _st + 1
'                         m&aa = _mTot(_kt)
'                     Else
'                        exit
'                     End If
'                 Next
'
'
'                 index on &mtotal to dok&xeirisths
'                 do case
'                    case _st=1
'                         total on  &b fields &m1 to dok&xeirisths for &mfor_tot
'                    case _st=2
'                         total on  &b fields &m1,&m2 to dok&xeirisths for &mfor_tot
'                    case _ST=3
'                         total on  &b fields &m1,&m2,&m3 to dok&xeirisths for &mfor_tot
'                    case _ST=4
'                         total on  &b fields &m1,&m2,&m3,&m4 to dok&xeirisths for &mfor_tot
'                    case _ST=5
'                         total on  &b fields &m1,&m2,&m3,&m4,&m5     to dok&xeirisths for &mfor_tot
'                    case _ST=6
'                         total on  &b fields &m1,&m2,&m3,&m4,&m5,&m6 to dok&xeirisths for &mfor_tot
'                    otherwise
'                         wait ' δεν όρισες πεδία που θα αθροίσω '
'                         Close Data
'                         Return
'                 endcase
'
'
'              use dok&xeirisths alias &marx1 exclu
'            * 20-2-2000    index on &b tag dok&xeirisths to dok&xeirisths
'
'              if  len(mRelation2)>0
'                  sele &mARX1
'                  set relation to &mrelation2 into &mARX2 ADDIT
'              End If
'              if  len(mRelation3)>0
'                  sele &mARX1
'                  set relation to &mrelation3 into &mARX3 ADDIT
'              End If
'              index on &b tag dok&xeirisths to dok&xeirisths
'
'
'           Else
'              Print b
'              Print xeirisths
'              Print dok
'              index on &b to dok___&xeirisths
'           End If
'           **************** κανω τις σούμες  **********************
'       End If
'
'Return
'
'
'proc make_mis_totals
'*********************************************************
'       if len(b)>1
'           gm_fILTER = TRIM(KODIK("SYNTIF)
'           if val(b)>0
'               set order to val(b)
'           Else
'               index on &b to dok&xeirisths for &gm_fILTER      && tag dok&xeirisths to dok&xeirisths
'           End If
'
'           **************** κανω τις σούμες  **********************
'           if kodik("total<>'     '
'              mtotal=trim(kodik("FOR_tot)
'              *  total on &b fields &mtotal  to dok&xeirisths
'
'
'
'              total on  &b fields     w1,w2,w3,w4,w5,w6,w7,w8,w9,w10,;
'                                      w11,w12,w13,w14,w15,w16,w17,w18,w19,w20,;
'                                      w21,w22,w23,w24,w25,w26,w27,w28,w29,w30,;
'                                      w31,w32,w33,w34,w35,w36,w37,w38,w39,w40,;
'                                      w41,w42,w43,w44,w45,w46,w47,w48,w49,w50,;
'                                      w51,w52,w53,w54,w55,w56,w57,w58,w59,w60,;
'                                      w61,w62,w63,w64,w65,w66,w67,w68,w69,w70,;
'                                      w71,w72,w73,w74,w75,w76,w77,w78,w79,w80,;
'                                      w81,w82,w83,w84,w85,w86,w87,w88,w89,w90,;
'                                      w91,w92,w93,w94,w95,w96,w97,w98,w99     ;
'                                  to dok&xeirisths for &mtotal
'              use dok&xeirisths alias &marx1 exclu
'              index on &b to dok&xeirisths   && tag dok&xeirisths to dok&xeirisths
'
'              if  len(mRelation2)>0
'                  SELE &mARX2
'                  set order to 1  &&  tag &mrelation2
'                  sele &mARX1
'                  set relation to &mrelation2 into &mARX2 ADDIT
'              End If
'           End If
'           **************** κανω τις σούμες  **********************
'
'       End If
'Return
'
'
'func errorMessage(e)
'local cMessage
'***********************************************
'
'        // start error message
'        cMessage := if( e:severity > ES_WARNING, "Error ", "Warning " )
'
'
'        // add subsystem name if available
'        if ( ValType(e:subsystem) == "C" )
'                cMessage += e:subsystem()
'        Else
'                cMessage += "???"
'        End
'
'
'        // add subsystem's error code if available
'        if ( ValType(e:subCode) == "N" )
'                cMessage += ("/" + LTRIM(STR(e:subCode) ))
'        Else
'                cMessage += "/???"
'        End
'
'
'        // add error description if available
'        if ( ValType(e:description) == "C" )
'                cMessage += ("  " + e:description)
'        End
'
'
'        // add either filename or operation
'        if ( !Empty(e:filename) )
'                cMessage += (": " + e:filename)
'
'        elseif ( !Empty(e:operation) )
'                cMessage += (": " + e:operation)
'
'        End
'
'
'return (cMessage)
'
'
'Function Find_timol(key)
'*****************************************
'* key   =  atim+dtos(hme)
'priv aa
'aa=select()
'do open_data with .f.,2,'EID','EGGTIM'
'
'sele EGGTIM
'
'seek key
'do while .not. eof() .and. key=eggtim("atim+dtos(eggtim("hme)
'  sele EID
'  seek eggtim("kode
'  sele EGGTIM
'  printer.print+1,15 say '- '+ eid("ono+'  '+trans(poso,'9999.9')+'  '+trans(timm,'9999,999.99')+'  '+trans(timm*poso,'999,999.99')
'  Skip
'enddo
'aa = alltrim(Str(aa))
'sele &aa
'
'Return '  '
'
'Function db17_udf()
'*****************************************
'para dbmode, fld
'getit=arr2(fld)
'
'do case
'  Case dbmode = 0
'     return 1
'  Case lastkey() = 27
'     return 0
'  Case lastkey() = K_INS
'       ADD_REC (0)
'       RETURN 1
'  Case lastkey() = K_ENTER
'     set cursor on
'     @ row(),col() get &getit
'     if rec_lock(0)
'       Read
'       unlock
'     End If
'     set cursor off
'     return 1
'  otherwise
'     return 1
'endcase
'
'
'
'
'
'
Function to928(ByVal string_ As String) As String
        '<EhHeader>
        On Error GoTo to928_Err
        '</EhHeader>

        Dim a$, k As Integer, s As String, T As Integer, s928 As String, s437 As String

        'metatrepei eggrafo apo 437->928
100     s928 = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ-αβγδεζηθικλμνξοπρστυφχψω-ςάέήίόύώ"
102     s437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰ-ªαβγεζηι"    ' saehioyv
        's437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰª" 'αβγεζηι"

104     a$ = string_

        '                                                        saehioyv
        'GoTo 11
        'Open Text2.Text For Output As #2
        'Open Text1.Text For Input As #1
        'Do While Not EOF(1)
        '  Line Input #1, a$
106     For k = 1 To Len(a$)
108         s = mID(a$, k, 1)
110         T = InStr(s437, s)

112         If T > 0 Then
114             Mid$(a$, k, 1) = mID$(s928, T, 1)
            End If

        Next

116     to928 = a$


        '<EhFooter>
        Exit Function

to928_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.to928 " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function

Function toGreekLish(ByVal string_ As String) As String
        '<EhHeader>
        On Error GoTo toGreeklish_Err
        '</EhHeader>

        Dim a$, k As Integer, s As String, T As Integer, s928 As String, s437 As String

        'metatrepei eggrafo apo 437->928
        
        
      


        
        'ΑΒΓΔΕΖΗ Θ  ΙΚΛΜΝΞΟΠΡΣΤΥΦΧ Ψ  Ωαβγδεζη θ  ικλμνξοπρστυφχ ψ  ωάέήίόύώ
'100     s928 = "ΑΒΓΔΕΖΗΙΚΛΜΝΞΟΠΡΣΤΥΦΧΩαβγδεζηικλμνξοπρστυφχωάέήίόύώΘθΨψ "
'"ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ-αβγδεζηθικλμνξοπρστυφχψω-ςάέήίόύώ"
'102     s437 = "AVGDEZIIKLMNXOPRSTIFHOavgdeziiklmnxoprstifhoaeiioioΘθΨψ-"
'"ABGDEZH8IKLMNXOPRSTYFXCW-abgdezh8iklmnxoprstyfxcw-saehioyw"    ' saehioyv
        's437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰª" 'αβγεζηι"
        '€‚ƒ„…†‡‰‹‘’“”•–—™› ΅Ά£¤¥¦§¨©«¬­®―ΰ

104     a$ = string_

        '                                                        saehioyv
        'GoTo 11
        'Open Text2.Text For Output As #2
        'Open Text1.Text For Input As #1
        'Do While Not EOF(1)
        '  Line Input #1, a$
106     For k = 1 To Len(a$)
108         s = mID(a$, k, 1)
110        ' T = InStr(s928, s)
            If InStr("ΑΆαά", s) > 0 Then s = "a"
            If InStr("Ββ", s) > 0 Then s = "b"
            If InStr("Γγ", s) > 0 Then s = "g"
            If InStr("Δδ", s) > 0 Then s = "d"
            If InStr("ΕΈεέ", s) > 0 Then s = "e"
            If InStr("ΗΉηήΙΊιίΪΐ", s) > 0 Then s = "i"
            If InStr("Ζζ", s) > 0 Then s = "z"
            If InStr("Κκ", s) > 0 Then s = "k"
            If InStr("Λλ", s) > 0 Then s = "l"
            If InStr("Μμµ", s) > 0 Then s = "m"
            If InStr("Νν", s) > 0 Then s = "n"
            If InStr("ΟΌοόΩΏωώ", s) > 0 Then s = "o"
            If InStr("Ππ", s) > 0 Then s = "p"
            If InStr("Ρρ", s) > 0 Then s = "r"
            If InStr("Σσς", s) > 0 Then s = "s"
            If InStr("Ττ", s) > 0 Then s = "t"
            If InStr("Χχ*", s) > 0 Then s = "x"
            If InStr("Φφ", s) > 0 Then s = "f"
            If InStr("ΥΎυύΫϋ", s) > 0 Then s = "u"
            If InStr(",_ #""/()·:", s) > 0 Then s = "-"
            If InStr("Φφ", s) > 0 Then s = "f"
            If InStr("Φφ", s) > 0 Then s = "f"
                
            Mid$(a$, k, 1) = s
112       '  If T > 0 Then
114
          '  End If

        Next
        
        
        a$ = Replace$(a$, "Θ", "th")
        a$ = Replace$(a$, "θ", "th")
        
        a$ = Replace$(a$, "Ψ", "ps")
        a$ = Replace$(a$, "ψ", "ps")
        
         a$ = Replace$(a$, "Ξ", "ks")
        a$ = Replace$(a$, "ξ", "ks")
        
        
        a$ = Replace$(a$, "™", "")
        a$ = Replace$(a$, "®", "")
        a$ = Replace$(a$, "΄", "")
        a$ = Replace$(a$, "'", "")
        a$ = Replace$(a$, "\", "")
         a$ = Replace$(a$, "/", "")
         
         a$ = Replace$(a$, "----", "-")
         
         a$ = Replace$(a$, "---", "-")
         a$ = Replace$(a$, "--", "-")
         
       
        
        a$ = Replace$(a$, "’", "")
       

116     toGreekLish = a$
        
         
        '<EhFooter>
        Exit Function

toGreeklish_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.toGreeklish_Err " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function




Function to437(ByVal string_ As String) As String
        '<EhHeader>
        On Error GoTo to437_Err
        '</EhHeader>

        Dim a$, k As Integer, s As String, T As Integer, s928 As String, s437 As String

        'metatrepei eggrafo apo 437->928
100     s928 = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ-αβγδεζηθικλμνξοπρστυφχψω-ςάέήίόύώ"
102     s437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰ-ªαβγεζηι"    ' saehioyv
        's437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰª" 'αβγεζηι"
        '€‚ƒ„…†‡‰‹‘’“”•–—™› ΅Ά£¤¥¦§¨©«¬­®―ΰ

104     a$ = string_

        '                                                        saehioyv
        'GoTo 11
        'Open Text2.Text For Output As #2
        'Open Text1.Text For Input As #1
        'Do While Not EOF(1)
        '  Line Input #1, a$
106     For k = 1 To Len(a$)
108         s = mID(a$, k, 1)
110         T = InStr(s928, s)

112         If T > 0 Then
114             Mid$(a$, k, 1) = mID$(s437, T, 1)
            End If

        Next

116     to437 = a$


        '<EhFooter>
        Exit Function

to437_Err:
        SAVE_ERROR Err.Description & vbCrLf & _
               "in ADOMERCNEW.Module5.to437 " & _
               "at line " & Erl


        Resume Next
        '</EhFooter>
End Function

'Sub g_ektyp(sql As String)
''POSITION EINAI TO RECORD ΤΟΥ REPORT
''notepad  hmm ok i checked the binary a bit out now..
''here are all command line arguments notepad takes
''
''/A <filename> open file as ansi
''/W <filename> open file as unicode
''/P <filename> print filename
''/PT <filename> <printername> <driverdll> <port>
''/.SETUP some weird stuff is happening i cant identify =) its enumerating the systemdir and opening notepad without the minimize option.. but dont ask me why
''
''hmm i think the best solution to your windowing problem would be to just move the window to where you want to have it
''close notepad so notepad saves the position size etc to HKEY_CURRENT_USER\Software\Microsoft\Notepad in the registry..
''then just export that registry.. create a batch file or something then that just imports that reg entry (regedit /s filename.reg) and opens your notepad after..
'
'    Dim DUM, N
'    Dim TT, F_T As New ADODB.Recordset
'
'
'    TT = GetCurrentTime()
'
'
'    F_T.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly
'
'    MDIForm1.Caption = (GetCurrentTime() - TT)
'
'
'    N = F_T.FIELDS.Count
'
'
'
'    f_tab(0) = 0
'    'f_tab(1) = 15
'    'f_tab(2) = 50
'    'f_tab(3) = 60
'    'f_tab(4) = 70
'
'    Dim k
'    Dim m_s
'
'    ' m_s = f_tab(0)
'
'
'    'For K = 1 To f_t.Fields.Count - 1
'    '  f_tab(K) = 2 + m_s + f_t.Fields(K - 1).Size
'    'Next
'
'    'm_s = f_tab(0)
'
'
'
'
'
'
'
'
'    ' Dim F_T As New ADODB.Recordset
'
'
'    TT = GetCurrentTime()
'
'
'    F_T.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly
'
'    N = F_T.FIELDS.Count - 1
'
'
'
'    For k = 1 To N
'        If F_T(k - 1).Type = 8 Then    'DATE
'            f_tab(k) = 2 + f_tab(k - 1) + 10
'        Else
'            f_tab(k) = 2 + f_tab(k - 1) + F_T.FIELDS(k - 1).DefinedSize
'        End If
'    Next
'
'
'    'dum = print3_xar(n, "0", "κωδικός", "F_KODIK!FIELDS", "xxxxxxx")
'
'
'
'
'
'
'
'
'End Sub

