Attribute VB_Name = "NOTES"




'mydata 11/9/2021
' ХЕКОУЛЕ МА ЛПОЯЕИ МА ДИОЯХЫМЕИ ТОМ КОЦО АПАККАЦГР ЭПЫР ЕПИСГР ЙАИ ТО СВЕТИЙО ПАЯАСТАТИЙО СТА ПИСТЫТИЙА
' ЕПИСГР СТО АЙУЯЫТИЙО kai sto pistvtiko  МА ЙЯАТАЕИ ТО marknumber




'10-2-14  бектиыхгйе г диояхысг & летасглатислос ма доукеуоум ле ема п╒тгла пкгйтяоу

'ver 78
'ЕЙТУПЫСГ СЕ ПОККАПКщР СЕИЯщР

'VER.13.0.20
'26-7-13
'стяоццукела сто 0.01
'тис паяаццекиес тис апохгйеуеи ейтос апо то C:\PDF KAI сто C:\PDFPARAGG
'ам дысы паяалетяо F_PDFPARAGG=1

'VER 13.0.12
' думатотгта ма яухлифы том аяихло тым амтитупым еиспяангс / пкгяылгс

'VER 13.0.11 17-6-13  стгм ыяиламсг упокоипым диояхыхгйе то пяобкгла ле тоус лецакоус йыдийоус пекатым > 6 ьгжха

'VER 13.0.10 паяоусиасг тилокоциоу (йаятекка,паяоус.емос пая/йоу,пая.йата  глея/миа)
' ле амакусг йахаягс аниас,жпа йаи сумокоу
' диацяажг епитацгс ле вягсг то  ID тгс епитацгс

'21-3-2013
' a = a + "where HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'  GROUP BY KODE"   'AND ascii(left(ATIM,1)) in (" + pol + ")

'----------------------------------------------------------------------------------------------------------------------
'UTILITIES
' http://www.mztools.com/v3/download.aspx

'24/6/2007 ВЯЫЛАТИФЫ ТО f10
'24/6 ЕПАМАЖЕЯЫ СЕ ЕЙЙЯЕЛОТГТА -БЦАФЫ АПЭ ЕЙЙЯЕЛЭТГТА

'25-5-07 лпояы ма лпкойаяы ама цяаллг паяастатийоу то жпа ч ТГМ ЕЙПТЫСГ

'17-8-06 ЙИМГСГ ЕИДОУС АПО ЕОС

'1-7-2006 аутолата амтицяажа
'2-8-2006 NA EPILEGEI SAN DEFAULT TO пAяAсTATIKO пOY хекеи
'25/12/2005  бкепеи пакиа вяомиа стом пекатг
'            & аявийо йаи текийо упокоипо та бкепы стгм еисацыцг-диояхысг пекатг
'23/5/2005 ЦЯАЖИЙч ПАЯэСТАСГ

'11/5/2008  LIST BOX ME SYSRIMO ALLAZO THN SEIRA EKTYPOSIS

'vivotek   kamera  asyrmath kinhth

'ALTER TABLE MEM
'ADD HME2 SMALLDATETIME NULL
' СГЛЕИОСГ 26-6-2010
' 2h shmeiosh 26/6/2010

'Sub sql_COMMAND_LINE()
'    Dim a
'    a = 0
'
'
'    '26/11/05
'    'Command LINE
'    'doksql:  Script file
'    ' select * FROM PINAKES
'    ' GO
'
'    'apotel: apotelesma toy select
'
'    ' ISQLW -S FLAGAKIS -d MERCURY -U sa -P 38983 -i DOKSQL -o apotel
'    'ГЛЕЯОЛГМИА ЫЯЕР АПАСВЭКГСГР  ЕЯЦАСИА    ТЕВМИТГР    ЕИДОР( ЕЦЙАТАСТАСГ - ЕПИСЙЕУч ОМ СИТЕ )
'    '                                         ( ЕЯЦАСИА АКТЕЙ - ЕПИСЙЕУГ СЕ СЕЯБИР )
'    '                                           ( ЕЦЙАТАСТАСГ ПЯОЦЯАЛЛэТЫМ КОЦИСТИЙЧМ
'    '(ЕЙПАъДЕУСГ
'    '
'
'
'
'    ' KINHSH ANA MHNA    KYKLOFORIAKI TAXITITA
'    ' select ROUND(sum(XRE)/3+0.5,0) AS POL,EG.KODE,E.ONO  FROM EGGTIM EG INNER JOIN EID E  ON EG.KODE=E.KOD  WHERE  E.AEG=11 AND E.KODLOG='1' GROUP BY EG.KODE,E.ONO
'    '  select ROUND(sum(XRE)/3+0.5,0) AS POL,EG.KODE,E.ONO  INTO DOKMEL FROM EGGTIM EG INNER JOIN EID E  ON EG.KODE=E.KOD  WHERE  E.AEG=11 AND E.KODLOG='1' GROUP BY EG.KODE,E.ONO
'
'    'UPDATE EID SET EPIUYP=DOKMEL.POL FROM EID INNER JOIN DOKMEL ON EID.KOD=DOKMEL.KOD WHERE
'
'End Sub

'Simple Example: Lets say you have a Students table, and a Lockers table.
'
'Each student can be assigned to a locker, so there is a "LockerNumber" column in the student table. More than one student could potentially be in a single locker, but especially at the begining of the school year, you may have some incoming students without lockers and some lockers that have no students assigned.
'
'For the sake of this example, lets say you have 100 students, 70 of which have lockers. You have a total of 50 lockers, 40 of which have at least 1 student.
'
'INNER JOIN is equivalent to "show me all students with lockers".
'Any students without lockers, or any lockers without students are missing.
'Returns 70 rows
'
'LEFT OUTER JOIN would be "show me all students, with their corresponding locker if they have one".
'This might be a general student list, or could be used to identify students with no locker.
'Returns 100 rows
'
'RIGHT OUTER JOIN would be "show me all lockers, and the students assigned to them if there are any".
'This could be used to identify lockers that have no students assigned, or lockers that have too many students.
'Returns 80 rows (list of 70 students in the 40 lockers, plus the 10 lockers with no student)
'
'FULL OUTER JOIN would be silly and probably not much use.
'Something like "show me all students and all lockers, and match them up where you can"
'Returns 110 rows (all 100 students, including those without lockers. Plus the 10 lockers with no student)
'
'CROSS JOIN is also fairly silly in this scenario.
'It doesn 't use the linked "lockernumber" field in the students table, so you basically end up with a big giant list of every possible student-to-locker pairing, whether or not it actually exists.
'Returns 5000 rows (100 students x 50 lockers). Could be useful (with filtering) as a starting point to match up the new students with the empty lockers.

'============================================================================================
'CREATE PROC SearchAllTables
'(
'@SearchStr nvarchar(100)
')
'AS
'BEGIN
'
'-- Copyright ╘ 2002 Narayana Vyas Kondreddi. All rights reserved.
'-- Purpose: To search all columns of all tables for a given search string
'-- Written by: Narayana Vyas Kondreddi
'-- Site: http://vyaskn.tripod.com
'-- Tested on: SQL Server 7.0 and SQL Server 2000
'-- Date modified: 28th July 2002 22:50 GMT
'
'
'CREATE TABLE #Results (ColumnName nvarchar(370), ColumnValue nvarchar(3630))
'
'SET NOCOUNT ON
'
'DECLARE @TableName nvarchar(256), @ColumnName nvarchar(128), @SearchStr2 nvarchar(110)
'SET  @TableName = ''
'SET @SearchStr2 = QUOTENAME('%' + @SearchStr + '%','''')
'
'WHILE @TableName IS NOT NULL
'BEGIN
'    SET @ColumnName = ''
'    SET @TableName =
'    (
'        SELECT MIN(QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME))
'        From INFORMATION_SCHEMA.Tables
'        WHERE       TABLE_TYPE = 'BASE TABLE'
'            AND QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME) > @TableName
'            AND OBJECTPROPERTY(
'                    OBJECT_ID(
'                        QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME)
'                         ), 'IsMSShipped'
'                           ) = 0
'    )
'
'    WHILE (@TableName IS NOT NULL) AND (@ColumnName IS NOT NULL)
'    BEGIN
'        SET @ColumnName =
'        (
'            SELECT MIN(QUOTENAME(COLUMN_NAME))
'            From INFORMATION_SCHEMA.Columns
'            WHERE       TABLE_SCHEMA    = PARSENAME(@TableName, 2)
'                AND TABLE_NAME  = PARSENAME(@TableName, 1)
'                AND QUOTENAME(COLUMN_NAME) > @ColumnName
'        )
'
'        IF @ColumnName IS NOT NULL
'        BEGIN
'            INSERT INTO #Results
'            EXEC
'            (
'                'SELECT ''' + @TableName + '.' + @ColumnName + ''', LEFT(CONVERT(varchar(max), ' + @ColumnName + '), 3630)
'                From ' + @TableName + ' (NOLOCK) ' +
'                ' WHERE CONVERT(varchar(max), ' + @ColumnName + ') LIKE ' + @SearchStr2
'            )
'        End
'    End
'End
'
'SELECT ColumnName, ColumnValue FROM #Results
'End

' 'ьавмы ма ды ам упаявеи то паяастатийо гдг ( апо кахос)
'                 Dim r60 As New ADODB.Recordset
'                 r60.Open "select * FROM TIM WHERE ATIM='" + Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text + "'", Gdb, adOpenDynamic, adLockOptimistic
'                 If r60.EOF Then
'                    ' OK
'                 Else
'                    'намаупаявеи то тилокоцио опоте паиямеи то еполемо
'                    r60.Close
'                    r60.Open "select MAX(ATIM) FROM TIM WHERE LEFT(ATIM,1)='" + Right(PARAS.Text, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
'                    Text1.Text = LTrim(str(1 + Val(mID(r60.Atim, 2, 6))))
'                 End If
'                r60.Close
'
'ALTER TABLE EGG
'ADD CONSTRAINT FIDEGG
'Foreign Key(IDtim)
'References TIM(ID_NUM)

'ALTER TABLE EGGTIM
'ADD CONSTRAINT FIDNUM
'Foreign Key(ID_NUM)
'References TIM(ID_NUM)

'EGG: стис вяеыпистысеис йаи стис епитацес евеи то IDTIM NULL
'опоте дем епгяеафеи то CONSTARINT TOY TIM
'дгкадг ои еццяажес поу пяоеявомтаи апо тгм тилокоцгсг
'йяеломтаи апо то тил

'пяосовг ле йимгсеис талеиайгс поу евоум нейяеластес KINHSEIS EGGTIM
'циа ма евы тгм ломадийотгта тоу TIM (KLEIDI)
' ALTER TABLE TIM  ADD CONSTRAINT KLEIDI  UNIQUE ( KLEIDI);



