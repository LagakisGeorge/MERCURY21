Attribute VB_Name = "Rtf2Html"
'**********************************************************
'*             Rich HTML by Joseph Huntley               *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'*                                                        *
'*  Made:  October 4, 1999                                *
'*  Level: Beginner                                       *
'**********************************************************
'*   The form here are only used to demonstrate how to    *
'* use the function 'RichToHTML' and 'HTMLToRich'. You    *
'* may copy the function into your project for use. If    *
'* you need any help please e-mail me.                    *                            *
'**********************************************************
'* Notes: None                                            *
'**********************************************************

Function RichToHTML(rtbRichTextBox As RichTextLib.RichTextBox, _
                    Optional lngStartPosition As Long, _
                    Optional lngEndPosition As Long) As String

    '**********************************************************
    '*            Rich To HTML by Joseph Huntley              *
    '*               joseph_huntley@email.com                 *
    '*                http://joseph.vr9.com                   *
    '**********************************************************
    '*   You may use this code freely as long as credit is    *
    '* given to the author, and the header remains intact.    *
    '**********************************************************

    '--------------------- The Arguments -----------------------
    'rtbRichTextBox     - The rich textbox control to convert.
    'lngStartPosition   - The character position to start from.
    'lngEndPosition     - The character position to end at.
    '-----------------------------------------------------------
    'Returns:     The rich text converted to HTML.

    'Description: Converts rich text to HTML.

    Dim blnBold   As Boolean, blnUnderline As Boolean, blnStrikeThru As Boolean

    Dim blnItalic As Boolean, strLastFont As String, lngLastFontColor As Long

    Dim strHTML   As String, lngColor As Long, lngRed As Long, lngGreen As Long

    Dim lngBlue   As Long, lngCurText As Long, strHex As String, intLastAlignment As Integer

    Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2

    'check for lngStartPosition ad lngEndPosition

    If IsMissing(lngStartPosition&) Then lngStartPosition& = 0
    If IsMissing(lngEndPosition&) Then lngEndPosition& = Len(rtbRichTextBox.Text)

    lngLastFontColor& = -1 'no color

    For lngCurText& = lngStartPosition& To lngEndPosition&
        rtbRichTextBox.SelStart = lngCurText&
        rtbRichTextBox.SelLength = 1
   
        If intLastAlignment% <> rtbRichTextBox.SelAlignment Then
            intLastAlignment% = rtbRichTextBox.SelAlignment
              
            Select Case rtbRichTextBox.SelAlignment

                Case AlignLeft: strHTML$ = strHTML$ & "<p align=left>"

                Case AlignRight: strHTML$ = strHTML$ & "<p align=right>"

                Case AlignCenter: strHTML$ = strHTML$ & "<p align=center>"
            End Select
                
        End If
   
        If blnBold <> rtbRichTextBox.SelBold Then
            If rtbRichTextBox.SelBold = True Then
                strHTML$ = strHTML$ & "<b>"
            Else
                strHTML$ = strHTML$ & "</b>"
            End If

            blnBold = rtbRichTextBox.SelBold
        End If

        If blnUnderline <> rtbRichTextBox.SelUnderline Then
            If rtbRichTextBox.SelUnderline = True Then
                strHTML$ = strHTML$ & "<u>"
            Else
                strHTML$ = strHTML$ & "</u>"
            End If

            blnUnderline = rtbRichTextBox.SelUnderline
        End If

        If blnItalic <> rtbRichTextBox.SelItalic Then
            If rtbRichTextBox.SelItalic = True Then
                strHTML$ = strHTML$ & "<i>"
            Else
                strHTML$ = strHTML$ & "</i>"
            End If

            blnItalic = rtbRichTextBox.SelItalic
        End If

        If blnStrikeThru <> rtbRichTextBox.SelStrikeThru Then
            If rtbRichTextBox.SelStrikeThru = True Then
                strHTML$ = strHTML$ & "<s>"
            Else
                strHTML$ = strHTML$ & "</s>"
            End If

            blnStrikeThru = rtbRichTextBox.SelStrikeThru
        End If

        If strLastFont$ <> rtbRichTextBox.SelFontName Then
            strLastFont$ = rtbRichTextBox.SelFontName
            strHTML$ = strHTML$ + "<font face=""" & strLastFont$ & """>"
        End If

        If lngLastFontColor& <> rtbRichTextBox.SelColor Then
            lngLastFontColor& = rtbRichTextBox.SelColor
            
            ''Get hexidecimal value of color
            strHex$ = Hex(rtbRichTextBox.SelColor)
            strHex$ = String$(6 - Len(strHex$), "0") & strHex$
            strHex$ = Right$(strHex$, 2) & mID$(strHex$, 3, 2) & Left$(strHex$, 2)
            
            strHTML$ = strHTML$ + "<font color=#" & strHex$ & ">"
        End If
 
        'strHTML$ = strHTML$ + rtbRichTextBox.SelText
        If Len(rtbRichTextBox.SelText) = 0 Then
            strHTML$ = strHTML$ + "<p align=left>"
        Else
            strHTML$ = strHTML$ + rtbRichTextBox.SelText
        End If

    Next lngCurText&

    RichToHTML = strHTML$

End Function

Sub HTMLToRich(strHTML As String, rtbRichTextBox As RichTextLib.RichTextBox)
    '**********************************************************
    '*            HTML To Rich by Joseph Huntley              *
    '*               joseph_huntley@email.com                 *
    '*                http://joseph.vr9.com                   *
    '**********************************************************
    '*   You may use this code freely as long as credit is    *
    '* given to the author, and the header remains intact.    *
    '**********************************************************

    '--------------------- The Arguments -----------------------
    'strHTML          - The html you want to convert.
    'rtbRichTextBox   - The Rich Textbox you want it edit.
    '-----------------------------------------------------------

    'Description: Takes a HTML string and then edit's a richtext
    '             control.

    Dim blnBold   As Boolean, blnUnderline As Boolean, blnStrikeThru As Boolean

    Dim blnItalic As Boolean, strLastFont As String, lngLastFontColor As Long, lngLastFontSize As Long

    Dim lngChar   As Long, strTag As String, lngSpot As Long, strChar As String

    Dim lngAlign  As Long, strBuf As String, strBuf2 As String, lngBuf As Long, strBuf3 As String

    Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2

    'set default values
    strLastFont$ = rtbRichTextBox.Font.Name
    lngLastFontColor& = -1

    'clear richtextbox
    rtbRichTextBox.Text = ""

    'Loop through string. If finds an HTML string
    For lngChar& = 1 To Len(strHTML$)
        strChar$ = mID$(strHTML$, lngChar&, 1)
      
        If strChar$ = "<" Then
            lngSpot& = InStr(lngChar& + 1, strHTML$, ">")

            If lngSpot& Then
              
                strTag$ = LCase$(mID$(strHTML$, lngChar& + 1, lngSpot& - lngChar& - 1))
                 
                If strTag$ = "b" Then
                    blnBold = True
                ElseIf strTag$ = "/b" Then
                    blnBold = False
                ElseIf strTag$ = "u" Then
                    blnUnderline = True
                ElseIf strTag$ = "/u" Then
                    blnUnderline = False
                ElseIf strTag$ = "i" Then
                    blnItalic = True
                ElseIf strTag$ = "/i" Then
                    blnItalic = False
                ElseIf strTag$ = "s" Then
                    blnStrikeThru = True
                ElseIf strTag$ = "/s" Then
                    blnStrikeThru = False
                ElseIf Left$(strTag$, 8) = "p align=" Then
                    strBuf$ = Right$(strTag$, Len(strTag$) - 8)
                    strBuf3$ = ""
                      
                    For lngBuf& = 1 To Len(strBuf$)
                        strBuf2$ = mID$(strBuf$, lngBuf&, 1)

                        If strBuf2$ <> """" Then strBuf3$ = strBuf3$ & strBuf2$
                    Next lngBuf&
                         
                    Select Case strBuf3$

                        Case "left":   lngAlign& = AlignLeft

                        Case "right":  lngAlign& = AlignRight

                        Case "center": lngAlign& = AlignCenter
                    End Select
                         
                ElseIf Left$(strTag$, 5) = "font " Then
                    strBuf$ = Right$(strTag$, Len(strTag$) - 5)
                         
                    Select Case Left$(strBuf$, InStr(strBuf$, "=") - 1)
                            
                        Case "color":
                            strBuf$ = Right$(strBuf$, Len(strBuf$) - InStr(strBuf$, "="))
                            strBuf3$ = ""

                            For lngBuf& = 1 To Len(strBuf$)
                                strBuf2$ = mID$(strBuf$, lngBuf&, 1)

                                If strBuf2$ <> """" And strBuf2$ <> "#" Then strBuf3$ = strBuf3$ & strBuf2$
                            Next lngBuf&

                            lngLastFontColor& = HexToDecimal(strBuf3$)
                            
                        Case "face":
                            strBuf$ = Right$(strBuf$, Len(strBuf$) - InStr(strBuf$, "="))
                            strBuf3$ = ""

                            For lngBuf& = 1 To Len(strBuf$)
                                strBuf2$ = mID$(strBuf$, lngBuf&, 1)

                                If strBuf2$ <> """" Then strBuf3$ = strBuf3$ & strBuf2$
                            Next lngBuf&

                            strLastFont$ = strBuf3$
                            
                        Case "size":
                            strBuf$ = Right$(strBuf$, Len(strBuf$) - InStr(strBuf$, "="))
                            strBuf3$ = ""

                            For lngBuf& = 1 To Len(strBuf$)
                                strBuf2$ = mID$(strBuf$, lngBuf&, 1)

                                If strBuf2$ <> """" Then strBuf3$ = strBuf3$ & strBuf2$
                            Next lngBuf&
                                   
                            Select Case strBuf3$

                                Case "1": lngLastFontSize& = 4

                                Case "2": lngLastFontSize& = 8

                                Case "3": lngLastFontSize& = 10

                                Case "4": lngLastFontSize& = 14

                                Case "5": lngLastFontSize& = 18

                                Case "6": lngLastFontSize& = 20

                                Case "7": lngLastFontSize& = 72
                            End Select
                                   
                    End Select

                End If
                   
                'skip over html tag
                lngChar& = lngSpot&
            End If 'for: If lngSpot& Then

        Else
            'set character with curretn artributes.
            rtbRichTextBox.SelStart = Len(rtbRichTextBox.Text)
            rtbRichTextBox.SelLength = 0
            rtbRichTextBox.SelText = strChar$
            rtbRichTextBox.SelStart = Len(rtbRichTextBox.Text) - 1
            rtbRichTextBox.SelLength = 1
            rtbRichTextBox.SelBold = blnBold
            rtbRichTextBox.SelUnderline = blnUnderline
            rtbRichTextBox.SelItalic = blnItalic
            rtbRichTextBox.SelStrikeThru = blnStrikeThru
            rtbRichTextBox.SelFontName = strLastFont$
            rtbRichTextBox.SelFontSize = lngLastFontSize&
            rtbRichTextBox.SelAlignment = lngAlign&
            rtbRichTextBox.SelColor = lngLastFontColor&
        End If 'for: If rtbRichTextBox.SelText = "<" Then
      
    Next lngChar&

End Sub

Function HexToDecimal(ByVal strHex As String) As Long

    'This function is required by the function 'HTMLToRich'

    'this function converts any hexidecimal color value
    '(e.g. "0000FF" = Blue) to decimal color value.

    Dim lngDecimal As Long, strCharHex As String, lngColor As Long

    Dim lngChar    As Long

    If Left$(strHex$, 1) = "#" Then strHex$ = Right$(strHex$, 6)
  
    strHex$ = Right$(strHex$, 2) & mID$(strHex$, 3, 2) & Left$(strHex$, 2)

    For lngChar& = Len(strHex$) To 1 Step -1
        strCharHex$ = mID$(UCase$(strHex$), lngChar&, 1)
    
        Select Case strCharHex$

            Case 0 To 9
                lngDecimal& = CLng(strCharHex$)

            Case Else 'A,B,C,D,E,F
                lngDecimal& = CLng(Chr$((Asc(strCharHex$) - 17))) + 10
        End Select
       
        lngColor& = lngColor& + lngDecimal& * 16 ^ (Len(strHex$) - lngChar&)
    Next lngChar&
  
    HexToDecimal = lngColor&

End Function

'Private Sub cmdConvertToHTML_Click()
'  txtHTML.Text = RichToHTML(rtbRichText, 0&, Len(rtbRichText.Text))
'End Sub
'
'Private Sub cmdConvertToRichText_Click()
'  Call HTMLToRich(txtHTML.Text, rtbRichText)
'End Sub
'
'Private Sub Form_Load()
'
'  'set the text in rtbRichTextBox
'
'  With rtbRichText
'     .Text = "Click on the 'convert' button to convert this richtext to HTML."
'     .SelStart = 0
'     .SelLength = Len(.Text)
'     .SelFontName = "Arial"
'     .SelFontSize = 10
'     .SelAlignment = rtfCenter
'     .SelStart = InStr(.Text, "convert") - 1
'     .SelLength = Len("convert")
'     .SelFontName = "Courier New"
'     .SelColor = vbBlue
'     .SelStart = InStr(.Text, "HTML") - 1
'     .SelLength = 4
'     .SelFontName = "Courier New"
'     .SelUnderline = True
'     .SelStart = .SelStart + 1
'     .SelLength = 1
'     .SelColor = vbRed
'     .SelStart = .SelStart + 1
'     .SelLength = 1
'     .SelColor = vbBlue
'     .SelStart = .SelStart + 1
'     .SelLength = 1
'     .SelColor = vbGreen
'     .SelStart = 0
'     .SelLength = 0
'  End With
'
'
'End Sub
'
