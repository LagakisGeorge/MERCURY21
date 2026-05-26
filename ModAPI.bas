Attribute VB_Name = "ModAPI"

Option Explicit

'===============================
'API Declarations and Constant
'===============================

'For tracking mouse cursor position
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
            
Public Type POINTAPI

    X As Long
    Y As Long

End Type

Public BImg() As Byte

'For memory status
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
                
Private Type MEMORYSTATUS 'Type variable for memory info

    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long

End Type

Private Type BITMAP '14 bytes

    bmType                      As Long
    bmWidth                     As Long
    bmHeight                    As Long
    bmWidthBytes                As Long
    bmPlanes                    As Integer
    bmBitsPixel                 As Integer
    bmBits                      As Long

End Type

'Win32 API Function Declarations
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetObject _
                Lib "gdi32" _
                Alias "GetObjectA" (ByVal hObject As Long, _
                                    ByVal nCount As Long, _
                                    lpObject As Any) As Long

Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long

Private Declare Function SetStretchBltMode _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nStretchMode As Long) As Long

Private Declare Function StretchBlt _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal nSrcWidth As Long, _
                             ByVal nSrcHeight As Long, _
                             ByVal dwRop As Long) As Long

'Win32 API Constant Declarations
Private Const STRETCH_HALFTONE = 4

'Use to get the window top,left,right and buttom position
Public Declare Function GetWindowRect _
               Lib "user32" (ByVal hWnd As Long, _
                             lpRect As Rect) As Long
                
Public Type Rect 'Type variable for window rect

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type

Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hWnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long
                
Public Const WM_CLOSE = &H10 'Message use to perform close

Public Const WM_ACTIVATE = &H6 'Message use to perform activate

'Use to set the parent
Public Declare Function SetParent _
               Lib "user32" (ByVal hWndChild As Long, _
                             ByVal hWndNewParent As Long) As Long

'Use for setting windows on top
Public Declare Function SetWindowPos _
               Lib "user32" (ByVal hWnd As Long, _
                             ByVal hWndInsertAfter As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal cx As Long, _
                             ByVal cy As Long, _
                             ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1

Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOSIZE = &H1

Public Const SWP_NOMOVE = &H2

Public Const SWP_NOACTIVATE = &H10

Public Const SWP_SHOWWINDOW = &H40

'API for opening a browser
Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

Public MEM_STAT As MEMORYSTATUS

'API used to change the form border
Public Declare Function GetWindowLong _
               Lib "user32" _
               Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                       ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong _
               Lib "user32" _
               Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long

'Related contstant (see API used to change the form border)
Public Const WS_CAPTION = &HC00000

Public Const WS_THICKFRAME = &H40000

Public Const WS_DLGFRAME = &H400000

Public Const GWL_STYLE = (-16)

Public Const GWL_EXSTYLE = (-20)

Public Const SWP_FRAMECHANGED = &H20

Public Const SWP_NOZORDER = &H4

Public Const SWPFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE

'Function used to change the form border (see API used to change the form border)
'Public Function FormStyle(ByRef srcHWND As Long, ByVal Bit As Long, ByVal Value As Boolean) As Boolean
'    Dim nStyle As Long
'
'    nStyle = GetWindowLong(srcHWND, GWL_STYLE)
'
'    If Value Then
'        nStyle = nStyle Or Bit
'    Else
'        nStyle = nStyle And Not Bit
'    End If
'    Call SetWindowLong(srcHWND, GWL_STYLE, nStyle)
'    SetWindowPos srcHWND, 0, 0, 0, 0, 0, SWPFlags
'
'    FormStyle = (nStyle = GetWindowLong(srcHWND, GWL_STYLE))
'End Function

'Public Sub OpenURL(urlADD As String, sourceHWND As Long)
'     Call ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
'End Sub

'Public Function OnTop(ByRef hWnd As Long, Optional BringOnTop As Boolean = True) As Long
'    If BringOnTop = True Then
'        SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'    Else
'        SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'    End If
'End Function

' http://www.codenewsgroups.net/group/microsoft.public.vb.general.discussion/topic3286.aspx
' http://edais.mvps.org/

'---------------------------------Start Code------------------

'Public Function StretchSourcePictureFromFile(FILENAME As String, picDest As PictureBox) As StdPicture
'    Dim hMemDC      As Long
'    Dim hOldBmp     As Long
'    Dim hMemWdth    As Long
'    Dim hMemHght    As Long
'    Dim Bmp         As BITMAP
'    Dim nRetVal     As Long
'    Dim picSrc      As StdPicture
'    Dim OldSM       As ScaleModeConstants
'    Dim OldAR       As Boolean
'    Dim ScaleFactor As Double
'    Dim ShowLeft    As Long
'    Dim ShowTop     As Long
'    Dim ShowWidth   As Long
'    Dim ShowHeight  As Long
'
'    If Len(FILENAME) = 0 Then
'        Beep
'        Exit Function
'    End If
'
'    'Create the memory DC
'    hMemDC = CreateCompatibleDC(GetDC(GetDesktopWindow()))
'    'Load the picture
'    Set picSrc = LoadPicture(FILENAME)
'    'Assign the picture to the memory DC
'    hOldBmp = SelectObject(hMemDC, picSrc.Handle)
'
'    'Get the sizes of the picture
'    nRetVal = GetObject(picSrc.Handle, Len(Bmp), Bmp)
'    hMemWdth = Bmp.bmWidth
'    hMemHght = Bmp.bmHeight
'
'    'Make sure there is a picture
'    If (hMemWdth > 0) And (hMemHght > 0) Then
'
'        'Stretch the picture to the picturebox
'        With picDest
'            'Save the PictureBox's ScaleMode and set it to vbPixels
'            OldSM = .ScaleMode
'            .ScaleMode = vbPixels
'
'            'Get the largest possible scaling factor
'            ScaleFactor = Biggest(hMemWdth / .ScaleWidth, hMemHght / _
'                    .ScaleHeight)
'
'            'Get the positions and sizes for the destination picture
'            ShowWidth = hMemWdth / ScaleFactor
'            ShowHeight = hMemHght / ScaleFactor
'            ShowLeft = (.ScaleWidth - ShowWidth) / 2
'            ShowTop = (.ScaleHeight - ShowHeight) / 2
'
'            'Save the PictureBox's AutoRedraw and set it to True
'            OldAR = .AutoRedraw
'            .AutoRedraw = True
'
'            .Picture = LoadPicture()
'            .Cls
'            nRetVal = SetStretchBltMode(.hDC, STRETCH_HALFTONE)
'            nRetVal = StretchBlt(.hDC, ShowLeft, ShowTop, ShowWidth, _
'                    ShowHeight, hMemDC, 0, 0, hMemWdth, hMemHght, _
'                    vbSrcCopy)
'
'            If (nRetVal = 0) Then
'                Debug.Print "StretchBlt() Error Code " & _
'                        Err.LastDllError
'            End If
'
'            .Refresh
'        End With
'
'        'Reset the PictureBox's ScaleMode
'        picDest.ScaleMode = OldSM
'
'        'Reset the PictureBox's AutoRedraw
'        picDest.AutoRedraw = OldAR
'    End If
'
'    'Return the picture object
'    Set StretchSourcePictureFromFile = picSrc
'
'    'Clean up the used memory
'    Call SelectObject(hMemDC, hOldBmp)
'    Call DeleteDC(hMemDC)
'    Set picSrc = Nothing
'End Function

Public Sub StretchSourcePictureFromPicture(picSrc As StdPicture, picDest As PictureBox)

        '<EhHeader>
        On Error GoTo StretchSourcePictureFromPicture_Err

        '</EhHeader>
        Dim hMemDC      As Long

        Dim hOldBmp     As Long

        Dim hMemWdth    As Long

        Dim hMemHght    As Long

        Dim Bmp         As BITMAP

        Dim nRetVal     As Long

        Dim OldSM       As ScaleModeConstants

        Dim OldAR       As Boolean

        Dim ScaleFactor As Double

        Dim ShowLeft    As Long

        Dim ShowTop     As Long

        Dim ShowWidth   As Long

        Dim ShowHeight  As Long

        'Make sure we have a valid picture
100     If picSrc.Handle = 0 Then
110         Beep

            Exit Sub

        End If

        'Create the memory DC
120     hMemDC = CreateCompatibleDC(GetDC(GetDesktopWindow()))
        'Assign the picture to the memory DC
130     hOldBmp = SelectObject(hMemDC, picSrc.Handle)

        'Get the sizes of the picture
140     nRetVal = GetObject(picSrc.Handle, Len(Bmp), Bmp)
150     hMemWdth = Bmp.bmWidth
160     hMemHght = Bmp.bmHeight

        'Make sure there is a picture
170     If (hMemWdth > 0) And (hMemHght > 0) Then

            'Stretch the picture to the picturebox
180         With picDest
                'Save the PictureBox's ScaleMode and set it to vbPixels
190             OldSM = .ScaleMode
200             .ScaleMode = vbPixels

                'Get the largest possible scaling factor
210             ScaleFactor = Biggest(hMemWdth / .ScaleWidth, hMemHght / .ScaleHeight)

                'Get the positions and sizes for the destination picture
220             ShowWidth = hMemWdth / ScaleFactor
230             ShowHeight = hMemHght / ScaleFactor
240             ShowLeft = (.ScaleWidth - ShowWidth) / 2
250             ShowTop = (.ScaleHeight - ShowHeight) / 2

                'Save the PictureBox's AutoRedraw and set it to True
260             OldAR = .AutoRedraw
270             .AutoRedraw = True

280             .Cls
290             nRetVal = SetStretchBltMode(.hDC, STRETCH_HALFTONE)
300             nRetVal = StretchBlt(.hDC, ShowLeft, ShowTop, ShowWidth, ShowHeight, hMemDC, 0, 0, hMemWdth, hMemHght, vbSrcCopy)

310             If (nRetVal = 0) Then
320                 Debug.Print "StretchBlt() Error Code " & Err.LastDllError
                End If

330             .Refresh
            End With

            'Reset the PictureBox's ScaleMode
340         picDest.ScaleMode = OldSM

            'Reset the PictureBox's AutoRedraw
350         picDest.AutoRedraw = OldAR
        End If

        'Clean up the used memory
360     Call SelectObject(hMemDC, hOldBmp)
370     Call DeleteDC(hMemDC)

        '<EhFooter>
        Exit Sub

StretchSourcePictureFromPicture_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.ModAPI.StretchSourcePictureFromPicture " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.ModAPI.StretchSourcePictureFromPicture " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Function Biggest(Val1 As Double, Val2 As Double) As Double

        '<EhHeader>
        On Error GoTo Biggest_Err

        '</EhHeader>

100     Biggest = IIf(Val1 >= Val2, Val1, Val2)

        '<EhFooter>
        Exit Function

Biggest_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.ModAPI.Biggest " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.ModAPI.Biggest " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'-------------------------------End Code-----------------

'Function mADD_FIELD(ByVal TABLE, ByVal FIELD, ByVal FIELDTYPE)
''-----------------------------------------------------------------------------------
'    Dim R As New ADODB.Recordset
'    Dim sql
'    sql = "SELECT TOP 1 * FROM " + TABLE + ";"
'    MDIForm1.Caption = TABLE + FIELD
'    R.Open sql, CN, adOpenForwardOnly, adLockReadOnly
'    Dim k As Long
'    Dim OK
'
'    OK = 0
'    For k = 0 To R.FIELDS.Count - 1
'        If R.FIELDS(k).Name = FIELD Then
'            OK = 1
'        End If
'    Next
'    ADD_FIELD = 0
'    If OK = 0 Then
'        sql = "alter table " + TABLE + " ADD " + FIELD + " " + FIELDTYPE
'        ADD_FIELD = 1
'        'On Error GoTo 0
'        CN.Execute sql
'    End If
'
'End Function

