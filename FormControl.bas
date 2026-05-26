Attribute VB_Name = "FormControl"

Private List()   As Control

Private curr_obj As Object

Private iHeight  As Integer

Private iWidth   As Integer

Private x_size   As Double

Private y_size   As Double
'*****************************************************************************************
'                           LICENSE INFORMATION
'*****************************************************************************************
'   FormControl Version 2.0
'   Code module for resizing a form based on screen size, then resizing the
'   controls based on the forms size
'http://www.dreamincode.net/forums/topic/34776-resize-form-controls-for-screen-size-vb6/
'   Copyright (C) 2007
'   Richard L. McCutchen
'   Email: richard@psychocoder.net
'   Created: AUG99
'
'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>.
'*****************************************************************************************

'Private Type Control
'    Index As Integer
'    Name As String
'    Left As Integer
'    Top As Integer
'    width As Integer
'    height As Integer
'End Type

'Private fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Public Sub mForm_Load(frm As Form, ByRef fh, ByRef fw, ByRef ft, ByRef fl)

        '<EhHeader>
        On Error GoTo mForm_Load_Err

        '</EhHeader>
        Dim k

        'ffproto = 1
'100     ffw = frm.Width
'110     ffh = frm.Height

        On Error Resume Next

120     For k = 0 To frm.Controls.Count - 1
130         fw(k) = frm.Controls(k).Width
140         fh(k) = frm.Controls(k).Height
           If Right(frm.Controls(k).Name, 5) = "parad" Then
               fl(k) = frm.Controls(k).Left
               ft(k) = frm.Controls(k).Top
           End If
           
           
150         fl(k) = frm.Controls(k).Left
160         ft(k) = frm.Controls(k).Top
        Next

        '<EhFooter>
        Exit Sub

mForm_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.FormControl.mForm_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.FormControl.mForm_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub mForm_Resize(frm As Form, _
                        fontHeight As Single, _
                        WW As Long, _
                        HH As Long, _
                        fh, _
                        fw, _
                        ft, _
                        fl)

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
    Dim k As Long

    Dim nW, nH

    nW = frm.Width
    nH = frm.Height

    On Error Resume Next
    Dim F_fontheight As Long
     F_fontheight = Val(FINDPARAMETROI(1, "MDIFORM1", "F_fontheight", "10", "лецисто уьос FONT(цяаллатосеияас) лета тгм лецистопоигсг(10-14)"))
     If F_fontheight > 16 Or F_fontheight < 8 Then
         F_fontheight = 10
     End If
    'If ffproto < 2 Then
    '   ffproto = ffproto + 1
    'If Me.width > ffW Then
    'Else
    '    ffW = Me.width
    '     ffH = Me.height
    'End If
    'Else
    For k = 0 To frm.Controls.Count - 1
      If Right(frm.Controls(k).Name, 5) = "parad" Then
         frm.Controls(k).Width = fw(k) * (nW / WW)
      Else
      End If
      
        frm.Controls(k).Width = fw(k) * (nW / WW)
        ' On Error Resume Next
        frm.Controls(k).Height = fh(k) * (nH / HH)
        frm.Controls(k).Top = ft(k) * (nH / HH)    ' ORIG10830)
        
        
        frm.Controls(k).Left = fl(k) * (nW / WW)    '13860)
        If frm.Controls(k).Left < 0 Then
            frm.Controls(k).Left = fl(k)
        End If

        x_size = nW / WW

        '   x_size = ((nH / HH) + (nW / WW)) / 2
        If x_size > 0 Then
            'On Error GoTo 0
            If x_size > 1.5 Then
               frm.Controls(k).Font.Size = F_fontheight ' 10 ' SetFontSize(fontHeight)
            ElseIf x_size < 1 Then
                frm.Controls(k).Font.Size = 8 ' SetFontSize(fontHeight)
            End If
            'On Error Resume Next
            
        End If

        'frm.Controls(k).Name
    Next

    'End If

End Sub

Public Sub ResizeControls(frm As Form)
    Dim i As Integer
    '   Get ratio of initial form size to current form size
    x_size = frm.Height / iHeight
    y_size = frm.Width / iWidth
    On Error Resume Next
    'Loop though all the objects on the form
    'Based on the upper bound of the # of controls
    For i = 0 To UBound(List)
        'Grad each control individually
        For Each curr_obj In frm
            'Check to make sure its the right control
            If curr_obj.TabIndex = List(i).index Then
                'Then resize the control
                With curr_obj
                    .Left = List(i).Left * y_size
                    .Width = List(i).Width * y_size
                    .Height = List(i).Height * x_size
                    .Top = List(i).Top * x_size
                   ' .FontSize = SetFontSize()
                End With
            End If
            'Get the next control
        Next curr_obj
    Next i
End Sub

Public Function SetFontSize(ByVal fontHeight As Single) As Integer

        'Make sure x_size is greater than 0
        '<EhHeader>
        On Error GoTo SetFontSize_Err

        '</EhHeader>
100     If Int(x_size) > 0 Then
            'Set the font size
110         SetFontSize = Int(x_size * fontHeight) '8   '12
        End If

        '<EhFooter>
        Exit Function

SetFontSize_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.FormControl.SetFontSize " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.FormControl.SetFontSize " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Public Sub GetLocation(frm As Form)
'    Dim I As Integer
'    '   Load the current positions of each object into a user defined type array.
'    '   This information will be used to rescale them in the Resize function.
'    'On Error Resume Next
'    'Loop through each control
'    For Each curr_obj In frm
'        'Resize the Array by 1, and preserve
'        'the original objects in the array
'        ReDim Preserve List(I)
'        With List(I)
'            '        .Name = curr_obj
'            .Index = curr_obj.TabIndex
'            .Left = curr_obj.Left
'            .Top = curr_obj.Top
'            .width = curr_obj.width
'            .height = curr_obj.height
'        End With
'        I = I + 1
'    Next curr_obj
'
'    '   This is what the object sizes will be compared to on rescaling.
'    iHeight = frm.height
'    iWidth = frm.width
'End Sub

'Public Sub CenterForm(frm As Form)
'    frm.Move (Screen.width - frm.width) \ 2, (Screen.height - frm.height) \ 2
'End Sub

Public Sub ResizeForm(frm As Form)

On Error Resume Next

'Set the forms height
    frm.Height = Screen.Height / 2
    'Set the forms width
    frm.Width = Screen.Width / 2
    'Resize all of the controls
    'based on the forms new size
    ResizeControls frm
End Sub

