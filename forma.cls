VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "forma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private List()   As Control

Private curr_obj As Object

Private iHeight  As Integer

Private iWidth   As Integer

Private x_size   As Double

Private y_size   As Double

Private fh()     As Long

Private fw()     As Long

Private ft()     As Long

Private fl()     As Long

Dim ffw, ffh

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

Private Type Control

    Index As Integer
    Name As String
    Left As Integer
    Top As Integer
    Width As Integer
    Height As Integer

End Type

'Private fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Public Sub mForm_Load(frm As Form)

        '<EhHeader>
        On Error GoTo mForm_Load_Err

        '</EhHeader>

100     nNumOfControls = frm.Controls.Count - 1
        'bFirstTime = True
102     ReDim fw(nNumOfControls)
104     ReDim fh(nNumOfControls)
106     ReDim fl(nNumOfControls)
108     ReDim ft(nNumOfControls)

        Dim k

110     ffh = frm.Height '* 1.4
112     ffw = frm.Width '* 1

        'ffW = frm.Width
        'ffH = frm.Height
        On Error Resume Next

114     For k = 0 To frm.Controls.Count - 1
116         fw(k) = frm.Controls(k).Width
118         fh(k) = frm.Controls(k).Height
120         fl(k) = frm.Controls(k).Left
122         ft(k) = frm.Controls(k).Top
        Next

        '<EhFooter>
        Exit Sub

mForm_Load_Err:
        Err.Raise vbObjectError + 100, "ADOMERCNEW.forma.mForm_Load", "forma component failure"

        '</EhFooter>

End Sub

Public Sub mForm_Resize(frm As Form, WW As Long, HH As Long)  ', fh, fw, ft, fl)

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
    Dim k As Long

    Dim nW, nH

    nW = frm.Width
    nH = frm.Height

    On Error Resume Next
  
    For k = 0 To frm.Controls.Count - 1
        frm.Controls(k).Width = fw(k) * (nW / WW)
        ' On Error Resume Next
        frm.Controls(k).Height = fh(k) * (nH / HH)
        frm.Controls(k).Top = ft(k) * (nH / HH)    ' ORIG10830)
        frm.Controls(k).Left = fl(k) * (nW / WW)    '13860)

        x_size = nW / WW

        If x_size > 0 Then
            'On Error GoTo 0
            frm.Controls(k).Font.Size = SetFontSize()
            'On Error Resume Next
            
        End If

        frm.Controls(k).Name
    Next

    'End If

End Sub

'Private Sub ResizeControls(frm As Form)
'    Dim I As Integer
'    '   Get ratio of initial form size to current form size
'    x_size = frm.height / iHeight
'    y_size = frm.width / iWidth
'    On Error Resume Next
'    'Loop though all the objects on the form
'    'Based on the upper bound of the # of controls
'    For I = 0 To UBound(List)
'        'Grad each control individually
'        For Each curr_obj In frm
'            'Check to make sure its the right control
'            If curr_obj.TabIndex = List(I).Index Then
'                'Then resize the control
'                With curr_obj
'                    .Left = List(I).Left * y_size
'                    .width = List(I).width * y_size
'                    .height = List(I).height * x_size
'                    .Top = List(I).Top * x_size
'                    .FONTSIZE = SetFontSize()
'                End With
'            End If
'            'Get the next control
'        Next curr_obj
'    Next I
'End Sub

Private Function SetFontSize() As Integer

        'Make sure x_size is greater than 0
        '<EhHeader>
        On Error GoTo SetFontSize_Err

        '</EhHeader>
100     If Int(x_size) > 0 Then
            'Set the font size
102         SetFontSize = Int(x_size * 8)
        End If

        '<EhFooter>
        Exit Function

SetFontSize_Err:
        Err.Raise vbObjectError + 100, "ADOMERCNEW.forma.SetFontSize", "forma component failure"

        '</EhFooter>

End Function

'Private Sub GetLocation(frm As Form)
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

'Private Sub CenterForm(frm As Form)
'    frm.Move (Screen.width - frm.width) \ 2, (Screen.height - frm.height) \ 2
'End Sub

'Private Sub ResizeForm(frm As Form)
''Set the forms height
'    frm.height = Screen.height / 2
'    'Set the forms width
'    frm.width = Screen.width / 2
'    'Resize all of the controls
'    'based on the forms new size
'    ResizeControls frm
'End Sub

