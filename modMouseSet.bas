Attribute VB_Name = "modMouseSet"

    Option Explicit
    DefLng A-Z
    ' Save Mouse Here
    Dim sMouseStack     As String

Private Function MouseFunction(Optional vntNewMouse As Variant) As Integer
' 95/11/24 Rewrite to use optional parameter. Larry.
    Const ciMouseMin As Integer = vbDefault   'default
    Const ciMouseMax As Integer = vbHourglass 'hourglass
    If Not IsMissing(vntNewMouse) Then
        If vntNewMouse >= ciMouseMin And vntNewMouse <= ciMouseMax Then
            sMouseStack = sMouseStack & Chr$(Screen.MousePointer)   'current mouse
            MouseFunction = vntNewMouse    'return it
        End If
    Else
        If sMouseStack <> "" Then       'Any? Zero is returned if there is none
            MouseFunction = Asc(Right$(sMouseStack, 1))  'return it
            sMouseStack = Left$(sMouseStack, Len(sMouseStack) - 1)  'truncate
        End If
    End If
End Function

Public Function MouseReset()
' 95/11/24 Another name for MouseRestore, more logical. Larry.
    MouseReset = MouseFunction()
End Function

Public Function MouseRestore()
    MouseRestore = MouseFunction()          'restore it
End Function

Public Function MouseSet(Optional NewMouse As Variant) As Integer
' 95/11/24 Can use MouseSet() to reset the mouse. Larry.
    If IsMissing(NewMouse) Then
        MouseSet = MouseFunction()
    Else
        MouseSet = MouseFunction(NewMouse)  'return the value found
    End If
End Function

