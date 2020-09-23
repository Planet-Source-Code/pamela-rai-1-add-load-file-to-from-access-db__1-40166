Attribute VB_Name = "modNotTooSmallBig"

    Option Explicit
    DefLng A-Z

Public Function NotTooSmall(lValue As Long, lMin As Long) As Long
' 1998/01/21 Can't Be Too Small
    If lValue < lMin Then
        NotTooSmall = lMin
    Else
        NotTooSmall = lValue
    End If
End Function

Public Function NotTooBig(lValue As Long, lMax As Long) As Long
' 1999/11/09 Added and used in picSplitMouseUp on the MDI form
    If lValue > lMax Then
        NotTooBig = lMax
    Else
        NotTooBig = lValue
    End If
End Function

