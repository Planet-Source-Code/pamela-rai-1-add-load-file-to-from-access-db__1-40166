Attribute VB_Name = "modAddBackSlash"

    Option Explicit
    DefLng A-Z
    
    Const mcsBkSlash = "\"
    
Public Function AddBackslash(sThePath As String) As String
' Add a backslash to a path if needed
' sPath contains the path
' Return a path with a backslash
    If Right$(sThePath, 1) <> mcsBkSlash Then
        sThePath = sThePath + mcsBkSlash
    End If
    AddBackslash = sThePath
End Function


