Attribute VB_Name = "modPath"

    Option Explicit
    DefLng A-Z


Public Function GetNameFromPathAndName(ByVal sThePathAndName As String) As String
' Return a string containing the file name from a fully qualified file name. Larry
' If no path then return the file name anyhow
    Dim sTemp1 As String                        'temporary string
    Dim sTemp2 As String
    sTemp1 = Trim$(sThePathAndName)             'save it
    sTemp2 = GetPathFromPathAndName(sTemp1)     'get path
    If InStr(sThePathAndName, sTemp2) > 0 Then
        sTemp1 = Mid$(sThePathAndName, Len(sTemp2) + 1)
    End If
    GetNameFromPathAndName = sTemp1             'now contains just file name
End Function

Public Function GetPathFromPathAndName(ByVal sThePathAndName As String) As String
' Return a string containing the file's path from a fully qualified file name. Larry
' Return "" if no path
    Dim i As Integer                    'used in for/next loops
    Dim sTemp As String
    sTemp = Trim$(sThePathAndName)      'trim it
    If InStr(sTemp, "\") = 0 Then       'any backslash?
        Exit Function
    End If
    For i = Len(sTemp) To 1 Step -1     'find the right most one
        If Mid$(sTemp, i, 1) = "\" Then
            GetPathFromPathAndName = Mid$(sTemp, 1, i) 'now have just path
            Exit Function
        End If
    Next
End Function

Public Function GetExtensionFromFileName(sTheFileAndExt As String, sTheFile As String) As String
' 95/09/15 Return the file's extension from the filename and extension. Larry
' Return just the file name in sTheFile
    Const csFrame = "+"                         'frame character
    Const csPeriod = "."                        'ext follows this
    Dim i As Integer
    Dim iLen As Integer
    Dim iLoc As Integer                         'location of the period
    Dim sTemp As String
    Dim sFil As String                          'file
    Dim sExt As String                          'extension
    
    sTemp = csFrame & Trim$(sTheFileAndExt) & csFrame 'work with it here
    If InStr(sTheFileAndExt, ".") = 0 Then      'none, return what we found
        sTheFile = sTheFileAndExt
        Exit Function
    End If
    iLen = Len(sTemp)
    For i = iLen To 1 Step -1                   'find the period
        If Mid$(sTemp, i, 1) = csPeriod Then
            iLoc = i                            'got it
            sExt = Mid$(sTemp, i + 1)           'got the extension
'            If Len(sExt) > 1 And Len(sExt) < 5 Then 'OK, good
            If Len(sExt) > 3 Then               '2000/02/01 OK, good
                sExt = Mid$(sExt, 1, Len(sExt) - 1) 'drop Frame character
            Else
                GoTo GetExtensionFromFileNameExit   'bad
            End If
            sFil = Mid$(sTemp, 1, i - 1)
            If Len(sFil) > 1 Then
                sFil = Mid$(sTemp, 2, i - 2)        'drop Frame character
            Else
                GoTo GetExtensionFromFileNameExit   'bad
            End If
            Exit For
        End If
    Next
    sTheFile = sFil                 'return what we found
    GetExtensionFromFileName = sExt
    Exit Function                   'bye
    
GetExtensionFromFileNameExit:
    sTheFile = ""                   'return blank
    GetExtensionFromFileName = ""   'not good
End Function

