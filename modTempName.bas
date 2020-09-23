Attribute VB_Name = "modTempName"

    Option Explicit
    DefLng A-Z

' Used to get a temporary file name
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" _
    (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'

Public Function TempNamePlease(sTheExtension As String, Optional sPath As Variant, _
    Optional sPrefix As Variant) As String
' Get a temporary file name, concept from MSDN CD7
' 1994/10/31 If sTheExtension provided then it replaces ".tmp" as the extenstion
' 1995/07/21 Modify for Win32 API's
' 1995/10/02 Add sPrefix to cater to building temporary datasets
    Dim sTemp As String         'temporary string
    Dim sTempPath As String     'temporary path
    Dim iBufSize As Integer     'buffer size
    Dim iRtn As Long            'return value
    Dim iFileHandle             'file handle
    Const csTmpExt = ".tmp"     'leave lower case please
    If IsMissing(sPrefix) Then
        sPrefix = "tmp"
    End If
    iBufSize = 256              'and a little extra, make sure have some spaces
    sTemp = String$(iBufSize, Chr$(0))          'load file name buffer
    sTempPath = String$(iBufSize, Chr$(0))      'load temp path buffer
    If IsMissing(sPath) Then
        iRtn = GetTempPath(iBufSize, sTempPath)     'get a temporary path
        If iRtn > 0 Then
            sTempPath = Left$(sTempPath, iRtn)      'get the path
        Else
            sTempPath = App.Path
            If Right$(sTempPath, 1) <> "\" Then
                sTempPath = sTempPath & "\"
            End If
        End If
    Else
        sTempPath = sPath                       'use the one provided by requestor
    End If
    iRtn = GetTempFileName(sTempPath, sPrefix, 0, sTemp) 'API to get the name
    If iRtn > 0 Then
        sTemp = Mid$(sTemp, 1, InStr(sTemp, Chr$(0)) - 1)
    End If
    If InStr(LCase$(sTemp), LCase$(csTmpExt)) Then  'end in temp?
        If sTheExtension <> "" Then             'replacement specified?
            Kill sTemp                          'yes, kill the one Windows created
            iBufSize = InStr(LCase$(sTemp), LCase$(csTmpExt))       'add the one passed in to this function
            sTemp = Mid$(sTemp, 1, iBufSize - 1) & sTheExtension    'add it to the name
            iFileHandle = FreeFile              'get a file handle
            Open sTemp For Output As #iFileHandle   'make empty file
            Close #iFileHandle                  'close the file
        End If
    End If
    TempNamePlease = sTemp     'return the name
End Function

Public Function TempPathPlease() As String
' 1998/02/28 Return the temporary path name
    Dim sTempPath As String     'temporary path
    Dim iBufSize As Integer     'buffer size
    Dim iRtn As Long            'return value
    
    iBufSize = 256              'and a little extra, make sure have some spaces
    sTempPath = String$(iBufSize, Chr$(0))      'load temp path buffer
    iRtn = GetTempPath(iBufSize, sTempPath)     'get a temporary path
    If iRtn > 0 Then
        sTempPath = Left$(sTempPath, iRtn)      'get the path
    Else
        sTempPath = ""
    End If
    TempPathPlease = sTempPath
End Function


