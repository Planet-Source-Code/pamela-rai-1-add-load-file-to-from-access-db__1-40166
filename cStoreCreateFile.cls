VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStoreCreateFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

    Option Explicit
    DefLng A-Z
    
    Const mclChunkSize As Long = 1024& * 32&    '32K Chunks

Public Function StoreFileIntoField(rs As Recordset, fldFileName As Field, fldLongBinary As Field, sFilename As String) As Boolean
    Dim iFN         As Integer
    Dim lFileLen    As Long
    Dim sChunk      As String
    
    With fldLongBinary
        If .Type = dbLongBinary Then 'only long binary supported
'            On Error GoTo StoreFileIntoFieldEH
            rs.Edit                 'edit the record in the recordset
            .Value = ""             'clear the field
            rs.Update               'update the record
            rs.Edit                 'now edit again
            
            iFN = FreeFile          'get a file handle
            Open sFilename For Binary As #iFN
            lFileLen = LOF(iFN)
            
            While lFileLen > 0      'loop until finished
                If mclChunkSize < lFileLen Then                 'chunk smaller than remaining
                    sChunk = String$(mclChunkSize, vbNullChar)
                Else                                            'chunk larger than remaining
                    sChunk = String$(lFileLen, vbNullChar)
                End If
                Get #iFN, , sChunk      'get a chunk
                .AppendChunk sChunk     'append to the field
                lFileLen = lFileLen - mclChunkSize              'handle the next chunk
            Wend
            
            Close #iFN                  'close the file
            rs.Update                   'store recordset onto database
            If StoreFileNameIntoField(rs, fldFileName, sFilename) Then
                StoreFileIntoField = True   'report success
            End If
        Else
            Err.Raise vbObjectError + gclErrFieldNotLongBinary, , "Field '" & .Name & "' is not Long Binary"
        End If
    End With
    Exit Function       'exit
'StoreFileIntoFieldEH:   'general error
'    Err.Raise Err.Number, "cStoreCreateFile, StoreFileIntoField", Err.Description
End Function

Public Function StoreFileNameIntoField(rs As Recordset, fldFileName As Field, sFilename As String) As Boolean
' Store the file name in fldFileName
'    On Error GoTo StoreFileNameIntoFieldEH
    With fldFileName
        If .Type = dbText Then
            With rs
                .Edit
                If sFilename = "" Then
                    fldFileName.Value = vbNull
                Else
                    fldFileName.Value = sFilename
                End If
                .Update
                StoreFileNameIntoField = True
            End With
        Else
            Err.Raise vbObjectError + gclErrFieldNotText, , "Field '" & .Name & "' is not Text"
        End If
    End With
    Exit Function
'StoreFileNameIntoFieldEH:
'    Err.Raise Err.Number, "cStoreCreateFile, StoreFileNameIntoField", Err.Description
End Function

Public Function CreateFileFromField(fld As Field, sFilename As String) As Boolean
    Dim iFN         As Integer
    Dim lFileLen    As Long
    Dim lOffset     As Long
    Dim sChunk      As String
    
    With fld
        If .Type = dbLongBinary Then            'only long binary supported
'            On Error GoTo CreateFileFromFieldEH 'general error
            If Dir$(sFilename) <> "" Then       'delete it if there already
                Kill sFilename
            End If
            iFN = FreeFile                      'get a file handle
            Open sFilename For Binary As #iFN   'open the file
            lFileLen = .FieldSize               'get the length of the field
            If lFileLen > 0 Then                '2000/09/28 Check for any data
                While lOffset < lFileLen            'loop getting chunks
                    sChunk = .GetChunk(lOffset, mclChunkSize)   'get chunk
                    Put #iFN, , sChunk              'write the chunk
                    lOffset = lOffset + mclChunkSize    'bump along the field
                Wend
                Close #iFN                          'close the newly created file
                CreateFileFromField = True          'report success
            Else
                Err.Raise vbObjectError + gclErrFieldNoLongBinaryData, , "Field '" & .Name & "' is not Long Binary"
            End If
        Else
            Err.Raise vbObjectError + gclErrFieldNotLongBinary, , "Field '" & .Name & "' is not Long Binary"
        End If
    End With
'    Exit Function           'exit
'CreateFileFromFieldEH:      'general error
'    Err.Raise Err.Number, "cStoreCreateFile, CreateFileFromField", Err.Description
End Function













