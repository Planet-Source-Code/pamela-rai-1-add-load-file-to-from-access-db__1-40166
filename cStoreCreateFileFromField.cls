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
' cStoreCreateFile
' 2000/08/29 Copyright © 2000, Larry Rebich, using the DELL7500, while in Pacifica, California
' 2000/08/29 larry@buygold.net, www.buygold.net, 760.771.4730

    Option Explicit
    DefLng A-Z

Public Function StoreFileIntoField(rs As Recordset, fld As Field, sFilename As String) As Boolean
    With fld
        If .Type = dbLongBinary Then
        Else
            Err.Raise vbObjectError + gclErrFieldNotLongBinary, "cSaveLoadAccessPicture, SaveImageIntoField", "Field is not Long Binary"
        End If
    End With
End Function


Public Function CreateFileFromField(fld As Field, sFilename As String) As Boolean

End Function
