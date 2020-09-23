Attribute VB_Name = "modStatusV2"
' modStatusV2
' 2000/04/13 Copyright © 2000, Larry Rebich
' 2000/04/13 larry@buygold.net, www.buygold.net, 760.771.4730
' 2000/10/01 Used in BrandingModel
' 2000/10/01 Used in Branding

    Option Explicit
    DefLng A-Z

Public Function StatusOff(frm As Form) As Boolean
    Dim objStatus As StatusBar
    
    Set objStatus = frm.StatusBar1
    With objStatus
        If .Panels("status").Text <> "" Then
            .Panels("status").Text = ""
            StatusOff = True
        End If
    End With
    frm.tmrStatus.Interval = 0
End Function

Public Function Status(frm As Form, sMsg As String, Optional vntCritical As Variant, Optional vntPersistent As Variant) As Boolean
    Dim objStatus As StatusBar
    Dim bCritical As Boolean
    Dim bPersistent As Boolean
    
    If Not IsMissing(vntCritical) Then
        bCritical = vntCritical
    End If
    If Not IsMissing(vntPersistent) Then
        bPersistent = vntPersistent
    End If
    
    Set objStatus = frm.StatusBar1
    With objStatus
        If .Panels("status").Text <> " " & sMsg Then
            .Panels("status").Text = " " & sMsg
            Status = True
        End If
        If bCritical Then
            .Font.Bold = True
            Beep
        Else
            .Font.Bold = False
        End If
    End With
    If Not bPersistent Then
        frm.tmrStatus.Interval = 2000
    End If
End Function

