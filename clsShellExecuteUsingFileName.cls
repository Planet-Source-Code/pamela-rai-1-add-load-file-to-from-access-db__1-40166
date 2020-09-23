VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShellEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

    Option Explicit
    DefLng A-Z

    Private mlHwnd          As Long     'owner's handle
    Private msFilename      As String   'file to launch
    Private msErrorMessage  As String   'error message, if any error
    Private mlRtn           As Long
    
    Private Const SE_ERR_NOASSOC = 31   'get this error when no application associated with the file
    Private Const SW_SHOWNORMAL As Long = 1
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    '

Public Property Let lRtn(lRtn As Long)
    mlRtn = lRtn
End Property
Public Property Get lRtn() As Long
    lRtn = mlRtn
End Property

Public Property Let lHwnd(lHwnd As Long)
    mlHwnd = lHwnd
End Property
Public Property Get lHwnd() As Long
    lHwnd = mlHwnd
End Property

Public Property Let sFileName(sFileName As String)
    msFilename = sFileName
End Property
Public Property Get sFileName() As String
    sFileName = msFilename
End Property

Public Property Let sErrorMessage(sErrorMessage As String)
    msErrorMessage = sErrorMessage
End Property
Public Property Get sErrorMessage() As String
    sErrorMessage = msErrorMessage
End Property

Public Function ShellExecuteUsingFileName() As Boolean
    mlRtn = ShellExecute(mlHwnd, "Open", msFilename, 0, 0, SW_SHOWNORMAL)    'launch it
    If mlRtn > 32 Then                       'OK
        ShellExecuteUsingFileName = True
    ElseIf mlRtn = SE_ERR_NOASSOC Then      'no association
        Me.OpenWith                         'so use OpenAs_RunDLL
        ShellExecuteUsingFileName = True    'report true, Open With is running.
    Else
        Me.sErrorMessage = "Can't launch '" & Me.sFileName & "'."
    End If
End Function

Public Function OpenWith() As Long
    ' Show the 'open with' dialog
    OpenWith = ShellExecute(mlHwnd, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & msFilename, "", vbNormalFocus)
End Function


