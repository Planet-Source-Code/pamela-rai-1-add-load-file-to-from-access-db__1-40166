VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStoreCreateFileDemo 
   Caption         =   "frmStoreCreateFileDemo"
   ClientHeight    =   4920
   ClientLeft      =   2430
   ClientTop       =   2475
   ClientWidth     =   10500
   Icon            =   "frmStoreCreateFileDemo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   10500
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Index           =   1
      Left            =   9780
      TabIndex        =   9
      Top             =   3780
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            ImageKey        =   "open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Index           =   0
      Left            =   4620
      TabIndex        =   8
      Top             =   3180
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            ImageKey        =   "open"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   60
      Top             =   2280
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000009&
      Height          =   285
      Index           =   0
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   360
      Width           =   9495
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   60
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   10500
   End
   Begin VB.PictureBox Picture1 
      Height          =   4395
      Left            =   5280
      ScaleHeight     =   4335
      ScaleWidth      =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&4. Save the Database Field as File"
      Height          =   495
      Index           =   3
      Left            =   5460
      TabIndex        =   3
      Top             =   3720
      Width           =   4875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&1. Get a New File"
      Height          =   495
      Index           =   2
      Left            =   360
      TabIndex        =   0
      Top             =   3120
      Width           =   4875
   End
   Begin VB.Timer tmrStatus 
      Left            =   60
      Top             =   2760
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4545
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18018
            Key             =   "status"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&3. Get the Stored File from the Database Field"
      Height          =   495
      Index           =   1
      Left            =   5460
      TabIndex        =   2
      Top             =   3120
      Width           =   4875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&2. Store the File in a Database Field"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   3660
      Width           =   4875
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":030A
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":041C
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":052E
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":0640
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":0752
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":0864
            Key             =   "save"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":0976
            Key             =   "print"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":0A88
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":0B9A
            Key             =   "iconslarge"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":0CAC
            Key             =   "iconssmall"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":0DBE
            Key             =   "iconslist"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreCreateFileDemo.frx":0ED0
            Key             =   "iconsdetails"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2730
      Index           =   1
      Left            =   5460
      Stretch         =   -1  'True
      ToolTipText     =   " Double-click to stretch or un-stretch. "
      Top             =   240
      Width           =   4830
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2730
      Index           =   0
      Left            =   360
      ToolTipText     =   " Double-click to stretch or un-stretch. "
      Top             =   240
      Width           =   4830
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowItem 
         Caption         =   "mnuWindowItem"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmStoreCreateFileDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Option Explicit
    DefLng A-Z
    
    Const mcsStartingFile   As String = "test.jpg"    'lovely daughter Lynette
    Const mcsSQL            As String = "Select * From tblPicture"      'get all data
    Const mcsDBName         As String = "StorePicture.mdb"              'sample database
    Dim msDBNameFull        As String               'sample database with folder [app.path]
    Dim mDB                 As Database             'open once, close upon unload
    Dim msLastSavedFileName As String               'use this file name if they open the saved file
    Dim mbInFormLoad        As Boolean
    '
    
Private Sub Command1_Click(index As Integer)
    Screen.MousePointer = MouseSet(vbHourglass) 'show working
    Select Case index
        Case 0  'store
            Me.Image1(1).Picture = LoadPicture("")      'clear the picture
            If DoStoreInDB() Then
                Status Me, "File stored in database '" & mDB.Name & "', field 'oPicture'.", , True
            End If
        Case 1  'get from database field
            If DoRetrieveFromDB() Then
                Status Me, "Picture retrieved from database '" & mDB.Name & "', field 'oPicture'.", , True
            Else
'                Status Me, "Picture not found in database '" & mDB.Name & "'.", True, True
            End If
        Case 2  'get new picture from disk
            If DoBrowse() Then
                Me.Image1(1).Picture = LoadPicture("")      'clear the picture
                Status Me, "File '" & Me.Text1(0).text & "' loaded.", , True
            End If
        Case 3  'store as a file
            If DoStorePictureAsFile() Then
            End If
    End Select
    SetButtons
    Screen.MousePointer = MouseSet()            'reset the mouse
End Sub

Private Function DoStorePictureAsFile() As Boolean
' Save the database field as a file
    Dim obj As New CDialog
    Dim pic As New SaveCreateFile.cStoreCreateFile
    Dim rs  As Recordset
    Dim fld As Field
    Dim sTemp As String
    Dim sFileName As String
    Dim sExt    As String
    
    Set rs = mDB.OpenRecordset(mcsSQL)
    With rs
        If Not .EOF Then
            Set fld = .Fields![oPicture]            'set a reference to the picture field
            sFileName = "" & .Fields![sFileName]    'get its 'stored as' file name
            sTemp = sFileName                       'save for later use
            sFileName = GetNameFromPathAndName(sFileName)   'strip off the path
            sExt = GetExtensionFromFileName(sFileName, sFileName)   'now get the extension
            If sExt = "" Then                       'if non then use bmp
                sExt = "bmp"
            End If
            With obj                                'browse for a file name
                .lHwnd = Me.hWnd    'we are the caller
                .Filename = sTemp   'initial file name
                .InitDir = sTemp    'initial directory
                .Flags = cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNOverwritePrompt + cdlOFNExplorer
                .Filter = "Graphics File (." & sExt & ")|*." & sExt & "|All Files|*.*"
                .DialogTitle = "Select a Folder and File Name"  'title
                .DefaultExt = "." & sExt    'default extension
                If .ShowSave() Then         'show the dialog
                    sTemp = .Filename       'get the file name
                    With pic
                        If .CreateFileFromField(fld, sTemp) Then
                            msLastSavedFileName = sTemp
                            Status Me, "File '" & sTemp & "' created.", , True
                        End If
                    End With
                End If
            End With
        Else
            Status Me, "No records in database '" & mDB.Name & "'.", True
        End If
        .Close
    End With
    Set obj = Nothing
    Set pic = Nothing
    Set rs = Nothing
End Function

Private Function DoBrowse() As Boolean
' Browse for a file
    Dim obj As New CDialog
    With obj                'find a graphics file
        .Flags = cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNFileMustExist + cdlOFNExplorer
        .Filter = "Graphics File|*.bmp;*.ico;*.emf;*.wmf;*.jpg;*.gif|All Files|*.*"
        .DialogTitle = "Select a File"
        .InitDir = Trim$(Me.Text1(0).text)
        .lHwnd = Me.hWnd
        .ShowOpen
        If Not .Cancelled Then
            Me.Text1(0).text = .Filename
            If LoadPictureUsingTextBox() Then
                DoBrowse = True
            End If
        End If
    End With
    Set obj = Nothing
End Function

Private Function LoadPictureUsingTextBox() As Boolean
    On Error GoTo LoadPictureUsingTextBoxEH
    With Me.Image1(0)
        .Stretch = True
        .Picture = LoadPicture(Me.Text1(0).text)
        .Visible = True
    End With
    LoadPictureUsingTextBox = True
    Exit Function
LoadPictureUsingTextBoxEH:
    Me.Image1(0).Visible = False
End Function

Private Function DoRetrieveFromDB() As Boolean
' Retrieve the file from the database field
    Dim rs      As Recordset
    Dim fld     As Field
    Dim sTemp   As String
    Dim sFileName As String
    Dim sExt    As String
    Dim obj     As New SaveCreateFile.cStoreCreateFile
    
    Set rs = mDB.OpenRecordset(mcsSQL)      'load the table into a recordset
    With rs                     'set a reference
        If Not .EOF Then        'any records
            sFileName = .Fields![sFileName]     'name the file save as
            sFileName = GetNameFromPathAndName(sFileName)   'strip off folder
            sExt = GetExtensionFromFileName(sFileName, sFileName)   'get extension
            sTemp = TempNamePlease("." & sExt, App.Path, "~tmp")    'make a temp name
            Set fld = .Fields![oPicture]        'set a reference
            With obj
                If .CreateFileFromField(fld, sTemp) Then    'call the dll and create the file
                    On Error GoTo DoRetrieveFromDBEH        'in case the file is not a picture
                    Me.Image1(1).Picture = LoadPicture(sTemp)   'show on form
                    Me.Image1(1).Visible = True             'make sure it is visible
                    DoRetrieveFromDB = True                 'report success
                End If
            End With
        Else
            Status Me, "No records in database '" & mDB.Name & "'.", True
        End If
        .Close
    End With
    GoTo DoRetrieveFromDBExit
DoRetrieveFromDBEH:
    Debug.Print Err.Number, Err.description
    If Err.Number = 481 Then    'not a picture
        Status Me, "Not a picture, probably a file.", , True
    End If
    Me.Image1(1).Visible = False
DoRetrieveFromDBExit:
    If sTemp <> "" Then 'any name
        If Dir$(sTemp) <> "" Then
            Kill sTemp  'dump the temp file if it still exists
        End If
    End If
    Set obj = Nothing
    Set fld = Nothing
    Set rs = Nothing
End Function

Private Function DoStoreInDB() As Boolean
' Store the file in a database field.
    Dim rs      As Recordset
    Dim fldLongBinary As Field
    Dim fldFileName   As Field
    Dim sTemp   As String
    Dim obj     As New SaveCreateFile.cStoreCreateFile
    
    sTemp = Trim$(Me.Text1(0).text)
    Set rs = mDB.OpenRecordset(mcsSQL)  'make sure there is at least one record and store the file name
    With rs
        If .EOF Then
            .AddNew         'add a one 'dummy' record if none
            .Update         'update the table
        End If
    End With
    Set rs = mDB.OpenRecordset(mcsSQL)  'file will be stored in the first record
    With rs
        If Not .EOF Then
            Set fldFileName = .Fields![sFileName]   'set reference to this field
            Set fldLongBinary = .Fields![oPicture]  'set a reference to the file field
            With obj                        'now call the dll
                If .StoreFileIntoField(rs, fldFileName, fldLongBinary, sTemp) Then  'call the dll
                    DoStoreInDB = True
                End If
            End With
        End If
        .Close
    End With
    Set rs = Nothing
    Set fldFileName = Nothing
    Set fldLongBinary = Nothing
    Set obj = Nothing
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
' Unload me if escape key pressed
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = MouseSet(vbHourglass) 'show working
    mbInFormLoad = True             'only in Form_Resize once
    Width = 10140                   'start with this size
    Height = 6100
    gsProductName = App.Title       'for the registry
    gsCompany = "BuyGold"           'for the registry
    SetScreenLocation               'get the screen location from the registry
    mbInFormLoad = False            'done with forcing size, allow Form_Resize to work
    TB_BuildWindowMenu Me, True     'set up the window menu
    MakeACopyOfTheDatabaseAndOpenIt 'make a copy of the database and open it
    Me.Text1(0).text = AddBackslash(App.Path) & mcsStartingFile 'start with the picture
    LoadPictureUsingTextBox                             'load the picture
    Caption = App.ProductName & ", Database: '" & mcsDBName & "'"
    Status Me, "1) First get a new file or 2) store the current file in the database.", , True
    SetButtons
    Screen.MousePointer = MouseSet()    'show normal mouse
End Sub

Private Sub SetScreenLocation()
' Get the screen location from the registry
    Dim l, t, w, h
    Dim iState As Integer   'window state
    If TB_GetFormInformation(Me, iState, l, t, w, h) Then   'get last saved screen location
        TB_MakeSureOnScreen l, t, w, h  'make sure it is in the desktop work area
        Me.Move l, t, w, h
        Me.WindowState = iState
    Else
        TB_CenterForm32 Me              'center in desktop work area if no registry entries
    End If
End Sub

Private Sub MakeACopyOfTheDatabaseAndOpenIt()
    On Error Resume Next
    ' delete the current copy of the sample database
    msDBNameFull = AddBackslash(App.Path) & mcsDBName   'db name
    If Dir$(msDBNameFull) <> "" Then    'does it exist?
        Kill msDBNameFull
    End If
    ' make a copy of the database
    FileCopy AddBackslash(App.Path) & "Copy of " & mcsDBName, msDBNameFull
    Set mDB = OpenDatabase(msDBNameFull)
End Sub

Private Sub Form_Resize()
' Resize the controls
    Dim l, t, w, h
    If mbInFormLoad Then Exit Sub
    Screen.MousePointer = MouseSet(vbHourglass)
    If WindowState <> vbMinimized Then
        If WindowState <> vbMaximized Then
            Width = NotTooSmall(Width, 6000)
            Height = NotTooSmall(Height, 4400)
        End If
        Me.Image1(0).Visible = False        'prevent flicker
        Me.Image1(1).Visible = False
        With Me.Picture1
            l = (ScaleWidth - .Width) \ 2
            t = Me.Picture2.Height
            w = .Width
            h = ScaleHeight - t - Me.StatusBar1.Height + 10
            .Move l, t, w, h
        End With
        With Me.Command1(0)
            h = .Height
            l = 120
            t = ScaleHeight - Me.StatusBar1.Height - h - 30
            w = (ScaleWidth \ 2) - (l * 3) / 2
            .Move l, t, w, h
        End With
        With Me.Command1(3)
            l = l + w + l
            .Move l, t, w, h
        End With
        With Me.Command1(2)
            l = Me.Command1(0).Left
            t = Me.Command1(0).Top - Me.Command1(0).Height - 30
            w = Me.Command1(0).Width
            h = Me.Command1(0).Height
            .Move l, t, w, h
        End With
        With Me.Command1(1)
            l = Me.Command1(3).Left
            .Move l, t, w, h
        End With
        
        With Me.Image1(0)
            l = Me.Command1(0).Left
            t = l
            w = Me.Command1(0).Width
            h = Me.Command1(2).Top - t * 2
            .Move l, t, w, h
        End With
        With Me.Image1(1)
            l = Me.Command1(1).Left
            .Move l, t, w, h
        End With
        With Me.Text1(0)
            l = Me.Image1(0).Left + 90
'            t = Me.Image1(0).Top + (Me.Image1(0).Height - .Height) \ 2 ' - 90
            t = Me.Image1(0).Top + 90
            w = (Me.Image1(0).Width * 2) - 60
            .Move l, t, w
        End With
        With Me.Toolbar1(0)
            l = Me.Command1(2).Left + Me.Command1(2).Width - .Width - 90
            t = Me.Command1(2).Top + (Me.Command1(2).Height - .Height) \ 2
            w = 340
            h = 300
            .Move l, t, w, h
            .ToolTipText = " Open with the file's associated editor. "
        End With
        With Me.Toolbar1(1)
            l = Me.Command1(3).Left + Me.Command1(3).Width - .Width - 90
            t = Me.Command1(3).Top + (Me.Command1(3).Height - .Height) \ 2
            .Move l, t, w, h
            .ToolTipText = " Open with the file's associated editor after saving it as a file. "
        End With
        
        Me.Image1(0).Visible = True
        Me.Image1(1).Visible = True
    End If
    Screen.MousePointer = MouseSet()    'reset
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TB_SaveFormInformation Me   'save the screen location
    mDB.Close                   'close the database
    Set mDB = Nothing           'release resources
End Sub

Private Sub Image1_DblClick(index As Integer)
' stretch or un-stretch
    With Me.Image1(index)
        .Stretch = Not .Stretch
        Form_Resize
    End With
    Me.Refresh
End Sub

Private Sub mnuFileItem_Click(index As Integer)
    Select Case index
        Case 0  'exit
            Unload Me
    End Select
End Sub



Private Sub mnuWindow_Click()
    TB_SetWindowMenu Me         'setup the window menu
End Sub

Private Sub mnuWindowItem_Click(index As Integer)
    TB_WindowItem Me, index     'process the window menu item
End Sub

Private Sub Text1_GotFocus(index As Integer)
    With Text1(index)       'highlight the text
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub tmrStatus_Timer()
    StatusOff Me        'turn off the status, if enabled
End Sub

Private Sub Timer1_Timer()
    SetButtons      'set command buttons
End Sub

Private Sub SetButtons()
' Set the command buttons enabled or disabled
    Dim bValue As Boolean
    Dim rs As Recordset
    Dim bValueToolBar As Boolean
    
    If Me.Text1(0).text <> "" Then              'any file name
        If Dir$(Me.Text1(0).text) <> "" Then
            bValue = True                       'set enabled
        End If
    End If
    If bValue Then                              'if have a file name then do we have a record in the database
        Set rs = mDB.OpenRecordset(mcsSQL)
        With rs
            If Not .EOF Then                    'any records?
                If "" & .Fields![sFileName] = "" Then   'any file name?
                    bValue = False              'no file name
                End If
            Else
                bValue = False                  'no records
            End If
            .Close      'done with the recordset
        End With
    End If
    bValueToolBar = bValue
    Me.Command1(1).Enabled = bValue 'set buttons
    Me.Command1(3).Enabled = bValue
    If bValueToolBar Then
        If msLastSavedFileName = "" Then
            bValueToolBar = False
        End If
    End If
    Me.Toolbar1(1).Visible = bValueToolBar
    Set rs = Nothing    'release resources
End Sub

Private Sub Toolbar1_ButtonClick(index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case index
        Case 0  'input file
            Select Case Button.Key
                Case "open"
                    DoOpenFileWithAssociatedEditor Me.Text1(0).text
            End Select
        Case 1  'created file
            Select Case Button.Key
                Case "open"
                    If msLastSavedFileName <> "" Then
                        DoOpenFileWithAssociatedEditor msLastSavedFileName
                    End If
            End Select
    End Select
End Sub

Private Sub DoOpenFileWithAssociatedEditor(sFileName As String)
    Dim obj As New clsShellEX
    With obj
        .lHwnd = Me.hWnd
        .sFileName = sFileName
        .ShellExecuteUsingFileName
    End With
End Sub








