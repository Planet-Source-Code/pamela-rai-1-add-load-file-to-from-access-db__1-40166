VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

    Option Explicit
    
    ' API declarations
    

' Common Dialog Constants

    Const cdlCancel As Long = &H7FF3

'    Public Const cdlOFNExplorer = &H80000
'    Public Const cdlOFNHideReadOnly = &H4
'    Public Const cdlOFNAllowMultiselect = &H200
'    Public Const cdlOFNReadOnly = &H1
'    Public Const cdlOFNFileMustExist = &H1000
'    Public Const cdlOFNPathMustExist = &H800

' File Open/Save Dialog Box Flags
    Const cdlOFNReadOnly = &H1           'Checks Read-Only check box for Open and Save As dialog boxes.
    Const cdlOFNOverwritePrompt = &H2    'Causes the Save As dialog box to generate a message box if the selected file already exists.
    Const cdlOFNHideReadOnly = &H4       'Hides the Read-Only check box.
    Const cdlOFNNoChangeDir = &H8        'Sets the current directory to what it was when the dialog box was invoked.
    Const cdlOFNHelpButton = &H10        'Causes the dialog box to display the Help button.
    Const cdlOFNNoValidate = &H100       'Allows invalid characters in the returned filename.
    Const cdlOFNAllowMultiselect = &H200 'Allows the File Name list box to have multiple selections.
    Const cdlOFNExtensionDifferent = &H400 'The extension of the returned filename is different from the extension set by the DefaultExt property.

    Const cdlOFNPathMustExist = &H800    'User can enter only valid path names.
    Const cdlOFNFileMustExist = &H1000   'User can enter only names of existing files.
    Const cdlOFNCreatePrompt = &H2000    'Sets the dialog box to ask if the user wants to create a file that doesn't currently exist.
    Const cdlOFNShareAware = &H4000      'Sharing violation errors will be ignored.
    Const cdlOFNNoReadOnlyReturn = &H8000 'The returned file doesn't have the Read-Only attribute set and won't be in a write-protected directory.
    Const cdlOFNExplorer = &H80000       'Use the Explorer-like Open A File dialog box template.  (Windows 95 only.)
    Const cdlOFNNoDereferenceLinks = &H100000 'Do not dereference shortcuts (shell links).  By default, choosing a shortcut causes it to be dereferenced by the shell.  (Windows 95 only.)

    Const cdlOFNLongNames = &H200000     'Use Long filenames.  (Windows 95 only.)

' Color Dialog Box Flags
    Const cdlCCRGBInit = &H1             'Sets initial color value for the dialog box.
    Const cdlCCFullOpen = &H2            'Entire dialog box is displayed, including the Define Custom Colors section.
    Const cdlCCPreventFullOpen = &H4     'Disables the Define Custom Colors section of the dialog box.
    Const cdlCCHelpButton = &H8          'Dialog box displays a Help button.

' Fonts Dialog Box Flags
    Const cdlCFScreenFonts = &H1         'Dialog box lists only screen fonts supported by the system.
    Const cdlCFPrinterFonts = &H2        'Dialog box lists only fonts supported by the printer.
    Const cdlCFBoth = &H3                'Dialog box lists available screen and printer fonts.
    Const cdlCFHelpButton = &H4          'Dialog box displays a Help button.
    Const cdlCFEffects = &H100           'Dialog box enables strikeout, underline, and color effects.
    Const cdlCFApply = &H200             'Dialog box enables the Apply button.
    Const cdlCFANSIOnly = &H400          'Dialog box allows only a selection of fonts that use the Windows character set.
    Const cdlCFNoVectorFonts = &H800     'Dialog box should not allow vector-font selections.

    Const cdlCFNoSimulations = &H1000    'Dialog box should not allow graphic device interface (GDI) font simulations.
    Const cdlCFLimitSize = &H2000        'Dialog box should select only font sizes within the range specified by the Min and Max properties.
    Const cdlCFFixedPitchOnly = &H4000   'Dialog box should select only fixed-pitch fonts.
    Const cdlCFWYSIWYG = &H8000          'Dialog box should allow only the selection of fonts available to both the screen and printer.
    Const cdlCFForceFontExist = &H10000  'An error dialog box is displayed if a user selects a font or style that doesn't exist.
    Const cdlCFScalableOnly = &H20000    'Dialog box should allow only the selection of scalable fonts.

    Const cdlCFTTOnly = &H40000          'Dialog box should allow only the selection of TrueType fonts.
    Const cdlCFNoFaceSel = &H80000       'No font name selected.
    Const cdlCFNoStyleSel = &H100000     'No font style selected.
    Const cdlCFNoSizeSel = &H200000      'No font size selected.

'Printer Dialog Box Flags
    Const cdlPDAllPages = &H0            'Returns or sets state of All Pages option button.
    Const cdlPDCollate = &H10            'Returns or sets state of Collate check box.
    Const cdlPDDisablePrintToFile = &H80000 'Disables the Print To File check box.
    Const cdlPDHidePrintToFile = &H100000    'The Print To File check box isn't displayed.
    Const cdlPDNoPageNums = &H8          'Returns or sets the state of the Pages option button.
    Const cdlPDNoSelection = &H4         'Disables the Selection option button.
    Const cdlPDNoWarning = &H80          'Prevents a warning message when there is no default printer.
    Const cdlPDPageNums = &H2            'Returns or sets the state of the Pages option button.

    Const cdlPDPrintSetup = &H40         'Displays the Print Setup dialog box rather than the Print dialog box.
    Const cdlPDPrintToFile = &H20        'Returns or sets the state of the Print To File check box.
    Const cdlPDReturnDC = &H100          'Returns a device context for the printer selection value returned in the hDC property of the dialog box.
    Const cdlPDReturnDefault = &H400     'Returns default printer name.
    Const cdlPDReturnIC = &H200          'Returns an information context for the printer selection value returned in the hDC property of the dialog box.
    Const cdlPDSelection = &H1           'Returns or sets the state of the Selection option button.
    Const cdlPDHelpButton = &H800        'Dialog box displays the Help button.
    Const cdlPDUseDevModeCopies = &H40000 'Sets support for multiple copies action; depends upon whether or not printer supports multiple copies.
#If Win16 Then
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Integer
        hInstance As Integer
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
    Private Declare Function GetOpenFileName Lib "commdlg.dll" (pOpenfilename As OPENFILENAME) As Integer
    Private Declare Function GetSaveFileName Lib "commdlg.dll" (pOpenfilename As OPENFILENAME) As Integer
    Private Declare Function CommDlgExtendedError Lib "commdlg.dll" () As Long
    
#Else
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
    
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
    Private Declare Function GetVersion Lib "kernel32" () As Long
#End If

    ' Public Properties common to all common dialogs
    Public CancelError As Boolean
    Public Cancelled   As Boolean   '1998/01/18 Added by Larry.
    Public Flags As Long
    Public Tag As String
    Private m_Action As Integer
    
    ' Help Properties (only for compatibility with OCX)
    Public HelpFile As String
    Public HelpCommand As Long
    Public HelpContext As Long
    Public HelpKey As String
    
    ' Open/Save File Common Dialog properties
    Public DialogTitle As String
    Public Filename As String
    Public FileTitle As String
    Public InitDir As String
    Public FilterIndex As Integer
    Public DefaultExt As String
    Public MaxFileSize As Integer
    
    Public CustomFilter As String
    Public CustomFilterFilespec As String
    Public Error As Integer
    
    Private mlHwnd   As Long
    
    
    Private m_SelectedFileCount As Integer
    Private m_SelectedFile() As String
    Private m_Filter As String
    
    Dim sepChar As String
    Dim clearFilters As Boolean
    '

Private Sub Class_Initialize()

    ' default value for a few properties
    MaxFileSize = 256
    FilterIndex = 1
    ReDim m_SelectedFile(1 To 1) As String

End Sub

Public Property Let Action(newValue As Integer)

    'This behaves exactly as the old VB3 Action property

    Select Case newValue
    Case 1
        ShowOpen
    Case 2
        ShowSave
    End Select

End Property

Public Property Get MultiSelect() As Integer

    ' Single or multi-selection Open File Dialog
    '   0 = single selection
    '   1 = multiple selection, non-Explorer style
    '   2 = multiple selection, Explorer style

    If (Flags And cdlOFNAllowMultiselect) = 0 Then
        MultiSelect = 0
    ElseIf (Flags And cdlOFNExplorer) Then
        MultiSelect = 2
    Else
        MultiSelect = 1
    End If

End Property

Public Property Let MultiSelect(newValue As Integer)

    ' Single or multi-selection Open File Dialog
    '   0 = single selection
    '   1 = multiple selection, non-Explorer style
    '   2 = multiple selection, Explorer style

    Select Case newValue
    Case 0
        Flags = Flags And Not cdlOFNAllowMultiselect And Not cdlOFNExplorer
    Case 1
        Flags = (Flags And Not cdlOFNExplorer) Or cdlOFNAllowMultiselect
    Case 2
        Flags = Flags Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    End Select
    
End Property

Public Property Get Filter() As String

    Filter = m_Filter

End Property

Public Property Let Filter(newValue As String)

    m_Filter = newValue
    clearFilters = False

End Property

Public Sub AddFilter(description As String, filespec As String, Optional defaultFilter)

    ' Add a new filter to the list
    ' multiple Filespecs must be separated using semicolons
    ' if defaultFilter = True this becomes the default filter
    
    If IsMissing(defaultFilter) Then
        defaultFilter = False
    End If
    
    ' clear filters if called after a Show command
    If clearFilters Then Filter = ""
    
    ' add to existing filter
    If Len(Filter) > 0 And Right$(Filter, 1) <> "|" Then
        Filter = Filter & "|"
    End If
    ' add filespec to description only if it is not a multiple description
    Filter = Filter & description & IIf(InStr(filespec, ";"), "", " (" & filespec & ")") & "|" & filespec
    
    If defaultFilter <> 0 Then
        ' we must count existing filters
        Dim i As Integer
        FilterIndex = 0
        Do
            FilterIndex = FilterIndex + 1
            i = InStr(i + 1, Filter, "|")
            i = InStr(i + 1, Filter, "|")
        Loop While i
    End If

End Sub

Public Property Get FileReadOnly() As Integer

    'The visibility and state of the ReadOnly checkbox
    '  0   = unchecked
    '  1   = checked
    '  2   = hidden

    If Flags And cdlOFNHideReadOnly Then
        FileReadOnly = 2
    ElseIf Flags And cdlOFNReadOnly Then
        FileReadOnly = 1
    Else
        FileReadOnly = 0
    End If

End Property

Public Property Let FileReadOnly(newValue As Integer)

    ' The visibility and state of the ReadOnly checkbox
    '  0   = unchecked
    '  1   = checked
    '  2   = hidden
    
    Select Case newValue
    Case 0
        Flags = Flags And Not cdlOFNHideReadOnly And Not cdlOFNReadOnly
    Case 1
        Flags = (Flags And Not cdlOFNHideReadOnly) Or cdlOFNReadOnly
    Case 2
        Flags = Flags Or cdlOFNHideReadOnly
    End Select

End Property

Public Property Get FileMustExist() As Integer

    ' What to do if a file does not exists
    '    0 = no special action
    '    1 = file must exist
    '    2 = path must exist
    '    3 = prompt to create new file
    '    4 = prompt to create new file, but path must exist
    '    5 = prompt before overwriting (File Save)
    '    6 = prompt before overwriting, but path must exist
    
    If (Flags And cdlOFNFileMustExist) > 0 And (Flags And cdlOFNPathMustExist) > 0 Then
        FileMustExist = 1
    ElseIf (Flags And cdlOFNCreatePrompt) > 0 And (Flags And cdlOFNPathMustExist) > 0 Then
        FileMustExist = 4
    ElseIf (Flags And cdlOFNCreatePrompt) > 0 Then
        FileMustExist = 3
    ElseIf (Flags And cdlOFNOverwritePrompt) > 0 And (Flags And cdlOFNPathMustExist) > 0 Then
        FileMustExist = 6
    ElseIf (Flags And cdlOFNOverwritePrompt) > 0 Then
        FileMustExist = 5
    ElseIf (Flags And cdlOFNPathMustExist) > 0 Then
        FileMustExist = 2
    Else
        FileMustExist = 0
    End If

End Property

Public Property Let FileMustExist(newValue As Integer)

    ' What to do if a file does not exists
    '    0 = no special action
    '    1 = file must exist
    '    2 = path must exist
    '    3 = prompt to create new file
    '    4 = prompt to create new file, but path must exist
    '    5 = prompt before overwriting (File Save)
    '    6 = prompt before overwriting, but path must exist
    
    Select Case newValue
    Case 1
        Flags = Flags Or cdlOFNPathMustExist Or cdlOFNFileMustExist
    Case 2
        Flags = Flags Or cdlOFNPathMustExist
    Case 3
        Flags = (Flags And Not cdlOFNPathMustExist And Not cdlOFNFileMustExist) Or cdlOFNCreatePrompt
    Case 4
        Flags = (Flags And Not cdlOFNFileMustExist) Or cdlOFNPathMustExist Or cdlOFNCreatePrompt
    Case 5
        Flags = (Flags And Not cdlOFNPathMustExist) Or cdlOFNOverwritePrompt
    Case 6
        Flags = Flags Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
    End Select

End Property

Public Property Get FileMustBeWritable() As Boolean

    FileMustBeWriteable = Flags And cdlOFNNoReadOnlyReturn

End Property

Public Property Let FileMustBeWriteable(newValue As Boolean)

    If newValue Then
        Flags = Flags Or cdlOFNNoReadOnlyReturn
    Else
        Flags = Flags And Not cdlOFNNoReadOnlyReturn
    End If

End Property

Public Sub AddDefaultFilters(ctrl As Variant)

    ' add default filters for a control
    
    If TestProperties(ctrl, "TextRTF") Then
        AddFilter "All Document files", "*.rtf;*.txt"
        AddFilter "Rich Text Box files", "*.rtf", True
        AddFilter "Text Files", "*.txt"
    ElseIf TestProperties(ctrl, "Picture") Then
        AddFilter "All Image files", "*.bmp;*.dib;*.wmf;*.emf;*.ico;*.cur"
        AddFilter "Bitmaps", "*.bmp;*.dib"
        AddFilter "Metafiles", "*.wmf;*.emf"
        AddFilter "Icons", "*.ico;*.cur"
    ElseIf TestProperties(ctrl, "ListIndex") Then
        AddFilter "Data Files", "*.dat;*.lst"
    ElseIf TestProperties(ctrl, "Text") Or VarType(ctrl) = vbString Then
        AddFilter "Text files ", "*.txt"
    Else
        Exit Sub
    End If
    
    If Len(Filter) Then
        Filter = "All Files (*.*)|*.*|" & Filter
        FilterIndex = 2
    End If

End Sub

Public Property Get SelectedFileCount() As Integer

    ' Returns the number of selected files

    SelectedFileCount = m_SelectedFileCount

End Property

Public Property Get SelectedFile(ByVal index As Integer) As String

    ' Returns the N-th selected file

    If index < 1 Or index > m_SelectedFileCount Then
        SelectedFile = ""
    Else
        SelectedFile = m_SelectedFile(index)
    End If

End Property

Private Function ShowFileDialog() As Boolean

    ' Show an Open/Save File Common Dialog (private procedure)
    '    action = 1  -> Open
    '    action = 2  -> Save

    Dim ofn As OPENFILENAME
    Dim tmpFilter As String
    Dim tmpFlags As Long
    Dim i As Integer
    Dim ok As Integer
    Dim fileIndex As Integer
    Dim extIndex As Integer
    Dim saveCustomFilter As String
    
    ' convert all "|"'s into null characters
    tmpFilter = Filter
    Do
        i = InStr(tmpFilter, "|")
        If i = 0 Then Exit Do
        Mid$(tmpFilter, i, 1) = vbNullChar
    Loop
    
    ' adjust Flag to avoid problems under WinNT 3.51
    tmpFlags = Flags And Not cdlOFNHelpButton And Not cdlOFNLongNames
    sepChar = " "
#If Win32 Then
    If (Flags And cdlOFNAllowMultiselect) > 0 And (Flags And cdlOFNExplorer) > 0 Then
        If (GetVersion() And &H800000FF) = 3 Then
            ' do not use Explorer style under Windows NT 3.51
            tmpFlags = tmpFlags And Not cdlOFNExplorer
        Else
            sepChar = vbNullChar
        End If
    End If
#End If
    
    With ofn
        .lStructSize = Len(ofn)
'        .hwndOwner = Screen.ActiveForm.hWnd
        .hwndOwner = mlHwnd
        .nFilterIndex = FilterIndex
        .lpstrFilter = tmpFilter
        .lpstrDefExt = DefaultExt
        .Flags = tmpFlags
        .lpstrTitle = DialogTitle
        .nMaxFile = MaxFileSize
        .lpstrFile = Left$(Filename & vbNullChar & Space$(MaxFileSize), MaxFileSize)
        .lpstrInitialDir = InitDir

        If Len(CustomFilter) Then
            .lpstrCustomFilter = CustomFilter & vbNullChar & IIf(CustomFilterFilespec <> "", CustomFilterFilespec, "*") & vbNullChar & Space$(99)
            .nMaxCustFilter = Len(.lpstrCustomFilter)
            ' save for later comparison
            saveCustomFilter = .lpstrCustomFilter
        End If

        ' call one of the two dialogs
        If m_Action = 1 Then
            ok = GetOpenFileName(ofn)
        ElseIf m_Action = 2 Then
            ok = GetSaveFileName(ofn)
        End If
        
        ShowFileDialog = ok
        Error = 0
        
        ' modify custom filter filespec only if user selected
        ' the custom filter and typed a new filespec
        If .nFilterIndex = 0 And .lpstrCustomFilter <> saveCustomFilter Then
            i = InStr(.lpstrCustomFilter, vbNullChar)
            CustomFilterFilespec = Mid$(.lpstrCustomFilter, i + 1, InStr(i + 1, .lpstrCustomFilter, vbNullChar) - i - 1)
        End If
        
        If ok Then
            ' modify properties if everything was OK
            Err = 0
            Flags = .Flags
            FilterIndex = .nFilterIndex
            ' get filename
            If sepChar = " " Then
                ' filename is usually terminated by a single null byte
                Filename = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
            Else
                ' unless the null character is used to delimit multiple files
                Filename = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar & vbNullChar) - 1)
            End If
            ' get new value for InitDir (for compatibility with Common Dialog control)
            If (Flags And cdlOFNNoChangeDir) = 0 Then
                InitDir = CurDir$
            End If
            ' get new value for FileTitle property
            fileIndex = 1
            extIndex = Len(Filename) + 1
            For i = Len(Filename) To 1 Step -1
                If Mid$(Filename, i, 1) = "." Then
                    extIndex = i
                ElseIf InStr(":\", Mid$(Filename, i, 1)) Then
                    fileIndex = i + 1
                    Exit For
                End If
            Next
'            FileTitle = Mid$(Filename, fileIndex, extIndex - fileIndex)
            FileTitle = Mid$(Filename, fileIndex)   '1998/01/18 Was not returning extension. Larry
            
        Else
            ' something did not work
            Error = CommDlgExtendedError()
            If Error = 0 Then
                ' user pressed Cancel
                ' raise an error if the user requested it
                If CancelError Then
                    Cancelled = True
'                    Err.Clear              '1998/01/18 Can't raise error if no reference to CommonDialog OCX.
'                    Err.Raise cdlCancel
                End If
            End If
        End If
    End With
    
    ' remember to clear filter
    clearFilters = True

End Function

Public Function ShowOpen(Optional target As Variant) As Boolean

    Cancelled = False

    'Show the Open File Common Dialog
    
    ' build filters if possible
    If Not IsMissing(target) Then
        If Filter = "" Then
            AddDefaultFilters target
        End If
    End If
    
    ' exit right now if user did not select any file
    m_Action = 1
    If Not ShowFileDialog() Then
        ShowOpen = False
        Cancelled = True
        Exit Function
    End If
    ShowOpen = True
    
    If Flags And cdlOFNAllowMultiselect Then
        ' Multiple file selection
        Dim fileNameList As String
        Dim dirName As String
        Dim sepIndex As Integer
        
        ' zero number of selected files
        m_SelectedFileCount = 0
        fileNameList = Filename
        sepIndex = InStr(fileNameList, sepChar)
        If sepIndex = 0 Then
            ' user selected only one file
            dirName = ""
        Else
            ' user selected multiple file
            dirName = Left$(fileNameList, sepIndex - 1)
            If Right$(dirName, 1) <> "\" Then dirName = dirName & "\"
        End If
        
        ' extract file list
        Do While sepIndex < Len(fileNameList)
            fileNameList = Mid$(fileNameList, sepIndex + 1)
            sepIndex = InStr(fileNameList & sepChar, sepChar)
            m_SelectedFileCount = m_SelectedFileCount + 1
            If UBound(m_SelectedFile) < m_SelectedFileCount Then
                ReDim Preserve m_SelectedFile(1 To m_SelectedFileCount + 50) As String
            End If
            m_SelectedFile(m_SelectedFileCount) = dirName & Left$(fileNameList, sepIndex - 1)
        Loop
        ReDim Preserve m_SelectedFile(1 To m_SelectedFileCount)
    
    Else
        ' single file selection
        m_SelectedFileCount = 1
        ReDim m_SelectedFile(1) As String
        m_SelectedFile(1) = Filename
        ' is there a target to load ?
        If Not IsMissing(target) Then
            LoadInto target
        End If
    End If

End Function

Public Function ShowSave(Optional target As Variant) As Boolean

    Cancelled = False
    
    ' Show the Save File Common Dialog
    
    ' build filters if possible
    If Not IsMissing(target) Then
        If Filter = "" Then
            AddDefaultFilters target
        End If
    End If
    
    ' exit right now if user did not select any file
    Flags = Flags Or cdlOFNHideReadOnly
    m_Action = 2
    If Not ShowFileDialog() Then
        Cancelled = True
        ShowSave = False
        Exit Function
    End If
    ShowSave = True
    
    ' Process user's choices
    m_SelectedFileCount = 1
    ReDim m_SelectedFile(1) As String
    m_SelectedFile(1) = Filename
    
    ' is there a target to save ?
    If Not IsMissing(target) Then
        SaveFrom target
    End If
    
End Function

Private Function TestProperties(ByVal ctrl As Variant, ParamArray props()) As Integer

    ' return True if the control supports all the specified properties
    
    On Error Resume Next
    
    Dim dummy As Variant
    Dim index As Integer
    
    TestProperties = 0
    
    For index = LBound(props) To UBound(props)
        Select Case LCase$(props(index))
        Case "text"
            dummy = ctrl.text
        Case "picture"
            dummy = ctrl.Picture
        Case "textrtf"
            dummy = ctrl.TextRTF
        Case "forecolor"
            dummy = ctrl.ForeColor
        Case "backcolor"
            dummy = ctrl.BackColor
        Case "color"
            dummy = ctrl.Color
        Case "listindex"
            dummy = ctrl.ListIndex
        Case "font"
            Set dummy = ctrl.Font
        End Select
        If Err Then Exit Function
    Next

    ' all tests were OK
    TestProperties = True

End Function

Public Sub SaveFrom(target As Variant)

    ' load the contents of a file into a control
    Dim fileHandle  As Integer
    Dim index       As Integer
    
    If m_Action = 2 And m_SelectedFileCount = 1 Then
        ' File Open common dialog
        fileHandle = FreeFile()
        
        If TestProperties(target, "TextRTF") Then
            target.SaveFile Filename
        
        ElseIf TestProperties(target, "Picture") Then
            SavePicture target.Picture, Filename
        
        ElseIf TestProperties(target, "ListIndex") Then
            Open Filename For Output As fileHandle
            For index = 0 To target.ListCount - 1
                Print #fileHandle, target.List(index)
            Next
            Close #fileHandle
        
        ElseIf TestProperties(target, "Text") Then
            Open Filename For Output As fileHandle
            Print #fileHandle, target.text;
            Close #fileHandle
        
        ElseIf VarType(target) = vbString Then
            Open Filename For Output As fileHandle
            Print #fileHandle, target;
            Close #fileHandle
        End If
    End If

End Sub
Public Sub LoadInto(target As Variant)

    ' load the contents of a file into a control
    Dim Filename As String
    Dim fileHandle As Integer
    Dim text As String
    
    If m_Action = 1 And m_SelectedFileCount = 1 Then
        ' File Open common dialog
        Filename = SelectedFile(1)
        fileHandle = FreeFile()
        
        If TestProperties(target, "TextRTF") Then
            target.LoadFile Filename
        
        ElseIf TestProperties(target, "Picture") Then
            target.Picture = LoadPicture(Filename)
        
        ElseIf TestProperties(target, "ListIndex") Then
            target.Clear
            Open Filename For Input As fileHandle
            Do Until EOF(fileHandle)
                Line Input #fileHandle, text
                target.AddItem text
            Loop
            Close #fileHandle
            If target.ListCount > 0 Then target.ListIndex = 0
        
        ElseIf TestProperties(target, "Text") Then
            Open Filename For Input As fileHandle
            target.text = Input$(LOF(fileHandle), fileHandle)
            Close #fileHandle
        
        ElseIf VarType(target) = vbString Then
            Open Filename For Input As fileHandle
            target = Input$(LOF(fileHandle), fileHandle)
            Close #fileHandle
        End If
    
    End If

End Sub

Public Property Let lHwnd(lHwnd As Long)
    mlHwnd = lHwnd
End Property
Public Property Get lHwnd() As Long
    lHwnd = mlHwnd
End Property

