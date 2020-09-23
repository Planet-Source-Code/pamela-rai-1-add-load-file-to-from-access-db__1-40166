Attribute VB_Name = "modCommonDialogConstants"

    
    Option Explicit
    DefLng A-Z

' Common Dialog Constants

    Public Const cdlCancel As Long = &H7FF3

'    Public Const cdlOFNExplorer = &H80000
'    Public Const cdlOFNHideReadOnly = &H4
'    Public Const cdlOFNAllowMultiselect = &H200
'    Public Const cdlOFNReadOnly = &H1
'    Public Const cdlOFNFileMustExist = &H1000
'    Public Const cdlOFNPathMustExist = &H800

' File Open/Save Dialog Box Flags
    Public Const cdlOFNReadOnly = &H1           'Checks Read-Only check box for Open and Save As dialog boxes.
    Public Const cdlOFNOverwritePrompt = &H2    'Causes the Save As dialog box to generate a message box if the selected file already exists.
    Public Const cdlOFNHideReadOnly = &H4       'Hides the Read-Only check box.
    Public Const cdlOFNNoChangeDir = &H8        'Sets the current directory to what it was when the dialog box was invoked.
    Public Const cdlOFNHelpButton = &H10        'Causes the dialog box to display the Help button.
    Public Const cdlOFNNoValidate = &H100       'Allows invalid characters in the returned filename.
    Public Const cdlOFNAllowMultiselect = &H200 'Allows the File Name list box to have multiple selections.
    Public Const cdlOFNExtensionDifferent = &H400 'The extension of the returned filename is different from the extension set by the DefaultExt property.

    Public Const cdlOFNPathMustExist = &H800    'User can enter only valid path names.
    Public Const cdlOFNFileMustExist = &H1000   'User can enter only names of existing files.
    Public Const cdlOFNCreatePrompt = &H2000    'Sets the dialog box to ask if the user wants to create a file that doesn't currently exist.
    Public Const cdlOFNShareAware = &H4000      'Sharing violation errors will be ignored.
    Public Const cdlOFNNoReadOnlyReturn = &H8000 'The returned file doesn't have the Read-Only attribute set and won't be in a write-protected directory.
    Public Const cdlOFNExplorer = &H80000       'Use the Explorer-like Open A File dialog box template.  (Windows 95 only.)
    Public Const cdlOFNNoDereferenceLinks = &H100000 'Do not dereference shortcuts (shell links).  By default, choosing a shortcut causes it to be dereferenced by the shell.  (Windows 95 only.)

    Public Const cdlOFNLongNames = &H200000     'Use Long filenames.  (Windows 95 only.)

' Color Dialog Box Flags
    Public Const cdlCCRGBInit = &H1             'Sets initial color value for the dialog box.
    Public Const cdlCCFullOpen = &H2            'Entire dialog box is displayed, including the Define Custom Colors section.
    Public Const cdlCCPreventFullOpen = &H4     'Disables the Define Custom Colors section of the dialog box.
    Public Const cdlCCHelpButton = &H8          'Dialog box displays a Help button.

' Fonts Dialog Box Flags
    Public Const cdlCFScreenFonts = &H1         'Dialog box lists only screen fonts supported by the system.
    Public Const cdlCFPrinterFonts = &H2        'Dialog box lists only fonts supported by the printer.
    Public Const cdlCFBoth = &H3                'Dialog box lists available screen and printer fonts.
    Public Const cdlCFHelpButton = &H4          'Dialog box displays a Help button.
    Public Const cdlCFEffects = &H100           'Dialog box enables strikeout, underline, and color effects.
    Public Const cdlCFApply = &H200             'Dialog box enables the Apply button.
    Public Const cdlCFANSIOnly = &H400          'Dialog box allows only a selection of fonts that use the Windows character set.
    Public Const cdlCFNoVectorFonts = &H800     'Dialog box should not allow vector-font selections.

    Public Const cdlCFNoSimulations = &H1000    'Dialog box should not allow graphic device interface (GDI) font simulations.
    Public Const cdlCFLimitSize = &H2000        'Dialog box should select only font sizes within the range specified by the Min and Max properties.
    Public Const cdlCFFixedPitchOnly = &H4000   'Dialog box should select only fixed-pitch fonts.
    Public Const cdlCFWYSIWYG = &H8000          'Dialog box should allow only the selection of fonts available to both the screen and printer.
    Public Const cdlCFForceFontExist = &H10000  'An error dialog box is displayed if a user selects a font or style that doesn't exist.
    Public Const cdlCFScalableOnly = &H20000    'Dialog box should allow only the selection of scalable fonts.

    Public Const cdlCFTTOnly = &H40000          'Dialog box should allow only the selection of TrueType fonts.
    Public Const cdlCFNoFaceSel = &H80000       'No font name selected.
    Public Const cdlCFNoStyleSel = &H100000     'No font style selected.
    Public Const cdlCFNoSizeSel = &H200000      'No font size selected.

'Printer Dialog Box Flags
    Public Const cdlPDAllPages = &H0            'Returns or sets state of All Pages option button.
    Public Const cdlPDCollate = &H10            'Returns or sets state of Collate check box.
    Public Const cdlPDDisablePrintToFile = &H80000 'Disables the Print To File check box.
    Public Const cdlPDHidePrintToFile = &H100000    'The Print To File check box isn't displayed.
    Public Const cdlPDNoPageNums = &H8          'Returns or sets the state of the Pages option button.
    Public Const cdlPDNoSelection = &H4         'Disables the Selection option button.
    Public Const cdlPDNoWarning = &H80          'Prevents a warning message when there is no default printer.
    Public Const cdlPDPageNums = &H2            'Returns or sets the state of the Pages option button.

    Public Const cdlPDPrintSetup = &H40         'Displays the Print Setup dialog box rather than the Print dialog box.
    Public Const cdlPDPrintToFile = &H20        'Returns or sets the state of the Print To File check box.
    Public Const cdlPDReturnDC = &H100          'Returns a device context for the printer selection value returned in the hDC property of the dialog box.
    Public Const cdlPDReturnDefault = &H400     'Returns default printer name.
    Public Const cdlPDReturnIC = &H200          'Returns an information context for the printer selection value returned in the hDC property of the dialog box.
    Public Const cdlPDSelection = &H1           'Returns or sets the state of the Selection option button.
    Public Const cdlPDHelpButton = &H800        'Dialog box displays the Help button.
    Public Const cdlPDUseDevModeCopies = &H40000 'Sets support for multiple copies action; depends upon whether or not printer supports multiple copies.










