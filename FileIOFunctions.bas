Attribute VB_Name = "FileIOFunctions"
Option Explicit

Private Const LANG_FILE_ERROR_STATEMENT_ID_OFFSET = 20000
Private Const LANG_FILE_CAUTION_STATEMENT_ID_OFFSET = 22000

Private Const CAP_FLOW_FILE_VERSIONTWO = 2

Private Function BackupFile(strFilePath As String) As String
    ' Creates a backup of the file given by strFilePath by copying the file
    '  to a new file, wherein the file's extension has been replaced with .Bak
    ' Returns the path of the backup file if success; otherwise, returns ""
    
    Dim strBackupPath As String
    
On Error GoTo BackupFileErrorHandler
    
    If gBlnWriteFilesOnDrive Then
        If FileExists(strFilePath) Then
            strBackupPath = FileExtensionForce(strFilePath, "bak", True)
            
            FileCopy strFilePath, strBackupPath
            BackupFile = strFilePath
        Else
            BackupFile = ""
        End If
    Else
        BackupFile = ""
    End If
    
    Exit Function

BackupFileErrorHandler:
    Debug.Assert False
    BackupFile = ""
    
End Function

Public Function BuildPath(strParentDirectory As String, strFileName As String) As String
    Dim fso As New FileSystemObject
    
    BuildPath = fso.BuildPath(strParentDirectory, strFileName)
    
    Set fso = Nothing
    
End Function

Private Function CheckBoxToIntegerString(chkThisCheckBox As CheckBox) As String
    CheckBoxToIntegerString = Trim(Str(Val(chkThisCheckBox.value)))
End Function

Public Function ConstructFileDialogFilterMask(strFileTypeDescription As String, strFileExtension As String) As String
    ' Returns a properly formatted mask string for the Open or Save common dialog
    ' For example: "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    Dim strMask As String, strAllFiles As String
    
    strAllFiles = LookupMessage(1500)
    strMask = strFileTypeDescription & " (*." & strFileExtension & ")|*." & strFileExtension & "|" & strAllFiles & " (*.*)|*.*"
    ConstructFileDialogFilterMask = strMask
    
End Function

Public Function IsComment(ByVal strTestString As String) As Boolean
    ' Returns True if strTestString starts with ; or '
    
    strTestString = Trim(strTestString)
    If Left(strTestString, 1) = COMMENT_CHAR Or Left(strTestString, 1) = "'" Then
        IsComment = True
    Else
        IsComment = False
    End If
End Function

Public Sub LoadAbbreviations(blnResetToDefaultAbbreviations As Boolean)
    ' blnResetToDefaultAbbreviations = False will load abbreviations from disk
    ' blnResetToDefaultAbbreviations = True will reset abbreviations to default and update the file on disk
    
    Dim intIndex As Integer, intAbbrevFound As Integer
    Dim blnFileNotFound As Boolean, blnRecreateFile As Boolean
    Dim strWork As String
    Dim strFilePath As String, strBackupFilePath  As String
    Dim InFileNum As Integer
    
    Dim lngErrorID As Long, lngAbbreviationID As Long
    Dim intInvalidAbbreviationCount As Integer
    Dim strSymbol As String, strFormula As String, strOneLetterSymbol As String, strComment As String
    Dim blnIsAminoAcid As Boolean, blnInvalidSymbolOrFormula As Boolean
    Dim sngCharge As Single

On Error GoTo LoadAbbreviationsErrorHandler
    
    If Not gBlnAccessFilesOnDrive Then
        If blnResetToDefaultAbbreviations Then
            ' Load the abbreviations from memory
            objMwtWin.ResetAbbreviations
        End If
        Exit Sub
    End If
    
    ' Load Abbreviations
    AddToIntro LookupLanguageCaption(3800, "Loading Abbreviations") & " ...", False, False

    strFilePath = BuildPath(gCurrentPath, ABBREVIATIONS_FILENAME)
    blnFileNotFound = Not FileExists(strFilePath)

    If blnFileNotFound Or blnResetToDefaultAbbreviations Then
        
        ' Set the default abbreviations
        objMwtWin.ResetAbbreviations
        
        If gBlnWriteFilesOnDrive Then
            ' Re-create the abbreviations file
            SaveAbbreviations True, False, strBackupFilePath
            If Len(strBackupFilePath) > 0 Then
                AddToIntro LookupMessage(110) & "  " & LookupMessage(115) & ": " & strBackupFilePath
            Else
                AddToIntro LookupMessage(110)
            End If
        End If
    Else
    
        ' Load from disk
        InFileNum = FreeFile()
        Open strFilePath For Input As #InFileNum

        ' Read the first line and make sure it's a valid version (5.0 or greater)
        Line Input #InFileNum, strWork
        
        intIndex = InStr(strWork, "(v")
        If intIndex = 0 Then
            ' Missing version number, so re-create file
            blnRecreateFile = True
        Else
            If Val(Mid(strWork, intIndex + 2)) < 5 Then
                ' Version is before 5.0, re-create file
                blnRecreateFile = True
            End If
        End If

        If blnRecreateFile Then
            Close InFileNum
            ' Set the default abbreviations
            objMwtWin.ResetAbbreviations
            SaveAbbreviations
        Else
            
            objMwtWin.RemoveAllAbbreviations
            
            Do
                Line Input #InFileNum, strWork
                strWork = Trim(strWork)
    
                If strWork <> "" And Not IsComment(strWork) Then
                    Select Case intAbbrevFound
                    Case 0
                        If Left(strWork, 13) = "[AMINO ACIDS]" Then intAbbrevFound = 1
                    Case 1
                        If Left(strWork, 15) = "[ABBREVIATIONS]" Then
                            intAbbrevFound = 2
                        Else
                            LoadAbbreviationsParse strWork, True
                        End If
                    Case Else
                        Debug.Assert intAbbrevFound = 2
                        LoadAbbreviationsParse strWork, False
                    End Select
                End If
            Loop Until EOF(InFileNum)
        
            Select Case intAbbrevFound
            Case 0
                ' Amino Acids not found
                AddToIntro LookupMessage(120)
            Case 1
                ' Abbreviations not found
                AddToIntro LookupMessage(130)
                AddToIntro LookupMessage(135)
            Case Else
                ' Everything is fine
            End Select
        
            Close InFileNum
        End If
    End If
    
    ' Validate all of the abbreviations
    intInvalidAbbreviationCount = objMwtWin.ValidateAllAbbreviations()
    If intInvalidAbbreviationCount > 0 Then
        For lngAbbreviationID = 1 To objMwtWin.GetAbbreviationCount
            lngErrorID = objMwtWin.GetAbbreviation(lngAbbreviationID, strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment, blnInvalidSymbolOrFormula)
            If blnInvalidSymbolOrFormula Then
                AddToIntro LookupMessage(160) & ": " & strSymbol & "  " & strFormula
            End If
        Next lngAbbreviationID
    End If
    
    Exit Sub

LoadAbbreviationsErrorHandler:
    Close InFileNum
    AddToIntro LookupMessage(150) & " (" & strFilePath & "): " & Err.Description

    ' Set the default abbreviations
    objMwtWin.ResetAbbreviations
    
End Sub

Private Sub LoadAbbreviationsParse(strWork As String, blnAminoAcidAbbreviation As Boolean)
    
    Const MAX_PARSE_VALS = 4
    Dim strParsedVals(MAX_PARSE_VALS) As String, strRemaining As String
    Dim strComment As String
    Dim intParseCount As Integer
    Dim intParsedValIndex As Integer
    
    Dim strAbbrevSymbol As String, strThisAbbrevData As String
    Dim strFormula As String, strOneLetterSymbol As String
    Dim sngCharge As Single
    Dim lngErrorID As Long
    
    strWork = FormatForLocale(strWork)
    
    ' Look for a comment at the end of strWork and store in strComment
    strComment = StripComment(strWork)
    
    intParseCount = ParseString(strWork, strParsedVals(), MAX_PARSE_VALS, " ", strRemaining, True, True)
    
    For intParsedValIndex = 1 To MAX_PARSE_VALS
        strParsedVals(intParsedValIndex) = Trim(strParsedVals(intParsedValIndex))
    Next intParsedValIndex
    
    ' Make sure strParsedVals(2) contains useful information (i.e. doesn't start with ' or ; and isn't blank)
    If intParseCount > 0 And Len(strParsedVals(2)) > 0 Then
        
        strAbbrevSymbol = strParsedVals(1)
        If Len(strAbbrevSymbol) > 6 Then
            strAbbrevSymbol = Left(strAbbrevSymbol, 6)
            AddToIntro LookupMessage(190) & ": " & strWork
        End If
            
        For intParsedValIndex = 1 To 3
            ' strParsedVals(2) contains the formula
            ' strParsedVals(3) contains the charge
            ' strParsedVals(4) contains the 1 letter abbreviation (for amino acids only, and not all have one)
            strThisAbbrevData = strParsedVals(intParsedValIndex + 1)
                
            If Not IsComment(strThisAbbrevData) Then
                Select Case intParsedValIndex
                Case 1
                    ' Formula
                    strFormula = strThisAbbrevData
                Case 2
                    ' Charge
                    sngCharge = CSngSafe(strThisAbbrevData)
                Case 3
                    ' One letter abbreviation
                    If blnAminoAcidAbbreviation Then
                        ' Single letter abbreviation for amino acids
                        ' Limit to just 1 letter
                        strOneLetterSymbol = UCase(Left(strThisAbbrevData, 1))
                    End If
                End Select
            End If
        Next intParsedValIndex
        
        ' Note: Passing False to blnValidateFormula so that all abbreviations are added
        '       We later call .ValidateAllAbbreviations to validate them
        lngErrorID = objMwtWin.SetAbbreviation(strAbbrevSymbol, strFormula, sngCharge, blnAminoAcidAbbreviation, strOneLetterSymbol, strComment, False)
        
        If lngErrorID <> 0 Then
            ' Ignore the error for now; we'll validate all of the abbreviations later
            ' Reset objMwtWin.ErrorID
            objMwtWin.ClearError
        End If
    Else
        AddToIntro LookupMessage(200) & ": " & strWork
    End If

End Sub

Public Sub LoadCapillaryFlowInfo()
    ' Loads Capillary Flow values from an Info file
    
    Dim strInfoFilePath As String, strMessage As String
    Dim blnMatched As Boolean
    Dim strLineIn As String, intEqualLoc As Integer
    Dim strSettingInFile As String, strIDStringInFile As String
    Dim InFileNum As Integer
    Dim intCapillaryFlowFileFormatVersion As Integer
    
    ' 1550 = Capillary Flow Info Files, 1555 = .cap
    strInfoFilePath = SelectFile(frmCapillaryCalcs.hwnd, "Select File", gLastFileOpenSaveFolder, False, "", ConstructFileDialogFilterMask(LookupMessage(1550), LookupMessage(1555)), 1, True)
    If Len(strInfoFilePath) = 0 Then
        ' No file selected (or other error)
        Exit Sub
    End If
    
On Error GoTo LoadCapillaryFlowInfoErrorHandler
    
    ' Open the file for input
    InFileNum = FreeFile()
    Open strInfoFilePath For Input As #InFileNum
    
    Do While Not EOF(InFileNum)
        Line Input #InFileNum, strLineIn
        
        If Len(strLineIn) > 0 Then
            If Not IsComment(strLineIn) Then
                
                intEqualLoc = InStr(strLineIn, "=")
                If intEqualLoc > 0 Then
                    strIDStringInFile = UCase(Left(strLineIn, intEqualLoc - 1))
                    strSettingInFile = Mid(strLineIn, intEqualLoc + 1)
                    
                    If UCase(Left(strLineIn, 13)) = "CAPILLARYFLOW" Then
                        blnMatched = ParseCapillaryFlowSetting(strIDStringInFile, strSettingInFile, intCapillaryFlowFileFormatVersion)
                    Else
                        blnMatched = False
                    End If

                    If Not blnMatched Then
                        ' Not matched, error
                        ' Stop in IDE but ignore when compiled
                        Debug.Assert False
                    End If
                    
                End If
            End If
        End If
    Loop
    
    Close InFileNum
    Exit Sub
    
LoadCapillaryFlowInfoErrorHandler:
    Close InFileNum
    strMessage = LookupMessage(330) & ": " & strInfoFilePath
    strMessage = strMessage & vbCrLf & Err.Description
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)

End Sub

Public Sub LoadDefaultOptions(blnResetToDefaults As Boolean, Optional blnShowDebugPrompts As Boolean = False)
    ' blnResetToDefaults = False will load defaults from disk
    ' blnResetToDefaults = True will reset defaults to default
    
    Dim intIndex As Integer, intCharLoc As Integer
    Dim strWork As String, strMessage As String, strWorkVal As String, intWorkVal As Integer
    Dim blnFileNotFound As Boolean
    Dim strNewFontName As String
    Dim intNewFontSize As Integer
    Dim intSeparatorLoc As Integer, intSavedMaxAllowableIndex As Integer
    Dim strFilePath As String
    Dim InFileNum As Integer
    Dim eResponse As VbMsgBoxResult
    
On Error GoTo LoadDefaultOptionsErrorHandler
    
    If Not gBlnAccessFilesOnDrive Then
        ' Load the options from memory
        SetDefaultOptions
        SetAllTooltips
        Exit Sub
    End If
    
    strFilePath = BuildPath(gCurrentPath, INI_FILENAME)
    blnFileNotFound = Not FileExists(strFilePath)

    If blnFileNotFound Or blnResetToDefaults Then
        
        ' Set the default options
        SetDefaultOptions
        SetAllTooltips
        
        If gBlnWriteFilesOnDrive Then
            If blnShowDebugPrompts Then MsgBox "LoadDefaultOptions: Re-creating the options file (" & strFilePath & ")"
            SaveDefaultOptions
        End If
    
    Else
    
        ' Load from disk
        InFileNum = FreeFile()
        
        If blnShowDebugPrompts Then MsgBox "LoadDefaultOptions: Reading options file (" & strFilePath & ")"
        Open strFilePath For Input As #InFileNum

        Do
            Line Input #InFileNum, strWork
            strWork = Trim(strWork)

            If strWork <> "" And Not IsComment(strWork) Then
                intCharLoc = InStr(strWork, "=")
                If intCharLoc > 0 Then
                    strWorkVal = Mid(strWork, intCharLoc + 1)
                    intWorkVal = CIntSafe(strWorkVal)
                    Select Case UCase(Left(strWork, intCharLoc - 1))
                    Case "VIEW":
                        If intWorkVal = vmdSingleView Then
                            frmMain.SetViewMode vmdSingleView
                        Else
                            frmMain.SetViewMode vmdMultiView
                        End If
                    Case "CONVERT"
                        If intWorkVal >= 0 And intWorkVal <= 2 Then
                            frmProgramPreferences.optConvertType(intWorkVal).value = True
                        End If
                    Case "ABBREV":
                        If intWorkVal >= 0 And intWorkVal <= 2 Then
                            frmProgramPreferences.optAbbrevType(intWorkVal).value = True
                        End If
                    Case "STDDEV":
                        If intWorkVal >= 0 And intWorkVal <= 3 Then
                            frmProgramPreferences.optStdDevType(intWorkVal).value = True
                        End If
                    Case "CAUTION": SetCheckBoxValue frmProgramPreferences.chkShowCaution, intWorkVal
                    Case "ADVANCE": SetCheckBoxValue frmProgramPreferences.chkAdvanceOnCalculate, intWorkVal
                    Case "CHARGE": SetCheckBoxValue frmProgramPreferences.chkComputeCharge, intWorkVal
                    Case "QUICKSWITCH": SetCheckBoxValue frmProgramPreferences.chkShowQuickSwitch, intWorkVal
                    Case "FONT":
                        ' Also get fontsize before reformatting the objects
                        strNewFontName = strWorkVal
                    Case "FONTSIZE":
                        If intWorkVal >= 7 And intWorkVal <= 64 Then
                            intNewFontSize = intWorkVal
                        Else
                            intNewFontSize = 10
                        End If
                    Case "EXITCONFIRM":
                        If intWorkVal >= 0 And intWorkVal <= 3 Then
                            frmProgramPreferences.optExitConfirmation(intWorkVal).value = True
                        End If
                    Case "FINDERWEIGHTMODEWARN":
                        If intWorkVal >= -1 And intWorkVal <= 1 Then
                            With frmProgramPreferences
                                Select Case intWorkVal
                                Case 1
                                    .chkAlwaysSwitchToIsotopic.value = vbChecked
                                    ' This also checks never show
                                Case -1
                                    .chkAlwaysSwitchToIsotopic.value = vbUnchecked
                                    .chkNeverShowFormulaFinderWarning.value = vbChecked
                                Case Else
                                    .chkAlwaysSwitchToIsotopic.value = vbUnchecked
                                    .chkNeverShowFormulaFinderWarning.value = vbUnchecked
                                End Select
                            End With
                        End If
                    Case "TOOLTIPS"
                        SetCheckBoxValue frmProgramPreferences.chkShowToolTips, intWorkVal
                    Case "HIDEINACTIVEFORMS"
                        SetCheckBoxValue frmProgramPreferences.chkHideInactiveForms, intWorkVal
                    Case "STARTUPMODULE"
                        With frmProgramPreferences.cboStartupModule
                            If intWorkVal < .ListCount Then
                                .ListIndex = intWorkVal
                            End If
                        End With
                    Case "AUTOSAVEVALUES":
                        SetCheckBoxValue frmProgramPreferences.chkAutosaveValues, intWorkVal
                    Case "BRACKETSASPARENTHESES"
                        SetCheckBoxValue frmProgramPreferences.chkBracketsAsParentheses, intWorkVal
                    Case "AUTOCOPYCURRENTMWT"
                        SetCheckBoxValue frmProgramPreferences.chkAutoCopyCurrentMWT, intWorkVal
                    Case "MAXIMUMFORMULASTOSHOW"
                        ' Note that the cboMaximumFormulasToShow combo box is initialized to the allowable
                        ' values for the current resolution when frmProgramPreferences is loaded

                        ' Parse the two values stored on this line
                        intSeparatorLoc = InStr(strWorkVal, "::")
                        If intSeparatorLoc > 0 Then
                            ' intWorkVal holds the user's desired max formula index
                            intWorkVal = CIntSafe(Left(strWorkVal, intSeparatorLoc - 1))
                            intSavedMaxAllowableIndex = CIntSafe(Mid(strWorkVal, intSeparatorLoc + 2))
                            With frmProgramPreferences
                                If intSavedMaxAllowableIndex = CIntSafeDbl(.cboMaximumFormulasToShow.List(.cboMaximumFormulasToShow.ListCount - 1)) - 1 Then
                                    ' Only use the saved desired max formula index if the screen resolution
                                    '  has not changed since the program was last exited
                                    ' This is done so that the user will realize that more formulas can be displayed at higher resolutions
                                    For intIndex = 0 To .cboMaximumFormulasToShow.ListCount - 1
                                        If .cboMaximumFormulasToShow.List(intIndex) = Trim(Str(intWorkVal)) + 1 Then
                                            .cboMaximumFormulasToShow.ListIndex = intIndex
                                            If frmMain.GetTopFormulaIndex <= intWorkVal Then
                                                gMaxFormulaIndex = intWorkVal
                                            End If
                                            Exit For
                                        End If
                                    Next intIndex
                                End If
                            End With
                        End If
                    Case "FINDERBOUNDEDSEARCH"
                        If intWorkVal = 0 Or intWorkVal = 1 Then
                            frmFinderOptions.cboSearchType.ListIndex = intWorkVal
                        End If
                    Case "LANGUAGE"
                        gCurrentLanguage = strWorkVal
                    Case "LANGUAGEFILE"
                        gCurrentLanguageFileName = strWorkVal
                    Case "LASTOPENSAVEFOLDER"
                          gLastFileOpenSaveFolder = strWorkVal
                    Case Else
                        ' Not matched, error
                        ' Stop in IDE but ignore when compiled
                        Debug.Assert False
                    End Select
                Else
                    ' Not matched, error
                    ' Stop in IDE but ignore when compiled
                    Debug.Assert False
                End If
            End If
        Loop Until EOF(InFileNum)
        Close InFileNum
        
        If blnShowDebugPrompts Then MsgBox "LoadDefaultOptions: Set Fonts"
        
        ' Tasks that need to be done now that the options have been loaded
        SetFonts strNewFontName, intNewFontSize
        
        If blnShowDebugPrompts Then
            eResponse = MsgBox("Show detailed debugging information when setting ToolTips?", vbQuestion + vbYesNo + vbDefaultButton2)
            SetAllTooltips (eResponse = vbYes)
        Else
            SetAllTooltips
        End If

    End If

    Exit Sub

LoadDefaultOptionsErrorHandler:
    Close InFileNum
    
    strMessage = LookupMessage(400) & " (" & strFilePath & "): " & Err.Description
    strMessage = strMessage & vbCrLf & LookupMessage(410) & vbCrLf & LookupMessage(345)
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)

End Sub

Public Sub LoadElements(intNewElementMode As Integer, Optional blnShowFrmIntro As Boolean = True)
    ' intNewElementMode = 0 will load elements from disk
    ' intNewElementMode = 1 will reset elements to default (average weights)
    ' intNewElementMode = 2 will change elements to isotopic weights
    ' intNewElementMode = 3 will change the elements to their integer weights

    ' If loading elements from memory, re-creates the element file
    ' Otherwise, loads from disk and updates the values in objMwtWin
    
    Const MAX_PARSE_COUNT = 4
    Dim intParseCount As Integer, intParsedValIndex As Integer
    Dim strParsedVals(MAX_PARSE_COUNT) As String        ' 0-based array
    
    Dim intCharLoc As Integer
    Dim blnElementsHeaderFound As Boolean, blnFileNotFound As Boolean, blnRecreateFile As Boolean
    Dim InFileNum As Integer
    Dim eNewElementWeightType As emElementModeConstants
    
    Dim strLineIn As String, strRemaining As String
    Dim strSymbol As String
    Dim lngElementID As Long, lngErrorID As Long
    Dim dblMass As Double, dblUncertainty As Double, sngCharge As Single, intIsotopeCount As Integer
    Dim dblNewMass As Double, dblNewUncertainty As Double, sngNewCharge As Single
    
    Dim strFilePath As String, strBackupFilePath As String

On Error GoTo LoadElementsErrorHandler

    If Not gBlnAccessFilesOnDrive Then
        ' Load the element weights from memory
        SwitchWeightModeInteger intNewElementMode
        Exit Sub
    End If
    
    ' Load Elements
    If blnShowFrmIntro Then frmIntro.Show vbModeless
    frmIntro.lblLoadStatus.Caption = LookupLanguageCaption(3810, "Loading Elements") & " ..."
    
    strFilePath = BuildPath(gCurrentPath, ELEMENTS_FILENAME)
    blnFileNotFound = Not FileExists(strFilePath)

    If blnFileNotFound Or intNewElementMode >= 1 Then
        If blnFileNotFound Then
            AddToIntro LookupMessage(270) & " (" & strFilePath & ")"
        End If
        
        SwitchWeightModeInteger intNewElementMode
    
        If gBlnWriteFilesOnDrive Then
            SaveElements strBackupFilePath
            If Len(strBackupFilePath) > 0 Then
                AddToIntro LookupMessage(210) & "  " & LookupMessage(115) & ": " & strBackupFilePath
            Else
                AddToIntro LookupMessage(210)
            End If
        End If
    Else
        
        ' Load from disk
        InFileNum = FreeFile()
        Open strFilePath For Input As #InFileNum

        ' Read the first line and make sure it's a valid version (5.0 or greater)
        Line Input #InFileNum, strLineIn
        
        intCharLoc = InStr(strLineIn, "(v")
        If intCharLoc = 0 Then
            ' Missing version number, so re-create file
            blnRecreateFile = True
        Else
            If Val(Mid(strLineIn, intCharLoc + 2)) < 5 Then
                ' Version is before 5.0, re-create file
                blnRecreateFile = True
            End If
        End If

        If blnRecreateFile Then
            Close InFileNum
            SwitchWeightModeInteger intNewElementMode
            SaveElements
        Else
        
            blnElementsHeaderFound = False
            Do
                Line Input #InFileNum, strLineIn
                strLineIn = Trim(strLineIn)
                strRemaining = ""
    
                If strLineIn <> "" And Not IsComment(strLineIn) Then
                    If Not blnElementsHeaderFound Then
                        If Left(strLineIn, 19) = "[ELEMENTWEIGHTTYPE]" Then
                            eNewElementWeightType = CIntSafe(Trim(Mid(strLineIn, 20)))
                            If eNewElementWeightType < 1 Or eNewElementWeightType > 3 Then
                                eNewElementWeightType = 1
                            End If
                            SwitchWeightMode eNewElementWeightType
                            gElementWeightTypeInFile = eNewElementWeightType
                        End If
                        If Left(strLineIn, 13) = "[ELEMENTS]" Then
                            blnElementsHeaderFound = True
                            If eNewElementWeightType = 0 Then
                                ' No gElementWeightType statement present, assume type 1
                                eNewElementWeightType = 1
                                SwitchWeightMode eNewElementWeightType
                                gElementWeightTypeInFile = eNewElementWeightType
                            End If
                        End If
                    Else
                        strLineIn = FormatForLocale(strLineIn)
                        
                        ' Note: by using a delimeter of " ;'" and setting MatchWholeDelimeter to false, then strLineIn will be split based on a space, semicolon, or apostrophe
                        intParseCount = ParseString(strLineIn, strParsedVals(), MAX_PARSE_COUNT, " ;'", strRemaining, False, True, False)
                        
                        If intParseCount >= 3 Then
                            For intParsedValIndex = 0 To MAX_PARSE_COUNT - 1
                                strParsedVals(intParsedValIndex) = Trim(strParsedVals(intParsedValIndex))
                            Next intParsedValIndex
                            
                            If IsNumeric(strParsedVals(1)) Then
                                strSymbol = strParsedVals(0)
                                
                                dblNewMass = CDblSafe(strParsedVals(1))
                                If IsNumeric(strParsedVals(3)) Then
                                    dblNewUncertainty = CDblSafe(strParsedVals(2))
                                Else
                                    dblNewUncertainty = 0
                                End If
                                
                                If IsNumeric(strParsedVals(3)) Then
                                    sngNewCharge = CSngSafe(strParsedVals(3))
                                Else
                                    sngNewCharge = 0
                                End If
                                
                                If dblNewMass >= 0 And dblNewUncertainty >= 0 Then
                                    ' Make sure strSymbol is valid, and grab the current values for the element
                                    lngElementID = objMwtWin.GetElementID(strSymbol)
                                    
                                    If lngElementID > 0 Then
                                        ' Get the current element values
                                        lngErrorID = objMwtWin.GetElement(lngElementID, strSymbol, dblMass, dblUncertainty, sngCharge, intIsotopeCount)
                                        Debug.Assert lngErrorID = 0
                                        
                                        ' See if new mass is more than 20% different than old mass
                                        If dblNewMass > 1.2 * dblMass Or _
                                           dblNewMass < 0.8 * dblMass Then
                                            AddToIntro LookupMessage(220, ": " & strSymbol & ", " & CStr(dblNewMass))
                                        End If
                                        
                                        ' See if uncertainty is more than 10 times different than old uncertainty
                                        If gElementWeightTypeInFile = emAverageMass And _
                                           (dblNewUncertainty > 10 * dblUncertainty Or _
                                            dblNewUncertainty < 0.1 * dblUncertainty) Then
                                            AddToIntro LookupMessage(230, ": " & strSymbol & ", " & CStr(dblNewUncertainty))
                                        End If
                                        
                                        lngErrorID = objMwtWin.SetElement(strSymbol, dblNewMass, dblNewUncertainty, sngNewCharge, False)
                                        
                                        If lngErrorID <> 0 Then
                                            AddToIntro LookupMessage(lngErrorID) & ": " & strLineIn
                                        End If
                                    Else
                                         AddToIntro LookupMessage(250) & ": " & strLineIn
                                    End If
                                Else
                                    AddToIntro LookupMessage(200) & ": " & strLineIn
                                End If
                            Else
                                AddToIntro LookupMessage(200) & ": " & strLineIn
                            End If
                        Else
                            AddToIntro LookupMessage(200) & ": " & strLineIn
                        End If
                    
                    End If
                End If
            Loop Until EOF(InFileNum)
    
            If Not blnElementsHeaderFound Then
                ' Elements not found
                AddToIntro LookupMessage(260)
                AddToIntro LookupMessage(265)
            End If
    
            Close InFileNum
        End If
    End If
 
LoadElementsExit:

    ' Recompute the abbreviation masses
    objMwtWin.RecomputeAbbreviationMasses
    
    ' Make sure QuickSwitch Element Mode value is correct
    frmMain.ShowHideQuickSwitch frmProgramPreferences.chkShowQuickSwitch.value

    Exit Sub

LoadElementsErrorHandler:
    Close InFileNum
    AddToIntro LookupMessage(280) & " (" & strFilePath & ")"
    AddToIntro Err.Description
    AddToIntro LookupMessage(265)
    Resume LoadElementsExit
    
End Sub

Public Function LoadLanguageSettings(strLangFilename As String, strNewLanguage As String) As Boolean
    Dim strFilePath As String
    Dim blnSuccess As Boolean
    Dim strSearchForFile As String, strMessage As String
        
    ' See if the language file exists
    strFilePath = BuildPath(gCurrentPath, strLangFilename)
    
    strSearchForFile = Dir(strFilePath)
    If Len(strSearchForFile) > 0 Then
        ' Load the new language file into form frmStrings
        blnSuccess = LoadLanguageFile(strFilePath, frmStrings.grdLanguageStrings, frmStrings.grdLanguageStringsCrossRef, True)
        
        If blnSuccess Then
            ' Reset menu captions to numeric values
            ResetMenuCaptions False
            
            ' Load the captions into controls on all forms
            LoadLanguageCaptions
            
            ' Load the captions into the dynamic text fields on the Formula Finder form
            frmFinder.LoadDynamicTextCaptions
            
            ' Add shortcut keys to menus
            AppendShortcutKeysToMenuCaptions
            
            ' Save new language in gCurrentLanguage
            gCurrentLanguage = strNewLanguage
            gCurrentLanguageFileName = strLangFilename
            
            ' Save new value for gMWAbbreviation
            gMWAbbreviation = LookupLanguageCaption(4040, "MW")
            If Len(gMWAbbreviation) <> 2 Then
                gMWAbbreviation = "MW"
            End If
            
            blnSuccess = True
        Else
            ' Problem loading settings
            strMessage = LookupMessage(440) & " (" & strFilePath & ")"
            strMessage = strMessage & vbCrLf & LookupMessage(450)
            MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)
            blnSuccess = False
        End If
    Else
        strMessage = LookupMessage(460) & " (" & strFilePath & ")"
        strMessage = strMessage & vbCrLf & LookupMessage(450)
        MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)
        blnSuccess = False
    End If
    
    LoadLanguageSettings = blnSuccess
    
End Function

Private Sub LoadGridColumnTitles(grdThisFlexGrid As MSFlexGrid)
    
    Dim intIndex As Integer
    
On Error GoTo LoadGridColumnTitlesErrorHandler

    ' Need to update Column titles in various MSFlexGrids in program
    With grdThisFlexGrid
        Select Case LCase(.Name)
        Case "grdamino"
            .TextMatrix(0, 0) = "AbbrevID (Hidden)"
            .TextMatrix(0, 1) = LookupLanguageCaption(9180, .TextMatrix(0, 0))
            .TextMatrix(0, 2) = LookupLanguageCaption(9160, .TextMatrix(0, 1))
            .TextMatrix(0, 3) = LookupLanguageCaption(9150, .TextMatrix(0, 2))
            .TextMatrix(0, 4) = LookupLanguageCaption(9190, .TextMatrix(0, 3))
            .TextMatrix(0, 5) = LookupLanguageCaption(9195, "Comment")
        Case "grdnormal"
            .TextMatrix(0, 0) = "AbbrevID (Hidden)"
            .TextMatrix(0, 1) = LookupLanguageCaption(9170, .TextMatrix(0, 0))
            .TextMatrix(0, 2) = LookupLanguageCaption(9160, .TextMatrix(0, 1))
            .TextMatrix(0, 3) = LookupLanguageCaption(9150, .TextMatrix(0, 2))
            .TextMatrix(0, 4) = LookupLanguageCaption(9195, "Comment")
        Case "grdelem"
            .TextMatrix(0, 0) = LookupLanguageCaption(9350, .TextMatrix(0, 0))
            .TextMatrix(0, 1) = LookupLanguageCaption(9360, .TextMatrix(0, 1))
            .TextMatrix(0, 2) = LookupLanguageCaption(9370, .TextMatrix(0, 2))
            .TextMatrix(0, 3) = LookupLanguageCaption(9150, .TextMatrix(0, 3))
        Case "grdmodsymbols"
            .TextMatrix(0, 0) = "ModSymbolID (Hidden)"
            .TextMatrix(0, 1) = LookupLanguageCaption(15610, "Symbol")
            .TextMatrix(0, 2) = LookupLanguageCaption(15620, "Mass")
            .TextMatrix(0, 3) = LookupLanguageCaption(15640, "Comment")
        Case "grdionlist"
            .TextMatrix(0, 0) = LookupLanguageCaption(12550, "Mass")
            .TextMatrix(0, 1) = LookupLanguageCaption(12560, "Intensity")
            .TextMatrix(0, 2) = LookupLanguageCaption(12570, "Symbol")
        Case "grdfragmasses"
            .TextMatrix(0, 0) = LookupLanguageCaption(12500, "#")
            .TextMatrix(0, 1) = LookupLanguageCaption(12510, "Immon.")
            ' We can only update the Seq column if the initial language is English
            For intIndex = 0 To .Cols - 1
                If .TextMatrix(0, intIndex) = "Seq." Then
                    .TextMatrix(0, intIndex) = LookupLanguageCaption(12520, "Seq.")
                    Exit For
                End If
            Next intIndex
            .TextMatrix(0, .Cols - 1) = .TextMatrix(0, 0)
        Case "grdpc", "grdlanguagestrings", "grdmenuinfo", "grdlanguagestringscrossref"
            ' No column titles in grdPC or grdLanguageStrings or grdMenuInfo or grdLanguageStringsCrossRef
        Case Else
            ' Unknown Grid: do not set any titles
            ' Add cases above for the other grids so this assertion is note reached
            Debug.Assert False
        End Select
    End With

    Exit Sub
    
LoadGridColumnTitlesErrorHandler:
    Debug.Assert False
    GeneralErrorHandler "MwtWinProcedures|LoadGridColumnTitles", Err.Number, Err.Description
End Sub
    
Private Sub LoadLanguageCaptions(Optional boolSingleFormOnly As Boolean = False, Optional frmThisSingleForm As VB.Form)
    Dim frmThisForm As VB.Form
        
    If boolSingleFormOnly Then
        LoadLanguageCaptionsIntoForm frmThisSingleForm
        frmThisSingleForm.Refresh
    Else
        ' Load new captions into all forms
        For Each frmThisForm In Forms
            LoadLanguageCaptionsIntoForm frmThisForm
            frmThisForm.Refresh
        Next
                
        LoadLanguageCaptionsIntoAppWideDynamicLabels
        
        frmEditElem.PositionFormControls
        frmEditAbbrev.PositionFormControls
    End If

    If frmMain.mnuShowTips.Checked = True Then
        ' Load all ToolTips
        SetAllTooltips
    End If

    Dim strMinutesElapsedRemaining As String, strClickToPause As String
    Dim strPreparingToPause As String, strPaused As String
    Dim strResuming As String, strPressEscapeToAbort As String
    
    ' Update language captions for frmProgress and objMwtWin.frmProgress
    strMinutesElapsedRemaining = LookupLanguageCaption(14740, "min. elapsed/remaining")
    strClickToPause = LookupLanguageCaption(14710, "Click to Pause")
    strPreparingToPause = LookupLanguageCaption(14715, "Preparing to Pause")
    strPaused = LookupLanguageCaption(14720, "Paused")
    strResuming = LookupLanguageCaption(14725, "Resuming")
    strPressEscapeToAbort = LookupLanguageCaption(14730, "(Press Escape to Abort)")
        
    frmProgress.SetStandardCaptionText strMinutesElapsedRemaining, strPreparingToPause, strResuming, strClickToPause, strPaused, strPressEscapeToAbort
    
    objMwtWin.SetStandardProgressCaptionText strMinutesElapsedRemaining, strPreparingToPause, strResuming, strClickToPause, strPaused, strPressEscapeToAbort

End Sub

Private Sub LoadLanguageCaptionsIntoAppWideDynamicLabels()
    Dim intIndex As Integer
    
    ' Need to update Formula 1, Formula 2, etc. in frmMain since they are
    '   added dynamically and do not have a LanguageID# in their .Tag
    With frmMain
        For intIndex = 0 To frmMain.GetTopFormulaIndex
            .lblFormula(intIndex).Caption = ConstructFormulaLabel(intIndex)
        Next intIndex
    End With

End Sub

Private Sub LoadLanguageCaptionsIntoForm(frmThisForm As VB.Form)
    Dim ctlThisControl As Control, strControlType As String
    Dim boolComboBox As Boolean
    
    ' Load the caption for the form
    frmThisForm.Caption = LookupLanguageCaption(frmThisForm.Tag, frmThisForm.Caption)

    ' Load captions for each control on form (if appropriate)
    ' Note that ToolTips are loaded separately using SetAllTooltips
    '  Attempting to set ToolTips using this Sub results in very high processor usage
    For Each ctlThisControl In frmThisForm.Controls
        boolComboBox = False
        With ctlThisControl
            strControlType = TypeName(ctlThisControl)
            Select Case strControlType
            Case "Label"
                .Caption = LookupLanguageCaption(.Tag, .Caption)
            Case "Menu"
                .Caption = LookupLanguageCaption(.Caption, .Caption, True, .Name)
            Case "TabStrip"
                ' Not implemented: use .Tag to set .Caption and .ToolTipText to set .ToolTipText
            Case "Toolbar"
                ' Not implemented: use .ToolTipText to set .ToolTipText
            Case "ListView"
                ' Not implemented: use .Tag to set .Text
            Case "CommonDialog"
                ' Not used
            Case "StatusBar"
                ' MsgBox "status bar fontsize=" & ctl.Font.Size
            Case "SplitFrame"
            Case "ScrollPanel"
            Case "PictureBox"
            Case "ProgressBar"
            Case "TextBox"
            Case "Timer"
            Case "Shape"
            Case "ComboBox"
                boolComboBox = True
            Case "Image"
            Case "ListBox"
            Case "Line"
            Case "MSFlexGrid"
                LoadGridColumnTitles ctlThisControl
            Case "RichTextBox"
                ' I use the .Tag property of RichTextBox controls to store useful values
                ' Do not attempt to set the caption
            Case Else
                ' Regular control (and Frames)
                .Caption = LookupLanguageCaption(.Tag, .Caption)
            End Select
        End With
        If boolComboBox Then
            ' Do not Re-populate if control is cboSortResults (.Tag = 10850)
            ' The control's list items are added dynamically depending on the checkboxes on the form
            If ctlThisControl.Tag <> "10850" Then
                PopulateComboBox ctlThisControl, True
            End If
        End If
    Next

End Sub

Private Function LoadLanguageFile(strLangFilePath As String, grdThisLanguageGrid As MSFlexGrid, grdThisLanguageGridCrossRef As MSFlexGrid, blnForeignLanguage As Boolean) As Boolean
    ' This sub assumes file strLangFilePath exists
    ' Check for this before calling
    
    Dim strWork As String, intEqualPos As Integer, lngKeyValue As Long
    Dim strNewSymbolCombo As String, strCaption As String
    Dim blnCautionFound As Boolean, intCrossRefTrack As Integer
    Dim lngError As Long
    Dim InFileNum As Integer
    
    Const CROSS_REF_TRACK_STEP = 1000
    
On Error GoTo LoadLanguageFileErrorHandler
    
    If Not FileExists(strLangFilePath) Then
        ' File not found
        AddToIntro LookupMessage(460) & " (" & strLangFilePath & ")"
        AddToIntro LookupMessage(305)
        ResetMenuCaptions True
        
        LoadLanguageFile = False
        Exit Function
    End If
    
    InFileNum = FreeFile()
    Open strLangFilePath For Input As #InFileNum

    With grdThisLanguageGrid
        .Clear
        .Rows = 0
        .ColWidth(0) = 650
        .ColWidth(1) = 8000
    End With
    
    With grdThisLanguageGridCrossRef
        .Clear
        .Rows = 0
        .ColWidth(0) = 600
        .ColWidth(1) = 600
    End With
    
    intCrossRefTrack = CROSS_REF_TRACK_STEP
    Do
        ' Read each line and store in grdThisLanguageGrid if it starts with
        '  a number followed by an = sign
        Line Input #InFileNum, strWork
        intEqualPos = InStr(strWork, "=")
        
        If intEqualPos > 0 Then
            If IsNumeric(Left(strWork, intEqualPos - 1)) Then
                lngKeyValue = CLng(Left(strWork, intEqualPos - 1))
                strCaption = Mid(strWork, intEqualPos + 1)
                
                With grdThisLanguageGrid
                    .AddItem CStr(lngKeyValue)
                    .TextMatrix(.Rows - 1, 1) = strCaption
                End With
                
                If lngKeyValue >= LANG_FILE_ERROR_STATEMENT_ID_OFFSET And _
                   lngKeyValue < LANG_FILE_CAUTION_STATEMENT_ID_OFFSET Then
                    lngError = objMwtWin.SetMessageStatement(lngKeyValue - LANG_FILE_ERROR_STATEMENT_ID_OFFSET, strCaption)
                    Debug.Assert lngError = 0
                End If
                
                If lngKeyValue >= intCrossRefTrack Then
                    Do While intCrossRefTrack + 1000 <= lngKeyValue
                        intCrossRefTrack = intCrossRefTrack + CROSS_REF_TRACK_STEP
                        With grdThisLanguageGridCrossRef
                            .AddItem intCrossRefTrack
                            .TextMatrix(.Rows - 1, 1) = grdThisLanguageGrid.Rows - 1
                        End With
                    Loop
                    With grdThisLanguageGridCrossRef
                        .AddItem intCrossRefTrack
                        .TextMatrix(.Rows - 1, 1) = grdThisLanguageGrid.Rows - 1
                    End With
                    intCrossRefTrack = intCrossRefTrack + CROSS_REF_TRACK_STEP
                End If
                
                If lngKeyValue = LANG_FILE_CAUTION_STATEMENT_ID_OFFSET And blnForeignLanguage Then
                    blnCautionFound = True
                End If
            ElseIf blnCautionFound And intEqualPos > 2 Then
                Select Case Asc(Left(strWork, 1))
                Case 97 To 122, 65 To 90    ' a to z and A to Z
                    strNewSymbolCombo = Left(strWork, intEqualPos - 1)
                    strCaption = Mid(strWork, intEqualPos + 1)
                    
                    lngError = objMwtWin.SetCautionStatement(strNewSymbolCombo, strCaption)
                    Debug.Assert lngError = 0
                End Select
                
            End If
        End If
    
    Loop Until EOF(InFileNum)
    Close InFileNum
        
    LoadLanguageFile = True
    Exit Function
    
LoadLanguageFileErrorHandler:
    Close InFileNum
    AddToIntro LookupMessage(440) & " (" & strLangFilePath & "): " & Err.Description
    AddToIntro LookupMessage(305)
    LoadLanguageFile = False
    
End Function

Public Sub LoadSequenceInfo()
    ' Loads Fragmentation Modelling sequence info from a file
    
    Dim strSequenceFilePath As String, strMessage As String
    Dim blnMatched As Boolean
    Dim strLineIn As String, intEqualLoc As Integer
    Dim strSettingInFile As String, strIDStringInFile As String
    Dim SeqFileNum As Integer
    Dim blnSkipNextLineRead As Boolean
    
    ' 1530 = Sequence Files, 1535 = .seq
    strSequenceFilePath = SelectFile(frmFragmentationModelling.hwnd, "Select File", gLastFileOpenSaveFolder, False, "", ConstructFileDialogFilterMask(LookupMessage(1530), LookupMessage(1535)), 1)
    If Len(strSequenceFilePath) = 0 Then
        ' No file selected (or other error)
        Exit Sub
    End If
    
On Error GoTo LoadSequenceInfoErrorHandler
    
    ' Open the file for input
    SeqFileNum = FreeFile()
    Open strSequenceFilePath For Input As #SeqFileNum
    
    Do While Not EOF(SeqFileNum)
        If blnSkipNextLineRead Then
            blnSkipNextLineRead = False
        Else
            Line Input #SeqFileNum, strLineIn
            strLineIn = Trim(strLineIn)
        End If
        
        If Len(strLineIn) > 0 Then
            If Not IsComment(strLineIn) Then
                
                intEqualLoc = InStr(strLineIn, "=")
                If intEqualLoc > 0 Then
                    strIDStringInFile = UCase(Left(strLineIn, intEqualLoc - 1))
                    strSettingInFile = Mid(strLineIn, intEqualLoc + 1)
                    
                    If UCase(Left(strLineIn, 9)) = "FRAGMODEL" Then
                        blnMatched = ParseFragModelSetting(SeqFileNum, strIDStringInFile, strSettingInFile)
                    Else
                        blnMatched = False
                    End If

                    If Not blnMatched Then
                        If UCase(Left(strLineIn, 7)) = "IONPLOT" Then
                            ' This is an old setting; we can ignore it
                        Else
                            ' Not matched, error
                            ' Stop in IDE but ignore when compiled
                            Debug.Assert False
                        End If
                    End If
                Else
                    If Not PossiblySkipCWSpectrumSection(strLineIn, SeqFileNum, blnSkipNextLineRead) Then
                        ' Not matched, error
                        ' Stop in IDE but ignore when compiled
                        Debug.Assert False
                    End If
                End If
            End If
        End If
    Loop
    
    Close SeqFileNum
    
    ' Load the CWSpectrum Options
    frmFragmentationModelling.LoadCWSpectrumOptions strSequenceFilePath

    ' Rematch the loaded ions in case txtAlignment is nonzero
    ' Necessary since txtAlignment's value is set after the IonMatchList is Loaded
    frmFragmentationModelling.AlignmentOffsetValidate
    
    Exit Sub
    
LoadSequenceInfoErrorHandler:
    Close SeqFileNum
    strMessage = LookupMessage(900) & ": " & strSequenceFilePath
    strMessage = strMessage & vbCrLf & Err.Description
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)

End Sub

Public Sub LoadValuesAndFormulas(blnRecreateDefaultValuesFile As Boolean)
    ' blnRecreateDefaultValuesFile = False will load values from disk
    ' blnRecreateDefaultValuesFile = True will reset values to default
            
    Dim intEqualLoc As Integer, intIndex As Integer
    Dim blnMatched As Boolean, blnFileNotFound As Boolean
    Dim strLineIn As String, strMessage As String, strSettingInFile As String
    Dim strIDStringInFile As String, strIDStringToMatch As String
    Dim dblSettingInFileValue As Double, lngSettingInFile As Long
    Dim strFilePath As String
    Dim strNewFormulaIndex As String, intNewFormulaIndex As Integer
    Dim InFileNum As Integer
    Dim intCapillaryFlowFileFormatVersion As Integer
    Dim eViewModeSaved As vmdViewModeConstants
    Dim blnSkipNextLineRead As Boolean
        
On Error GoTo LoadValuesAndFormulasErrorHandler

    If Not gBlnAccessFilesOnDrive Then
        SetDefaultValuesAndFormulas blnRecreateDefaultValuesFile
        Exit Sub
    End If
        
    strFilePath = BuildPath(gCurrentPath, VALUES_FILENAME)
    blnFileNotFound = Not FileExists(strFilePath)
    
    If blnFileNotFound Or blnRecreateDefaultValuesFile Then
        
        SetDefaultValuesAndFormulas True
        
        If gBlnWriteFilesOnDrive Then
            
            If Not blnFileNotFound Then
                BackupFile strFilePath
            End If

            ' Re-create the values file
            SaveValuesAndFormulas
            
            ' Wait 100 msec, then call frmFragmentationModelling.SaveCWSpectrumOptions
            Sleep 100
            frmFragmentationModelling.SaveCWSpectrumOptions strFilePath
        End If
    Else
        InFileNum = FreeFile()
        Open strFilePath For Input As #InFileNum

        ' Set blnDelayUpdate to True on frmMMconvert to delay auto-updating
        frmMMConvert.SetDelayUpdate True
                    
        eViewModeSaved = frmMain.GetViewMode()
        If eViewModeSaved = vmdSingleView Then frmMain.SetViewMode vmdMultiView
        
        intNewFormulaIndex = -1
        Do
            If blnSkipNextLineRead Then
                blnSkipNextLineRead = False
            Else
                Line Input #InFileNum, strLineIn
                strLineIn = Trim(strLineIn)
            End If
            
            If strLineIn <> "" And Not IsComment(strLineIn) Then
                intEqualLoc = InStr(strLineIn, "=")
                If intEqualLoc > 0 Then
                    blnMatched = False
                    strSettingInFile = Mid(strLineIn, intEqualLoc + 1)
                    strIDStringInFile = UCase(Left(strLineIn, intEqualLoc - 1))
                    If IsNumeric(strSettingInFile) Then
                        dblSettingInFileValue = CDblSafe(FormatForLocale(strSettingInFile))
                        lngSettingInFile = CLngSafe(Str(dblSettingInFileValue))
                    Else
                        dblSettingInFileValue = 0
                        lngSettingInFile = 0
                    End If
                    
                    If UCase(Left(strLineIn, 7)) = "FORMULA" Then
                        strNewFormulaIndex = Mid(strLineIn, 8, intEqualLoc - 8)
                        If IsNumeric(strNewFormulaIndex) Then
                            strSettingInFile = FormatForLocale(strSettingInFile)
                            If strSettingInFile = "" Then
                                ' Skip it if it's blank
                                blnMatched = True
                            Else
                                If intNewFormulaIndex < gMaxFormulaIndex Then
                                    intNewFormulaIndex = intNewFormulaIndex + 1
                                    If intNewFormulaIndex > frmMain.GetTopFormulaIndex Then
                                        frmMain.AddNewFormulaWrapper
                                    End If
                                    frmMain.rtfFormula(intNewFormulaIndex).Text = strSettingInFile
                                    
                                    If frmMain.GetViewMode = vmdSingleView Then
                                        ' Need to manually update the formula's formatting
                                        frmMain.UpdateAndFormatFormula intNewFormulaIndex, True
                                    End If

                                    ' Must assign value to rtfFormulaSingle to allow new formula button_click to work
                                    frmMain.rtfFormulaSingle.Text = strSettingInFile
                                    frmMain.rtfFormula(intNewFormulaIndex).Tag = FORMULA_CHANGED
                                    frmMain.rtfFormulaSingle.Tag = FORMULA_CHANGED
                                End If
                            End If
                            blnMatched = True
                        End If
                    End If
                    
                    For intIndex = 0 To 9
                        strIDStringToMatch = "FINDERMIN" & Trim(Str(intIndex))
                        If strIDStringInFile = strIDStringToMatch Then
                            frmFinder.txtMin(intIndex).Text = strSettingInFile
                            blnMatched = True
                        End If
                        strIDStringToMatch = "FINDERMAX" & Trim(Str(intIndex))
                        If strIDStringInFile = strIDStringToMatch Then
                            frmFinder.txtMax(intIndex).Text = strSettingInFile
                            blnMatched = True
                        End If
                        strIDStringToMatch = "FINDERCHECKELEMENTS" & Trim(Str(intIndex))
                        If strIDStringInFile = strIDStringToMatch Then
                            frmFinder.chkElements(intIndex).value = lngSettingInFile
                            blnMatched = True
                        End If
                        strIDStringToMatch = "FINDERPERCENTVALUE" & Trim(Str(intIndex))
                        If strIDStringInFile = strIDStringToMatch Then
                            frmFinder.txtPercent(intIndex).Text = strSettingInFile
                            blnMatched = True
                        End If
                        strIDStringToMatch = "FINDERCUSTOMWEIGHT" & Trim(Str(intIndex - 3))
                        If strIDStringInFile = strIDStringToMatch Then
                            frmFinder.txtWeight(intIndex).Text = strSettingInFile
                            blnMatched = True
                        End If
                        If blnMatched Then Exit For
                    Next intIndex
                    
                    If Not blnMatched Then
                        Select Case strIDStringInFile
                        Case "AMINOACIDCONVERTONELETTER": frmAminoAcidConverter.txt1LetterSequence = strSettingInFile
                        Case "AMINOACIDCONVERTTHREELETTER":  frmAminoAcidConverter.txt3LetterSequence = strSettingInFile
                        Case "AMINOACIDCONVERTSPACEONELETTER":  frmAminoAcidConverter.chkSpaceEvery10.value = lngSettingInFile
                        Case "AMINOACIDCONVERTDASHTHREELETTER": frmAminoAcidConverter.chkSeparateWithDash.value = lngSettingInFile
                        
                        Case "MOLE/MASSWEIGHTSOURCE": SetDualOptionGroup frmMMConvert.optWeightSource, CInt(lngSettingInFile)
                        Case "MOLE/MASSCUSTOMMASS": frmMMConvert.txtCustomMass.Text = strSettingInFile
                        Case "MOLE/MASSACTION": frmMMConvert.cboAction.ListIndex = lngSettingInFile
                        Case "MOLE/MASSFROM": frmMMConvert.txtFromNum.Text = strSettingInFile
                        Case "MOLE/MASSFROMUNITS": frmMMConvert.cboFrom.ListIndex = lngSettingInFile
                        Case "MOLE/MASSDENSITY": frmMMConvert.txtDensity.Text = strSettingInFile
                        Case "MOLE/MASSTOUNITS": frmMMConvert.cboTo.ListIndex = lngSettingInFile
                        Case "MOLE/MASSVOLUME": frmMMConvert.txtVolume.Text = strSettingInFile
                        Case "MOLE/MASSVOLUMEUNITS": frmMMConvert.cboVolume.ListIndex = lngSettingInFile
                        Case "MOLE/MASSMOLARITY": frmMMConvert.txtConcentration.Text = strSettingInFile
                        Case "MOLE/MASSMOLARITYUNITS": frmMMConvert.cboConcentration.ListIndex = lngSettingInFile
                        
                        Case "MOLE/MASSLINKCONCENTRATIONS": frmMMConvert.chkLinkMolarities = lngSettingInFile
                        Case "MOLE/MASSLINKDILUTIONVOLUMEUNITS": frmMMConvert.chkLinkDilutionVolumeUnits = lngSettingInFile
                        
                        Case "MOLE/MASSDILUTIONMODE": frmMMConvert.cboDilutionMode.ListIndex = lngSettingInFile
                        Case "MOLE/MASSMOLARITYINITIAL": frmMMConvert.txtDilutionConcentrationInitial.Text = strSettingInFile
                        Case "MOLE/MASSMOLARITYINITIALUNITS": frmMMConvert.cboDilutionConcentrationInitial.ListIndex = lngSettingInFile
                        Case "MOLE/MASSVOLUMEINITIAL": frmMMConvert.txtStockSolutionVolume.Text = strSettingInFile
                        Case "MOLE/MASSVOLUMEINITIALUNITS": frmMMConvert.cboStockSolutionVolume.ListIndex = lngSettingInFile
                        Case "MOLE/MASSMOLARITYFINAL": frmMMConvert.txtDilutionConcentrationFinal.Text = strSettingInFile
                        Case "MOLE/MASSMOLARITYFINALUNITS": frmMMConvert.cboDilutionConcentrationFinal.ListIndex = lngSettingInFile
                        Case "MOLE/MASSVOLUMESOLVENT": frmMMConvert.txtDilutingSolventVolume.Text = strSettingInFile
                        Case "MOLE/MASSVOLUMESOLVENTUNITS": frmMMConvert.cboDilutingSolventVolume.ListIndex = lngSettingInFile
                        Case "MOLE/MASSVOLUMETOTAL": frmMMConvert.txtTotalVolume.Text = strSettingInFile
                        Case "MOLE/MASSVOLUMETOTALUNITS": frmMMConvert.cboTotalVolume.ListIndex = lngSettingInFile
                        
                        Case "VISCOSITYMECNPERCENTACETONTRILE": frmViscosityForMeCN.txtPercentAcetonitrile = strSettingInFile
                        Case "VISCOSITYMECNTEMPERATURE": frmViscosityForMeCN.txtTemperature = strSettingInFile
                        Case "VISCOSITYMECNTEMPERATUREUNITS": frmViscosityForMeCN.cboTemperatureUnits.ListIndex = lngSettingInFile
                        
                        Case "CALCULATOR": frmCalculator.rtfExpression.Text = strSettingInFile
                        
                        Case "IONPLOTCOLOR": frmIsotopicDistribution.lblPlotColor.BackColor = lngSettingInFile
                        Case "IONPLOTSHOWPLOT": frmIsotopicDistribution.chkPlotResults.value = lngSettingInFile
                        Case "IONPLOTTYPE"
                            ValidateValueLng lngSettingInFile, 0, 1, 1          ' ipmGaussian
                            frmIsotopicDistribution.cboPlotType.ListIndex = lngSettingInFile
                        Case "IONPLOTRESOLUTION"
                            ValidateValueLng lngSettingInFile, 1, 1000000000#, 5000
                            frmIsotopicDistribution.txtEffectiveResolution = strSettingInFile
                        Case "IONPLOTRESOLUTIONMASS": frmIsotopicDistribution.txtEffectiveResolutionMass = strSettingInFile
                        Case "IONPLOTGAUSSIANQUALITY": frmIsotopicDistribution.txtGaussianQualityFactor = strSettingInFile
                        Case "IONCOMPARISONPLOTCOLOR": frmIsotopicDistribution.lblComparisonListPlotColor.BackColor = lngSettingInFile
                        Case "IONCOMPARISONPLOTTYPE"
                            ValidateValueLng lngSettingInFile, 0, 2, 0          ' ipmSticksToZero
                            frmIsotopicDistribution.cboComparisonListPlotType.ListIndex = lngSettingInFile
                        Case "IONCOMPARISONPLOTNORMALIZE": frmIsotopicDistribution.chkComparisonListNormalize.value = lngSettingInFile
                        
                        Case "FINDERACTION": SetDualOptionGroup frmFinder.optType, CInt(lngSettingInFile)
                        Case "FINDERMWT": frmFinder.txtMWT.Text = strSettingInFile
                        Case "FINDERPERCENTMAXWEIGHT": frmFinder.txtPercentMaxWeight.Text = strSettingInFile
                        Case "FINDERPPM": frmFinder.chkPpmMode.value = lngSettingInFile
                        Case "FINDERSHOWDELTAMASS": frmFinder.chkShowDeltaMass.value = lngSettingInFile
                        Case "FINDERWTTOLERANCE": frmFinder.txtWeightTolerance.Text = strSettingInFile
                        Case "FINDERPERCENTTOLERANCE": frmFinder.txtPercentTolerance.Text = strSettingInFile
                        Case "FINDERMAXHITS": frmFinder.txtHits.Text = strSettingInFile
                        Case "FINDERSORTRESULTS": frmFinderOptions.chkSort.value = lngSettingInFile
                        Case "FINDERSMARTH": frmFinderOptions.chkVerifyHydrogens.value = lngSettingInFile
                        Case "FINDERFINDCHARGE": frmFinderOptions.chkFindCharge.value = lngSettingInFile
                            frmFinderOptions.UpdateCheckBoxes
                        Case "FINDERFINDMTOZ": frmFinderOptions.chkFindMtoZ.value = lngSettingInFile
                            frmFinderOptions.UpdateCheckBoxes
                        Case "FINDERLIMITCHARGERANGE": frmFinderOptions.chkLimitChargeRange.value = lngSettingInFile
                            frmFinderOptions.UpdateCheckBoxes
                        Case "FINDERCHARGERANGEMIN": frmFinderOptions.txtChargeMin.Text = strSettingInFile
                        Case "FINDERCHARGERANGEMAX": frmFinderOptions.txtChargeMax.Text = strSettingInFile
                        Case "FINDERFINDTARGETMTOZ": frmFinderOptions.chkFindTargetMtoZ.value = lngSettingInFile
                            frmFinderOptions.UpdateCheckBoxes
                        Case "FINDERHIGHLIGHTTEXT": frmProgramPreferences.chkHighlightTextFields.value = lngSettingInFile
                        Case "FINDERAUTOBOUNDSSET": frmFinderOptions.chkAutoSetBounds.value = lngSettingInFile
                        Case "FINDERSORTMODE":
                            If dblSettingInFileValue <= frmFinderOptions.cboSortResults.ListCount - 1 Then
                                frmFinderOptions.cboSortResults.ListIndex = dblSettingInFileValue
                            End If
                        Case Else
                            If UCase(Left(strLineIn, 13)) = "CAPILLARYFLOW" Then
                                ParseCapillaryFlowSetting strIDStringInFile, strSettingInFile, intCapillaryFlowFileFormatVersion
                            ElseIf UCase(Left(strLineIn, 9)) = "FRAGMODEL" Or UCase(Left(strLineIn, 7)) = "IONPLOT" Then
                                ParseFragModelSetting InFileNum, strIDStringInFile, strSettingInFile
                            Else
                                If Not blnMatched Then
                                    ' Not matched, error
                                    ' Stop in IDE but ignore when compiled
                                    Debug.Assert False
                                End If
                            End If
                        End Select
                    End If
                Else
                    If UCase(Left(strLineIn, 21)) = "FRAGMODELIONMATCHLIST" Then
                        ' Typically this line is read by frmFragmentationModelling.LoadIonListToMatch()
                        ' However, if no ions are present, this line won't be read
                        ' Simply ignore it
                    Else
                        If Not PossiblySkipCWSpectrumSection(strLineIn, InFileNum, blnSkipNextLineRead) Then
                            ' Not matched, error
                            ' Stop in IDE but ignore when compiled
                            Debug.Assert False
                        End If
                    End If
                End If
            End If
        Loop Until EOF(InFileNum)
        
        Close InFileNum
        
        ' Load the CWSpectrum Options
        frmFragmentationModelling.LoadCWSpectrumOptions strFilePath

        If eViewModeSaved = vmdSingleView Then frmMain.SetViewMode vmdSingleView
        
        ' Rematch the loaded ions in case txtAlignment is nonzero
        ' Necessary since txtAlignment's value is set after the IonMatchList is Loaded
        frmFragmentationModelling.AlignmentOffsetValidate
    End If
    
LoadValuesAndFormulasExit:
    
    ' Restore blnDelayUpdate on frmMMConvert
    frmMMConvert.SetDelayUpdate False
    Exit Sub

LoadValuesAndFormulasErrorHandler:
    Close InFileNum
    ' Error Loading/Creating Values File
    strMessage = LookupMessage(330) & " (" & strFilePath & "): " & Err.Description
    strMessage = strMessage & vbCrLf & LookupMessage(340) & vbCrLf & LookupMessage(345)
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)
    Resume LoadValuesAndFormulasExit
            
End Sub

Private Sub OutputToFile(OutFileNum As Integer, strMessage As String, Optional strMessage2 As String = "", Optional strMessage3 As String = "", Optional strMessage4 As String = "", Optional strMessage5 As String = "")
    If Len(strMessage2) <= 12 Then
        Print #OutFileNum, strMessage; Tab(9); strMessage2; Tab(30); strMessage3; Tab(40); strMessage4; Tab(50); strMessage5
    Else
        Print #OutFileNum, strMessage; Tab(9); strMessage2; "   "; strMessage3; "     "; strMessage4; "     "; strMessage5
    End If
End Sub

Private Function ParseCapillaryFlowSetting(ByVal strIDStringInFile As String, ByVal strSettingInFile As String, ByRef intCapillaryFlowFileFormatVersion As Integer) As Boolean
    ' Returns true if setting is found
    Dim blnSettingFound As Boolean
    Dim intSettingInFileValue As Integer, dblSettingInFileValue As Double
    Dim lngIndex As Long
    Dim strIDStringToMatch As String
    
On Error GoTo ParseCapillaryFlowSettingErrorHandler

    If IsNumeric(strSettingInFile) Then
        dblSettingInFileValue = CDblSafe(strSettingInFile)
        intSettingInFileValue = CIntSafe(strSettingInFile)
    Else
        strSettingInFile = 0
    End If
    
    blnSettingFound = False
    
    For lngIndex = 0 To CapTextBoxMaxIndex
        strIDStringToMatch = "CAPILLARYFLOWOPENTEXT" & Trim(Str(lngIndex))
        If strIDStringInFile = strIDStringToMatch Then
            OpenCapVals.TextValues(lngIndex) = dblSettingInFileValue
            blnSettingFound = True
        End If
        If Not blnSettingFound Then
            strIDStringToMatch = "CAPILLARYFLOWPACKEDTEXT" & Trim(Str(lngIndex))
            If strIDStringInFile = strIDStringToMatch Then
                PackedCapVals.TextValues(lngIndex) = dblSettingInFileValue
                blnSettingFound = True
            End If
        End If
        If blnSettingFound Then Exit For
    Next lngIndex

    If Not blnSettingFound Then
        For lngIndex = 0 To CapComboBoxMaxIndex
            strIDStringToMatch = "CAPILLARYFLOWOPENCOMBO" & Trim(Str(lngIndex))
            blnSettingFound = ParseCapillaryFlowSettingCombo(lngIndex, strIDStringInFile, strIDStringToMatch, intSettingInFileValue, intCapillaryFlowFileFormatVersion, False)
            
            If Not blnSettingFound Then
                strIDStringToMatch = "CAPILLARYFLOWPACKEDCOMBO" & Trim(Str(lngIndex))
                blnSettingFound = ParseCapillaryFlowSettingCombo(lngIndex, strIDStringInFile, strIDStringToMatch, intSettingInFileValue, intCapillaryFlowFileFormatVersion, True)
            End If
            
            If blnSettingFound Then Exit For
        Next lngIndex
    End If
    
    If Not blnSettingFound Then
        ' Assume true for now
        blnSettingFound = True
        
        Select Case strIDStringInFile
        Case "CAPILLARYFLOWMODE": frmCapillaryCalcs.cboCapillaryType.ListIndex = intSettingInFileValue
        Case "CAPILLARYFLOWCOMPUTATIONTYPE": gCapFlowComputationTypeSave = intSettingInFileValue
        Case "CAPILLARYFLOWLINKFLOWRATE": gCapFlowLinkMassRateFlowRateSave = intSettingInFileValue
        Case "CAPILLARYFLOWLINKLINEARVELOCITY": gCapFlowLinkBdLinearVelocitySave = intSettingInFileValue
        Case "CAPILLARYFLOWSHOWPEAKBROADENINGSAVE": gCapFlowShowPeakBroadeningSave = intSettingInFileValue
        Case "CAPILLARYFLOWWEIGHTSOURCE": SetDualOptionGroup frmCapillaryCalcs.optWeightSource, intSettingInFileValue
        Case "CAPILLARYFLOWCUSTOMMASS": frmCapillaryCalcs.txtCustomMass.Text = strSettingInFile
        Case "CAPILLARYFLOWFILEFORMATVERSION": intCapillaryFlowFileFormatVersion = intSettingInFileValue
        Case Else
            ' Not found, set back to false
            blnSettingFound = False
        End Select
    End If
    
    ParseCapillaryFlowSetting = blnSettingFound
    Exit Function

ParseCapillaryFlowSettingErrorHandler:
    GeneralErrorHandler "FileIOFunctions|ParseCapillaryFlowSetting", Err.Number, Err.Description
    
End Function

Private Function ParseCapillaryFlowSettingCombo(lngComboBoxIndex As Long, strIDStringInFile As String, strIDStringToMatch As String, intSettingInFileValue As Integer, intCapillaryFlowFileFormatVersion As Integer, blnPackedCombo As Boolean) As Boolean
    Dim lngTargetIndex As Long
    Dim intTargetSetting As Integer
    
On Error GoTo ParseCapillaryFlowSettingComboErrorHandler

    If strIDStringInFile = strIDStringToMatch Then
        lngTargetIndex = lngComboBoxIndex
        intTargetSetting = intSettingInFileValue
        
        If intCapillaryFlowFileFormatVersion < CAP_FLOW_FILE_VERSIONTWO Then
            ' Old file version, may need to use a different save index or adjust SettingInFileValue
            Select Case lngComboBoxIndex
            Case 0          ' No conversion needed
            Case 1:         lngTargetIndex = lngComboBoxIndex + 1:   intTargetSetting = intSettingInFileValue + 3
            Case 2:         lngTargetIndex = lngComboBoxIndex + 1
            Case 3 To 12:   lngTargetIndex = lngComboBoxIndex + 2
            Case 13:        lngTargetIndex = lngComboBoxIndex + 3:   intTargetSetting = intSettingInFileValue + 3
            Case Else
                '  This is unexpected
                Debug.Assert False
            End Select
        End If
        
        If Not blnPackedCombo Then
            OpenCapVals.ComboValues(lngTargetIndex) = intTargetSetting
        Else
            PackedCapVals.ComboValues(lngTargetIndex) = intTargetSetting
        End If
        
        ParseCapillaryFlowSettingCombo = True
    Else
        ParseCapillaryFlowSettingCombo = False
    End If

    Exit Function

ParseCapillaryFlowSettingComboErrorHandler:
    GeneralErrorHandler "FileIOFunctions|ParseCapillaryFlowSettingCombo", Err.Number, Err.Description

End Function

Private Function ParseFragModelSetting(SeqFileNum As Integer, strIDStringInFile As String, strSettingInFile As String) As Boolean
    ' Returns true if setting is found
    
    Const MODIFICATION_SYMBOL_KEY = "FRAGMODELMODIFICATIONSYMBOL"
    Const MAX_PARSE_COUNT = 4
    
    Dim blnSettingFound As Boolean, blnSuccess As Boolean
    Dim lngSettingInFile As Long
    Dim lngModSymbolIndex As Long
    Dim lngParseCount As Long
    Dim strParsedVals(MAX_PARSE_COUNT) As String
    Dim strSymbol As String, strComment As String
    Dim dblMass As Double, blnIndicatesPhosphorylation As Boolean
    
On Error GoTo ParseFragModelSettingErrorHandler

    If IsNumeric(strSettingInFile) Then
        lngSettingInFile = CLngSafe(strSettingInFile)
    Else
        lngSettingInFile = 0
    End If
    
    blnSettingFound = True
    
    Select Case strIDStringInFile
    Case "FRAGMODELNOTATIONMODE": frmFragmentationModelling.cboNotation.ListIndex = lngSettingInFile
    Case "FRAGMODELSEQUENCE": frmFragmentationModelling.txtSequence = strSettingInFile
    Case "FRAGMODELNTERMINUS": frmFragmentationModelling.cboNTerminus.ListIndex = lngSettingInFile
    Case "FRAGMODELCTERMINUS": frmFragmentationModelling.cboCTerminus.ListIndex = lngSettingInFile
    Case "FRAGMODELIONTYPE0":  frmFragmentationModelling.chkIonType(0) = lngSettingInFile
    Case "FRAGMODELIONTYPE1": frmFragmentationModelling.chkIonType(1) = lngSettingInFile
    Case "FRAGMODELIONTYPE2": frmFragmentationModelling.chkIonType(2) = lngSettingInFile
    Case "FRAGMODELIONTYPE3": frmFragmentationModelling.chkIonType(3) = lngSettingInFile
    Case "FRAGMODELIONTYPE4": frmFragmentationModelling.chkIonType(4) = lngSettingInFile
    Case "FRAGMODELIONSTOMODIFY0": frmFragmentationModelling.lstIonsToModify.Selected(0) = lngSettingInFile
    Case "FRAGMODELIONSTOMODIFY1": frmFragmentationModelling.lstIonsToModify.Selected(1) = lngSettingInFile
    Case "FRAGMODELIONSTOMODIFY2": frmFragmentationModelling.lstIonsToModify.Selected(2) = lngSettingInFile
    Case "FRAGMODELWATERLOSS": frmFragmentationModelling.chkWaterLoss = lngSettingInFile
    Case "FRAGMODELAMMONIALOSS": frmFragmentationModelling.chkAmmoniaLoss = lngSettingInFile
    Case "FRAGMODELPHOSPHATELOSS": frmFragmentationModelling.chkPhosphateLoss = lngSettingInFile
    Case "FRAGMODELDOUBLECHARGE": frmFragmentationModelling.chkDoubleCharge = lngSettingInFile
    Case "FRAGMODELDOUBLECHARGETHRESHOLD": frmFragmentationModelling.cboDoubleCharge.ListIndex = lngSettingInFile
    Case "FRAGMODELTRIPLECHARGE": frmFragmentationModelling.chkTripleCharge = lngSettingInFile
    Case "FRAGMODELTRIPLECHARGETHRESHOLD": frmFragmentationModelling.cboTripleCharge.ListIndex = lngSettingInFile
    Case "FRAGMODELPRECURSORIONREMOVE": frmFragmentationModelling.chkRemovePrecursorIon = lngSettingInFile
    Case "FRAGMODELPRECURSORIONMASS": frmFragmentationModelling.txtPrecursorIonMass = strSettingInFile
    Case "FRAGMODELPRECURSORIONMASSWINDOW": frmFragmentationModelling.txtPrecursorMassWindow = strSettingInFile
    Case "FRAGMODELIONMATCHWINDOW": frmFragmentationModelling.txtIonMatchingWindow = strSettingInFile
    Case "FRAGMODELIONINTENSITY0": frmIonMatchOptions.txtIonIntensity(0) = strSettingInFile
    Case "FRAGMODELIONINTENSITY1": frmIonMatchOptions.txtIonIntensity(1) = strSettingInFile
    Case "FRAGMODELIONINTENSITY2": frmIonMatchOptions.txtIonIntensity(2) = strSettingInFile
    Case "FRAGMODELSHOULDERINTENSITY": frmIonMatchOptions.txtBYIonShoulders = strSettingInFile
    Case "FRAGMODELNEUTRALLOSSINTENSITY": frmIonMatchOptions.txtNeutralLosses = strSettingInFile
    Case "FRAGMODELGROUPIONS": frmIonMatchOptions.chkGroupSimilarIons = lngSettingInFile
    Case "FRAGMODELGROUPIONSMASSWINDOW": frmIonMatchOptions.txtGroupIonMassWindow = strSettingInFile
    Case "FRAGMODELNORMALIZEDINTENSITY": frmIonMatchOptions.txtNormalizedIntensity = strSettingInFile
    Case "FRAGMODELNORMALIZATIONIONSUSECOUNT": frmIonMatchOptions.txtIonCountToUse = strSettingInFile
    Case "FRAGMODELNORMALIZATIONMASSREGIONS": frmIonMatchOptions.txtMassRegions = strSettingInFile
    Case "FRAGMODELLABELMAINIONS": frmIonMatchOptions.chkFragSpecLabelMainIons = strSettingInFile
    Case "FRAGMODELLABELOTHERIONS": frmIonMatchOptions.chkFragSpecLabelOtherIons = strSettingInFile
    Case "FRAGMODELEMPHASIZEPROLINEYIONS": frmIonMatchOptions.chkFragSpecEmphasizeProlineYIons = strSettingInFile
    Case "FRAGMODELPLOTPREDICTEDSPECTRUMINVERTED": frmIonMatchOptions.chkPlotSpectrumInverted.value = lngSettingInFile
    Case "FRAGMODELAUTOLABELMASS": frmIonMatchOptions.chkAutoLabelMass.value = lngSettingInFile
    Case "FRAGMODELFRAGSPECTRUMCOLOR": frmIonMatchOptions.lblFragSpectrumColor.BackColor = lngSettingInFile
    Case "FRAGMODELMATCHINGIONDATACOLOR": frmIonMatchOptions.lblMatchingIonDataColor.BackColor = lngSettingInFile
    Case "FRAGMODELIONMATCHLIST"
        ' An ion match list is present in the column
        ' Read numbers in list until FragModelIonMatchListEnd is found (or a non-numeric line is found)
        ' Store results in IonMatchList
        blnSuccess = frmFragmentationModelling.LoadIonListToMatch(SeqFileNum, CLngSafe(strSettingInFile))
        If blnSuccess Then
            frmFragmentationModelling.UpdateIonMatchListWrapper
        End If
    Case "FRAGMODELIONMATCHLISTCAPTION": frmFragmentationModelling.SetIonMatchListCaption strSettingInFile

    Case "FRAGMODELIONALIGNMENT": frmFragmentationModelling.txtAlignment = strSettingInFile
    Case Else
        If Left(strIDStringInFile, Len(MODIFICATION_SYMBOL_KEY)) = MODIFICATION_SYMBOL_KEY Then
            lngModSymbolIndex = CLngSafe(Mid(strIDStringInFile, Len(MODIFICATION_SYMBOL_KEY) + 1))
            If lngModSymbolIndex > 0 Then
                lngParseCount = ParseString(strSettingInFile, strParsedVals(), MAX_PARSE_COUNT, ",")

                If lngParseCount >= 3 Then
                    strSymbol = strParsedVals(1)
                    dblMass = CDblSafe(strParsedVals(2))
                    blnIndicatesPhosphorylation = CBoolSafe(strParsedVals(3))
                    If lngParseCount >= 4 Then
                        strComment = strParsedVals(4)
                    Else
                        strComment = ""
                    End If
                    
                    If Len(strSymbol) > 0 Then
                        If lngModSymbolIndex = 1 Then
                            ' Clear any existing modification symbols
                            objMwtWin.Peptide.RemoveAllModificationSymbols
                        End If
                    
                        objMwtWin.Peptide.SetModificationSymbol strSymbol, dblMass, blnIndicatesPhosphorylation, strComment
                    End If
                End If
            End If
        Else
            blnSettingFound = False
        End If
    End Select
    
    ParseFragModelSetting = blnSettingFound

    Exit Function

ParseFragModelSettingErrorHandler:
    GeneralErrorHandler "FileIOFunctions|ParseFragModelSetting", Err.Number, Err.Description

End Function

Private Function PossiblySkipCWSpectrumSection(ByRef strCurrentLine As String, ByVal InFileNum As Integer, ByRef blnSkipNextLineRead As Boolean) As Boolean
    ' If strCurrentLine is a valid section header from CWSpectrum Options, then
    '  reads all of the subsequent lines belonging to the section
    ' Returns True if a valid section header was found, false otherwise
    
    Dim blnValidHeaderFound As Boolean
    
    blnSkipNextLineRead = False
    
    If InFileNum < 0 Then
        PossiblySkipCWSpectrumSection = False
        Exit Function
    End If
    
    If UCase(Left(strCurrentLine, 19)) = "[PLOTOPTIONS_SERIES" Or _
       UCase(strCurrentLine) = "[GLOBALPLOTOPTIONS]" Then
        
        blnValidHeaderFound = True
        
        ' Ignore this entire section
        Do
            Line Input #InFileNum, strCurrentLine
            strCurrentLine = Trim(strCurrentLine)
            If Len(strCurrentLine) = 0 Then
                Exit Do
            ElseIf Left(strCurrentLine, 1) = "[" Or InStr(strCurrentLine, "=") <= 0 Then
                blnSkipNextLineRead = True
                Exit Do
            End If
        Loop Until EOF(InFileNum)
    End If
    
    PossiblySkipCWSpectrumSection = blnValidHeaderFound
    
End Function

Public Sub SaveAbbreviations(Optional blnSortAbbreviations As Boolean = True, Optional blnUseMessageBoxForErrors As Boolean = False, Optional ByRef strBackupFilePath As String)
    
    Dim lngIndex As Long
    Dim strSymbol As String, strFormula As String, strOneLetterSymbol As String, strComment As String
    Dim blnIsAminoAcid As Boolean
    Dim sngCharge As Single
    Dim lngError As Long
    
    Dim strFilePath As String
    Dim OutFileNum As Integer
    
On Error GoTo SaveAbbreviationsErrorHandler
    
    If blnSortAbbreviations Then
        ' Sort the abbreviations before we save them
        objMwtWin.SortAbbreviations
    End If

    If gBlnWriteFilesOnDrive Then
        
        strFilePath = BuildPath(gCurrentPath, ABBREVIATIONS_FILENAME)
        strBackupFilePath = BackupFile(strFilePath)
        
        OutFileNum = FreeFile()
        Open strFilePath For Output As #OutFileNum
        
        ' Intro
        Print #OutFileNum, COMMENT_CHAR & "       Abbreviations File for MWTWIN program (v" & PROGRAM_VERSION & ")"
        Print #OutFileNum, COMMENT_CHAR
        Print #OutFileNum, COMMENT_CHAR & " Comments may be added by preceding with a semicolon"
        Print #OutFileNum, COMMENT_CHAR & " Two group headings must exist: [AMINO ACIDS] and [ABBREVIATIONS]"
        Print #OutFileNum, COMMENT_CHAR & " Abbreviations may be added; simply type the abbreviation and molecular"
        Print #OutFileNum, COMMENT_CHAR & "   formula under the appropriate column in either section"
        Print #OutFileNum, COMMENT_CHAR & " Note:  Only the first letter of an abbreviation is allowed to be capitalized"
        Print #OutFileNum, COMMENT_CHAR
        Print #OutFileNum, COMMENT_CHAR & " The abbreviations in the Amino Acid section are treated as extended level"
        Print #OutFileNum, COMMENT_CHAR & "   abbreviations:  they are only recognized when extended abbreviations are on"
        Print #OutFileNum, COMMENT_CHAR & " If this file becomes corrupted, the MWTWIN program will inform the user"
        Print #OutFileNum, COMMENT_CHAR & "    and ignore incorrect lines"
        Print #OutFileNum, COMMENT_CHAR & " If this file becomes deleted, the MWTWIN program will create a new file"
        Print #OutFileNum, COMMENT_CHAR & "    with the default abbreviations"
        Print #OutFileNum, COMMENT_CHAR
        Print #OutFileNum, COMMENT_CHAR & " Default Amino Acids are in their ionic form"
        Print #OutFileNum, COMMENT_CHAR & " Amino Acid abbreviation names may be up to 6 characters long"
    
        ' Amino Acids
        Print #OutFileNum, "[AMINO ACIDS]"
    
        For lngIndex = 1 To objMwtWin.GetAbbreviationCount
            lngError = objMwtWin.GetAbbreviation(lngIndex, strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment)
            If blnIsAminoAcid Then
                OutputToFile OutFileNum, strSymbol, strFormula, CStr(sngCharge), strOneLetterSymbol, COMMENT_CHAR & " " & strComment
            End If
        Next lngIndex
            
        ' Other Abbreviations
        Print #OutFileNum, ""
        Print #OutFileNum, COMMENT_CHAR & " Normal abbreviation names may be up to 6 characters long"
        Print #OutFileNum, "[ABBREVIATIONS]"
            
        For lngIndex = 1 To objMwtWin.GetAbbreviationCount
            lngError = objMwtWin.GetAbbreviation(lngIndex, strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment)
            If Not blnIsAminoAcid Then
                OutputToFile OutFileNum, strSymbol, strFormula, CStr(sngCharge), strOneLetterSymbol, COMMENT_CHAR & " " & strComment
            End If
        Next lngIndex
        
        Close #OutFileNum
    End If
    
    Exit Sub

SaveAbbreviationsErrorHandler:
    Debug.Assert False
    Debug.Print "Error in SaveAbbreviations: " & Err.Description
    AddToIntro LookupMessage(100) & ": " & Err.Description, blnUseMessageBoxForErrors, True
    
End Sub

Public Sub SaveCapillaryFlowInfo(Optional intCurrentFileNumber As Integer = 0)
    ' Saves capillary flow info to a file
    ' Can also save the loaded ion list to the .Seq file
    '
    ' If intCurrentFileNumber is not specified then creates a new file
    ' Otherwise, saves to the file given by intCurrentFileNumber
    
    Dim strInfoFilePath As String, strMessage As String
    Dim lngIndex As Long, strLineOut As String
    Dim InfoFileNum As Integer

    If intCurrentFileNumber = 0 Then
        ' 1550 = Capillary Flow Info Files, 1555 = .cap
        strInfoFilePath = SelectFile(frmCapillaryCalcs.hwnd, "Save File", gLastFileOpenSaveFolder, True, "", ConstructFileDialogFilterMask(LookupMessage(1550), LookupMessage(1555)), 1)
        If Len(strInfoFilePath) = 0 Then
            ' No file selected (or other error)
            Exit Sub
        End If
    End If
    
    On Error GoTo SaveFlowProblem
    
    If intCurrentFileNumber = 0 Then
        ' Open the file for output
        InfoFileNum = FreeFile()
        Open strInfoFilePath For Output As #InfoFileNum
    Else
        InfoFileNum = intCurrentFileNumber
    End If
    
    With frmCapillaryCalcs
        If intCurrentFileNumber = 0 Then
            Print #InfoFileNum, COMMENT_CHAR & " Molecular Weight Calculator Capillary Flow Information File"
        End If
        Print #InfoFileNum, "CapillaryFlowFileFormatVersion=" & Trim(CAP_FLOW_FILE_VERSIONTWO)
        Print #InfoFileNum, "CapillaryFlowMode=" & Trim(frmCapillaryCalcs.cboCapillaryType.ListIndex)
        Print #InfoFileNum, "CapillaryFlowComputationType=" & CStr(gCapFlowComputationTypeSave)
        Print #InfoFileNum, "CapillaryFlowLinkFlowRate=" & CStr(gCapFlowLinkMassRateFlowRateSave)
        Print #InfoFileNum, "CapillaryFlowLinkLinearVelocity=" & CStr(gCapFlowLinkBdLinearVelocitySave)
        Print #InfoFileNum, "CapillaryFlowShowPeakBroadeningSave=" & CStr(gCapFlowShowPeakBroadeningSave)
        
        strLineOut = "CapillaryFlowWeightSource="
        If frmCapillaryCalcs.optWeightSource(0) = True Then
            strLineOut = strLineOut & "0"
        Else
            strLineOut = strLineOut & "1"
        End If
        Print #InfoFileNum, strLineOut
        
        Print #InfoFileNum, "CapillaryFlowCustomMass=" & Trim(frmCapillaryCalcs.txtCustomMass)

        For lngIndex = 0 To CapTextBoxMaxIndex
            Print #InfoFileNum, "CapillaryFlowOpenText" & CStr(lngIndex) & "=" & CStr(OpenCapVals.TextValues(lngIndex))
            Print #InfoFileNum, "CapillaryFlowPackedText" & CStr(lngIndex) & "=" & CStr(PackedCapVals.TextValues(lngIndex))
        Next lngIndex
        
        For lngIndex = 0 To CapComboBoxMaxIndex
            Print #InfoFileNum, "CapillaryFlowOpenCombo" & CStr(lngIndex) & "=" & CStr(OpenCapVals.ComboValues(lngIndex))
            Print #InfoFileNum, "CapillaryFlowPackedCombo" & CStr(lngIndex) & "=" & CStr(PackedCapVals.ComboValues(lngIndex))
        Next lngIndex
    End With

    If intCurrentFileNumber = 0 Then
        Close InfoFileNum
    End If
    
    Exit Sub
    
SaveFlowProblem:
    Close
    strMessage = LookupMessage(330) & ": " & strInfoFilePath
    strMessage = strMessage & vbCrLf & Err.Description
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)

End Sub

Public Sub SaveDefaultOptions()
    Dim strMessage As String, strLineOut As String, strFilePath As String
    Dim FinderWeightModeWarn As Integer
    Dim OutFileNum As Integer
    
On Error GoTo SaveDefaultOptionsErrorHandler
    
    If gBlnWriteFilesOnDrive Then
        strFilePath = BuildPath(gCurrentPath, INI_FILENAME)
        
        OutFileNum = FreeFile()
        Open strFilePath For Output As #OutFileNum
            
            ' Intro
            Print #OutFileNum, COMMENT_CHAR & "       Options File for MWTWIN Program (v" & PROGRAM_VERSION & ")"
            Print #OutFileNum, COMMENT_CHAR
            Print #OutFileNum, COMMENT_CHAR & " File Automatically Created -- Select Save Options As Defaults in Preferences under the Options Menu"
            Print #OutFileNum, COMMENT_CHAR
            
            strLineOut = "View="
            If frmMain.GetViewMode() = vmdMultiView Then
                strLineOut = strLineOut & "0"
            Else
                strLineOut = strLineOut & "1"
            End If
            
            Print #OutFileNum, strLineOut
            
            With frmProgramPreferences
                strLineOut = "Convert="
                If .optConvertType(0).value = True Then
                    strLineOut = strLineOut & "0"
                ElseIf .optConvertType(1).value = True Then
                    strLineOut = strLineOut & "1"
                Else
                    strLineOut = strLineOut & "2"
                End If
                Print #OutFileNum, strLineOut
                
                strLineOut = "Abbrev="
                If .optAbbrevType(0).value = True Then
                    strLineOut = strLineOut & "0"
                ElseIf .optAbbrevType(1).value = True Then
                    strLineOut = strLineOut & "1"
                Else
                    strLineOut = strLineOut & "2"
                End If
                Print #OutFileNum, strLineOut
                
                strLineOut = "StdDev="
                If .optStdDevType(0).value = True Then
                    strLineOut = strLineOut & "0"
                ElseIf .optStdDevType(1).value = True Then
                    strLineOut = strLineOut & "1"
                ElseIf .optStdDevType(2).value = True Then
                    strLineOut = strLineOut & "2"
                Else
                    strLineOut = strLineOut & "3"
                End If
                Print #OutFileNum, strLineOut
                
                Print #OutFileNum, "Caution=" & CheckBoxToIntegerString(.chkShowCaution)
            
                Print #OutFileNum, "Advance=" & CheckBoxToIntegerString(.chkAdvanceOnCalculate)
                
                Print #OutFileNum, "Charge=" & CheckBoxToIntegerString(.chkComputeCharge)
                
                Print #OutFileNum, "QuickSwitch=" & CheckBoxToIntegerString(.chkShowQuickSwitch)
                
                Print #OutFileNum, "Font=" & objMwtWin.RtfFontName
                Print #OutFileNum, "FontSize=" & objMwtWin.RtfFontSize
                
                strLineOut = "ExitConfirm="
                If .optExitConfirmation(exmEscapeKeyConfirmExit).value = True Then
                    strLineOut = strLineOut & "0"
                ElseIf .optExitConfirmation(exmEscapeKeyDoNotConfirmExit).value = True Then
                    strLineOut = strLineOut & "1"
                ElseIf .optExitConfirmation(exmIgnoreEscapeKeyConfirmExit).value = True Then
                    strLineOut = strLineOut & "2"
                Else
                    strLineOut = strLineOut & "3"
                End If
                Print #OutFileNum, strLineOut
                
                Print #OutFileNum, "ToolTips=" & CheckBoxToIntegerString(.chkShowToolTips)
                
                Print #OutFileNum, "HideInactiveForms=" & CheckBoxToIntegerString(.chkHideInactiveForms)
                
                Print #OutFileNum, "AutoSaveValues=" & CheckBoxToIntegerString(.chkAutosaveValues)
    
                Print #OutFileNum, "BracketsAsParentheses=" & CheckBoxToIntegerString(.chkBracketsAsParentheses)
    
                Print #OutFileNum, "AutoCopyCurrentMWT=" & CheckBoxToIntegerString(.chkAutoCopyCurrentMWT)
    
                ' It is important that this line be present after the Hide Inactive program windows and Exit Program behavior options
                Print #OutFileNum, "StartupModule=" & frmProgramPreferences.cboStartupModule.ListIndex
                
                strLineOut = "MaximumFormulasToShow="
                strLineOut = strLineOut & FormatMaximumFormulasToShowString()
                Print #OutFileNum, strLineOut

                If cChkBox(.chkAlwaysSwitchToIsotopic) Then
                    FinderWeightModeWarn = 1
                ElseIf cChkBox(.chkNeverShowFormulaFinderWarning) Then
                    FinderWeightModeWarn = -1
                Else
                    FinderWeightModeWarn = 0
                End If
                
                Print #OutFileNum, "FinderWeightModeWarn=" & FinderWeightModeWarn
                Print #OutFileNum, "FinderBoundedSearch=" & frmFinderOptions.cboSearchType.ListIndex
                Print #OutFileNum, "Language=" & gCurrentLanguage
                Print #OutFileNum, "LanguageFile=" & gCurrentLanguageFileName
                Print #OutFileNum, "LastOpenSaveFolder=" & gLastFileOpenSaveFolder
            End With
            
        Close OutFileNum
        
        frmMain.lblStatus.ForeColor = vbWindowText
        frmMain.lblStatus.Caption = LookupLanguageCaption(11930, "Default options saved.")
    Else
        frmMain.lblStatus.ForeColor = vbWindowText
        frmMain.lblStatus.Caption = LookupLanguageCaption(11935, "Default options NOT saved since /X command line option was used.")
    End If
    
    Exit Sub

SaveDefaultOptionsErrorHandler:
    Close OutFileNum
    strMessage = LookupMessage(360) & " (" & strFilePath & "): " & Err.Description
    strMessage = strMessage & vbCrLf & LookupMessage(345)
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)
    
End Sub

Public Sub SaveElements(Optional ByRef strBackupFilePath As String)
    ' Saves the elements to the default element file
    ' Returns the backup file path byref in strBackupFilePath
    
    Dim strFilePath As String
    Dim OutFileNum As Integer
    
On Error GoTo SaveElementsErrorHandler
    
    If gBlnWriteFilesOnDrive Then
        strFilePath = BuildPath(gCurrentPath, ELEMENTS_FILENAME)
        strBackupFilePath = BackupFile(strFilePath)
        
        OutFileNum = FreeFile()
        Open strFilePath For Output As #OutFileNum
        
        SaveElementsWork OutFileNum
        gElementWeightTypeInFile = objMwtWin.GetElementMode
        
        Close OutFileNum
    End If
    
    Exit Sub

SaveElementsErrorHandler:
    Close OutFileNum
    AddToIntro LookupMessage(320) & " (" & strFilePath & "): " & Err.Description

End Sub

Private Sub SaveElementsWork(OutFileNum As Integer)
    Dim intIndex As Integer
    Dim strSymbol As String, dblMass As Double, dblUncertainty As Double
    Dim sngCharge As Single, intIsotopeCount As Integer
    Dim lngError As Long
    
    ' Intro
    Print #OutFileNum, COMMENT_CHAR & "       Elements File for MWTWIN program (v" & PROGRAM_VERSION & ")"
    Print #OutFileNum, COMMENT_CHAR
    Print #OutFileNum, COMMENT_CHAR & " Comments may be added by preceding with a semicolon"
    Print #OutFileNum, COMMENT_CHAR & " The heading [ELEMENTWEIGHTTYPE] 1 signifies that Average Elemental Weights are"
    Print #OutFileNum, COMMENT_CHAR & "   being used while [ELEMENTWEIGHTTYPE] 2 signifies the use of Isotopic Weights"
    Print #OutFileNum, COMMENT_CHAR & "   and [ELEMENTWEIGHTTYPE] 3 signifies the use of Integer Weights"
    Print #OutFileNum, COMMENT_CHAR & " The group heading [ELEMENTS] must exist to signify the start of the elements"
    Print #OutFileNum, COMMENT_CHAR & " Elemental values may be changed, but new elements may not be added"
    Print #OutFileNum, COMMENT_CHAR & " If you wish to add new elements, simply add them as abbreviations"
    Print #OutFileNum, COMMENT_CHAR & " If this file becomes deleted, the MWTWIN program will create a new file"
    Print #OutFileNum, COMMENT_CHAR & "   with the default values"
    Print #OutFileNum, COMMENT_CHAR & " Uncertainties from CRC Handbook of Chemistry and Physics"
    Print #OutFileNum, COMMENT_CHAR & "   For Radioactive elements, the most stable isotope is NOT used;"
    Print #OutFileNum, COMMENT_CHAR & "   instead, an average Mol. Weight is used, just like with other elements."
    Print #OutFileNum, COMMENT_CHAR & "   Data obtained from the Perma-Chart Science Series periodic table, 1993."
    Print #OutFileNum, COMMENT_CHAR & "   Uncertainties from CRC Handoobk of Chemistry and Physics, except for"
    Print #OutFileNum, COMMENT_CHAR & "   Radioactive elements, where uncertainty was estimated to be .n5 where"
    Print #OutFileNum, COMMENT_CHAR & "   n represents the number digits after the decimal point but before the last"
    Print #OutFileNum, COMMENT_CHAR & "   number of the molecular weight."
    Print #OutFileNum, COMMENT_CHAR & "   For example, for No, MW = 259.1009 (±0.0005)"
    Print #OutFileNum, COMMENT_CHAR
    Print #OutFileNum, "[ELEMENTWEIGHTTYPE] " & objMwtWin.GetElementMode
    Print #OutFileNum, COMMENT_CHAR & " The values signify:"
    Print #OutFileNum, COMMENT_CHAR & "       Weight           Uncertainty        Charge"
                
    ' The Elements
    Print #OutFileNum, "[ELEMENTS]"
                
    For intIndex = 1 To objMwtWin.GetElementCount
        lngError = objMwtWin.GetElement(intIndex, strSymbol, dblMass, dblUncertainty, sngCharge, intIsotopeCount)
        Debug.Assert lngError = 0
        
        Print #OutFileNum, strSymbol; Tab(8); dblMass; Tab(25); dblUncertainty; Tab(45); sngCharge
    Next intIndex

End Sub

Public Sub SaveSequenceInfo(IonMatchList() As Double, IonMatchListCount As Long, strIonMatchListCaption As String, Optional intCurrentFileNumber As Integer = 0)
    ' Saves Fragmentation Modelling sequence info to a file
    ' Can also save the loaded ion list to the .Seq file
    '
    ' If intCurrentFileNumber is not specified then creates a new file
    ' Otherwise, saves to the file given by intCurrentFileNumber
    
    Dim strSequenceFilePath As String, strMessage As String
    Dim lngIndex As Long
    Dim SeqFileNum As Integer
    Dim lngModificationSymbolCount As Long
    Dim lngErrorID As Long
    Dim strModSymbol As String, strComment As String
    Dim dblModMass As Double
    Dim blnIndicatesPhosphorylation As Boolean
    
    If intCurrentFileNumber = 0 Then
        ' 1530 = Sequence Files, 1535 = .seq
        strSequenceFilePath = SelectFile(frmFragmentationModelling.hwnd, "Select File", gLastFileOpenSaveFolder, True, "", ConstructFileDialogFilterMask(LookupMessage(1530), LookupMessage(1535)), 1)
        If Len(strSequenceFilePath) = 0 Then
            ' No file selected (or other error)
            Exit Sub
        End If
    End If
    
On Error GoTo SaveSequenceInfoErrorHandler
    
    If intCurrentFileNumber = 0 Then
        ' Open the file for output
        SeqFileNum = FreeFile()
        Open strSequenceFilePath For Output As #SeqFileNum
        Print #SeqFileNum, COMMENT_CHAR & " Molecular Weight Calculator Sequence Information File"
    Else
        SeqFileNum = intCurrentFileNumber
    End If
    
    ' Save the currently defined modification symbols
    lngModificationSymbolCount = objMwtWin.Peptide.GetModificationSymbolCount
    Print #SeqFileNum, "FragModelModificationSymbolsCount=" & lngModificationSymbolCount
    For lngIndex = 1 To lngModificationSymbolCount
        lngErrorID = objMwtWin.Peptide.GetModificationSymbol(lngIndex, strModSymbol, dblModMass, blnIndicatesPhosphorylation, strComment)
        Debug.Assert lngErrorID = 0
        
        Print #SeqFileNum, "FragModelModificationSymbol" & Trim(lngIndex) & "=" & strModSymbol & "," & CStr(dblModMass) & "," & CStr(blnIndicatesPhosphorylation) & "," & strComment
    Next lngIndex
    
    With frmFragmentationModelling
        Print #SeqFileNum, "FragModelNotationMode=" & .cboNotation.ListIndex
        ' Make sure no carriage returns are present in .txtSequence
        Print #SeqFileNum, "FragModelSequence=" & Trim(Replace(.txtSequence, vbCrLf, " "))
        Print #SeqFileNum, "FragModelNTerminus=" & .cboNTerminus.ListIndex
        Print #SeqFileNum, "FragModelCTerminus=" & .cboCTerminus.ListIndex
        Print #SeqFileNum, "FragModelIonType0=" & CheckBoxToIntegerString(.chkIonType(0))
        Print #SeqFileNum, "FragModelIonType1=" & CheckBoxToIntegerString(.chkIonType(1))
        Print #SeqFileNum, "FragModelIonType2=" & CheckBoxToIntegerString(.chkIonType(2))
        Print #SeqFileNum, "FragModelIonType3=" & CheckBoxToIntegerString(.chkIonType(3))
        Print #SeqFileNum, "FragModelIonType4=" & CheckBoxToIntegerString(.chkIonType(4))
        Print #SeqFileNum, "FragModelIonsToModify0=" & Abs(CIntSafe(.lstIonsToModify.Selected(0)))
        Print #SeqFileNum, "FragModelIonsToModify1=" & Abs(CIntSafe(.lstIonsToModify.Selected(1)))
        Print #SeqFileNum, "FragModelIonsToModify2=" & Abs(CIntSafe(.lstIonsToModify.Selected(2)))
        Print #SeqFileNum, "FragModelWaterLoss=" & CheckBoxToIntegerString(.chkWaterLoss)
        Print #SeqFileNum, "FragModelAmmoniaLoss=" & CheckBoxToIntegerString(.chkAmmoniaLoss)
        Print #SeqFileNum, "FragModelPhosphateLoss=" & CheckBoxToIntegerString(.chkPhosphateLoss)
        Print #SeqFileNum, "FragModelDoubleCharge=" & CheckBoxToIntegerString(.chkDoubleCharge)
        Print #SeqFileNum, "FragModelDoubleChargeThreshold=" & .cboDoubleCharge.ListIndex
        Print #SeqFileNum, "FragModelTripleCharge=" & CheckBoxToIntegerString(.chkTripleCharge)
        Print #SeqFileNum, "FragModelTripleChargeThreshold=" & .cboTripleCharge.ListIndex
        Print #SeqFileNum, "FragModelPrecursorIonRemove=" & CheckBoxToIntegerString(.chkRemovePrecursorIon)
        Print #SeqFileNum, "FragModelPrecursorIonMass=" & .txtPrecursorIonMass
        Print #SeqFileNum, "FragModelPrecursorIonMassWindow=" & .txtPrecursorMassWindow
        Print #SeqFileNum, "FragModelIonMatchWindow=" & .txtIonMatchingWindow
    End With
    
    With frmIonMatchOptions
        Print #SeqFileNum, "FragModelIonIntensity0=" & .txtIonIntensity(0)
        Print #SeqFileNum, "FragModelIonIntensity1=" & .txtIonIntensity(1)
        Print #SeqFileNum, "FragModelIonIntensity2=" & .txtIonIntensity(2)
        Print #SeqFileNum, "FragModelShoulderIntensity=" & .txtBYIonShoulders
        Print #SeqFileNum, "FragModelNeutralLossIntensity=" & .txtNeutralLosses
        Print #SeqFileNum, "FragModelGroupIons=" & CheckBoxToIntegerString(.chkGroupSimilarIons)
        Print #SeqFileNum, "FragModelGroupIonsMassWindow=" & .txtGroupIonMassWindow
        Print #SeqFileNum, "FragModelNormalizedIntensity=" & .txtNormalizedIntensity
        Print #SeqFileNum, "FragModelNormalizationIonsUseCount=" & .txtIonCountToUse
        Print #SeqFileNum, "FragModelNormalizationMassRegions=" & .txtMassRegions
        Print #SeqFileNum, "FragModelLabelMainIons=" & CheckBoxToIntegerString(.chkFragSpecLabelMainIons)
        Print #SeqFileNum, "FragModelLabelOtherIons=" & CheckBoxToIntegerString(.chkFragSpecLabelOtherIons)
        Print #SeqFileNum, "FragModelEmphasizeProlineYIons=" & CheckBoxToIntegerString(.chkFragSpecEmphasizeProlineYIons)
        Print #SeqFileNum, "FragModelPlotPredictedSpectrumInverted=" & CheckBoxToIntegerString(.chkPlotSpectrumInverted)
        Print #SeqFileNum, "FragModelAutoLabelMass=" & CheckBoxToIntegerString(.chkAutoLabelMass)
        Print #SeqFileNum, "FragModelFragSpectrumColor=" & .lblFragSpectrumColor.BackColor
        Print #SeqFileNum, "FragModelMatchingIonDataColor=" & .lblMatchingIonDataColor.BackColor
    End With
    
    If IonMatchListCount > 0 Then
        Print #SeqFileNum, "FragModelIonMatchList=" & Trim(Str(IonMatchListCount))
        For lngIndex = 1 To IonMatchListCount
            Print #SeqFileNum, CStr(IonMatchList(lngIndex, 0)) & "," & CStr(IonMatchList(lngIndex, 1))
        Next lngIndex
        Print #SeqFileNum, "FragModelIonMatchListEnd"
        Print #SeqFileNum, "FragModelIonMatchListCaption=" & strIonMatchListCaption
    End If
    
    ' Note: Must save/load the Ion Alignment value after the IonMatchList since the loading of the list sets the value to 0
    Print #SeqFileNum, "FragModelIonAlignment=" & frmFragmentationModelling.txtAlignment

    If intCurrentFileNumber = 0 Then
        Close SeqFileNum
        
        ' Wait 100 msec, then call .SaveCWSpectrumOptions
        Sleep 100
        frmFragmentationModelling.SaveCWSpectrumOptions strSequenceFilePath
    End If
    
    Exit Sub
    
SaveSequenceInfoErrorHandler:
    Close SeqFileNum
    strMessage = LookupMessage(900) & ": " & strSequenceFilePath
    strMessage = strMessage & vbCrLf & Err.Description
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)

End Sub

Public Sub SaveSingleDefaultOption(strOptionName As String, strOptionSetting As String)
    ' Code from Sub SaveDefaultOptions
    
    Const MAX_INI_FILENAME_LINES = 50
    Dim strFilePath As String, strLinesRead(MAX_INI_FILENAME_LINES) As String, intCharLoc As Integer
    Dim lngLineCount As Long
    Dim strLineIn As String, strMessage As String
    Dim blnReplaced As Boolean
    Dim InFileNum As Integer, OutFileNum As Integer
    
    ' Exit sub if program is loading
    If frmMain.lblHiddenFormStatus = "Loading" Then Exit Sub
    
On Error GoTo SaveSingleDefaultOptionsErrorHandler
    
    If gBlnWriteFilesOnDrive Then
        strFilePath = BuildPath(gCurrentPath, INI_FILENAME)
        
        InFileNum = FreeFile()
        Open strFilePath For Input As #InFileNum

        ' Code from LoadDefaultOptions subroutine
        lngLineCount = 0
        blnReplaced = False
        Do
            lngLineCount = lngLineCount + 1
            Line Input #InFileNum, strLinesRead(lngLineCount)
            strLineIn = Trim(strLinesRead(lngLineCount))

            If strLineIn <> "" And Not IsComment(strLineIn) Then
                intCharLoc = InStr(strLineIn, "=")
                If intCharLoc > 0 Then
                    If UCase(Left(strLineIn, intCharLoc - 1)) = UCase(strOptionName) Then
                        strLineIn = strOptionName & "=" & strOptionSetting
                        strLinesRead(lngLineCount) = strLineIn
                        blnReplaced = True
                    End If
                End If
            End If
        Loop Until EOF(InFileNum) Or lngLineCount > MAX_INI_FILENAME_LINES
        Close InFileNum
        
        If Not blnReplaced Then
            ' The desired line wasn't found, add it to the end
        
            lngLineCount = lngLineCount + 1
            strMessage = strOptionName & "=" & strOptionSetting
            strLinesRead(lngLineCount) = strLineIn
        End If
        
        OutFileNum = FreeFile
        Open strFilePath For Output As #OutFileNum
            For intCharLoc = 1 To lngLineCount
                Print #OutFileNum, strLinesRead(intCharLoc)
            Next intCharLoc
        Close OutFileNum
            
    End If

    Exit Sub

SaveSingleDefaultOptionsErrorHandler:
    Close
    strMessage = LookupMessage(360) & " (" & strFilePath & "): " & Err.Description
    strMessage = strMessage & vbCrLf & LookupMessage(345)
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)
    
End Sub

Public Sub SaveValuesAndFormulas()
    Dim strMessage As String, strLineOut As String, strFilePath As String
    Dim intIndex As Integer
    
    Dim ThisIonMatchListCount As Long
    Dim ThisIonMatchList() As Double        ' 1 based in the first dimension and 0-based in the second, using columns 0, 1, and 2
    Dim ThisIonMatchListCaption As String
    
    Dim OutFileNum As Integer
    
    ReDim ThisIonMatchList(1, 3)

On Error GoTo SaveValuesErrorHandler

    If gBlnWriteFilesOnDrive Then
        strFilePath = BuildPath(gCurrentPath, VALUES_FILENAME)
        BackupFile strFilePath
        
        OutFileNum = FreeFile()
        Open strFilePath For Output As OutFileNum
            ' Intro
            Print #OutFileNum, COMMENT_CHAR & "       Values File for MWTWIN Program (v" & PROGRAM_VERSION & ")"
            Print #OutFileNum, COMMENT_CHAR
            Print #OutFileNum, COMMENT_CHAR & " File Automatically Created -- Select Save Values and Formulas under the Options Menu"
            Print #OutFileNum, COMMENT_CHAR
            
            For intIndex = 0 To frmMain.GetTopFormulaIndex
                strLineOut = "Formula" & CStr(intIndex) & "="
                If frmMain.rtfFormula(intIndex).Text <> "" Then
                    strLineOut = strLineOut & frmMain.rtfFormula(intIndex).Text
                Else
                    ' nothing, just print the label string
                End If
                RemoveHeightAdjustChar strLineOut
                Print #OutFileNum, strLineOut
            Next intIndex
            
            With frmAminoAcidConverter
                Print #OutFileNum, "AminoAcidConvertOneLetter=" & Trim(.txt1LetterSequence)
                Print #OutFileNum, "AminoAcidConvertThreeLetter=" & Trim(.txt3LetterSequence)
                Print #OutFileNum, "AminoAcidConvertSpaceOneLetter=" & CheckBoxToIntegerString(.chkSpaceEvery10)
                Print #OutFileNum, "AminoAcidConvertDashThreeLetter=" & CheckBoxToIntegerString(.chkSeparateWithDash)
            End With
            
            With frmMMConvert
                
                strLineOut = "Mole/MassWeightSource="
                If .optWeightSource(0) = True Then
                    strLineOut = strLineOut & "0"
                Else
                    strLineOut = strLineOut & "1"
                End If
                Print #OutFileNum, strLineOut
                
                Print #OutFileNum, "Mole/MassCustomMass=" & Trim(.txtCustomMass)
                Print #OutFileNum, "Mole/MassAction=" & Trim(.cboAction.ListIndex)
                Print #OutFileNum, "Mole/MassFrom=" & Trim(.txtFromNum.Text)
                Print #OutFileNum, "Mole/MassFromUnits=" & Trim(.cboFrom.ListIndex)
                Print #OutFileNum, "Mole/MassDensity=" & Trim(.txtDensity.Text)
                Print #OutFileNum, "Mole/MassToUnits=" & Trim(.cboTo.ListIndex)
                Print #OutFileNum, "Mole/MassVolume=" & Trim(.txtVolume.Text)
                Print #OutFileNum, "Mole/MassVolumeUnits=" & Trim(.cboVolume.ListIndex)
                Print #OutFileNum, "Mole/MassMolarity=" & Trim(.txtConcentration.Text)
                Print #OutFileNum, "Mole/MassMolarityUnits=" & Trim(.cboConcentration.ListIndex)
                
                Print #OutFileNum, "Mole/MassDilutionMode=" & Trim(.cboDilutionMode.ListIndex)
                Print #OutFileNum, "Mole/MassMolarityInitial=" & Trim(.txtDilutionConcentrationInitial.Text)
                Print #OutFileNum, "Mole/MassMolarityInitialUnits=" & Trim(.cboDilutionConcentrationInitial.ListIndex)
                Print #OutFileNum, "Mole/MassVolumeInitial=" & Trim(.txtStockSolutionVolume.Text)
                Print #OutFileNum, "Mole/MassVolumeInitialUnits=" & Trim(.cboStockSolutionVolume.ListIndex)
                Print #OutFileNum, "Mole/MassMolarityFinal=" & Trim(.txtDilutionConcentrationFinal.Text)
                Print #OutFileNum, "Mole/MassMolarityFinalUnits=" & Trim(.cboDilutionConcentrationFinal.ListIndex)
                Print #OutFileNum, "Mole/MassVolumeSolvent=" & Trim(.txtDilutingSolventVolume.Text)
                Print #OutFileNum, "Mole/MassVolumeSolventUnits=" & Trim(.cboDilutingSolventVolume.ListIndex)
                Print #OutFileNum, "Mole/MassVolumeTotal=" & Trim(.txtTotalVolume.Text)
                Print #OutFileNum, "Mole/MassVolumeTotalUnits=" & Trim(.cboTotalVolume.ListIndex)
                
                Print #OutFileNum, "Mole/MassLinkConcentrations=" & CheckBoxToIntegerString(.chkLinkMolarities)
                Print #OutFileNum, "Mole/MassLinkDilutionVolumeUnits=" & CheckBoxToIntegerString(.chkLinkDilutionVolumeUnits)
            End With
            
            SaveCapillaryFlowInfo OutFileNum
            
            With frmViscosityForMeCN
                Print #OutFileNum, "ViscosityMeCNPercentAcetontrile=" & Trim(.txtPercentAcetonitrile)
                Print #OutFileNum, "ViscosityMeCNTemperature=" & Trim(.txtTemperature)
                Print #OutFileNum, "ViscosityMeCNTemperatureUnits=" & Trim(.cboTemperatureUnits.ListIndex)
            End With
            
            strLineOut = Trim(frmCalculator.rtfExpression.Text)
            RemoveHeightAdjustChar strLineOut
            Print #OutFileNum, "Calculator=" & strLineOut
            
            With frmIsotopicDistribution
                Print #OutFileNum, "IonPlotColor=" & Trim(.lblPlotColor.BackColor)
                Print #OutFileNum, "IonPlotShowPlot=" & CheckBoxToIntegerString(.chkPlotResults)
                Print #OutFileNum, "IonPlotType=" & Trim(.cboPlotType.ListIndex)
                Print #OutFileNum, "IonPlotResolution=" & Trim(.txtEffectiveResolution.Text)
                Print #OutFileNum, "IonPlotResolutionMass=" & Trim(.txtEffectiveResolutionMass.Text)
                Print #OutFileNum, "IonPlotGaussianQuality=" & Trim(.txtGaussianQualityFactor.Text)
                Print #OutFileNum, "IonComparisonPlotColor=" & Trim(.lblComparisonListPlotColor.BackColor)
                Print #OutFileNum, "IonComparisonPlotType=" & Trim(.cboComparisonListPlotType.ListIndex)
                Print #OutFileNum, "IonComparisonPlotNormalize=" & CheckBoxToIntegerString(.chkComparisonListNormalize)
            End With
            
            With frmFinder
                strLineOut = "FinderAction="
                If .optType(0) = True Then
                    strLineOut = strLineOut & "0"
                Else
                    strLineOut = strLineOut & "1"
                End If
                Print #OutFileNum, strLineOut
    
                Print #OutFileNum, "FinderMWT=" & Trim(.txtMWT.Text)
                Print #OutFileNum, "FinderPercentMaxWeight=" & Trim(.txtPercentMaxWeight.Text)
                Print #OutFileNum, "FinderPPM=" & CheckBoxToIntegerString(.chkPpmMode)
                Print #OutFileNum, "FinderShowDeltaMass=" & CheckBoxToIntegerString(.chkShowDeltaMass)
                Print #OutFileNum, "FinderWtTolerance=" & Trim(.txtWeightTolerance.Text)
                Print #OutFileNum, "FinderPercentTolerance=" & Trim(.txtPercentTolerance.Text)
                Print #OutFileNum, "FinderMaxHits=" & Trim(.txtHits.Text)
            End With
            
            With frmFinderOptions
                Print #OutFileNum, "FinderSortResults=" & CheckBoxToIntegerString(.chkSort)
                Print #OutFileNum, "FinderSmartH=" & CheckBoxToIntegerString(.chkVerifyHydrogens)
                Print #OutFileNum, "FinderFindCharge=" & CheckBoxToIntegerString(.chkFindCharge)
                Print #OutFileNum, "FinderFindMtoZ=" & CheckBoxToIntegerString(.chkFindMtoZ)
                Print #OutFileNum, "FinderLimitChargeRange=" & CheckBoxToIntegerString(.chkLimitChargeRange)
                Print #OutFileNum, "FinderChargeRangeMin=" & Trim(.txtChargeMin.Text)
                Print #OutFileNum, "FinderChargeRangeMax=" & Trim(.txtChargeMax.Text)
                Print #OutFileNum, "FinderFindTargetMtoZ=" & CheckBoxToIntegerString(.chkFindTargetMtoZ)
                Print #OutFileNum, "FinderHighlightText=" & CheckBoxToIntegerString(frmProgramPreferences.chkHighlightTextFields)
                Print #OutFileNum, "FinderAutoBoundsSet=" & CheckBoxToIntegerString(.chkAutoSetBounds)
                Print #OutFileNum, "FinderSortMode=" & Trim(.cboSortResults.ListIndex)
            End With
            
            With frmFinder
                For intIndex = 0 To 9
                    Print #OutFileNum, "FinderMin" & CStr(intIndex) & "=" & Trim(.txtMin(intIndex).Text)
                    Print #OutFileNum, "FinderMax" & CStr(intIndex) & "=" & Trim(.txtMax(intIndex).Text)
                    Print #OutFileNum, "FinderCheckElements" & CStr(intIndex) & "=" & CheckBoxToIntegerString(.chkElements(intIndex))
                    Print #OutFileNum, "FinderPercentValue" & CStr(intIndex) & "=" & Trim(.txtPercent(intIndex).Text)
                    If intIndex >= 4 Then
                        Print #OutFileNum, "FinderCustomWeight" & CStr(intIndex - 3) & "=" & .txtWeight(intIndex).Text
                    End If
                Next intIndex
            End With
            
            ThisIonMatchListCount = frmFragmentationModelling.GetIonMatchList(ThisIonMatchList(), ThisIonMatchListCaption)
            SaveSequenceInfo ThisIonMatchList(), ThisIonMatchListCount, ThisIonMatchListCaption, OutFileNum

        Close OutFileNum
        
        ' Wait 100 msec, then call frmFragmentationModelling.SaveCWSpectrumOptions
        Sleep 100
        frmFragmentationModelling.SaveCWSpectrumOptions strFilePath
        
        frmMain.lblStatus.ForeColor = vbWindowText
        frmMain.lblStatus.Caption = LookupLanguageCaption(11920, "Values and formulas saved.")
    Else
        frmMain.lblStatus.ForeColor = vbWindowText
        frmMain.lblStatus.Caption = LookupLanguageCaption(11925, "Values and formulas NOT saved since /X command line option was used.")
    End If

    Exit Sub

SaveValuesErrorHandler:
    Close OutFileNum
    strMessage = LookupMessage(380) & " (" & strFilePath & "): " & Err.Description
    strMessage = strMessage & vbCrLf & LookupMessage(345)
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)
    
End Sub

Private Sub SetCheckBoxValue(chkThisCheckBox As CheckBox, intValue As Integer)
    If intValue <> 0 Then
        chkThisCheckBox.value = vbChecked
    Else
        chkThisCheckBox.value = vbUnchecked
    End If
End Sub

Public Sub SetDualOptionGroup(optThisOptionGroup As Variant, intButtonToSet As Integer)
    ' Note: optThisOptionGroup is an array of OptionButtons
    '       Due to constraints in VB, I cannot declare as optThisOptionGroup() as OptionButton
    '       Instead, I must declare as a Variant
    If intButtonToSet >= 0 And intButtonToSet <= 1 Then
        optThisOptionGroup(intButtonToSet).value = True
    End If
End Sub

Private Function StripComment(ByRef strWork As String) As String
    ' Looks for a comment at the end of strWork
    ' If found, removes the comment from strWork and returns the comment
    ' If not found, returns "" and leaves strWork untouched
    
    Dim strComment As String
    Dim lngCharLoc As Long
    
    lngCharLoc = InStr(strWork, COMMENT_CHAR)
    If lngCharLoc > 0 Then
        strComment = Trim(Mid(strWork, lngCharLoc + 1))
        strWork = Trim(Left(strWork, lngCharLoc - 1))
    Else
        lngCharLoc = InStr(strWork, "'")
        If lngCharLoc > 0 Then
            strComment = Trim(Mid(strWork, lngCharLoc + 1))
            strWork = Trim(Left(strWork, lngCharLoc - 1))
        End If
    End If
    
    StripComment = strComment
End Function
