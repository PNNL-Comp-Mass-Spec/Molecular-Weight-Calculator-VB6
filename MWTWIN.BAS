Attribute VB_Name = "MwtWinProcedures"
Option Explicit
'   A chemistry molecular weight calculator
'       Windows 95/98/NT Version
'
'   By Matthew Monroe
'   Richland, WA
'
'   Dos Version completed on August 9, 1995
'   Windows Version 3.0 Completed on June 10, 1996
'   Windows 95 version 4.0 Completed August 25, 1997
'   Windows 95 version 4.6 final, Completed January 23, 1999
'   Windows 95/98/NT version 5.11, Last modified January 26, 2000  (VB v6.0, sp3)
'   Windows 9x/NT/2000 version 5.2, Last modified July 9, 2000  (VB v6.0, sp5)
'   Windows 9x/ME/NT/2000/XP version 6.23, Last modified November 11, 2002 (VB v6.0, sp5)
'   Primary computation algorithms moved to an ActiveX Dll in December 2002, allowing the computational routines to be accessed by other developers' programs
'
'   E-mail: Matt@Alchemistmatt.Com or AlchemistMatt@Yahoo.com
'   WWW:    http://www.alchemistmatt.com/ and
'           http://ncrr.pnl.gov/software/
'           http://www.geocities.com/alchemistmatt/
'           http://come.to/alchemistmatt/
'
'   As of version 4.61 (2/1/99) there are 10,300 lines of code
'   As of version 4.62 (2/8/99) there are  9,505 lines of code (combined redundant functions)
'   As of version 5.11 (1/25/00) there are 13,092 lines of code (added numerous items, including Capillary Calculations form, plus made more readable)
'   As of version 6.14 (5/9/02) there are 25,300 lines of code; 9,200 in .Bas files and 16,100 in .Frm files  (numerous new features added since last count)

'   version 6.35

' *** Be sure to update version in Project | Properties also ***
Public Const PROGRAM_VERSION = "6.45"
Public Const PROGRAM_DATE = "July 6, 2007"
' ***

'
'
' Test String: HHeLiBeBCNoFNeNaMgAlSiPSClArKCaScTiVCrMnFeCoNiCu[xZnGaGe]AsSeBrKrRbSrYZrNbMoTcRuRhPdAgCdInSnSbTeIXeCsBaLaHfTaWReOsIrPtAuHgTiPbBiPoAtRnFrRaAcCePrNdPmSmEuGdTbDyHoErTmYbLuThPaUNpPuAmCmBkCfEsFmMdNoLr
'       Should give answer MW = 13140.9(±2) in Average Mode
'       and                MW = 13167.4185807(±0) in Isotopic mode
'       and                MW = 13172(±0) in Integer Mode


' ToDo for next version:
'  Language Captions for all the new labels and buttons and forms and tooltips
'  Tooltips for new items
'  Update help file for isotopic distribution simulator and Iso option in formula finder (Ctrl+D)
'  Compute multiplicity in isotopic distribution
'  Update the Overview.htm help file

' June 2003:
' Allow up to 3 custom neutral losses in Peptide fragmentation modeller
' Allow multiple neutral losses (see protein prospector for example)


'
' File Name Constants
'
Public Const INI_FILENAME = "MWTWIN.INI"
Public Const HELP_FILENAME = "MWTWIN.CHM"
Public Const ELEMENTS_FILENAME = "MWT_ELEM.DAT"
Public Const ABBREVIATIONS_FILENAME = "MWT_ABBR.DAT"
Public Const VALUES_FILENAME = "MWT_VALU.INI"
Public Const DEFAULT_LANGUAGE_FILENAME = "Lang_English.ini"
Public Const MAX_LANGUAGE_FILE_COUNT = 200                         ' Maximum number of language files to find in a given directory

Public Const FORMULA_CHANGED = "Changed"

' Formula and element Constants
'
Public Const MAX_FORMULAS = 25                  ' Maximum number of formulas that will ever be shown simultaneously
Public Const MAX_ELEMENT_INDEX = 103

'
' Public Constants
'
Public Const RTF_HEIGHT_ADJUSTCHAR = "~"           ' A hidden character to adjust the height of Rtf Text Boxes when using superscripts
Public Const EMPTY_STRINGCHAR = "~"
Public Const COMMENT_CHAR = ";"
Public Const LowestValueForDoubleDataType = -1.79E+308      ' Use -3.4E+38 for Single
Public Const HighestValueForDoubleDataType = 1.79E+308      ' Use  3.4E+38 for Single

'
' Enumerated Constants
'
Public Enum vmdViewModeConstants
    vmdMultiView = 0
    vmdSingleView = 1
End Enum

Public Enum psmPercentSolverModeConstants
    psmPercentSolverOff = 0
    psmPercentSolverOn = 1
End Enum

' Constants for Text Boxes on Capillary Calcs form
' The text boxes are part of an array;
' the following constants are used to reference the appropriate text box in the array
Public Const CapTextBoxMaxIndex = 16
Public Enum cctCapCalcTextBoxIDConstants
     cctPressure = 0
     cctColumnLength = 1
     cctColumnID = 2
     cctViscosity = 3
     cctParticleDiamter = 4
     cctFlowRate = 5
     cctDeadTime = 6
     cctPorosity = 7
     cctMassRateConcentration = 8
     cctMassRateVolFlowRate = 9
     cctMassRateInjectionTime = 10
     cctBdLinearVelocity = 11
     cctBdDiffusionCoefficient = 12
     cctBdOpenTubeLength = 13
     cctBdOpenTubeID = 14
     cctBdInitialPeakWidth = 15
     cctBdAdditionalVariance = 16
End Enum

' Constants for cboCapValue() array of Combo Boxes on Capillary Calcs form
Public Const CapComboBoxMaxIndex = 18
Public Enum cccCapCalComboBoxIDConstants
     cccPressureUnits = 0
     cccColumnLengthUnits = 1
     cccColumnIDUnits = 2
     cccViscosityUnits = 3
     cccParticleDiameterUnits = 4
     cccFlowRateUnits = 5
     cccLinearVelocityUnits = 6
     cccDeadTimeUnits = 7
     cccVolumeUnits = 8
     cccMassRateConcentrationUnits = 9
     cccMassRateVolFlowRateUnits = 10
     cccMassRateInjectionTimeUnits = 11
     cccMassFlowRateUnits = 12
     cccMassRateMolesInjectedUnits = 13
     cccBdLinearVelocityUnits = 14
     cccBdOpenTubeLengthUnits = 15
     cccBdOpenTubeIDUnits = 16
     cccBdInitialPeakWidthUnits = 17
     cccBdResultantPeakWidthUnits = 18
End Enum

Public Enum mmcMassModeConstants
    mmcComputedMass = 0
    mmcCustomMass = 1
End Enum

Public Enum exmExitModeConstants
    exmEscapeKeyConfirmExit = 0
    exmEscapeKeyDoNotConfirmExit
    exmIgnoreEscapeKeyConfirmExit
    exmIgnoreEscapeKeyDoNotConfirmExit
End Enum

Public Enum gcmGridCopyModeConstants
    gcmText = 0
    gcmRTF = 1
    gcmHTML = 2
End Enum

Public Enum smcFinderResultsSortModeConstants
    smcSortByFormula = 0
    smcSortByCharge = 1
    smcSortByMWT = 2
    smcSortByMZ = 3
    smcSortByDeltaMass = 4
End Enum

' Choose Color API's
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColorType) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'structure used for Color Dialog
Private Type ChooseColorType
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


' Constants for Button Return Status between forms
Public Const BUTTON_NOT_CLICKED_YET = -1
Public Const BUTTON_OK = 0
Public Const BUTTON_CANCEL = 1
Public Const BUTTON_RESET = 2           ' Also used for Remove Abbreviation or Remove Mod Symbol

' Screen Colors
Public Const COLOR_ERR = 12           ' Light Red
Public Const COLOR_WARN = 13          ' Light Yellow
Public Const COLOR_SOLVER = 2         ' Green
Public Const COLOR_CALC = 2          ' Bright White
Public Const COLOR_DIREC = 9          ' Light Blue
Public Const COLOR_WHITE = 15         ' White
Public Const COLOR_COMPUTEDQUANTITY = 14   ' Bright yellow

' Constants for loading an HTML Help file
Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_HELP_CONTEXT = &HF          ' Display mapped numeric value in dwData
''Public Const HH_SET_WIN_TYPE = &H4          ' Unused
''Public Const HH_GET_WIN_TYPE = &H5          ' Unused
''Public Const HH_GET_WIN_HANDLE = &H6        ' Unused
''Public Const HH_DISPLAY_TEXT_POPUP = &HE    ' Unused - Display string resource ID or text in a pop-up window
''Public Const HH_TP_HELP_CONTEXTMENU = &H10  ' Unused - Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU
''Public Const HH_TP_HELP_WM_HELP = &H11      ' Unused - text pop-up help, similar to WinHelp's HELP_WM_HELP
                                                                        
' Declare the HTML help function
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
    (ByVal hwndCaller As Long, ByVal pszFile As String, _
    ByVal uCommand As Long, ByVal dwData As Long) As Long

'
' User Defined Types (UDT's)
'
Public Type udtCapillaryFlowValuesType
    ' Text Boxes
    TextValues(CapTextBoxMaxIndex + 1) As Double
    
    ' Combo Boxes
    ComboValues(CapComboBoxMaxIndex + 1) As Integer
End Type

'
' Public Non-Array Variables
'
Public gCurrentPath As String
Public gMaxFormulaIndex As Integer                 ' Maximum number of formulas to display for the given screen resolution; absolute max is MAX_FORMULAS
Public gBlnShowStdDevWithMass As Boolean
Public gBlnErrorPresent As Boolean
Public gBlnStatusCaution As Boolean
Public gBlnLoadStatusOK As Boolean
Public gBlnAccessFilesOnDrive As Boolean            ' If False, then drive is never accessed
Public gBlnWriteFilesOnDrive As Boolean             ' If false, then files are read, but not written

Public gKeyPressAbortFormulaFinder As Integer       ' Used with frmProgress

Public gElementWeightTypeInFile As emElementModeConstants

Public gCurrentLanguage As String
Public gCurrentLanguageFileName As String
Public gMWAbbreviation As String
Public gNonSaveExitApp As Boolean
Public gCommandLineInstructionsDisplayed As Boolean

Public gLastFileOpenSaveFolder As String
'
' Public Array variables
'

' The following variables are used to remember unique settings for
'   Open and Packed capillaries on frmCapillaryCalcs
Public OpenCapVals As udtCapillaryFlowValuesType
Public PackedCapVals As udtCapillaryFlowValuesType

Public gCapFlowComputationTypeSave As Integer
Public gCapFlowLinkMassRateFlowRateSave As Integer
Public gCapFlowLinkBdLinearVelocitySave As Integer
Public gCapFlowShowPeakBroadeningSave As Integer

' Reference to the MwtWinDll.Dll
Public objMwtWin As New MwtWinDll.MolecularWeightCalculator
Public objCompounds(MAX_FORMULAS) As New MWCompoundClass

' Purpose: Append a string to frmIntro.lblLoadStatus.Caption
'          When blnErrorMessage is true, then sets gBlnLoadStatusOK to False and shows cmdOK
Public Sub AddToIntro(strAdd As String, Optional blnUseMsgBoxInstead As Boolean = False, Optional blnErrorMessage As Boolean = True)
    If blnUseMsgBoxInstead Then
        MsgBox strAdd, vbOKOnly + vbInformation, frmMain.Caption
    Else
        frmIntro.lblLoadStatus.Caption = frmIntro.lblLoadStatus.Caption & vbCrLf & strAdd
        If blnErrorMessage Then
            gBlnLoadStatusOK = False
            frmIntro.cmdOK.Visible = True
        End If
    End If
    
End Sub

' Purpose: Append ellipses to a menu
Public Function AppendEllipsesToSingleMenu(strMenuCaption As String, boolDynamicMenu As Boolean, Optional boolAlwaysShowExclamationpoint As Boolean = False) As String
    ' Append ... to menu, making sure not to add ... if already present
    ' If boolDynamicMenu = True, then add either ... or ! depending on
    '   frmProgramPreferences.optExitConfirmation
    
    Dim strEllipsesToAppend As String, intOldCaptionEllipsesStartIndex As Integer
    Dim strOriginalCaption As String
    
    If Not boolDynamicMenu Then
        strEllipsesToAppend = "..."
    Else
        If (frmProgramPreferences.optExitConfirmation(exmEscapeKeyConfirmExit).value = True Or frmProgramPreferences.optExitConfirmation(exmIgnoreEscapeKeyConfirmExit).value = True) And _
           Not boolAlwaysShowExclamationpoint Then
            strEllipsesToAppend = "..."
        Else
            strEllipsesToAppend = "!"
        End If
    End If
    
    ' Strip to current caption
    ' First look for ellipses
    intOldCaptionEllipsesStartIndex = InStr(strMenuCaption, "...")
    If intOldCaptionEllipsesStartIndex <= 0 Then
        ' Ellipses not found, look for !
        intOldCaptionEllipsesStartIndex = InStr(strMenuCaption, "!")
    End If
    
    ' If either ... or ! was found, remove them
    If intOldCaptionEllipsesStartIndex > 0 Then
        strOriginalCaption = Left(strMenuCaption, intOldCaptionEllipsesStartIndex - 1)
    Else
        strOriginalCaption = strMenuCaption
    End If
    
    AppendEllipsesToSingleMenu = strOriginalCaption & strEllipsesToAppend
End Function

Public Sub AppendShortcutKeysToMenuCaptions()
    
    With frmMain
        ' First add ... or ! to corresponding menus
        .mnuEditElements.Caption = AppendEllipsesToSingleMenu(.mnuEditElements.Caption, False)
        .mnuEditAbbrev.Caption = AppendEllipsesToSingleMenu(.mnuEditAbbrev.Caption, False)
        .mnuCalculateFile.Caption = AppendEllipsesToSingleMenu(.mnuCalculateFile.Caption, False)
        .mnuPrint.Caption = AppendEllipsesToSingleMenu(.mnuPrint.Caption, False)
        .mnuExit.Caption = AppendEllipsesToSingleMenu(.mnuExit.Caption, True)
        .mnuEraseAll.Caption = AppendEllipsesToSingleMenu(.mnuEraseAll.Caption, False)
        .mnuEraseCurrent.Caption = AppendEllipsesToSingleMenu(.mnuEraseCurrent.Caption, False)
        .mnuExpandAbbrev.Caption = AppendEllipsesToSingleMenu(.mnuExpandAbbrev.Caption, True)
        .mnuEmpirical.Caption = AppendEllipsesToSingleMenu(.mnuEmpirical.Caption, True)
        .mnuMMConvert.Caption = AppendEllipsesToSingleMenu(.mnuMMConvert.Caption, False)
        .mnuFinder.Caption = AppendEllipsesToSingleMenu(.mnuFinder.Caption, False)
        .mnuAminoAcidNotationConverter.Caption = AppendEllipsesToSingleMenu(.mnuAminoAcidNotationConverter.Caption, False)
        .mnuPeptideSequenceFragmentation.Caption = AppendEllipsesToSingleMenu(.mnuPeptideSequenceFragmentation.Caption, False)
        .mnuIsotopicDistribution.Caption = AppendEllipsesToSingleMenu(.mnuIsotopicDistribution.Caption, False)
        .mnuCalculator.Caption = AppendEllipsesToSingleMenu(.mnuCalculator.Caption, False)
        .mnuCapillaryFlow.Caption = AppendEllipsesToSingleMenu(.mnuCapillaryFlow.Caption, False)
        .mnuChooseLanguage.Caption = AppendEllipsesToSingleMenu(.mnuChooseLanguage.Caption, False)
        .mnuProgramOptions.Caption = AppendEllipsesToSingleMenu(.mnuProgramOptions.Caption, False)
        .mnuChangeFont.Caption = AppendEllipsesToSingleMenu(.mnuChangeFont.Caption, False)
        .mnuRestoreValues.Caption = AppendEllipsesToSingleMenu(.mnuRestoreValues.Caption, False)
        .mnuAbout.Caption = AppendEllipsesToSingleMenu(.mnuAbout.Caption, False)
        
        ' Special handling for mnuSaveValues
        ' Always append title with an exclamation point
        .mnuSaveValues.Caption = AppendEllipsesToSingleMenu(.mnuSaveValues.Caption, True, True)
        
        ' Now add shortcut key tips to menus
        .mnuExit.Caption = AppendShortcutKeyToDynamicMenu(.mnuExit.Caption)
        .mnuCut.Caption = AppendShortcutKeyToSingleMenu(.mnuCut.Caption, "Ctrl+X")
        .mnuCopy.Caption = AppendShortcutKeyToSingleMenu(.mnuCopy.Caption, "Ctrl+C")
        .mnuPaste.Caption = AppendShortcutKeyToSingleMenu(.mnuPaste.Caption, "Ctrl+V")
        .mnuDelete.Caption = AppendShortcutKeyToSingleMenu(.mnuDelete.Caption, "Del")
        
        .mnuRightClickUndo.Caption = AppendShortcutKeyToSingleMenu(.mnuRightClickUndo.Caption, "Ctrl+Z")
        .mnuRightClickCut.Caption = AppendShortcutKeyToSingleMenu(.mnuRightClickCut.Caption, "Ctrl+X")
        .mnuRightClickCopy.Caption = AppendShortcutKeyToSingleMenu(.mnuRightClickCopy.Caption, "Ctrl+C")
        .mnuRightClickPaste.Caption = AppendShortcutKeyToSingleMenu(.mnuRightClickPaste.Caption, "Ctrl+V")
        .mnuRightClickDelete.Caption = AppendShortcutKeyToSingleMenu(.mnuRightClickDelete.Caption, "Del")
        .mnuRightClickSelectAll.Caption = AppendShortcutKeyToSingleMenu(.mnuRightClickSelectAll.Caption, "Ctrl+A")
        
        .mnuEmpirical.Caption = AppendShortcutKeyToSingleMenu(.mnuEmpirical.Caption, "Ctrl+E")
        
        .mnuViewType(vmdMultiView).Caption = AppendShortcutKeyToSingleMenu(.mnuViewType(vmdMultiView).Caption, "F8")
        .mnuViewType(vmdSingleView).Caption = AppendShortcutKeyToSingleMenu(.mnuViewType(vmdSingleView).Caption, "F8")
        .mnuPercentSolver.Caption = AppendShortcutKeyToSingleMenu(.mnuPercentSolver.Caption, "F11")
    End With
    
    With frmCapillaryCalcs
        .mnuLoadCapValues.Caption = AppendEllipsesToSingleMenu(.mnuLoadCapValues.Caption, False)
        .mnuSaveCapValues.Caption = AppendEllipsesToSingleMenu(.mnuSaveCapValues.Caption, False)
    End With
    
    With frmFragmentationModelling
        .mnuLoadSequenceInfo.Caption = AppendEllipsesToSingleMenu(.mnuLoadSequenceInfo.Caption, False)
        .mnuSaveSequenceInfo.Caption = AppendEllipsesToSingleMenu(.mnuSaveSequenceInfo.Caption, False)
        .mnuLoadIonList.Caption = AppendEllipsesToSingleMenu(.mnuLoadIonList.Caption, False)
        .mnuIonMatchListOptions.Caption = AppendEllipsesToSingleMenu(.mnuIonMatchListOptions.Caption, False)
        .mnuShowMassSpectrum.Caption = AppendEllipsesToSingleMenu(.mnuShowMassSpectrum.Caption, False)
        .mnuAutoAlign.Caption = AppendEllipsesToSingleMenu(.mnuAutoAlign.Caption, False)
    End With
    
'''    With frmMsPlot
'''        .mnuExportData.Caption = AppendEllipsesToSingleMenu(.mnuExportData.Caption, False)
'''        .mnuSetResolution.Caption = AppendEllipsesToSingleMenu(.mnuSetResolution.Caption, False)
'''        .mnuTicksXAxis.Caption = AppendEllipsesToSingleMenu(.mnuTicksXAxis.Caption, False)
'''        .mnuTicksYAxis.Caption = AppendEllipsesToSingleMenu(.mnuTicksYAxis.Caption, False)
'''        .mnuGaussianQuality.Caption = AppendEllipsesToSingleMenu(.mnuGaussianQuality.Caption, False)
'''        .mnuApproximationFactor.Caption = AppendEllipsesToSingleMenu(.mnuApproximationFactor.Caption, False)
'''        .mnuSetRangeX.Caption = AppendEllipsesToSingleMenu(.mnuSetRangeX.Caption, False)
'''        .mnuSetRangeY.Caption = AppendEllipsesToSingleMenu(.mnuSetRangeY.Caption, False)
'''        .mnuResetToDefaults.Caption = AppendEllipsesToSingleMenu(.mnuResetToDefaults.Caption, False)
'''
'''        ' Special shortcut keys with language-specific captions
'''        .mnuZoomOutToPrevious.Caption = .mnuZoomOutToPrevious.Caption & vbTab & LookupLanguageCaption(13295, "Ctrl+Z or Right Click")
'''        .mnuCursorMode.Caption = .mnuCursorMode.Caption & vbTab & LookupLanguageCaption(13315, "Space Enables Move")
'''        .mnuZoomIn.Caption = .mnuZoomIn.Caption & vbTab & LookupLanguageCaption(13365, "Left Click")
'''    End With
    
End Sub

Private Function AppendShortcutKeyToDynamicMenu(strMenuCaption As String, Optional intExitMode As Integer = -1) As String
    ' This sub either appends a new shortcut key to the Exit menu
    ' Must use different logic than in sub AppendShortcutKeyToSingleMenu
    '  since sometimes the exit menu is Exit! and other times Exit... (! means immediate exit)
    
    Dim strNewShortcut As String, eWorkingExitMode As exmExitModeConstants, intIndex As Integer
    Dim intOldCaptionShortcutStartIndex As Integer, strOriginalCaption As String
    
    If intExitMode >= exmEscapeKeyConfirmExit And intExitMode <= exmIgnoreEscapeKeyDoNotConfirmExit Then
        ' Sub called due to changing exit mode
        eWorkingExitMode = intExitMode
    Else
        ' Sub called due to changing caption or appending shortcut to caption
        eWorkingExitMode = exmEscapeKeyConfirmExit
        For intIndex = exmEscapeKeyConfirmExit To exmIgnoreEscapeKeyDoNotConfirmExit
            If frmProgramPreferences.optExitConfirmation(intIndex).value = True Then
                eWorkingExitMode = intIndex
                Exit For
            End If
        Next intIndex
    End If
    
    ' Construct new shortcut
    Select Case eWorkingExitMode
    Case exmEscapeKeyConfirmExit
        strNewShortcut = "..." & vbTab & "Esc or Alt+F4"
    Case exmEscapeKeyDoNotConfirmExit
        strNewShortcut = "!" & vbTab & "Esc or Alt+F4"
    Case exmIgnoreEscapeKeyConfirmExit
        strNewShortcut = "..." & vbTab & "Alt+F4"
    Case Else
        ' Includes exmIgnoreEscapeKeyDoNotConfirmExit
        strNewShortcut = "!" & vbTab & "Alt+F4"
    End Select
    
    ' Find current caption
    intOldCaptionShortcutStartIndex = InStr(strMenuCaption, "...")
    If intOldCaptionShortcutStartIndex <= 0 Then
        intOldCaptionShortcutStartIndex = InStr(strMenuCaption, "!")
    End If
    
    If intOldCaptionShortcutStartIndex > 0 Then
        strOriginalCaption = Left(strMenuCaption, intOldCaptionShortcutStartIndex - 1)
    Else
        strOriginalCaption = strMenuCaption
    End If
    
    ' Combine the current caption with the correct shortcut
    AppendShortcutKeyToDynamicMenu = strOriginalCaption & strNewShortcut

End Function

Public Function AppendShortcutKeyToSingleMenu(strMenuCaption As String, strShortcutKey As String) As String
    ' This sub appends shortcut key to a menu
    
    Dim intTabIndex As Integer, strOriginalCaption  As String
    
    intTabIndex = InStr(strMenuCaption, vbTab)
    If intTabIndex > 0 Then
        strOriginalCaption = Left(strMenuCaption, intTabIndex - 1)
    Else
        strOriginalCaption = strMenuCaption
    End If
    
    AppendShortcutKeyToSingleMenu = strOriginalCaption & vbTab & strShortcutKey
    
End Function

Public Sub BatchProcessTextFile(Optional strInputFilename As String = "", Optional strOutputFilename As String = "", Optional boolOverwriteWithoutAsking As Boolean = False)
    ' Use Open dialog to choose file
    Dim strFileName As String
    Dim strWork As String, strPeptide3Letter As String
    Dim strOutLine As String
    Dim strMessage As String
    Dim lngBytesRead As Long, lngFormulasProcessed As Long
    Dim eResponse As VbMsgBoxResult
    
    Dim intIndex As Integer
    Dim strCommand As String, strSettings As String, strSettingsUCase As String
    Dim strStatus As String
    Dim strCustomElementFormula As String
    Dim strDelimeter As String
    Dim blnShowCapitalizedFormula As Boolean
    Dim blnConvertToEmpiricalFormula As Boolean
    Dim blnExpandAbbreviations As Boolean
    Dim blnStdDevModeEnabled As Boolean
    Dim eStdDevModeSaved As smStdDevModeConstants
    Dim eElementModeSaved As emElementModeConstants
    Dim blnShowWeight As Boolean
    Dim blnOneLetterPeptideWeightMode As Boolean
    Dim strPeptideWeightModePrefixFormula As String      ' Atoms that make up the prefix group for Peptide weight mode; default is H
    Dim strPeptideWeightModeSuffixFormula As String      ' Atoms that make up the suffix group for Peptide weight mode; default is OH
    Dim blnAAConvert1to3 As Boolean, blnAAConvert3to1 As Boolean
    Dim blnSpaceEvery10 As Boolean, blnSeparateWithDash As Boolean
    Dim blnShowInputSequence As Boolean
    Dim blnShowSourceFormula As Boolean
    Dim blnEchoComments As Boolean
    Dim blnVerboseMode As Boolean
    
    Dim blnFFMode As Boolean
    Dim intValsFound As Integer
    Dim intElementsChecked As Integer, intCustomElementIndex As Integer
    Dim intMatchLoc As Integer
    Dim lngNewMaxHits As Long
    
    Dim blnIsotopicDistributionMode As Boolean
    Dim intIsotopicDistributionChargeState As Integer
    Dim strResults As String
    Dim dblDummyArray() As Double
    Dim lngErrorID As Long
    
    Dim strHeaderIsotopicAbundace As String
    Dim strHeaderMass As String
    Dim strHeaderFraction As String
    Dim strHeaderIntensity As String
    
    Dim InFileNum As Integer, OutFileNum As Integer
    
    Dim objCompound As New MWCompoundClass
    
    Const MAX_SEARCH_ELEMENTS = 20
    Dim strSearchElements(MAX_SEARCH_ELEMENTS) As String, strRemaining As String
    Dim strThisElement As String, strSearchElementsSelected As String
    
    ' Save the current standard deviation mode and retrieve the current state of displaying the StdDev Mode
    eElementModeSaved = objMwtWin.GetElementMode
    eStdDevModeSaved = objMwtWin.StdDevMode
    blnStdDevModeEnabled = gBlnShowStdDevWithMass
    
    ' Set default output options
    strDelimeter = vbTab
    blnShowWeight = True
    blnShowInputSequence = True
    blnShowSourceFormula = True
    blnEchoComments = False
    blnVerboseMode = True
    
    ' Default AA weight mode prefix and suffix
    strPeptideWeightModePrefixFormula = "H"
    strPeptideWeightModeSuffixFormula = "OH"
    
    ' Isotopic Distribution Options
    intIsotopicDistributionChargeState = 1
    strHeaderIsotopicAbundace = LookupLanguageCaption(15200, "Isotopic Abundances for")
    strHeaderMass = LookupLanguageCaption(15210, "Mass/Charge")
    strHeaderFraction = LookupLanguageCaption(15220, "Fraction")
    strHeaderIntensity = LookupLanguageCaption(15230, "Intensity")

    If strInputFilename = "" Then
        ' 1510 = Text Files, 1515 = .txt
        strFileName = SelectFile(frmMain.hwnd, "Select File", gLastFileOpenSaveFolder, False, "", ConstructFileDialogFilterMask(LookupMessage(1510), LookupMessage(1515)), 1)
        If Len(strFileName) = 0 Then
            ' No file selected (or other error)
            Exit Sub
        End If
        strOutputFilename = strFileName & ".out"
    Else
        strFileName = strInputFilename
        
        If Not FileExists(strInputFilename) Then
            MsgBox LookupMessage(470, " (" & strInputFilename & ")"), vbOKOnly + vbExclamation, LookupMessage(480)
            Exit Sub
        End If
        If strOutputFilename = "" Or strOutputFilename = strFileName Then
            strOutputFilename = strFileName & ".out"
        End If
    End If
    
    On Error GoTo BatchFileProb
    
    ' Open the file for input
    InFileNum = FreeFile()
    Open strFileName For Input As #InFileNum

    If Not boolOverwriteWithoutAsking And (frmProgramPreferences.optExitConfirmation(exmEscapeKeyConfirmExit).value = True Or frmProgramPreferences.optExitConfirmation(exmIgnoreEscapeKeyConfirmExit).value = True) Then
        If FileExists(strOutputFilename) Then
            eResponse = MsgBox(strOutputFilename & ": " & LookupMessage(490), vbYesNoCancel + vbDefaultButton2 + vbExclamation, LookupMessage(500))
            If eResponse <> vbYes Then Exit Sub
        End If
    End If
    
    OutFileNum = FreeFile()
    Open strOutputFilename For Output As #OutFileNum
    
    ' Change mouse pointer to hourglass
    frmMain.MousePointer = vbArrowHourglass
    
    blnFFMode = False
    blnIsotopicDistributionMode = False
    
    frmProgress.InitializeForm "Batch analyzing", 0, FileLen(strFileName), True, False, True
    frmProgress.ToggleAlwaysOnTop True

    Do Until EOF(InFileNum)
        Line Input #InFileNum, strWork
        lngBytesRead = lngBytesRead + Len(strWork) + 2
        
        If Trim(strWork) = "" Or IsComment(strWork) Then
            If blnEchoComments Then
                Print #OutFileNum, strWork
            End If
        Else
            
            ' Examine strWork to see if it contains an = sign
            ' If so, then it is a batch processing Command
            intMatchLoc = InStr(strWork, "=")
            If intMatchLoc > 0 Then
                ' Command Found
                strCommand = Trim(UCase(Left(strWork, intMatchLoc - 1)))
                strSettings = Trim(Mid(strWork, intMatchLoc + 1))
                strSettingsUCase = UCase(strSettings)
                
                Select Case Trim(strCommand)
            ' Output option commands
                Case "VERBOSEMODE"
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        Print #OutFileNum, COMMENT_CHAR & " Verbose mode is now on"
                        blnVerboseMode = True
                    Case Else
                        blnVerboseMode = False
                    End Select
                
                Case "DELIMETER"
                    If Left(strSettings, 1) = "<" And Right(strSettings, 1) = ">" Then
                        ' Special delimeter
                        Select Case UCase(Mid(strSettings, 2, Len(strSettings) - 2))
                        Case "TAB"
                            strDelimeter = vbTab
                            If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Delimeter now a Tab"
                        Case "SPACE"
                            strDelimeter = " "
                            If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Delimeter now a Space"
                        Case "ENTER", "CRLF"
                            strDelimeter = vbCrLf
                            If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Delimeter now a Carriage Return (Enter)"
                        Case Else
                            Print #OutFileNum, COMMENT_CHAR & " Unknown delimeter code: " & strSettings & " -- Should be one of the following: <TAB>, <SPACE>, <ENTER>, <CRLF>"
                        End Select
                    ElseIf Len(Mid(strWork, intMatchLoc + 1)) > 0 Then
                        ' Normal text delimeter (one or more characters)
                        strDelimeter = Mid(strWork, intMatchLoc + 1)
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Delimeter now " & strDelimeter
                    Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Delimeter reset to default (Tab)"
                        strDelimeter = vbTab
                    End If
                Case "ECHOCOMMENTS"
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Will now write comments present in the source file to the output file"
                        blnEchoComments = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Comments found in the source file will not be written to the output file"
                        blnEchoComments = False
                    End Select
                    
            ' Molecular Weight commands
                Case "MW"
                    If blnFFMode Then
                        frmFinder.Hide
                        blnFFMode = False
                    End If
                    blnIsotopicDistributionMode = False
                    blnShowWeight = True
                    blnAAConvert3to1 = False
                    blnAAConvert1to3 = False
                    blnConvertToEmpiricalFormula = False
                    blnExpandAbbreviations = False
                    blnOneLetterPeptideWeightMode = False
                    If blnVerboseMode Then
                        Print #OutFileNum, ""
                        Print #OutFileNum, COMMENT_CHAR & " Normal Molecular Weight Mode Enabled (other modes turned Off)"
                    End If
                Case "CAPITALIZED"
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Source formula will be displayed with proper capitalization"
                        blnShowCapitalizedFormula = True
                        blnShowSourceFormula = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Source formula will be displayed exactly as found in the input file"
                        blnShowCapitalizedFormula = False
                    End Select
                Case "MWSHOWSOURCEFORMULA"
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Display of source formula is now On"
                        blnShowSourceFormula = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Display of source formula is now Off"
                        blnShowSourceFormula = False
                    End Select
                Case "EMPIRICAL"
                    If blnVerboseMode Then Print #OutFileNum, ""
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        ' Enable conversion of formulas to their empirical formulas
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Converting formulas to empirical formulas now On"
                        blnConvertToEmpiricalFormula = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Converting formulas to empirical formulas now Off"
                        blnConvertToEmpiricalFormula = False
                    End Select
                Case "EXPANDABBREVIATIONS"
                    If blnVerboseMode Then Print #OutFileNum, ""
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        ' Enable expansion of abbreviations
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Abbreviation expansion now On"
                        blnExpandAbbreviations = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Abbreviation expansion now Off"
                        blnExpandAbbreviations = False
                    End Select
                Case "SHOWWEIGHT"
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        ' Show molecular weight
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Will display the molecular weight (mass) of each formula"
                        blnShowWeight = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Will not display the molecular weight (mass) of each formula"
                        blnShowWeight = False
                    End Select
                Case "STDDEVMODE"
                    Select Case strSettingsUCase
                    Case "SHORT"
                        objMwtWin.StdDevMode = smShort
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Standard deviation mode using Short display"
                        blnStdDevModeEnabled = True
                    Case "SCIENTIFIC"
                        objMwtWin.StdDevMode = smScientific
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Standard deviation mode using Scientific display"
                        blnStdDevModeEnabled = True
                    Case "DECIMAL"
                        objMwtWin.StdDevMode = smDecimal
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Standard deviation mode using Decimal display"
                        blnStdDevModeEnabled = True
                    Case "OFF"
                        blnStdDevModeEnabled = False
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Standard deviations will not be displayed"
                    Case Else
                        Print #OutFileNum, COMMENT_CHAR & " Warning: Invalid standard deviation mode: " & strSettingsUCase & " -- Should be one of the following: SHORT, SCIENTIFIC, DECIMAL, OFF"
                    End Select
                Case "WEIGHTMODE"
                    Select Case strSettingsUCase
                    Case "AVERAGE"
                        SwitchWeightMode emAverageMass, False
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Average Weight Mode Enabled"
                    Case "ISOTOPIC"
                        SwitchWeightMode emIsotopicMass, False
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Isotopic Weight Mode Enabled"
                    Case "INTEGER"
                        SwitchWeightMode emIntegerMass, False
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Integer Weight Mode Enabled"
                    Case Else
                        Print #OutFileNum, COMMENT_CHAR & " Warning: Invalid elemental weight mode: " & strSettingsUCase & " -- Should be one of the following: AVERAGE, ISOTOPIC, INTEGER"
                    End Select
                Case "ONELETTERPEPTIDEWEIGHTMODE"
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        ' Show molecular weight
                        If blnVerboseMode Then
                            Print #OutFileNum, ""
                            Print #OutFileNum, COMMENT_CHAR & " One letter Amino Acid weight mode: input formulas are assumed to be peptides in one-letter notation"
                        End If
                        
                        If blnFFMode Then
                            frmFinder.Hide
                            blnFFMode = False
                        End If
                        blnIsotopicDistributionMode = False
                        blnShowWeight = True
                        blnAAConvert3to1 = False
                        blnAAConvert1to3 = False
                        blnConvertToEmpiricalFormula = False
                        blnExpandAbbreviations = False
                        blnOneLetterPeptideWeightMode = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " "
                        blnOneLetterPeptideWeightMode = False
                    End Select
                Case "PEPTIDEWEIGHTMODEPEPTIDEPREFIX"
                    strPeptideWeightModePrefixFormula = strSettings
                    
                Case "PEPTIDEWEIGHTMODEPEPTIDESUFFIX"
                    strPeptideWeightModeSuffixFormula = strSettings
                    
            ' Amino acid notation conversion commands
                Case "AACONVERT3TO1"
                    ' Treat strSettings as an amino acid with 3 letter symbols
                    ' Convert to 1 letter symbols
                    If blnVerboseMode Then Print #OutFileNum, ""
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " 3 letter to 1 letter amino acid symbol conversion now On"
                        blnAAConvert3to1 = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " 3 letter to 1 letter amino acid symbol conversion now Off"
                        blnAAConvert3to1 = True
                    End Select
                    blnAAConvert1to3 = False
                    blnFFMode = False
                    blnIsotopicDistributionMode = False
                Case "AACONVERT1TO3"
                    ' Treat strSettings as an amino acid with 1 letter symbols
                    ' Convert to 3 letter symbols
                    If blnVerboseMode Then Print #OutFileNum, ""
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " 1 letter to 3 letter amino acid symbol conversion now On"
                        blnAAConvert1to3 = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " 1 letter to 3 letter amino acid symbol conversion Off"
                        blnAAConvert1to3 = False
                    End Select
                    blnAAConvert3to1 = False
                    blnFFMode = False
                    blnIsotopicDistributionMode = False
                Case "AASPACEEVERY10"
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Will add a space every 10 amino acids"
                        blnSpaceEvery10 = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Will not add a space every 10 amino acids"
                        blnSpaceEvery10 = False
                    End Select
                Case "AA1TO3USEDASH"
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Will separate residues with a dash"
                        blnSeparateWithDash = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Will not separate residues with a dash"
                        blnSeparateWithDash = False
                    End Select
                Case "AASHOWSEQUENCEBEINGCONVERTED"
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Will show sequence being converted, in addition to the converted sequence"
                        blnShowInputSequence = True
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Will only show the converted sequence, not the sequence being converted"
                        blnShowInputSequence = False
                    End Select
            ' Formula Finder commands
                Case "FF"
                    If Not blnFFMode Then
                        blnFFMode = True
                        frmFinder.Show
                        ' Need to re-initialize to make sure the progress window is on top
                        frmProgress.UpdateCurrentTask "Batch analyzing"
                    End If
                    
                    frmFinder.SetWeightMatchingMode 0
                    ' Examine strSettings to see if any elements/abbreviations are specified
                    ' If yes, construct an array of the elements/abbreviations
                    ' Once constructed, select the appropriate elements on the form

                    If Len(strSettings) > 0 Then
                        intValsFound = ParseString(strSettings, strSearchElements, MAX_SEARCH_ELEMENTS, ",", strRemaining, True, True)
                        If intValsFound > 0 Then
                            ' Uncheck all of the currently selected search elements on frmFinder
                            For intIndex = 0 To 9
                                frmFinder.chkElements(intIndex).value = vbUnchecked
                            Next intIndex
                            
                            intCustomElementIndex = 3
                            For intIndex = 1 To intValsFound
                                strCustomElementFormula = strSearchElements(intIndex)
                                
                                Select Case Trim(UCase(strCustomElementFormula))
                                Case "C":  frmFinder.chkElements(0).value = vbChecked
                                Case "H":  frmFinder.chkElements(1).value = vbChecked
                                Case "N":  frmFinder.chkElements(2).value = vbChecked
                                Case "O":  frmFinder.chkElements(3).value = vbChecked
                                Case Else
                                    If intCustomElementIndex < 9 Then
                                        ' See if strCustomElementFormula is valid by checking if it has a weight of 0 or more
                                        strCustomElementFormula = UCase(Left(strCustomElementFormula, 1)) & Mid(strCustomElementFormula, 2)
                                        
                                        objCompound.Formula = strCustomElementFormula
                                        If objCompound.ErrorID = 0 Then
                                            intCustomElementIndex = intCustomElementIndex + 1
                                            frmFinder.chkElements(intCustomElementIndex).value = vbChecked
                                            
                                            frmFinder.txtWeight(intCustomElementIndex).Text = objCompound.FormulaCapitalized
                                        Else
                                            Print #OutFileNum, COMMENT_CHAR & " Error: Invalid formula or abbreviation: " & strCustomElementFormula
                                        End If
                                    Else
                                        Print #OutFileNum, COMMENT_CHAR & " Error: Too many custom search elements (" & intValsFound & ")"
                                        Exit For
                                    End If
                                End Select
                            Next intIndex
                        End If
                    End If
                    
                    strSearchElementsSelected = ""
                    For intIndex = 0 To 9
                        If cChkBox(frmFinder.chkElements(intIndex)) Then
                            Select Case intIndex
                            Case 0: strThisElement = "C"
                            Case 1: strThisElement = "H"
                            Case 2: strThisElement = "N"
                            Case 3: strThisElement = "O"
                            Case Else
                                strThisElement = frmFinder.txtWeight(intIndex)
                            End Select
                            If Len(strSearchElementsSelected) > 0 Then
                                strSearchElementsSelected = strSearchElementsSelected & ", " & strThisElement
                            Else
                                strSearchElementsSelected = strThisElement
                            End If
                        End If
                    Next intIndex

                    If blnVerboseMode Then Print #OutFileNum, ""
                    If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Formula Finder Mode Enabled.  Search elements/abbreviations: " & strSearchElementsSelected

                Case "MAXHITS"
                    If IsNumeric(strSettings) Then
                        lngNewMaxHits = CLng(strSettings)
                        If lngNewMaxHits >= 1 And lngNewMaxHits <= 15000 Then
                            frmFinder.txtHits = lngNewMaxHits
                            If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " FF Maximum Hits set to " & Trim(lngNewMaxHits)
                        End If
                    End If
                Case "TOLERANCE"
                    If IsNumeric(strSettings) Then
                        frmFinder.txtWeightTolerance = CSng(strSettings)
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " FF Tolerance set to " & Trim(CSng(strSettings))
                    End If
                
                Case "ISOTOPICDISTRIBUTION"
                    ' Isotopic Distribution calculations
                    If blnVerboseMode Then Print #OutFileNum, ""
                    Select Case strSettingsUCase
                    Case "TRUE", "1", "ON"
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Isotopic Distribution calculations now On"
                        blnIsotopicDistributionMode = True
                        blnShowWeight = False
                    Case Else
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Isotopic Distribution calculations now Off"
                        blnIsotopicDistributionMode = False
                        blnShowWeight = True
                    End Select
                    
                    blnFFMode = False
                    blnAAConvert3to1 = False
                    blnAAConvert1to3 = False
                    blnConvertToEmpiricalFormula = False
                    blnExpandAbbreviations = False
                    blnOneLetterPeptideWeightMode = False
                    
                Case "ISOTOPICDISTRIBUTIONCHARGE"
                    ' Charge state for Isotopic Distribution calculations
                    If IsNumeric(strSettings) Then
                        intIsotopicDistributionChargeState = CInt(strSettings)
                        If blnVerboseMode Then Print #OutFileNum, COMMENT_CHAR & " Isotopic Distribution charge set to " & Trim(intIsotopicDistributionChargeState)
                    End If
                Case Else
                    Print #OutFileNum, COMMENT_CHAR & " Error: Unknown Command: " & strCommand
                End Select
            Else
                If KeyPressAbortProcess > 1 Then
                    Print #OutFileNum, COMMENT_CHAR & " Error: Processing aborted"
                    Exit Do
                End If
                
                lngFormulasProcessed = lngFormulasProcessed + 1
                If blnFFMode Then
                    ' Using Formula Finder
                    ' Only parse line if it contains a number
                    If IsNumeric(strWork) Then
                        ' Make sure at least one element is checked
                        intElementsChecked = 0
                        For intIndex = 0 To 9
                            If cChkBox(frmFinder.chkElements(intIndex)) Then
                                intElementsChecked = intElementsChecked + 1
                            End If
                        Next intIndex
                        If intElementsChecked = 0 Then
                            Print #OutFileNum, COMMENT_CHAR & " Warning: Formula Finder Mode is Enabled, but no elements were chosen."
                            Print #OutFileNum, COMMENT_CHAR & "          Will use C, H, N, and O by default"
                            frmFinder.chkElements(0).value = vbChecked
                            frmFinder.chkElements(1).value = vbChecked
                            frmFinder.chkElements(2).value = vbChecked
                            frmFinder.chkElements(3).value = vbChecked
                        End If
                        
                        If CSng(strWork) > 0 Then
                            Print #OutFileNum, COMMENT_CHAR & " FF Searching: " & CSng(strWork)
                            frmFinder.txtMWT = CSng(strWork)
                            frmFinder.rtfResults = ""
                            
                            ' Start the Formula Finder in batch mode
                            strStatus = frmFinder.FormulaFinderCalculate(True)
                            
                            If strStatus <> "" Then
                                intMatchLoc = InStr(strStatus, vbCrLf)
                                If intMatchLoc > 0 Then
                                    strStatus = Left(strStatus, intMatchLoc - 1)
                                End If
                                Print #OutFileNum, COMMENT_CHAR & " Error: " & strStatus
                            End If
                            ' Now print the results
                            Print #OutFileNum, frmFinder.rtfResults.Text
                            Print #OutFileNum, ""
                        Else
                            Print #OutFileNum, COMMENT_CHAR & " Error: Cannot use Formula Finder with a search value of 0"
                        End If
                    Else
                        Print #OutFileNum, COMMENT_CHAR & " Error: Formula Finder Mode is Enabled, but number not found: " & strWork
                    End If
                Else
                    ' Not formula finder mode
                    strOutLine = ""
                    If blnAAConvert3to1 Then
                        ' Convert 3 letter to 1 letter amino acids
                        objMwtWin.Peptide.SetSequence strWork
                        
                        If blnShowInputSequence Then strOutLine = strWork & strDelimeter
                        strOutLine = strOutLine & objMwtWin.Peptide.GetSequence(False, blnSpaceEvery10, False, False, True)
                        
                    ElseIf blnAAConvert1to3 Then
                        ' Convert 1 letter to 3 letter amino acids
                        objMwtWin.Peptide.SetSequence strWork, ntgHydrogen, ctgHydroxyl, False
                        
                        If blnShowInputSequence Then strOutLine = strWork & strDelimeter
                        strOutLine = strOutLine & objMwtWin.Peptide.GetSequence(True, blnSpaceEvery10, blnSeparateWithDash, False, True)
                        
                    ElseIf blnIsotopicDistributionMode Then
                        If intIsotopicDistributionChargeState < 1 Then
                            intIsotopicDistributionChargeState = 1
                        End If
                        
                        lngErrorID = objMwtWin.ComputeIsotopicAbundances(strWork, intIsotopicDistributionChargeState, strResults, dblDummyArray, 0, strHeaderIsotopicAbundace, strHeaderMass, strHeaderFraction, strHeaderIntensity)

                        If lngErrorID = 0 Then
                             ' Formula parsed successfuly
                             strOutLine = strResults
                        ElseIf lngErrorID <> -1 Then
                             strOutLine = LookupMessage(350) & ": " & LookupMessage(lngErrorID)
                        End If
                        
                    Else
                        ' Compute the mass of a formula or of a 1-letter peptide sequence
                        If blnOneLetterPeptideWeightMode Then
                            ' Convert strWork to 3-letter notation, then add
                            '   strPeptideWeightModePrefixFormula and strPeptideWeightModeSuffixFormula
                            
                            objMwtWin.Peptide.SetSequence strWork, ntgHydrogen, ctgHydroxyl, False
                            strPeptide3Letter = objMwtWin.Peptide.GetSequence(True, False, False, False, False)
                            strPeptide3Letter = strPeptideWeightModePrefixFormula & "-" & strPeptide3Letter & "-" & strPeptideWeightModeSuffixFormula
                            
                            objCompound.Formula = strPeptide3Letter
                        Else
                            objCompound.Formula = strWork
                        End If
                        
                        If objCompound.ErrorID = 0 Then
                            
                            If Not (blnShowCapitalizedFormula Or blnShowSourceFormula Or blnConvertToEmpiricalFormula Or blnExpandAbbreviations Or blnShowWeight) Then
                                ' Need to make sure blnShowWeight = True so that something gets displayed
                                blnShowWeight = True
                            End If
                            
                            If blnShowSourceFormula Then
                                If blnShowCapitalizedFormula Then
                                    strOutLine = objCompound.FormulaCapitalized
                                Else
                                    strOutLine = strWork
                                End If
                            Else
                                strOutLine = ""
                            End If
                            
                            If blnConvertToEmpiricalFormula Then
                                If Len(strOutLine) > 0 Then strOutLine = strOutLine & strDelimeter
                                strOutLine = strOutLine & objCompound.ConvertToEmpirical
                            ElseIf blnExpandAbbreviations Then
                                If Len(strOutLine) > 0 Then strOutLine = strOutLine & strDelimeter
                                strOutLine = strOutLine & objCompound.ExpandAbbreviations
                            End If
                            
                            If blnShowWeight Then
                                If Len(strOutLine) > 0 Then strOutLine = strOutLine & strDelimeter
                                If blnStdDevModeEnabled Then
                                    strOutLine = strOutLine & objCompound.MassAndStdDevString
                                Else
                                    strOutLine = strOutLine & objCompound.Mass
                                End If
                            End If
                        Else
                            ' If an error, output the error to the file
                            strOutLine = strWork & strDelimeter & LookupMessage(objCompound.ErrorID)
                        End If
                    End If
                    
                    ' Write strOutLine to the output file
                    Print #OutFileNum, strOutLine
                
                End If
            End If
        End If
        
        If lngFormulasProcessed Mod 50 = 0 Then
            frmProgress.UpdateProgressBar lngBytesRead
            frmProgress.UpdateCurrentSubTask "Processed " & Trim(lngFormulasProcessed) & " formulas"
        End If
    Loop
    Close #InFileNum
    Close #OutFileNum
    
    frmProgress.HideForm
    
    If blnFFMode Then frmFinder.Hide
    
    Set objCompound = Nothing
    
    ' Change mouse pointer to default
    frmMain.MousePointer = vbDefault
   
ExitBatch:
    ' Restore the standard deviation mode to the saved value
    objMwtWin.StdDevMode = eStdDevModeSaved
    SwitchWeightMode eElementModeSaved, False

    Exit Sub

BatchFileProb:
    Close
    strMessage = LookupMessage(510) & ": " & strFileName & " or " & strOutputFilename
    strMessage = strMessage & vbCrLf & Err.Description
    strMessage = strMessage & vbCrLf & LookupMessage(515)
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)
    Resume ExitBatch

End Sub

Public Function ChargeValueToString(sngChargeValue As Single) As String
    Dim strChargeString As String
    
    strChargeString = Trim(CStr(sngChargeValue))
    If InStr(strChargeString, glbDecimalSeparator) Then
        strChargeString = Trim(Format(sngChargeValue, "#0.0##"))
    End If
    
    If sngChargeValue > 0 Then
        ChargeValueToString = strChargeString & "+"
    ElseIf sngChargeValue < 0 Then
        ChargeValueToString = Trim(CStr(Abs(sngChargeValue))) & "-"
    Else
        ChargeValueToString = "0"
    End If

End Function

Public Function ConstructAddFormulaToolTip() As String
    If frmMain.GetViewMode = vmdMultiView Then
        ConstructAddFormulaToolTip = LookupToolTipLanguageCaption(5110, "Adds a new, blank formula to the list")
    Else
        ConstructAddFormulaToolTip = LookupToolTipLanguageCaption(5115, "Displays a new, blank formula")
    End If
End Function

Public Function ConstructFormulaLabel(intFormulaID As Integer) As String
    ' Note: intFormulaID is 0 based, ranging from 0 to MAX_FORMULAS-1
    ConstructFormulaLabel = LookupLanguageCaption(5000, "Formula") & " " & Trim(Str(intFormulaID + 1)) & ":"
End Function

Public Function ConstructMWEqualsCaption() As String
    ConstructMWEqualsCaption = gMWAbbreviation & " = "
End Function

Public Sub CopyCurrentMWT(blnNumberOnly As Boolean)
    ' If blnNumberOnly is True then only the numerical value is copied
    Dim strWork As String, EqualLoc As Integer, ParenthLoc As Integer
    
    Clipboard.Clear
    
    ' If standard deviation mode is set to off, copy the molecular weight as just
    ' a value, rather than MW = 00000.000
    
    strWork = frmMain.txtMWT(frmMain.GetCurrentFormulaIndex).Text
    If gBlnShowStdDevWithMass = False Or blnNumberOnly Then
        EqualLoc = InStr(strWork, "=")
        If EqualLoc > 0 Then
            strWork = Trim(Mid(strWork, EqualLoc + 1))
        End If
        If blnNumberOnly Then
            ' Must remove the standard deviation value from strWork
            ParenthLoc = InStr(strWork, "(")
            If ParenthLoc > 0 Then strWork = Left(strWork, ParenthLoc - 1)
        End If
    End If
    
    Clipboard.SetText strWork, vbCFText

End Sub

Private Sub FlexGridAddItems(grdThisGrid As MSFlexGrid, strItem1 As String, strItem2 As String, Optional strItem3 As String = "", Optional strItem4 As String = "", Optional strItem5 As String = "")
    With grdThisGrid
        .AddItem strItem1
        .TextMatrix(.Rows - 1, 1) = strItem2
        If Len(strItem3) > 0 Then .TextMatrix(.Rows - 1, 2) = strItem3
        If Len(strItem4) > 0 Then .TextMatrix(.Rows - 1, 3) = strItem4
        If Len(strItem5) > 0 Then .TextMatrix(.Rows - 1, 4) = strItem5
    End With
End Sub

Public Function FlexGridCopyInfo(frmThisForm As VB.Form, grdThisGrid As MSFlexGrid, Optional eCopyMode As gcmGridCopyModeConstants = gcmText, Optional intRowStart As Integer = -1, Optional intRowEnd As Integer = -1, Optional intColStart As Integer = -1, Optional intColEnd As Integer = -1, Optional ByVal blnIncludeHeaders As Boolean = True) As String
    ' eCopyMode can be 0 = text, 1 = RTF, or 2 = Html
    
    Const FONT_SIZE = 10
    
    Dim intRowSaved As Integer, intColSaved As Integer
    Dim intRowIndex As Integer, intColIndex As Integer
    Dim strCopyText As String, lngCopyFormat As Long
    Dim blnHeadersCopied As Boolean
    
    Dim strRowPrefix As String, strRowSuffix As String, blnIncludeSuffixInFinalRow As Boolean
    Dim strCellPrefix As String, strCellSuffix As String, blnIncludeSuffixInFinalCol As Boolean
    Dim strHighlightPrefix As String, strHighlightSuffix As String, blnHighlightCell As Boolean
    
    Select Case eCopyMode
    Case gcmHTML
        ' Html Mode
        strRowPrefix = "<TR>" & vbCrLf:  strRowSuffix = "</TR>" & vbCrLf: blnIncludeSuffixInFinalRow = True
        strCellPrefix = "<TD>": strCellSuffix = "</TD>":                  blnIncludeSuffixInFinalCol = True
        strHighlightPrefix = "<B><FONT COLOR=#FF0000>": strHighlightSuffix = "</FONT></B>"
        strCopyText = "<TABLE BORDER=1 CELLSPACING=0 CELLPADDING=2>" & vbCrLf
    Case gcmRTF
        ' RTF Mode
        strRowSuffix = " \par ": blnIncludeSuffixInFinalRow = True
        strCellSuffix = " \tab "
        strHighlightPrefix = "{\b\cf1 ": strHighlightSuffix = "}"
        ' Add the necessary RTF header string to strCopyText
        strCopyText = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman " & "Courier New" & ";}{\f3\froman Times New Roman;}{\f4\fswiss\fprq2 System;}}{\colortbl\red0\green0\blue0;\red255\green0\blue0;\red255\green255\blue255;}\deflang1033\pard\plain\f2\fs" & Trim(Str$(Int(FONT_SIZE * 2))) & " "
    Case Else
        ' Includes gcmText: Text mode
        strCellSuffix = vbTab
        strRowSuffix = vbCrLf
    End Select
    
    frmThisForm.MousePointer = vbHourglass
    
    With grdThisGrid
        intRowSaved = .Row
        intColSaved = .Col
        
        If intRowStart < 0 Then intRowStart = .Row
        If intRowEnd < 0 Then intRowEnd = .RowSel
        
        If intColStart < 0 Then intColStart = .Col
        If intColEnd < 0 Then intColEnd = .ColSel
        
        If intRowStart > intRowEnd Then SwapValues intRowStart, intRowEnd
        If intColStart > intColEnd Then SwapValues intColStart, intColEnd
        
        If (intRowEnd > intRowStart Or blnIncludeHeaders) And eCopyMode = gcmText Then
            blnIncludeSuffixInFinalRow = True
        End If
        
        If intRowStart = 0 Then
            ' Headers will automatically be included
            blnIncludeHeaders = False
        End If
        
        ' Copy info, separating cells in a row using tabs, and separating rows by vbCrLf
        intRowIndex = intRowStart
        Do While intRowIndex <= intRowEnd
            If blnIncludeHeaders And Not blnHeadersCopied Then
                intRowIndex = 0
            End If
            
            strCopyText = strCopyText & strRowPrefix
            For intColIndex = intColStart To intColEnd
                strCopyText = strCopyText & strCellPrefix
                If eCopyMode <> gcmText Then
                    ' Check cell back color
                    .Row = intRowIndex: .Col = intColIndex
                    If .CellBackColor > 0 Then
                        blnHighlightCell = True
                    Else
                        blnHighlightCell = False
                    End If
                End If
                If blnHighlightCell Then strCopyText = strCopyText & strHighlightPrefix
                
                strCopyText = strCopyText & .TextMatrix(intRowIndex, intColIndex)
                
                If blnHighlightCell Then strCopyText = strCopyText & strHighlightSuffix
                
                If intColIndex < intColEnd Or blnIncludeSuffixInFinalCol Then
                    strCopyText = strCopyText & strCellSuffix
                End If
            Next intColIndex
            
            If blnIncludeHeaders And Not blnHeadersCopied Then
                intRowIndex = intRowStart - 1
                If blnIncludeSuffixInFinalRow Then
                    strCopyText = strCopyText & strRowSuffix
                End If
                blnHeadersCopied = True
            Else
                If intRowIndex < intRowEnd Or blnIncludeSuffixInFinalRow Then
                    strCopyText = strCopyText & strRowSuffix
                End If
            End If
            intRowIndex = intRowIndex + 1
        Loop
        
        .Row = intRowSaved
        .Col = intColSaved
    End With
    
    Select Case eCopyMode
    Case gcmHTML
        ' Html Mode
        strCopyText = strCopyText & "</TABLE>" & vbCrLf
        lngCopyFormat = vbCFText
    Case gcmRTF
        ' RTF Mode
        strCopyText = strCopyText & " }"
        lngCopyFormat = vbCFRTF
    Case Else
        ' gcmText
        lngCopyFormat = vbCFText
    End Select
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText strCopyText, lngCopyFormat
    
    FlexGridCopyInfo = strCopyText
    
    frmThisForm.MousePointer = vbDefault

End Function

Public Sub CopyRoutine(frmCallingForm As VB.Form, blnCalledFromMenu As Boolean)
    ' If blnCalledFromMenu = true, then need to manually copy
    ' Otherwise, CTRL+C was pressed and the RTF box already copied the text; must validate it, though
    Dim strCopyText As String
    
On Error GoTo CopyRoutineErrorHandler

    If blnCalledFromMenu Then
        ' Called from menu
        strCopyText = frmCallingForm.ActiveControl.SelText
    Else
        ' Validate copied text; for safety, add a 50 msec delay
        ' This helps to prevent error code 521 = "Can't open clipboard" from occurring
        Sleep 50
        strCopyText = GetClipboardTextSmart()
        If Len(strCopyText) = 0 Then
            ' Abort copy; probably couldn't open clipboard
            Exit Sub
        End If
    End If
        
    RemoveHeightAdjustChar strCopyText
    
    ' Copy corrected text to Clipboard.
    Clipboard.Clear
    Clipboard.SetText strCopyText, vbCFText
    
    Exit Sub

CopyRoutineErrorHandler:
    GeneralErrorHandler "MwtWinProcedures|CopyRoutine", Err.Number, Err.Description

End Sub

Public Sub CutRoutine(frmCallingForm As VB.Form, blnCalledFromMenu As Boolean)
    CopyRoutine frmCallingForm, blnCalledFromMenu
    
    ' If called from a menu, then manually need to clear the selected text
    ' If called by the user pressing Ctrl+x in the RTF box, then selected text is automatically cleared
    If blnCalledFromMenu Then
        frmCallingForm.ActiveControl.SelText = ""
    End If

End Sub

Public Function DetermineMaxAllowableFormulaIndex() As Integer
    Dim intMaxAllowableFormulaIndex As Integer
    
    '    The following table shows the ideal number of formulas to show for standard resolutions:
    '   640x480     6
    '   800x600     8
    '   1024x768    11
    '   1152x864    13
    '   1280x960    15
    '   1280x1024   17
    '   1600x1200   20
    '
    '   I plotted Screen.Height / 15 vs. the above desired values to obtain the
    '   equation used to compute intMaxAllowableFormulaIndex
    
    intMaxAllowableFormulaIndex = Round((Screen.Height / 15) * 0.0199 - 3.8675, 0)
    If intMaxAllowableFormulaIndex < 6 Then
        intMaxAllowableFormulaIndex = 6
    End If
    
    DetermineMaxAllowableFormulaIndex = intMaxAllowableFormulaIndex
    
End Function
    
Public Sub DisplayCurrentFormulaOnSubForm(frmThisForm As VB.Form)
    ' Used by frmCapillaryCalcs and frmMMConvert
    
    Dim intEqualLoc As Integer
    
    With frmThisForm
        .lblFormula.Caption = LookupLanguageCaption(7460, "Current formula is") & ":"
        .rtfCurrentFormula.TextRTF = frmMain.rtfFormula(frmMain.GetCurrentFormulaIndex).TextRTF
        .lblMWTValue.Caption = frmMain.txtMWT(frmMain.GetCurrentFormulaIndex).Text
        intEqualLoc = InStr(.lblMWTValue.Caption, "=")
        If intEqualLoc > 0 Then
            .lblMWTValue.Caption = Mid(.lblMWTValue.Caption, intEqualLoc + 1)
            intEqualLoc = InStr(.lblMWTValue.Caption, "(")
            If intEqualLoc > 0 Then
                .lblMWTValue.Caption = Left(.lblMWTValue.Caption, intEqualLoc - 1)
            End If
            .lblMWTValue.Caption = CStr(CDblSafe(.lblMWTValue.Caption))
        End If
        .lblMWT = LookupLanguageCaption(4040, "MW") & " = "
    End With
    
End Sub

Public Sub ExitProgram()
    Dim eResponse As VbMsgBoxResult
    
    If frmProgramPreferences.optExitConfirmation(exmEscapeKeyConfirmExit).value = True Or frmProgramPreferences.optExitConfirmation(exmIgnoreEscapeKeyConfirmExit).value = True Then
        ' Display the dialog box and get user's Response.
        eResponse = YesNoBox(LookupLanguageCaption(3790, "Are you sure you want quit?"), _
                            LookupLanguageCaption(3795, "Exiting Program"))

        ' Evaluate the user's Response.
        If eResponse = vbYes Then
            ' It's good practice to unload the form before ending the program
            ' Plus, it causes the default settings files to be saved
            Unload frmMain
        End If
    Else
        ' It's good practice to unload the form before ending the program
        ' Plus, it causes the default settings files to be saved
        Unload frmMain
    End If
    
End Sub

Public Function FormatMaximumFormulasToShowString(Optional intDesiredMaxIndex As Integer = -1) As String
    
    If intDesiredMaxIndex = -1 Then
        With frmProgramPreferences
            intDesiredMaxIndex = Val(.cboMaximumFormulasToShow.List(.cboMaximumFormulasToShow.ListIndex)) - 1
        End With
    End If
    
    FormatMaximumFormulasToShowString = CStr(intDesiredMaxIndex) & "::" & CStr(Val(frmProgramPreferences.cboMaximumFormulasToShow.List(frmProgramPreferences.cboMaximumFormulasToShow.ListCount - 1)) - 1)
End Function

Public Function FormatForLocale(strWork As String) As String
    ' Searches through strWork, changing , to . (if decimal point is .) or . to , (if decimal point is ,)
    
    Dim strSearchChar As String, strReplaceChar As String, intDecPtLoc As Integer
    
    If glbDecimalSeparator = "," Then
        strSearchChar = ".": strReplaceChar = ","
    Else
        strSearchChar = ",": strReplaceChar = "."
    End If
    
    intDecPtLoc = InStr(strWork, strSearchChar)
    Do While intDecPtLoc > 0
        strWork = Left(strWork, intDecPtLoc - 1) & strReplaceChar & Mid(strWork, intDecPtLoc + 1)
        intDecPtLoc = InStr(strWork, strSearchChar)
    Loop
    
    FormatForLocale = strWork
    
End Function

Public Sub FormatLabel(lblThisLabel As Label, dblNewValue As Double)
    
    If frmProgramPreferences.optStdDevType(1).value = True Or dblNewValue < 0.0001 Or dblNewValue > 100000 Then
        ' Use scientific notation
        lblThisLabel.Caption = Trim(Format(dblNewValue, "0.0###E+00"))
    Else
        ' Display result in 0.000 notation rather than exponential
        lblThisLabel.Caption = Trim(Format(dblNewValue, "#0.0###"))
    End If

End Sub

Public Sub FormatTextBox(txtThisTextBox As TextBox, dblNewValue As Double)
    
    If dblNewValue = 0 Then
        txtThisTextBox.Text = "0"
    Else
        If frmProgramPreferences.optStdDevType(1).value = True Or dblNewValue < 0.0001 Or dblNewValue > 100000 Then
            ' Use scientific notation
            txtThisTextBox.Text = Format(dblNewValue, "0.0###E+00")
        Else
            ' Display result in 0.000 notation rather than exponential
            txtThisTextBox.Text = Format(dblNewValue, "#0.0#####")
        End If
    End If
    
End Sub

Public Sub GeneralErrorHandler(strSourceForm As String, Optional lngErrorNumber As Long = 0, Optional strErrorDescription As String = "")
    Dim strMessage As String
    
    If lngErrorNumber = 6 Then
        MsgBox LookupMessage(590), vbOKOnly + vbExclamation, LookupMessage(350)
    Else
        strMessage = LookupMessage(600) & ": " & strErrorDescription & vbCrLf & _
                   " (" & strSourceForm & " handler, #" & Trim(lngErrorNumber) & ")"
        strMessage = strMessage & vbCrLf & LookupMessage(605)
        MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)
    End If
End Sub

Private Sub GetColorAPIDlg(ByVal Ownerhwnd As Long, ThingColor As Long)
    Dim ChClr As ChooseColorType
    Dim CustColor(15) As Long
    Dim I As Integer
    Dim Res As Long
    
    For I = 0 To 15
        CustColor(I) = GetSysColor(I)
    Next I
    
    ChClr.lStructSize = Len(ChClr)
    ChClr.hwndOwner = Ownerhwnd
    ChClr.rgbResult = ThingColor
    ChClr.hInstance = 0
    ChClr.lpCustColors = VarPtr(CustColor(0))
    ChClr.flags = 0
    Res = ChooseColor(ChClr)
    If Res = 1 Then
        ThingColor = ChClr.rgbResult
    Else
        ' return 0 means Canceled, other Error
        ThingColor = ThingColor
    End If
End Sub

Public Sub HideFormShowMain(frmThisForm As VB.Form, Optional PossiblyExitProgram As Boolean = False)
    Dim boolStartupFormClosed As Boolean
    
    frmThisForm.Hide
    frmMain.Visible = True
    If frmMain.WindowState = vbMinimized Then frmMain.WindowState = vbNormal
    
    ' If the user has a specific form showing at startup, then exit entire program if user clicks
    ' the X on the given startup form
    If PossiblyExitProgram Then
        If frmProgramPreferences.cboStartupModule.ListIndex > 0 Then
            boolStartupFormClosed = False
            ' 0: Main
            ' 1: Formula Finder
            ' 2: Capillary Flow Calculator
            ' 3: Mole/Mass Converter
            ' 4: Peptide Sequence Fragmentation Modeller
            ' 5: Amino Acid Notation Converter
            ' 6: Isotopic Distribution Modeller
            Select Case frmProgramPreferences.cboStartupModule.ListIndex
            Case 1: If frmThisForm.Name = "frmFinder" Then boolStartupFormClosed = True
            Case 2: If frmThisForm.Name = "frmCapillaryCalcs" Then boolStartupFormClosed = True
            Case 3: If frmThisForm.Name = "frmMMConvert" Then boolStartupFormClosed = True
            Case 4: If frmThisForm.Name = "frmFragmentationModelling" Then boolStartupFormClosed = True
            Case 5: If frmThisForm.Name = "frmAminoAcidConverter" Then boolStartupFormClosed = True
            Case 6: If frmThisForm.Name = "frmIsotopicDistribution" Then boolStartupFormClosed = True
            Case Else
                ' A request was given to close a form that cannot be auto-shown at startup
                ' Do not exit entire program
            End Select
            
            If boolStartupFormClosed Then ExitProgram
        End If
    End If
End Sub

Public Sub HighlightOnFocus(ThisControl As TextBox)
    If cChkBox(frmProgramPreferences.chkHighlightTextFields) Then
        ThisControl.SelStart = 0
        ThisControl.SelLength = Len(ThisControl.Text)
    End If
    SetMostRecentTextBoxValue ThisControl.Text
End Sub

Public Sub InitializeGlobalVariables(Optional blnShowDebugPrompts As Boolean = False)
    Dim intMaxAllowableFormulaIndex As Integer
    
On Error GoTo InitializeGlobalVariablesErrorHandler

    ' Store program path
    gCurrentPath = App.Path
    
    App.HelpFile = App.Path
    If Right(App.HelpFile, 1) <> "\" Then App.HelpFile = App.HelpFile & "\"
    App.HelpFile = App.HelpFile & HELP_FILENAME

    gBlnStatusCaution = False           ' No caution warnings yet
    gBlnAccessFilesOnDrive = True       ' Allow program to read elements and abbreviations from drive
    gBlnWriteFilesOnDrive = True        ' Allow program to write files to drive
    gBlnLoadStatusOK = True             ' For now, assume the LoadElements and LoadAbbreviations procedures will be successfull
    gCurrentLanguage = "English"        ' Current language
    gCurrentLanguageFileName = DEFAULT_LANGUAGE_FILENAME
    gLastFileOpenSaveFolder = App.Path
    gMWAbbreviation = "MW"              ' Abbreviation for Molecular Weight; must be exactly two letters long
    
    If blnShowDebugPrompts Then MsgBox "Determine decimal point based on locale"
    glbDecimalSeparator = DetermineDecimalPoint()

    gMaxFormulaIndex = MAX_FORMULAS - 1          ' Max number of formulas (counting the first formula as 0)
    
    ' Need to make sure Screen is large enough to fit gMaxFormulaIndex formulas in Window
    ' Make smaller if not big enough
    intMaxAllowableFormulaIndex = DetermineMaxAllowableFormulaIndex()
    
    If gMaxFormulaIndex > intMaxAllowableFormulaIndex Then
        gMaxFormulaIndex = intMaxAllowableFormulaIndex
    End If

    If blnShowDebugPrompts Then MsgBox "First call to MwtWinDll"
    ' Make sure the maximum element index defined in MwtWinProcedures matches that in objMwtWin
    If objMwtWin.GetElementCount <> MAX_ELEMENT_INDEX Then
        Debug.Assert False
        If blnShowDebugPrompts Then MsgBox "Error: objMwtWin.GetElementCount returned " & Trim(objMwtWin.GetElementCount) & " rather than " & Trim(MAX_ELEMENT_INDEX)
    End If
    
    objMwtWin.BracketsTreatedAsParentheses = False
    objMwtWin.ShowErrorDialogs = blnShowDebugPrompts
    
    Exit Sub
    
InitializeGlobalVariablesErrorHandler:
    If Err.Number = -2147024770 Or Err.Number = 429 Then
        MsgBox "Unable to connect to MwtWinDll.Dll -- you probably need to reinstall the application to properly register the file.  The MwtwinDll.Dll file is required for program operation, and the program will therefore be terminated.", _
                vbOKOnly + vbExclamation, "Dll Not found"
        End
    Else
        GeneralErrorHandler "frmFinder|cmdCalculate", Err.Number, Err.Description
    End If

End Sub

Public Function LookupToolTipLanguageCaption(strCaptionID As String, strDefaultTooltip As String) As String
    ' ToolTip captions are equal to the CaptionID value plus 1
    If IsNumeric(strCaptionID) And cChkBox(frmProgramPreferences.chkShowToolTips) Then
        LookupToolTipLanguageCaption = LookupLanguageCaption(CStr(CLng(strCaptionID) + 1), strDefaultTooltip)
    Else
        LookupToolTipLanguageCaption = strDefaultTooltip
    End If
End Function

Public Function LookupLanguageCaption(ByVal strCaptionID As String, ByVal strOldCaption As String, Optional ByVal boolMenuItem As Boolean = False, Optional ByVal strMenuControlName As String) As String
    Dim lngCaptionID As Long, strNewCaption As String, strMenuCaption As String
    
    If IsNumeric(strCaptionID) Then
        lngCaptionID = CLng(strCaptionID)
        
        ' Use lngCaptionID to search grdLanguageStrings for caption
        strNewCaption = LookupLanguageCaptionInGrid(lngCaptionID, frmStrings.grdLanguageStrings, frmStrings.grdLanguageStringsCrossRef)
        
        If Len(strNewCaption) = 0 Then
            ' Caption not found
            ' If caption is for a Menu then load English version of caption from frmStrings
            ' Note that Menu id values on frmMain are 1000 to 3599
            '  while menu ID's on other forms are larger numbers; thus the check for >= 1000 and <= 3599 will only work for frmMain)
            If (lngCaptionID >= 1000 And lngCaptionID <= 3599 And IsNumeric(strOldCaption)) Or boolMenuItem Then
                strMenuCaption = LookupMenuCaption(lngCaptionID, strMenuControlName)
                If strMenuCaption <> "" Then
                    strNewCaption = strMenuCaption
                End If
            Else
                strNewCaption = strOldCaption
            End If
        End If
    Else
        If Len(strCaptionID) > 0 And strCaptionID <> "-" And strCaptionID <> "RightClickMenu" Then
            ' This code should not be reached
            Debug.Assert False
        End If
        strNewCaption = strOldCaption
    End If
    
    ' Caption still not found, return strOldCaption
    LookupLanguageCaption = strNewCaption

End Function

Private Function LookupLanguageCaptionInGrid(lngCaptionID As Long, grdThisLanguageGrid As MSFlexGrid, grdThisLanguageGridCrossRef As MSFlexGrid) As String
    Dim lngIndex As Long, lngStartRow As Long, lngLookupRow As Long
    Dim lngNewCaption As String
    
    If lngCaptionID > 0 Then
        
        ' The CrossRef grids are lookup tables that contain info on what row in
        '  grdThisLanguageGrid to start searching for a given lngCaptionID
        ' For example, ID's between 1000 and 1999 would have lngStartRow =0
        '  2000 to 2999 might have lngStartRow = 20
        '  3000 to 3999 might have lngStartRow = 50
        '   etc.
        ' This way, when searching for a large lngCaptionID (like 20000), we don't start
        '  searching at row 0 of grdThisLanguageGrid.
        With grdThisLanguageGridCrossRef
            ' The lookup row in the CrossRef grid is the integer portion of lngCaptionID/1000
            ' Must use the Int() function to assure lngCaptionID / 1000 rounds down
            lngLookupRow = Int(lngCaptionID / 1000) - 1
            ' Make sure lngLookupRow isn't invalid
            If lngLookupRow < .Rows Then
                lngStartRow = grdThisLanguageGridCrossRef.TextMatrix(lngLookupRow, 1)
            Else
                ' Most likely the CrossRef grid hasn't been filled yet
                lngStartRow = 0
            End If
        End With
        
        With grdThisLanguageGrid
            For lngIndex = lngStartRow To .Rows - 1
                If IsNumeric(.TextMatrix(lngIndex, 0)) Then
                    If Val(.TextMatrix(lngIndex, 0)) = lngCaptionID Then
                        lngNewCaption = .TextMatrix(lngIndex, 1)
                        Exit For
                    End If
                End If
            Next lngIndex
        End With
    Else
        lngNewCaption = ""
    End If
        
    LookupLanguageCaptionInGrid = lngNewCaption

End Function

Private Function LookupMenuCaption(lngMenuCaptionID As Long, strMenuControlName As String) As String
    Dim strMenuCaption As String, lngIndex As Long
    
    With frmStrings.grdMenuInfo
        For lngIndex = 0 To .Rows - 1
            If CLng(.TextMatrix(lngIndex, 0)) = lngMenuCaptionID Or _
               LCase(.TextMatrix(lngIndex, 1)) = LCase(strMenuControlName) Then
                strMenuCaption = .TextMatrix(lngIndex, 2)
                Exit For
            End If
        Next lngIndex
    End With
        
    LookupMenuCaption = strMenuCaption

End Function

Public Function LookupMessage(lngMessageID As Long, Optional strAppendText As String = "") As String
    Dim strMessage As String
    
On Error GoTo LookupMessageErrorHandler

    With objMwtWin
        strMessage = .GetMessageStatement(lngMessageID, strAppendText)
        
        ' lngMessageID's 1 and 18 may need to have an addendum added
        If lngMessageID = 1 Then
            If frmProgramPreferences.optConvertType(1).value = True Then
                strMessage = strMessage & " (" & .GetMessageStatement(680) & ")"
            End If
        ElseIf lngMessageID = 18 Then
            If Not cChkBox(frmProgramPreferences.chkBracketsAsParentheses) Then
                strMessage = strMessage & " (" & .GetMessageStatement(685) & ")"
            Else
                strMessage = strMessage & " (" & .GetMessageStatement(690) & ")"
            End If
        End If
    End With
    
    If Len(strMessage) = 0 Then
        strMessage = "Unknown error"
        Debug.Assert False
    End If
    
    LookupMessage = strMessage
    Exit Function

LookupMessageErrorHandler:
    GeneralErrorHandler "MwtWinProcedures|LookupMessage", Err.Number, Err.Description
    LookupMessage = "Unknown message"
    
End Function

Public Sub Main()
    
    ' Only need this if we don't use the New keyword when declaring objMwtWin
    ' Set objMwtWin = New MolecularWeightCalculator
    
    frmMain.Show
    
    If Not gNonSaveExitApp Then
        If Not gBlnLoadStatusOK Then
            frmIntro.cmdOK.SetFocus
        End If
    Else
        If gCommandLineInstructionsDisplayed Then
            frmMain.Hide
        Else
            Unload frmMain
            
            ' This shouldn't ever get reached, since the program should end when the
            '   Unload event on frmMain finishes
            Debug.Assert False
            Set objMwtWin = Nothing
        End If
    End If
End Sub

Public Sub MemoryLoadAllStringConstants()
    ' Load Menu captions and codes into frmStrings
    MemoryLoadMenuCaptions

End Sub

Private Sub MemoryLoadMenuCaptions()
    With frmStrings
        With .grdMenuInfo
            .Clear
            .Rows = 0
            .Cols = 3
            .ColWidth(0) = 650
            .ColWidth(1) = 2000
            .ColWidth(2) = 3000
        End With
        
        ' Note that we must add the items in order numerically
        FlexGridAddItems .grdMenuInfo, "1000", "mnuFile", "&File"
        FlexGridAddItems .grdMenuInfo, "1010", "mnuEditElements", "Edit &Elements Table..."
        FlexGridAddItems .grdMenuInfo, "1020", "mnuEditAbbrev", "Edit &Abbreviations..."
        FlexGridAddItems .grdMenuInfo, "1030", "mnuCalculateFile", "&Calculate weights from text file"
        FlexGridAddItems .grdMenuInfo, "1040", "mnuPrint", "&Print Results..."
        FlexGridAddItems .grdMenuInfo, "1050", "mnuExit", "E&xit..."
        FlexGridAddItems .grdMenuInfo, "1500", "mnuEdit", "&Edit"
        FlexGridAddItems .grdMenuInfo, "1500", "mnuRightClick", "Right Click Menu"      ' This menu is hidden, so its caption doesn't matter; I'll use the Edit menu's menu code here
        FlexGridAddItems .grdMenuInfo, "1505", "mnuRightClickUndo", "&Undo"
        FlexGridAddItems .grdMenuInfo, "1510", "mnuCut", "Cu&t"
        FlexGridAddItems .grdMenuInfo, "1510", "mnuRightClickCut", "Cu&t"           ' Duplicate of 1510 above
        FlexGridAddItems .grdMenuInfo, "1520", "mnuCopy", "&Copy"
        FlexGridAddItems .grdMenuInfo, "1520", "mnuRightClickCopy", "&Copy"         ' Duplicate of 1520 above
        FlexGridAddItems .grdMenuInfo, "1530", "mnuPaste", "&Paste"
        FlexGridAddItems .grdMenuInfo, "1530", "mnuRightClickPaste", "&Paste"       ' Duplicate of 1530 above
        FlexGridAddItems .grdMenuInfo, "1540", "mnuDelete", "&Delete"
        FlexGridAddItems .grdMenuInfo, "1540", "mnuRightClickDelete", "&Delete"     ' Duplicate of 1540 above
        FlexGridAddItems .grdMenuInfo, "1545", "mnuRightClickSelectAll", "Select &All"
        FlexGridAddItems .grdMenuInfo, "1550", "mnuCopyRTF", "Copy Current Formula as &RTF"
        FlexGridAddItems .grdMenuInfo, "1560", "mnuCopyMWT", "Copy Current &Molecular Weight"
        FlexGridAddItems .grdMenuInfo, "1570", "mnuCopyPC", "Copy P&ercent Composition Data"
        FlexGridAddItems .grdMenuInfo, "1580", "mnuCopyCurrent", "Duplicate Current &Formula"
        FlexGridAddItems .grdMenuInfo, "1590", "mnuEraseAll", "Erase &All Formulas..."
        FlexGridAddItems .grdMenuInfo, "1600", "mnuEraseCurrent", "Erase Current Formula..."
        FlexGridAddItems .grdMenuInfo, "1610", "mnuExpandAbbrev", "E&xpand Abbreviations..."
        FlexGridAddItems .grdMenuInfo, "1620", "mnuEmpirical", "Convert to Empirical F&ormula..."
        FlexGridAddItems .grdMenuInfo, "2000", "mnuView", "&View"
        FlexGridAddItems .grdMenuInfo, "2010", "mnuViewType", "&Multi View"
        FlexGridAddItems .grdMenuInfo, "2020", "mnuViewType", "&Single View"
        FlexGridAddItems .grdMenuInfo, "2030", "mnuPercentSolver", "&Percent Solver"
        FlexGridAddItems .grdMenuInfo, "2040", "mnuPercentType", "O&ff"
        FlexGridAddItems .grdMenuInfo, "2050", "mnuPercentType", "&On"
        FlexGridAddItems .grdMenuInfo, "2500", "mnuTools", "&Tools"
        FlexGridAddItems .grdMenuInfo, "2510", "mnuMMConvert", "&Mole/Mass Converter..."
        FlexGridAddItems .grdMenuInfo, "2520", "mnuFinder", "&Formula Finder..."
        FlexGridAddItems .grdMenuInfo, "2530", "mnuAminoAcidNotationConverter", "&Amino Acid Notation Converter..."
        
        FlexGridAddItems .grdMenuInfo, "2533", "mnuPeptideSequenceFragmentation", "&Peptide Sequence Fragmentation Modelling..."
        FlexGridAddItems .grdMenuInfo, "2536", "mnuIsotopicDistribution", "&Isotopic Distribution Modelling..."
        FlexGridAddItems .grdMenuInfo, "2538", "mnuDisplayIsotopicDistribution", "Show Isotopic &Distribution for Current Formula"
        
        FlexGridAddItems .grdMenuInfo, "2540", "mnuCalculator", "Math &Calculator..."
        FlexGridAddItems .grdMenuInfo, "2550", "mnuCapillaryFlow", "Capillar&y Flow Calculator"
        
        FlexGridAddItems .grdMenuInfo, "3000", "mnuOptions", "&Options"
        FlexGridAddItems .grdMenuInfo, "3010", "mnuChooseLanguage", "Choose &Language..."
        FlexGridAddItems .grdMenuInfo, "3020", "mnuProgramOptions", "Change Program &Preferences..."
        FlexGridAddItems .grdMenuInfo, "3030", "mnuChangeFont", "Change &Formula Font..."
        FlexGridAddItems .grdMenuInfo, "3040", "mnuStayOnTop", "Stay on &Top"
        FlexGridAddItems .grdMenuInfo, "3050", "mnuDefaultsOptions", "&Save and Restore Default Values"
        FlexGridAddItems .grdMenuInfo, "3060", "mnuRestoreValues", "&Restore Default Values and Formulas..."
        FlexGridAddItems .grdMenuInfo, "3070", "mnuSaveValues", "Save &Values and Formulas Now!"
        FlexGridAddItems .grdMenuInfo, "3500", "mnuHelp", "&Help"
        FlexGridAddItems .grdMenuInfo, "3510", "mnuOverview", "&Program Overview"
        FlexGridAddItems .grdMenuInfo, "3530", "mnuShowTips", "&Show Tool Tips"
        FlexGridAddItems .grdMenuInfo, "3540", "mnuAbout", "&About MWT..."
            
        ' frmCapillaryFlow
        FlexGridAddItems .grdMenuInfo, "7900", "mnuLoadCapValues", "&Load Values"
        FlexGridAddItems .grdMenuInfo, "7910", "mnuSaveCapValues", "&Save Values"
        FlexGridAddItems .grdMenuInfo, "7950", "mnuCapillaryFlowHelp", "&Capillary Flow Calculations"

        ' frmFragmentationModelling
        FlexGridAddItems .grdMenuInfo, "12800", "mnuLoadSequenceInfo", "&Load Sequence Info"
        FlexGridAddItems .grdMenuInfo, "12810", "mnuSaveSequenceInfo", "&Save Sequence Info"
        FlexGridAddItems .grdMenuInfo, "12820", "mnuLoadIonList", "Load List of &Ions or .Dta file to Match"
        FlexGridAddItems .grdMenuInfo, "12830", "mnuClose", "&Close"
        FlexGridAddItems .grdMenuInfo, "12840", "mnuCopyPredictedIons", "&Copy Predicted Ions"
        FlexGridAddItems .grdMenuInfo, "12850", "mnuCopyPredictedIonsAsRTF", "Copy Predicted Ions as &RTF"
        FlexGridAddItems .grdMenuInfo, "12855", "mnuCopyPredictedIonsAsHtml", "Copy Predicted Ions as Html"
        FlexGridAddItems .grdMenuInfo, "12860", "mnuPasteIonList", "&Paste List of Ions to Match"
        FlexGridAddItems .grdMenuInfo, "12865", "mnuCopySequenceMW", "Copy Sequence Molecular &Weight"
        FlexGridAddItems .grdMenuInfo, "12870", "mnuClearMatchIonList", "Clear Match Ion &List"
        FlexGridAddItems .grdMenuInfo, "12880", "mnuViewMatchIonList", "List of &Ions to Match"
        FlexGridAddItems .grdMenuInfo, "12900", "mnuShowMassSpectrum", "&Mass Spectrum"
        FlexGridAddItems .grdMenuInfo, "12910", "mnuUpdateSpectrum", "&Update Spectrum on Change"
        FlexGridAddItems .grdMenuInfo, "12915", "mnuViewDtaTxtBrowser", "&Dta.Txt File Browser"
        FlexGridAddItems .grdMenuInfo, "12920", "mnuIonMatchListOptions", "Ion Match List &Options"
        FlexGridAddItems .grdMenuInfo, "12925", "mnuEditModificationSymbols", "Edit Residue &Modification Symbols"
        FlexGridAddItems .grdMenuInfo, "12930", "mnuAutoAlign", "&Automatically Align Ions to Match"
        FlexGridAddItems .grdMenuInfo, "12940", "mnuFragmentationModellingHelp", "&Fragmentation Modelling"
        
        FlexGridAddItems .grdMenuInfo, "12950", "mnuIonMatchListRightClick", "IonMatchListRightClickMenu"      ' This menu is hidden, so its caption doesn't matter; I'll use 12950 here
        FlexGridAddItems .grdMenuInfo, "12950", "mnuIonMatchListRightClickCopy", "&Copy Selected Ions"
        FlexGridAddItems .grdMenuInfo, "12860", "mniIonMatchListRightClickPaste", "&Paste List of Ions to Match"
        FlexGridAddItems .grdMenuInfo, "12960", "mnuIonMatchListRightClickDeleteAll", "&Delete All (Clear list)"
        FlexGridAddItems .grdMenuInfo, "1545", "mnuIonMatchListRightClickSelectAll", "Select &All"
                
        ' frmMsPlot
        FlexGridAddItems .grdMenuInfo, "13100", "mnuExportData", "&Export Data"
        FlexGridAddItems .grdMenuInfo, "13150", "mnuPlotType", "&Plot Type"
        FlexGridAddItems .grdMenuInfo, "13160", "mnuPlotTypeSticksToZero", "&Sticks To Zero"
        FlexGridAddItems .grdMenuInfo, "13170", "mnuPlotTypeGaussianPeaks", "&Gaussian Peaks"
        FlexGridAddItems .grdMenuInfo, "13180", "mnuSetResolution", "Set Effective &Resolution"
        FlexGridAddItems .grdMenuInfo, "13190", "mnuGridLinesXAxis", "X Axis Gridlines"
        FlexGridAddItems .grdMenuInfo, "13200", "mnuGridLinesYAxis", "Y Axis Gridlines"
        FlexGridAddItems .grdMenuInfo, "13210", "mnuTicks", "&Ticks to label (approx.)"
        FlexGridAddItems .grdMenuInfo, "13220", "mnuTicksXAxis", "&X Axis"
        FlexGridAddItems .grdMenuInfo, "13230", "mnuTicksYAxis", "&Y Axis"
        FlexGridAddItems .grdMenuInfo, "13235", "mnuPlotQuality", "Plot &Quality (affects speed)"
        FlexGridAddItems .grdMenuInfo, "13240", "mnuGaussianQuality", "&Gaussian Representation Quality"
        FlexGridAddItems .grdMenuInfo, "13245", "mnuApproximationFactor", "&Approximation Factor"
        FlexGridAddItems .grdMenuInfo, "13250", "mnuSetRangeX", "Set &X Range"
        FlexGridAddItems .grdMenuInfo, "13260", "mnuSetRangeY", "Set &Y Range"
        FlexGridAddItems .grdMenuInfo, "13270", "mnuAutoScaleYAxis", "&Autoscale Y Axis"
        FlexGridAddItems .grdMenuInfo, "13280", "mnuFixMinimumYAtZero", "&Fix mimimum Y at zero"
        FlexGridAddItems .grdMenuInfo, "13290", "mnuZoomOutToPrevious", "&Zoom Out to Previous"
        FlexGridAddItems .grdMenuInfo, "13300", "mnuZoomOutFullScale", "Zoom Out to Show All"
        FlexGridAddItems .grdMenuInfo, "13310", "mnuCursorMode", "&Cursor Mode"
        FlexGridAddItems .grdMenuInfo, "13320", "mnuCursorModeZoom", "&Zoom"
        FlexGridAddItems .grdMenuInfo, "13330", "mnuCursorModeMove", "&Move"
        FlexGridAddItems .grdMenuInfo, "13340", "mnuShowCurrentPosition", "&Show Current Position"
        FlexGridAddItems .grdMenuInfo, "13342", "mnuShowLegend", "Show &Legend"
        FlexGridAddItems .grdMenuInfo, "13345", "mnuResetToDefaults", "Reset to &Default Options"
        FlexGridAddItems .grdMenuInfo, "13350", "mnuZoomOptions", "&Zoom Box"
        FlexGridAddItems .grdMenuInfo, "13360", "mnuZoomIn", "Zoom &In"
        FlexGridAddItems .grdMenuInfo, "13370", "mnuZoomInHorizontal", "Zoom In Horizontal"
        FlexGridAddItems .grdMenuInfo, "13380", "mnuZoomInVertical", "Zoom In Vertical"
        FlexGridAddItems .grdMenuInfo, "13390", "mnuZoomOut", "Zoom &Out"
        FlexGridAddItems .grdMenuInfo, "13400", "mnuZoomOutHorizontal", "Zoom Out Horizontal"
        FlexGridAddItems .grdMenuInfo, "13410", "mnuZoomOutVertical", "Zoom Out Vertical"
    
    End With
End Sub

Public Sub ModSymbolKeyPressHandler(txtThisTextBox As VB.TextBox, ByRef KeyAscii As Integer)

    ' Checks KeyAscii to see if it's valid
    ' If it isn't, it is set to 0
    
    ' Valid Mod Symbols are ! # $ % & ' * + ? ^ _ ` ~
    If Not objMwtWin.IsModSymbol(Chr(KeyAscii)) Then
        
        Select Case KeyAscii
        Case 1          ' Ctrl+A -- select entire textbox
            txtThisTextBox.SelStart = 0
            txtThisTextBox.SelLength = Len(txtThisTextBox.Text)
            KeyAscii = 0
        Case 24, 3, 22  ' Cut, Copy, Paste is allowed
        Case 26
            ' Ctrl+Z = Undo
            KeyAscii = 0
            txtThisTextBox.Text = GetMostRecentTextBoxValue()
        Case 8          ' Backspace key is allowed
        Case 48 To 57   ' Numbers are not allowed       ' ToDo: Possibly allow them for side-chains
            KeyAscii = 0
        Case 32         ' Spaces are not allowed
            KeyAscii = 0
        Case 40 To 41   ' Parentheses are not allowed       ' ToDo: Possibly allow them for side-chains
            KeyAscii = 0
        Case 43:        ' Plus sign is allowed
        Case 45:        ' Negative sign is not allowed
            KeyAscii = 0
        Case 44, 46:    ' Decimal point (. or ,) is not allowed
            KeyAscii = 0
        Case 65 To 90, 97 To 122    ' Characters are not allowed
            KeyAscii = 0
        Case 95:        ' Underscore is allowed
        Case Else
            KeyAscii = 0
        End Select
    End If
    
End Sub

Public Sub PasteRoutine(frmCallingForm As VB.Form, blnCalledFromMenu As Boolean)
    ' If blnCalledFromMenu = True, then need to manually paste
    ' Otherwise, CTRL+V was pressed and the RTF box already pasted the text for me
    
    Dim strPasteText As String
    
On Error GoTo PasteRoutineErrorHandler

    If blnCalledFromMenu Then
        ' Place text from Clipboard into active control.
        If Clipboard.GetFormat(vbCFRTF) Then
            ' rtf on clipboard
            strPasteText = GetClipboardTextSmart(vbCFRTF)
            If TypeOf frmCallingForm.ActiveControl Is RichTextBox Then
                frmCallingForm.ActiveControl.SelRTF = strPasteText
            End If
        Else
            strPasteText = GetClipboardTextSmart()
            frmCallingForm.ActiveControl.SelText = strPasteText
        End If
    End If
    
    Exit Sub

PasteRoutineErrorHandler:
    GeneralErrorHandler "MwtWinProcedures|PasteRoutine", Err.Number, Err.Description

End Sub

Public Sub PopulateComboBox(cboThisComboBox As ComboBox, boolLookupLanguageItems As Boolean, Optional strDefaultItemsToPopulate As String = "", Optional intListIndexToSet As Integer = -1)
    ' Note: This sub assumes the maximum number of items assigned to a combo box is 30
    
    Dim intNewListIndex As Integer, intOldListCount As Integer
    Dim strPotentialItemsToPopulate As String
    Dim strItemsArray(30) As String, intItemsCount As Integer, intIndex As Integer
    Dim lngComboBoxID As Long, intSepLoc As Integer
    
    With cboThisComboBox
        If intListIndexToSet >= 0 Then
            intNewListIndex = intListIndexToSet
        Else
            intNewListIndex = .ListIndex
        End If
        intOldListCount = .ListCount
        
        If boolLookupLanguageItems Then
            ' See if the combo box has an ID and see if it's in frmstrings.grdLanguageStrings
            If IsNumeric(.Tag) Then
                lngComboBoxID = CLng(.Tag)
                
                ' Use lngCaptionID to search grdLanguageStrings for caption
                strPotentialItemsToPopulate = ""
                strPotentialItemsToPopulate = LookupLanguageCaptionInGrid(lngComboBoxID, frmStrings.grdLanguageStrings, frmStrings.grdLanguageStringsCrossRef)
            End If
        End If
        
        If Len(strPotentialItemsToPopulate) = 0 Then
            strPotentialItemsToPopulate = strDefaultItemsToPopulate
        End If
        
        If Len(strPotentialItemsToPopulate) > 0 Then
            intSepLoc = InStr(strPotentialItemsToPopulate, "|")
            If intSepLoc Then
                Do While intSepLoc > 0
                    intItemsCount = intItemsCount + 1
                    strItemsArray(intItemsCount) = Left(strPotentialItemsToPopulate, intSepLoc - 1)
                    strPotentialItemsToPopulate = Mid(strPotentialItemsToPopulate, intSepLoc + 1)
                    intSepLoc = InStr(strPotentialItemsToPopulate, "|")
                Loop
                If Len(strPotentialItemsToPopulate) > 0 Then
                    intItemsCount = intItemsCount + 1
                    strItemsArray(intItemsCount) = strPotentialItemsToPopulate
                End If
                
                ' If there were more items in the old list than were found in strPotentialItemsToPopulate,
                '  then grab items from the old list and add to strItemsArray()
                Do While intOldListCount > intItemsCount
                    intItemsCount = intItemsCount + 1
                    strItemsArray(intItemsCount) = .List(intItemsCount - 1)
                Loop
                
                ' Finally, update the list
                .Clear
                For intIndex = 1 To intItemsCount
                    .AddItem (strItemsArray(intIndex))
                Next intIndex
            Else
                ' Item separator (|) not found
                ' If intOldListCount = 0 or intOldListCount = 1 then update list
                If intOldListCount <= 1 Then
                    .Clear
                    .AddItem strPotentialItemsToPopulate
                End If
            End If
            If intNewListIndex < .ListCount Then
                .ListIndex = intNewListIndex
            End If
        Else
            ' Nothing to populate list with
        End If
    End With
End Sub

Public Sub PositionWeightSourceframeControls(frmThisForm As VB.Form)
        
    With frmThisForm
        .optWeightSource(0).Top = 0
        .optWeightSource(0).Left = 0
        .optWeightSource(1).Top = 260
        .optWeightSource(1).Left = .optWeightSource(0).Left
        
        .lblFormula.Top = 720
        .lblFormula.Left = 120
        With .rtfCurrentFormula
            .Top = 600
            .Left = 1800
            .Width = 3850
            .Height = 500
        End With
        .lblMWT.Top = 1200
        .lblMWT.Left = .lblFormula.Left
        
        .lblCustomMass.Top = .lblFormula.Top
        .lblCustomMass.Left = .lblFormula.Left
        .txtCustomMass.Top = .lblCustomMass.Top - 40
        .txtCustomMass.Left = .rtfCurrentFormula.Left
        .lblCustomMassUnits.Top = .lblCustomMass.Top
        .lblCustomMassUnits.Left = .txtCustomMass.Left + .txtCustomMass.Width + 100
    End With
End Sub

' Purpose: Called by the Form_Activate function of each form to hide frmMain if .chkHideInactiveForms = vbChecked
Public Sub PossiblyHideMainWindow()
    If cChkBox(frmProgramPreferences.chkHideInactiveForms) Then
        frmMain.Visible = False
    End If
End Sub

Public Sub RecalculateAllFormulas(Optional blnSetFocusToCurrentFormula As Boolean = True)
    Dim intIndex As Integer
    
    With frmMain
        ' ReCalculate
        For intIndex = 0 To frmMain.GetTopFormulaIndex
            .rtfFormula(intIndex).Tag = FORMULA_CHANGED
        Next intIndex
        .Calculate True, True, True, 0, False, False, False, 1, blnSetFocusToCurrentFormula
        .LabelStatus
    End With
    
End Sub

Public Sub QueryUnloadFormHandler(frmThisForm As VB.Form, ByRef boolCancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        boolCancel = True
        HideFormShowMain frmThisForm, True
    End If
End Sub

Public Sub RaiseLabelStatusError(ThisErrorID As Long)
    frmMain.lblStatus.ForeColor = QBColor(COLOR_ERR)
    frmMain.lblStatus.Caption = LookupMessage(ThisErrorID)
    gBlnErrorPresent = True
End Sub

Public Sub RemoveHeightAdjustChar(ByRef strText As String)
    If InStr(strText, RTF_HEIGHT_ADJUSTCHAR) > 0 Then
        ' Remove the tilde sign
        strText = Left(strText, InStr(strText, RTF_HEIGHT_ADJUSTCHAR) - 1)
    End If

End Sub

Public Function RemoveRTFHeightAdjust(rtfText As String, Optional strReplaceText As String = "") As String
        
    ' Search for tilde signs and get rid of them, replacing "\cf1 ~" with strReplaceText
    rtfText = Replace(rtfText, "\cf1 " & RTF_HEIGHT_ADJUSTCHAR, strReplaceText)
    
    If InStr(rtfText, RTF_HEIGHT_ADJUSTCHAR) > 0 Then
        rtfText = Replace(rtfText, "\cf2 " & RTF_HEIGHT_ADJUSTCHAR, strReplaceText)
    End If

    RemoveRTFHeightAdjust = rtfText

End Function

Public Sub ResetMenuCaptions(boolResetToText As Boolean)
    ' If boolResetToText = True then captions are reset to default text
    ' Otherwise, they are reset to default numeric values
    
    ' Reset menus on frmMain
    ResetMenuCaptionsWork frmMain, boolResetToText
    
    ' Reset on other forms
    ResetMenuCaptionsWork frmFinder, boolResetToText
    ResetMenuCaptionsWork frmCalculator, boolResetToText
    ResetMenuCaptionsWork frmCapillaryCalcs, boolResetToText
    ResetMenuCaptionsWork frmFragmentationModelling, boolResetToText
'''    ResetMenuCaptionsWork frmMsPlot, boolResetToText
    
    If boolResetToText = True Then
        ' Add shortcut keys to menu captions
        AppendShortcutKeysToMenuCaptions
    End If
    
End Sub

Private Sub ResetMenuCaptionsWork(frmThisForm As VB.Form, boolResetToText As Boolean)
    ' Reset the menu captions to either and ID number (so that an international text value can be loaded)
    '  or to the default English text
    
    Dim ctlThisControl As Control
    Dim intViewTypeChecked As Integer, intPercentTypeChecked As Integer
    Dim lngIndex As Long, lngOffSet As Long
    Dim boolReset As Boolean
    
    ' Note: intViewTypeChecked and intPercentTypeChecked apply to frmMain only
    intViewTypeChecked = 0
    intPercentTypeChecked = 0
    For Each ctlThisControl In frmThisForm.Controls
        If TypeName(ctlThisControl) = "Menu" Then
            With frmStrings.grdMenuInfo
                Select Case ctlThisControl.Name
                Case "mnuViewType"
                    lngOffSet = intViewTypeChecked * 10
                    intViewTypeChecked = intViewTypeChecked + 1
                Case "mnuPercentType"
                    lngOffSet = intPercentTypeChecked * 10
                    intPercentTypeChecked = intPercentTypeChecked + 1
                Case Else
                    lngOffSet = 0
                End Select
                
                boolReset = False
                For lngIndex = 0 To .Rows - 1
                    If .TextMatrix(lngIndex, 1) = ctlThisControl.Name Then
                        boolReset = True
                        If boolResetToText Then
                            ' Reset to text value
                            ctlThisControl.Caption = .TextMatrix(lngIndex + lngOffSet / 10, 2)
                        Else
                            ' Reset to numeric ID
                            ctlThisControl.Caption = CStr(CLng(.TextMatrix(lngIndex, 0)) + lngOffSet)
                        End If
                        Exit For
                    End If
                Next lngIndex
                
                ' Unless the menu caption is a dash (which means a separator menu), make sure
                '  that boolReset = True
                ' If it isn't, then there probably is not an entry for ctlThisControl.Name in grdMenuInfo
                ' Need to edit MemoryLoadMenuCaptions
                If ctlThisControl.Caption <> "-" Then
                    Debug.Assert boolReset
                End If
            End With
        End If
    Next

End Sub

Public Sub RTFBoxKeyPressHandler(frmCallingForm As VB.Form, ThisRichTextBox As RichTextBox, ByRef KeyAscii As Integer, Optional blnDisplayErrorsInStatusLabel As Boolean = True, Optional blnAllowAminoAcidModSymbols As Boolean = False)
    Dim blnCharOK As Boolean
    
    If blnAllowAminoAcidModSymbols Then
        If objMwtWin.IsModSymbol(Chr(KeyAscii)) Then blnCharOK = True
    End If
    
    If Not blnCharOK Then
        Select Case KeyAscii
        Case 44, 46
            Select Case glbDecimalSeparator
            Case ","
                If KeyAscii = 46 Then
                    KeyAscii = 0
                    If blnDisplayErrorsInStatusLabel Then RaiseLabelStatusError 94
                End If
            Case Else   ' includes "."
                If KeyAscii = 44 Then
                    KeyAscii = 0
                    If blnDisplayErrorsInStatusLabel Then RaiseLabelStatusError 93
                End If
            End Select
        Case 40 To 41, 43, 45, 48 To 57, 62, 65 To 90, 91, 93, 94, 95, 97 To 122, 123, 125
            ' Valid Characters are:
            ' ( and ) are 40 and 41   - is 45   + is 43
            ' { and } are 123 and 125
            ' Numbers are 48 to 57    . is 46
            ' Lowercase letters are 97 to 122
            ' Uppercase letters are 65 to 90
            ' > is 62
            ' _ is 95
        Case 8
            ' Backspace; it is valid
        Case 9
            ' Tab; it is valid
        Case 32
            ' Space
            If blnDisplayErrorsInStatusLabel Then RaiseLabelStatusError 610
            KeyAscii = 0
        Case 3
            ' User pressed CTRL-C (copy)
            CopyRoutine frmCallingForm, False
        Case 22
            ' User pressed CTRL-V (paste)
            PasteRoutine frmCallingForm, False
        Case 24
            ' User pressed CTRL-X (cut)
            CutRoutine frmCallingForm, False
        Case 1
            ' Ctrl+A - rtfTextBox automatically highlights entire box
        Case 26
            ' Ctrl+Z - Undo
            ThisRichTextBox.Text = GetMostRecentTextBoxValue()
        Case 0
            ' This is set by Form_KeyDown, we should just ignore the entry
        Case vbKeyEscape
            ' Ignore the key if it got this far
            KeyAscii = 0
        Case Else
            ' All other ASCII characters are invalid
            If blnDisplayErrorsInStatusLabel Then
                frmMain.lblStatus.ForeColor = QBColor(COLOR_ERR)
                If KeyAscii >= 1 And KeyAscii <= 26 Then
                    frmMain.lblStatus.Caption = LookupMessage(620) & " Ctrl+" & Chr$(KeyAscii + 64)   ' & "#" & Str(keyascii)
                Else
                    frmMain.lblStatus.Caption = LookupMessage(620) & " " & Chr$(KeyAscii)               ' & "#" & Str(keyascii)
                End If
            End If
            KeyAscii = 0
        End Select
    End If
    
End Sub

Public Sub SelectCustomColor(frmCallingForm As VB.Form, lblThisLabel As Label)
    Dim lngTemporaryColor As Long
    
    lngTemporaryColor = lblThisLabel.BackColor
    Call GetColorAPIDlg(frmCallingForm.hwnd, lngTemporaryColor)
    If lngTemporaryColor >= 0 Then
        lblThisLabel.BackColor = lngTemporaryColor
    End If
    
End Sub

Public Sub SetAllTooltips(Optional blnShowDebugPrompts As Boolean = False)
    Dim intIndex As Integer
    
On Error GoTo SetAllTooltipsErrorHandler
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmAminoAcidConverter"
    With frmAminoAcidConverter
        .txt1LetterSequence.ToolTipText = LookupToolTipLanguageCaption(6030, "Enter sequence using 1 letter abbreviations here")
        .txt3LetterSequence.ToolTipText = LookupToolTipLanguageCaption(6040, "Enter sequence using 3 letter abbreviations here")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmCalculator"
    With frmCalculator
        .rtfExpression.ToolTipText = LookupToolTipLanguageCaption(6510, "Enter a mathematical expression to evaluate here")
        .cmdCalculate.ToolTipText = LookupToolTipLanguageCaption(6520, "Evaluates the current expression")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmDiff"
    frmDiff.cmdCopy.ToolTipText = LookupToolTipLanguageCaption(8610, "Copies the results to the clipboard")
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmEditElem"
    With frmEditElem
        .cmdReset.ToolTipText = LookupToolTipLanguageCaption(9210, "Resets elemental weights to their average weights")
        .cmdAverageMass.ToolTipText = LookupToolTipLanguageCaption(9230, "Sets all elemental weights to their average weights found in nature")
        .cmdIsotopicMass.ToolTipText = LookupToolTipLanguageCaption(9240, "Sets all elemental weights to the weight of the element's most common isotope (for high resolution mass spectrometry)")
        .cmdIntegerMass.ToolTipText = LookupToolTipLanguageCaption(9245, "Sets all elemental weights to the nominal integer weight of the element's most common isotope (for low resolution mass spectrometry)")
        .grdElem.ToolTipText = LookupToolTipLanguageCaption(9340, "Click to change an element's weight or uncertainty")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmEditAbbrev"
    With frmEditAbbrev
        .cmdReset.ToolTipText = LookupToolTipLanguageCaption(9010, "Resets the abbreviations to the program defaults")
        .grdAmino.ToolTipText = LookupToolTipLanguageCaption(9140, "Click to change an abbreviation")
        .grdNormal.ToolTipText = LookupToolTipLanguageCaption(9140, "Click to change an abbreviation")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmFinder"
    With frmFinder
        .txtHits.ToolTipText = LookupToolTipLanguageCaption(10240, "Maximum number of target compounds to find")
        .txtWeightTolerance.ToolTipText = LookupToolTipLanguageCaption(10220, "Amount that target compound's weight can be from the target weight")
        .txtPercentTolerance.ToolTipText = LookupToolTipLanguageCaption(10230, "Amount that elemental percent compositions can be from the target percentage")
        .cmdCalculate.ToolTipText = LookupToolTipLanguageCaption(10310, "Find the compounds that match the specified parameters")
        .cmdCopyAsRTF.ToolTipText = LookupToolTipLanguageCaption(10330, "Copy results to clipboard in Rich Text Format")
        .cmdCopy.ToolTipText = LookupToolTipLanguageCaption(10340, "Copy results to clipboard")
        .cmdDisplayIsotopicAbundance.ToolTipText = LookupToolTipLanguageCaption(10345, "Display the Isotopic Distribution of the currently selected compound (Ctrl+D)")
        .cmdFinderOptions.ToolTipText = LookupToolTipLanguageCaption(10300, "Shortcut: Ctrl+O")
        For intIndex = 0 To 9
            .txtMin(intIndex).ToolTipText = LookupToolTipLanguageCaption(10250, "Minimum number of atoms in target compound") & _
                                     " (" & .chkElements(intIndex).Caption & ")"
            .txtMax(intIndex).ToolTipText = LookupToolTipLanguageCaption(10255, "Maximum number of atoms in target compound") & _
                                     " (" & .chkElements(intIndex).Caption & ")"
            .txtPercent(intIndex).ToolTipText = LookupToolTipLanguageCaption(10260, "Percent composition of target compound that is") & _
                                         " " & .chkElements(intIndex).Caption
        Next intIndex
        For intIndex = 4 To 9
            .txtWeight(intIndex).ToolTipText = LookupToolTipLanguageCaption(10270, "Type a weight for the custom element, an elemental symbol, or an abbreviation")
        Next intIndex
    End With

    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmFinderOptions"
    With frmFinderOptions
        .chkFindCharge.ToolTipText = LookupToolTipLanguageCaption(10900, "Compute the overall charge of each compound found")
        .chkLimitChargeRange.ToolTipText = LookupToolTipLanguageCaption(10910, "Limit the displayed compounds to a specific charge range")
        .chkFindMtoZ.ToolTipText = LookupToolTipLanguageCaption(10920, "Compute the mass to charge ratio for each compound found")
        .chkFindTargetMtoZ.ToolTipText = LookupToolTipLanguageCaption(10930, "Find compounds with m/z values equivalent to the target")
        .chkSort.ToolTipText = LookupToolTipLanguageCaption(10940, "Convert results to empirical formulas and sort them")
        .chkVerifyHydrogens.ToolTipText = LookupToolTipLanguageCaption(10950, "Limit number of hydrogen atoms in found compounds to a realistic number")
        .chkAutoSetBounds.ToolTipText = LookupToolTipLanguageCaption(10960, "Automatically adjust the Min and Max search values to a valid range for the given target weight")
        .txtChargeMin.ToolTipText = LookupToolTipLanguageCaption(10820, "Minimum charge to limit compounds to")
        .txtChargeMax.ToolTipText = LookupToolTipLanguageCaption(10830, "Maximum charge to limit compounds to")
        .cboSearchType.ToolTipText = LookupToolTipLanguageCaption(10840, "A thorough search finds all matching compounds while a bounded search finds compounds within a specific atomic range (thorough is usually faster).")
        .cboSortResults.ToolTipText = LookupToolTipLanguageCaption(10850, "Method to sort results.  Re-calculate to re-sort.")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmMain"
    With frmMain
        .rtfFormulaSingle.ToolTipText = LookupToolTipLanguageCaption(5050, "Type the molecular formula here")
        .grdPC.ToolTipText = LookupToolTipLanguageCaption(5200, "Click to set or reset a target value")
        .cmdNewFormula.ToolTipText = ConstructAddFormulaToolTip
        .cmdCalculate.ToolTipText = LookupToolTipLanguageCaption(5100, "Determines the molecular weight of the current formula")
        .lblStatus.ToolTipText = LookupToolTipLanguageCaption(5300, "Double click the status line to expand it")
        For intIndex = 0 To frmMain.GetTopFormulaIndex
            .rtfFormula(intIndex).ToolTipText = LookupToolTipLanguageCaption(5050, "Type the molecular formula here")
        Next intIndex
        .optElementMode(0).ToolTipText = LookupToolTipLanguageCaption(5020, "Use average weights of elements")
        .optElementMode(1).ToolTipText = LookupToolTipLanguageCaption(5030, "Use weight of most common isotope")
        .optElementMode(2).ToolTipText = LookupToolTipLanguageCaption(5040, "Use nominal integer weight of most common isotope")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmMMConvert"
    With frmMMConvert
        .cboAction.ToolTipText = LookupToolTipLanguageCaption(11010, "Perform conversions between different amounts of compound or perform molarity-related calculations")
        .SetDynamicToolTips
        .txtCustomMass.ToolTipText = LookupToolTipLanguageCaption(7600, "Enter custom numerical mass for use in computations")
        .cmdFindAmount.ToolTipText = LookupToolTipLanguageCaption(11150, "Calculate the amount using the volume and molarity")
        .cmdFindVolume.ToolTipText = LookupToolTipLanguageCaption(11160, "Calculate the volume using the amount and molarity")
        .cmdFindConcentration.ToolTipText = LookupToolTipLanguageCaption(11170, "Calculate the concentration using the amount and volume")
        .txtDensity.ToolTipText = LookupToolTipLanguageCaption(11110, "Density of the compound")
                ' Note: cboTo has .Tag = 11020 so that combo box elements match cboFrom
                '       However, its .ToolTipText phrase is stored in 11041, thus we use 11040 in the next statement
        .cboTo.ToolTipText = LookupToolTipLanguageCaption(11040, "Units of amount to convert to")
        .txtVolume.ToolTipText = LookupToolTipLanguageCaption(11120, "Volume of solvent the compound is dissolved in")
        .cboVolume.ToolTipText = LookupToolTipLanguageCaption(11030, "Units of volume")
        .txtConcentration.ToolTipText = LookupToolTipLanguageCaption(11130, "Concentration of compound in the solvent")
        .cboConcentration.ToolTipText = LookupToolTipLanguageCaption(11050, "Units of concentration")
        
        .cboDilutionMode.ToolTipText = LookupToolTipLanguageCaption(11210, "Quantity to find for dilution calculations")
        .chkLinkMolarities.ToolTipText = LookupToolTipLanguageCaption(11220, "Copy the Computed Concentration for converting amounts to the Initial Concentration for dilutions and vice versa if either changes")
        .chkLinkDilutionVolumeUnits.ToolTipText = LookupToolTipLanguageCaption(11230, "Synchronize the units for the Volume of Stock, Volume of Solvent, and Final Total Volume")
        
        .txtDilutionConcentrationInitial.ToolTipText = LookupToolTipLanguageCaption(11255, "Concentration of solute in stock solution")
        .txtStockSolutionVolume.ToolTipText = LookupToolTipLanguageCaption(11265, "Volume (aliquot) of stock solution to remove when performing dilution")
        .txtDilutionConcentrationFinal.ToolTipText = LookupToolTipLanguageCaption(11275, "Concentration of solute in final solution following dilution")
        .txtDilutingSolventVolume.ToolTipText = LookupToolTipLanguageCaption(11285, "Volume of solvent to mix with the stock solution (aliquot) removed for dilution")
        .txtTotalVolume.ToolTipText = LookupToolTipLanguageCaption(11295, "Total volume of the final solution following mixing of stock and diluting solvent")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmProgramPreferences"
    With frmProgramPreferences
        .optAbbrevType(0).ToolTipText = LookupToolTipLanguageCaption(11650, "Recognize normal abbreviations, but not amino acids")
        .optAbbrevType(1).ToolTipText = LookupToolTipLanguageCaption(11655, "Recognize normal abbreviations and amino acids")
        .optAbbrevType(2).ToolTipText = LookupToolTipLanguageCaption(11660, "Ignore all abbreviations")
        .optConvertType(0).ToolTipText = LookupToolTipLanguageCaption(11665, "Correctly capitalize formulas while parsing")
        .optConvertType(1).ToolTipText = LookupToolTipLanguageCaption(11670, "Require user to type formulas with correct capitalization")
        .optConvertType(2).ToolTipText = LookupToolTipLanguageCaption(11675, "Interpret lowercase formulas, and do not capitalize them")
        .optStdDevType(0).ToolTipText = LookupToolTipLanguageCaption(11680, "Display standard deviations in abbreviated form")
        .optStdDevType(1).ToolTipText = LookupToolTipLanguageCaption(11685, "Display standard deviations in scientific notation")
        .optStdDevType(2).ToolTipText = LookupToolTipLanguageCaption(11690, "Display standard deviations in long decimal form")
        .optStdDevType(3).ToolTipText = LookupToolTipLanguageCaption(11695, "Do not display standard deviations")
        For intIndex = 0 To 3
            .optExitConfirmation(intIndex).ToolTipText = LookupToolTipLanguageCaption(11700, "Determines if the escape key exits program and whether to confirm program exit")
        Next intIndex
        .chkAdvanceOnCalculate.ToolTipText = LookupToolTipLanguageCaption(11750, "Move to a new formula line after calculating a formula's weight")
        .chkBracketsAsParentheses.ToolTipText = LookupToolTipLanguageCaption(11760, "Treat brackets, [ and ], as parentheses, rather than as percent solver placeholders")
        .chkAutoCopyCurrentMWT.ToolTipText = LookupToolTipLanguageCaption(11770, "Automatically copy the selected formula's molecular weight value to the clipboard after each calculation")
        .chkComputeCharge.ToolTipText = LookupToolTipLanguageCaption(11780, "Compute charge of compounds (very basic rules, cannot correct for double or triple bonds, etc.)")
        .chkAlwaysSwitchToIsotopic.ToolTipText = LookupToolTipLanguageCaption(11800, "Switch to isotopic weights automatically upon entering the formula finder or peptide fragmentation module")
        .chkNeverShowFormulaFinderWarning.ToolTipText = LookupToolTipLanguageCaption(11810, "Never prompt about current weight mode when entering the formula finder or peptide fragmentation module")
        .chkAutosaveValues.ToolTipText = LookupToolTipLanguageCaption(11820, "Automatically save options, values, and formulas on program exit")
        .chkShowCaution.ToolTipText = LookupToolTipLanguageCaption(11830, "Caution when possibly confusing elemental combinations are in formulas (like Co vs. CO)")
        .chkShowQuickSwitch.ToolTipText = LookupToolTipLanguageCaption(11840, "Show option to quickly change elemental weight modes")
        .chkShowToolTips.ToolTipText = LookupToolTipLanguageCaption(11850, "Show short help messages when mouse passes over certain buttons and areas")
        .chkHighlightTextFields.ToolTipText = LookupToolTipLanguageCaption(11860, "Highlight the entire text field when moving to it")
        .chkHideInactiveForms.ToolTipText = LookupToolTipLanguageCaption(11870, "Hide the main program window when using the Formula Finder, Mole/Mass calculator, etc.")
        .cboStartupModule.ToolTipText = LookupToolTipLanguageCaption(11895, "Module to show when program starts")
        .cboMaximumFormulasToShow.ToolTipText = LookupToolTipLanguageCaption(11880, "Choose a smaller number to prevent the formula window from filling the screen.  If lowering, must exit and restart program before takes effect.  Maximum is dependent upon screen resolution.")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmCapillaryCalcs"
    With frmCapillaryCalcs
        .cboCapillaryType.ToolTipText = LookupToolTipLanguageCaption(7010, "Toggle between open and packed capillaries")
        .txtCapValue(cctViscosity).ToolTipText = LookupToolTipLanguageCaption(7850, "Typical viscosity value is 0.0089 poise")
        .txtCapValue(cctPorosity).ToolTipText = LookupToolTipLanguageCaption(7860, "Typical porosity value is 0.4")
        .txtCapValue(cctBdDiffusionCoefficient).ToolTipText = LookupToolTipLanguageCaption(7870, "Typical diffusion coefficient for small organics is 0.00001, i.e. 1E-5; Typical value for peptides is 0.000005, i.e. 5E-6")
        .txtCustomMass.ToolTipText = LookupToolTipLanguageCaption(7600, "Enter custom numerical mass for use in computations")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmFragmentationModelling"
    With frmFragmentationModelling
        .cboNotation.ToolTipText = LookupToolTipLanguageCaption(12010, "Amino acid sequence notation type")
        .txtSequence.ToolTipText = LookupToolTipLanguageCaption(12050, "Enter amino acid sequence here")
        .txtAlignment.ToolTipText = LookupToolTipLanguageCaption(12060, "Shifts the loaded ions to be matched by the given amount to correct for post-translational modifications.")
        .lstIonsToModify.ToolTipText = LookupToolTipLanguageCaption(12235, "Choose ions to which losses will be applied")
        .cboDoubleCharge.ToolTipText = LookupToolTipLanguageCaption(12285, "The 2+ m/z value will be computed for ions above this m/z")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmFragmentationModelling"
    With frmIonMatchOptions
        .lblFragSpectrumColor.ToolTipText = LookupToolTipLanguageCaption(14257, "Click to change")
        .lblMatchingIonDataColor.ToolTipText = LookupToolTipLanguageCaption(14257, "Click to change")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmAminoAcidModificationSymbols"
    With frmAminoAcidModificationSymbols
        .grdModSymbols.ToolTipText = LookupToolTipLanguageCaption(15100, "Click to change a modification symbol")
        .cmdReset.ToolTipText = LookupToolTipLanguageCaption(15300, "Resets modification symbols to defaults")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmDtaTxtFileBrowser"
    With frmDtaTxtFileBrowser
        .cmdJumpToScan.ToolTipText = LookupToolTipLanguageCaption(16100, "Shortcut is Ctrl+J")
    End With
    
    If blnShowDebugPrompts Then MsgBox "Set ToolTips: frmIsotopicDistribution"
    With frmIsotopicDistribution
        .lblPlotColor.ToolTipText = LookupToolTipLanguageCaption(15125, "Click to change")
        .lblComparisonListPlotColor.ToolTipText = LookupToolTipLanguageCaption(15125, "Click to change")
        .txtResults.ToolTipText = LookupToolTipLanguageCaption(15150, "Isotopic distribution results")
    End With
    
    Exit Sub

SetAllTooltipsErrorHandler:
    GeneralErrorHandler "MwtWinProcedures|SetAllTooltips", Err.Number, Err.Description

End Sub

Public Sub SetDefaultOptions()
        
On Error GoTo SetDefaultsErrorHandler
    
    With frmMain
        .SetViewMode vmdMultiView
        .ShowHideQuickSwitch vbChecked
    End With

    With frmProgramPreferences
        .optConvertType(0).value = True
        .optAbbrevType(1).value = True
        .optStdDevType(0).value = True
        gBlnShowStdDevWithMass = True
        
        .chkShowCaution.value = vbChecked
        .chkAdvanceOnCalculate.value = vbUnchecked
        .chkComputeCharge.value = vbUnchecked
        .chkShowQuickSwitch.value = vbChecked
        
        SetFonts "Arial", 10
        
        .optExitConfirmation(exmEscapeKeyConfirmExit).value = True
        .chkShowToolTips.value = vbChecked
        .chkHideInactiveForms.value = vbUnchecked
        .chkAutosaveValues.value = vbChecked
        .chkBracketsAsParentheses = vbUnchecked
        .chkAutoCopyCurrentMWT.value = vbUnchecked
        
        .cboStartupModule.ListIndex = 0
        .cboMaximumFormulasToShow = "9"
        
        .chkAlwaysSwitchToIsotopic.value = vbUnchecked
        .chkNeverShowFormulaFinderWarning.value = vbUnchecked
        
        .chkHighlightTextFields = vbUnchecked
    End With
    
    frmFinderOptions.cboSearchType.ListIndex = 0
    
    Exit Sub
    
SetDefaultsErrorHandler:
    GeneralErrorHandler "MwtWinProcedures|SetDefaults", Err.Number, Err.Description
    Resume Next
    
End Sub

Private Sub SetDefaultCapillaryFlowValues()

    gCapFlowComputationTypeSave = 3                     ' cptVolFlowRate = 3
    gCapFlowLinkMassRateFlowRateSave = 1                ' Link Flow Rate
    gCapFlowLinkBdLinearVelocitySave = 1                ' Link Linear Velocity
    gCapFlowShowPeakBroadeningSave = 0                  ' Hide the peak broadening screen
    
    OpenCapVals.TextValues(cctPressure) = 50                        ' Pressure
    OpenCapVals.ComboValues(cccPressureUnits) = uprPsi              ' Pressure Units
    OpenCapVals.TextValues(cctColumnLength) = 50                    ' Col Length
    OpenCapVals.ComboValues(cccColumnLengthUnits) = ulnCM           ' Col Length
    OpenCapVals.TextValues(cctColumnID) = 30                        ' Col ID
    OpenCapVals.ComboValues(cccColumnIDUnits) = ulnMicrons          ' Column Inner Diameter Units
    OpenCapVals.TextValues(cctViscosity) = 0.0089                   ' Viscosity
    OpenCapVals.ComboValues(cccViscosityUnits) = uviPoise           ' Viscosity Units
    OpenCapVals.TextValues(cctParticleDiamter) = 5                  ' Particle Diameter
    OpenCapVals.ComboValues(cccParticleDiameterUnits) = ulnMicrons  ' Particle Diameter Units
    OpenCapVals.TextValues(cctFlowRate) = 100                       ' Flow Rate
    OpenCapVals.ComboValues(cccFlowRateUnits) = ufrNLPerMin         ' Flow Rate Units
    OpenCapVals.ComboValues(cccLinearVelocityUnits) = ulvCmPerSec   ' Linear Velocity Units
    OpenCapVals.TextValues(cctDeadTime) = 1                         ' Dead time
    OpenCapVals.ComboValues(cccDeadTimeUnits) = utmSeconds          ' Dead Time Units
    OpenCapVals.TextValues(cctPorosity) = 0.4                       ' Porosity
    OpenCapVals.ComboValues(cccVolumeUnits) = uvoNL                 ' Volume Units
    
    OpenCapVals.TextValues(cctMassRateConcentration) = 1                    ' Concentration
    OpenCapVals.ComboValues(cccMassRateConcentrationUnits) = ucoMicroMolar  ' Concentration Units
    OpenCapVals.TextValues(cctMassRateVolFlowRate) = 100                    ' MassRateVolFlowRate
    OpenCapVals.ComboValues(cccMassRateVolFlowRateUnits) = ufrNLPerMin      ' MassRateVolFlowRate Units
    OpenCapVals.TextValues(cctMassRateInjectionTime) = 30                   ' Injection Time
    OpenCapVals.ComboValues(cccMassRateInjectionTimeUnits) = utmSeconds     ' Injection Time Units
    OpenCapVals.ComboValues(cccMassFlowRateUnits) = umfFmolPerSec           ' Mass Flow Rate
    OpenCapVals.ComboValues(cccMassRateMolesInjectedUnits) = umaFemtoMoles  ' Moles Injected Units
    
    OpenCapVals.TextValues(cctBdLinearVelocity) = 1                         ' Broadening Linear velocity
    OpenCapVals.ComboValues(cccBdLinearVelocityUnits) = ulvCmPerSec         ' Broadening Linear Velocity Units
    OpenCapVals.TextValues(cctBdDiffusionCoefficient) = 0.000005            ' Broadening diffusion coefficient
    OpenCapVals.TextValues(cctBdOpenTubeLength) = 5                         ' Open Tube Length
    OpenCapVals.ComboValues(cccBdOpenTubeLengthUnits) = ulnCM               ' Open Tube Col Length
    OpenCapVals.TextValues(cctBdOpenTubeID) = 50                            ' Open Tube Inner Diameter
    OpenCapVals.ComboValues(cccBdOpenTubeIDUnits) = ulnMicrons              ' Open Tube Inner Diameter Units
    OpenCapVals.TextValues(cctBdInitialPeakWidth) = 30                      ' Initial Peak Width
    OpenCapVals.ComboValues(cccBdInitialPeakWidthUnits) = utmSeconds        ' Initial Peak Width
    OpenCapVals.TextValues(cctBdAdditionalVariance) = 0                     ' Additional Variance (in square seconds)
    OpenCapVals.ComboValues(cccBdResultantPeakWidthUnits) = utmSeconds      ' Resultant Peak Width Units
    
    ' Most of the packed capillary values are identical to the open capillary values
    '
    PackedCapVals.TextValues(cctPressure) = 3000                            ' Pressure
    PackedCapVals.ComboValues(cccPressureUnits) = OpenCapVals.ComboValues(cccPressureUnits)
    PackedCapVals.TextValues(cctColumnLength) = 50                          ' Column Length
    PackedCapVals.ComboValues(cccColumnLengthUnits) = OpenCapVals.ComboValues(cccColumnLengthUnits)
    PackedCapVals.TextValues(cctColumnID) = 75                              ' Column ID
    PackedCapVals.ComboValues(cccColumnIDUnits) = OpenCapVals.ComboValues(cccColumnIDUnits)
    PackedCapVals.TextValues(cctViscosity) = OpenCapVals.TextValues(cctViscosity)
    PackedCapVals.ComboValues(cccViscosityUnits) = OpenCapVals.ComboValues(cccViscosityUnits)
    PackedCapVals.TextValues(cctParticleDiamter) = OpenCapVals.TextValues(cctParticleDiamter)
    PackedCapVals.ComboValues(cccParticleDiameterUnits) = OpenCapVals.ComboValues(cccParticleDiameterUnits)
    PackedCapVals.TextValues(cctFlowRate) = OpenCapVals.TextValues(cctFlowRate)
    PackedCapVals.ComboValues(cccFlowRateUnits) = OpenCapVals.ComboValues(cccFlowRateUnits)
    PackedCapVals.ComboValues(cccLinearVelocityUnits) = OpenCapVals.ComboValues(cccLinearVelocityUnits)
    PackedCapVals.TextValues(cctDeadTime) = OpenCapVals.TextValues(cctDeadTime)
    PackedCapVals.ComboValues(cccDeadTimeUnits) = OpenCapVals.ComboValues(cccDeadTimeUnits)
    PackedCapVals.TextValues(cctPorosity) = OpenCapVals.TextValues(cctPorosity)
    PackedCapVals.ComboValues(cccVolumeUnits) = OpenCapVals.ComboValues(cccVolumeUnits)
    
    PackedCapVals.TextValues(cctMassRateConcentration) = OpenCapVals.TextValues(cctMassRateConcentration)
    PackedCapVals.ComboValues(cccMassRateConcentrationUnits) = OpenCapVals.ComboValues(cccMassRateConcentrationUnits)
    PackedCapVals.TextValues(cctMassRateVolFlowRate) = OpenCapVals.TextValues(cctMassRateVolFlowRate)
    PackedCapVals.ComboValues(cccMassRateVolFlowRateUnits) = OpenCapVals.ComboValues(cccMassRateVolFlowRateUnits)
    PackedCapVals.TextValues(cctMassRateInjectionTime) = OpenCapVals.TextValues(cctMassRateInjectionTime)
    PackedCapVals.ComboValues(cccMassRateInjectionTimeUnits) = OpenCapVals.ComboValues(cccMassRateInjectionTimeUnits)
    PackedCapVals.ComboValues(cccMassFlowRateUnits) = OpenCapVals.ComboValues(cccMassFlowRateUnits)
    PackedCapVals.ComboValues(cccMassRateMolesInjectedUnits) = OpenCapVals.ComboValues(cccMassRateMolesInjectedUnits)
   
    PackedCapVals.TextValues(cctBdLinearVelocity) = OpenCapVals.TextValues(cctBdLinearVelocity)
    PackedCapVals.ComboValues(cccBdLinearVelocityUnits) = OpenCapVals.ComboValues(cccBdLinearVelocityUnits)
    PackedCapVals.TextValues(cctBdDiffusionCoefficient) = OpenCapVals.TextValues(cctBdDiffusionCoefficient)
    PackedCapVals.TextValues(cctBdOpenTubeLength) = OpenCapVals.TextValues(cctBdOpenTubeLength)
    PackedCapVals.ComboValues(cccBdOpenTubeLengthUnits) = OpenCapVals.ComboValues(cccBdOpenTubeLengthUnits)
    PackedCapVals.TextValues(cctBdOpenTubeID) = OpenCapVals.TextValues(cctBdOpenTubeID)
    PackedCapVals.ComboValues(cccBdOpenTubeIDUnits) = OpenCapVals.ComboValues(cccBdOpenTubeIDUnits)
    PackedCapVals.TextValues(cctBdInitialPeakWidth) = OpenCapVals.TextValues(cctBdInitialPeakWidth)
    PackedCapVals.ComboValues(cccBdInitialPeakWidthUnits) = OpenCapVals.ComboValues(cccBdInitialPeakWidthUnits)
    PackedCapVals.TextValues(cctBdAdditionalVariance) = OpenCapVals.TextValues(cctBdAdditionalVariance)
    PackedCapVals.ComboValues(cccBdLinearVelocityUnits) = OpenCapVals.ComboValues(cccBdLinearVelocityUnits)
    PackedCapVals.ComboValues(cccBdOpenTubeIDUnits) = OpenCapVals.ComboValues(cccBdOpenTubeIDUnits)
    PackedCapVals.ComboValues(cccBdResultantPeakWidthUnits) = OpenCapVals.ComboValues(cccBdResultantPeakWidthUnits)
    
End Sub

Public Sub SetDefaultValuesAndFormulas(blnResetToDefaultFormulas As Boolean, Optional blnShowDebugPrompts As Boolean = False)
    
    Dim intFormulaCount As Integer
    Dim blnSuccess As Boolean
    Dim intIndex As Integer
    Dim eViewModeSaved As vmdViewModeConstants
    
On Error GoTo SetDefaultValuesAndFormulasErrorHandler

    ' Make sure we have at least 6 formulas
    If blnResetToDefaultFormulas Then
        With frmMain
            frmMain.Show

            eViewModeSaved = .GetViewMode()
            
            If eViewModeSaved <> vmdMultiView Then
                .SetViewMode vmdMultiView
            End If
            
            intFormulaCount = .GetTopFormulaIndex() + 1
            Do While intFormulaCount < 6
                blnSuccess = .AddNewFormulaWrapper(True)
                Debug.Assert blnSuccess
                intFormulaCount = intFormulaCount + 1
            Loop
    
            .EraseAllFormulas False
    
            .rtfFormula(0).Text = "BrCH2(CH2)7CH2Br"
            .rtfFormula(1).Text = "FeCl3-6H2O"
            .rtfFormula(2).Text = "Co(Bpy)(CO)4"
            .rtfFormula(3).Text = "^13C6H6-.1H2O"
            .rtfFormula(4).Text = "HGlyLeuTyrOH"
            .rtfFormula(5).Text = "BrCH2(CH2)7CH2Br>CH8"
    
            .SetFocusToFormulaByIndex 0
            
            If eViewModeSaved = vmdSingleView Then
                .SetViewMode vmdSingleView
            End If
        End With
    End If
    
    With frmAminoAcidConverter
        .txt1LetterSequence = "GLY"
        .txt3LetterSequence = "Gly-Leu-Tyr"
        SetCheckBox .chkSpaceEvery10, True
        SetCheckBox .chkSeparateWithDash, True
    End With
    
    With frmMMConvert
        SetCheckBox .chkLinkMolarities, False
        SetCheckBox .chkLinkDilutionVolumeUnits, True
        
        .optWeightSource(0).value = True
        .txtCustomMass.Text = "100"
        .cboAction.ListIndex = 2                        ' mmaDilutionCalculations
        .txtFromNum.Text = "1"
        .cboFrom.ListIndex = uamMoles
        .txtDensity.Text = "1"
        .cboTo.ListIndex = uamGrams
        .txtVolume.Text = "100"
        .cboVolume.ListIndex = uevML
        .txtConcentration.Text = "1"
        .cboConcentration.ListIndex = ummcMolar
        
        .cboDilutionMode.ListIndex = 0                  ' dmFindRequiredDilutionVolumes
        .txtDilutionConcentrationInitial.Text = "10"
        .cboDilutionConcentrationInitial.ListIndex = ummcMolar
        .txtStockSolutionVolume.Text = "3"
        .cboStockSolutionVolume.ListIndex = uevML
        .txtDilutionConcentrationFinal.Text = "2"
        .cboDilutionConcentrationFinal.ListIndex = ummcMolar
        .txtDilutingSolventVolume.Text = "12"
        .cboDilutingSolventVolume.ListIndex = uevML
        .txtTotalVolume.Text = "15"
        .cboTotalVolume.ListIndex = uevML
    End With
    
    frmCalculator.rtfExpression.Text = "6^3-8.5*3"
    
    With frmFinder
        .optType(0).value = True
        .txtMWT.Text = "200"
        .txtPercentMaxWeight.Text = "400"
        SetCheckBox .chkPpmMode, False
        SetCheckBox .chkShowDeltaMass, True
        .txtWeightTolerance.Text = "0.05"
        .txtPercentTolerance.Text = "1"
        .txtHits.Text = "100"
    End With
    
    With frmFinderOptions
        SetCheckBox .chkAutoSetBounds, True
        .cboSortResults.ListIndex = 0
        SetCheckBox .chkSort, True
        SetCheckBox .chkVerifyHydrogens, True
        SetCheckBox .chkFindCharge, False
        SetCheckBox .chkFindMtoZ, False
        SetCheckBox .chkLimitChargeRange, False
        .txtChargeMin.Text = "-4"
        .txtChargeMax.Text = "4"
        SetCheckBox .chkFindTargetMtoZ, False
        .UpdateCheckBoxes
    End With
    
    With frmFinder
        '.InitializeForm
        For intIndex = 0 To .GetMaxElementCount() - 1
            .txtMin(intIndex).Text = 0
            .txtMax(intIndex).Text = 10
            
            ' By default, check the first 4 elements
            SetCheckBox .chkElements(intIndex), (intIndex <= 3)
            
            .txtPercent(intIndex).Text = LookupLanguageCaption(10260, "Percent")
            If intIndex >= 4 Then
                .txtWeight(intIndex).Text = LookupLanguageCaption(10270, "# or Element or Abbrev.")
            End If
        Next intIndex
    End With
    
    SetDefaultCapillaryFlowValues
    If frmCapillaryCalcs.cboCapillaryType.ListIndex <> 0 Then
        ' Setting .ListIndex to 0 will automatically call UpdateCapillaryType
        frmCapillaryCalcs.cboCapillaryType.ListIndex = 0
    Else
        frmCapillaryCalcs.UpdateCapillaryType
    End If

    With frmFragmentationModelling
        .cboNotation.ListIndex = 1
        .txtSequence = "Arg-His-Pro-Glu-Tyr-Ala-Val"
        .cboNTerminus.ListIndex = ntgHydrogen
        .cboCTerminus.ListIndex = ctgHydroxyl
        SetCheckBox .chkIonType(0), False
        SetCheckBox .chkIonType(1), True
        SetCheckBox .chkIonType(2), True
        .lstIonsToModify.Selected(0) = False
        .lstIonsToModify.Selected(1) = True
        .lstIonsToModify.Selected(2) = True
        SetCheckBox .chkWaterLoss, True
        SetCheckBox .chkAmmoniaLoss, False
        SetCheckBox .chkPhosphateLoss, False
        SetCheckBox .chkDoubleCharge, False
        SetCheckBox .chkTripleCharge, False
        .cboDoubleCharge.ListIndex = 8
        .cboTripleCharge.ListIndex = 9
        SetCheckBox .chkRemovePrecursorIon, True
        .txtPrecursorIonMass = "300"
        .txtPrecursorMassWindow = "2"
        .txtIonMatchingWindow = "0.5"
        .txtAlignment = "0"
    End With

    With frmIonMatchOptions
        .txtIonIntensity(0) = "10"
        .txtIonIntensity(1) = "50"
        .txtIonIntensity(2) = "50"
        .txtBYIonShoulders = "25"
        .txtNeutralLosses = "10"
        SetCheckBox .chkGroupSimilarIons, False
        .txtGroupIonMassWindow = ".5"
        .txtNormalizedIntensity = "100"
        .txtIonCountToUse = "200"
        .txtMassRegions = "1"
        SetCheckBox .chkFragSpecLabelMainIons, True
        SetCheckBox .chkFragSpecLabelOtherIons, False
        SetCheckBox .chkFragSpecEmphasizeProlineYIons, False
        SetCheckBox .chkPlotSpectrumInverted, True
        SetCheckBox .chkAutoLabelMass, False
        .lblFragSpectrumColor.BackColor = RGB(0, 0, 255)
        .lblMatchingIonDataColor.BackColor = RGB(0, 128, 0)
    End With
    
    frmIsotopicDistribution.EnableDisablePlotUpdates True
    With frmIsotopicDistribution
        .cboPlotType.ListIndex = 1                      ' ipmGaussian
        .lblPlotColor.BackColor = vbBlue
        .txtEffectiveResolution = 5000
        .txtEffectiveResolutionMass = 1000
        .txtGaussianQualityFactor = 50
        
        .lblComparisonListPlotColor.BackColor = vbRed
        .cboComparisonListPlotType.ListIndex = 0        ' ipmSticksToZero
        SetCheckBox .chkComparisonListNormalize, True
        
        SetCheckBox .chkAutoLabelPeaks, False
        SetCheckBox .chkPlotResults, True
    End With
    
    If blnShowDebugPrompts Then MsgBox "Call objMwtWin.Peptide.SetDefaultModificationSymbols"
    objMwtWin.Peptide.SetDefaultModificationSymbols

    If blnShowDebugPrompts Then MsgBox "Call frmFragmentationModelling.ResetCWSpectrumOptions"
    frmFragmentationModelling.ResetCWSpectrumOptions
    
    If blnShowDebugPrompts Then MsgBox "Call frmIsotopicDistribution.ResetCWSpectrumOptions"
    frmIsotopicDistribution.ResetCWSpectrumOptions

    frmIsotopicDistribution.EnableDisablePlotUpdates False
    
    Exit Sub

SetDefaultValuesAndFormulasErrorHandler:
    Debug.Assert False
    GeneralErrorHandler "MwtWinProcedures|SetDefaultValuesAndFormulas", Err.Number, Err.Description
    Resume Next
    
End Sub

Public Sub SetFonts(strFontName As String, intFontSize As Integer)
    Dim intIndex As Integer
    
On Error GoTo SetFontsErrorHandler

    If Len(strFontName) = 0 Then Exit Sub
    
    objMwtWin.RtfFontName = strFontName
    objMwtWin.RtfFontSize = intFontSize
    
    With frmMain
        .rtfFormulaSingle.Font.Name = objMwtWin.RtfFontName
        .rtfFormulaSingle.Font.Size = objMwtWin.RtfFontSize
        
        With .txtMWTSingle
            .Font = objMwtWin.RtfFontName
            .FontSize = objMwtWin.RtfFontSize
            .FontItalic = False
        End With
        For intIndex = 0 To frmMain.GetTopFormulaIndex
            .rtfFormula(intIndex).Font.Name = objMwtWin.RtfFontName
            .rtfFormula(intIndex).Font.Size = objMwtWin.RtfFontSize
            With .txtMWT(intIndex)
                .Font = objMwtWin.RtfFontName
                .FontSize = objMwtWin.RtfFontSize
                .FontItalic = False
            End With
            ' Make changes appear
            .rtfFormula(intIndex).TextRTF = objMwtWin.TextToRTF(.rtfFormula(intIndex).Text)
        Next intIndex
    End With
    
    frmMain.Refresh
    
    frmFinder.lstResults.FontName = objMwtWin.RtfFontName
    frmFinder.rtfResults.Font.Name = objMwtWin.RtfFontName
    frmFinder.rtfResults.Font.Size = objMwtWin.RtfFontSize
    frmFinder.Refresh
    
    Exit Sub
    
SetFontsErrorHandler:
    GeneralErrorHandler "MwtWinProcedures|SetFonts", Err.Number, Err.Description
    
End Sub

Public Sub ShowDefaultFormAtLoad()
    Dim intListIndex As Integer
    
    ' 0: Main
    ' 1: Formula Finder
    ' 2: Capillary Flow Calculator
    ' 3: Mole/Mass Converter
    ' 4: Peptide Sequence Fragmentation Modeller
    ' 5: Amino Acid Notation Converter
    ' 6: Isotopic Distribution Modeller
    
    intListIndex = frmProgramPreferences.cboStartupModule.ListIndex
    Select Case intListIndex
    Case 1: frmMain.ShowFormulaFinder
    Case 2: frmMain.ShowCapillaryFlowModule
    Case 3: frmMain.ShowMoleMassConverter
    Case 4: frmMain.ShowPeptideSequenceModeller
    Case 5: frmMain.ShowAminoAcidNotationModule
    Case 6: frmMain.ShowIsoDistributionModeller
    Case Else
        ' Do nothing, let frmMain be shown
    End Select

End Sub

Public Sub ShowHelpPage(hwnd As Long, lngContextID As Long)
    On Error GoTo ErrHandler

    'hWnd is a Long defined elsewhere to be the window handle
    'that will be the parent to the help window.
    Dim hwndHelp As Long
    
    'The return value is the window handle of the created help window.
    hwndHelp = HtmlHelp(hwnd, App.HelpFile, HH_HELP_CONTEXT, lngContextID)
    
    Exit Sub
    
ErrHandler:
    ' User pressed cancel button
    Exit Sub

End Sub

Public Sub ShowHideMassInputControlsGlobal(frmThisForm As VB.Form)
    Dim blnCustomVisible As Boolean
    
    With frmThisForm
        If .optWeightSource(0).value = True Then
            .eMassMode = mmcComputedMass
        Else
            .eMassMode = mmcCustomMass
        End If
        
        If .eMassMode = mmcComputedMass Then
            blnCustomVisible = False
        Else
            blnCustomVisible = True
        End If
        
        .lblFormula.Visible = Not blnCustomVisible
        .rtfCurrentFormula.Visible = Not blnCustomVisible
        .lblMWT.Visible = Not blnCustomVisible
        .lblMWTValue.Visible = Not blnCustomVisible
        
        .lblCustomMass.Visible = blnCustomVisible
        .txtCustomMass.Visible = blnCustomVisible
    End With
End Sub

Public Sub SwitchWeightMode(eNewElementMode As emElementModeConstants, Optional blnRecalculateAllFormulas As Boolean = True, Optional blnSetFocusToCurrentFormula As Boolean = True)
    ' eNewElementMode = 1 is average weights        (emAverageMass)
    ' eNewElementMode = 2 is isotopic weights       (emIsotopicMass)
    ' eNewElementMode = 3 is integer weights        (emIntegerMass)
    
    If eNewElementMode < emAverageMass Or eNewElementMode > emIntegerMass Then
        eNewElementMode = emAverageMass
    End If
    
    objMwtWin.SetElementMode eNewElementMode
    
    If blnRecalculateAllFormulas Then
        RecalculateAllFormulas blnSetFocusToCurrentFormula
    End If

    ' Make sure QuickSwitch Element Mode value is correct
    frmMain.ShowHideQuickSwitch frmProgramPreferences.chkShowQuickSwitch.value

End Sub

Public Sub SwitchWeightModeInteger(intNewElementMode As Integer, Optional blnRecalculateAllFormulas As Boolean = True)
    ' This sub is used to call SwitchWeightMode when the integer equivalent of the weight mode is known
    
    If intNewElementMode = emIsotopicMass Then
        SwitchWeightMode emIsotopicMass, blnRecalculateAllFormulas
    ElseIf intNewElementMode = emIntegerMass Then
        SwitchWeightMode emIntegerMass, blnRecalculateAllFormulas
    Else
        ' Assume intNewElementMode = emAverageMass
        SwitchWeightMode emAverageMass, blnRecalculateAllFormulas
    End If
    
End Sub

Public Sub ShowPopupRightClickMenu(frmParentForm As VB.Form, rtfThisTextBox As RichTextBox, Optional blnEnableUndoMenu As Boolean = True, Optional blnAllowPasteAndDelete As Boolean = True)
    ' Enable/Disable the appropriate menus
    
    With frmParentForm
        .mnuRightClickUndo.Enabled = blnEnableUndoMenu
        
        .mnuRightClickCut.Enabled = (rtfThisTextBox.SelLength > 0)
        .mnuRightClickCopy.Enabled = (rtfThisTextBox.SelLength > 0)
        
        If blnAllowPasteAndDelete Then
            If Clipboard.GetFormat(vbCFText) Or Clipboard.GetFormat(vbCFRTF) Then
                .mnuRightClickPaste.Enabled = True
            Else
                .mnuRightClickPaste.Enabled = False
            End If
            
            .mnuRightClickDelete.Enabled = (rtfThisTextBox.SelLength > 0)
        Else
            .mnuRightClickPaste.Enabled = False
            .mnuRightClickDelete.Enabled = False
        End If
        
        .mnuRightClickSelectAll.Enabled = (Len(rtfThisTextBox.Text) > 0)
    
        frmParentForm.PopupMenu .mnuRightClick
    End With
    
End Sub

Public Sub SwitchTips(blnShowToolTips As Boolean)
    Dim ThisForm As VB.Form, ThisControl As Control
    
    If blnShowToolTips Then
        ' Set all the tooltips
        SetAllTooltips
    Else
        ' Cancel all the tooltips
        On Error Resume Next
        For Each ThisForm In Forms
            For Each ThisControl In ThisForm
                If Not (TypeOf ThisControl Is CommonDialog Or TypeOf ThisControl Is Menu) Then
                    ThisControl.ToolTipText = ""
                End If
            Next ThisControl
        Next ThisForm
    End If
    
    frmMain.mnuShowTips.Checked = blnShowToolTips
        
End Sub

Public Sub ZoomLine(strCaption As String)
    MsgBox LookupLanguageCaption(3600, "The entire line is") & vbCrLf & strCaption, vbOKOnly, LookupLanguageCaption(3605, "Zooming")
End Sub

