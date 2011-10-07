VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MolecularWeightCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Molecular Weight Calculator routines with ActiveX Class interfaces
'
' Based on Molecular Weight Calculator, v6.20 code, written by Matthew Monroe 1995-2002
'
' ActiveX Dll version written by Matthew Monroe in Richland, WA
'
' Project started: November 12, 2002

Private Const PROGRAM_DATE = "August 5, 2011"
Private Const PROGRAM_VERSION = "2.37"

' Duplicate of enum statements in ElementAndMassRoutines.bas
Public Enum emElementModeConstants
    emAverageMass = 1
    emIsotopicMass = 2
    emIntegerMass = 3
End Enum

Public Enum smStdDevModeConstants
    smShort = 0
    smScientific = 1
    smDecimal = 2
End Enum

Public Enum ccCaseConversionConstants
    ccConvertCaseUp = 0
    ccExactCase = 1
    ccSmartCase = 2
End Enum

Public Enum arAbbrevRecognitionModeConstants
    arNormalOnly = 0
    arNormalPlusAminoAcids = 1
    arNoAbbreviations = 2
End Enum

Public Enum esElementStatsConstants
    esMass = 0
    esUncertainty = 1
    esCharge = 2
End Enum

Private mDataInitialized As Boolean

Public Compound As New MWCompoundClass
Public CapFlow As New MWCapillaryFlowClass
Public Peptide As New MWPeptideClass

Public Sub ClearError()
    ResetErrorParamsInternal
End Sub

Public Function ComputeMass(ByVal strFormula As String) As Double
    Dim objCompound As New MWCompoundClass
  
    With objCompound
        ' Simply assigning strFormula to .Formula will update the Mass
        .Formula = strFormula
        ComputeMass = .Mass
    End With
    
    Set objCompound = Nothing

End Function

Public Function ComputeIsotopicAbundances(ByRef strFormulaIn As String, ByVal intChargeState As Integer, ByRef strResults As String, ByRef ConvolutedMSData2DOneBased() As Double, ByRef ConvolutedMSDataCount As Long, Optional ByVal strHeaderIsotopicAbundances As String = "Isotopic Abundances for", Optional ByVal strHeaderMass As String = "Mass", Optional ByVal strHeaderFraction As String = "Fraction", Optional ByVal strHeaderIntensity As String = "Intensity") As Integer
    ' Computes the Isotopic Distribution for a formula
    ' Returns 0 if success, or -1 if an error
    
    ComputeIsotopicAbundances = ComputeIsotopicAbundancesInternal(strFormulaIn, intChargeState, strResults, ConvolutedMSData2DOneBased(), ConvolutedMSDataCount, strHeaderIsotopicAbundances, strHeaderMass, strHeaderFraction, strHeaderIntensity, False)
End Function

Public Function ConvoluteMass(ByVal dblMassMZ As Double, ByVal intCurrentCharge As Integer, Optional ByVal intDesiredCharge As Integer = 1, Optional ByVal dblChargeCarrierMass As Double = 0) As Double
    ConvoluteMass = ConvoluteMassInternal(dblMassMZ, intCurrentCharge, intDesiredCharge, dblChargeCarrierMass)
End Function

Public Function GetAbbreviation(ByVal lngAbbreviationID As Long, ByRef strSymbol As String, ByRef strFormula As String, ByRef sngCharge As Single, ByRef blnIsAminoAcid As Boolean, Optional ByRef strOneLetterSymbol As String, Optional ByRef strComment As String, Optional ByRef blnInvalidSymbolOrFormula As Boolean) As Long
    ' Returns 0 if success, 1 if failure
    GetAbbreviation = GetAbbreviationInternal(lngAbbreviationID, strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment, blnInvalidSymbolOrFormula)
End Function

Public Function GetAbbreviationCount() As Long
    GetAbbreviationCount = GetAbbreviationCountInternal()
End Function

Public Function GetAbbreviationCountMax() As Long
    GetAbbreviationCountMax = MAX_ABBREV_COUNT
End Function

Public Function GetAbbreviationID(ByVal strSymbol As String) As Long
    ' Returns 0 if not found, the ID if found
    GetAbbreviationID = GetAbbreviationIDInternal(strSymbol)
End Function

Public Function GetAminoAcidSymbolConversion(strSymbolToFind As String, bln1LetterTo3Letter As Boolean) As String
    ' If bln1LetterTo3Letter = True, then converting 1 letter codes to 3 letter codes
    ' Returns the symbol, if found
    ' Otherwise, returns ""
    GetAminoAcidSymbolConversion = GetAminoAcidSymbolConversionInternal(strSymbolToFind, bln1LetterTo3Letter)
End Function

Public Function GetCautionStatement(ByVal lngCautionStatementID As Long, ByRef strSymbolCombo As String, ByRef strCautionStatement As String) As Long
    ' Returns the contents of CautionStatements() in the ByRef variables
    ' Returns 0 if success, 1 if failure
    GetCautionStatement = GetCautionStatementInternal(lngCautionStatementID, strSymbolCombo, strCautionStatement)
End Function

Public Function GetCautionStatementCount() As Long
    ' Returns the number of Caution Statements in memory
    GetCautionStatementCount = GetCautionStatementCountInternal()
End Function

Public Function GetCautionStatementID(ByVal strSymbolCombo As String) As Long
    ' Returns -1 if not found, the ID if found
    GetCautionStatementID = GetCautionStatementIDInternal(strSymbolCombo)
End Function

Public Function GetChargeCarrierMass() As Double
    GetChargeCarrierMass = GetChargeCarrierMassInternal()
End Function

Public Function GetElement(ByVal intElementID As Integer, ByRef strSymbol As String, ByRef dblMass As Double, ByRef dblUncertainty As Double, ByRef sngCharge As Single, ByRef intIsotopeCount As Integer) As Long
    ' Returns the settings for the element with intElementID in the ByRef variables
    ' Returns 0 if success, 1 if failure
    GetElement = GetElementInternal(intElementID, strSymbol, dblMass, dblUncertainty, sngCharge, intIsotopeCount)
End Function

Public Function GetElementCount() As Long
    ' Returns the number of elements in memory
    GetElementCount = GetElementCountInternal()
End Function

Public Function GetElementID(ByVal strSymbol As String) As Long
    ' Returns 0 if not found, the ID if found
    GetElementID = GetElementIDInternal(strSymbol)
End Function

Public Function GetElementIsotopes(ByVal intElementID As Integer, ByRef intIsotopeCount As Integer, ByRef dblIsotopeMasses() As Double, ByRef sngIsotopeAbundances() As Single) As Long
    ' Returns the Isotope masses for the element with intElementID
    GetElementIsotopes = GetElementIsotopesInternal(intElementID, intIsotopeCount, dblIsotopeMasses(), sngIsotopeAbundances())
End Function

Public Function GetElementMode() As emElementModeConstants
    ' Returns the element mode:
    '    emAverageMass  = 1
    '    emIsotopicMass = 2
    '    emIntegerMass  = 3
    GetElementMode = GetElementModeInternal()
End Function

Public Function GetElementSymbol(ByVal intElementID As Integer) As String
    ' Returns the symbol for the given element ID
    GetElementSymbol = GetElementSymbolInternal(intElementID)
End Function

Public Function GetElementStat(ByVal intElementID As Integer, ByVal eElementStat As esElementStatsConstants) As Double
    ' Returns a single bit of information about a single element
    GetElementStat = GetElementStatInternal(intElementID, eElementStat)
End Function

Public Function GetMessageStatement(lngMessageID As Long, Optional strAppendText As String = "") As String
    ' Returns the message for lngMessageID
    GetMessageStatement = GetMessageStatementInternal(lngMessageID, strAppendText)
End Function

Public Function GetMessageStatementCount() As Long
    GetMessageStatementCount = GetMessageStatementCountInternal()
End Function

Public Function IsModSymbol(strSymbol As String) As Boolean
    ' Returns True if strSymbol starts with a ModSymbol
    IsModSymbol = IsModSymbolInternal(strSymbol)
End Function

Private Sub LoadDefaults()
    MemoryLoadAll emAverageMass
    
    Me.SetElementMode emAverageMass
    Me.AbbreviationRecognitionMode = arNormalPlusAminoAcids
    Me.BracketsTreatedAsParentheses = True
    Me.CaseConversionMode = ccConvertCaseUp
    Me.DecimalSeparator = "."
    Me.RtfFontName = "Arial"
    Me.RtfFontSize = 10
    Me.StdDevMode = smDecimal
    
    gComputationOptions.DecimalSeparator = DetermineDecimalPoint()
    
    mDataInitialized = True
End Sub

Public Sub RemoveAllAbbreviations()
    RemoveAllAbbreviationsInternal
End Sub

Public Sub RemoveAllCautionStatements()
    RemoveAllCautionStatementsInternal
End Sub

Public Function MassToPPM(ByVal dblMassToConvert As Double, ByVal dblCurrentMZ As Double) As Double
    MassToPPM = MassToPPMInternal(dblMassToConvert, dblCurrentMZ)
End Function

Public Function MonoMassToMZ(ByVal dblMonoisotopicMass As Double, ByVal intCharge As Integer, Optional ByVal dblChargeCarrierMass As Double = 0) As Double
    MonoMassToMZ = MonoMassToMZInternal(dblMonoisotopicMass, intCharge, dblChargeCarrierMass)
End Function

Public Sub RecomputeAbbreviationMasses()
    ' Use this sub to manually recompute the masses of the abbreviations
    ' Useful if we just finished setting lots of element masses, and
    '  had blnRecomputeAbbreviationMasses = False when calling .SetElement()
    RecomputeAbbreviationMassesInternal
End Sub

Public Function RemoveAbbreviation(ByVal strAbbreviationSymbol As String) As Long
    RemoveAbbreviation = RemoveAbbreviationInternal(strAbbreviationSymbol)
End Function

Public Function RemoveAbbreviationByID(ByVal lngAbbreviationID As Long) As Long
    RemoveAbbreviationByID = RemoveAbbreviationByIDInternal(lngAbbreviationID)
End Function

Public Function RemoveCautionStatement(ByVal strCautionSymbol As String) As Long
    RemoveCautionStatement = RemoveCautionStatementInternal(strCautionSymbol)
End Function

Public Sub ResetAbbreviations()
    MemoryLoadAbbreviations
End Sub

Public Sub ResetCautionStatements()
    MemoryLoadCautionStatements
End Sub

Public Sub ResetElement(intElementID As Integer, eSpecificStatToReset As esElementStatsConstants)
    MemoryLoadElements GetElementMode(), intElementID, eSpecificStatToReset
End Sub

Public Sub ResetMessageStatements()
    MemoryLoadMessageStatements
End Sub

Public Function SetAbbreviation(strSymbol As String, strFormula As String, sngCharge As Single, blnIsAminoAcid As Boolean, Optional strOneLetterSymbol As String = "", Optional strComment As String = "", Optional blnValidateFormula As Boolean = True) As Long
    ' Adds a new abbreviation or updates an existing one (based on strSymbol)
    ' If blnValidateFormula = True, then makes sure the formula is valid
    ' It is useful to set blnValidateFormula = False when you're defining all of the abbreviations at once,
    '  since one abbreviation can depend upon another, and if the second abbreviation hasn't yet been
    '  defined, then the parsing of the first abbreviation will fail
    ' Returns 0 if successful, otherwise, returns an Error ID
    SetAbbreviation = SetAbbreviationInternal(strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment, blnValidateFormula)
End Function

Public Function SetAbbreviationByID(lngAbbrevID As Long, strSymbol As String, strFormula As String, sngCharge As Single, blnIsAminoAcid As Boolean, Optional strOneLetterSymbol As String = "", Optional strComment As String = "", Optional blnValidateFormula As Boolean = True) As Long
    ' Adds a new abbreviation or updates an existing one (based on strSymbol)
    ' If blnValidateFormula = True, then makes sure the formula is valid
    ' It is useful to set blnValidateFormula = False when you're defining all of the abbreviations at once,
    '  since one abbreviation can depend upon another, and if the second abbreviation hasn't yet been
    '  defined, then the parsing of the first abbreviation will fail
    ' Returns 0 if successful, otherwise, returns an Error ID
    SetAbbreviationByID = SetAbbreviationByIDInternal(lngAbbrevID, strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment, blnValidateFormula)
End Function

Public Function SetCautionStatement(strNewSymbolCombo As String, strNewCautionStatement As String) As Long
    ' Adds a new caution statement or updates an existing one (based on strSymbol)
    ' Returns 0 if successful, otherwise, returns an Error ID
    SetCautionStatement = SetCautionStatementInternal(strNewSymbolCombo, strNewCautionStatement)
End Function

Public Function SetChargeCarrierMass(dblMass As Double)
    SetChargeCarrierMassInternal dblMass
End Function

Public Function SetElement(strSymbol As String, dblMass As Double, dblUncertainty As Double, sngCharge As Single, Optional blnRecomputeAbbreviationMasses As Boolean = True) As Long
    ' Used to update the values for a single element (based on strSymbol)
    SetElement = SetElementInternal(strSymbol, dblMass, dblUncertainty, sngCharge, blnRecomputeAbbreviationMasses)
End Function

Public Function SetElementIsotopes(ByVal strSymbol As String, ByVal intIsotopeCount As Integer, ByRef dblIsotopeMassesOneBased() As Double, ByRef sngIsotopeAbundancesOneBased() As Single) As Long
    SetElementIsotopes = SetElementIsotopesInternal(strSymbol, intIsotopeCount, dblIsotopeMassesOneBased(), sngIsotopeAbundancesOneBased())
End Function

Public Sub SetElementMode(NewElementMode As emElementModeConstants, Optional blnMemoryLoadElementValues As Boolean = True)
    SetElementModeInternal NewElementMode, blnMemoryLoadElementValues
End Sub

Public Sub SetStandardProgressCaptionText(Optional ByVal strMinutesElapsedRemaining As String = "min. elapsed/remaining", Optional ByVal strPreparingToPause As String = "Preparing to Pause", Optional ByVal strResuming As String = "Resuming", Optional ByVal strClickToPause As String = "Click to Pause", Optional ByVal strPaused As String = "Paused", Optional ByVal strPressEscapeToAbort As String = "(Press Escape to abort)")
    frmProgress.SetStandardCaptionText strMinutesElapsedRemaining, strPreparingToPause, strResuming, strClickToPause, strPaused, strPressEscapeToAbort
End Sub

Public Function SetMessageStatement(lngMessageID As Long, strNewMessage As String) As Long
    ' Used to replace the default message strings with foreign language equivalent ones
    SetMessageStatement = SetMessageStatementInternal(lngMessageID, strNewMessage)
End Function

Public Sub SortAbbreviations()
    SortAbbreviationsInternal
End Sub

Public Function TextToRTF(strTextToConvert As String, Optional CalculatorMode As Boolean = False, Optional blnHighlightCharFollowingPercentSign As Boolean = True, Optional blnOverrideErrorID As Boolean = False, Optional lngErrorIDOverride As Long = 0) As String
    ' Converts an RTF string for the given text
    ' If blnHighlightCharFollowingPercentSign is true, then changes the character
    '  following a percent sign to red (and removes the percent sign)
        
    ' When blnCalculatorMode = True, then does not superscript + signs and numbers following + signs
    TextToRTF = PlainTextToRtfInternal(strTextToConvert, CalculatorMode, blnHighlightCharFollowingPercentSign, blnOverrideErrorID, lngErrorIDOverride)
End Function

Public Function ValidateAllAbbreviations() As Long
    ' Checks the formula of all abbreviations to make sure it's valid
    ' Marks any abbreviations as Invalid if a problem is found or a circular reference exists
    ' Returns a count of the number of invalid abbreviations found
    
    ValidateAllAbbreviations = ValidateAllAbbreviationsInternal()
End Function

Private Sub Class_Initialize()
    If Not mDataInitialized Then LoadDefaults
End Sub

Private Sub Class_Terminate()
    Set Peptide = Nothing
    Set CapFlow = Nothing
    Set Compound = Nothing
End Sub

Public Property Get AbbreviationRecognitionMode() As arAbbrevRecognitionModeConstants
    AbbreviationRecognitionMode = gComputationOptions.AbbrevRecognitionMode
End Property

Public Property Let AbbreviationRecognitionMode(eAbbrevRecognitionMode As arAbbrevRecognitionModeConstants)
    If eAbbrevRecognitionMode >= arNormalOnly And eAbbrevRecognitionMode <= arNoAbbreviations Then
        gComputationOptions.AbbrevRecognitionMode = eAbbrevRecognitionMode
        ConstructMasterSymbolsList
    End If
End Property

Public Property Get AppDate() As String
    AppDate = PROGRAM_DATE
End Property

Public Property Get AppVersion() As String
    AppVersion = PROGRAM_VERSION
End Property

Public Property Let BracketsTreatedAsParentheses(boolBracketSet As Boolean)
    gComputationOptions.BracketsAsParentheses = boolBracketSet
End Property

Public Property Get BracketsTreatedAsParentheses() As Boolean
    BracketsTreatedAsParentheses = gComputationOptions.BracketsAsParentheses
End Property

Public Property Let CaseConversionMode(NewCaseConversionMode As ccCaseConversionConstants)
    If NewCaseConversionMode >= ccConvertCaseUp And NewCaseConversionMode <= ccSmartCase Then
        gComputationOptions.CaseConversion = NewCaseConversionMode
    End If
End Property

Public Property Get CaseConversionMode() As ccCaseConversionConstants
    CaseConversionMode = gComputationOptions.CaseConversion
End Property

Public Property Let DecimalSeparator(strNewDecimalSeparator As String)
    If Len(strNewDecimalSeparator) = 1 Then
        gComputationOptions.DecimalSeparator = strNewDecimalSeparator
    End If
End Property

Public Property Get DecimalSeparator() As String
    DecimalSeparator = gComputationOptions.DecimalSeparator
End Property

Public Property Get ErrorDescription() As String
    ErrorDescription = GetErrorDescription()
End Property

Public Property Get ErrorID() As Long
    ErrorID = GetErrorID()
End Property

Public Property Get ErrorCharacter() As String
    ErrorCharacter = GetErrorCharacter()
End Property

Public Property Get ErrorPosition() As Long
    ErrorPosition = GetErrorPosition()
End Property

Public Property Let RtfFontName(strNewFontName As String)
    If Len(strNewFontName) > 0 Then
        gComputationOptions.RtfFontName = strNewFontName
    End If
End Property

Public Property Get RtfFontName() As String
    RtfFontName = gComputationOptions.RtfFontName
End Property

Public Property Let RtfFontSize(intNewFontSize As Integer)
    If intNewFontSize > 0 Then
        gComputationOptions.RtfFontSize = intNewFontSize
    End If
End Property

Public Property Get RtfFontSize() As Integer
    RtfFontSize = gComputationOptions.RtfFontSize
End Property

Public Property Get ShowErrorDialogs() As Boolean
    ShowErrorDialogs = ShowErrorMessageDialogs()
End Property

Public Property Let ShowErrorDialogs(blnValue As Boolean)
    SetShowErrorMessageDialogs blnValue
End Property

Public Property Let StdDevMode(NewStdDevMode As smStdDevModeConstants)
    If NewStdDevMode >= smShort And NewStdDevMode <= smDecimal Then
        gComputationOptions.StdDevMode = NewStdDevMode
    End If
End Property

Public Property Get StdDevMode() As smStdDevModeConstants
    StdDevMode = gComputationOptions.StdDevMode
End Property

