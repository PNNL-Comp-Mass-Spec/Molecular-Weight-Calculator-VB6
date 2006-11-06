Attribute VB_Name = "ElementAndMassRoutines"
Option Explicit

' Last Modified: November 22, 2003

Public Const ELEMENT_COUNT = 103
Public Const MAX_ABBREV_COUNT = 500

Private Const MESSAGE_STATEMENT_DIMCOUNT = 1600
Private Const MAX_ABBREV_LENGTH = 6
Private Const MAX_ISOTOPES = 11
Private Const MAX_CAUTION_STATEMENTS = 100

Private Const EMPTY_STRINGCHAR As String = "~"
Private Const RTF_HEIGHT_ADJUSTCHAR As String = "~"          ' A hidden character to adjust the height of Rtf Text Boxes when using superscripts
Private Const LowestValueForDoubleDataType = -1.79E+308      ' Use -3.4E+38 for Single
Private Const HighestValueForDoubleDataType = 1.79E+308      ' Use  3.4E+38 for Single

' Note that these Enums are duplicated in MolecularWeightCalculator.cls as Public versions
Public Enum emElementModeConstantsPrivate
    emAverageMass = 1
    emIsotopicMass = 2
    emIntegerMass = 3
End Enum

Public Enum smStdDevModeConstantsPrivate
    smShort = 0
    smScientific = 1
    smDecimal = 2
End Enum

Public Enum ccCaseConversionConstantsPrivate
    ccConvertCaseUp = 0
    ccExactCase = 1
    ccSmartCase = 2
End Enum

Private Enum smtSymbolMatchTypeConstants
    smtUnknown = 0
    smtElement = 1
    smtAbbreviation = 2
End Enum

Public Type udtOptionsType
    AbbrevRecognitionMode As arAbbrevRecognitionModeConstants
    BracketsAsParentheses As Boolean
    CaseConversion As ccCaseConversionConstantsPrivate
    DecimalSeparator As String
    RtfFontName As String
    RtfFontSize As Integer
    StdDevMode As smStdDevModeConstantsPrivate   ' Can be 0, 1, or 2 (see smStdDevModeConstants)
End Type

Public Type usrIsotopicAtomInfoType
    Count As Double                     ' Can have non-integer counts of atoms, eg. ^13C5.5
    Mass As Double
End Type

Public Type udtElementUseStatsType
    Used As Boolean
    Count As Double                        ' Can have non-integer counts of atoms, eg. C5.5
    IsotopicCorrection As Double
    IsotopeCount As Integer                ' Number of specific isotopes defined
    Isotopes() As usrIsotopicAtomInfoType
End Type

Public Type udtPctCompType
    PercentComposition As Double
    StdDeviation As Double
End Type

Public Type udtComputationStatsType
    Elements(ELEMENT_COUNT) As udtElementUseStatsType           ' 1-based array
    TotalMass As Double
    PercentCompositions(ELEMENT_COUNT) As udtPctCompType
    Charge As Single
    StandardDeviation As Double
End Type

Public Type udtIsotopeInfoType
    Mass As Double
    Abundance As Single
End Type

Public Type udtElementStatsType
    Symbol As String
    Mass As Double
    Uncertainty As Double
    Charge As Single
    IsotopeCount As Integer                         ' # of isotopes an element has
    Isotopes(MAX_ISOTOPES) As udtIsotopeInfoType        ' Masses and Abundances of the isotopes; 1-based array
End Type

Public Type udtAbbrevStatsType
    Symbol As String            ' The symbol for the abbreviation, e.g. Ph for the phenyl group or Ala for alanine (3 letter codes for amino acids)
    Formula As String           ' Cannot contain other abbreviations
    Mass As Double              ' Computed mass for quick reference
    Charge As Single
    IsAminoAcid As Boolean      ' True if an amino acid
    OneLetterSymbol As String   ' Only used for amino acids
    Comment As String           ' Description of the abbreviation
    InvalidSymbolOrFormula As Boolean
End Type

Private Type udtErrorDescriptionType
    ErrorID As Long                  ' Contains the error number (used in the LookupMessage function).  In addition, if a program error occurs, ErrorParams.ErrorID = -10
    ErrorPosition As Integer
    ErrorCharacter As String
End Type

Private Type udtIsoResultsByElementType
    ElementIndex As Integer             ' Index of element in ElementStats() array; look in ElementStats() to get information on its isotopes
    boolExplicitIsotope As Boolean      ' True if an explicitly defined isotope
    ExplicitMass As Double
    AtomCount As Long               ' Number of atoms of this element in the formula being parsed
    ResultsCount As Long            ' Number of masses in MassAbundances
    StartingResultsMass As Long     ' Starting mass of the results for this element
    MassAbundances() As Single      ' Abundance of each mass, starting with StartingResultsMass
End Type

Private Type udtIsoResultsOverallType
    Abundance As Single
    Multiplicity As Long
End Type

Private Type udtAbbrevSymbolStackType
    Count As Integer
    SymbolReferenceStack() As Integer       ' 0-based array
End Type

Public gComputationOptions As udtOptionsType

Private ElementAlph(ELEMENT_COUNT) As String               ' Stores the elements in alphabetical order; used for constructing empirical formulas; 1 to ELEMENT_COUNT
Private ElementStats(ELEMENT_COUNT) As udtElementStatsType      ' 1 to ELEMENT_COUNT

' No number for array size since I dynamically allocate memory for it
Private MasterSymbolsList() As String            ' Stores the element symbols, abbreviations, & amino acids in order of longest symbol length to shortest length, non-alphabatized, for use in symbol matching when parsing a formula; 1 To MasterSymbolsListcount
Private MasterSymbolsListCount As Integer

Private AbbrevStats(MAX_ABBREV_COUNT) As udtAbbrevStatsType         ' Includes both abbreviations and amino acids; 1-based array
Private AbbrevAllCount As Integer

Private CautionStatements(MAX_CAUTION_STATEMENTS, 2) As String       ' CautionStatements(x,0) holds the symbol combo to look for, CautionStatements(x, 1) holds the caution statement; 1-based array
Private CautionStatementCount As Long

Private MessageStatements(MESSAGE_STATEMENT_DIMCOUNT) As String                  ' Holds error messages; 1-based array
Private MessageStatmentCount As Long

Private ErrorParams As udtErrorDescriptionType

Private mChargeCarrierMass As Double                ' 1.00727649 for monoisotopic mass or 1.00739 for average mass

Private mCurrentElementMode As emElementModeConstantsPrivate
Private mStrCautionDescription As String
Private mComputationStatsSaved As udtComputationStatsType

Private mShowErrorMessageDialogs As Boolean

Private Sub AbbrevSymbolStackAdd(udtAbbrevSymbolStack As udtAbbrevSymbolStackType, SymbolReference As Integer)

On Error GoTo AbbrevSymbolStackAddErrorHandler

    With udtAbbrevSymbolStack
        .Count = .Count + 1
        ReDim Preserve .SymbolReferenceStack(.Count - 1)
        .SymbolReferenceStack(.Count - 1) = SymbolReference
    End With

    Exit Sub
    
AbbrevSymbolStackAddErrorHandler:
    Debug.Assert False

End Sub

Private Sub AbbrevSymbolStackAddRemoveMostRecent(udtAbbrevSymbolStack As udtAbbrevSymbolStackType)
    With udtAbbrevSymbolStack
        If .Count > 0 Then
            .Count = .Count - 1
        End If
    End With
End Sub

Private Sub AddAbbreviationWork(intAbbrevIndex As Integer, strSymbol As String, strFormula As String, sngCharge As Single, blnIsAminoAcid As Boolean, Optional strOneLetter As String = "", Optional strComment As String = "", Optional blnInvalidSymbolOrFormula As Boolean = False)
    With AbbrevStats(intAbbrevIndex)
        .InvalidSymbolOrFormula = blnInvalidSymbolOrFormula
        .Symbol = strSymbol
        .Formula = strFormula
        .Mass = ComputeFormulaWeight(strFormula)
        If .Mass < 0 Then
            ' Error occurred computing mass for abbreviation
            .Mass = 0
            .InvalidSymbolOrFormula = True
        End If
        .Charge = sngCharge
        .OneLetterSymbol = UCase(strOneLetter)
        .IsAminoAcid = blnIsAminoAcid
        .Comment = strComment
    End With
End Sub

Private Sub AddToCautionDescription(strTextToAdd As String)
    If Len(mStrCautionDescription) > 0 Then
        mStrCautionDescription = mStrCautionDescription
    End If
    mStrCautionDescription = mStrCautionDescription & strTextToAdd
End Sub

Private Sub CheckCaution(strFormulaExcerpt As String)
    Dim strTest As String
    Dim strNewCaution As String
    Dim intIndex As Integer
    
    For intIndex = 1 To MAX_ABBREV_LENGTH
        strTest = Left(strFormulaExcerpt, intIndex)
        strNewCaution = LookupCautionStatement(strTest)
        If Len(strNewCaution) > 0 Then
            AddToCautionDescription strNewCaution
            Exit For
        End If
    Next intIndex

End Sub

Private Sub CatchParsenumError(AdjacentNum As Double, numsizing As Integer, curcharacter As Integer, symbollength As Integer)

    If AdjacentNum < 0 And numsizing = 0 Then
        Select Case AdjacentNum
        Case -1
            ' No number, but no error
            ' That's ok
        Case -3
            ' Error: No number after decimal point
            ErrorParams.ErrorID = 12: ErrorParams.ErrorPosition = curcharacter + symbollength
        Case -4
            ' Error: More than one decimal point
            ErrorParams.ErrorID = 27: ErrorParams.ErrorPosition = curcharacter + symbollength
        Case Else
            ' Error: General number error
            ErrorParams.ErrorID = 14: ErrorParams.ErrorPosition = curcharacter + symbollength
        End Select
    End If
    
End Sub

Private Function CheckElemAndAbbrev(strFormulaExcerpt As String, ByRef SymbolReference As Integer) As smtSymbolMatchTypeConstants
    ' Returns smtElement if matched an element
    ' Returns smtAbbreviation if matched an abbreviation or amino acid
    ' Returns smtUnknown if no match
    
    ' SymbolReference is the index of the matched element or abbreviation in MasterSymbolsList()
    
    Dim intIndex As Integer
    Dim eSymbolMatchType As smtSymbolMatchTypeConstants
    
    ' MasterSymbolsList() stores the element symbols, abbreviations, & amino acids in order of longest length to
    '   shortest length, non-alphabatized, for use in symbol matching when parsing a formula

    ' MasterSymbolsList(intIndex,0) contains the symbol to be matched
    ' MasterSymbolsList(intIndex,1) contains E for element, A for amino acid, or N for normal abbreviation, followed by
    '     the reference number in the master list
    ' For example for Carbon, MasterSymbolsList(intIndex,0) = "C" and MasterSymbolsList(intIndex,1) = "E6"

    ' Look for match, stepping directly through MasterSymbolsList()
    ' List is sorted by reverse length, so can do all at once
    
    For intIndex = 1 To MasterSymbolsListCount
        If Len(MasterSymbolsList(intIndex, 0)) > 0 Then
            If Left(strFormulaExcerpt, Len(MasterSymbolsList(intIndex, 0))) = MasterSymbolsList(intIndex, 0) Then
                ' Matched a symbol
                Select Case UCase(Left(MasterSymbolsList(intIndex, 1), 1))
                Case "E"    ' An element
                    eSymbolMatchType = smtElement
                Case "A"    ' An abbreviation or amino acid
                    eSymbolMatchType = smtAbbreviation
                Case Else
                    ' error
                    eSymbolMatchType = smtUnknown
                    SymbolReference = -1
                End Select
                If eSymbolMatchType <> smtUnknown Then
                    SymbolReference = Val(Mid(MasterSymbolsList(intIndex, 1), 2))
                End If
                Exit For
            End If
        Else
            Debug.Assert False
        End If
    Next intIndex
    
    CheckElemAndAbbrev = eSymbolMatchType

End Function

Public Function ComputeFormulaWeight(strFormula As String) As Double
    ' Simply returns the weight of a formula (or abbreviation)
    ' Returns -1 if an error occurs
    ' Error information is stored in ErrorParams
    
    Dim udtComputationStats As udtComputationStatsType
    Dim dblMass As Double

    dblMass = ParseFormulaPublic(strFormula, udtComputationStats, False)
    
    If ErrorParams.ErrorID = 0 Then
        ComputeFormulaWeight = udtComputationStats.TotalMass
    Else
        ComputeFormulaWeight = -1
    End If
End Function

Public Function TestComputeIsotopicAbundances(ByRef strFormulaIn As String, ByVal intChargeState As Integer)

    Dim strResults As String
    
    Dim ConvolutedMSData2DOneBased() As Double
    Dim ConvolutedMSDataCount As Long
    
    Dim intReturn As Integer
    Static ProgramInitialized As Boolean
    
    If Not ProgramInitialized Then
        MemoryLoadAll emElementModeConstantsPrivate.emIsotopicMass
        ProgramInitialized = True
    End If
    
    intReturn = ComputeIsotopicAbundancesInternal(strFormulaIn, intChargeState, strResults, ConvolutedMSData2DOneBased, ConvolutedMSDataCount)
    
    Debug.Print strResults
    
End Function

Public Function ComputeIsotopicAbundancesInternal(ByRef strFormulaIn As String, ByVal intChargeState As Integer, ByRef strResults As String, ByRef ConvolutedMSData2DOneBased() As Double, ByRef ConvolutedMSDataCount As Long, Optional ByVal strHeaderIsotopicAbundances As String = "Isotopic Abundances for", Optional ByVal strHeaderMassToCharge As String = "Mass/Charge", Optional ByVal strHeaderFraction As String = "Fraction", Optional ByVal strHeaderIntensity As String = "Intensity", Optional blnUseFactorials As Boolean = False) As Integer
    ' Computes the Isotopic Distribution for a formula, returns uncharged mass vlaues if intChargeState=0,
    '  M+H values if intChargeState=1, and convoluted m/z if intChargeState is > 1
    ' Updates strFormulaIn to the properly formatted formula
    ' Returns the results in strResults
    ' Returns 0 if success, or -1 if an error
    
    Dim strFormula As String, strModifiedFormula As String
    Dim dblWorkingFormulaMass As Double
    Dim dblExactBaseIsoMass As Double, dblMassDefect As Double, dblMaxPercentDifference As Double
    Dim intElementIndex As Integer, intElementCount As Integer
    Dim lngMassIndex As Long, lngRowIndex As Long
    Dim udtComputationStats As udtComputationStatsType
    Dim dblTemp As Double
    
    Dim IsoStats() As udtIsoResultsByElementType
    
    Dim IsotopeCount As Integer, IsotopeStartingMass As Integer, IsotopeEndingMass As Integer
    Dim MasterElementIndex As Integer, AtomCount As Long, dblCount As Double
    Dim PredictedCombos As Long, CombosFound As Long
    Dim ComboIndex As Long, IsotopeIndex As Integer, intIndex As Integer
    Dim IndexToStoreAbundance As Long
    Dim dblThisComboFractionalAbundance As Double, dblNextComboFractionalAbundance As Double
    Dim blnRatioMethodUsed As Boolean, blnRigorousMethodUsed As Boolean
    
    Const strDeuteriumEquiv = "^2.014H"
    Dim blnReplaceDeuterium As Boolean, intAsciiOfNext As Integer

    Dim IsotopeCountInThisCombo As Long
    Dim strOutput As String
    
    Dim blnShowProgressForm As Boolean
    Dim PredictedConvIterations As Long
    Dim PredictedTotalComboCalcs As Long, CompletedComboCalcs As Long
    Const COMBO_ITERATIONS_SHOW_PROGRESS = 500
    Const MIN_ABUNDANCE_TO_KEEP = 0.000001
    Const CUTOFF_FOR_RATIO_METHOD = 0.00001
    
    ' AbundDenom  and  AbundSuffix are only needed if using the easily-overflowed factorial method
    Dim AbundDenom As Double, AbundSuffix As Double
    
    Dim AtomTrackHistory() As Long
    
    Dim IsoCombos() As Long     ' 2D array: Holds the # of each isotope for each combination
    ' For example, Two chlorines, Cl2, has at most 6 combos since Cl isotopes are 35, 36, and 37
    ' m1  m2  m3
    ' 2   0   0
    ' 1   1   0
    ' 1   0   1
    ' 0   2   0
    ' 0   1   1
    ' 0   0   2
    
    Dim ConvolutedAbundances() As udtIsoResultsOverallType     ' Fractional abundance at each mass; 1-based array
    Dim ConvolutedAbundanceStartMass As Long
    
    Dim MinWeight As Long, MaxWeight As Long, ResultingMassCountForElement As Long
    Dim blnExplicitIsotopesPresent As Boolean, ExplicitIsotopeCount As Integer
    
    Dim SubIndex As Long, lngSigma As Long
    Dim dblLogSigma  As Double, dblSumI As Double, dblSumF As Double
    Dim dblWorkingSum As Double
    Dim dblLogFreq As Double
    
    Dim dblFractionalAbundanceSaved As Double, dblLogRho As Double, dblRho As Double
    Dim lngM As Double, lngMPrime As Double
    Dim dblRatioOfFreqs As Double
    
    ' Make sure formula is not blank
    If Len(strFormulaIn) = 0 Then
        ComputeIsotopicAbundancesInternal = -1
        Exit Function
    End If

    On Error GoTo IsoAbundanceErrorHandler
    
    ' Change strHeaderMassToCharge to "Neutral Mass" if intChargeState = 0 and strHeaderMassToCharge is "Mass/Charge"
    If intChargeState = 0 Then
        If strHeaderMassToCharge = "Mass/Charge" Then
            strHeaderMassToCharge = "Neutral Mass"
        End If
    End If
    
    ' Parse Formula to determine if valid and number of each element
    strFormula = strFormulaIn
    dblWorkingFormulaMass = ParseFormulaPublic(strFormula, udtComputationStats, False)
    
    If dblWorkingFormulaMass < 0 Then
        ' Error occurred; information is stored in ErrorParams
        ComputeIsotopicAbundancesInternal = -1
        strResults = LookupMessage(350) & ": " & LookupMessage(ErrorParams.ErrorID)
        Exit Function
    End If
        
    ' See if Deuterium is present by looking for a fractional amount of Hydrogen
    ' strFormula will contain a capital D followed by a number or another letter (or the end of formula)
    ' If found, replace each D with ^2.014H and re-compute
    dblCount = udtComputationStats.Elements(1).Count
    If dblCount <> CLng(dblCount) Then
        ' Deuterium is present
        strModifiedFormula = ""
        intIndex = 1
        Do While intIndex <= Len(strFormula)
            blnReplaceDeuterium = False
            If Mid(strFormula, intIndex, 1) = "D" Then
                If intIndex = Len(strFormula) Then
                    blnReplaceDeuterium = True
                Else
                    intAsciiOfNext = Asc(Mid(strFormula, intIndex + 1, 1))
                    If intAsciiOfNext < 97 Or intAsciiOfNext > 122 Then
                        blnReplaceDeuterium = True
                    End If
                End If
                If blnReplaceDeuterium Then
                    If intIndex > 1 Then
                        strModifiedFormula = Left(strFormula, intIndex - 1)
                    End If
                    strModifiedFormula = strModifiedFormula & strDeuteriumEquiv
                    If intIndex < Len(strFormula) Then
                        strModifiedFormula = strModifiedFormula & Mid(strFormula, intIndex + 1)
                    End If
                    strFormula = strModifiedFormula
                    intIndex = 0
                End If
            End If
            intIndex = intIndex + 1
        Loop
        
        ' Re-Parse Formula since D's are now ^2.014H
        dblWorkingFormulaMass = ParseFormulaPublic(strFormula, udtComputationStats, False)
        
        If dblWorkingFormulaMass < 0 Then
            ' Error occurred; information is stored in ErrorParams
            ComputeIsotopicAbundancesInternal = -1
            strResults = LookupMessage(350) & ": " & LookupMessage(ErrorParams.ErrorID)
            Exit Function
        End If
    End If
    
    ' Make sure there are no fractional atoms present (need to specially handle Deuterium)
    For intElementIndex = 1 To ELEMENT_COUNT
        dblCount = udtComputationStats.Elements(intElementIndex).Count
        If dblCount <> CLng(dblCount) Then
            ComputeIsotopicAbundancesInternal = -1
            strResults = LookupMessage(350) & ": " & LookupMessage(805) & ": " & ElementStats(intElementIndex).Symbol & dblCount
            Exit Function
        End If
    Next intElementIndex
    
    ' Remove occurrences of explicitly defined isotopes from the formula
    For intElementIndex = 1 To ELEMENT_COUNT
        With udtComputationStats.Elements(intElementIndex)
            If .IsotopeCount > 0 Then
                blnExplicitIsotopesPresent = True
                ExplicitIsotopeCount = ExplicitIsotopeCount + .IsotopeCount
                For IsotopeIndex = 1 To .IsotopeCount
                    .Count = .Count - .Isotopes(IsotopeIndex).Count
                Next IsotopeIndex
            End If
        End With
    Next intElementIndex
    
    ' Determine the number of elements present in strFormula
    intElementCount = 0
    For intElementIndex = 1 To ELEMENT_COUNT
        If udtComputationStats.Elements(intElementIndex).Used Then
            intElementCount = intElementCount + 1
        End If
    Next intElementIndex
    
    If blnExplicitIsotopesPresent Then
        intElementCount = intElementCount + ExplicitIsotopeCount
    End If
    
    If intElementCount = 0 Or dblWorkingFormulaMass = 0 Then
        ' No elements or no weight
        ComputeIsotopicAbundancesInternal = -1
        Exit Function
    End If
    
    ' The formula seems valid, so update strFormulaIn
    strFormulaIn = strFormula
    
    ' Reserve memory for IsoStats() array
    ReDim IsoStats(intElementCount)

    ' Step through udtComputationStats.Elements() again and copy info into IsoStats()
    ' In addition, determine minimum and maximum weight for the molecule
    intElementCount = 0
    MinWeight = 0
    MaxWeight = 0
    For intElementIndex = 1 To ELEMENT_COUNT
        If udtComputationStats.Elements(intElementIndex).Used Then
            If udtComputationStats.Elements(intElementIndex).Count > 0 Then
                intElementCount = intElementCount + 1
                IsoStats(intElementCount).ElementIndex = intElementIndex
                IsoStats(intElementCount).AtomCount = udtComputationStats.Elements(intElementIndex).Count   ' Note: Ignoring .Elements(intElementIndex).IsotopicCorrection
                IsoStats(intElementCount).ExplicitMass = ElementStats(intElementIndex).Mass
                
                With ElementStats(intElementIndex)
                    MinWeight = MinWeight + IsoStats(intElementCount).AtomCount * Round(.Isotopes(1).Mass, 0)
                    MaxWeight = MaxWeight + IsoStats(intElementCount).AtomCount * Round(.Isotopes(.IsotopeCount).Mass, 0)
                End With
            End If
        End If
    Next intElementIndex
    
    If blnExplicitIsotopesPresent Then
        ' Add the isotopes, pretending they are unique elements
        For intElementIndex = 1 To ELEMENT_COUNT
            With udtComputationStats.Elements(intElementIndex)
                If .IsotopeCount > 0 Then
                    For IsotopeIndex = 1 To .IsotopeCount
                        intElementCount = intElementCount + 1
                        
                        IsoStats(intElementCount).boolExplicitIsotope = True
                        IsoStats(intElementCount).ElementIndex = intElementIndex
                        IsoStats(intElementCount).AtomCount = .Isotopes(IsotopeIndex).Count
                        IsoStats(intElementCount).ExplicitMass = .Isotopes(IsotopeIndex).Mass
                        
                        With IsoStats(intElementCount)
                            MinWeight = MinWeight + .AtomCount * .ExplicitMass
                            MaxWeight = MaxWeight + .AtomCount * .ExplicitMass
                        End With
                        
                    Next IsotopeIndex
                End If
            End With
        Next intElementIndex
    End If
    
    If MinWeight < 0 Then MinWeight = 0
    
    ' Create an array to hold the Fractional Abundances for all the masses
    ConvolutedMSDataCount = MaxWeight - MinWeight + 1
    ConvolutedAbundanceStartMass = MinWeight
    ReDim ConvolutedAbundances(ConvolutedMSDataCount)
    
    ' Predict the total number of computations required; show progress if necessary
    PredictedTotalComboCalcs = 0
    For intElementIndex = 1 To intElementCount
        MasterElementIndex = IsoStats(intElementIndex).ElementIndex
        AtomCount = IsoStats(intElementIndex).AtomCount
        IsotopeCount = ElementStats(MasterElementIndex).IsotopeCount
        
        PredictedCombos = FindCombosPredictIterations(AtomCount, IsotopeCount)
        PredictedTotalComboCalcs = PredictedTotalComboCalcs + PredictedCombos
    Next intElementIndex
         
    If PredictedTotalComboCalcs > COMBO_ITERATIONS_SHOW_PROGRESS Then
        blnShowProgressForm = True
        frmProgress.InitializeForm "Finding Isotopic Abundances", 0, PredictedTotalComboCalcs, False
        frmProgress.UpdateCurrentSubTask "Computing abundances"
    End If
    
    ' For each element, compute all of the possible combinations
    CompletedComboCalcs = 0
    For intElementIndex = 1 To intElementCount
       
        MasterElementIndex = IsoStats(intElementIndex).ElementIndex
        AtomCount = IsoStats(intElementIndex).AtomCount
        
        If IsoStats(intElementIndex).boolExplicitIsotope Then
            IsotopeCount = 1
            IsotopeStartingMass = IsoStats(intElementIndex).ExplicitMass
            IsotopeEndingMass = IsotopeStartingMass
        Else
            With ElementStats(MasterElementIndex)
                IsotopeCount = .IsotopeCount
                IsotopeStartingMass = Round(.Isotopes(1).Mass, 0)
                IsotopeEndingMass = Round(.Isotopes(IsotopeCount).Mass, 0)
            End With
        End If
        
        PredictedCombos = FindCombosPredictIterations(AtomCount, IsotopeCount)
        
        If PredictedCombos > 10000000 Then
            MsgBox "Too many combinations necessary for prediction of isotopic distribution: " & Format(PredictedCombos, "#,##0") & vbCrLf & "Please use a simpler formula or reduce the isotopic range defined for the element (currently " & IsotopeCount & ")"
            ComputeIsotopicAbundancesInternal = -1
            Exit Function
        End If
        
        ReDim IsoCombos(PredictedCombos, IsotopeCount)
        
        ReDim AtomTrackHistory(IsotopeCount)
        AtomTrackHistory(1) = AtomCount
        
        CombosFound = FindCombosRecurse(IsoCombos(), AtomCount, IsotopeCount, IsotopeCount, 1, 1, AtomTrackHistory())
        
        ' The predicted value should always match the actual value, unless blnExplicitIsotopesPresent = True
        If Not blnExplicitIsotopesPresent Then
            Debug.Assert PredictedCombos = CombosFound
        End If
        
        ' Reserve space for the abundances based on the minimum and maximum weight of the isotopes of the element
        MinWeight = AtomCount * IsotopeStartingMass
        MaxWeight = AtomCount * IsotopeEndingMass
        ResultingMassCountForElement = MaxWeight - MinWeight + 1
        IsoStats(intElementIndex).StartingResultsMass = MinWeight
        IsoStats(intElementIndex).ResultsCount = ResultingMassCountForElement
        ReDim IsoStats(intElementIndex).MassAbundances(ResultingMassCountForElement)
        
        If IsoStats(intElementIndex).boolExplicitIsotope Then
            ' Explicitly defined isotope; there is only one "combo" and its abundance = 1
            IsoStats(intElementIndex).MassAbundances(1) = 1
        Else
            dblFractionalAbundanceSaved = 0
            For ComboIndex = 1 To CombosFound
                CompletedComboCalcs = CompletedComboCalcs + 1
                If blnShowProgressForm Then
                    If CompletedComboCalcs Mod 10 = 0 Then
                        frmProgress.UpdateProgressBar CompletedComboCalcs
                        If KeyPressAbortProcess > 1 Then Exit For
                    End If
                End If
                
                dblThisComboFractionalAbundance = -1
                blnRatioMethodUsed = False
                blnRigorousMethodUsed = False
                
                If blnUseFactorials Then
                    ' #######
                    ' Rigorous, slow, easily overflowed method
                    ' #######
                    '
                    blnRigorousMethodUsed = True
                    
                    AbundDenom = 1
                    AbundSuffix = 1
                    With ElementStats(MasterElementIndex)
                        For IsotopeIndex = 1 To IsotopeCount
                            IsotopeCountInThisCombo = IsoCombos(ComboIndex, IsotopeIndex)
                            If IsotopeCountInThisCombo > 0 Then
                                AbundDenom = AbundDenom * Factorial(CDbl(IsotopeCountInThisCombo))
                                AbundSuffix = AbundSuffix * .Isotopes(IsotopeIndex).Abundance ^ IsotopeCountInThisCombo
                            End If
                        Next IsotopeIndex
                    End With
                    
                    dblThisComboFractionalAbundance = Factorial(CDbl(AtomCount)) / (AbundDenom) * AbundSuffix
                Else
                    If dblFractionalAbundanceSaved < CUTOFF_FOR_RATIO_METHOD Then
                        ' #######
                        ' Equivalent of rigorous method, but uses logarithms
                        ' #######
                        '
                        blnRigorousMethodUsed = True
                        
                        dblLogSigma = 0
                        For lngSigma = 1 To AtomCount
                            dblLogSigma = dblLogSigma + Log(CDbl(lngSigma))
                        Next lngSigma
                        
                        dblSumI = 0
                        For IsotopeIndex = 1 To IsotopeCount
                            If IsoCombos(ComboIndex, IsotopeIndex) > 0 Then
                                dblWorkingSum = 0
                                For SubIndex = 1 To IsoCombos(ComboIndex, IsotopeIndex)
                                        dblWorkingSum = dblWorkingSum + Log(CDbl(SubIndex))
                                Next SubIndex
                                dblSumI = dblSumI + dblWorkingSum
                            End If
                        Next IsotopeIndex
                        
                        With ElementStats(MasterElementIndex)
                            dblSumF = 0
                            For IsotopeIndex = 1 To IsotopeCount
                                If .Isotopes(IsotopeIndex).Abundance > 0 Then
                                    dblSumF = dblSumF + IsoCombos(ComboIndex, IsotopeIndex) * Log(CDbl(.Isotopes(IsotopeIndex).Abundance))
                                End If
                            Next IsotopeIndex
                        End With
                        
                        dblLogFreq = dblLogSigma - dblSumI + dblSumF
                        dblThisComboFractionalAbundance = Exp(dblLogFreq)
                    
                        dblFractionalAbundanceSaved = dblThisComboFractionalAbundance
                    End If
                    
                    ' Use dblThisComboFractionalAbundance to predict
                    ' the Fractional Abundance of the Next Combo
                    If ComboIndex < CombosFound And dblFractionalAbundanceSaved >= CUTOFF_FOR_RATIO_METHOD Then
                        ' #######
                        ' Third method, determines the ratio of this combo's abundance and the next combo's abundance
                        ' #######
                        '
                        dblRatioOfFreqs = 1
                        
                        For IsotopeIndex = 1 To IsotopeCount
                            lngM = IsoCombos(ComboIndex, IsotopeIndex)
                            lngMPrime = IsoCombos(ComboIndex + 1, IsotopeIndex)
                            
                            If lngM > lngMPrime Then
                                dblLogSigma = 0
                                For SubIndex = lngMPrime + 1 To lngM
                                    dblLogSigma = dblLogSigma + Log(SubIndex)
                                Next SubIndex
                                
                                With ElementStats(MasterElementIndex)
                                    dblLogRho = dblLogSigma - (lngM - lngMPrime) * Log(CDbl(.Isotopes(IsotopeIndex).Abundance))
                                End With
                            ElseIf lngM < lngMPrime Then
                                dblLogSigma = 0
                                For SubIndex = lngM + 1 To lngMPrime
                                    dblLogSigma = dblLogSigma + Log(SubIndex)
                                Next SubIndex
                                
                                With ElementStats(MasterElementIndex)
                                    If .Isotopes(IsotopeIndex).Abundance > 0 Then
                                        dblLogRho = (lngMPrime - lngM) * Log(CDbl(.Isotopes(IsotopeIndex).Abundance)) - dblLogSigma
                                    End If
                                End With
                            Else
                                ' lngM = lngMPrime
                                dblLogRho = 0
                            End If
                            
                            dblRho = Exp(dblLogRho)
                            dblRatioOfFreqs = dblRatioOfFreqs * dblRho
                            
                        Next IsotopeIndex
                        
                        dblNextComboFractionalAbundance = dblFractionalAbundanceSaved * dblRatioOfFreqs
                        
                        dblFractionalAbundanceSaved = dblNextComboFractionalAbundance
                        blnRatioMethodUsed = True
                    End If
                End If
                
                If blnRigorousMethodUsed Then
                    ' Determine nominal mass of this combination; depends on number of atoms of each isotope
                    IndexToStoreAbundance = FindIndexForNominalMass(IsoCombos(), ComboIndex, IsotopeCount, AtomCount, ElementStats(MasterElementIndex).Isotopes())
                
                    ' Store the abundance in .MassAbundances() at location IndexToStoreAbundance
                    With IsoStats(intElementIndex)
                        .MassAbundances(IndexToStoreAbundance) = .MassAbundances(IndexToStoreAbundance) + dblThisComboFractionalAbundance
                    End With
                End If
                
                If blnRatioMethodUsed Then
                    ' Store abundance for next Combo
                    IndexToStoreAbundance = FindIndexForNominalMass(IsoCombos(), ComboIndex + 1, IsotopeCount, AtomCount, ElementStats(MasterElementIndex).Isotopes())
                    
                    ' Store the abundance in .MassAbundances() at location IndexToStoreAbundance
                    With IsoStats(intElementIndex)
                        .MassAbundances(IndexToStoreAbundance) = .MassAbundances(IndexToStoreAbundance) + dblNextComboFractionalAbundance
                    End With
                End If
                
                If blnRatioMethodUsed And ComboIndex + 1 = CombosFound Then
                    ' No need to compute the last combo since we just did it
                    Exit For
                End If
    
            Next ComboIndex
        End If
        
        If KeyPressAbortProcess > 1 Then Exit For
   
    Next intElementIndex
    
    If KeyPressAbortProcess > 1 Then
        ' Process Aborted
        strResults = LookupMessage(940)
        If blnShowProgressForm Then frmProgress.HideForm
        ComputeIsotopicAbundancesInternal = -1
        Exit Function
    End If
    
    ' Step Through IsoStats from the end to the beginning, shortening the length to the
    ' first value greater than MIN_ABUNDANCE_TO_KEEP
    ' This greatly speeds up the convolution
    For intElementIndex = 1 To intElementCount
        With IsoStats(intElementIndex)
            lngRowIndex = .ResultsCount
            Do While .MassAbundances(lngRowIndex) < MIN_ABUNDANCE_TO_KEEP
                lngRowIndex = lngRowIndex - 1
                If lngRowIndex = 1 Then Exit Do
            Loop
            .ResultsCount = lngRowIndex
        End With
    Next intElementIndex
    
    ' Examine IsoStats() to predict the number of ConvolutionIterations
    PredictedConvIterations = IsoStats(1).ResultsCount
    For intElementIndex = 2 To intElementCount
        PredictedConvIterations = PredictedConvIterations * IsoStats(2).ResultsCount
    Next intElementIndex
    
    If blnShowProgressForm Then
        frmProgress.InitializeForm "Finding Isotopic Abundances", 0, IsoStats(1).ResultsCount, False
        frmProgress.UpdateCurrentSubTask "Convoluting results"
    End If
    
    ' Convolute the results for each element using a recursive convolution routine
    Dim ConvolutionIterations As Long
    ConvolutionIterations = 0
    For lngRowIndex = 1 To IsoStats(1).ResultsCount
        If blnShowProgressForm Then
            frmProgress.UpdateProgressBar lngRowIndex - 1
            If KeyPressAbortProcess > 1 Then Exit For
        End If
        ConvoluteMasses ConvolutedAbundances(), ConvolutedAbundanceStartMass, lngRowIndex, 1, 0, 1, IsoStats(), intElementCount, ConvolutionIterations
    Next lngRowIndex
    
    If KeyPressAbortProcess > 1 Then
        ' Process Aborted
        strResults = LookupMessage(940)
        If blnShowProgressForm Then frmProgress.HideForm
        ComputeIsotopicAbundancesInternal = -1
        Exit Function
    End If
    
    ' Compute mass defect (difference of initial isotope's mass from integer mass)
    dblExactBaseIsoMass = 0
    For intElementIndex = 1 To intElementCount
        With IsoStats(intElementIndex)
            If .boolExplicitIsotope Then
                dblExactBaseIsoMass = dblExactBaseIsoMass + .AtomCount * .ExplicitMass
            Else
                dblExactBaseIsoMass = dblExactBaseIsoMass + .AtomCount * ElementStats(.ElementIndex).Isotopes(1).Mass
            End If
        End With
    Next intElementIndex
    
    dblMassDefect = Round(dblExactBaseIsoMass - ConvolutedAbundanceStartMass, 5)
    
    ' Assure that the mass defect is only a small percentage of the total mass
    ' This won't be true for very small compounds so dblTemp is set to at least 10
    If ConvolutedAbundanceStartMass < 10 Then
        dblTemp = 10
    Else
        dblTemp = ConvolutedAbundanceStartMass
    End If
    dblMaxPercentDifference = 10 ^ -(3 - Round(Log10(CDbl(dblTemp)), 0))
    Debug.Assert Abs(dblMassDefect / dblExactBaseIsoMass) < dblMaxPercentDifference
    
    ' Step Through ConvolutedAbundances(), starting at the end, and find the first value above MIN_ABUNDANCE_TO_KEEP
    ' Decrease ConvolutedMSDataCount to remove the extra values below MIN_ABUNDANCE_TO_KEEP
    For lngMassIndex = ConvolutedMSDataCount To 1 Step -1
        If ConvolutedAbundances(lngMassIndex).Abundance > MIN_ABUNDANCE_TO_KEEP Then
            ConvolutedMSDataCount = lngMassIndex
            Exit For
        End If
    Next lngMassIndex
    
    strOutput = strHeaderIsotopicAbundances & " " & strFormula & vbCrLf
    strOutput = strOutput & SpacePad("  " & strHeaderMassToCharge, 12) & vbTab & SpacePad(strHeaderFraction, 9) & vbTab & strHeaderIntensity & vbCrLf
    
    ' Initialize ConvolutedMSData2DOneBased()
    ReDim ConvolutedMSData2DOneBased(ConvolutedMSDataCount, 2)
    
    ' Find Maximum Abundance
    Dim dblMaxAbundance As Double
    dblMaxAbundance = 0
    For lngMassIndex = 1 To ConvolutedMSDataCount
        If ConvolutedAbundances(lngMassIndex).Abundance > dblMaxAbundance Then
            dblMaxAbundance = ConvolutedAbundances(lngMassIndex).Abundance
        End If
    Next lngMassIndex
    
    ' Populate the results array with the masses and abundances
    ' Also, if intChargeState is >= 1, then convolute the mass to the appropriate m/z
    If dblMaxAbundance = 0 Then dblMaxAbundance = 1
    For lngMassIndex = 1 To ConvolutedMSDataCount
        With ConvolutedAbundances(lngMassIndex)
            ConvolutedMSData2DOneBased(lngMassIndex, 0) = (ConvolutedAbundanceStartMass + lngMassIndex - 1) + dblMassDefect
            ConvolutedMSData2DOneBased(lngMassIndex, 1) = .Abundance / dblMaxAbundance * 100
            
            If intChargeState >= 1 Then
                ConvolutedMSData2DOneBased(lngMassIndex, 0) = ConvoluteMassInternal(ConvolutedMSData2DOneBased(lngMassIndex, 0), 0, intChargeState)
            End If
        End With
    Next lngMassIndex
    
    ' Step through ConvolutedMSData2DOneBased() from the beginning to find the
    ' first value greater than MIN_ABUNDANCE_TO_KEEP
    lngRowIndex = 1
    Do While ConvolutedMSData2DOneBased(lngRowIndex, 1) < MIN_ABUNDANCE_TO_KEEP
        lngRowIndex = lngRowIndex + 1
        If lngRowIndex = ConvolutedMSDataCount Then Exit Do
    Loop

    ' If lngRowIndex > 1 then remove rows from beginning since value is less than MIN_ABUNDANCE_TO_KEEP
    If lngRowIndex > 1 And lngRowIndex < ConvolutedMSDataCount Then
        lngRowIndex = lngRowIndex - 1
        ' Remove rows from the start of ConvolutedMSData2DOneBased() since 0 mass
        For lngMassIndex = lngRowIndex + 1 To ConvolutedMSDataCount
            ConvolutedMSData2DOneBased(lngMassIndex - lngRowIndex, 0) = ConvolutedMSData2DOneBased(lngMassIndex, 0)
            ConvolutedMSData2DOneBased(lngMassIndex - lngRowIndex, 1) = ConvolutedMSData2DOneBased(lngMassIndex, 1)
        Next lngMassIndex
        ConvolutedMSDataCount = ConvolutedMSDataCount - lngRowIndex
    End If
    
    ' Write to strOutput
    For lngMassIndex = 1 To ConvolutedMSDataCount
        strOutput = strOutput & SpacePadFront(Format(ConvolutedMSData2DOneBased(lngMassIndex, 0), "#0.00000"), 12) & vbTab
        strOutput = strOutput & Format(ConvolutedMSData2DOneBased(lngMassIndex, 1) * dblMaxAbundance / 100, "0.0000000") & vbTab
        strOutput = strOutput & SpacePadFront(Format(ConvolutedMSData2DOneBased(lngMassIndex, 1), "##0.00"), 7) & vbCrLf
        ''ToDo: Fix Multiplicity
        '' strOutput = strOutput & Format(.Multiplicity, "##0") & vbCrLf
    Next lngMassIndex
    
    strResults = strOutput
    
    If blnShowProgressForm Then frmProgress.HideForm
    
    ComputeIsotopicAbundancesInternal = 0     ' Success
    
LeaveSub:
    Exit Function
    
IsoAbundanceErrorHandler:
    MwtWinDllErrorHandler "MwtWinDll|ComputeIsotopicAbundances"
    ErrorParams.ErrorID = 590
    ErrorParams.ErrorPosition = 0
    ComputeIsotopicAbundancesInternal = -1
    Resume LeaveSub

End Function

Public Sub ComputePercentComposition(udtComputationStats As udtComputationStatsType)
                
    Dim intElementIndex As Integer, dblElementTotalMass As Double
    Dim dblPercentComp As Double, dblStdDeviation As Double
    
    With udtComputationStats
        ' Determine the number of elements in the formula
        
        For intElementIndex = 1 To ELEMENT_COUNT
            If .TotalMass > 0 Then
                dblElementTotalMass = .Elements(intElementIndex).Count * ElementStats(intElementIndex).Mass + _
                                      .Elements(intElementIndex).IsotopicCorrection
                
                ' Percent is the percent composition
                dblPercentComp = dblElementTotalMass / .TotalMass * 100#
                .PercentCompositions(intElementIndex).PercentComposition = dblPercentComp
            
            
                ' Calculate standard deviation
                If .Elements(intElementIndex).IsotopicCorrection = 0 Then
                    ' No isotopic mass correction factor exists
                    dblStdDeviation = dblPercentComp * Sqr((ElementStats(intElementIndex).Uncertainty / ElementStats(intElementIndex).Mass) ^ 2 + _
                              (.StandardDeviation / .TotalMass) ^ 2)
                Else
                    ' Isotopic correction factor exists, assume no error in it
                    dblStdDeviation = dblPercentComp * Sqr((.StandardDeviation / .TotalMass) ^ 2)
                End If
                
                If dblElementTotalMass = .TotalMass And dblPercentComp = 100 Then
                    dblStdDeviation = 0
                End If
                
                .PercentCompositions(intElementIndex).StdDeviation = dblStdDeviation
            Else
                .PercentCompositions(intElementIndex).PercentComposition = 0
                .PercentCompositions(intElementIndex).StdDeviation = 0
            End If
        Next intElementIndex
    
    End With
                
End Sub

Public Sub ConstructMasterSymbolsList()
    ' Call after loading or changing abbreviations or elements
    ' Call after loading or setting abbreviation mode
    
    ReDim MasterSymbolsList(ELEMENT_COUNT + AbbrevAllCount, 1)
    Dim intIndex As Integer
    Dim blnIncludeAmino As Boolean
    
    ' MasterSymbolsList(,0) contains the symbol to be matched
    ' MasterSymbolsList(,1) contains E for element, A for amino acid, or N for normal abbreviation, followed by
    '     the reference number in the master list
    ' For example for Carbon, MasterSymbolsList(intIndex,0) = "C" and MasterSymbolsList(intIndex,1) = "E6"
    
    ' Construct search list
    For intIndex = 1 To ELEMENT_COUNT
        MasterSymbolsList(intIndex, 0) = ElementStats(intIndex).Symbol
        MasterSymbolsList(intIndex, 1) = "E" & Trim(Str(intIndex))
    Next intIndex
    MasterSymbolsListCount = ELEMENT_COUNT
    
    ' Note: AbbrevStats is 1-based
    If gComputationOptions.AbbrevRecognitionMode <> arNoAbbreviations Then
        If gComputationOptions.AbbrevRecognitionMode = arNormalPlusAminoAcids Then
            blnIncludeAmino = True
        Else
            blnIncludeAmino = False
        End If
        
        For intIndex = 1 To AbbrevAllCount
            With AbbrevStats(intIndex)
                ' If blnIncludeAmino = False then do not include amino acids
                If blnIncludeAmino Or (Not blnIncludeAmino And Not .IsAminoAcid) Then
                    ' Do not include if the formula is invalid
                    If Not .InvalidSymbolOrFormula Then
                        MasterSymbolsListCount = MasterSymbolsListCount + 1
                        
                        MasterSymbolsList(MasterSymbolsListCount, 0) = .Symbol
                        MasterSymbolsList(MasterSymbolsListCount, 1) = "A" & Trim(Str(intIndex))
                    End If
                End If
            End With
        Next intIndex
    End If
    
    ' Sort the search list
    ShellSortSymbols 1, CLng(MasterSymbolsListCount)

End Sub

Public Function ConvoluteMassInternal(ByVal dblMassMZ As Double, ByVal intCurrentCharge As Integer, Optional ByVal intDesiredCharge As Integer = 1, Optional ByVal dblChargeCarrierMass As Double = 0) As Double
    ' Converts dblMassMZ to the MZ that would appear at the given intDesiredCharge
    ' To return the neutral mass, set intDesiredCharge to 0

    ' If dblChargeCarrierMass is 0, then uses mChargeCarrierMass
    
    Const DEFAULT_CHARGE_CARRIER_MASS_MONOISO = 1.00727649

    Dim dblNewMZ As Double

    If dblChargeCarrierMass = 0 Then dblChargeCarrierMass = mChargeCarrierMass
    If dblChargeCarrierMass = 0 Then dblChargeCarrierMass = DEFAULT_CHARGE_CARRIER_MASS_MONOISO

    If intCurrentCharge = intDesiredCharge Then
        dblNewMZ = dblMassMZ
    Else
        If intCurrentCharge = 1 Then
            dblNewMZ = dblMassMZ
        ElseIf intCurrentCharge > 1 Then
            ' Convert dblMassMZ to M+H
            dblNewMZ = (dblMassMZ * intCurrentCharge) - dblChargeCarrierMass * (intCurrentCharge - 1)
        ElseIf intCurrentCharge = 0 Then
            ' Convert dblMassMZ (which is neutral) to M+H and store in dblNewMZ
            dblNewMZ = dblMassMZ + dblChargeCarrierMass
        Else
            ' Negative charges are not supported; return 0
            ConvoluteMassInternal = 0
            Exit Function
        End If

        If intDesiredCharge > 1 Then
            dblNewMZ = (dblNewMZ + dblChargeCarrierMass * (intDesiredCharge - 1)) / intDesiredCharge
        ElseIf intDesiredCharge = 1 Then
            ' Return M+H, which is currently stored in dblNewMZ
        ElseIf intDesiredCharge = 0 Then
            ' Return the neutral mass
            dblNewMZ = dblNewMZ - dblChargeCarrierMass
        Else
            ' Negative charges are not supported; return 0
            dblNewMZ = 0
        End If
    End If

    ConvoluteMassInternal = dblNewMZ

End Function

' Converts strFormula to its corresponding empirical formula
Public Function ConvertFormulaToEmpirical(ByVal strFormula As String) As String
    ' Returns the empirical formula for strFormula
    ' If an error occurs, returns -1
    
    Dim dblMass As Double
    Dim udtComputationStats As udtComputationStatsType
    Dim strEmpiricalFormula As String
    Dim intElementIndex As Integer, intElementSearchIndex As Integer
    Dim intElementIndexToUse As Integer
    Dim dblThisElementCount As Double
    
    ' Call ParseFormulaPublic to compute the formula's mass and fill udtComputationStats
    dblMass = ParseFormulaPublic(strFormula, udtComputationStats)
        
    If ErrorParams.ErrorID = 0 Then
        ' Convert to empirical formula
        strEmpiricalFormula = ""
        ' Carbon first, then hydrogen, then the rest alphabetically
        ' This is correct to start at -1
        For intElementIndex = -1 To ELEMENT_COUNT
            If intElementIndex = -1 Then
                ' Do Carbon first
                intElementIndexToUse = 6
            ElseIf intElementIndex = 0 Then
                ' Then do Hydrogen
                intElementIndexToUse = 1
            Else
                ' Then do the rest alphabetically
                If ElementAlph(intElementIndex) = "C" Or ElementAlph(intElementIndex) = "H" Then
                    ' Increment intElementIndex when we encounter carbon or hydrogen
                    intElementIndex = intElementIndex + 1
                End If
                For intElementSearchIndex = 2 To ELEMENT_COUNT    ' Start at 2 to since we've already done hydrogen
                    ' find the element in the numerically ordered array that corresponds to the alphabetically ordered array
                    If ElementStats(intElementSearchIndex).Symbol = ElementAlph(intElementIndex) Then
                        intElementIndexToUse = intElementSearchIndex
                        Exit For
                    End If
                Next intElementSearchIndex
            End If
            
            ' Only display the element if it's in the formula
            With mComputationStatsSaved
                dblThisElementCount = .Elements(intElementIndexToUse).Count
                If dblThisElementCount = 1# Then
                    strEmpiricalFormula = strEmpiricalFormula & ElementStats(intElementIndexToUse).Symbol
                ElseIf dblThisElementCount > 0 Then
                    strEmpiricalFormula = strEmpiricalFormula & ElementStats(intElementIndexToUse).Symbol & Trim(CStr(dblThisElementCount))
                End If
            End With
        Next intElementIndex
        
        ConvertFormulaToEmpirical = strEmpiricalFormula
    Else
        ConvertFormulaToEmpirical = -1
    End If
    
End Function

' Expands abbreviations in formula to elemental equivalent
Public Function ExpandAbbreviationsInFormula(ByVal strFormula As String) As String
    ' Returns the result
    ' If an error occurs, returns -1

    Dim dblMass As Double
    Dim udtComputationStats As udtComputationStatsType
    
    ' Call ExpandAbbreviationsInFormula to compute the formula's mass
    dblMass = ParseFormulaPublic(strFormula, udtComputationStats, True)
        
    If ErrorParams.ErrorID = 0 Then
        ExpandAbbreviationsInFormula = strFormula
      Else
        ExpandAbbreviationsInFormula = -1
    End If

End Function

Private Function FindIndexForNominalMass(IsoCombos() As Long, ComboIndex As Long, IsotopeCount As Integer, AtomCount As Long, ThisElementsIsotopes() As udtIsotopeInfoType) As Long
    Dim lngWorkingMass As Long
    Dim IsotopeIndex As Integer
    
    lngWorkingMass = 0
    For IsotopeIndex = 1 To IsotopeCount
        lngWorkingMass = lngWorkingMass + IsoCombos(ComboIndex, IsotopeIndex) * (Round(ThisElementsIsotopes(IsotopeIndex).Mass, 0))
    Next IsotopeIndex
    
'                             (lngWorkingMass - IsoStats(ElementIndex).StartingResultsMass) + 1
    FindIndexForNominalMass = (lngWorkingMass - AtomCount * Round(ThisElementsIsotopes(1).Mass, 0)) + 1
End Function

Private Sub ConvoluteMasses(ByRef ConvolutedAbundances() As udtIsoResultsOverallType, ByRef ConvolutedAbundanceStartMass As Long, ByRef WorkingRow As Long, ByRef WorkingAbundance As Single, ByRef WorkingMassTotal As Long, ByRef ElementTrack As Integer, ByRef IsoStats() As udtIsoResultsByElementType, ByRef ElementCount As Integer, ByRef Iterations As Long)
    ' Recursive function to Convolute the Results in IsoStats() and store in ConvolutedAbundances(); 1-based array
    
    Dim IndexToStoreResult As Long, RowIndex As Long
    Dim NewAbundance As Single, NewMassTotal As Long
    
    If KeyPressAbortProcess > 1 Then Exit Sub
    
    Iterations = Iterations + 1
    If Iterations Mod 10000 = 0 Then
        DoEvents
    End If
    
    NewAbundance = WorkingAbundance * IsoStats(ElementTrack).MassAbundances(WorkingRow)
    NewMassTotal = WorkingMassTotal + (IsoStats(ElementTrack).StartingResultsMass + WorkingRow - 1)
    
    If ElementTrack >= ElementCount Then
        IndexToStoreResult = NewMassTotal - ConvolutedAbundanceStartMass + 1
        With ConvolutedAbundances(IndexToStoreResult)
            If NewAbundance > 0 Then
                .Abundance = .Abundance + NewAbundance
                .Multiplicity = .Multiplicity + 1
            End If
        End With
    Else
        For RowIndex = 1 To IsoStats(ElementTrack + 1).ResultsCount
            ConvoluteMasses ConvolutedAbundances(), ConvolutedAbundanceStartMass, RowIndex, NewAbundance, NewMassTotal, ElementTrack + 1, IsoStats(), ElementCount, Iterations
        Next RowIndex
    End If
End Sub

' Note: This function is unused
Private Function FindCombinations(Optional AtomCount As Long = 2, Optional IsotopeCount As Integer = 2, Optional boolPrintOutput As Boolean = False) As Long
    ' Find Combinations of atoms for a given number of atoms and number of potential isotopes
    ' Can print results to debug window
    
    Dim ComboResults() As Long
    Dim AtomTrackHistory() As Long
    Dim PredictedCombos As Long, CombosFound As Long
    
    PredictedCombos = FindCombosPredictIterations(AtomCount, IsotopeCount)
    
    If PredictedCombos > 10000000 Then
        MsgBox "Too many combinations necessary for prediction of isotopic distribution: " & Format(PredictedCombos, "#,##0") & vbCrLf & "Please use a simpler formula or reduce the isotopic range defined for the element (currently " & IsotopeCount & ")"
        FindCombinations = -1
        Exit Function
    End If
    
    On Error GoTo FindCombinationsErrorHandler
    
    ReDim ComboResults(PredictedCombos, IsotopeCount)
    
    ReDim AtomTrackHistory(IsotopeCount)
    AtomTrackHistory(1) = AtomCount
    
    CombosFound = FindCombosRecurse(ComboResults(), AtomCount, IsotopeCount, IsotopeCount, 1, 1, AtomTrackHistory())
    
    If boolPrintOutput Then
        Dim RowIndex As Long, ColIndex As Integer
        Dim strOutput As String, strHeader As String

        strHeader = CombosFound & " combos found for " & AtomCount & " atoms for element with " & IsotopeCount & " isotopes"
        If CombosFound > 5000 Then
            strHeader = strHeader & vbCrLf & "Only displaying the first 5000 combinations"
        End If
        
        Debug.Print strHeader
        
        For RowIndex = 1 To CombosFound
            strOutput = ""
            For ColIndex = 1 To IsotopeCount
                strOutput = strOutput & ComboResults(RowIndex, ColIndex) & vbTab
            Next ColIndex
            Debug.Print strOutput
            If RowIndex > 5000 Then Exit For
        Next RowIndex
        
        If CombosFound > 50 Then Debug.Print strHeader
    
    End If
    
    FindCombinations = CombosFound
LeaveSub:
    Exit Function
    
FindCombinationsErrorHandler:
    MwtWinDllErrorHandler "MwtWinDll|FindCombinations"
    ErrorParams.ErrorID = 590
    ErrorParams.ErrorPosition = 0
    FindCombinations = -1
    Resume LeaveSub

End Function

Private Function FindCombosPredictIterations(AtomCount As Long, IsotopeCount As Integer) As Long
    ' Determines the number of Combo results (iterations) for the given
    ' number of Atoms for an element with the given number of Isotopes
    
    ' Empirically determined the following results and figured out that the RunningSum()
    '  method correctly predicts the results
    
    ' IsotopeCount   AtomCount    Total Iterations
    '     2             1               2
    '     2             2               3
    '     2             3               4
    '     2             4               5
    '
    '     3             1               3
    '     3             2               6
    '     3             3               10
    '     3             4               15
    '
    '     4             1               4
    '     4             2               10
    '     4             3               20
    '     4             4               35
    '
    '     5             1               5
    '     5             2               15
    '     5             3               35
    '     5             4               70
    '
    '     6             1               6
    '     6             2               21
    '     6             3               56
    '     6             4               126
    
    Dim IsotopeIndex As Integer, AtomIndex As Long
    Dim PredictedCombos As Long
    Dim RunningSum() As Long, PreviousComputedValue As Long
    
    ReDim RunningSum(AtomCount)
    
    On Error GoTo PredictCombinationsErrorHandler
    
    If AtomCount = 1 Or IsotopeCount = 1 Then
        PredictedCombos = IsotopeCount
    Else
        ' Initialize RunningSum()
        For AtomIndex = 1 To AtomCount
            RunningSum(AtomIndex) = AtomIndex + 1
        Next AtomIndex
        
        For IsotopeIndex = 3 To IsotopeCount
            PreviousComputedValue = IsotopeIndex
            For AtomIndex = 2 To AtomCount
                ' Compute new count for this AtomIndex
                RunningSum(AtomIndex) = PreviousComputedValue + RunningSum(AtomIndex)
                
                ' Also place result in RunningSum(AtomIndex)
                PreviousComputedValue = RunningSum(AtomIndex)
            Next AtomIndex
        Next IsotopeIndex
        
        PredictedCombos = RunningSum(AtomCount)
    End If

    FindCombosPredictIterations = PredictedCombos
    
LeaveSub:
    Exit Function
    
PredictCombinationsErrorHandler:
    MwtWinDllErrorHandler "MwtWinDll|FindCombosPredictIterations"
    ErrorParams.ErrorID = 590
    ErrorParams.ErrorPosition = 0
    FindCombosPredictIterations = -1
    Resume LeaveSub
    
End Function

Private Function FindCombosRecurse(ByRef ComboResults() As Long, ByRef AtomCount As Long, ByRef MaxIsotopeCount As Integer, CurrentIsotopeCount As Integer, CurrentRow As Long, CurrentCol As Integer, AtomTrackHistory() As Long) As Long
    ' Recursive function to find all the combinations
    ' of a number of atoms with the given maximum isotopic count
    
    ' IsoCombos() is a 2D array holding the # of each isotope for each combination
    ' For example, Two chlorines, Cl2, has at most 6 combos since Cl isotopes are 35, 36, and 37
    ' m1  m2  m3
    ' 2   0   0
    ' 1   1   0
    ' 1   0   1
    ' 0   2   0
    ' 0   1   1
    ' 0   0   2

    ' Returns the number of combinations found, or -1 if an error
    
    Dim ColIndex As Integer
    Dim AtomTrack As Long
    Dim intNewColumn As Integer
    
    If CurrentIsotopeCount = 1 Or AtomCount = 0 Then
        ' End recursion
        ComboResults(CurrentRow, CurrentCol) = AtomCount
    Else
        AtomTrack = AtomCount
        
        ' Store AtomTrack value at current position
        ComboResults(CurrentRow, CurrentCol) = AtomTrack
        
        Do While AtomTrack > 0
            CurrentRow = CurrentRow + 1
            
            ' Went to a new row; if CurrentCol > 1 then need to assign previous values to previous columns
            If (CurrentCol > 1) Then
                For ColIndex = 1 To CurrentCol - 1
                    ComboResults(CurrentRow, ColIndex) = AtomTrackHistory(ColIndex)
                Next ColIndex
            End If
            
            AtomTrack = AtomTrack - 1
            ComboResults(CurrentRow, CurrentCol) = AtomTrack
            
            If CurrentCol < MaxIsotopeCount Then
                intNewColumn = CurrentCol + 1
                AtomTrackHistory(intNewColumn - 1) = AtomTrack
                FindCombosRecurse ComboResults(), AtomCount - AtomTrack, MaxIsotopeCount, CurrentIsotopeCount - 1, CurrentRow, intNewColumn, AtomTrackHistory
            Else
                Debug.Assert "Program bug. This line should not be reached."
            End If
        Loop
        
        ' Reached AtomTrack = 0; end recursion
    End If
    
    FindCombosRecurse = CurrentRow
    
End Function

Public Sub GeneralErrorHandler(strCallingProcedure As String, lngErrorNumber As Long, Optional strErrorDescriptionAddnl As String = "")
    Dim strMessage As String
    Dim strErrorFilePath As String
    Dim intOutFileNum As Integer
    
    strMessage = "Error in " & strCallingProcedure & ": " & Error(lngErrorNumber) & " (#" & Trim(lngErrorNumber) & ")"
    If Len(strErrorDescriptionAddnl) > 0 Then strMessage = strMessage & vbCrLf & strErrorDescriptionAddnl
    
    If mShowErrorMessageDialogs Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error in MwtWinDll"
    Else
        Debug.Print strMessage
        Debug.Assert False
    End If
    
    On Error Resume Next
    strErrorFilePath = App.Path
    If Right(strErrorFilePath, 1) <> "\" Then strErrorFilePath = strErrorFilePath & "\"
    strErrorFilePath = strErrorFilePath & "ErrorLog.txt"
    
    intOutFileNum = FreeFile()
    Open strErrorFilePath For Append As #intOutFileNum
    
    Print #intOutFileNum, Now() & " -- " & strMessage & vbCrLf
    Close #intOutFileNum

End Sub

Public Function GetAbbreviationCountInternal() As Long
    GetAbbreviationCountInternal = AbbrevAllCount
End Function

Public Function GetAbbreviationIDInternal(ByVal strSymbol As String, Optional blnAminoAcidsOnly As Boolean = False) As Long
    ' Returns 0 if not found, the ID if found
    
    Dim lngIndex As Long
    
    For lngIndex = 1 To AbbrevAllCount
        If LCase(AbbrevStats(lngIndex).Symbol) = LCase(strSymbol) Then
            If Not blnAminoAcidsOnly Or (blnAminoAcidsOnly And AbbrevStats(lngIndex).IsAminoAcid) Then
                GetAbbreviationIDInternal = lngIndex
                Exit Function
            End If
        End If
    Next lngIndex
    
    GetAbbreviationIDInternal = 0
End Function

Public Function GetAbbreviationInternal(ByVal lngAbbreviationID As Long, ByRef strSymbol As String, ByRef strFormula As String, ByRef sngCharge As Single, ByRef blnIsAminoAcid As Boolean, Optional ByRef strOneLetterSymbol As String, Optional ByRef strComment As String, Optional ByRef blnInvalidSymbolOrFormula As Boolean) As Long
    ' Returns the contents of AbbrevStats() in the ByRef variables
    ' Returns 0 if success, 1 if failure
    
    If lngAbbreviationID >= 1 And lngAbbreviationID <= AbbrevAllCount Then
        With AbbrevStats(lngAbbreviationID)
            strSymbol = .Symbol
            strFormula = .Formula
            sngCharge = .Charge
            blnIsAminoAcid = .IsAminoAcid
            strOneLetterSymbol = .OneLetterSymbol
            strComment = .Comment
            blnInvalidSymbolOrFormula = .InvalidSymbolOrFormula
        End With
        GetAbbreviationInternal = 0
    Else
        GetAbbreviationInternal = 1
    End If
    
End Function

Public Function GetAbbreviationMass(ByVal lngAbbreviationID As Long) As Double
    ' Returns the mass if success, 0 if failure
    ' Could return -1 if failure, but might mess up some calculations
    
    ' This function does not recompute the abbreviation mass each time it is called
    ' Rather, it uses the .Mass member of AbbrevStats
    ' This requires that .Mass be updated if the abbreviation is changed, if an element is changed, or if the element mode is changed
    
    If lngAbbreviationID >= 1 And lngAbbreviationID <= AbbrevAllCount Then
        GetAbbreviationMass = AbbrevStats(lngAbbreviationID).Mass
    Else
        GetAbbreviationMass = 0
    End If
    
End Function

Public Function GetAminoAcidSymbolConversionInternal(strSymbolToFind As String, bln1LetterTo3Letter As Boolean) As String
    ' If bln1LetterTo3Letter = True, then converting 1 letter codes to 3 letter codes
    ' Returns the symbol, if found
    ' Otherwise, returns ""
    Dim lngIndex As Long, strReturnSymbol As String, strCompareSymbol As String
    
    strReturnSymbol = ""
    ' Use AbbrevStats() array to lookup code
    For lngIndex = 1 To AbbrevAllCount
        If AbbrevStats(lngIndex).IsAminoAcid Then
            If bln1LetterTo3Letter Then
                strCompareSymbol = AbbrevStats(lngIndex).OneLetterSymbol
            Else
                strCompareSymbol = AbbrevStats(lngIndex).Symbol
            End If
            
            If LCase(strCompareSymbol) = LCase(strSymbolToFind) Then
                If bln1LetterTo3Letter Then
                    strReturnSymbol = AbbrevStats(lngIndex).Symbol
                Else
                    strReturnSymbol = AbbrevStats(lngIndex).OneLetterSymbol
                End If
                Exit For
            End If
        End If
    Next lngIndex
    
    GetAminoAcidSymbolConversionInternal = strReturnSymbol
    
End Function

Public Function GetCautionStatementCountInternal() As Long
    GetCautionStatementCountInternal = CautionStatementCount
End Function

Public Function GetCautionStatementIDInternal(ByVal strSymbolCombo As String) As Long
    ' Returns -1 if not found, the ID if found
    
    Dim intIndex As Integer
    
    For intIndex = 1 To CautionStatementCount
        If CautionStatements(intIndex, 0) = strSymbolCombo Then
            GetCautionStatementIDInternal = intIndex
            Exit Function
        End If
    Next intIndex
    
    GetCautionStatementIDInternal = -1
End Function

Public Function GetCautionStatementInternal(ByVal lngCautionStatementID As Long, ByRef strSymbolCombo As String, ByRef strCautionStatement As String) As Long
    ' Returns the contents of CautionStatements() in the ByRef variables
    ' Returns 0 if success, 1 if failure
    
    If lngCautionStatementID >= 1 And lngCautionStatementID <= CautionStatementCount Then
        strSymbolCombo = CautionStatements(lngCautionStatementID, 0)
        strCautionStatement = CautionStatements(lngCautionStatementID, 1)
        GetCautionStatementInternal = 0
    Else
        GetCautionStatementInternal = 1
    End If
End Function

Public Function GetCautionDescription() As String
    GetCautionDescription = mStrCautionDescription
End Function

Public Function GetChargeCarrierMassInternal() As Double
    GetChargeCarrierMassInternal = mChargeCarrierMass
End Function

Public Function GetElementCountInternal() As Long
    GetElementCountInternal = ELEMENT_COUNT
End Function

Public Function GetElementInternal(ByVal intElementID As Integer, ByRef strSymbol As String, ByRef dblMass As Double, ByRef dblUncertainty As Double, ByRef sngCharge As Single, ByRef intIsotopeCount As Integer) As Long
    ' Returns the contents of ElementStats() in the ByRef variables
    ' Returns 0 if success, 1 if failure
    
    If intElementID >= 1 And intElementID <= ELEMENT_COUNT Then
        strSymbol = ElementAlph(intElementID)
        With ElementStats(intElementID)
            strSymbol = .Symbol
            dblMass = .Mass
            dblUncertainty = .Uncertainty
            sngCharge = .Charge
            intIsotopeCount = .IsotopeCount
        End With
        GetElementInternal = 0
    Else
        GetElementInternal = 1
    End If
    
End Function

Public Function GetElementIDInternal(ByVal strSymbol As String) As Long
    ' Returns 0 if not found, the ID if found
    Dim intIndex As Integer
    
    For intIndex = 1 To ELEMENT_COUNT
        If LCase(ElementStats(intIndex).Symbol) = LCase(strSymbol) Then
            GetElementIDInternal = intIndex
            Exit Function
        End If
    Next intIndex
    
    GetElementIDInternal = 0
End Function

Public Function GetElementIsotopesInternal(ByVal intElementID As Integer, ByRef intIsotopeCount As Integer, ByRef dblIsotopeMasses() As Double, ByRef sngIsotopeAbundances() As Single) As Long
    Dim intIsotopeindex As Integer

    If intElementID >= 1 And intElementID <= ELEMENT_COUNT Then
        With ElementStats(intElementID)
            intIsotopeCount = .IsotopeCount
            For intIsotopeindex = 1 To .IsotopeCount
                dblIsotopeMasses(intIsotopeindex) = .Isotopes(intIsotopeindex).Mass
                sngIsotopeAbundances(intIsotopeindex) = .Isotopes(intIsotopeindex).Abundance
            Next intIsotopeindex
        End With
        GetElementIsotopesInternal = 0
    Else
        GetElementIsotopesInternal = 1
    End If
End Function

Public Function GetElementModeInternal() As emElementModeConstantsPrivate
    GetElementModeInternal = mCurrentElementMode
End Function

Public Function GetElementSymbolInternal(ByVal intElementID As Integer) As String
    
    If intElementID >= 1 And intElementID <= ELEMENT_COUNT Then
        GetElementSymbolInternal = ElementStats(intElementID).Symbol
    Else
        GetElementSymbolInternal = ""
    End If
    
End Function

Public Function GetElementStatInternal(ByVal intElementID As Integer, ByVal eElementStat As esElementStatsConstants) As Double
    ' Returns a single bit of information about a single element
    ' Since a value may be negavite, simply returns 0 if an error
    
    If intElementID >= 1 And intElementID <= ELEMENT_COUNT Then
        Select Case eElementStat
        Case esMass
            GetElementStatInternal = ElementStats(intElementID).Mass
        Case esCharge
            GetElementStatInternal = ElementStats(intElementID).Charge
        Case esUncertainty
            GetElementStatInternal = ElementStats(intElementID).Uncertainty
        Case Else
            GetElementStatInternal = 0
        End Select
    Else
        GetElementStatInternal = 0
    End If
    
End Function

Public Function GetErrorDescription() As String
    If ErrorParams.ErrorID <> 0 Then
        GetErrorDescription = LookupMessage(ErrorParams.ErrorID)
    Else
        GetErrorDescription = ""
    End If
End Function

Public Function GetErrorID() As Long
    GetErrorID = ErrorParams.ErrorID
End Function

Public Function GetErrorCharacter() As String
    GetErrorCharacter = ErrorParams.ErrorCharacter
End Function

Public Function GetErrorPosition() As Long
    GetErrorPosition = ErrorParams.ErrorPosition
End Function

Public Function GetMessageStatementCountInternal() As Long
    GetMessageStatementCountInternal = MessageStatmentCount
End Function

Public Function GetMessageStatementInternal(lngMessageID As Long, Optional strAppendText As String = "") As String
' GetMessageStringInternal simply returns the message for lngMessageID
' LookupMessage formats the message, and possibly combines multiple messages, depending on the message number
    Dim strMessage As String
    
    If lngMessageID > 0 And lngMessageID <= MessageStatmentCount Then
        strMessage = MessageStatements(lngMessageID)
        
        ' Append Prefix to certain strings
        Select Case lngMessageID
            ' Add Error: to the front of certain error codes
        Case 1 To 99, 120, 130, 140, 260, 270, 300
            strMessage = GetMessageStatementInternal(350) & ": " & strMessage
        End Select
        
        ' Now append strAppendText
        GetMessageStatementInternal = strMessage & strAppendText
    Else
        GetMessageStatementInternal = ""
    End If
End Function

Private Function IsPresentInAbbrevSymbolStack(udtAbbrevSymbolStack As udtAbbrevSymbolStackType, SymbolReference As Integer) As Boolean
    ' Checks for presence of SymbolReference in udtAbbrevSymbolStack
    ' If found, returns true
    
    Dim intIndex As Integer
    Dim blnFound As Boolean
    
On Error GoTo IsPresentInAbbrevSymbolStackErrorHandler
    
    With udtAbbrevSymbolStack
        blnFound = False
        For intIndex = 0 To .Count - 1
            If .SymbolReferenceStack(intIndex) = SymbolReference Then
                blnFound = True
                Exit For
            End If
        Next intIndex
    End With
    
    IsPresentInAbbrevSymbolStack = blnFound
    Exit Function
    
IsPresentInAbbrevSymbolStackErrorHandler:
    Debug.Assert False
    IsPresentInAbbrevSymbolStack = False
    
End Function

Public Function IsModSymbolInternal(strTestChar As String) As Boolean
    ' Returns True if the first letter of strTestChar is a ModSymbol
    ' Invalid Mod Symbols are letters, numbers, ., -, space, (, or )
    ' Valid Mod Symbols are ! # $ % & ' * + ? ^ _ ` ~

    Dim strFirstChar As String, blnIsModSymbol As Boolean
    
    blnIsModSymbol = False
    If Len(strTestChar) > 0 Then
        strFirstChar = Left(strTestChar, 1)
    
        Select Case Asc(strFirstChar)
        Case 34           ' " is not allowed
            blnIsModSymbol = False
        Case 40 To 41     ' ( and ) are not allowed
            blnIsModSymbol = False
        Case 44 To 62     ' . and - and , and / and numbers and : and ; and < and = and > are not allowed
            blnIsModSymbol = False
        Case 33 To 43, 63 To 64, 94 To 96, 126
            blnIsModSymbol = True
        Case Else
            blnIsModSymbol = False
        End Select
    Else
        blnIsModSymbol = False
    End If

    IsModSymbolInternal = blnIsModSymbol
    
End Function

Private Function IsStringAllLetters(strTest As String) As Boolean
    ' Tests if all of the characers in strTest are letters
    
    Dim blnAllLetters As Boolean, intIndex As Integer
    
    ' Assume true until proven otherwise
    blnAllLetters = True
    For intIndex = 1 To Len(strTest)
        If Not IsCharacter(Mid(strTest, intIndex, 1)) Then
            blnAllLetters = False
            Exit For
        End If
    Next intIndex
    
    IsStringAllLetters = blnAllLetters
End Function

Private Function LookupCautionStatement(strCompareText As String) As String
    Dim intIndex As Integer
    
    For intIndex = 1 To CautionStatementCount
        If strCompareText = CautionStatements(intIndex, 0) Then
            LookupCautionStatement = CautionStatements(intIndex, 1)
            Exit For
        End If
    Next intIndex
End Function

Private Function LookupMessage(lngMessageID As Long, Optional strAppendText As String) As String
    ' Looks up the message for lngMessageID
    ' Also appends any data in strAppendText to the message
    ' Returns the complete message
    
    Dim strMessage As String
    
    If MessageStatmentCount = 0 Then MemoryLoadMessageStatements
    
    ' First assume we can't find the message number
    strMessage = "General unspecified error"
    
    ' Now try to find it
    If lngMessageID < MESSAGE_STATEMENT_DIMCOUNT Then
        If Len(MessageStatements(lngMessageID)) > 0 Then
            strMessage = MessageStatements(lngMessageID)
        End If
    End If
    
    ' Now prepend Prefix to certain strings
    Select Case lngMessageID
        ' Add Error: to the front of certain error codes
    Case 1 To 99, 120, 130, 140, 260, 270, 300
        strMessage = LookupMessage(350) & ": " & strMessage
    End Select
    
    ' Now append strAppendText
    strMessage = strMessage & strAppendText
    
    ' lngMessageID's 1 and 18 may need to have an addendum added
    If lngMessageID = 1 Then
        If gComputationOptions.CaseConversion = ccExactCase Then
            strMessage = strMessage & " (" & LookupMessage(680) & ")"
        End If
    ElseIf lngMessageID = 18 Then
        If Not gComputationOptions.BracketsAsParentheses Then
            strMessage = strMessage & " (" & LookupMessage(685) & ")"
        Else
            strMessage = strMessage & " (" & LookupMessage(690) & ")"
        End If
    End If
    
    LookupMessage = strMessage
End Function

Public Function MassToPPMInternal(ByVal dblMassToConvert As Double, ByVal dblCurrentMZ As Double) As Double
    ' Converts dblMassToConvert to ppm, based on the value of dblCurrentMZ
    
    If dblCurrentMZ > 0 Then
        MassToPPMInternal = dblMassToConvert * 1000000# / dblCurrentMZ
    Else
        MassToPPMInternal = 0
    End If
End Function

Public Function MonoMassToMZInternal(ByVal dblMonoisotopicMass As Double, ByVal intCharge As Integer, Optional ByVal dblChargeCarrierMass As Double = 0) As Double
    ' If dblChargeCarrierMass is 0, then uses mChargeCarrierMass
    If dblChargeCarrierMass = 0 Then dblChargeCarrierMass = mChargeCarrierMass
    
    ' Call ConvoluteMass to convert to the desired charge state
    MonoMassToMZInternal = ConvoluteMassInternal(dblMonoisotopicMass + dblChargeCarrierMass, 1, intCharge, dblChargeCarrierMass)
End Function

Public Sub MemoryLoadAll(eElementMode As emElementModeConstantsPrivate)
    
    MemoryLoadElements eElementMode
    
    ' Reconstruct master symbols list
    ConstructMasterSymbolsList
    
    MemoryLoadIsotopes
    
    MemoryLoadAbbreviations
   
    ' Reconstruct master symbols list
    ConstructMasterSymbolsList
    
    MemoryLoadCautionStatements
    
    MemoryLoadMessageStatements
    
End Sub

Public Sub MemoryLoadAbbreviations()
    Dim intIndex As Integer
    
    ' Symbol                            Formula            1 letter abbreviation
    Const AminoAbbrevCount = 28
    
    AbbrevAllCount = AminoAbbrevCount
    For intIndex = 1 To AbbrevAllCount
        AbbrevStats(intIndex).IsAminoAcid = True
    Next intIndex
    
    AddAbbreviationWork 1, "Ala", "C3H5NO", 0, True, "A", "Alanine"
    AddAbbreviationWork 2, "Arg", "C6H12N4O", 0, True, "R", "Arginine, (unprotonated NH2)"
    AddAbbreviationWork 3, "Asn", "C4H6N2O2", 0, True, "N", "Asparagine"
    AddAbbreviationWork 4, "Asp", "C4H5NO3", 0, True, "D", "Aspartic acid (undissociated COOH)"
    AddAbbreviationWork 5, "Cys", "C3H5NOS", 0, True, "C", "Cysteine (no disulfide link)"
    AddAbbreviationWork 6, "Gla", "C6H7NO5", 0, True, "U", "gamma-Carboxyglutamate"
    AddAbbreviationWork 7, "Gln", "C5H8N2O2", 0, True, "Q", "Glutamine"
    AddAbbreviationWork 8, "Glu", "C5H7NO3", 0, True, "E", "Glutamic acid (undissociated COOH)"
    AddAbbreviationWork 9, "Gly", "C2H3NO", 0, True, "G", "Glycine"
    AddAbbreviationWork 10, "His", "C6H7N3O", 0, True, "H", "Histidine (unprotonated NH)"
    AddAbbreviationWork 11, "Hse", "C4H7NO2", 0, True, "", "Homoserine"
    AddAbbreviationWork 12, "Hyl", "C6H12N2O2", 0, True, "", "Hydroxylysine"
    AddAbbreviationWork 13, "Hyp", "C5H7NO2", 0, True, "", "Hydroxyproline"
    AddAbbreviationWork 14, "Ile", "C6H11NO", 0, True, "I", "Isoleucine"
    AddAbbreviationWork 15, "Leu", "C6H11NO", 0, True, "L", "Leucine"
    AddAbbreviationWork 16, "Lys", "C6H12N2O", 0, True, "K", "Lysine (unprotonated NH2)"
    AddAbbreviationWork 17, "Met", "C5H9NOS", 0, True, "M", "Methionine"
    AddAbbreviationWork 18, "Orn", "C5H10N2O", 0, True, "O", "Ornithine"
    AddAbbreviationWork 19, "Phe", "C9H9NO", 0, True, "F", "Phenylalanine"
    AddAbbreviationWork 20, "Pro", "C5H7NO", 0, True, "P", "Proline"
    AddAbbreviationWork 21, "Pyr", "C5H5NO2", 0, True, "", "Pyroglutamic acid"
    AddAbbreviationWork 22, "Sar", "C3H5NO", 0, True, "", "Sarcosine"
    AddAbbreviationWork 23, "Ser", "C3H5NO2", 0, True, "S", "Serine"
    AddAbbreviationWork 24, "Thr", "C4H7NO2", 0, True, "T", "Threonine"
    AddAbbreviationWork 25, "Trp", "C11H10N2O", 0, True, "W", "Tryptophan"
    AddAbbreviationWork 26, "Tyr", "C9H9NO2", 0, True, "Y", "Tyrosine"
    AddAbbreviationWork 27, "Val", "C5H9NO", 0, True, "V", "Valine"
    AddAbbreviationWork 28, "Xxx", "C6H12N2O", 0, True, "X", "Unknown"
    
    Const NormalAbbrevCount = 16
    AbbrevAllCount = AbbrevAllCount + NormalAbbrevCount
    For intIndex = AminoAbbrevCount + 1 To AbbrevAllCount
        AbbrevStats(intIndex).IsAminoAcid = False
    Next intIndex
    
    AddAbbreviationWork AminoAbbrevCount + 1, "Bpy", "C10H8N2", 0, False, "", "Bipyridine"
    AddAbbreviationWork AminoAbbrevCount + 2, "Bu", "C4H9", 1, False, "", "Butyl"
    AddAbbreviationWork AminoAbbrevCount + 3, "D", "^2.014H", 1, False, "", "Deuterium"
    AddAbbreviationWork AminoAbbrevCount + 4, "En", "C2H8N2", 0, False, "", "Ethylenediamine"
    AddAbbreviationWork AminoAbbrevCount + 5, "Et", "CH3CH2", 1, False, "", "Ethyl"
    AddAbbreviationWork AminoAbbrevCount + 6, "Me", "CH3", 1, False, "", "Methyl"
    AddAbbreviationWork AminoAbbrevCount + 7, "Ms", "CH3SOO", -1, False, "", "Mesyl"
    AddAbbreviationWork AminoAbbrevCount + 8, "Oac", "C2H3O2", -1, False, "", "Acetate"
    AddAbbreviationWork AminoAbbrevCount + 9, "Otf", "OSO2CF3", -1, False, "", "Triflate"
    AddAbbreviationWork AminoAbbrevCount + 10, "Ox", "C2O4", -2, False, "", "Oxalate"
    AddAbbreviationWork AminoAbbrevCount + 11, "Ph", "C6H5", 1, False, "", "Phenyl"
    AddAbbreviationWork AminoAbbrevCount + 12, "Phen", "C12H8N2", 0, False, "", "Phenanthroline"
    AddAbbreviationWork AminoAbbrevCount + 13, "Py", "C5H5N", 0, False, "", "Pyridine"
    AddAbbreviationWork AminoAbbrevCount + 14, "Tpp", "(C4H2N(C6H5C)C4H2N(C6H5C))2", 0, False, "", "Tetraphenylporphyrin"
    AddAbbreviationWork AminoAbbrevCount + 15, "Ts", "CH3C6H4SOO", -1, False, "", "Tosyl"
    AddAbbreviationWork AminoAbbrevCount + 16, "Urea", "H2NCONH2", 0, False, "", "Urea"
                                              
'    ' Note Asx or B is often used for Asp or Asn
'    ' Note Glx or Z is often used for Glu or Gln
'    ' Note X is often used for "unknown"
'
'    ' Other amino acids without widely agreed upon 1 letter codes
'
'    FlexGridAddItems .grdAminoAcids, "Aminosuberic Acid", "Asu"     ' A pair of Cys residues bonded by S-S
'    FlexGridAddItems .grdAminoAcids, "Cystine", "Cyn"
'    FlexGridAddItems .grdAminoAcids, "Homocysteine", "Hcy"
'    FlexGridAddItems .grdAminoAcids, "Homoserine", "Hse"            ' 101.04 (C4H7NO2)
'    FlexGridAddItems .grdAminoAcids, "Hydroxylysine", "Hyl"         ' 144.173 (C6H12N2O2)
'    FlexGridAddItems .grdAminoAcids, "Hydroxyproline", "Hyp"        ' 113.116 (C5H7NO2)
'    FlexGridAddItems .grdAminoAcids, "Norleucine", "Nle"            ' 113.06
'    FlexGridAddItems .grdAminoAcids, "Norvaline", "Nva"
'    FlexGridAddItems .grdAminoAcids, "Pencillamine", "Pen"
'    FlexGridAddItems .grdAminoAcids, "Phosphoserine", "Sep"
'    FlexGridAddItems .grdAminoAcids, "Phosphothreonine", "Thp"
'    FlexGridAddItems .grdAminoAcids, "Phosphotyrosine", "Typ"
'    FlexGridAddItems .grdAminoAcids, "Pyroglutamic Acid", "Pyr"     ' 111.03 (C5H5NO2) (also Glp in some tables)
'    FlexGridAddItems .grdAminoAcids, "Sarcosine", "Sar"             ' 71.08 (C3H5NO)
'    FlexGridAddItems .grdAminoAcids, "Statine", "Sta"
'    FlexGridAddItems .grdAminoAcids, "b-[2-Thienyl]Ala", "Thi"
        

' Need to explore http://www.abrf.org/ABRF/ResearchCommittees/deltamass/deltamass.html
    
    ' Isoelectric points
'   TYR   Y   C9H9NO2     163.06333  163.1760      0               9.8
'   HIS   H   C6H7N3O     137.05891  137.1411      1               6.8
'   LYS   K   C6H12N2O    128.09496  128.1741      1              10.1
'   ASP   D   C4H5NO3     115.02694  115.0886      1               4.5
'   GLU   E   C5H7NO3     129.04259  129.1155      1               4.5
'   CYS   C   C3H5NOS     103.00919  103.1388      0               8.6
'   ARG   R   C6H12N4O    156.10111  156.1875      1              12.0

                                              
End Sub

' Use objMwtWin.ClearCautionStatements and objMwtWin.AddCautionStatement to
'  set these based on language
Public Sub MemoryLoadCautionStatements()
        
    CautionStatementCount = 41
    
    CautionStatements(1, 0) = "Bi":   CautionStatements(1, 1) = "Bi means bismuth; BI means boron-iodine.  "
    CautionStatements(2, 0) = "Bk":   CautionStatements(2, 1) = "Bk means berkelium; BK means boron-potassium.  "
    CautionStatements(3, 0) = "Bu":   CautionStatements(3, 1) = "Bu means the butyl group; BU means boron-uranium.  "
    CautionStatements(4, 0) = "Cd":   CautionStatements(4, 1) = "Cd means cadmium; CD means carbon-deuterium.  "
    CautionStatements(5, 0) = "Cf":   CautionStatements(5, 1) = "Cf means californium; CF means carbon-fluorine.  "
    CautionStatements(6, 0) = "Co":   CautionStatements(6, 1) = "Co means cobalt; CO means carbon-oxygen.  "
    CautionStatements(7, 0) = "Cs":   CautionStatements(7, 1) = "Cs means cesium; CS means carbon-sulfur.  "
    CautionStatements(8, 0) = "Cu":   CautionStatements(8, 1) = "Cu means copper; CU means carbon-uranium.  "
    CautionStatements(9, 0) = "Dy":   CautionStatements(9, 1) = "Dy means dysprosium; DY means deuterium-yttrium.  "
    CautionStatements(10, 0) = "Hf":  CautionStatements(10, 1) = "Hf means hafnium; HF means hydrogen-fluorine.  "
    CautionStatements(11, 0) = "Ho":  CautionStatements(11, 1) = "Ho means holmium; HO means hydrogen-oxygen.  "
    CautionStatements(12, 0) = "In":  CautionStatements(12, 1) = "In means indium; IN means iodine-nitrogen.  "
    CautionStatements(13, 0) = "Nb":  CautionStatements(13, 1) = "Nb means niobium; NB means nitrogen-boron.  "
    CautionStatements(14, 0) = "Nd":  CautionStatements(14, 1) = "Nd means neodymium; ND means nitrogen-deuterium.  "
    CautionStatements(15, 0) = "Ni":  CautionStatements(15, 1) = "Ni means nickel; NI means nitrogen-iodine.  "
    CautionStatements(16, 0) = "No":  CautionStatements(16, 1) = "No means nobelium; NO means nitrogen-oxygen.  "
    CautionStatements(17, 0) = "Np":  CautionStatements(17, 1) = "Np means neptunium; NP means nitrogen-phosphorus.  "
    CautionStatements(18, 0) = "Os":  CautionStatements(18, 1) = "Os means osmium; OS means oxygen-sulfur.  "
    CautionStatements(19, 0) = "Pd":  CautionStatements(19, 1) = "Pd means palladium; PD means phosphorus-deuterium.  "
    CautionStatements(20, 0) = "Ph":  CautionStatements(20, 1) = "Ph means phenyl, PH means phosphorus-hydrogen.  "
    CautionStatements(21, 0) = "Pu":  CautionStatements(21, 1) = "Pu means plutonium; PU means phosphorus-uranium.  "
    CautionStatements(22, 0) = "Py":  CautionStatements(22, 1) = "Py means pyridine; PY means phosphorus-yttrium.  "
    CautionStatements(23, 0) = "Sb":  CautionStatements(23, 1) = "Sb means antimony; SB means sulfor-boron.  "
    CautionStatements(24, 0) = "Sc":  CautionStatements(24, 1) = "Sc means scandium; SC means sulfur-carbon.  "
    CautionStatements(25, 0) = "Si":  CautionStatements(25, 1) = "Si means silicon; SI means sulfur-iodine.  "
    CautionStatements(26, 0) = "Sn":  CautionStatements(26, 1) = "Sn means tin; SN means sulfor-nitrogen.  "
    CautionStatements(27, 0) = "TI":  CautionStatements(27, 1) = "TI means tritium-iodine, Ti means titanium.  "
    CautionStatements(28, 0) = "Yb":  CautionStatements(28, 1) = "Yb means ytterbium; YB means yttrium-boron.  "
    CautionStatements(29, 0) = "BPY": CautionStatements(29, 1) = "BPY means boron-phosphorus-yttrium; Bpy means bipyridine.  "
    CautionStatements(30, 0) = "BPy": CautionStatements(30, 1) = "BPy means boron-pyridine; Bpy means bipyridine.  "
    CautionStatements(31, 0) = "Bpy": CautionStatements(31, 1) = "Bpy means bipyridine.  "
    CautionStatements(32, 0) = "Cys": CautionStatements(32, 1) = "Cys means cysteine; CYS means carbon-yttrium-sulfur.  "
    CautionStatements(33, 0) = "His": CautionStatements(33, 1) = "His means histidine; HIS means hydrogen-iodine-sulfur.  "
    CautionStatements(34, 0) = "Hoh": CautionStatements(34, 1) = "HoH means holmium-hydrogen; HOH means hydrogen-oxygen-hydrogen (aka water).  "
    CautionStatements(35, 0) = "Hyp": CautionStatements(35, 1) = "Hyp means hydroxyproline; HYP means hydrogen-yttrium-phosphorus.  "
    CautionStatements(36, 0) = "OAc": CautionStatements(36, 1) = "OAc means oxygen-actinium; Oac means acetate.  "
    CautionStatements(37, 0) = "Oac": CautionStatements(37, 1) = "Oac means acetate.  "
    CautionStatements(38, 0) = "Pro": CautionStatements(38, 1) = "Pro means proline; PrO means praseodymium-oxygen.  "
    CautionStatements(39, 0) = "PrO": CautionStatements(39, 1) = "Pro means proline; PrO means praseodymium-oxygen.  "
    CautionStatements(40, 0) = "Val": CautionStatements(40, 1) = "Val means valine; VAl means vanadium-aluminum.  "
    CautionStatements(41, 0) = "VAl": CautionStatements(41, 1) = "Val means valine; VAl means vanadium-aluminum.  "
    
End Sub

Public Sub MemoryLoadElements(eElementMode As emElementModeConstantsPrivate, Optional intSpecificElement As Integer = 0, Optional eSpecificStatToReset As esElementStatsConstants = esMass)
   ' intSpecificElement and intSpecificElementProperty are zero when updating all of the elements
   ' nonzero intSpecificElement and intSpecificElementProperty values will set just that specific value to the default
   ' eElementMode = 0 should not occur
   ' eElementMode = 1 means to use the average elemental weights
   ' eElementMode = 2 means to use isotopic elemental weights
   ' eElementMode = 3 means to use integer isotopic weights

    Const DEFAULT_CHARGE_CARRIER_MASS_AVG = 1.00739
    Const DEFAULT_CHARGE_CARRIER_MASS_MONOISO = 1.00727649

   ' This array stores the element names
   Dim strElementNames(ELEMENT_COUNT) As String
   
    ' dblElemVals(intElementIndex,1) stores the element's weight
    ' dblElemVals(intElementIndex,2) stores the element's uncertainty
    ' dblElemVals(intElementIndex,3) stores the element's charge
    ' Note: I could make this array of type udtElementStatsType, but the size of this sub would increase dramatically
   Dim dblElemVals(ELEMENT_COUNT, 3) As Double
   
   Dim intElementIndex As Integer, intIndex As Integer, intCompareIndex As Integer
   Dim strSwap As String

   ' Data Load Statements
        ' Uncertainties from CRC Handbook of Chemistry and Physics
        ' For Radioactive elements, the most stable isotope is NOT used;
        ' instead, an average Mol. Weight is used, just like with other elements.
        ' Data obtained from the Perma-Chart Science Series periodic table, 1993.
        ' Uncertainties from CRC Handoobk of Chemistry and Physics, except for
        ' Radioactive elements, where uncertainty was estimated to be .n5 where
        ' intSpecificElementProperty represents the number digits after the decimal point but before the last
        ' number of the molecular weight.
        ' For example, for No, MW = 259.1009 (±0.0005)
    
   ' Define the charge carrier mass
    If eElementMode = emAverageMass Then
        SetChargeCarrierMassInternal DEFAULT_CHARGE_CARRIER_MASS_AVG
    Else
        SetChargeCarrierMassInternal DEFAULT_CHARGE_CARRIER_MASS_MONOISO
    End If
   
   ' Assigning element names,        Charges
    strElementNames(1) = "H": dblElemVals(1, 3) = 1
    strElementNames(2) = "He": dblElemVals(2, 3) = 0
    strElementNames(3) = "Li": dblElemVals(3, 3) = 1
    strElementNames(4) = "Be": dblElemVals(4, 3) = 2
    strElementNames(5) = "B": dblElemVals(5, 3) = 3
    strElementNames(6) = "C": dblElemVals(6, 3) = 4
    strElementNames(7) = "N": dblElemVals(7, 3) = -3
    strElementNames(8) = "O": dblElemVals(8, 3) = -2
    strElementNames(9) = "F": dblElemVals(9, 3) = -1
    strElementNames(10) = "Ne": dblElemVals(10, 3) = 0
    strElementNames(11) = "Na": dblElemVals(11, 3) = 1
    strElementNames(12) = "Mg": dblElemVals(12, 3) = 2
    strElementNames(13) = "Al": dblElemVals(13, 3) = 3
    strElementNames(14) = "Si": dblElemVals(14, 3) = 4
    strElementNames(15) = "P": dblElemVals(15, 3) = -3
    strElementNames(16) = "S": dblElemVals(16, 3) = -2
    strElementNames(17) = "Cl": dblElemVals(17, 3) = -1
    strElementNames(18) = "Ar": dblElemVals(18, 3) = 0
    strElementNames(19) = "K": dblElemVals(19, 3) = 1
    strElementNames(20) = "Ca": dblElemVals(20, 3) = 2
    strElementNames(21) = "Sc": dblElemVals(21, 3) = 3
    strElementNames(22) = "Ti": dblElemVals(22, 3) = 4
    strElementNames(23) = "V": dblElemVals(23, 3) = 5
    strElementNames(24) = "Cr": dblElemVals(24, 3) = 3
    strElementNames(25) = "Mn": dblElemVals(25, 3) = 2
    strElementNames(26) = "Fe": dblElemVals(26, 3) = 3
    strElementNames(27) = "Co": dblElemVals(27, 3) = 2
    strElementNames(28) = "Ni": dblElemVals(28, 3) = 2
    strElementNames(29) = "Cu": dblElemVals(29, 3) = 2
    strElementNames(30) = "Zn": dblElemVals(30, 3) = 2
    strElementNames(31) = "Ga": dblElemVals(31, 3) = 3
    strElementNames(32) = "Ge": dblElemVals(32, 3) = 4
    strElementNames(33) = "As": dblElemVals(33, 3) = -3
    strElementNames(34) = "Se": dblElemVals(34, 3) = -2
    strElementNames(35) = "Br": dblElemVals(35, 3) = -1
    strElementNames(36) = "Kr": dblElemVals(36, 3) = 0
    strElementNames(37) = "Rb": dblElemVals(37, 3) = 1
    strElementNames(38) = "Sr": dblElemVals(38, 3) = 2
    strElementNames(39) = "Y": dblElemVals(39, 3) = 3
    strElementNames(40) = "Zr": dblElemVals(40, 3) = 4
    strElementNames(41) = "Nb": dblElemVals(41, 3) = 5
    strElementNames(42) = "Mo": dblElemVals(42, 3) = 6
    strElementNames(43) = "Tc": dblElemVals(43, 3) = 7
    strElementNames(44) = "Ru": dblElemVals(44, 3) = 4
    strElementNames(45) = "Rh": dblElemVals(45, 3) = 3
    strElementNames(46) = "Pd": dblElemVals(46, 3) = 2
    strElementNames(47) = "Ag": dblElemVals(47, 3) = 1
    strElementNames(48) = "Cd": dblElemVals(48, 3) = 2
    strElementNames(49) = "In": dblElemVals(49, 3) = 3
    strElementNames(50) = "Sn": dblElemVals(50, 3) = 4
    strElementNames(51) = "Sb": dblElemVals(51, 3) = -3
    strElementNames(52) = "Te": dblElemVals(52, 3) = -2
    strElementNames(53) = "I": dblElemVals(53, 3) = -1
    strElementNames(54) = "Xe": dblElemVals(54, 3) = 0
    strElementNames(55) = "Cs": dblElemVals(55, 3) = 1
    strElementNames(56) = "Ba": dblElemVals(56, 3) = 2
    strElementNames(57) = "La": dblElemVals(57, 3) = 3
    strElementNames(58) = "Ce": dblElemVals(58, 3) = 3
    strElementNames(59) = "Pr": dblElemVals(59, 3) = 4
    strElementNames(60) = "Nd": dblElemVals(60, 3) = 3
    strElementNames(61) = "Pm": dblElemVals(61, 3) = 3
    strElementNames(62) = "Sm": dblElemVals(62, 3) = 3
    strElementNames(63) = "Eu": dblElemVals(63, 3) = 3
    strElementNames(64) = "Gd": dblElemVals(64, 3) = 3
    strElementNames(65) = "Tb": dblElemVals(65, 3) = 3
    strElementNames(66) = "Dy": dblElemVals(66, 3) = 3
    strElementNames(67) = "Ho": dblElemVals(67, 3) = 3
    strElementNames(68) = "Er": dblElemVals(68, 3) = 3
    strElementNames(69) = "Tm": dblElemVals(69, 3) = 3
    strElementNames(70) = "Yb": dblElemVals(70, 3) = 3
    strElementNames(71) = "Lu": dblElemVals(71, 3) = 3
    strElementNames(72) = "Hf": dblElemVals(72, 3) = 4
    strElementNames(73) = "Ta": dblElemVals(73, 3) = 5
    strElementNames(74) = "W": dblElemVals(74, 3) = 6
    strElementNames(75) = "Re": dblElemVals(75, 3) = 7
    strElementNames(76) = "Os": dblElemVals(76, 3) = 4
    strElementNames(77) = "Ir": dblElemVals(77, 3) = 4
    strElementNames(78) = "Pt": dblElemVals(78, 3) = 4
    strElementNames(79) = "Au": dblElemVals(79, 3) = 3
    strElementNames(80) = "Hg": dblElemVals(80, 3) = 2
    strElementNames(81) = "Tl": dblElemVals(81, 3) = 1
    strElementNames(82) = "Pb": dblElemVals(82, 3) = 2
    strElementNames(83) = "Bi": dblElemVals(83, 3) = 3
    strElementNames(84) = "Po": dblElemVals(84, 3) = 4
    strElementNames(85) = "At": dblElemVals(85, 3) = -1
    strElementNames(86) = "Rn": dblElemVals(86, 3) = 0
    strElementNames(87) = "Fr": dblElemVals(87, 3) = 1
    strElementNames(88) = "Ra": dblElemVals(88, 3) = 2
    strElementNames(89) = "Ac": dblElemVals(89, 3) = 3
    strElementNames(90) = "Th": dblElemVals(90, 3) = 4
    strElementNames(91) = "Pa": dblElemVals(91, 3) = 5
    strElementNames(92) = "U": dblElemVals(92, 3) = 6
    strElementNames(93) = "Np": dblElemVals(93, 3) = 5
    strElementNames(94) = "Pu": dblElemVals(94, 3) = 4
    strElementNames(95) = "Am": dblElemVals(95, 3) = 3
    strElementNames(96) = "Cm": dblElemVals(96, 3) = 3
    strElementNames(97) = "Bk": dblElemVals(97, 3) = 3
    strElementNames(98) = "Cf": dblElemVals(98, 3) = 3
    strElementNames(99) = "Es": dblElemVals(99, 3) = 3
    strElementNames(100) = "Fm": dblElemVals(100, 3) = 3
    strElementNames(101) = "Md": dblElemVals(101, 3) = 3
    strElementNames(102) = "No": dblElemVals(102, 3) = 3
    strElementNames(103) = "Lr": dblElemVals(103, 3) = 3

    ' Set uncertainty to 0 for all elements if using exact isotopic or integer isotopic weights
   If eElementMode = 2 Or eElementMode = 3 Then
        For intIndex = 1 To ELEMENT_COUNT
            dblElemVals(intIndex, 2) = 0
        Next intIndex
   End If

   Select Case eElementMode
   Case 3
        ' Integer Isotopic Weights
        dblElemVals(1, 1) = 1
        dblElemVals(2, 1) = 4
        dblElemVals(3, 1) = 7
        dblElemVals(4, 1) = 9
        dblElemVals(5, 1) = 11
        dblElemVals(6, 1) = 12
        dblElemVals(7, 1) = 14
        dblElemVals(8, 1) = 16
        dblElemVals(9, 1) = 19
        dblElemVals(10, 1) = 20
        dblElemVals(11, 1) = 23
        dblElemVals(12, 1) = 24
        dblElemVals(13, 1) = 27
        dblElemVals(14, 1) = 28
        dblElemVals(15, 1) = 31
        dblElemVals(16, 1) = 32
        dblElemVals(17, 1) = 35
        dblElemVals(18, 1) = 40
        dblElemVals(19, 1) = 39
        dblElemVals(20, 1) = 40
        dblElemVals(21, 1) = 45
        dblElemVals(22, 1) = 48
        dblElemVals(23, 1) = 51
        dblElemVals(24, 1) = 52
        dblElemVals(25, 1) = 55
        dblElemVals(26, 1) = 56
        dblElemVals(27, 1) = 59
        dblElemVals(28, 1) = 58
        dblElemVals(29, 1) = 63
        dblElemVals(30, 1) = 64
        dblElemVals(31, 1) = 69
        dblElemVals(32, 1) = 72
        dblElemVals(33, 1) = 75
        dblElemVals(34, 1) = 80
        dblElemVals(35, 1) = 79
        dblElemVals(36, 1) = 84
        dblElemVals(37, 1) = 85
        dblElemVals(38, 1) = 88
        dblElemVals(39, 1) = 89
        dblElemVals(40, 1) = 90
        dblElemVals(41, 1) = 93
        dblElemVals(42, 1) = 98
        dblElemVals(43, 1) = 98
        dblElemVals(44, 1) = 102
        dblElemVals(45, 1) = 103
        dblElemVals(46, 1) = 106
        dblElemVals(47, 1) = 107
        dblElemVals(48, 1) = 114
        dblElemVals(49, 1) = 115
        dblElemVals(50, 1) = 120
        dblElemVals(51, 1) = 121
        dblElemVals(52, 1) = 130
        dblElemVals(53, 1) = 127
        dblElemVals(54, 1) = 132
        dblElemVals(55, 1) = 133
        dblElemVals(56, 1) = 138
        dblElemVals(57, 1) = 139
        dblElemVals(58, 1) = 140
        dblElemVals(59, 1) = 141
        dblElemVals(60, 1) = 142
        dblElemVals(61, 1) = 145
        dblElemVals(62, 1) = 152
        dblElemVals(63, 1) = 153
        dblElemVals(64, 1) = 158
        dblElemVals(65, 1) = 159
        dblElemVals(66, 1) = 164
        dblElemVals(67, 1) = 165
        dblElemVals(68, 1) = 166
        dblElemVals(69, 1) = 169
        dblElemVals(70, 1) = 174
        dblElemVals(71, 1) = 175
        dblElemVals(72, 1) = 180
        dblElemVals(73, 1) = 181
        dblElemVals(74, 1) = 184
        dblElemVals(75, 1) = 187
        dblElemVals(76, 1) = 192
        dblElemVals(77, 1) = 193
        dblElemVals(78, 1) = 195
        dblElemVals(79, 1) = 197
        dblElemVals(80, 1) = 202
        dblElemVals(81, 1) = 205
        dblElemVals(82, 1) = 208
        dblElemVals(83, 1) = 209
        dblElemVals(84, 1) = 209
        dblElemVals(85, 1) = 210
        dblElemVals(86, 1) = 222
        dblElemVals(87, 1) = 223
        dblElemVals(88, 1) = 227
        dblElemVals(89, 1) = 227
        dblElemVals(90, 1) = 232
        dblElemVals(91, 1) = 231
        dblElemVals(92, 1) = 238
        dblElemVals(93, 1) = 237
        dblElemVals(94, 1) = 244
        dblElemVals(95, 1) = 243
        dblElemVals(96, 1) = 247
        dblElemVals(97, 1) = 247
        dblElemVals(98, 1) = 251
        dblElemVals(99, 1) = 252
        dblElemVals(100, 1) = 257
        dblElemVals(101, 1) = 258
        dblElemVals(102, 1) = 269
        dblElemVals(103, 1) = 260
        
        ' Unused elements
        ' data 104,Unq,Unnilquadium,261.11,.05, 105,Unp,Unnilpentium,262.114,005, 106,Unh,Unnilhexium,263.118,.005, 107,Uns,Unnilseptium,262.12,.05

   Case 2
        ' isotopic Element Weights
        dblElemVals(1, 1) = 1.0078246
        dblElemVals(2, 1) = 4.0026029
        dblElemVals(3, 1) = 7.016005
        dblElemVals(4, 1) = 9.012183
        dblElemVals(5, 1) = 11.009305
        dblElemVals(6, 1) = 12
        dblElemVals(7, 1) = 14.003074
        dblElemVals(8, 1) = 15.994915
        dblElemVals(9, 1) = 18.9984032
        dblElemVals(10, 1) = 19.992439
        dblElemVals(11, 1) = 22.98977
        dblElemVals(12, 1) = 23.98505
        dblElemVals(13, 1) = 26.981541
        dblElemVals(14, 1) = 27.976928
        dblElemVals(15, 1) = 30.973763
        dblElemVals(16, 1) = 31.972072
        dblElemVals(17, 1) = 34.968853
        dblElemVals(18, 1) = 39.962383
        dblElemVals(19, 1) = 38.963708
        dblElemVals(20, 1) = 39.962591
        dblElemVals(21, 1) = 44.955914
        dblElemVals(22, 1) = 47.947947
        dblElemVals(23, 1) = 50.943963
        dblElemVals(24, 1) = 51.94051
        dblElemVals(25, 1) = 54.938046
        dblElemVals(26, 1) = 55.934939
        dblElemVals(27, 1) = 58.933198
        dblElemVals(28, 1) = 57.935347
        dblElemVals(29, 1) = 62.929599
        dblElemVals(30, 1) = 63.929145
        dblElemVals(31, 1) = 68.925581
        dblElemVals(32, 1) = 71.92208
        dblElemVals(33, 1) = 74.921596
        dblElemVals(34, 1) = 79.916521
        dblElemVals(35, 1) = 78.918336
        dblElemVals(36, 1) = 83.911506
        dblElemVals(37, 1) = 84.9118
        dblElemVals(38, 1) = 87.905625
        dblElemVals(39, 1) = 88.905856
        dblElemVals(40, 1) = 89.904708
        dblElemVals(41, 1) = 92.906378
        dblElemVals(42, 1) = 97.905405
        dblElemVals(43, 1) = 98
        dblElemVals(44, 1) = 101.90434
        dblElemVals(45, 1) = 102.905503
        dblElemVals(46, 1) = 105.903475
        dblElemVals(47, 1) = 106.905095
        dblElemVals(48, 1) = 113.903361
        dblElemVals(49, 1) = 114.903875
        dblElemVals(50, 1) = 119.902199
        dblElemVals(51, 1) = 120.903824
        dblElemVals(52, 1) = 129.906229
        dblElemVals(53, 1) = 126.904477
        dblElemVals(54, 1) = 131.904148
        dblElemVals(55, 1) = 132.905433
        dblElemVals(56, 1) = 137.905236
        dblElemVals(57, 1) = 138.906355
        dblElemVals(58, 1) = 139.905442
        dblElemVals(59, 1) = 140.907657
        dblElemVals(60, 1) = 141.907731
        dblElemVals(61, 1) = 145
        dblElemVals(62, 1) = 151.919741
        dblElemVals(63, 1) = 152.921243
        dblElemVals(64, 1) = 157.924111
        dblElemVals(65, 1) = 158.92535
        dblElemVals(66, 1) = 163.929183
        dblElemVals(67, 1) = 164.930332
        dblElemVals(68, 1) = 165.930305
        dblElemVals(69, 1) = 168.934225
        dblElemVals(70, 1) = 173.938873
        dblElemVals(71, 1) = 174.940785
        dblElemVals(72, 1) = 179.946561
        dblElemVals(73, 1) = 180.948014
        dblElemVals(74, 1) = 183.950953
        dblElemVals(75, 1) = 186.955765
        dblElemVals(76, 1) = 191.960603
        dblElemVals(77, 1) = 192.962942
        dblElemVals(78, 1) = 194.964785
        dblElemVals(79, 1) = 196.96656
        dblElemVals(80, 1) = 201.970632
        dblElemVals(81, 1) = 204.97441
        dblElemVals(82, 1) = 207.976641
        dblElemVals(83, 1) = 208.980388
        dblElemVals(84, 1) = 209
        dblElemVals(85, 1) = 210
        dblElemVals(86, 1) = 222
        dblElemVals(87, 1) = 223
        dblElemVals(88, 1) = 227
        dblElemVals(89, 1) = 227
        dblElemVals(90, 1) = 232.038054
        dblElemVals(91, 1) = 231
        dblElemVals(92, 1) = 238.050786
        dblElemVals(93, 1) = 237
        dblElemVals(94, 1) = 244
        dblElemVals(95, 1) = 243
        dblElemVals(96, 1) = 247
        dblElemVals(97, 1) = 247
        dblElemVals(98, 1) = 251
        dblElemVals(99, 1) = 252
        dblElemVals(100, 1) = 257
        dblElemVals(101, 1) = 258
        dblElemVals(102, 1) = 269
        dblElemVals(103, 1) = 260
        
        ' Unused elements
        ' data 104,Unq,Unnilquadium,261.11,.05, 105,Unp,Unnilpentium,262.114,005, 106,Unh,Unnilhexium,263.118,.005, 107,Uns,Unnilseptium,262.12,.05
    
    Case Else
        '                         Weight                           Uncertainty
        ' Average Element Weights
        dblElemVals(1, 1) = 1.00794: dblElemVals(1, 2) = 0.00007
        dblElemVals(2, 1) = 4.002602: dblElemVals(2, 2) = 0.000002
        dblElemVals(3, 1) = 6.941: dblElemVals(3, 2) = 0.002
        dblElemVals(4, 1) = 9.012182: dblElemVals(4, 2) = 0.000003
        dblElemVals(5, 1) = 10.811: dblElemVals(5, 2) = 0.007
        dblElemVals(6, 1) = 12.0107: dblElemVals(6, 2) = 0.0008
        dblElemVals(7, 1) = 14.00674: dblElemVals(7, 2) = 0.00007
        dblElemVals(8, 1) = 15.9994: dblElemVals(8, 2) = 0.0003
        dblElemVals(9, 1) = 18.9984032: dblElemVals(9, 2) = 0.0000005
        dblElemVals(10, 1) = 20.1797: dblElemVals(10, 2) = 0.0006
        dblElemVals(11, 1) = 22.98977: dblElemVals(11, 2) = 0.000002
        dblElemVals(12, 1) = 24.305: dblElemVals(12, 2) = 0.0006
        dblElemVals(13, 1) = 26.981538: dblElemVals(13, 2) = 0.000002
        dblElemVals(14, 1) = 28.0855: dblElemVals(14, 2) = 0.0003
        dblElemVals(15, 1) = 30.973761: dblElemVals(15, 2) = 0.000002
        dblElemVals(16, 1) = 32.066: dblElemVals(16, 2) = 0.006
        dblElemVals(17, 1) = 35.4527: dblElemVals(17, 2) = 0.0009
        dblElemVals(18, 1) = 39.948: dblElemVals(18, 2) = 0.001
        dblElemVals(19, 1) = 39.0983: dblElemVals(19, 2) = 0.0001
        dblElemVals(20, 1) = 40.078: dblElemVals(20, 2) = 0.004
        dblElemVals(21, 1) = 44.95591: dblElemVals(21, 2) = 0.000008
        dblElemVals(22, 1) = 47.867: dblElemVals(22, 2) = 0.001
        dblElemVals(23, 1) = 50.9415: dblElemVals(23, 2) = 0.0001
        dblElemVals(24, 1) = 51.9961: dblElemVals(24, 2) = 0.0006
        dblElemVals(25, 1) = 54.938049: dblElemVals(25, 2) = 0.000009
        dblElemVals(26, 1) = 55.845: dblElemVals(26, 2) = 0.002
        dblElemVals(27, 1) = 58.9332: dblElemVals(27, 2) = 0.000009
        dblElemVals(28, 1) = 58.6934: dblElemVals(28, 2) = 0.0002
        dblElemVals(29, 1) = 63.546: dblElemVals(29, 2) = 0.003
        dblElemVals(30, 1) = 65.39: dblElemVals(30, 2) = 0.02
        dblElemVals(31, 1) = 69.723: dblElemVals(31, 2) = 0.001
        dblElemVals(32, 1) = 72.61: dblElemVals(32, 2) = 0.02
        dblElemVals(33, 1) = 74.9216: dblElemVals(33, 2) = 0.00002
        dblElemVals(34, 1) = 78.96: dblElemVals(34, 2) = 0.03
        dblElemVals(35, 1) = 79.904: dblElemVals(35, 2) = 0.001
        dblElemVals(36, 1) = 83.8: dblElemVals(36, 2) = 0.01
        dblElemVals(37, 1) = 85.4678: dblElemVals(37, 2) = 0.0003
        dblElemVals(38, 1) = 87.62: dblElemVals(38, 2) = 0.01
        dblElemVals(39, 1) = 88.90585: dblElemVals(39, 2) = 0.00002
        dblElemVals(40, 1) = 91.224: dblElemVals(40, 2) = 0.002
        dblElemVals(41, 1) = 92.90638: dblElemVals(41, 2) = 0.00002
        dblElemVals(42, 1) = 95.94: dblElemVals(42, 2) = 0.01
        dblElemVals(43, 1) = 97.9072: dblElemVals(43, 2) = 0.0005
        dblElemVals(44, 1) = 101.07: dblElemVals(44, 2) = 0.02
        dblElemVals(45, 1) = 102.9055: dblElemVals(45, 2) = 0.00002
        dblElemVals(46, 1) = 106.42: dblElemVals(46, 2) = 0.01
        dblElemVals(47, 1) = 107.8682: dblElemVals(47, 2) = 0.0002
        dblElemVals(48, 1) = 112.411: dblElemVals(48, 2) = 0.008
        dblElemVals(49, 1) = 114.818: dblElemVals(49, 2) = 0.003
        dblElemVals(50, 1) = 118.71: dblElemVals(50, 2) = 0.007
        dblElemVals(51, 1) = 121.76: dblElemVals(51, 2) = 0.001
        dblElemVals(52, 1) = 127.6: dblElemVals(52, 2) = 0.03
        dblElemVals(53, 1) = 126.90447: dblElemVals(53, 2) = 0.00003
        dblElemVals(54, 1) = 131.29: dblElemVals(54, 2) = 0.02
        dblElemVals(55, 1) = 132.90545: dblElemVals(55, 2) = 0.00002
        dblElemVals(56, 1) = 137.327: dblElemVals(56, 2) = 0.007
        dblElemVals(57, 1) = 138.9055: dblElemVals(57, 2) = 0.0002
        dblElemVals(58, 1) = 140.116: dblElemVals(58, 2) = 0.001
        dblElemVals(59, 1) = 140.90765: dblElemVals(59, 2) = 0.00002
        dblElemVals(60, 1) = 144.24: dblElemVals(60, 2) = 0.03
        dblElemVals(61, 1) = 144.9127: dblElemVals(61, 2) = 0.0005
        dblElemVals(62, 1) = 150.36: dblElemVals(62, 2) = 0.03
        dblElemVals(63, 1) = 151.964: dblElemVals(63, 2) = 0.001
        dblElemVals(64, 1) = 157.25: dblElemVals(64, 2) = 0.03
        dblElemVals(65, 1) = 158.92534: dblElemVals(65, 2) = 0.00002
        dblElemVals(66, 1) = 162.5: dblElemVals(66, 2) = 0.03
        dblElemVals(67, 1) = 164.93032: dblElemVals(67, 2) = 0.00002
        dblElemVals(68, 1) = 167.26: dblElemVals(68, 2) = 0.03
        dblElemVals(69, 1) = 168.93421: dblElemVals(69, 2) = 0.00002
        dblElemVals(70, 1) = 173.04: dblElemVals(70, 2) = 0.03
        dblElemVals(71, 1) = 174.967: dblElemVals(71, 2) = 0.001
        dblElemVals(72, 1) = 178.49: dblElemVals(72, 2) = 0.02
        dblElemVals(73, 1) = 180.9479: dblElemVals(73, 2) = 0.0001
        dblElemVals(74, 1) = 183.84: dblElemVals(74, 2) = 0.01
        dblElemVals(75, 1) = 186.207: dblElemVals(75, 2) = 0.001
        dblElemVals(76, 1) = 190.23: dblElemVals(76, 2) = 0.03
        dblElemVals(77, 1) = 192.217: dblElemVals(77, 2) = 0.03
        dblElemVals(78, 1) = 195.078: dblElemVals(78, 2) = 0.002
        dblElemVals(79, 1) = 196.96655: dblElemVals(79, 2) = 0.00002
        dblElemVals(80, 1) = 200.59: dblElemVals(80, 2) = 0.02
        dblElemVals(81, 1) = 204.3833: dblElemVals(81, 2) = 0.0002
        dblElemVals(82, 1) = 207.2: dblElemVals(82, 2) = 0.1
        dblElemVals(83, 1) = 208.98038: dblElemVals(83, 2) = 0.00002
        dblElemVals(84, 1) = 208.9824: dblElemVals(84, 2) = 0.0005
        dblElemVals(85, 1) = 209.9871: dblElemVals(85, 2) = 0.0005
        dblElemVals(86, 1) = 222.0176: dblElemVals(86, 2) = 0.0005
        dblElemVals(87, 1) = 223.0197: dblElemVals(87, 2) = 0.0005
        dblElemVals(88, 1) = 226.0254: dblElemVals(88, 2) = 0.0001
        dblElemVals(89, 1) = 227.0278: dblElemVals(89, 2) = 0.00001
        dblElemVals(90, 1) = 232.0381: dblElemVals(90, 2) = 0.0001
        dblElemVals(91, 1) = 231.03588: dblElemVals(91, 2) = 0.00002
        dblElemVals(92, 1) = 238.0289: dblElemVals(92, 2) = 0.0001
        dblElemVals(93, 1) = 237.0482: dblElemVals(93, 2) = 0.0005
        dblElemVals(94, 1) = 244.0642: dblElemVals(94, 2) = 0.0005
        dblElemVals(95, 1) = 243.0614: dblElemVals(95, 2) = 0.0005
        dblElemVals(96, 1) = 247.0703: dblElemVals(96, 2) = 0.0005
        dblElemVals(97, 1) = 247.0703: dblElemVals(97, 2) = 0.0005
        dblElemVals(98, 1) = 251.0796: dblElemVals(98, 2) = 0.0005
        dblElemVals(99, 1) = 252.083: dblElemVals(99, 2) = 0.005
        dblElemVals(100, 1) = 257.0951: dblElemVals(100, 2) = 0.0005
        dblElemVals(101, 1) = 258.1: dblElemVals(101, 2) = 0.05
        dblElemVals(102, 1) = 259.1009: dblElemVals(102, 2) = 0.0005
        dblElemVals(103, 1) = 262.11: dblElemVals(103, 2) = 0.05
    
        ' Unused elements
        ' data 104,Unq,Unnilquadium,261,1, 105,Unp,Unnilpentium,262,1, 106,Unh,Unnilhexium,263,1
    End Select
    
    If intSpecificElement = 0 Then
        ' Updating all the elements
        For intElementIndex = 1 To ELEMENT_COUNT
            With ElementStats(intElementIndex)
                .Symbol = strElementNames(intElementIndex)
                .Mass = dblElemVals(intElementIndex, 1)
                .Uncertainty = dblElemVals(intElementIndex, 2)
                .Charge = dblElemVals(intElementIndex, 3)
                
                ElementAlph(intElementIndex) = .Symbol
            End With
        Next intElementIndex
        
        ' Alphabatize ElementAlph() array via bubble sort
        For intCompareIndex = ELEMENT_COUNT To 2 Step -1       ' Sort from end to start
            For intIndex = 1 To intCompareIndex - 1
                If ElementAlph(intIndex) > ElementAlph(intIndex + 1) Then
                    ' Swap them
                    strSwap = ElementAlph(intIndex)
                    ElementAlph(intIndex) = ElementAlph(intIndex + 1)
                    ElementAlph(intIndex + 1) = strSwap
                End If
            Next intIndex
        Next intCompareIndex
        
    Else
        If intSpecificElement >= 1 And intSpecificElement <= ELEMENT_COUNT Then
            With ElementStats(intSpecificElement)
                Select Case eSpecificStatToReset
                Case esMass
                    .Mass = dblElemVals(intSpecificElement, 1)
                Case esUncertainty
                    .Uncertainty = dblElemVals(intSpecificElement, 2)
                Case esCharge
                    .Charge = dblElemVals(intSpecificElement, 3)
                Case Else
                    ' Ignore it
                End Select
            End With
        End If
    End If

End Sub

Private Sub MemoryLoadIsotopes()
    ' Stores isotope information in ElementStats()
    
    Dim intElementIndex As Integer, intIsotopeindex As Integer
    
    ' The dblIsoMasses() array holds the mass of each isotope
    ' starting with dblIsoMasses(x,1), dblIsoMasses(x, 2), etc.
    Dim dblIsoMasses(ELEMENT_COUNT, MAX_ISOTOPES) As Double
    
    ' The sngIsoAbun() array holds the isotopic abundances of each of the isotopes,
    ' starting with sngIsoAbun(x,1) and corresponding to dblIsoMasses()
    Dim sngIsoAbun(ELEMENT_COUNT, MAX_ISOTOPES) As Single
    
    dblIsoMasses(1, 1) = 1.0078246: sngIsoAbun(1, 1) = 0.99985
    dblIsoMasses(1, 2) = 2.014: sngIsoAbun(1, 2) = 0.00015
    dblIsoMasses(2, 1) = 3.01603: sngIsoAbun(2, 1) = 0.00000137
    dblIsoMasses(2, 2) = 4.0026029: sngIsoAbun(2, 2) = 0.99999863
    dblIsoMasses(3, 1) = 6.01512: sngIsoAbun(3, 1) = 0.0759
    dblIsoMasses(3, 2) = 7.016005: sngIsoAbun(3, 2) = 0.9241
    dblIsoMasses(4, 1) = 9.012183: sngIsoAbun(4, 1) = 1
    dblIsoMasses(5, 1) = 10.0129: sngIsoAbun(5, 1) = 0.199
    dblIsoMasses(5, 2) = 11.009305: sngIsoAbun(5, 2) = 0.801
    dblIsoMasses(6, 1) = 12: sngIsoAbun(6, 1) = 0.9893
    dblIsoMasses(6, 2) = 13.00335: sngIsoAbun(6, 2) = 0.0107
    dblIsoMasses(7, 1) = 14.003074: sngIsoAbun(7, 1) = 0.99632
    dblIsoMasses(7, 2) = 15.00011: sngIsoAbun(7, 2) = 0.00368
    dblIsoMasses(8, 1) = 15.994915: sngIsoAbun(8, 1) = 0.99757
    dblIsoMasses(8, 2) = 16.999131: sngIsoAbun(8, 2) = 0.00038
    dblIsoMasses(8, 3) = 17.99916: sngIsoAbun(8, 3) = 0.00205
    dblIsoMasses(9, 1) = 18.9984032: sngIsoAbun(9, 1) = 1
    dblIsoMasses(10, 1) = 19.992439: sngIsoAbun(10, 1) = 0.9048
    dblIsoMasses(10, 2) = 20.99395: sngIsoAbun(10, 2) = 0.0027
    dblIsoMasses(10, 3) = 21.99138: sngIsoAbun(10, 3) = 0.0925
    dblIsoMasses(11, 1) = 22.98977: sngIsoAbun(11, 1) = 1
    dblIsoMasses(12, 1) = 23.98505: sngIsoAbun(12, 1) = 0.7899
    dblIsoMasses(12, 2) = 24.98584: sngIsoAbun(12, 2) = 0.1
    dblIsoMasses(12, 3) = 25.98259: sngIsoAbun(12, 3) = 0.1101
    dblIsoMasses(13, 1) = 26.981541: sngIsoAbun(13, 1) = 1
    dblIsoMasses(14, 1) = 27.976928: sngIsoAbun(14, 1) = 0.922297
    dblIsoMasses(14, 2) = 28.97649: sngIsoAbun(14, 2) = 0.046832
    dblIsoMasses(14, 3) = 29.97376: sngIsoAbun(14, 3) = 0.030871
    dblIsoMasses(15, 1) = 30.973763: sngIsoAbun(15, 1) = 1
    dblIsoMasses(16, 1) = 31.972072: sngIsoAbun(16, 1) = 0.9493
    dblIsoMasses(16, 2) = 32.97146: sngIsoAbun(16, 2) = 0.0076
    dblIsoMasses(16, 3) = 33.96786: sngIsoAbun(16, 3) = 0.0429
    dblIsoMasses(16, 4) = 35.96709: sngIsoAbun(16, 4) = 0.0002
    dblIsoMasses(17, 1) = 34.968853: sngIsoAbun(17, 1) = 0.7578
    dblIsoMasses(17, 2) = 36.99999: sngIsoAbun(17, 2) = 0.2422
    dblIsoMasses(18, 1) = 35.96755: sngIsoAbun(18, 1) = 0.003365
    dblIsoMasses(18, 2) = 37.96272: sngIsoAbun(18, 2) = 0.000632
    dblIsoMasses(18, 3) = 39.96999: sngIsoAbun(18, 3) = 0.996003        ' Note: Alternate mass is 39.962383
    dblIsoMasses(19, 1) = 38.963708: sngIsoAbun(19, 1) = 0.932581
    dblIsoMasses(19, 2) = 39.963999: sngIsoAbun(19, 2) = 0.000117
    dblIsoMasses(19, 3) = 40.961825: sngIsoAbun(19, 3) = 0.067302
    dblIsoMasses(20, 1) = 39.962591: sngIsoAbun(20, 1) = 0.96941
    dblIsoMasses(20, 2) = 41.958618: sngIsoAbun(20, 2) = 0.00647
    dblIsoMasses(20, 3) = 42.958766: sngIsoAbun(20, 3) = 0.00135
    dblIsoMasses(20, 4) = 43.95548: sngIsoAbun(20, 4) = 0.02086
    dblIsoMasses(20, 5) = 45.953689: sngIsoAbun(20, 5) = 0.00004
    dblIsoMasses(20, 6) = 47.952533: sngIsoAbun(20, 6) = 0.00187
    dblIsoMasses(21, 1) = 44.959404: sngIsoAbun(21, 1) = 1             ' Note: Alternate mass is 44.955914
    dblIsoMasses(22, 1) = 45.952629: sngIsoAbun(22, 1) = 0.0825
    dblIsoMasses(22, 2) = 46.951764: sngIsoAbun(22, 2) = 0.0744
    dblIsoMasses(22, 3) = 47.947947: sngIsoAbun(22, 3) = 0.7372
    dblIsoMasses(22, 4) = 48.947871: sngIsoAbun(22, 4) = 0.0541
    dblIsoMasses(22, 5) = 49.944792: sngIsoAbun(22, 5) = 0.0518
    dblIsoMasses(23, 1) = 49.947161: sngIsoAbun(23, 1) = 0.0025
    dblIsoMasses(23, 2) = 50.943963: sngIsoAbun(23, 2) = 0.9975
    dblIsoMasses(24, 1) = 49.946046: sngIsoAbun(24, 1) = 0.04345
    dblIsoMasses(24, 2) = 51.940509: sngIsoAbun(24, 2) = 0.83789
    dblIsoMasses(24, 3) = 52.940651: sngIsoAbun(24, 3) = 0.09501
    dblIsoMasses(24, 4) = 53.938882: sngIsoAbun(24, 4) = 0.02365
    dblIsoMasses(25, 1) = 54.938046: sngIsoAbun(25, 1) = 1
    dblIsoMasses(26, 1) = 53.939612: sngIsoAbun(26, 1) = 0.05845
    dblIsoMasses(26, 2) = 55.934939: sngIsoAbun(26, 2) = 0.91754
    dblIsoMasses(26, 3) = 56.935396: sngIsoAbun(26, 3) = 0.02119
    dblIsoMasses(26, 4) = 57.933277: sngIsoAbun(26, 4) = 0.00282
    dblIsoMasses(27, 1) = 58.933198: sngIsoAbun(27, 1) = 1
    dblIsoMasses(28, 1) = 57.935347: sngIsoAbun(28, 1) = 0.680769
    dblIsoMasses(28, 2) = 59.930788: sngIsoAbun(28, 2) = 0.262231
    dblIsoMasses(28, 3) = 60.931058: sngIsoAbun(28, 3) = 0.011399
    dblIsoMasses(28, 4) = 61.928346: sngIsoAbun(28, 4) = 0.036345
    dblIsoMasses(28, 5) = 63.927968: sngIsoAbun(28, 5) = 0.009256
    dblIsoMasses(29, 1) = 62.939598: sngIsoAbun(29, 1) = 0.6917         ' Note: Alternate mass is 62.929599
    dblIsoMasses(29, 2) = 64.927793: sngIsoAbun(29, 2) = 0.3083
    dblIsoMasses(30, 1) = 63.929145: sngIsoAbun(30, 1) = 0.4863
    dblIsoMasses(30, 2) = 65.926034: sngIsoAbun(30, 2) = 0.279
    dblIsoMasses(30, 3) = 66.927129: sngIsoAbun(30, 3) = 0.041
    dblIsoMasses(30, 4) = 67.924846: sngIsoAbun(30, 4) = 0.1875
    dblIsoMasses(30, 5) = 69.925325: sngIsoAbun(30, 5) = 0.0062
    dblIsoMasses(31, 1) = 68.925581: sngIsoAbun(31, 1) = 0.60108
    dblIsoMasses(31, 2) = 70.9247: sngIsoAbun(31, 2) = 0.39892
    dblIsoMasses(32, 1) = 69.92425: sngIsoAbun(32, 1) = 0.2084
    dblIsoMasses(32, 2) = 71.922079: sngIsoAbun(32, 2) = 0.2754
    dblIsoMasses(32, 3) = 72.923463: sngIsoAbun(32, 3) = 0.0773
    dblIsoMasses(32, 4) = 73.921177: sngIsoAbun(32, 4) = 0.3628
    dblIsoMasses(32, 5) = 75.921401: sngIsoAbun(32, 5) = 0.0761
    dblIsoMasses(33, 1) = 74.921596: sngIsoAbun(33, 1) = 1
    dblIsoMasses(34, 1) = 73.922475: sngIsoAbun(34, 1) = 0.0089
    dblIsoMasses(34, 2) = 75.919212: sngIsoAbun(34, 2) = 0.0937
    dblIsoMasses(34, 3) = 76.919912: sngIsoAbun(34, 3) = 0.0763
    dblIsoMasses(34, 4) = 77.919: sngIsoAbun(34, 4) = 0.2377
    dblIsoMasses(34, 5) = 79.916521: sngIsoAbun(34, 5) = 0.4961
    dblIsoMasses(34, 6) = 81.916698: sngIsoAbun(34, 6) = 0.0873
    dblIsoMasses(35, 1) = 78.918336: sngIsoAbun(35, 1) = 0.5069
    dblIsoMasses(35, 2) = 80.916289: sngIsoAbun(35, 2) = 0.4931
    dblIsoMasses(36, 1) = 77.92: sngIsoAbun(36, 1) = 0.0035
    dblIsoMasses(36, 2) = 79.91638: sngIsoAbun(36, 2) = 0.0228
    dblIsoMasses(36, 3) = 81.913482: sngIsoAbun(36, 3) = 0.1158
    dblIsoMasses(36, 4) = 82.914135: sngIsoAbun(36, 4) = 0.1149
    dblIsoMasses(36, 5) = 83.911506: sngIsoAbun(36, 5) = 0.57
    dblIsoMasses(36, 6) = 85.910616: sngIsoAbun(36, 6) = 0.173
    dblIsoMasses(37, 1) = 84.911794: sngIsoAbun(37, 1) = 0.7217
    dblIsoMasses(37, 2) = 86.909187: sngIsoAbun(37, 2) = 0.2783
    dblIsoMasses(38, 1) = 83.91343: sngIsoAbun(38, 1) = 0.0056
    dblIsoMasses(38, 2) = 85.909267: sngIsoAbun(38, 2) = 0.0986
    dblIsoMasses(38, 3) = 86.908884: sngIsoAbun(38, 3) = 0.07
    dblIsoMasses(38, 4) = 87.905625: sngIsoAbun(38, 4) = 0.8258
    dblIsoMasses(39, 1) = 88.905856: sngIsoAbun(39, 1) = 1
    dblIsoMasses(40, 1) = 89.904708: sngIsoAbun(40, 1) = 0.5145
    dblIsoMasses(40, 2) = 90.905644: sngIsoAbun(40, 2) = 0.1122
    dblIsoMasses(40, 3) = 91.905039: sngIsoAbun(40, 3) = 0.1715
    dblIsoMasses(40, 4) = 93.906314: sngIsoAbun(40, 4) = 0.1738
    dblIsoMasses(40, 5) = 95.908275: sngIsoAbun(40, 5) = 0.028
    dblIsoMasses(41, 1) = 92.906378: sngIsoAbun(41, 1) = 1
    dblIsoMasses(42, 1) = 91.906808: sngIsoAbun(42, 1) = 0.1484
    dblIsoMasses(42, 2) = 93.905085: sngIsoAbun(42, 2) = 0.0925
    dblIsoMasses(42, 3) = 94.90584: sngIsoAbun(42, 3) = 0.1592
    dblIsoMasses(42, 4) = 95.904678: sngIsoAbun(42, 4) = 0.1668
    dblIsoMasses(42, 5) = 96.90602: sngIsoAbun(42, 5) = 0.0955
    dblIsoMasses(42, 6) = 97.905405: sngIsoAbun(42, 6) = 0.2413
    dblIsoMasses(42, 7) = 99.907477: sngIsoAbun(42, 7) = 0.0963
    dblIsoMasses(43, 1) = 97.9072: sngIsoAbun(43, 1) = 1
    dblIsoMasses(44, 1) = 95.907599: sngIsoAbun(44, 1) = 0.0554
    dblIsoMasses(44, 2) = 97.905287: sngIsoAbun(44, 2) = 0.0187
    dblIsoMasses(44, 3) = 98.905939: sngIsoAbun(44, 3) = 0.1276
    dblIsoMasses(44, 4) = 99.904219: sngIsoAbun(44, 4) = 0.126
    dblIsoMasses(44, 5) = 100.905582: sngIsoAbun(44, 5) = 0.1706
    dblIsoMasses(44, 6) = 101.904348: sngIsoAbun(44, 6) = 0.3155
    dblIsoMasses(44, 7) = 103.905424: sngIsoAbun(44, 7) = 0.1862
    dblIsoMasses(45, 1) = 102.905503: sngIsoAbun(45, 1) = 1
    dblIsoMasses(46, 1) = 101.905634: sngIsoAbun(46, 1) = 0.0102
    dblIsoMasses(46, 2) = 103.904029: sngIsoAbun(46, 2) = 0.1114
    dblIsoMasses(46, 3) = 104.905079: sngIsoAbun(46, 3) = 0.2233
    dblIsoMasses(46, 4) = 105.903475: sngIsoAbun(46, 4) = 0.2733
    dblIsoMasses(46, 5) = 107.903895: sngIsoAbun(46, 5) = 0.2646
    dblIsoMasses(46, 6) = 109.905167: sngIsoAbun(46, 6) = 0.1172
    dblIsoMasses(47, 1) = 106.905095: sngIsoAbun(47, 1) = 0.51839
    dblIsoMasses(47, 2) = 108.904757: sngIsoAbun(47, 2) = 0.48161
    dblIsoMasses(48, 1) = 105.906461: sngIsoAbun(48, 1) = 0.0125
    dblIsoMasses(48, 2) = 107.904176: sngIsoAbun(48, 2) = 0.0089
    dblIsoMasses(48, 3) = 109.903005: sngIsoAbun(48, 3) = 0.1249
    dblIsoMasses(48, 4) = 110.904182: sngIsoAbun(48, 4) = 0.128
    dblIsoMasses(48, 5) = 111.902758: sngIsoAbun(48, 5) = 0.2413
    dblIsoMasses(48, 6) = 112.9044: sngIsoAbun(48, 6) = 0.1222
    dblIsoMasses(48, 7) = 113.903361: sngIsoAbun(48, 7) = 0.2873
    dblIsoMasses(48, 8) = 115.904754: sngIsoAbun(48, 8) = 0.0749
    dblIsoMasses(49, 1) = 112.904061: sngIsoAbun(49, 1) = 0.0429
    dblIsoMasses(49, 2) = 114.903875: sngIsoAbun(49, 2) = 0.9571
    dblIsoMasses(50, 1) = 111.904826: sngIsoAbun(50, 1) = 0.0097
    dblIsoMasses(50, 2) = 113.902784: sngIsoAbun(50, 2) = 0.0066
    dblIsoMasses(50, 3) = 114.903348: sngIsoAbun(50, 3) = 0.0034
    dblIsoMasses(50, 4) = 115.901747: sngIsoAbun(50, 4) = 0.1454
    dblIsoMasses(50, 5) = 116.902956: sngIsoAbun(50, 5) = 0.0768
    dblIsoMasses(50, 6) = 117.901609: sngIsoAbun(50, 6) = 0.2422
    dblIsoMasses(50, 7) = 118.90331: sngIsoAbun(50, 7) = 0.0859
    dblIsoMasses(50, 8) = 119.902199: sngIsoAbun(50, 8) = 0.3258
    dblIsoMasses(50, 9) = 121.90344: sngIsoAbun(50, 9) = 0.0463
    dblIsoMasses(50, 10) = 123.905274: sngIsoAbun(50, 10) = 0.0579
    dblIsoMasses(51, 1) = 120.903824: sngIsoAbun(51, 1) = 0.5721
    dblIsoMasses(51, 2) = 122.904216: sngIsoAbun(51, 2) = 0.4279
    dblIsoMasses(52, 1) = 119.904048: sngIsoAbun(52, 1) = 0.0009
    dblIsoMasses(52, 2) = 121.903054: sngIsoAbun(52, 2) = 0.0255
    dblIsoMasses(52, 3) = 122.904271: sngIsoAbun(52, 3) = 0.0089
    dblIsoMasses(52, 4) = 123.902823: sngIsoAbun(52, 4) = 0.0474
    dblIsoMasses(52, 5) = 124.904433: sngIsoAbun(52, 5) = 0.0707
    dblIsoMasses(52, 6) = 125.903314: sngIsoAbun(52, 6) = 0.1884
    dblIsoMasses(52, 7) = 127.904463: sngIsoAbun(52, 7) = 0.3174
    dblIsoMasses(52, 8) = 129.906229: sngIsoAbun(52, 8) = 0.3408
    dblIsoMasses(53, 1) = 126.904477: sngIsoAbun(53, 1) = 1
    dblIsoMasses(54, 1) = 123.905894: sngIsoAbun(54, 1) = 0.0009
    dblIsoMasses(54, 2) = 125.904281: sngIsoAbun(54, 2) = 0.0009
    dblIsoMasses(54, 3) = 127.903531: sngIsoAbun(54, 3) = 0.0192
    dblIsoMasses(54, 4) = 128.90478: sngIsoAbun(54, 4) = 0.2644
    dblIsoMasses(54, 5) = 129.903509: sngIsoAbun(54, 5) = 0.0408
    dblIsoMasses(54, 6) = 130.905072: sngIsoAbun(54, 6) = 0.2118
    dblIsoMasses(54, 7) = 131.904148: sngIsoAbun(54, 7) = 0.2689
    dblIsoMasses(54, 8) = 133.905395: sngIsoAbun(54, 8) = 0.1044
    dblIsoMasses(54, 9) = 135.907214: sngIsoAbun(54, 9) = 0.0887
    dblIsoMasses(55, 1) = 132.905433: sngIsoAbun(55, 1) = 1
    dblIsoMasses(56, 1) = 129.906282: sngIsoAbun(56, 1) = 0.00106
    dblIsoMasses(56, 2) = 131.905042: sngIsoAbun(56, 2) = 0.00101
    dblIsoMasses(56, 3) = 133.904486: sngIsoAbun(56, 3) = 0.02417
    dblIsoMasses(56, 4) = 134.905665: sngIsoAbun(56, 4) = 0.06592
    dblIsoMasses(56, 5) = 135.904553: sngIsoAbun(56, 5) = 0.07854
    dblIsoMasses(56, 6) = 136.905812: sngIsoAbun(56, 6) = 0.11232
    dblIsoMasses(56, 7) = 137.905236: sngIsoAbun(56, 7) = 0.71698
    dblIsoMasses(57, 1) = 137.907105: sngIsoAbun(57, 1) = 0.0009
    dblIsoMasses(57, 2) = 138.906355: sngIsoAbun(57, 2) = 0.9991
    dblIsoMasses(58, 1) = 135.90714: sngIsoAbun(58, 1) = 0.00185
    dblIsoMasses(58, 2) = 137.905985: sngIsoAbun(58, 2) = 0.00251
    dblIsoMasses(58, 3) = 139.905442: sngIsoAbun(58, 3) = 0.8845
    dblIsoMasses(58, 4) = 141.909241: sngIsoAbun(58, 4) = 0.11114
    dblIsoMasses(59, 1) = 140.907657: sngIsoAbun(59, 1) = 1
    dblIsoMasses(60, 1) = 141.907731: sngIsoAbun(60, 1) = 0.272
    dblIsoMasses(60, 2) = 142.90981: sngIsoAbun(60, 2) = 0.122
    dblIsoMasses(60, 3) = 143.910083: sngIsoAbun(60, 3) = 0.238
    dblIsoMasses(60, 4) = 144.91257: sngIsoAbun(60, 4) = 0.083
    dblIsoMasses(60, 5) = 145.913113: sngIsoAbun(60, 5) = 0.172
    dblIsoMasses(60, 6) = 147.916889: sngIsoAbun(60, 6) = 0.057
    dblIsoMasses(60, 7) = 149.920887: sngIsoAbun(60, 7) = 0.056
    dblIsoMasses(61, 1) = 144.9127: sngIsoAbun(61, 1) = 1
    dblIsoMasses(62, 1) = 143.911998: sngIsoAbun(62, 1) = 0.0307
    dblIsoMasses(62, 2) = 146.914895: sngIsoAbun(62, 2) = 0.1499
    dblIsoMasses(62, 3) = 147.91482: sngIsoAbun(62, 3) = 0.1124
    dblIsoMasses(62, 4) = 148.917181: sngIsoAbun(62, 4) = 0.1382
    dblIsoMasses(62, 5) = 149.917273: sngIsoAbun(62, 5) = 0.0738
    dblIsoMasses(62, 6) = 151.919741: sngIsoAbun(62, 6) = 0.2675
    dblIsoMasses(62, 7) = 153.922206: sngIsoAbun(62, 7) = 0.2275
    dblIsoMasses(63, 1) = 150.919847: sngIsoAbun(63, 1) = 0.4781
    dblIsoMasses(63, 2) = 152.921243: sngIsoAbun(63, 2) = 0.5219
    dblIsoMasses(64, 1) = 151.919786: sngIsoAbun(64, 1) = 0.002
    dblIsoMasses(64, 2) = 153.920861: sngIsoAbun(64, 2) = 0.0218
    dblIsoMasses(64, 3) = 154.922618: sngIsoAbun(64, 3) = 0.148
    dblIsoMasses(64, 4) = 155.922118: sngIsoAbun(64, 4) = 0.2047
    dblIsoMasses(64, 5) = 156.923956: sngIsoAbun(64, 5) = 0.1565
    dblIsoMasses(64, 6) = 157.924111: sngIsoAbun(64, 6) = 0.2484
    dblIsoMasses(64, 7) = 159.927049: sngIsoAbun(64, 7) = 0.2186
    dblIsoMasses(65, 1) = 158.92535: sngIsoAbun(65, 1) = 1
    dblIsoMasses(66, 1) = 155.925277: sngIsoAbun(66, 1) = 0.0006
    dblIsoMasses(66, 2) = 157.924403: sngIsoAbun(66, 2) = 0.001
    dblIsoMasses(66, 3) = 159.925193: sngIsoAbun(66, 3) = 0.0234
    dblIsoMasses(66, 4) = 160.92693: sngIsoAbun(66, 4) = 0.1891
    dblIsoMasses(66, 5) = 161.926795: sngIsoAbun(66, 5) = 0.2551
    dblIsoMasses(66, 6) = 162.928728: sngIsoAbun(66, 6) = 0.249
    dblIsoMasses(66, 7) = 163.929183: sngIsoAbun(66, 7) = 0.2818
    dblIsoMasses(67, 1) = 164.930332: sngIsoAbun(67, 1) = 1
    dblIsoMasses(68, 1) = 161.928775: sngIsoAbun(68, 1) = 0.0014
    dblIsoMasses(68, 2) = 163.929198: sngIsoAbun(68, 2) = 0.0161
    dblIsoMasses(68, 3) = 165.930305: sngIsoAbun(68, 3) = 0.3361
    dblIsoMasses(68, 4) = 166.932046: sngIsoAbun(68, 4) = 0.2293
    dblIsoMasses(68, 5) = 167.932368: sngIsoAbun(68, 5) = 0.2678
    dblIsoMasses(68, 6) = 169.935461: sngIsoAbun(68, 6) = 0.1493
    dblIsoMasses(69, 1) = 168.934225: sngIsoAbun(69, 1) = 1
    dblIsoMasses(70, 1) = 167.932873: sngIsoAbun(70, 1) = 0.0013
    dblIsoMasses(70, 2) = 169.934759: sngIsoAbun(70, 2) = 0.0304
    dblIsoMasses(70, 3) = 170.936323: sngIsoAbun(70, 3) = 0.1428
    dblIsoMasses(70, 4) = 171.936387: sngIsoAbun(70, 4) = 0.2183
    dblIsoMasses(70, 5) = 172.938208: sngIsoAbun(70, 5) = 0.1613
    dblIsoMasses(70, 6) = 173.938873: sngIsoAbun(70, 6) = 0.3183
    dblIsoMasses(70, 7) = 175.942564: sngIsoAbun(70, 7) = 0.1276
    dblIsoMasses(71, 1) = 174.940785: sngIsoAbun(71, 1) = 0.9741
    dblIsoMasses(71, 2) = 175.942679: sngIsoAbun(71, 2) = 0.0259
    dblIsoMasses(72, 1) = 173.94004: sngIsoAbun(72, 1) = 0.0016
    dblIsoMasses(72, 2) = 175.941406: sngIsoAbun(72, 2) = 0.0526
    dblIsoMasses(72, 3) = 176.943217: sngIsoAbun(72, 3) = 0.186
    dblIsoMasses(72, 4) = 177.943696: sngIsoAbun(72, 4) = 0.2728
    dblIsoMasses(72, 5) = 178.945812: sngIsoAbun(72, 5) = 0.1362
    dblIsoMasses(72, 6) = 179.946561: sngIsoAbun(72, 6) = 0.3508
    dblIsoMasses(73, 1) = 179.947462: sngIsoAbun(73, 1) = 0.00012
    dblIsoMasses(73, 2) = 180.948014: sngIsoAbun(73, 2) = 0.99988
    dblIsoMasses(74, 1) = 179.946701: sngIsoAbun(74, 1) = 0.0012
    dblIsoMasses(74, 2) = 181.948202: sngIsoAbun(74, 2) = 0.265
    dblIsoMasses(74, 3) = 182.95022: sngIsoAbun(74, 3) = 0.1431
    dblIsoMasses(74, 4) = 183.950953: sngIsoAbun(74, 4) = 0.3064
    dblIsoMasses(74, 5) = 185.954357: sngIsoAbun(74, 5) = 0.2843
    dblIsoMasses(75, 1) = 184.952951: sngIsoAbun(75, 1) = 0.374
    dblIsoMasses(75, 2) = 186.955765: sngIsoAbun(75, 2) = 0.626
    dblIsoMasses(76, 1) = 183.952488: sngIsoAbun(76, 1) = 0.0002
    dblIsoMasses(76, 2) = 185.95383: sngIsoAbun(76, 2) = 0.0159
    dblIsoMasses(76, 3) = 186.955741: sngIsoAbun(76, 3) = 0.0196
    dblIsoMasses(76, 4) = 187.95586: sngIsoAbun(76, 4) = 0.1324
    dblIsoMasses(76, 5) = 188.958137: sngIsoAbun(76, 5) = 0.1615
    dblIsoMasses(76, 6) = 189.958436: sngIsoAbun(76, 6) = 0.2626
    dblIsoMasses(76, 7) = 191.961467: sngIsoAbun(76, 7) = 0.4078            ' Note: Alternate mass is 191.960603
    dblIsoMasses(77, 1) = 190.960584: sngIsoAbun(77, 1) = 0.373
    dblIsoMasses(77, 2) = 192.962942: sngIsoAbun(77, 2) = 0.627
    dblIsoMasses(78, 1) = 189.959917: sngIsoAbun(78, 1) = 0.00014
    dblIsoMasses(78, 2) = 191.961019: sngIsoAbun(78, 2) = 0.00782
    dblIsoMasses(78, 3) = 193.962655: sngIsoAbun(78, 3) = 0.32967
    dblIsoMasses(78, 4) = 194.964785: sngIsoAbun(78, 4) = 0.33832
    dblIsoMasses(78, 5) = 195.964926: sngIsoAbun(78, 5) = 0.25242
    dblIsoMasses(78, 6) = 197.967869: sngIsoAbun(78, 6) = 0.07163
    dblIsoMasses(79, 1) = 196.966543: sngIsoAbun(79, 1) = 1
    dblIsoMasses(80, 1) = 195.965807: sngIsoAbun(80, 1) = 0.0015
    dblIsoMasses(80, 2) = 197.966743: sngIsoAbun(80, 2) = 0.0997
    dblIsoMasses(80, 3) = 198.968254: sngIsoAbun(80, 3) = 0.1687
    dblIsoMasses(80, 4) = 199.9683: sngIsoAbun(80, 4) = 0.231
    dblIsoMasses(80, 5) = 200.970277: sngIsoAbun(80, 5) = 0.1318
    dblIsoMasses(80, 6) = 201.970632: sngIsoAbun(80, 6) = 0.2986
    dblIsoMasses(80, 7) = 203.973467: sngIsoAbun(80, 7) = 0.0687
    dblIsoMasses(81, 1) = 202.97232: sngIsoAbun(81, 1) = 0.29524
    dblIsoMasses(81, 2) = 204.974401: sngIsoAbun(81, 2) = 0.70476
    dblIsoMasses(82, 1) = 203.97302: sngIsoAbun(82, 1) = 0.014
    dblIsoMasses(82, 2) = 205.97444: sngIsoAbun(82, 2) = 0.241
    dblIsoMasses(82, 3) = 206.975872: sngIsoAbun(82, 3) = 0.221
    dblIsoMasses(82, 4) = 207.976641: sngIsoAbun(82, 4) = 0.524
    dblIsoMasses(83, 1) = 208.980388: sngIsoAbun(83, 1) = 1
    dblIsoMasses(84, 1) = 209: sngIsoAbun(84, 1) = 1
    dblIsoMasses(85, 1) = 210: sngIsoAbun(85, 1) = 1
    dblIsoMasses(86, 1) = 222: sngIsoAbun(86, 1) = 1
    dblIsoMasses(87, 1) = 223: sngIsoAbun(87, 1) = 1
    dblIsoMasses(88, 1) = 226: sngIsoAbun(88, 1) = 1
    dblIsoMasses(89, 1) = 227: sngIsoAbun(89, 1) = 1
    dblIsoMasses(90, 1) = 232.038054: sngIsoAbun(90, 1) = 1
    dblIsoMasses(91, 1) = 231: sngIsoAbun(91, 1) = 1
    dblIsoMasses(92, 1) = 234.041637: sngIsoAbun(92, 1) = 0.000055
    dblIsoMasses(92, 2) = 235.043924: sngIsoAbun(92, 2) = 0.0072
    dblIsoMasses(92, 3) = 238.050786: sngIsoAbun(92, 3) = 0.992745
    dblIsoMasses(93, 1) = 237: sngIsoAbun(93, 1) = 1
    dblIsoMasses(94, 1) = 244: sngIsoAbun(94, 1) = 1
    dblIsoMasses(95, 1) = 243: sngIsoAbun(95, 1) = 1
    dblIsoMasses(96, 1) = 247: sngIsoAbun(96, 1) = 1
    dblIsoMasses(97, 1) = 247: sngIsoAbun(97, 1) = 1
    dblIsoMasses(98, 1) = 251: sngIsoAbun(98, 1) = 1
    dblIsoMasses(99, 1) = 252: sngIsoAbun(99, 1) = 1
    dblIsoMasses(100, 1) = 257: sngIsoAbun(100, 1) = 1
    dblIsoMasses(101, 1) = 258: sngIsoAbun(101, 1) = 1
    dblIsoMasses(102, 1) = 259: sngIsoAbun(102, 1) = 1
    dblIsoMasses(103, 1) = 262: sngIsoAbun(103, 1) = 1
        
    ' Note: I chose to store the desired values in the dblIsoMasses() and sngIsoAbun() 2D arrays
    '       then copy to the ElementStats() array since this method actually decreases
    '       the size of this subroutine
    For intElementIndex = 1 To ELEMENT_COUNT
        With ElementStats(intElementIndex)
            intIsotopeindex = 1
            Do While dblIsoMasses(intElementIndex, intIsotopeindex) > 0
                .Isotopes(intIsotopeindex).Abundance = sngIsoAbun(intElementIndex, intIsotopeindex)
                .Isotopes(intIsotopeindex).Mass = dblIsoMasses(intElementIndex, intIsotopeindex)
                intIsotopeindex = intIsotopeindex + 1
                If intIsotopeindex > MAX_ISOTOPES Then Exit Do
            Loop
            .IsotopeCount = intIsotopeindex - 1
        End With
    Next intElementIndex
End Sub

Public Sub MemoryLoadMessageStatements()

    MessageStatmentCount = 1555
    
    MessageStatements(1) = "Unknown element"
    MessageStatements(2) = "Obsolete msg: Cannot handle more than 4 layers of embedded parentheses"
    MessageStatements(3) = "Missing closing parentheses"
    MessageStatements(4) = "Unmatched parentheses"
    MessageStatements(5) = "Cannot have a 0 directly after an element or dash (-)"
    MessageStatements(6) = "Number too large or must only be after [, -, ), or caret (^)"
    MessageStatements(7) = "Number too large"
    MessageStatements(8) = "Obsolete msg: Cannot start formula with a number; use parentheses, brackets, or dash (-)"
    MessageStatements(9) = "Obsolete msg: Decimal numbers cannot be used after parentheses; use a [ or a caret (^)"
    MessageStatements(10) = "Obsolete msg: Decimal numbers less than 1 must be in the form .5 and not 0.5"
    MessageStatements(11) = "Numbers should follow left brackets, not right brackets (unless 'treat brackets' as parentheses is on)"
    MessageStatements(12) = "A number must be present after a bracket and/or after the decimal point"
    MessageStatements(13) = "Missing closing bracket, ]"
    MessageStatements(14) = "Misplaced number; should only be after an element, [, ), -, or caret (^)"
    MessageStatements(15) = "Unmatched bracket"
    MessageStatements(16) = "Cannot handle nested brackets or brackets inside multiple hydrates (unless 'treat brackets as parentheses' is on)"
    MessageStatements(17) = "Obsolete msg: Cannot handle multiple hydrates (extras) in brackets"
    MessageStatements(18) = "Unknown element "
    MessageStatements(19) = "Obsolete msg: Cannot start formula with a dash (-)"
    MessageStatements(20) = "There must be an isotopic mass number following the caret (^)"
    MessageStatements(21) = "Obsolete msg: Zero after caret (^); an isotopic mass of zero is not allowed"
    MessageStatements(22) = "An element must be present after the isotopic mass after the caret (^)"
    MessageStatements(23) = "Negative isotopic masses are not allowed after the caret (^)"
    MessageStatements(24) = "Isotopic masses are not allowed for abbreviations"
    MessageStatements(25) = "An element must be present after the leading coefficient of the dash"
    MessageStatements(26) = "Isotopic masses are not allowed for abbreviations; D is an abbreviation"
    MessageStatements(27) = "Numbers cannot contain more than one decimal point"
    MessageStatements(28) = "Circular abbreviation reference; can't have an abbreviation referencing a second abbreviation that depends upon the first one"
    MessageStatements(29) = "Obsolete msg: Cannot run percent solver until one or more lines are locked to a value."
    MessageStatements(30) = "Invalid formula subtraction; one or more atoms (or too many atoms) in the right-hand formula are missing (or less abundant) in the left-hand formula"
    
    ' Cases 50 through 74 are used during the % Solver routine
    MessageStatements(50) = "Target value is greater than 100%, an impossible value."
    
    ' Cases 75 through 99 are used in frmCalculator
    MessageStatements(75) = "Letters are not allowed in the calculator line"
    MessageStatements(76) = "Missing closing parenthesis"
    MessageStatements(77) = "Unmatched parentheses"
    MessageStatements(78) = "Misplaced number; or number too large, too small, or too long"
    MessageStatements(79) = "Obsolete msg: Misplaced parentheses"
    MessageStatements(80) = "Misplaced operator"
    MessageStatements(81) = "Track variable is less than or equal to 1; program bug; please notify programmer"
    MessageStatements(82) = "Missing operator. Note: ( is not needed OR allowed after a + or -"
    MessageStatements(83) = "Obsolete msg: Brackets not allowed in calculator; simply use nested parentheses"
    MessageStatements(84) = "Obsolete msg: Decimal numbers less than 1 must be in the form .5 and not 0.5"
    MessageStatements(85) = "Cannot take negative numbers to a decimal power"
    MessageStatements(86) = "Cannot take zero to a negative power"
    MessageStatements(87) = "Cannot take zero to the zeroth power"
    MessageStatements(88) = "Obsolete msg: Only a single positive or negative number is allowed after a caret (^)"
    MessageStatements(89) = "A single positive or negative number must be present after a caret (^)"
    MessageStatements(90) = "Numbers cannot contain more than one decimal point"
    MessageStatements(91) = "You tried to divide a number by zero.  Please correct the problem and recalculate."
    MessageStatements(92) = "Spaces are not allowed in mathematical expressions"

    ' Note that tags 93 and 94 are also used on frmMain
    MessageStatements(93) = "Use a period for a decimal point"
    MessageStatements(94) = "Use a comma for a decimal point"
    MessageStatements(95) = "A number must be present after a decimal point"
    
    
    ' Cases 100 and up are shown when loading data from files and starting application
    MessageStatements(100) = "Error Saving Abbreviation File"
    MessageStatements(110) = "The default abbreviation file has been re-created."
    MessageStatements(115) = "The old file has been renamed"
    MessageStatements(120) = "[AMINO ACIDS] heading not found in MWT_ABBR.DAT file.  This heading must be located before/above the [ABBREVIATIONS] heading."
    MessageStatements(125) = "Obsolete msg: Select OK to continue without any abbreviations."
    MessageStatements(130) = "[ABBREVIATIONS] heading not found in MWT_ABBR.DAT file.  This heading must be located before/above the [AMINO ACIDS] heading."
    MessageStatements(135) = "Select OK to continue with amino acids abbreviations only."
    MessageStatements(140) = "The Abbreviations File was not found in the program directory"
    MessageStatements(150) = "Error Loading/Creating Abbreviation File"
    MessageStatements(160) = "Ignoring Abbreviation -- Invalid Formula"
    MessageStatements(170) = "Ignoring Duplicate Abbreviation"
    MessageStatements(180) = "Ignoring Abbreviation; Invalid Character"
    MessageStatements(190) = "Ignoring Abbreviation; too long"
    MessageStatements(192) = "Ignoring Abbreviation; symbol length cannot be 0"
    MessageStatements(194) = "Ignoring Abbreviation; symbol most only contain letters"
    MessageStatements(196) = "Ignoring Abbreviation; Too many abbreviations in memory"
    MessageStatements(200) = "Ignoring Invalid Line"
    MessageStatements(210) = "The default elements file has been re-created."
    MessageStatements(220) = "Possibly incorrect weight for element"
    MessageStatements(230) = "Possibly incorrect uncertainty for element"
    MessageStatements(250) = "Ignoring Line; Invalid Element Symbol"
    MessageStatements(260) = "[ELEMENTS] heading not found in MWT_ELEM.DAT file.  This heading must be located in the file."
    MessageStatements(265) = "Select OK to continue with default Element values."
    MessageStatements(270) = "The Elements File was not found in the program directory"
    MessageStatements(280) = "Error Loading/Creating Elements File"
    MessageStatements(305) = "Continuing with default captions."
    MessageStatements(320) = "Error Saving Elements File"
    MessageStatements(330) = "Error Loading/Creating Values File"
    MessageStatements(340) = "Select OK to continue without loading default Values and Formulas."
    MessageStatements(345) = "If using a Read-Only drive, use the /X switch at the command line to prevent this error."
    MessageStatements(350) = "Error"
    MessageStatements(360) = "Error Saving Default Options File"
    MessageStatements(370) = "Obsolete msg: If using a Read-Only drive, you cannot save the default options."
    MessageStatements(380) = "Error Saving Values and Formulas File"
    MessageStatements(390) = "Obsolete msg: If using a Read-Only drive, you cannot save the values and formulas."
    MessageStatements(400) = "Error Loading/Creating Default Options File"
    MessageStatements(410) = "Select OK to continue without loading User Defaults."
    MessageStatements(420) = "Obsolete msg: The Default Options file was corrupted; it will be re-created."
    MessageStatements(430) = "Obsolete msg: The Values and Formulas file was corrupted; it will be re-created."
    MessageStatements(440) = "The language file could not be successfully opened or was formatted incorrectly."
    MessageStatements(450) = "Unable to load language-specific captions"
    MessageStatements(460) = "The language file could not be found in the program directory"
    MessageStatements(470) = "The file requested for molecular weight processing was not found"
    MessageStatements(480) = "File Not Found"
    MessageStatements(490) = "This file already exists.  Replace it?"
    MessageStatements(500) = "File Exists"
    MessageStatements(510) = "Error Reading/Writing files for batch processing"
    MessageStatements(515) = "Select OK to abort batch file processing."
    MessageStatements(520) = "Error in program"
    MessageStatements(530) = "These lines of code should not have been encountered.  Please notify programmer."
    MessageStatements(540) = "Obsolete msg: You can't edit elements because the /X switch was used at the command line."
    MessageStatements(545) = "Obsolete msg: You can't edit abbreviations because the /X switch was used at the command line."
    MessageStatements(550) = "Percent solver cannot be used when brackets are being treated as parentheses.  You can change the bracket recognition mode by choosing Change Program Preferences under the Options menu."
    MessageStatements(555) = "Percent Solver not Available"
    MessageStatements(560) = "Maximum number of formula fields exist."
    MessageStatements(570) = "Current formula is blank."
    MessageStatements(580) = "Turn off Percent Solver (F11) before creating a new formula."
    MessageStatements(590) = "An overflow error has occured.  Please reduce number sizes and recalculate."
    MessageStatements(600) = "An error has occured"
    MessageStatements(605) = "Please exit the program and report the error to the programmer.  Select About from the Help menu to see the E-mail address."
    MessageStatements(610) = "Spaces are not allowed in formulas"
    MessageStatements(620) = "Invalid Character"
    MessageStatements(630) = "Cannot copy to new formula."
    MessageStatements(645) = "Obsolete msg: Maximum number of formulas is 7"
    MessageStatements(650) = "Current formula is blank."
    MessageStatements(655) = "Percent Solver mode is on (F11 to exit mode)."
    MessageStatements(660) = "Warning, isotopic mass is probably too large for element"
    MessageStatements(662) = "Warning, isotopic mass is probably too small for element"
    MessageStatements(665) = "vs avg atomic wt of"
    MessageStatements(670) = "Warning, isotopic mass is impossibly small for element"
    MessageStatements(675) = "protons"
    MessageStatements(680) = "Note: Exact Mode is on"
    MessageStatements(685) = "Note: for % Solver, a left bracket must precede an x"
    MessageStatements(690) = "Note: brackets are being treated as parentheses"
    MessageStatements(700) = "One or more elements must be checked."
    MessageStatements(705) = "Maximum hits must be greater than 0."
    MessageStatements(710) = "Maximum hits must be less than "
    MessageStatements(715) = "Minimum number of elements must be 0 or greater."
    MessageStatements(720) = "Minimum number of elements must be less than maximum number of elements."
    MessageStatements(725) = "Maximum number of elements must be less than 65,025"
    MessageStatements(730) = "An atomic weight must be entered for custom elements."
    MessageStatements(735) = "Atomic Weight must be greater than 0 for custom elements."
    MessageStatements(740) = "Target molecular weight must be entered."
    MessageStatements(745) = "Target molecular weight must be greater than 0."
    MessageStatements(750) = "Obsolete msg: Weight tolerance must be 0 or greater."
    MessageStatements(755) = "A maximum molecular weight must be entered."
    MessageStatements(760) = "Maximum molecular weight must be greater than 0."
    MessageStatements(765) = "Target percentages must be entered for element"
    MessageStatements(770) = "Target percentage must be greater than 0."
    MessageStatements(775) = "Custom elemental weights must contain only numbers or only letters.  If letters are used, they must be for a single valid elemental symbol or abbreviation."
    MessageStatements(780) = "Custom elemental weight is empty.  If letters are used, they must be for a single valid elemental symbol or abbreviation."
    MessageStatements(785) = "Unknown element or abbreviation for custom elemental weight"
    MessageStatements(790) = "Only single elemental symbols or abbreviations are allowed."
    MessageStatements(800) = "Caution, no abbreviations were loaded -- Command has no effect."
    MessageStatements(805) = "Cannot handle fractional numbers of atoms"
    MessageStatements(910) = "Ions are already present in the ion list.  Replace with new ions?"
    MessageStatements(920) = "Replace Existing Ions"
    MessageStatements(930) = "Loading Ion List"
    MessageStatements(940) = "Process aborted"
    MessageStatements(945) = " aborted"
    MessageStatements(950) = "Normalizing ions"
    MessageStatements(960) = "Normalizing by region"
    MessageStatements(965) = "Sorting by Intensity"
    MessageStatements(970) = "Matching Ions"
    MessageStatements(980) = "The clipboard is empty.  No ions to paste."
    MessageStatements(985) = "No ions"
    MessageStatements(990) = "Pasting ion list"
    MessageStatements(1000) = "Determining number of ions in list"
    MessageStatements(1010) = "Parsing list"
    MessageStatements(1020) = "No valid ions were found on the clipboard.  A valid ion list is a list of mass and intensity pairs, separated by commas, tabs, or spaces.  One mass/intensity pair should be present per line."
    
    MessageStatements(1030) = "Error writing data to file"
    MessageStatements(1040) = "Set Range"
    MessageStatements(1050) = "Start Val"
    MessageStatements(1055) = "End Val"
    MessageStatements(1060) = "Set X Axis Range"
    MessageStatements(1065) = "Set Y Axis Range"
    MessageStatements(1070) = "Enter a new Gaussian Representation quality factor.  Higher numbers result in smoother Gaussian curves, but slower updates.  Valid range is 1 to 50, default is 20."
    MessageStatements(1072) = "Gaussian Representation Quality"
    MessageStatements(1075) = "Enter a new plotting approximation factor. Higher numbers result in faster updates, but give a less accurate graphical representation when viewing a wide mass range (zoomed out).  Valid range is 1 to 50, default is 10."
    MessageStatements(1077) = "Plotting Approximation Factor"
    MessageStatements(1080) = "Resolving Power Specifications"
    MessageStatements(1090) = "Resolving Power"
    MessageStatements(1100) = "X Value of Specification"
    MessageStatements(1110) = "Please enter the approximate number of ticks to show on the axis"
    MessageStatements(1115) = "Axis Ticks"
    MessageStatements(1120) = "Creating Gaussian Representation"
    MessageStatements(1130) = "Preparing plot"
    MessageStatements(1135) = "Drawing plot"
    MessageStatements(1140) = "Are you sure you want to restore the default plotting options?"
    MessageStatements(1145) = "Restore Default Options"
    MessageStatements(1150) = "Auto Align Ions"
    MessageStatements(1155) = "Maximum Offset"
    MessageStatements(1160) = "Offset Increment"
    MessageStatements(1165) = "Aligning Ions"
    
    MessageStatements(1200) = "Caution symbol must be 1 to " & MAX_ABBREV_LENGTH & " characters long"
    MessageStatements(1205) = "Caution symbol most only contain letters"
    MessageStatements(1210) = "Caution description length cannot be 0"
    MessageStatements(1215) = "Too many caution statements.  Unable to add another one."
    
    MessageStatements(1500) = "All Files"
    MessageStatements(1510) = "Text Files"
    MessageStatements(1515) = "txt"
    MessageStatements(1520) = "Data Files"
    MessageStatements(1525) = "csv"
    MessageStatements(1530) = "Sequence Files"
    MessageStatements(1535) = "seq"
    MessageStatements(1540) = "Ion List Files"
    MessageStatements(1545) = "txt"
    MessageStatements(1550) = "Capillary Flow Info Files"
    MessageStatements(1555) = "cap"
            
End Sub

Private Sub MwtWinDllErrorHandler(strSourceForm As String)
    Dim strMessage As String
    Dim blnShowErrorMessageDialogsSaved As Boolean
    
    If Err.Number = 6 Then
        MsgBox LookupMessage(590), vbOKOnly + vbExclamation, LookupMessage(350)
    Else
        strMessage = LookupMessage(600) & ": " & Err.Description & vbCrLf & _
                   " (" & strSourceForm & " handler)"
        strMessage = strMessage & vbCrLf & LookupMessage(605)
        MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)
        
        blnShowErrorMessageDialogsSaved = mShowErrorMessageDialogs
        mShowErrorMessageDialogs = False
        
        ' Call GeneralErrorHandler so that the error gets logged to ErrorLog.txt
        GeneralErrorHandler strSourceForm, Err.Number
        
        mShowErrorMessageDialogs = blnShowErrorMessageDialogsSaved
        
    End If
End Sub

Private Sub InitializeAbbrevSymbolStack(ByRef udtAbbrevSymbolStack As udtAbbrevSymbolStackType)
    With udtAbbrevSymbolStack
        .Count = 0
        ReDim .SymbolReferenceStack(0)
    End With
End Sub
    
Private Sub InitializeComputationStats(ByRef udtComputationStats As udtComputationStatsType)
    Dim intElementIndex As Integer
    
    With udtComputationStats
        .Charge = 0!
        .StandardDeviation = 0#
        .TotalMass = 0#
        
        For intElementIndex = 0 To ELEMENT_COUNT
            With .Elements(intElementIndex)
                .Used = False                ' whether element is present
                .Count = 0                   ' # of each element
                .IsotopicCorrection = 0      ' isotopic correction
                .IsotopeCount = 0            ' Count of the number of atoms defined as specific isotopes
                ReDim .Isotopes(2)           ' Default to have room for 2 explicitly defined isotopes
            End With
        Next intElementIndex
    End With

End Sub

Public Function ParseFormulaPublic(ByRef strFormula As String, ByRef udtComputationStats As udtComputationStatsType, Optional blnExpandAbbreviations As Boolean = False, Optional dblValueForX As Double = 1#) As Double
    ' Determines the molecular weight and elemental composition of strFormula
    ' Returns the computed molecular weight if no error; otherwise, returns -1
    ' ErrorParams will hold information on errors that occur (previous errors are cleared when this function is called)
    ' Returns other useful information in udtComputationStats()
    ' Use ComputeFormulaWeight if you only want to know the weight of a formula (it calls this function)
    
    Dim intElementIndex As Integer
    Dim dblStdDevSum As Double
    Dim udtAbbrevSymbolStack As udtAbbrevSymbolStackType
    
On Error GoTo ParseFormulaPublicErrorHandler

    InitializeComputationStats udtComputationStats
    InitializeAbbrevSymbolStack udtAbbrevSymbolStack
    
    dblStdDevSum = 0#
    
    ' Reset ErrorParams to clear any prior errors
    ResetErrorParamsInternal
    
    ' Reset Caution Description
    mStrCautionDescription = ""
    
    If Len(strFormula) > 0 Then
        strFormula = ParseFormulaRecursive(strFormula, udtComputationStats, udtAbbrevSymbolStack, blnExpandAbbreviations, dblStdDevSum, dblValueForX)
    End If
    
    ' Copy udtComputationStats to mComputationStatsSaved
    mComputationStatsSaved = udtComputationStats
    
    If ErrorParams.ErrorID = 0 Then
        With udtComputationStats
            
            ' Compute the standard deviation
            .StandardDeviation = Sqr(dblStdDevSum)
            
            ' Compute the total molecular weight
            .TotalMass = 0      ' Reset total weight of compound to 0 so we can add to it
            For intElementIndex = 1 To ELEMENT_COUNT
                ' Increase total weight by multipling the count of each element by the element's mass
                ' In addition, add in the Isotopic Correction value
                .TotalMass = .TotalMass + _
                               ElementStats(intElementIndex).Mass * .Elements(intElementIndex).Count + _
                               .Elements(intElementIndex).IsotopicCorrection
            Next intElementIndex
        End With
        ParseFormulaPublic = udtComputationStats.TotalMass
    Else
        ParseFormulaPublic = -1
    End If

    Exit Function

ParseFormulaPublicErrorHandler:
    GeneralErrorHandler "MWPeptideClass.ParseFormulaPublic", Err.Number

End Function

Private Function ParseFormulaRecursive(ByVal strFormula As String, ByRef udtComputationStats As udtComputationStatsType, ByRef udtAbbrevSymbolStack As udtAbbrevSymbolStackType, ByVal blnExpandAbbreviations As Boolean, ByRef dblStdDevSum As Double, Optional dblValueForX As Double = 1#, Optional ByVal intCharCountPrior As Integer = 0, Optional ByVal dblParenthMultiplier As Double = 1#, Optional ByVal dblDashMultiplierPrior As Double = 1#, Optional ByVal dblBracketMultiplierPrior As Double = 1#, Optional ByRef CarbonOrSiliconReturnCount As Long = 0, Optional ByVal intParenthLevelPrevious As Integer = 0) As String
    
    ' Determine elements in an abbreviation or elements and abbreviations in a formula
    ' Stores results in udtComputationStats
    ' ErrorParams will hold information on errors that occur
    ' Returns the formatted formula
    
    ' blnDisplayMessages indicates whether to display error messages
    ' dblParenthMultiplier is the value to multiply all values by if inside parentheses
    ' dblStdDevSum is the sum of the squares of the standard deviations
    ' CarbonOrSiliconReturnCount records the number of carbon and silicon atoms found; used when correcting for charge inside parentheses or inside an abbreviation
    
    ' ( and ) are 40 and 41   - is 45   { and } are 123 and 125
    ' Numbers are 48 to 57    . is 46
    ' Uppercase letters are 65 to 90
    ' Lowercase letters are 97 to 122
    ' [ and ] are 91 and 93
    ' ^ is 94
    ' _ is 95
    
    Dim intAddonCount As Integer, intSymbolLength As Integer
    Dim blnCaretPresent As Boolean, intElementIndex As Integer, intNumLength As Integer
    Dim intCharIndex As Integer, intMinusSymbolLoc As Integer
    Dim strLeftHalf As String, strRightHalf As String, blnMatchFound As Boolean
    Dim strNewFormulaRightHalf As String
    Dim udtComputationStatsRightHalf As udtComputationStatsType
    Dim udtAbbrevSymbolStackRightHalf As udtAbbrevSymbolStackType
    Dim dblStdDevSumRightHalf As Double
    Dim dblAdjacentNum As Double, dblCaretVal As Double, dblCaretValDifference As Double
    Dim dblAtomCountToAdd As Double
    Dim dblBracketMultiplier As Double, blnInsideBrackets As Boolean
    Dim intDashPos As Integer, dblDashMultiplier As Double
    Dim sngChargeSaved As Single
    Dim eSymbolMatchType As smtSymbolMatchTypeConstants
    Dim strChar1 As String, strChar2 As String, strChar3 As String, strCharRemain As String
    Dim strFormulaExcerpt As String
    Dim strCharVal As String, intCharAsc As Integer
    Dim LoneCarbonOrSilicon As Long
    Dim dblIsoDifferenceTop As Double, dblIsoDifferenceBottom As Double
    
    Dim SymbolReference As Integer, PrevSymbolReference As Integer
    Dim strReplace As String, strNewFormula As String, strSubFormula As String
    Dim intParenthClose As Integer, intParenthLevel As Integer
    Dim intExpandAbbrevAdd As Integer
    
On Error GoTo ErrorHandler
    
    dblDashMultiplier = dblDashMultiplierPrior     ' Leading coefficient position and default value
    dblBracketMultiplier = dblBracketMultiplierPrior    ' Bracket correction factor
    blnInsideBrackets = False                     ' Switch for in or out of brackets
    
    intDashPos = 0
    
    LoneCarbonOrSilicon = 0                 ' The number of carbon or silicon atoms
    CarbonOrSiliconReturnCount = 0
    
    ' Look for the > symbol
    ' If found, this means take First Part minus the Second Part
    intMinusSymbolLoc = InStr(strFormula, ">")
    If intMinusSymbolLoc > 0 Then
        ' Look for the first occurrence of >
        intCharIndex = 1
        blnMatchFound = False
        Do
            If Mid(strFormula, intCharIndex, 1) = ">" Then
                blnMatchFound = True
                strLeftHalf = Left(strFormula, intCharIndex - 1)
                strRightHalf = Mid(strFormula, intCharIndex + 1)
                
                ' Parse the first half
                strNewFormula = ParseFormulaRecursive(strLeftHalf, udtComputationStats, udtAbbrevSymbolStack, blnExpandAbbreviations, dblStdDevSum, dblValueForX, intCharCountPrior, dblParenthMultiplier, dblDashMultiplier, dblBracketMultiplier, CarbonOrSiliconReturnCount, intParenthLevelPrevious)
                
                ' Parse the second half
                InitializeComputationStats udtComputationStatsRightHalf
                InitializeAbbrevSymbolStack udtAbbrevSymbolStackRightHalf
                
                strNewFormulaRightHalf = ParseFormulaRecursive(strRightHalf, udtComputationStatsRightHalf, udtAbbrevSymbolStackRightHalf, blnExpandAbbreviations, dblStdDevSumRightHalf, dblValueForX, intCharCountPrior + intCharIndex, dblParenthMultiplier, dblDashMultiplier, dblBracketMultiplier, CarbonOrSiliconReturnCount, intParenthLevelPrevious)
                
                Exit Do
            End If
            intCharIndex = intCharIndex + 1
        Loop While intCharIndex <= Len(strFormula)
        ' blnMatchFound should always be true; thus, I can probably remove it
        Debug.Assert blnMatchFound
        
        If blnMatchFound Then
            ' Update strFormula
            strFormula = strNewFormula & ">" & strNewFormulaRightHalf
            
            ' Update udtComputationStats by subtracting the atom counts of the first half minus the second half
            ' If any atom counts become < 0 then, then raise an error
            For intElementIndex = 1 To ELEMENT_COUNT
                With udtComputationStats.Elements(intElementIndex)
                    If ElementStats(intElementIndex).Mass * .Count + .IsotopicCorrection >= _
                       ElementStats(intElementIndex).Mass * udtComputationStatsRightHalf.Elements(intElementIndex).Count + udtComputationStatsRightHalf.Elements(intElementIndex).IsotopicCorrection Then
                       
                        .Count = .Count - udtComputationStatsRightHalf.Elements(intElementIndex).Count
                        If .Count < 0 Then
                            ' This shouldn't happen
                            Debug.Assert False
                            .Count = 0
                        End If
                        
                        If udtComputationStatsRightHalf.Elements(intElementIndex).IsotopicCorrection <> 0 Then
                            ' This assertion is here simply because I want to check the code
                            Debug.Assert False
                            .IsotopicCorrection = .IsotopicCorrection - udtComputationStatsRightHalf.Elements(intElementIndex).IsotopicCorrection
                        End If
                    Else
                        ' Invalid Formula; raise error
                        ErrorParams.ErrorID = 30: ErrorParams.ErrorPosition = intCharIndex
                    End If
                End With
                If ErrorParams.ErrorID <> 0 Then Exit For
            Next intElementIndex
            
            ' Adjust the overall charge
            udtComputationStats.Charge = udtComputationStats.Charge - udtComputationStatsRightHalf.Charge
        End If
    Else
        
        ' Formula does not contain >
        ' Parse it
        intCharIndex = 1
        Do
            strChar1 = Mid(strFormula, intCharIndex, 1)
            strChar2 = Mid(strFormula, intCharIndex + 1, 1)
            strChar3 = Mid(strFormula, intCharIndex + 2, 1)
            strCharRemain = Mid(strFormula, intCharIndex + 3)
            If gComputationOptions.CaseConversion <> ccExactCase Then strChar1 = UCase(strChar1)
                
            If gComputationOptions.BracketsAsParentheses Then
                If strChar1 = "[" Then strChar1 = "("
                If strChar1 = "]" Then strChar1 = ")"
            End If
            
            eSymbolMatchType = smtUnknown
            If strChar1 = "" Then strChar1 = EMPTY_STRINGCHAR
            If strChar2 = "" Then strChar2 = EMPTY_STRINGCHAR
            If strChar3 = "" Then strChar3 = EMPTY_STRINGCHAR
            If strCharRemain = "" Then strCharRemain = EMPTY_STRINGCHAR
            
            strFormulaExcerpt = strChar1 & strChar2 & strChar3 & strCharRemain
            
            ' Check for needed caution statements
            CheckCaution strFormulaExcerpt
            
            Select Case Asc(strChar1)
            Case 40, 123         ' (    Record its position
                ' See if a number is present just after the opening parenthesis
                If IsNumeric(strChar2) Or strChar2 = "." Then
                    ' Misplaced number
                    ErrorParams.ErrorID = 14: ErrorParams.ErrorPosition = intCharIndex
                End If
                
                If ErrorParams.ErrorID = 0 Then
                    ' search for closing parenthesis
                    intParenthLevel = 1
                    For intParenthClose = intCharIndex + 1 To Len(strFormula)
                        Select Case Mid(strFormula, intParenthClose, 1)
                        Case "(", "{", "["
                            ' Another opening parentheses
                            ' increment parenthtrack
                            If Not gComputationOptions.BracketsAsParentheses And Mid(strFormula, intParenthClose, 1) = "[" Then
                                ' Do not count the bracket
                            Else
                                intParenthLevel = intParenthLevel + 1
                            End If
                        Case ")", "}", "]"
                            If Not gComputationOptions.BracketsAsParentheses And Mid(strFormula, intParenthClose, 1) = "]" Then
                                ' Do not count the bracket
                            Else
                                intParenthLevel = intParenthLevel - 1
                                If intParenthLevel = 0 Then
                                    dblAdjacentNum = ParseNum(Mid(strFormula, intParenthClose + 1), intNumLength)
                                    CatchParsenumError dblAdjacentNum, intNumLength, intCharIndex, intSymbolLength
                                    
                                    If dblAdjacentNum < 0 Then
                                        dblAdjacentNum = 1#
                                        intAddonCount = 0
                                    Else
                                        intAddonCount = intNumLength
                                    End If
                                    
                                    strSubFormula = Mid(strFormula, intCharIndex + 1, intParenthClose - (intCharIndex + 1))
                                    
                                    ' Note, must pass parenthnum * dblAdjacentNum to preserve previous parentheses stuff
                                    strNewFormula = ParseFormulaRecursive(strSubFormula, udtComputationStats, udtAbbrevSymbolStack, blnExpandAbbreviations, dblStdDevSum, dblValueForX, intCharCountPrior + intCharIndex, dblParenthMultiplier * dblAdjacentNum, dblDashMultiplier, dblBracketMultiplier, CarbonOrSiliconReturnCount, intParenthLevelPrevious + 1)
                                    
                                    ' If expanding abbreviations, then strNewFormula might be longer than strFormula, must add this onto intCharIndex also
                                    intExpandAbbrevAdd = Len(strNewFormula) - Len(strSubFormula)
                                    
                                    ' Must replace the part of the formula parsed with the strNewFormula part, in case the formula was expanded or elements were capitalized
                                    strFormula = Left(strFormula, intCharIndex) & strNewFormula & Mid(strFormula, intParenthClose)
                                    intCharIndex = intParenthClose + intAddonCount + intExpandAbbrevAdd
                                    
                                    ' Correct charge
                                    If CarbonOrSiliconReturnCount > 0 Then
                                        udtComputationStats.Charge = udtComputationStats.Charge - 2 * dblAdjacentNum
                                        If dblAdjacentNum > 1 And CarbonOrSiliconReturnCount > 1 Then
                                            udtComputationStats.Charge = udtComputationStats.Charge - 2 * (dblAdjacentNum - 1) * (CarbonOrSiliconReturnCount - 1)
                                        End If
                                    End If
                                    Exit For
                                End If
                            End If
                        End Select
                    Next intParenthClose
                End If
                
                If intParenthLevel > 0 And ErrorParams.ErrorID = 0 Then
                    ' Missing closing parenthesis
                    ErrorParams.ErrorID = 3: ErrorParams.ErrorPosition = intCharIndex
                End If
                PrevSymbolReference = 0
            Case 41, 125         ' )    Repeat a section of a formula
                ' Should have been skipped
                ' Unmatched closing parentheses
                ErrorParams.ErrorID = 4: ErrorParams.ErrorPosition = intCharIndex
                
            Case 45            ' -
                ' Used to denote a leading coefficient
                dblAdjacentNum = ParseNum(strChar2 & strChar3 & strCharRemain, intNumLength)
                CatchParsenumError dblAdjacentNum, intNumLength, intCharIndex, intSymbolLength
                
                If dblAdjacentNum > 0 Then
                    intDashPos = intCharIndex + intNumLength
                    dblDashMultiplier = dblAdjacentNum * dblDashMultiplierPrior
                    intCharIndex = intCharIndex + intNumLength
                Else
                    If dblAdjacentNum = 0 Then
                        ' Cannot have 0 after a dash
                        ErrorParams.ErrorID = 5: ErrorParams.ErrorPosition = intCharIndex + 1
                    Else
                        ' No number is present, that's just fine
                        ' Make sure defaults are set, though
                        intDashPos = 0
                        dblDashMultiplier = dblDashMultiplierPrior
                    End If
                End If
                PrevSymbolReference = 0
            Case 44, 46, 48 To 57      ' . or , and Numbers 0 to 9
                ' They should only be encountered as a leading coefficient
                ' Should have been bypassed when the coefficient was processed
                If intCharIndex = 1 Then
                    ' Formula starts with a number -- multiply section by number (until next dash)
                    dblAdjacentNum = ParseNum(strFormulaExcerpt, intNumLength)
                    CatchParsenumError dblAdjacentNum, intNumLength, intCharIndex, intSymbolLength
                    
                    If dblAdjacentNum >= 0 Then
                        intDashPos = intCharIndex + intNumLength - 1
                        dblDashMultiplier = dblAdjacentNum * dblDashMultiplierPrior
                        intCharIndex = intCharIndex + intNumLength - 1
                    Else
                        ' A number less then zero should have been handled by CatchParsenumError above
                        ' Make sure defaults are set, though
                        intDashPos = 0
                        dblDashMultiplier = dblDashMultiplierPrior
                    End If
                Else
                    If CDblSafe(Mid(strFormula, intCharIndex - 1, 1)) > 0 Then
                        ' Number too large
                        ErrorParams.ErrorID = 7: ErrorParams.ErrorPosition = intCharIndex
                    Else
                        ' Misplaced number
                        ErrorParams.ErrorID = 14: ErrorParams.ErrorPosition = intCharIndex
                    End If
                End If
                PrevSymbolReference = 0
            Case 91 ' [
                If UCase(strChar2) = "X" Then
                    If strChar3 = "e" Then
                        dblAdjacentNum = ParseNum(strChar2 & strChar3 & strCharRemain, intNumLength)
                        CatchParsenumError dblAdjacentNum, intNumLength, intCharIndex, intSymbolLength
                    Else
                        dblAdjacentNum = dblValueForX
                        intNumLength = 1
                    End If
                Else
                    dblAdjacentNum = ParseNum(strChar2 & strChar3 & strCharRemain, intNumLength)
                    CatchParsenumError dblAdjacentNum, intNumLength, intCharIndex, intSymbolLength
                    
                End If
    
                If ErrorParams.ErrorID = 0 Then
                    If blnInsideBrackets Then
                        ' No Nested brackets.
                        ErrorParams.ErrorID = 16: ErrorParams.ErrorPosition = intCharIndex
                    Else
                        If dblAdjacentNum < 0 Then
                            ' No number after bracket
                            ErrorParams.ErrorID = 12: ErrorParams.ErrorPosition = intCharIndex + 1
                        Else
                            ' Coefficient for bracketed section.
                            blnInsideBrackets = True
                            dblBracketMultiplier = dblAdjacentNum * dblBracketMultiplierPrior     ' Take times dblBracketMultiplierPrior in case it wasn't 1 to start with
                            intCharIndex = intCharIndex + intNumLength
                        End If
                    End If
                End If
                PrevSymbolReference = 0
            Case 93 ' ]
                dblAdjacentNum = ParseNum(strChar2 & strChar3 & strCharRemain, intNumLength)
                CatchParsenumError dblAdjacentNum, intNumLength, intCharIndex, intSymbolLength
                
                If dblAdjacentNum >= 0 Then
                    ' Number following bracket
                    ErrorParams.ErrorID = 11: ErrorParams.ErrorPosition = intCharIndex + 1
                Else
                    If blnInsideBrackets Then
                        If intDashPos > 0 Then
                            ' Need to set intDashPos and dblDashMultiplier back to defaults, since a dash number goes back to one inside brackets
                            intDashPos = 0
                            dblDashMultiplier = 1
                        End If
                        blnInsideBrackets = False
                        dblBracketMultiplier = dblBracketMultiplierPrior
                      Else
                        ' Unmatched bracket
                        ErrorParams.ErrorID = 15: ErrorParams.ErrorPosition = intCharIndex
                    End If
                End If
            Case 65 To 90, 97 To 122, 43, 95      ' Uppercase A to Z and lowercase a to z, and the plus (+) sign, and the underscore (_)
                eSymbolMatchType = smtUnknown
                intAddonCount = 0
                dblAdjacentNum = 0
                
                eSymbolMatchType = CheckElemAndAbbrev(strFormulaExcerpt, SymbolReference)
                
                Select Case eSymbolMatchType
                Case smtElement
                    ' Found an element
                    ' SymbolReference is the elemental number
                    intSymbolLength = Len(ElementStats(SymbolReference).Symbol)
                    If intSymbolLength = 0 Then
                        ' No elements in ElementStats yet
                        ' Set intSymbolLength to 1
                        intSymbolLength = 1
                    End If
                    ' Look for number after element
                    dblAdjacentNum = ParseNum(Mid(strFormula, intCharIndex + intSymbolLength), intNumLength)
                    CatchParsenumError dblAdjacentNum, intNumLength, intCharIndex, intSymbolLength
                    
                    If dblAdjacentNum < 0 Then
                        dblAdjacentNum = 1
                    End If
    
                    ' Note that intNumLength = 0 if dblAdjacentNum was -1 or otherwise < 0
                    intAddonCount = intNumLength + intSymbolLength - 1
                    
                    If dblAdjacentNum = 0 Then
                        ' Zero after element
                        ErrorParams.ErrorID = 5: ErrorParams.ErrorPosition = intCharIndex + intSymbolLength
                    Else
                        If Not blnCaretPresent Then
                            dblAtomCountToAdd = dblAdjacentNum * dblBracketMultiplier * dblParenthMultiplier * dblDashMultiplier
                            With udtComputationStats.Elements(SymbolReference)
                                .Count = .Count + dblAtomCountToAdd
                                .Used = True     ' Element is present tag
                                dblStdDevSum = dblStdDevSum + dblAtomCountToAdd * ((ElementStats(SymbolReference).Uncertainty) ^ 2)
                            End With
                            
                            With udtComputationStats
                                ' Compute charge
                                If SymbolReference = 1 Then
                                    ' Dealing with hydrogen
                                    Select Case PrevSymbolReference
                                    Case 1, 3 To 6, 11 To 14, 19 To 32, 37 To 50, 55 To 82, 87 To 109
                                        ' Hydrogen is -1 with metals (non-halides)
                                        .Charge = .Charge + dblAtomCountToAdd * (-1)
                                    Case Else
                                        .Charge = .Charge + dblAtomCountToAdd * (ElementStats(SymbolReference).Charge)
                                    End Select
                                Else
                                    .Charge = .Charge + dblAtomCountToAdd * (ElementStats(SymbolReference).Charge)
                                End If
                            End With
                            
                            If SymbolReference = 6 Or SymbolReference = 14 Then
                                ' Sum up number lone C and Si (not in abbreviations)
                                LoneCarbonOrSilicon = LoneCarbonOrSilicon + dblAdjacentNum
                                CarbonOrSiliconReturnCount = CarbonOrSiliconReturnCount + dblAdjacentNum
                            End If
                        Else
                            ' blnCaretPresent = True
                            ' Check to make sure isotopic mass is reasonable
                            dblIsoDifferenceTop = CIntSafeDbl(0.63 * SymbolReference + 6)
                            dblIsoDifferenceBottom = CIntSafeDbl(0.008 * SymbolReference ^ 2 - 0.4 * SymbolReference - 6)
                            dblCaretValDifference = dblCaretVal - SymbolReference * 2
                            
                            If dblCaretValDifference >= dblIsoDifferenceTop Then
                                ' Probably too high isotopic mass
                                AddToCautionDescription LookupMessage(660) & ": " & ElementStats(SymbolReference).Symbol & " - " & CStr(dblCaretVal) & " " & LookupMessage(665) & " " & CStr(ElementStats(SymbolReference).Mass)
                            ElseIf dblCaretVal < SymbolReference Then
                                ' Definitely too low isotopic mass
                                AddToCautionDescription LookupMessage(670) & ": " & ElementStats(SymbolReference).Symbol & " - " & CStr(SymbolReference) & " " & LookupMessage(675)
                            ElseIf dblCaretValDifference <= dblIsoDifferenceBottom Then
                                ' Probably too low isotopic mass
                                AddToCautionDescription LookupMessage(662) & ": " & ElementStats(SymbolReference).Symbol & " - " & CStr(dblCaretVal) & " " & LookupMessage(665) & " " & CStr(ElementStats(SymbolReference).Mass)
                            End If
                            
                            ' Put in isotopic correction factor
                            dblAtomCountToAdd = dblAdjacentNum * dblBracketMultiplier * dblParenthMultiplier * dblDashMultiplier
                            With udtComputationStats.Elements(SymbolReference)
                                ' Increment element counting bin
                                .Count = .Count + dblAtomCountToAdd
                                
                                ' Store information in .Isotopes()
                                ' Increment the isotope counting bin
                                .IsotopeCount = .IsotopeCount + 1
                                
                                If UBound(.Isotopes) < .IsotopeCount Then
                                    ReDim Preserve .Isotopes(UBound(.Isotopes) + 2)
                                End If
                                
                                With .Isotopes(.IsotopeCount)
                                    .Count = .Count + dblAtomCountToAdd
                                    .Mass = dblCaretVal
                                End With
                                
                                ' Add correction amount to udtComputationStats.elements(SymbolReference).IsotopicCorrection
                                .IsotopicCorrection = .IsotopicCorrection + (dblCaretVal * dblAtomCountToAdd - ElementStats(SymbolReference).Mass * dblAtomCountToAdd)
                                
                                ' Set bit that element is present
                                .Used = True
                                
                                ' Assume no error in caret value, no need to change dblStdDevSum
                            End With
                            
                            ' Reset blnCaretPresent
                            blnCaretPresent = False
                        End If
                        If gComputationOptions.CaseConversion = ccConvertCaseUp Then
                            strFormula = Left(strFormula, intCharIndex - 1) + UCase(Mid(strFormula, intCharIndex, 1)) + Mid(strFormula, intCharIndex + 1)
                        End If
                        intCharIndex = intCharIndex + intAddonCount
                    End If
        
                Case smtAbbreviation
                    ' Found an abbreviation or amino acid
                    ' SymbolReference is the abbrev or amino acid number
    
                    If IsPresentInAbbrevSymbolStack(udtAbbrevSymbolStack, SymbolReference) Then
                        ' Circular Reference: Can't have an abbreviation referencing an abbreviation that depends upon it
                        ' For example, the following is impossible:  Lor = C6H5Tal and Tal = H4O2Lor
                        ' Furthermore, can't have this either: Lor = C6H5Tal and Tal = H4O2Vin and Vin = S3Lor
                        ErrorParams.ErrorID = 28: ErrorParams.ErrorPosition = intCharIndex
                    Else
                        ' Found an abbreviation
                        If blnCaretPresent Then
                            ' Cannot have isotopic mass for an abbreviation, including deuterium
                            If UCase(strChar1) = "D" And strChar2 <> "y" Then
                                ' Isotopic mass used for Deuterium
                                ErrorParams.ErrorID = 26: ErrorParams.ErrorPosition = intCharIndex
                            Else
                                ErrorParams.ErrorID = 24: ErrorParams.ErrorPosition = intCharIndex
                            End If
                        Else
                            ' Parse abbreviation
                            ' Simply treat it like a formula surrounded by parentheses
                            ' Thus, find the number after the abbreviation, then call ParseFormulaRecursive, sending it the formula for the abbreviation
                            ' Update the udtAbbrevSymbolStack before calling so that we can check for circular abbreviation references
                            
                            ' Record the abbreviation length
                            intSymbolLength = Len(AbbrevStats(SymbolReference).Symbol)
                            
                            ' Look for number after abbrev/amino
                            dblAdjacentNum = ParseNum(Mid(strFormula, intCharIndex + intSymbolLength), intNumLength)
                            CatchParsenumError dblAdjacentNum, intNumLength, intCharIndex, intSymbolLength
                            
                            If dblAdjacentNum < 0 Then
                                dblAdjacentNum = 1
                                intAddonCount = intSymbolLength - 1
                            Else
                                intAddonCount = intNumLength + intSymbolLength - 1
                            End If
                            
                            ' Add this abbreviation symbol to the Abbreviation Symbol Stack
                            AbbrevSymbolStackAdd udtAbbrevSymbolStack, SymbolReference
                            
                            ' Compute the charge prior to calling ParseFormulaRecursive
                            ' During the call to ParseFormulaRecursive, udtComputationStats.Charge will be
                            '  modified according to the atoms in the abbreviations's formula
                            ' This is not what we want; instead, we want to use the defined charge for the abbreviation
                            ' We'll use the dblAtomCountToAdd variable here, though instead of an atom count, it's really an abbreviation occurrence count
                            dblAtomCountToAdd = dblAdjacentNum * dblBracketMultiplier * dblParenthMultiplier * dblDashMultiplier
                            sngChargeSaved = udtComputationStats.Charge + dblAtomCountToAdd * (AbbrevStats(SymbolReference).Charge)
                            
                            ' When parsing an abbreviation, do not pass on the value of blnExpandAbbreviations
                            ' This way, an abbreviation containing an abbreviation will only get expanded one level
                            ParseFormulaRecursive AbbrevStats(SymbolReference).Formula, udtComputationStats, udtAbbrevSymbolStack, False, dblStdDevSum, dblValueForX, intCharCountPrior + intCharIndex, dblParenthMultiplier * dblAdjacentNum, dblDashMultiplier, dblBracketMultiplier, CarbonOrSiliconReturnCount, intParenthLevelPrevious
                            
                            ' Update the charge to sngChargeSaved
                            udtComputationStats.Charge = sngChargeSaved
                            
                            ' Remove this symbol from the Abbreviation Symbol Stack
                            AbbrevSymbolStackAddRemoveMostRecent udtAbbrevSymbolStack
                            
                            If ErrorParams.ErrorID = 0 Then
                            
                                If blnExpandAbbreviations Then
                                    ' Replace abbreviation with empirical formula
                                    strReplace = AbbrevStats(SymbolReference).Formula
                                    
                                    ' Look for a number after the abbreviation or amino acid
                                    dblAdjacentNum = ParseNum(Mid(strFormula, intCharIndex + intSymbolLength), intNumLength)
                                    CatchParsenumError dblAdjacentNum, intNumLength, intCharIndex, intSymbolLength
                                    
                                    If InStr(strReplace, ">") Then
                                        ' The > symbol means take First Part minus the Second Part
                                        ' If we are parsing a subformula inside parentheses, or if there are
                                        '  symbols (elements, abbreviations, or numbers) after the abbreviation, then
                                        '  we cannot let the > symbol remain in the abbreviation
                                        ' For example, if Jk = C6H5Cl2>HCl
                                        '  and the user enters Jk2 then chooses Expand Abbreviations
                                        ' Then, naively we might replace this with (C6H5Cl2>HCl)2
                                        ' However, this will generate an error because (C6H5Cl2>HCl)2 gets split
                                        '  to (C6H5Cl2 and HCl)2 which will both generate an error
                                        ' The only option is to convert the abbreviation to its empirical formula
                                        If intParenthLevelPrevious > 0 Or intParenthLevel > 0 Or intCharIndex + intSymbolLength <= Len(strFormula) Then
                                            strReplace = ConvertFormulaToEmpirical(strReplace)
                                        End If
                                    End If
                                    
                                    If dblAdjacentNum < 0 Then
                                        ' No number after abbreviation
                                        strFormula = Left(strFormula, intCharIndex - 1) & strReplace & Mid(strFormula, intCharIndex + intSymbolLength)
                                        intSymbolLength = Len(strReplace)
                                        dblAdjacentNum = 1
                                        intAddonCount = intSymbolLength - 1
                                    Else
                                        ' Number after abbreviation -- must put abbreviation in parentheses
                                        ' Parentheses can handle integer or decimal number
                                        strReplace = "(" & strReplace & ")"
                                        strFormula = Left(strFormula, intCharIndex - 1) & strReplace & Mid(strFormula, intCharIndex + intSymbolLength)
                                        intSymbolLength = Len(strReplace)
                                        intAddonCount = intNumLength + intSymbolLength - 1
                                    End If
                                End If
                                
                                If gComputationOptions.CaseConversion = ccConvertCaseUp Then
                                    strFormula = Left(strFormula, intCharIndex - 1) + UCase(Mid(strFormula, intCharIndex, 1)) + Mid(strFormula, intCharIndex + 1)
                                End If
                            End If
                        End If
                    End If
                    intCharIndex = intCharIndex + intAddonCount
                Case Else
                    ' Element not Found
                    If UCase(strChar1) = "X" Then
                        ' X for solver but no preceding bracket
                        ErrorParams.ErrorID = 18
                    Else
                        ErrorParams.ErrorID = 1
                    End If
                    ErrorParams.ErrorPosition = intCharIndex
                End Select
                PrevSymbolReference = SymbolReference
            Case 94         ' ^ (caret)
                dblAdjacentNum = ParseNum(strChar2 & strChar3 & strCharRemain, intNumLength)
                CatchParsenumError dblAdjacentNum, intNumLength, intCharIndex, intSymbolLength
                
                If ErrorParams.ErrorID <> 0 Then
                    ' Problem, don't go on.
                Else
                    strCharVal = Mid(strFormula, intCharIndex + 1 + intNumLength, 1)
                    If Len(strCharVal) > 0 Then intCharAsc = Asc(strCharVal) Else intCharAsc = 0
                    If dblAdjacentNum >= 0 Then
                        If (intCharAsc >= 65 And intCharAsc <= 90) Or (intCharAsc >= 97 And intCharAsc <= 122) Then ' Uppercase A to Z and lowercase a to z
                            blnCaretPresent = True
                            dblCaretVal = dblAdjacentNum
                            intCharIndex = intCharIndex + intNumLength
                        Else
                            ' No letter after isotopic mass
                            ErrorParams.ErrorID = 22: ErrorParams.ErrorPosition = intCharIndex + intNumLength + 1
                        End If
                    Else
                        ' Adjacent number is < 0 or not present
                        ' Record error
                        blnCaretPresent = False
                        If Mid(strFormula, intCharIndex + 1, 1) = "-" Then
                            ' Negative number following caret
                            ErrorParams.ErrorID = 23: ErrorParams.ErrorPosition = intCharIndex + 1
                        Else
                            ' No number following caret
                            ErrorParams.ErrorID = 20: ErrorParams.ErrorPosition = intCharIndex + 1
                        End If
                    End If
                End If
            Case Else
                ' There shouldn't be anything else (except the ~ filler character). If there is, we'll just ignore it
            End Select
            
            If intCharIndex = Len(strFormula) Then
                ' Need to make sure compounds are present after a leading coefficient after a dash
                If dblDashMultiplier > 0 Then
                    If intCharIndex <> intDashPos Then
                        ' Things went fine, no need to set anything
                    Else
                        ' No compounds after leading coefficient after dash
                        ErrorParams.ErrorID = 25: ErrorParams.ErrorPosition = intCharIndex
                    End If
                End If
            End If
            
            If ErrorParams.ErrorID <> 0 Then
                intCharIndex = Len(strFormula)
            End If
            
            intCharIndex = intCharIndex + 1
        Loop While intCharIndex <= Len(strFormula)
    End If
    
    If blnInsideBrackets Then
        If ErrorParams.ErrorID = 0 Then
            ' Missing closing bracket
            ErrorParams.ErrorID = 13: ErrorParams.ErrorPosition = intCharIndex
        End If
    End If
    
    With ErrorParams
        If .ErrorID <> 0 And Len(.ErrorCharacter) = 0 Then
            .ErrorCharacter = strChar1
            .ErrorPosition = .ErrorPosition + intCharCountPrior
        End If
    End With
    

    If LoneCarbonOrSilicon > 1 Then
        ' Correct Charge for number of C and Si
        udtComputationStats.Charge = udtComputationStats.Charge - CLng(LoneCarbonOrSilicon - 1) * 2
        
        CarbonOrSiliconReturnCount = LoneCarbonOrSilicon
    Else
        CarbonOrSiliconReturnCount = 0
    End If
    
    ' Return strFormula, which is possibly now capitalized correctly
    ' It will also contain expanded abbreviations
    ParseFormulaRecursive = strFormula
    Exit Function

ErrorHandler:
    MwtWinDllErrorHandler "MwtWinDll_WorkingRoutines|ParseFormula"
    ErrorParams.ErrorID = -10
    ErrorParams.ErrorPosition = 0
    
End Function

Private Function ParseNum(strWork As String, ByRef intNumLength As Integer, Optional blnAllowNegative As Boolean = False) As Double
    ' Looks for a number and returns it if found
    ' If an error is found, it returns a negative number for the error code and sets intNumLength = 0
    '  -1 = No number
    '  -2 =                                             (unused)
    '  -3 = No number at all or (more likely) no number after decimal point
    '  -4 = More than one decimal point
    
    Dim strWorking As String, strFoundNum As String, intIndex As Integer, intDecPtCount As Integer

    If gComputationOptions.DecimalSeparator = "" Then
        gComputationOptions.DecimalSeparator = DetermineDecimalPoint()
    End If
    
    ' Set intNumLength to -1 for now
    ' If it doesn't get set to 0 (due to an error), it will get set to the
    '   length of the matched number before exiting the sub
    intNumLength = -1
    
    If strWork = "" Then strWork = EMPTY_STRINGCHAR
    If (Asc(Left(strWork, 1)) < 48 Or Asc(Left(strWork, 1)) > 57) And _
       Left(strWork, 1) <> gComputationOptions.DecimalSeparator And _
       Not (Left(strWork, 1) = "-" And blnAllowNegative = True) Then
        intNumLength = 0          ' No number found
        ParseNum = -1
    Else
        ' Start of string is a number or a decimal point, or (if allowed) a negative sign
        For intIndex = 1 To Len(strWork)
            strWorking = Mid(strWork, intIndex, 1)
            If IsNumeric(strWorking) Or strWorking = gComputationOptions.DecimalSeparator Or _
               (blnAllowNegative = True And strWorking = "-") Then
                strFoundNum = strFoundNum & strWorking
            Else
                Exit For
            End If
        Next intIndex
    
        If Len(strFoundNum) = 0 Or strFoundNum = gComputationOptions.DecimalSeparator Then
            ' No number at all or (more likely) no number after decimal point
            strFoundNum = -3
            intNumLength = 0
        Else
            ' Check for more than one decimal point (. or ,)
            intDecPtCount = 0
            For intIndex = 1 To Len(strFoundNum)
                If Mid(strFoundNum, intIndex, 1) = gComputationOptions.DecimalSeparator Then intDecPtCount = intDecPtCount + 1
            Next intIndex
            If intDecPtCount > 1 Then
                ' more than one intDecPtCount
                strFoundNum = -4
                intNumLength = 0
            Else
                ' All is fine
            End If
        End If
    
        If intNumLength < 0 Then intNumLength = Len(strFoundNum)
        ParseNum = CDblSafe(strFoundNum)
    End If
    
End Function

Public Function PlainTextToRtfInternal(strWorkText As String, Optional CalculatorMode As Boolean = False, Optional blnHighlightCharFollowingPercentSign As Boolean = True, Optional blnOverrideErrorID As Boolean = False, Optional lngErrorIDOverride As Long = 0) As String
    Dim strWorkChar As String, strWorkCharPrev As String, strRTF As String
    Dim intCharIndex As Integer, intCharIndex2 As Integer, blnSuperFound As Boolean
    Dim lngErrorID As Long
    
    ' Converts plain text to formatted rtf text.
    ' Rtf string must begin with {{\fonttbl{\f0\fcharset0\fprq2 Times New Roman;}}\pard\plain\fs25
    ' and must end with } or {\fs30  }} if superscript was used
    
    ' "{\super 3}C{\sub 6}H{\sub 6}{\fs30  }}"
    'strRTF = "{{\fonttbl{\f0\fcharset0\fprq2 " & rtfFormula(0).font & ";}}\pard\plain\fs25 "
    ' Old: strRTF = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman\fcharset2 Times New Roman;}{\f3\froman " & lblMWT(0).FontName & ";}}{\colortbl\red0\green0\blue0;\red255\green0\blue0;}\deflang1033\pard\plain\f3\fs25 "
    ' old: strRTF = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman " & lblMWT(0).FontName & ";}{\f3\fswiss\fprq2 System;}}{\colortbl\red0\green0\blue0;\red255\green0\blue0;}\deflang1033\pard\plain\f2\fs25 "
    '                                                            f0                               f1                                 f2                          f3                               f4                      cf0 (black)        cf1 (red)          cf3 (white)
    strRTF = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman " & gComputationOptions.RtfFontName & ";}{\f3\froman Times New Roman;}{\f4\fswiss\fprq2 System;}}{\colortbl\red0\green0\blue0;\red255\green0\blue0;\red255\green255\blue255;}\deflang1033\pard\plain\f2\fs" & Trim(Str(CIntSafeDbl(gComputationOptions.RtfFontSize * 2.5))) & " "
    
    If strWorkText = "" Then
        ' Return a blank RTF string
        PlainTextToRtfInternal = strRTF & "}"
        Exit Function
    End If
    
    blnSuperFound = False
    strWorkCharPrev = ""
    For intCharIndex = 1 To Len(strWorkText)
        strWorkChar = Mid(strWorkText, intCharIndex, 1)
        If strWorkChar = "%" And blnHighlightCharFollowingPercentSign Then
            ' An error was found and marked by a % sign
            ' Highlight the character at the % sign, and remove the % sign
            If intCharIndex = Len(strWorkText) Then
                ' At end of line
                If blnOverrideErrorID And lngErrorIDOverride <> 0 Then
                    lngErrorID = lngErrorIDOverride
                Else
                    lngErrorID = ErrorParams.ErrorID
                End If
                
                Select Case lngErrorID
                Case 2 To 4
                    ' Error involves a parentheses, find last opening parenthesis, (, or opening curly bracket, {
                    For intCharIndex2 = Len(strRTF) To 2 Step -1
                        If Mid(strRTF, intCharIndex2, 1) = "(" Then
                            strRTF = Left(strRTF, intCharIndex2 - 1) & "{\cf1 (}" & Mid(strRTF, intCharIndex2 + 1)
                            Exit For
                        ElseIf Mid(strRTF, intCharIndex2, 1) = "{" Then
                            strRTF = Left(strRTF, intCharIndex2 - 1) & "{\cf1 \{}" & Mid(strRTF, intCharIndex2 + 1)
                            Exit For
                        End If
                    Next intCharIndex2
                Case 13, 15
                    ' Error involves a bracket, find last opening bracket, [
                    For intCharIndex2 = Len(strRTF) To 2 Step -1
                        If Mid(strRTF, intCharIndex2, 1) = "[" Then
                            strRTF = Left(strRTF, intCharIndex2 - 1) & "{\cf1 [}" & Mid(strRTF, intCharIndex2 + 1)
                            Exit For
                        End If
                    Next intCharIndex2
                Case Else
                    ' Nothing to highlight
                End Select
            Else
                ' Highlight next character and skip % sign
                strWorkChar = Mid(strWorkText, intCharIndex + 1, 1)
                ' Handle curly brackets
                If strWorkChar = "{" Or strWorkChar = "}" Then strWorkChar = "\" & strWorkChar
                strRTF = strRTF & "{\cf1 " & strWorkChar & "}"
                intCharIndex = intCharIndex + 1
            End If
        ElseIf strWorkChar = "^" Then
            strRTF = strRTF & "{\super ^}"
            blnSuperFound = True
        ElseIf strWorkChar = "_" Then
            strRTF = strRTF & "{\super _}"
            blnSuperFound = True
        ElseIf strWorkChar = "+" And Not CalculatorMode Then
            strRTF = strRTF & "{\super +}"
            blnSuperFound = True
        ElseIf strWorkChar = EMPTY_STRINGCHAR Then
            ' skip it, the tilde sign is used to add additional height to the formula line when isotopes are used
            ' If it's here from a previous time, we ignore it, adding it at the end if needed (if blnSuperFound = true)
        ElseIf IsNumeric(strWorkChar) Or strWorkChar = gComputationOptions.DecimalSeparator Then
            ' Number or period, so super or subscript it if needed
            If intCharIndex = 1 Then
                ' at beginning of line, so leave it alone. Probably out of place
                strRTF = strRTF & strWorkChar
            ElseIf Not CalculatorMode And (IsCharacter(strWorkCharPrev) Or strWorkCharPrev = ")" Or strWorkCharPrev = "\}" Or strWorkCharPrev = "+" Or strWorkCharPrev = "_" Or Left(Right(strRTF, 6), 3) = "sub") Then
                ' subscript if previous character was a character, parentheses, curly bracket, plus sign, or was already subscripted
                ' But, don't use subscripts in calculator
                strRTF = strRTF & "{\sub " & strWorkChar & "}"
            ElseIf Not CalculatorMode And gComputationOptions.BracketsAsParentheses And strWorkCharPrev = "]" Then
                ' only subscript after closing bracket, ], if brackets are being treated as parentheses
                strRTF = strRTF & "{\sub " & strWorkChar & "}"
            ElseIf Left(Right(strRTF, 8), 5) = "super" Then
                ' if previous character was superscripted, then superscript this number too
                strRTF = strRTF & "{\super " & strWorkChar & "}"
                blnSuperFound = True
            Else
                strRTF = strRTF & strWorkChar
            End If
        ElseIf strWorkChar = " " Then
            ' Ignore it
        Else
            ' Handle curly brackets
            If strWorkChar = "{" Or strWorkChar = "}" Then strWorkChar = "\" & strWorkChar
            strRTF = strRTF & strWorkChar
        End If
        strWorkCharPrev = strWorkChar
    Next intCharIndex

    If blnSuperFound Then
        ' Add an extra tall character, the tilde sign (~, RTF_HEIGHT_ADJUSTCHAR)
        ' It is used to add additional height to the formula line when isotopes are used
        ' It is colored white so the user does not see it
        strRTF = strRTF & "{\fs" & Trim(Str(CIntSafeDbl(gComputationOptions.RtfFontSize * 3))) & "\cf2 " & RTF_HEIGHT_ADJUSTCHAR & "}}"
    Else
        strRTF = strRTF & "}"
    End If
    
    PlainTextToRtfInternal = strRTF

End Function

Public Sub RecomputeAbbreviationMassesInternal()
    ' Recomputes the Mass for all of the loaded abbreviations
    
    Dim lngIndex As Long
    
    For lngIndex = 1 To AbbrevAllCount
        With AbbrevStats(lngIndex)
            .Mass = ComputeFormulaWeight(.Formula)
        End With
    Next lngIndex
End Sub

Public Sub RemoveAllCautionStatementsInternal()
    CautionStatementCount = 0
End Sub

Public Sub RemoveAllAbbreviationsInternal()
    AbbrevAllCount = 0
    ConstructMasterSymbolsList
End Sub

Public Function RemoveAbbreviationInternal(ByVal strAbbreviationSymbol As String) As Long
    ' Look for the abbreviation and remove it
    ' Returns 0 if found and removed; 1 if error
    
    Dim lngIndex As Long
    Dim blnRemoved As Boolean
    
    strAbbreviationSymbol = LCase(strAbbreviationSymbol)
    
    For lngIndex = 1 To AbbrevAllCount
        If LCase(AbbrevStats(lngIndex).Symbol) = strAbbreviationSymbol Then
            RemoveAbbreviationByIDInternal lngIndex
            blnRemoved = True
        End If
    Next lngIndex
    
    If blnRemoved Then
        RemoveAbbreviationInternal = 0
    Else
        RemoveAbbreviationInternal = 1
    End If
End Function

Public Function RemoveAbbreviationByIDInternal(ByVal lngAbbreviationID As Long) As Long
    ' Remove the abbreviation at index lngAbbreviationID
    ' Returns 0 if found and removed; 1 if error
    
    Dim blnRemoved As Boolean
    Dim lngIndexRemove As Long
    
    If lngAbbreviationID >= 1 And lngAbbreviationID <= AbbrevAllCount Then
        For lngIndexRemove = lngAbbreviationID To AbbrevAllCount - 1
            AbbrevStats(lngIndexRemove) = AbbrevStats(lngIndexRemove + 1)
        Next lngIndexRemove
        AbbrevAllCount = AbbrevAllCount - 1
        ConstructMasterSymbolsList
        blnRemoved = True
    Else
        blnRemoved = False
    End If
    
    If blnRemoved Then
        RemoveAbbreviationByIDInternal = 0
    Else
        RemoveAbbreviationByIDInternal = 1
    End If
End Function

Public Function RemoveCautionStatementInternal(ByVal strCautionSymbol As String) As Long
    ' Look for the abbreviation
    ' Returns 0 if found and removed; 1 if error
    
    Dim intIndex As Integer, intIndexRemove As Integer
    Dim blnRemoved As Boolean
    
    For intIndex = 1 To CautionStatementCount
        If CautionStatements(intIndex, 0) = strCautionSymbol Then
            For intIndexRemove = intIndex To CautionStatementCount - 1
                CautionStatements(intIndexRemove, 0) = CautionStatements(intIndexRemove + 1, 0)
                CautionStatements(intIndexRemove, 1) = CautionStatements(intIndexRemove + 1, 1)
            Next intIndexRemove
            CautionStatementCount = CautionStatementCount - 1
            blnRemoved = True
        End If
    Next intIndex
    
    If blnRemoved Then
        RemoveCautionStatementInternal = 0
    Else
        RemoveCautionStatementInternal = 1
    End If
End Function

Public Sub ResetErrorParamsInternal()
    With ErrorParams
        .ErrorCharacter = ""
        .ErrorID = 0
        .ErrorPosition = 0
    End With
End Sub

Public Function ReturnFormattedMassAndStdDev(dblMass As Double, dblStdDev As Double, Optional blnIncludeStandardDeviation As Boolean = True, Optional blnIncludePctSign As Boolean = False) As String
    ' Plan:
      ' Round dblStdDev to 1 final digit.
      ' Round dblMass to the appropriate place based on stddev.

    ' dblMass is the main number
    ' dblStdDev is the standard deviation

    On Local Error GoTo StdErrorHandler

    Dim strResult As String
    Dim strStdDevShort As String
    Dim dblRoundedStdDev As Double, dblRoundedMain As Double
    Dim strWork As String, dblWork As Double, intExponentValue As Integer
    Dim strPctSign As String
    
    ' blnIncludePctSign is True when formatting Percent composition values
    If blnIncludePctSign Then
        strPctSign = "%"
    Else
        strPctSign = ""
    End If
    
    If Val(dblStdDev) = 0 Then
        ' Standard deviation value is 0; simply return the result
        strResult = Format(dblMass, "0.0####") & strPctSign & " (" & Chr$(177) & "0)"
        
        ' dblRoundedMain is used later, must copy dblMass to it
        dblRoundedMain = dblMass
    Else
        
        ' First round dblStdDev to show just one number
        dblRoundedStdDev = CDbl(Format(dblStdDev, "0E+000"))
        
        ' Now round dblMass
        ' Simply divide dblMass by 10^Exponent of the Standard Deviation
        ' Next round
        ' Now multiply to get back the rounded dblMass
        strWork = Format(dblStdDev, "0E+000")
        strStdDevShort = Left(strWork, 1)
        
        intExponentValue = CIntSafe(Right(strWork, 4))
        dblWork = dblMass / 10 ^ intExponentValue
        dblWork = Round(dblWork, 0)
        dblRoundedMain = dblWork * 10 ^ intExponentValue
        
        strWork = Format(dblRoundedMain, "0.0##E+00")
        
        If gComputationOptions.StdDevMode = smShort Then
            ' StdDevType Short (Type 0)
            strResult = CStr(dblRoundedMain)
            If blnIncludeStandardDeviation Then
                strResult = strResult & "(" & Chr$(177) & strStdDevShort & ")"
            End If
            strResult = strResult & strPctSign
        ElseIf gComputationOptions.StdDevMode = smScientific Then
            ' StdDevType Scientific (Type 1)
            strResult = CStr(dblRoundedMain) & strPctSign
            If blnIncludeStandardDeviation Then
                strResult = strResult & " (" & Chr$(177) & Trim(Format(dblStdDev, "0.000E+00")) & ")"
            End If
        Else
            ' StdDevType Decimal
            strResult = Format(dblMass, "0.0####") & strPctSign
            If blnIncludeStandardDeviation Then
                strResult = strResult & " (" & Chr$(177) & Trim(CStr(dblRoundedStdDev)) & ")"
            End If
        End If
    End If

    ReturnFormattedMassAndStdDev = strResult

StdStart:
    'This code skips the error handler
    Exit Function

StdErrorHandler:
    MwtWinDllErrorHandler "MwtWinDll_WorkingRoutines|ReturnFormattedMassAndStdDev"
    ErrorParams.ErrorID = -10
    ErrorParams.ErrorPosition = 0
    Resume StdStart
End Function

Public Function SetAbbreviationInternal(ByVal strSymbol As String, ByVal strFormula As String, ByVal sngCharge As Single, ByVal blnIsAminoAcid As Boolean, Optional ByVal strOneLetterSymbol As String = "", Optional strComment As String = "", Optional blnValidateFormula As Boolean = True) As Long
    ' Adds a new abbreviation or updates an existing one (based on strSymbol)
    ' Returns 0 if successful, otherwise, returns an error ID
    
    Dim blnAlreadyPresent As Boolean
    Dim lngIndex As Long, lngAbbrevID As Long
    
    ' See if the abbreviation is alrady present
    blnAlreadyPresent = False
    For lngIndex = 1 To AbbrevAllCount
        If UCase(AbbrevStats(lngIndex).Symbol) = UCase(strSymbol) Then
            blnAlreadyPresent = True
            lngAbbrevID = lngIndex
            Exit For
        End If
    Next lngIndex

    ' AbbrevStats is a 1-based array
    If Not blnAlreadyPresent Then
        If AbbrevAllCount < MAX_ABBREV_COUNT Then
            lngAbbrevID = AbbrevAllCount + 1
            AbbrevAllCount = AbbrevAllCount + 1
        Else
            ' Too many abbreviations
            ErrorParams.ErrorID = 196
        End If
    End If
    
    If lngAbbrevID >= 1 Then
        SetAbbreviationByIDInternal lngAbbrevID, strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment, blnValidateFormula
    End If

    SetAbbreviationInternal = ErrorParams.ErrorID

End Function

Public Function SetAbbreviationByIDInternal(ByVal intAbbrevID As Integer, ByVal strSymbol As String, ByVal strFormula As String, ByVal sngCharge As Single, ByVal blnIsAminoAcid As Boolean, Optional ByVal strOneLetterSymbol As String = "", Optional strComment As String = "", Optional blnValidateFormula As Boolean = True) As Long
    ' Adds a new abbreviation or updates an existing one (based on intAbbrevID)
    ' If intAbbrevID < 1 then adds as a new abbreviation
    ' Returns 0 if successful, otherwise, returns an error ID

    Dim udtComputationStats As udtComputationStatsType
    Dim udtAbbrevSymbolStack As udtAbbrevSymbolStackType
    Dim blnInvalidSymbolOrFormula As Boolean
    Dim eSymbolMatchType As smtSymbolMatchTypeConstants
    Dim intSymbolReference As Integer
    
    ResetErrorParamsInternal
    
    InitializeComputationStats udtComputationStats
    InitializeAbbrevSymbolStack udtAbbrevSymbolStack
    
    If Len(strSymbol) < 1 Then
        ' Symbol length is 0
        ErrorParams.ErrorID = 192
    ElseIf Len(strSymbol) > MAX_ABBREV_LENGTH Then
        ' Abbreviation symbol too long
        ErrorParams.ErrorID = 190
    Else
        If IsStringAllLetters(strSymbol) Then
            If Len(strFormula) > 0 Then
                ' Convert symbol to proper case mode
                strSymbol = UCase(Left(strSymbol, 1)) & LCase(Mid(strSymbol, 2))
            
                ' If intAbbrevID is < 1 or larger than AbbrevAllCount, then define it
                If intAbbrevID < 1 Or intAbbrevID > AbbrevAllCount + 1 Then
                    If AbbrevAllCount < MAX_ABBREV_COUNT Then
                        AbbrevAllCount = AbbrevAllCount + 1
                        intAbbrevID = AbbrevAllCount
                    Else
                        ' Too many abbreviations
                        ErrorParams.ErrorID = 196
                        intAbbrevID = -1
                    End If
                End If
            
                If intAbbrevID >= 1 Then
                    ' Make sure the abbreviation doesn't match one of the standard elements
                    eSymbolMatchType = CheckElemAndAbbrev(strSymbol, intSymbolReference)
                    
                    If eSymbolMatchType = smtElement Then
                        If ElementStats(intSymbolReference).Symbol = strSymbol Then
                            blnInvalidSymbolOrFormula = True
                        End If
                    End If
                    
                    If Not blnInvalidSymbolOrFormula And blnValidateFormula Then
                        ' Make sure the abbreviation's formula is valid
                        ' This will also auto-capitalize the formula if auto-capitalize is turned on
                        strFormula = ParseFormulaRecursive(strFormula, udtComputationStats, udtAbbrevSymbolStack, False, 0)
                        
                        If ErrorParams.ErrorID <> 0 Then
                            ' An error occurred while parsing
                            ' Already present in ErrorParams.ErrorID
                            ' We'll still add the formula, but mark it as invalid
                            blnInvalidSymbolOrFormula = True
                        End If
                    End If
                    
                    AddAbbreviationWork intAbbrevID, strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment, blnInvalidSymbolOrFormula
                    
                    ConstructMasterSymbolsList
                End If
            Else
                ' Invalid formula (actually, blank formula)
                ErrorParams.ErrorID = 160
            End If
        Else
            ' Symbol does not just contain letters
            ErrorParams.ErrorID = 194
        End If
    End If

    SetAbbreviationByIDInternal = ErrorParams.ErrorID

End Function

Public Function SetCautionStatementInternal(strSymbolCombo As String, strNewCautionStatement As String) As Long
    ' Adds a new caution statement or updates an existing one (based on strSymbolCombo)
    ' Returns 0 if successful, otherwise, returns an Error ID
    
    Dim blnAlreadyPresent As Boolean, intIndex As Integer
    
    ResetErrorParamsInternal
    
    If Len(strSymbolCombo) >= 1 And Len(strSymbolCombo) <= MAX_ABBREV_LENGTH Then
        ' Make sure all the characters in strSymbolCombo are letters
        If IsStringAllLetters(strSymbolCombo) Then
            If Len(strNewCautionStatement) > 0 Then
                ' See if strSymbolCombo is present in CautionStatements()
                For intIndex = 1 To CautionStatementCount
                    If CautionStatements(intIndex, 0) = strSymbolCombo Then
                        blnAlreadyPresent = True
                        Exit For
                    End If
                Next intIndex
                
                ' Caution statements is a 0-based array
                If Not blnAlreadyPresent Then
                    If CautionStatementCount < MAX_CAUTION_STATEMENTS Then
                        CautionStatementCount = CautionStatementCount + 1
                        intIndex = CautionStatementCount
                    Else
                        ' Too many caution statements
                        ErrorParams.ErrorID = 1215
                        intIndex = -1
                    End If
                End If
                
                If intIndex >= 1 Then
                    CautionStatements(intIndex, 0) = strSymbolCombo
                    CautionStatements(intIndex, 1) = strNewCautionStatement
                End If
            Else
                ' Caution description length is 0
                ErrorParams.ErrorID = 1210
            End If
        Else
            ' Caution symbol doesn't just contain letters
            ErrorParams.ErrorID = 1205
        End If
    Else
        ' Symbol length is 0 or is greater than MAX_ABBREV_LENGTH
        ErrorParams.ErrorID = 1200
    End If

    SetCautionStatementInternal = ErrorParams.ErrorID

End Function

Public Function SetChargeCarrierMassInternal(dblMass As Double)
    mChargeCarrierMass = dblMass
End Function

Public Function SetElementInternal(strSymbol As String, dblMass As Double, dblUncertainty As Double, sngCharge As Single, Optional blnRecomputeAbbreviationMasses As Boolean = True) As Long
    ' Used to update the values for a single element (based on strSymbol)
    ' Set blnRecomputeAbbreviationMasses to False if updating several elements
    
    Dim intIndex As Integer, blnFound As Boolean
    
    For intIndex = 1 To ELEMENT_COUNT
        If LCase(strSymbol) = LCase(ElementStats(intIndex).Symbol) Then
            With ElementStats(intIndex)
                .Mass = dblMass
                .Uncertainty = dblUncertainty
                .Charge = sngCharge
            End With
            blnFound = True
            Exit For
        End If
    Next intIndex
    
    If blnFound Then
        If blnRecomputeAbbreviationMasses Then RecomputeAbbreviationMassesInternal
        SetElementInternal = 0
    Else
        SetElementInternal = 1
    End If
End Function

Public Function SetElementIsotopesInternal(ByVal strSymbol As String, ByVal intIsotopeCount As Integer, ByRef dblIsotopeMassesOneBased() As Double, ByRef sngIsotopeAbundancesOneBased() As Single) As Long

    Dim intIndex As Integer, intIsotopeindex As Integer, blnFound As Boolean
    
    For intIndex = 1 To ELEMENT_COUNT
        If LCase(strSymbol) = LCase(ElementStats(intIndex).Symbol) Then
            With ElementStats(intIndex)
                If intIsotopeCount < 0 Then intIsotopeCount = 0
                .IsotopeCount = intIsotopeCount
                For intIsotopeindex = 1 To .IsotopeCount
                    If intIsotopeindex > MAX_ISOTOPES Then Exit For
                    .Isotopes(intIsotopeindex).Mass = dblIsotopeMassesOneBased(intIsotopeindex)
                    .Isotopes(intIsotopeindex).Abundance = sngIsotopeAbundancesOneBased(intIsotopeindex)
                Next intIsotopeindex
            End With
            blnFound = True
            Exit For
        End If
    Next intIndex
    
    If blnFound Then
        SetElementIsotopesInternal = 0
    Else
        SetElementIsotopesInternal = 1
    End If
End Function

Public Sub SetElementModeInternal(NewElementMode As emElementModeConstantsPrivate, Optional blnMemoryLoadElementValues As Boolean = True)
    ' The only time you would want blnMemoryLoadElementValues to be False is if you're
    '  manually setting element weight values, but want to let the software know that
    '  they're average, isotopic, or integer values
    
On Error GoTo SetElementModeInternalErrorHandler

    If NewElementMode >= emAverageMass And NewElementMode <= emIntegerMass Then
        If NewElementMode <> mCurrentElementMode Or blnMemoryLoadElementValues Then
            mCurrentElementMode = NewElementMode
            
            If blnMemoryLoadElementValues Then
                MemoryLoadElements mCurrentElementMode
            End If
            
            ConstructMasterSymbolsList
            RecomputeAbbreviationMassesInternal
        End If
    End If
    
    Exit Sub

SetElementModeInternalErrorHandler:
    GeneralErrorHandler "MWPeptideClass.SetElementModeInternal", Err.Number

End Sub

Public Function SetMessageStatementInternal(lngMessageID As Long, strNewMessage As String) As Long
    ' Used to replace the default message strings with foreign language equivalent ones
    ' Returns 0 if success; 1 if failure
    If lngMessageID >= 1 And lngMessageID <= MESSAGE_STATEMENT_DIMCOUNT And Len(strNewMessage) > 0 Then
        MessageStatements(lngMessageID) = strNewMessage
        SetMessageStatementInternal = 0
    Else
        SetMessageStatementInternal = 1
    End If
End Function

Private Sub ShellSortSymbols(lngLowIndex As Long, lngHighIndex As Long)
    Dim lngIndex As Long

    ReDim PointerArray(lngHighIndex) As Long
    ReDim SymbolsStore(lngHighIndex, 1) As String

    ' MasterSymbolsList starts at lngLowIndex
    For lngIndex = lngLowIndex To lngHighIndex
        PointerArray(lngIndex) = lngIndex
    Next lngIndex

    ShellSortSymbolsWork PointerArray(), lngLowIndex, lngHighIndex

    ' Reassign MasterSymbolsList array according to pointerarray order
    ' First, copy to a temporary array (I know it eats up memory, but I have no choice)
    For lngIndex = lngLowIndex To lngHighIndex
        SymbolsStore(lngIndex, 0) = MasterSymbolsList(lngIndex, 0)
        SymbolsStore(lngIndex, 1) = MasterSymbolsList(lngIndex, 1)
    Next lngIndex

    ' Now, put them back into the MasterSymbolsList() array in the correct order
    ' Use pointerarray() for this
    For lngIndex = lngLowIndex To lngHighIndex
        MasterSymbolsList(lngIndex, 0) = SymbolsStore(PointerArray(lngIndex), 0)
        MasterSymbolsList(lngIndex, 1) = SymbolsStore(PointerArray(lngIndex), 1)
    Next lngIndex

End Sub

Private Sub ShellSortSymbolsWork(ByRef PointerArray() As Long, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long)
    ' Sort the list using a shell sort
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim lngPointerSwap As Long
    Dim Length1 As Long, Length2 As Long
    
    ' Sort PointerArray[lngLowIndex..lngHighIndex] by comparing
    '   Len(MasterSymbolsList(PointerArray(x)) and sorting by decreasing length
    ' If same length, sort alphabetically (increasing)

    ' Compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' Sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            lngPointerSwap = PointerArray(lngIndex)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                ' Sort by decreasing length
                Length1 = Len(MasterSymbolsList(PointerArray(lngIndexCompare), 0))
                Length2 = Len(MasterSymbolsList(lngPointerSwap, 0))
                If Length1 > Length2 Then Exit For
                ' If same length, sort alphabetically
                If Length1 = Length2 Then
                    If UCase(MasterSymbolsList(PointerArray(lngIndexCompare), 0)) <= UCase(MasterSymbolsList(lngPointerSwap, 0)) Then Exit For
                End If
                PointerArray(lngIndexCompare + lngIncrement) = PointerArray(lngIndexCompare)
            Next lngIndexCompare
            PointerArray(lngIndexCompare + lngIncrement) = lngPointerSwap
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop

End Sub

Public Sub SetShowErrorMessageDialogs(blnValue As Boolean)
    mShowErrorMessageDialogs = blnValue
End Sub

Public Function ShowErrorMessageDialogs() As Boolean
    ShowErrorMessageDialogs = mShowErrorMessageDialogs
End Function

Public Sub SortAbbreviationsInternal()
    Dim lngLowIndex As Long, lngHighIndex As Long
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim udtCompare As udtAbbrevStatsType
    
    lngCount = AbbrevAllCount
    lngLowIndex = 1
    lngHighIndex = lngCount

    ' sort array[lngLowIndex..lngHighIndex]

    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            udtCompare = AbbrevStats(lngIndex)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If AbbrevStats(lngIndexCompare).Symbol <= udtCompare.Symbol Then Exit For
                AbbrevStats(lngIndexCompare + lngIncrement) = AbbrevStats(lngIndexCompare)
            Next lngIndexCompare
            AbbrevStats(lngIndexCompare + lngIncrement) = udtCompare
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop

    ' Need to re-construct the master symbols list
    ConstructMasterSymbolsList
    
End Sub

Private Function SpacePadFront(ByRef strWork As String, intLength As Integer) As String
    
    Do While Len(strWork) < intLength
        strWork = " " & strWork
    Loop
    SpacePadFront = strWork
    
End Function

Public Function ValidateAllAbbreviationsInternal() As Long
    ' Checks the formula of all abbreviations to make sure it's valid
    ' Marks any abbreviations as Invalid if a problem is found or a circular reference exists
    ' Returns a count of the number of invalid abbreviations found

    Dim intAbbrevIndex As Integer
    Dim intInvalidAbbreviationCount As Integer
    
    For intAbbrevIndex = 1 To AbbrevAllCount
        With AbbrevStats(intAbbrevIndex)
            SetAbbreviationByIDInternal intAbbrevIndex, .Symbol, .Formula, .Charge, .IsAminoAcid, .OneLetterSymbol, .Comment, True
            If .InvalidSymbolOrFormula Then
                intInvalidAbbreviationCount = intInvalidAbbreviationCount + 1
            End If
        End With
    Next intAbbrevIndex
    
    ValidateAllAbbreviationsInternal = intInvalidAbbreviationCount
    
End Function
