VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MWPeptideClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

' Molecular Weight Calculator routines with ActiveX Class interfaces: MWPeptideClass
' Written by Matthew Monroe in Richland, WA

Private Const RESIDUE_DIM_CHUNK = 50
Private Const MAX_MODIFICATIONS = 6             ' Maximum number of modifications for a single residue
Private Const UNKNOWN_SYMBOL = "Xxx"
Private Const UNKNOWN_SYMBOL_ONE_LETTER = "X"

Private Const TERMINII_SYMBOL = "-"
Private Const TRYPTIC_RULE_RESIDUES = "KR"
Private Const TRYPTIC_EXCEPTION_RESIDUES = "P"

Private Const SHOULDER_ION_PREFIX = "Shoulder-"

Public Enum ctgCTerminusGroupConstants
    ctgHydroxyl = 0
    ctgAmide = 1
    ctgNone = 2
End Enum

Public Enum ntgNTerminusGroupConstants
    ntgHydrogen = 0
    ntgHydrogenPlusProton = 1
    ntgAcetyl = 2
    ntgPyroGlu = 3
    ntgCarbamyl = 4
    ntgPTC = 5
    ntgNone = 6
End Enum

Private Const ION_TYPE_MAX As Integer = 4
Public Enum itIonTypeConstants
    itAIon = 0
    itBIon = 1
    itYIon = 2
    itCIon = 3
    itZIon = 4
End Enum

Private Type udtModificationSymbolType
    Symbol As String                        ' Symbol used for modification in formula; may be 1 or more characters; for example: + ++ * ** etc.
    ModificationMass As Double              ' Normally positive, but could be negative
    IndicatesPhosphorylation As Boolean     ' When true, then this symbol means a residue is phosphorylated
    Comment As String
End Type

Private Type udtResidueType
    Symbol As String                        ' 3 letter symbol
    Mass As Double                          ' The mass of the residue alone (excluding any modification)
    MassWithMods As Double                  ' The mass of the residue, including phosphorylation or any modification
    IonMass(ION_TYPE_MAX) As Double       ' 0-based array; the masses that the a, b, and y ions ending/starting with this residue will produce in the mass spectrum (includes H+)
    Phosphorylated As Boolean               ' Technically, only Ser, Thr, or Tyr residues can be phosphorylated (H3PO4), but if the user phosphorylates other residues, we'll allow that
    ModificationIDCount As Integer
    ModificationIDs(MAX_MODIFICATIONS) As Long  ' 1-based array
End Type

Private Type udtTerminusType
    Formula As String
    Mass As Double
    PrecedingResidue As udtResidueType      ' If the peptide sequence is part of a protein, the user can record the final residue of the previous peptide sequence here
    FollowingResidue As udtResidueType      ' If the peptide sequence is part of a protein, the user can record the first residue of the next peptide sequence here
End Type

Public Type udtFragmentionSpectrumIntensitiesType
    IonType(ION_TYPE_MAX) As Double       ' 0-based array
    BYIonShoulder As Double                 ' If > 0 then shoulder ions will be created by B and Y ions
    NeutralLoss As Double
End Type

' Note: A ions can have ammonia and phosphate loss, but not water loss, so this is set to false by default
'       The graphical version of MwtWin does not allow this to be overridden, but a programmer could do so via a call to this Dll
Public Type udtIonTypeOptionsType
    ShowIon As Boolean
    NeutralLossWater As Boolean
    NeutralLossAmmonia As Boolean
    NeutralLossPhosphate As Boolean
End Type

Public Type udtFragmentationSpectrumOptionsType
    IntensityOptions As udtFragmentionSpectrumIntensitiesType
    IonTypeOptions(ION_TYPE_MAX) As udtIonTypeOptionsType
    DoubleChargeIonsShow As Boolean
    DoubleChargeIonsThreshold As Single
    TripleChargeIonsShow As Boolean
    TripleChargeIonsThreshold As Single
End Type

Public Type udtFragmentationSpectrumDataType
    Mass As Double
    Intensity As Double
    Symbol As String                        ' The symbol, with the residue number (e.g. y1, y2, b3-H2O, Shoulder-y1, etc.)
    SymbolGeneric As String                 ' The symbol, without the residue number (e.g. a, b, y, b++, Shoulder-y, etc.)
    SourceResidueNumber As Long             ' The residue number that resulted in this mass
    SourceResidueSymbol3Letter As String    ' The residue symbol that resulted in this mass
    Charge As Integer
    IonType As itIonTypeConstants
    IsShoulderIon As Boolean                ' B and Y ions can have Shoulder ions at +-1
End Type

' Note: A peptide goes from N to C, eg. HGlyLeuTyrOH has N-Terminus = H and C-Terminus = OH
' Residue 1 would be Gly, Residue 2 would be Leu, Residue 3 would be Tyr
Private Residues() As udtResidueType        ' 1-based array
Private ResidueCount As Long
Private ResidueCountDimmed As Long

' ModificationSymbols() holds a list of the potential modification symbols and the mass of each modification
' Modification symbols can be 1 or more letters long
Private ModificationSymbols() As udtModificationSymbolType      ' 1-based array
Private ModificationSymbolCount As Long
Private ModificationSymbolCountDimmed As Long

Private mNTerminus As udtTerminusType       ' Formula on the N-Terminus
Private mCTerminus As udtTerminusType       ' Formula on the C-Terminus
Private mTotalMass As Double

Private mWaterLossSymbol As String          ' -H2O
Private mAmmoniaLossSymbol As String        ' -NH3
Private mPhosphoLossSymbol As String        ' -H3PO4

Private mFragSpectrumOptions As udtFragmentationSpectrumOptionsType

Private dblHOHMass As Double
Private dblNH3Mass As Double
Private dblH3PO4Mass As Double
Private dblPhosphorylationMass As Double    ' H3PO4 minus HOH = 79.9663326
Private dblHydrogenMass As Double           ' Mass of hydrogen
Private dblChargeCarrierMass As Double      ' H minus one electron

Private dblImmoniumMassDifference As Double ' CO minus H = 26.9871

Private dblHistidineFW As Double            ' 110
Private dblPhenylalanineFW As Double        ' 120
Private dblTyrosineFW As Double             ' 136

Private mDelayUpdateResidueMass As Boolean
'

Private Sub AppendDataToFragSpectrum(lngIonCount As Long, FragSpectrumWork() As udtFragmentationSpectrumDataType, sngMass As Single, sngIntensity As Single, strIonSymbol As String, strIonSymbolGeneric As String, lngSourceResidue As Long, strSourceResidueSymbol3Letter As String, intCharge As Integer, eIonType As itIonTypeConstants, blnIsShoulderIon As Boolean)
On Error GoTo AppendDataPointErrorHandler
    
    If lngIonCount > UBound(FragSpectrumWork()) Then
        ' This shouldn't happen
        Debug.Assert False
        ReDim Preserve FragSpectrumWork(UBound(FragSpectrumWork()) + 10)
    End If
    
    With FragSpectrumWork(lngIonCount)
        .Mass = sngMass
        .Intensity = sngIntensity
        .Symbol = strIonSymbol
        .SymbolGeneric = strIonSymbolGeneric
        .SourceResidueNumber = lngSourceResidue
        .SourceResidueSymbol3Letter = strSourceResidueSymbol3Letter
        .Charge = intCharge
        .IonType = eIonType
        .IsShoulderIon = blnIsShoulderIon
    End With
    lngIonCount = lngIonCount + 1
    Exit Sub

AppendDataPointErrorHandler:
    Debug.Print Err.Description
    Debug.Assert False
End Sub

Private Function CheckForModifications(strPartialSequence As String, lngResidueIndex As Long, Optional blnAddMissingModificationSymbols As Boolean = False) As Long
    ' Looks at strPartialSequence to see if it contains 1 or more modifications
    ' If any modification symbols are found, the modification is recorded in .ModificationIDs()
    ' If all or part of the modification symbol is not found in ModificationSymbols(), then a new entry
    '  is added to ModificationSymbols()
    ' Returns the total length of all modifications found
    
    Dim lngCompareIndex As Long, lngSequenceStrLength As Long
    Dim strModSymbolGroup As String
    Dim lngModificationID As Long, lngModSymbolLengthTotal As Long
    Dim strTestChar As String
    Dim lngIndex As Long
    Dim blnMatchFound As Boolean
    
    lngSequenceStrLength = Len(strPartialSequence)
    
    ' Find the entire group of potential modification symbols
    strModSymbolGroup = ""
    lngCompareIndex = 1
    Do While lngCompareIndex <= lngSequenceStrLength
        strTestChar = Mid(strPartialSequence, lngCompareIndex, 1)
        If IsModSymbolInternal(strTestChar) Then
            strModSymbolGroup = strModSymbolGroup & strTestChar
        Else
            Exit Do
        End If
        lngCompareIndex = lngCompareIndex + 1
    Loop
    
    lngModSymbolLengthTotal = Len(strModSymbolGroup)
    Do While Len(strModSymbolGroup) > 0
        ' Step through strModSymbolGroup to see if all of it or parts of it match any of the defined
        '  modification symbols
        
        blnMatchFound = False
        For lngIndex = Len(strModSymbolGroup) To 1 Step -1
            ' See if the modification is already defined
            lngModificationID = GetModificationSymbolID(Left(strModSymbolGroup, lngIndex))
            If lngModificationID > 0 Then
                blnMatchFound = True
                Exit For
            End If
        Next lngIndex
        
        If Not blnMatchFound Then
            If blnAddMissingModificationSymbols Then
                ' Add strModSymbolGroup as a new modification, using a mass of 0 since we don't know the modification mass
                SetModificationSymbol strModSymbolGroup, 0, False, ""
                blnMatchFound = True
            Else
                ' Ignore the modification
                strModSymbolGroup = 0
            End If
            strModSymbolGroup = ""
        End If
        
        If blnMatchFound Then
            ' Record the modification for this residue
            With Residues(lngResidueIndex)
                If .ModificationIDCount < MAX_MODIFICATIONS Then
                    .ModificationIDCount = .ModificationIDCount + 1
                    .ModificationIDs(.ModificationIDCount) = lngModificationID
                    If ModificationSymbols(lngModificationID).IndicatesPhosphorylation Then
                        .Phosphorylated = True
                    End If
                End If
            End With
            
            If lngIndex < Len(strModSymbolGroup) Then
                ' Remove the matched portion from strModSymbolGroup and test again
                strModSymbolGroup = Mid(strModSymbolGroup, lngIndex + 1)
            Else
                strModSymbolGroup = ""
            End If
        End If
    Loop
    
    CheckForModifications = lngModSymbolLengthTotal
    
End Function

Private Function ComputeMaxIonsPerResidue() As Integer
    ' Estimate the total ions per residue that will be created
    ' This number will nearly always be much higher than the number of ions that will actually
    '  be stored for a given sequence, since not all will be doubly charged, and not all will show
    '  all of the potential neutral losses
    
    Dim eIonIndex As itIonTypeConstants, intIonCount As Integer
    
    intIonCount = 0
    With mFragSpectrumOptions
        For eIonIndex = 0 To ION_TYPE_MAX
            If .IonTypeOptions(eIonIndex).ShowIon Then
                intIonCount = intIonCount + 1
                
                If Abs(.IntensityOptions.BYIonShoulder) > 0 Then
                    If eIonIndex = itIonTypeConstants.itBIon Or eIonIndex = itIonTypeConstants.itYIon Or _
                       eIonIndex = itIonTypeConstants.itCIon Or eIonIndex = itIonTypeConstants.itZIon Then
                        intIonCount = intIonCount + 2
                    End If
                End If
                
                If .IonTypeOptions(eIonIndex).NeutralLossAmmonia Then intIonCount = intIonCount + 1
                If .IonTypeOptions(eIonIndex).NeutralLossPhosphate Then intIonCount = intIonCount + 1
                If .IonTypeOptions(eIonIndex).NeutralLossWater Then intIonCount = intIonCount + 1
            End If
        Next eIonIndex
    
        ' Double Charge ions could be created for all ions, so simply double intIonCount
        If .DoubleChargeIonsShow Then
            intIonCount = intIonCount * 2
        End If
        
        If .TripleChargeIonsShow Then
            intIonCount = intIonCount * 2
        End If
        
    
    End With
    
    ComputeMaxIonsPerResidue = intIonCount

End Function

Private Function FillResidueStructureUsingSymbol(strSymbol As String, Optional blnUse3LetterCode As Boolean = True) As udtResidueType
    ' Returns a variable of type udtResidueType containing strSymbol as the residue symbol
    ' If strSymbol is a valid amino acid type, then also updates udtResidue with the default information
    
    Dim strSymbol3Letter As String
    Dim lngAbbrevID As Long
    Dim udtResidue As udtResidueType
    
    If Len(strSymbol) > 0 Then
        If blnUse3LetterCode Then
            strSymbol3Letter = strSymbol
        Else
            strSymbol3Letter = GetAminoAcidSymbolConversionInternal(strSymbol, True)
            If strSymbol3Letter = "" Then strSymbol3Letter = strSymbol
        End If
        
        lngAbbrevID = GetAbbreviationIDInternal(strSymbol3Letter, True)
    Else
        lngAbbrevID = 0
    End If
    
    With udtResidue
        .Symbol = strSymbol3Letter
        .ModificationIDCount = 0
        .Phosphorylated = False
        If lngAbbrevID > 0 Then
            .Mass = GetAbbreviationMass(lngAbbrevID)
        Else
            .Mass = 0
        End If
        .MassWithMods = .Mass
    End With
    
    FillResidueStructureUsingSymbol = udtResidue
    
End Function

'Public Function GetFragmentationMasses(ByVal lngMaxIonCount As Long, ByRef sngIonMassesZeroBased() As Single, ByRef sngIonIntensitiesZeroBased() As Single, ByRef strIonSymbolsZeroBased() As String) As Long
Public Function GetFragmentationMasses(ByRef udtFragSpectrum() As udtFragmentationSpectrumDataType) As Long
    ' Returns the number of ions in FragSpectrumWork()

    Const MAX_CHARGE As Integer = 3

    Dim lngResidueIndex As Long, intChargeIndex As Integer, intShoulderIndex As Integer
    Dim eIonType As itIonTypeConstants
    Dim lngIndex As Long
    Dim lngPredictedIonCount As Long, lngIonCount As Long
    Dim sngIonIntensities(ION_TYPE_MAX) As Single
    Dim sngIonShoulderIntensity As Single, sngNeutralLossIntensity As Single
    
    Dim blnShowCharge() As Boolean
    Dim sngChargeThreshold() As Single
    
    Dim sngBaseMass As Single, sngConvolutedMass As Single, sngObservedMass As Single
    Dim strResidues As String
    Dim blnPhosphorylated As Boolean
    Dim sngIntensity As Single
    Dim strIonSymbol As String, strIonSymbolGeneric As String
    Dim FragSpectrumWork() As udtFragmentationSpectrumDataType
    Dim PointerArray() As Long
    
    If ResidueCount = 0 Then
        ' No residues
        GetFragmentationMasses = 0
        Exit Function
    End If

    ReDim blnShowCharge(MAX_CHARGE)
    ReDim sngChargeThreshold(MAX_CHARGE)
    
    ' Copy some of the values from mFragSpectrumOptions to local variables to make things easier to read
    With mFragSpectrumOptions
        For eIonType = 0 To ION_TYPE_MAX
            sngIonIntensities(eIonType) = .IntensityOptions.IonType(eIonType)
        Next eIonType
        sngIonShoulderIntensity = .IntensityOptions.BYIonShoulder
        sngNeutralLossIntensity = .IntensityOptions.NeutralLoss
        
        If MAX_CHARGE >= 2 Then
            blnShowCharge(2) = .DoubleChargeIonsShow
            sngChargeThreshold(2) = .DoubleChargeIonsThreshold
        End If
        
        If MAX_CHARGE >= 3 Then
            blnShowCharge(3) = .TripleChargeIonsShow
            sngChargeThreshold(3) = .TripleChargeIonsThreshold
        End If
    End With

    ' Populate sngIonMassesZeroBased() and sngIonIntensitiesZeroBased()
    ' Put ion descriptions in strIonSymbolsZeroBased
    lngPredictedIonCount = GetFragmentationSpectrumRequiredDataPoints()
    
    If lngPredictedIonCount = 0 Then lngPredictedIonCount = ResidueCount
    ReDim FragSpectrumWork(lngPredictedIonCount)
    
    ' Need to update the residue masses in case the modifications have changed
    UpdateResidueMasses
    
    lngIonCount = 0
    For lngResidueIndex = 1 To ResidueCount
        With Residues(lngResidueIndex)
            
            For eIonType = 0 To ION_TYPE_MAX
                If mFragSpectrumOptions.IonTypeOptions(eIonType).ShowIon Then
                    If (lngResidueIndex = 1 Or lngResidueIndex = ResidueCount) And (eIonType = itIonTypeConstants.itAIon Or eIonType = itIonTypeConstants.itBIon Or eIonType = itIonTypeConstants.itCIon) Then
                        ' Don't include a or b ions in the output masses
                    Else
                        
                        ' Ion is used
                        sngBaseMass = .IonMass(eIonType)     ' Already in the H+ state
                        sngIntensity = sngIonIntensities(eIonType)
                        
                        ' Get the list of residues preceding or following this residue
                        ' Note that the residue symbols are separated by a space to avoid accidental matching by the InStr() functions below
                        strResidues = GetInternalResidues(lngResidueIndex, eIonType, blnPhosphorylated)
                        
                        For intChargeIndex = 1 To MAX_CHARGE
                            If intChargeIndex = 1 Or (intChargeIndex > 1 And blnShowCharge(intChargeIndex)) Then
                                If intChargeIndex = 1 Then
                                    sngConvolutedMass = sngBaseMass
                                Else
                                    ' Compute mass at higher charge
                                    sngConvolutedMass = ConvoluteMassInternal(sngBaseMass, 1, intChargeIndex, dblChargeCarrierMass)
                                End If
                                
                                If intChargeIndex > 1 And sngBaseMass < sngChargeThreshold(intChargeIndex) Then
                                    ' BaseMass is below threshold, do not add to Predicted Spectrum
                                Else
                                    ' Add ion to Predicted Spectrum
                                    
                                    ' Y and Z Ions are numbered in decreasing order: y5, y4, y3, y2, y1
                                    ' A and B, C ions are numbered in increasing order: a1, a2, etc.  or b1, b2, etc.
                                    strIonSymbolGeneric = LookupIonTypeString(eIonType)
                                    If eIonType = itIonTypeConstants.itYIon Or eIonType = itIonTypeConstants.itZIon Then
                                        strIonSymbol = strIonSymbolGeneric & Trim(Str(ResidueCount - lngResidueIndex + 1))
                                    Else
                                        strIonSymbol = strIonSymbolGeneric & Trim(Str(lngResidueIndex))
                                    End If
                                    
                                    If intChargeIndex > 1 Then
                                        strIonSymbol = strIonSymbol & String(intChargeIndex, "+")
                                        strIonSymbolGeneric = strIonSymbolGeneric & String(intChargeIndex, "+")
                                    End If
                                    
                                    AppendDataToFragSpectrum lngIonCount, FragSpectrumWork(), sngConvolutedMass, sngIntensity, strIonSymbol, strIonSymbolGeneric, lngResidueIndex, .Symbol, intChargeIndex, eIonType, False
                                    
                                    ' Add shoulder ions to PredictedSpectrum() if a B or Y ion and the shoulder intensity is > 0
                                    ' Need to use Abs() here since user can define negative theoretical intensities (which allows for plotting a spectrum inverted)
                                    If Abs(sngIonShoulderIntensity) > 0 And (eIonType = itIonTypeConstants.itBIon Or eIonType = itIonTypeConstants.itYIon Or eIonType = itIonTypeConstants.itCIon Or eIonType = itIonTypeConstants.itZIon) Then
                                        For intShoulderIndex = -1 To 1 Step 2
                                            sngObservedMass = sngConvolutedMass + intShoulderIndex * (1 / intChargeIndex)
                                            AppendDataToFragSpectrum lngIonCount, FragSpectrumWork(), sngObservedMass, sngIonShoulderIntensity, SHOULDER_ION_PREFIX & strIonSymbol, SHOULDER_ION_PREFIX & strIonSymbolGeneric, lngResidueIndex, .Symbol, intChargeIndex, eIonType, True
                                        Next intShoulderIndex
                                    End If
                                    
                                    ' Apply neutral loss modifications
                                    If mFragSpectrumOptions.IonTypeOptions(eIonType).NeutralLossWater Then
                                        ' Loss of water only affects Ser, Thr, Asp, or Glu (S, T, E, or D)
                                        ' See if the residues up to this point contain any of these residues
                                        If InStr(strResidues, "Ser") Or InStr(strResidues, "Thr") Or InStr(strResidues, "Glu") Or InStr(strResidues, "Asp") Then
                                            sngObservedMass = sngConvolutedMass - (dblHOHMass / intChargeIndex)
                                            AppendDataToFragSpectrum lngIonCount, FragSpectrumWork(), sngObservedMass, sngNeutralLossIntensity, strIonSymbol & mWaterLossSymbol, strIonSymbolGeneric & mWaterLossSymbol, lngResidueIndex, .Symbol, intChargeIndex, eIonType, False
                                        End If
                                    End If
                                    
                                    If mFragSpectrumOptions.IonTypeOptions(eIonType).NeutralLossAmmonia Then
                                        ' Loss of Ammonia only affects Arg, Lys, Gln, or Asn (R, K, Q, or N)
                                        ' See if the residues up to this point contain any of these residues
                                        If InStr(strResidues, "Arg") Or InStr(strResidues, "Lys") Or InStr(strResidues, "Gln") Or InStr(strResidues, "Asn") Then
                                            sngObservedMass = sngConvolutedMass - (dblNH3Mass / intChargeIndex)
                                            AppendDataToFragSpectrum lngIonCount, FragSpectrumWork(), sngObservedMass, sngNeutralLossIntensity, strIonSymbol & mAmmoniaLossSymbol, strIonSymbolGeneric & mAmmoniaLossSymbol, lngResidueIndex, .Symbol, intChargeIndex, eIonType, False
                                        End If
                                    End If
                                
                                    If mFragSpectrumOptions.IonTypeOptions(eIonType).NeutralLossPhosphate Then
                                        ' Loss of phosphate only affects phosphorylated residues
                                        ' Technically, only Ser, Thr, or Tyr (S, T, or Y) can be phosphorylated, but if the user marks other residues as phosphorylated, we'll allow that
                                        ' See if the residues up to this point contain phosphorylated residues
                                        If blnPhosphorylated Then
                                            sngObservedMass = sngConvolutedMass - (dblH3PO4Mass / intChargeIndex)
                                            AppendDataToFragSpectrum lngIonCount, FragSpectrumWork(), sngObservedMass, sngNeutralLossIntensity, strIonSymbol & mPhosphoLossSymbol, strIonSymbolGeneric & mPhosphoLossSymbol, lngResidueIndex, .Symbol, intChargeIndex, eIonType, False
                                        End If
                                    End If
                                
                                End If
                            End If
                        Next intChargeIndex
                    End If
                End If
            Next eIonType
        End With
    Next lngResidueIndex
    
    ' Sort arrays by mass (using a pointer array to synchronize the arrays)
    ReDim PointerArray(lngIonCount)
    
    For lngIndex = 0 To lngIonCount - 1
        PointerArray(lngIndex) = lngIndex
    Next lngIndex
    
    ShellSortFragSpectrum FragSpectrumWork(), PointerArray(), 0, lngIonCount - 1
    
    ' Copy the data from FragSpectrumWork() to udtFragSpectrum()
    
    ReDim udtFragSpectrum(lngIonCount)
    
    For lngIndex = 0 To lngIonCount - 1
        udtFragSpectrum(lngIndex) = FragSpectrumWork(PointerArray(lngIndex))
    Next lngIndex
    
    ' Return the actual number of ions computed
    GetFragmentationMasses = lngIonCount
    
End Function

Public Function GetFragmentationSpectrumRequiredDataPoints() As Long
    ' Determines the total number of data points that will be required for a theoretical fragmentation spectrum
    
    GetFragmentationSpectrumRequiredDataPoints = ResidueCount * ComputeMaxIonsPerResidue()
    
End Function

Public Function GetFragmentationSpectrumOptions() As udtFragmentationSpectrumOptionsType
    
On Error GoTo GetFragmentationSpectrumOptionsErrorHandler
    
    GetFragmentationSpectrumOptions = mFragSpectrumOptions

    Exit Function

GetFragmentationSpectrumOptionsErrorHandler:
    GeneralErrorHandler "MWPeptideClass.GetFragmentationSpectrumOptions", Err.Number

End Function

Public Function GetPeptideMass() As Double
    ' Returns the mass of the entire peptide
    
    ' Update the residue masses in order to update mTotalMass
    UpdateResidueMasses

    GetPeptideMass = mTotalMass
End Function

Private Function GetInternalResidues(lngCurrentResidueIndex As Long, eIonType As itIonTypeConstants, Optional ByRef blnPhosphorylated As Boolean) As String
    ' Determines the residues preceding or following the given residue (up to and including the current residue)
    ' If eIonType is a, b, or c ions, then returns residues from the N terminus
    ' If eIonType is y or z ions, then returns residues from the C terminus
    ' Also, set blnPhosphorylated to true if any of the residues is Ser, Thr, or Tyr and is phosphorylated
    '
    ' Note that the residue symbols are separated by a space to avoid accidental matching by the InStr() function

    Dim strInternalResidues As String
    Dim lngResidueIndex As Long
    
    strInternalResidues = ""
    blnPhosphorylated = False
    If eIonType = itIonTypeConstants.itYIon Or eIonType = itIonTypeConstants.itZIon Then
        For lngResidueIndex = lngCurrentResidueIndex To ResidueCount
            With Residues(lngResidueIndex)
                strInternalResidues = strInternalResidues & .Symbol & " "
                If .Phosphorylated Then blnPhosphorylated = True
            End With
        Next lngResidueIndex
    Else
        For lngResidueIndex = 1 To lngCurrentResidueIndex
            With Residues(lngResidueIndex)
                strInternalResidues = strInternalResidues & .Symbol & " "
                If .Phosphorylated Then blnPhosphorylated = True
            End With
        Next lngResidueIndex
    End If
    
    GetInternalResidues = strInternalResidues
    
End Function

Public Function GetModificationSymbol(ByVal lngModificationID As Long, ByRef strModSymbol As String, ByRef dblModificationMass As Double, ByRef blnIndicatesPhosphorylation As Boolean, ByRef strComment As String) As Long
    ' Returns information on the modification with lngModificationID
    ' Returns 0 if success, 1 if failure
    
    If lngModificationID >= 1 And lngModificationID <= ModificationSymbolCount Then
        With ModificationSymbols(lngModificationID)
            strModSymbol = .Symbol
            dblModificationMass = .ModificationMass
            blnIndicatesPhosphorylation = .IndicatesPhosphorylation
            strComment = .Comment
        End With
        GetModificationSymbol = 0
    Else
        strModSymbol = ""
        dblModificationMass = 0
        blnIndicatesPhosphorylation = False
        strComment = ""
        GetModificationSymbol = 1
    End If
    
End Function

Public Function GetModificationSymbolCount() As Long
    ' Returns the number of modifications defined
    
    GetModificationSymbolCount = ModificationSymbolCount
End Function

Public Function GetModificationSymbolID(strModSymbol As String) As Long
    ' Returns the ID for a given modification
    ' Returns 0 if not found, the ID if found
    
    Dim intIndex As Integer, lngModificationIDMatch As Long
    
    For intIndex = 1 To ModificationSymbolCount
        If ModificationSymbols(intIndex).Symbol = strModSymbol Then
            lngModificationIDMatch = intIndex
            Exit For
        End If
    Next intIndex
    
    GetModificationSymbolID = lngModificationIDMatch
    
End Function

Public Function GetResidue(ByVal lngResidueNumber As Long, ByRef strSymbol As String, ByRef dblMass As Double, ByRef blnIsModified As Boolean, ByRef intModificationCount As Integer) As Long
    ' Returns 0 if success, 1 if failure
    If lngResidueNumber >= 1 And lngResidueNumber <= ResidueCount Then
        With Residues(lngResidueNumber)
            strSymbol = .Symbol
            dblMass = .Mass
            blnIsModified = (.ModificationIDCount > 0)
            intModificationCount = .ModificationIDCount
        End With
        GetResidue = 0
    Else
        GetResidue = 1
    End If
End Function

Public Function GetResidueCount() As Long
    GetResidueCount = ResidueCount
End Function

Public Function GetResidueCountSpecificResidue(strResidueSymbol As String, Optional blnUse3LetterCode As Boolean = True) As Long
    ' Returns the number of occurrences of the given residue in the loaded sequence
    
    Dim strSearchResidue3Letter As String
    Dim lngResidueCount As Long
    Dim lngResidueIndex As Long
    
    If blnUse3LetterCode Then
        strSearchResidue3Letter = strResidueSymbol
    Else
        strSearchResidue3Letter = GetAminoAcidSymbolConversionInternal(strResidueSymbol, True)
    End If
    
    lngResidueCount = 0
    For lngResidueIndex = 0 To ResidueCount - 1
        If Residues(lngResidueIndex).Symbol = strSearchResidue3Letter Then
            lngResidueCount = lngResidueCount + 1
        End If
        
    Next lngResidueIndex
    
    GetResidueCountSpecificResidue = lngResidueCount
End Function

Public Function GetResidueModificationIDs(lngResidueNumber As Long, lngModificationIDsOneBased() As Long) As Long
    ' Returns the number of Modifications
    ' ReDims lngModificationIDsOneBased() to hold the values
    
    Dim intIndex As Integer
    
    If lngResidueNumber >= 1 And lngResidueNumber <= ResidueCount Then
        
        With Residues(lngResidueNumber)
        
            ' Need to use this in case the calling program is sending an array with fixed dimensions
            On Error Resume Next
            ReDim lngModificationIDsOneBased(.ModificationIDCount)
        
            For intIndex = 1 To .ModificationIDCount
                lngModificationIDsOneBased(intIndex) = .ModificationIDs(intIndex)
            Next intIndex
            
            GetResidueModificationIDs = .ModificationIDCount
        End With
    Else
        GetResidueModificationIDs = 0
    End If

End Function

Public Function GetResidueSymbolOnly(ByVal lngResidueNumber As Long, Optional blnUse3LetterCode As Boolean = True) As String
    ' Returns the symbol at the given residue number, or "" if an invalid residue number
    
    Dim strSymbol As String
    
    If lngResidueNumber >= 1 And lngResidueNumber <= ResidueCount Then
        With Residues(lngResidueNumber)
            strSymbol = .Symbol
        End With
        If Not blnUse3LetterCode Then strSymbol = GetAminoAcidSymbolConversionInternal(strSymbol, False)
    Else
        strSymbol = ""
    End If
    
    GetResidueSymbolOnly = strSymbol

End Function

Public Function GetSequence(Optional blnUse3LetterCode As Boolean = True, Optional blnAddSpaceEvery10Residues As Boolean = False, Optional blnSeparateResiduesWithDash As Boolean = False, Optional blnIncludeNandCTerminii As Boolean = False, Optional blnIncludeModificationSymbols As Boolean = True) As String
    ' Construct a text sequence using Residues() and the N and C Terminus info
    
    Dim strSequence As String, strSymbol3Letter As String, strSymbol1Letter As String
    Dim strDashAdd As String
    Dim strModSymbol As String, strModSymbolComment As String
    Dim blnIndicatesPhosphorylation As Boolean
    Dim dblModMass As Double
    Dim lngIndex As Long, intModIndex As Integer
    Dim lngError As Long
    
    If blnSeparateResiduesWithDash Then strDashAdd = "-" Else strDashAdd = ""
    
    strSequence = ""
    For lngIndex = 1 To ResidueCount
        With Residues(lngIndex)
            strSymbol3Letter = .Symbol
            If blnUse3LetterCode Then
                strSequence = strSequence & strSymbol3Letter
            Else
                strSymbol1Letter = GetAminoAcidSymbolConversionInternal(strSymbol3Letter, False)
                If strSymbol1Letter = "" Then strSymbol1Letter = UNKNOWN_SYMBOL_ONE_LETTER
                strSequence = strSequence & strSymbol1Letter
            End If
            
            If blnIncludeModificationSymbols Then
                For intModIndex = 1 To .ModificationIDCount
                    lngError = GetModificationSymbol(.ModificationIDs(intModIndex), strModSymbol, dblModMass, blnIndicatesPhosphorylation, strModSymbolComment)
                    If lngError = 0 Then
                        strSequence = strSequence & strModSymbol
                    Else
                        Debug.Assert False
                    End If
                Next intModIndex
            End If
            
        End With
        
        If lngIndex <> ResidueCount Then
            If blnAddSpaceEvery10Residues Then
                If lngIndex Mod 10 = 0 Then
                    strSequence = strSequence & " "
                Else
                    strSequence = strSequence & strDashAdd
                End If
            Else
                strSequence = strSequence & strDashAdd
            End If
        End If
        
    Next lngIndex
    
    If blnIncludeNandCTerminii Then
        strSequence = mNTerminus.Formula & strDashAdd & strSequence & strDashAdd & mCTerminus.Formula
    End If
    
    GetSequence = strSequence
End Function

Public Function GetSymbolWaterLoss() As String
    GetSymbolWaterLoss = mWaterLossSymbol
End Function

Public Function GetSymbolPhosphoLoss() As String
    GetSymbolPhosphoLoss = mPhosphoLossSymbol
End Function

Public Function GetSymbolAmmoniaLoss() As String
    GetSymbolAmmoniaLoss = mAmmoniaLossSymbol
End Function

Public Function GetTrypticName(ByVal strProteinResidues As String, ByVal strPeptideResidues As String, Optional ByRef lngReturnResidueStart As Long = 0, Optional ByRef lngReturnResidueEnd As Long = 0, Optional ByVal blnICR2LSCompatible As Boolean = False, Optional ByVal strRuleResidues As String = TRYPTIC_RULE_RESIDUES, Optional ByVal strExceptionResidues As String = TRYPTIC_EXCEPTION_RESIDUES, Optional ByVal strTerminiiSymbol As String = TERMINII_SYMBOL, Optional ByVal blnIgnoreCase As Boolean = True, Optional ByVal lngProteinSearchStartLoc As Long = 1) As String
    ' Examines strPeptideResidues to see where they exist in strProteinResidues
    ' Constructs a name string based on their position and based on whether the fragment is truly tryptic
    ' In addition, returns the position of the first and last residue in lngReturnResidueStart and lngReturnResidueEnd
    ' The tryptic name in the following format
    ' t1  indicates tryptic peptide 1
    ' t2 represents tryptic peptide 2, etc.
    ' t1.2  indicates tryptic peptide 1, plus one more tryptic peptide, i.e. t1 and t2
    ' t5.2  indicates tryptic peptide 5, plus one more tryptic peptide, i.e. t5 and t6
    ' t5.3  indicates tryptic peptide 5, plus two more tryptic peptides, i.e. t5, t6, and t7
    ' 40.52  means that the residues are not tryptic, and simply range from residue 40 to 52
    ' If the peptide residues are not present in strProteinResidues, then returns ""
    ' Since a peptide can occur multiple times in a protein, one can set lngProteinSearchStartLoc to a value larger than 1 to ignore previous hits
    
    ' If blnICR2LSCompatible is True, then the values returned when a peptide is not tryptic are modified to
    ' range from the starting residue, to the ending residue +1
    ' lngReturnResidueEnd is always equal to the position of the final residue, regardless of blnICR2LSCompatible
    
    ' For example, if strProteinResidues = "IGKANR"
    ' Then when strPeptideResidues = "IGK", the TrypticName is t1
    ' Then when strPeptideResidues = "ANR", the TrypticName is t2
    ' Then when strPeptideResidues = "IGKANR", the TrypticName is t1.2
    ' Then when strPeptideResidues = "IG", the TrypticName is 1.2
    ' Then when strPeptideResidues = "KANR", the TrypticName is 3.6
    ' Then when strPeptideResidues = "NR", the TrypticName is 5.6
    
    ' However, if blnICR2LSCompatible = True, then the last three are changed to:
    ' Then when strPeptideResidues = "IG", the TrypticName is 1.3
    ' Then when strPeptideResidues = "KANR", the TrypticName is 3.7
    ' Then when strPeptideResidues = "NR", the TrypticName is 5.7
    
    Dim intStartLoc As Integer, intEndLoc As Integer
    Dim strTrypticName As String
    Dim strPrefix As String, strSuffix As String
    Dim strResidueFollowingSearchResidues As String
    Dim blnMatchesCleavageRule As Boolean

    Dim intTrypticResidueNumber As Integer
    Dim intRuleResidueMatchCount As Integer
    Dim lngRuleResidueLoc As Long
    Dim strProteinResiduesBeforeStartLoc As String
    Dim lngPeptideResiduesLength As Long
    
    If blnIgnoreCase Then
        strProteinResidues = UCase(strProteinResidues)
        strPeptideResidues = UCase(strPeptideResidues)
    End If
    
    If lngProteinSearchStartLoc <= 1 Then
        intStartLoc = InStr(strProteinResidues, strPeptideResidues)
    Else
        intStartLoc = InStr(Mid(strProteinResidues, lngProteinSearchStartLoc), strPeptideResidues)
        If intStartLoc > 0 Then
            intStartLoc = intStartLoc + lngProteinSearchStartLoc - 1
        End If
    End If
    
    lngPeptideResiduesLength = Len(strPeptideResidues)
    
    If intStartLoc > 0 And Len(strProteinResidues) > 0 And lngPeptideResiduesLength > 0 Then
        intEndLoc = intStartLoc + lngPeptideResiduesLength - 1
        
        ' Determine if the residue is tryptic
        ' Use CheckSequenceAgainstCleavageRule() for this
        If intStartLoc > 1 Then
            strPrefix = Mid(strProteinResidues, intStartLoc - 1, 1)
        Else
            strPrefix = strTerminiiSymbol
        End If
    
        If intEndLoc = Len(strProteinResidues) Then
            strSuffix = strTerminiiSymbol
        Else
            strSuffix = Mid(strProteinResidues, intEndLoc + 1, 1)
        End If
        
        blnMatchesCleavageRule = CheckSequenceAgainstCleavageRule(strPrefix & "." & strPeptideResidues & "." & strSuffix, strRuleResidues, strExceptionResidues, False, ".", strTerminiiSymbol, blnIgnoreCase)
        
        If blnMatchesCleavageRule Then
            ' Construct strTrypticName
            
            ' Determine which tryptic residue strPeptideResidues is
            If intStartLoc = 1 Then
                intTrypticResidueNumber = 1
            Else
                intTrypticResidueNumber = 0
                strProteinResiduesBeforeStartLoc = Left(strProteinResidues, intStartLoc - 1)
                strResidueFollowingSearchResidues = Left(strPeptideResidues, 1)
                intTrypticResidueNumber = 0
                lngRuleResidueLoc = 0
                Do
                    lngRuleResidueLoc = GetTrypticNameFindNextCleavageLoc(strProteinResiduesBeforeStartLoc, strResidueFollowingSearchResidues, lngRuleResidueLoc + 1, strRuleResidues, strExceptionResidues, strTerminiiSymbol)
                    If lngRuleResidueLoc > 0 Then
                        intTrypticResidueNumber = intTrypticResidueNumber + 1
                    End If
                Loop While lngRuleResidueLoc > 0 And lngRuleResidueLoc + 1 < intStartLoc
                intTrypticResidueNumber = intTrypticResidueNumber + 1
            End If
            
            ' Determine number of K or R residues in strPeptideResidues
            ' Ignore K or R residues followed by Proline
            intRuleResidueMatchCount = 0
            lngRuleResidueLoc = 0
            Do
                lngRuleResidueLoc = GetTrypticNameFindNextCleavageLoc(strPeptideResidues, strSuffix, lngRuleResidueLoc + 1, strRuleResidues, strExceptionResidues, strTerminiiSymbol)
                If lngRuleResidueLoc > 0 Then
                    intRuleResidueMatchCount = intRuleResidueMatchCount + 1
                End If
            Loop While lngRuleResidueLoc > 0 And lngRuleResidueLoc < lngPeptideResiduesLength
            
            strTrypticName = "t" & Trim(Str(intTrypticResidueNumber))
            If intRuleResidueMatchCount > 1 Then
                strTrypticName = strTrypticName & "." & Trim(Str((intRuleResidueMatchCount)))
            End If
        Else
            If blnICR2LSCompatible Then
                strTrypticName = Trim(Str(intStartLoc)) & "." & Trim(Str(intEndLoc + 1))
            Else
                strTrypticName = Trim(Str(intStartLoc)) & "." & Trim(Str(intEndLoc))
            End If
        End If
        
        lngReturnResidueStart = intStartLoc
        lngReturnResidueEnd = intEndLoc
        GetTrypticName = strTrypticName
    Else
        ' Residues not found
        lngReturnResidueStart = 0
        lngReturnResidueEnd = 0
        GetTrypticName = ""
    End If
    
End Function

Public Function GetTrypticNameMultipleMatches(ByVal strProteinResidues As String, ByVal strPeptideResidues As String, Optional ByRef lngReturnMatchCount As Long, Optional ByRef lngReturnResidueStart As Long = 0, Optional ByRef lngReturnResidueEnd As Long = 0, Optional ByVal blnICR2LSCompatible As Boolean, Optional ByVal strRuleResidues As String = TRYPTIC_RULE_RESIDUES, Optional ByVal strExceptionResidues As String = TRYPTIC_EXCEPTION_RESIDUES, Optional ByVal strTerminiiSymbol As String = TERMINII_SYMBOL, Optional ByVal blnIgnoreCase As Boolean = True, Optional ByVal lngProteinSearchStartLoc As Long = 1, Optional ByVal strListDelimeter As String = ", ") As String
    ' Examines strPeptideResidues to see where they exist in strProteinResidues
    ' Looks for all possible matches, returning them as a comma separated list
    ' Returns the number of matches in lngReturnMatchCount
    ' lngReturnResidueStart contains the residue number of the start of the first match
    ' lngReturnResidueEnd contains the residue number of the end of the last match
    
    ' See GetTrypticName for additional information
    
    Dim strNameList As String, strCurrentName As String
    Dim lngCurrentSearchLoc As Long
    Dim lngCurrentResidueStart As Long, lngCurrentResidueEnd As Long
    
    lngCurrentSearchLoc = lngProteinSearchStartLoc
    lngReturnMatchCount = 0
    Do
        strCurrentName = GetTrypticName(strProteinResidues, strPeptideResidues, lngCurrentResidueStart, lngCurrentResidueEnd, blnICR2LSCompatible, strRuleResidues, strExceptionResidues, strTerminiiSymbol, blnIgnoreCase, lngCurrentSearchLoc)
        
        If Len(strCurrentName) > 0 Then
            If Len(strNameList) > 0 Then
                strNameList = strNameList & strListDelimeter
            End If
            strNameList = strNameList & strCurrentName
            lngCurrentSearchLoc = lngCurrentResidueEnd + 1
            lngReturnMatchCount = lngReturnMatchCount + 1
            
            If lngReturnMatchCount = 1 Then
                lngReturnResidueStart = lngCurrentResidueStart
            End If
            lngReturnResidueEnd = lngCurrentResidueEnd
            
            If lngCurrentSearchLoc > Len(strProteinResidues) Then Exit Do
        Else
            Exit Do
        End If
    Loop
    
    GetTrypticNameMultipleMatches = strNameList
    
End Function

Private Function GetTrypticNameFindNextCleavageLoc(strSearchResidues As String, strResidueFollowingSearchResidues As String, ByVal lngStartChar As Long, Optional ByVal strSearchChars As String = TRYPTIC_RULE_RESIDUES, Optional ByVal strExceptionSuffixResidues As String = TRYPTIC_EXCEPTION_RESIDUES, Optional ByVal strTerminiiSymbol As String = TERMINII_SYMBOL) As Long
    ' Finds the location of the next strSearchChar in strSearchResidues (K or R by default)
    ' Assumes strSearchResidues are already upper case
    ' Examines the residue following the matched residue
    '   If it matches one of the characters in strExceptionSuffixResidues, then the match is not counted
    ' Note that strResidueFollowingSearchResidues is necessary in case the potential cleavage residue is the final residue in strSearchResidues
    ' We need to know the next residue to determine if it matches an exception residue
    ' For example, if strSearchResidues =      "IGASGEHIFIIGVDKPNR"
    '  and the protein it is part of is: TNSANFRIGASGEHIFIIGVDKPNRQPDS
    '  and strSearchChars = "KR while strExceptionSuffixResidues  = "P"
    ' Then the K in IGASGEHIFIIGVDKPNR is ignored because the following residue is P,
    '  while the R in IGASGEHIFIIGVDKPNR is OK because strResidueFollowingSearchResidues is Q
    ' It is the calling function's responsibility to assign the correct residue to strResidueFollowingSearchResidues
    ' If no match is found, but strResidueFollowingSearchResidues is "-", then the cleavage location returned is Len(strSearchResidues) + 1
    
    Dim intCharLocInSearchChars As Integer
    Dim lngCharLoc As Long, lngMinCharLoc As Long
    Dim intExceptionSuffixResidueCount As Integer
    Dim intCharLocInExceptionChars As Integer
    Dim strResidueFollowingCleavageResidue As String
    Dim lngExceptionCharLocInSearchResidues As Long, lngCharLocViaRecursiveSearch As Long
    
    intExceptionSuffixResidueCount = Len(strExceptionSuffixResidues)
    
    lngMinCharLoc = -1
    For intCharLocInSearchChars = 1 To Len(strSearchChars)
        lngCharLoc = InStr(Mid(strSearchResidues, lngStartChar), Mid(strSearchChars, intCharLocInSearchChars, 1))
        
        If lngCharLoc > 0 Then
            lngCharLoc = lngCharLoc + lngStartChar - 1
            
            If intExceptionSuffixResidueCount > 0 Then
                ' Make sure strSuffixResidue does not match strExceptionSuffixResidues
                If lngCharLoc < Len(strSearchResidues) Then
                    lngExceptionCharLocInSearchResidues = lngCharLoc + 1
                    strResidueFollowingCleavageResidue = Mid(strSearchResidues, lngExceptionCharLocInSearchResidues, 1)
                Else
                    ' Matched the last residue in strSearchResidues
                    lngExceptionCharLocInSearchResidues = Len(strSearchResidues) + 1
                    strResidueFollowingCleavageResidue = strResidueFollowingSearchResidues
                End If
                
                For intCharLocInExceptionChars = 1 To intExceptionSuffixResidueCount
                    If strResidueFollowingCleavageResidue = Mid(strExceptionSuffixResidues, intCharLocInExceptionChars, 1) Then
                        ' Exception char is the following character; can't count this as the cleavage point
                        
                        If lngExceptionCharLocInSearchResidues < Len(strSearchResidues) Then
                            ' Recursively call this function to find the next cleavage position, using an updated lngStartChar position
                            lngCharLocViaRecursiveSearch = GetTrypticNameFindNextCleavageLoc(strSearchResidues, strResidueFollowingSearchResidues, lngExceptionCharLocInSearchResidues, strSearchChars, strExceptionSuffixResidues, strTerminiiSymbol)
                            
                            If lngCharLocViaRecursiveSearch > 0 Then
                                ' Found a residue further along that is a valid cleavage point
                                lngCharLoc = lngCharLocViaRecursiveSearch
                            Else
                                lngCharLoc = 0
                            End If
                        Else
                            lngCharLoc = 0
                        End If
                        Exit For
                    End If
                Next intCharLocInExceptionChars
            End If
        End If
        
        If lngCharLoc > 0 Then
            If lngMinCharLoc < 0 Then
                lngMinCharLoc = lngCharLoc
            Else
                If lngCharLoc < lngMinCharLoc Then
                    lngMinCharLoc = lngCharLoc
                End If
            End If
        End If
    Next intCharLocInSearchChars

    If lngMinCharLoc < 0 And strResidueFollowingSearchResidues = strTerminiiSymbol Then
        lngMinCharLoc = Len(strSearchResidues) + 1
    End If
    
    If lngMinCharLoc < 0 Then
        GetTrypticNameFindNextCleavageLoc = 0
    Else
        GetTrypticNameFindNextCleavageLoc = lngMinCharLoc
    End If
    
End Function

Public Function GetTrypticPeptideNext(ByVal strProteinResidues As String, ByVal lngSearchStartLoc As Long, Optional ByRef lngReturnResidueStart As Long, Optional ByRef lngReturnResidueEnd As Long, Optional ByVal strRuleResidues As String = TRYPTIC_RULE_RESIDUES, Optional ByVal strExceptionResidues As String = TRYPTIC_EXCEPTION_RESIDUES, Optional ByVal strTerminiiSymbol As String = TERMINII_SYMBOL) As String
    ' Returns the next tryptic peptide in strProteinResidues, starting the search as lngSearchStartLoc
    ' Useful when obtaining all of the tryptic peptides for a protein, since this function will operate
    '  much faster than repeatedly calling GetTrypticPeptideByFragmentNumber()
    
    ' Returns the position of the start and end residues using lngReturnResidueStart and lngReturnResidueEnd
    
    Dim lngRuleResidueLoc As Long
    Dim lngProteinResiduesLength As Long
    
    If lngSearchStartLoc < 1 Then lngSearchStartLoc = 1
    
    lngProteinResiduesLength = Len(strProteinResidues)
    If lngSearchStartLoc > lngProteinResiduesLength Then
        GetTrypticPeptideNext = ""
        Exit Function
    End If
    
    lngRuleResidueLoc = GetTrypticNameFindNextCleavageLoc(strProteinResidues, strTerminiiSymbol, lngSearchStartLoc, strRuleResidues, strExceptionResidues, strTerminiiSymbol)
    If lngRuleResidueLoc > 0 Then
        lngReturnResidueStart = lngSearchStartLoc
        If lngRuleResidueLoc > lngProteinResiduesLength Then
            lngReturnResidueEnd = lngProteinResiduesLength
        Else
            lngReturnResidueEnd = lngRuleResidueLoc
        End If
        GetTrypticPeptideNext = Mid(strProteinResidues, lngReturnResidueStart, lngReturnResidueEnd - lngReturnResidueStart + 1)
    Else
        lngReturnResidueStart = 1
        lngReturnResidueEnd = lngProteinResiduesLength
        GetTrypticPeptideNext = strProteinResidues
    End If
    
End Function

Public Function GetTrypticPeptideByFragmentNumber(ByVal strProteinResidues As String, ByVal intDesiredPeptideNumber As Integer, Optional ByRef lngReturnResidueStart As Long, Optional ByRef lngReturnResidueEnd As Long, Optional ByVal strRuleResidues As String = TRYPTIC_RULE_RESIDUES, Optional ByVal strExceptionResidues As String = TRYPTIC_EXCEPTION_RESIDUES, Optional ByVal strTerminiiSymbol As String = TERMINII_SYMBOL, Optional ByVal blnIgnoreCase As Boolean = True) As String
    ' Returns the desired tryptic peptide from strProteinResidues
    ' For example, if strProteinResidues = "IGKANRMTFGL" then
    '  when intDesiredPeptideNumber = 1, returns "IGK"
    '  when intDesiredPeptideNumber = 2, returns "ANR"
    '  when intDesiredPeptideNumber = 3, returns "MTFGL"
    
    ' Optionally, returns the position of the start and end residues
    '  using lngReturnResidueStart and lngReturnResidueEnd
    
    
    Dim lngStartLoc As Long, lngRuleResidueLoc As Long
    Dim lngPrevStartLoc As Long
    Dim lngProteinResiduesLength As Long
    Dim intCurrentTrypticPeptideNumber As Integer
    
    Dim strMatchingFragment As String
    
    If intDesiredPeptideNumber < 1 Then
        GetTrypticPeptideByFragmentNumber = ""
        Exit Function
    End If
    
    If blnIgnoreCase Then
        strProteinResidues = UCase(strProteinResidues)
    End If
    lngProteinResiduesLength = Len(strProteinResidues)
    
    lngStartLoc = 1
    lngRuleResidueLoc = 0
    intCurrentTrypticPeptideNumber = 0
    Do
        lngRuleResidueLoc = GetTrypticNameFindNextCleavageLoc(strProteinResidues, strTerminiiSymbol, lngStartLoc, strRuleResidues, strExceptionResidues, strTerminiiSymbol)
        If lngRuleResidueLoc > 0 Then
            intCurrentTrypticPeptideNumber = intCurrentTrypticPeptideNumber + 1
            lngPrevStartLoc = lngStartLoc
            lngStartLoc = lngRuleResidueLoc + 1
            
            If lngPrevStartLoc > lngProteinResiduesLength Then
                ' User requested a peptide number that doesn't exist
                GetTrypticPeptideByFragmentNumber = ""
                Exit Function
            End If
        Else
            ' I don't think I'll ever reach this code
            Debug.Assert False
            Exit Do
        End If
    Loop While intCurrentTrypticPeptideNumber < intDesiredPeptideNumber

    strMatchingFragment = ""
    If intCurrentTrypticPeptideNumber > 0 And lngPrevStartLoc > 0 Then
        If lngPrevStartLoc > Len(strProteinResidues) Then
            ' User requested a peptide number that is too high
            lngReturnResidueStart = 0
            lngReturnResidueEnd = 0
            strMatchingFragment = ""
        Else
            ' Match found, find the extent of this peptide
            lngReturnResidueStart = lngPrevStartLoc
            If lngRuleResidueLoc > lngProteinResiduesLength Then
                lngReturnResidueEnd = lngProteinResiduesLength
            Else
                lngReturnResidueEnd = lngRuleResidueLoc
            End If
            strMatchingFragment = Mid(strProteinResidues, lngPrevStartLoc, lngRuleResidueLoc - lngPrevStartLoc + 1)
        End If
    Else
        lngReturnResidueStart = 1
        lngReturnResidueEnd = lngProteinResiduesLength
        strMatchingFragment = strProteinResidues
    End If
    
    GetTrypticPeptideByFragmentNumber = strMatchingFragment
    
End Function

Public Function CheckSequenceAgainstCleavageRule(ByVal strSequence As String, ByVal strRuleResidues As String, ByVal strExceptionSuffixResidues As String, ByVal blnAllowPartialCleavage As Boolean, Optional ByVal strSeparationChar As String = ".", Optional ByVal strTerminiiSymbol As String = TERMINII_SYMBOL, Optional ByVal blnIgnoreCase As Boolean = True, Optional ByRef intRuleMatchCount As Integer) As Boolean
    ' Checks strSequence to see if it matches the cleavage rule
    ' Returns True if valid, False if invalid
    ' Returns True if doesn't contain any periods, and thus, can't be examined
    ' The ByRef variable intRuleMatchCount can be used to retrieve the number of ends that matched the rule (0, 1, or 2); terminii are counted as rule matches
    
    ' The residues in strRuleResidues specify the cleavage rule
    ' The peptide must end in one of the residues, or in -
    ' The preceding residue must be one of the residues or be -
    ' EXCEPTION: if blnAllowPartialCleavage = True then the rules need only apply to one end
    ' Finally, the suffix residue cannot match any of the residues in strExceptionSuffixResidues
    
    ' For example, if strRuleResidues = "KR" and strExceptionSuffixResidues = "P"
    ' Then if strSequence = "R.AEQDDLANYGPGNGVLPSAGSSISMEK.L" then blnMatchesCleavageRule = True
    ' However, if strSequence = "R.IGASGEHIFIIGVDK.P" then blnMatchesCleavageRule = False since strSuffix = "P"
    ' Finally, if strSequence = "R.IGASGEHIFIIGVDKPNR.Q" then blnMatchesCleavageRule = True since K is ignored, but the final R.Q is valid
    
    Dim strSequenceStart As String, strSequenceEnd As String
    Dim strPrefix As String, strSuffix As String
    Dim blnMatchesCleavageRule As Boolean, blnSkipThisEnd As Boolean
    Dim strTestResidue As String
    Dim intEndToCheck As Integer
    
    ' Need to reset this to zero since passed ByRef
    intRuleMatchCount = 0
    
    ' First, make sure the sequence is in the form A.BCDEFG.H or A.BCDEFG or BCDEFG.H
    ' If it isn't, then we can't check it (we'll return true)
    
    If Len(strRuleResidues) = 0 Then
        ' No rules
        CheckSequenceAgainstCleavageRule = True
        Exit Function
    End If
    
    If InStr(strSequence, strSeparationChar) = 0 Then
        ' No periods, can't check
        Debug.Assert False
        CheckSequenceAgainstCleavageRule = True
        Exit Function
    End If
    
    If blnIgnoreCase Then
        strSequence = UCase(strSequence)
    End If
    
    ' Find the prefix residue and starting residue
    If Mid(strSequence, 2, 1) = strSeparationChar Then
        strPrefix = Left(strSequence, 1)
        strSequenceStart = Mid(strSequence, 3, 1)
    Else
        strSequenceStart = Left(strSequence, 1)
    End If
    
    ' Find the suffix residue and the ending residue
    If Mid(strSequence, Len(strSequence) - 1, 1) = strSeparationChar Then
        strSuffix = Right(strSequence, 1)
        strSequenceEnd = Mid(strSequence, Len(strSequence) - 2, 1)
    Else
        strSequenceEnd = Right(strSequence, 1)
    End If
    
    If strRuleResidues = strTerminiiSymbol Then
        ' Peptide database rules
        ' See if prefix and suffix are "" or are strTerminiiSymbol
        If (strPrefix = strTerminiiSymbol And strSuffix = strTerminiiSymbol) Or _
           (strPrefix = "" And strSuffix = "") Then
            intRuleMatchCount = 2
            blnMatchesCleavageRule = True
        Else
            blnMatchesCleavageRule = False
        End If
    Else
        If blnIgnoreCase Then
            strRuleResidues = UCase(strRuleResidues)
        End If
        
        ' Test each character in strRuleResidues against both strPrefix and strSequenceEnd
        ' Make sure strSuffix does not match strExceptionSuffixResidues
        For intEndToCheck = 0 To 1
            blnSkipThisEnd = False
            If intEndToCheck = 0 Then
                strTestResidue = strPrefix
                If strPrefix = strTerminiiSymbol Then
                    intRuleMatchCount = intRuleMatchCount + 1
                    blnSkipThisEnd = True
                Else
                    ' See if strSequenceStart matches one of the exception residues
                    ' If it does, make sure strPrefix does not match one of the rule residues
                    If CheckSequenceAgainstCleavageRuleMatchTestResidue(strSequenceStart, strExceptionSuffixResidues) Then
                        ' Match found
                        ' Make sure strPrefix does not match one of the rule residues
                        If CheckSequenceAgainstCleavageRuleMatchTestResidue(strPrefix, strRuleResidues) Then
                            ' Match found; thus does not match cleavage rule
                            blnSkipThisEnd = True
                        End If
                    End If
                End If
            Else
                strTestResidue = strSequenceEnd
                If strSuffix = strTerminiiSymbol Then
                    intRuleMatchCount = intRuleMatchCount + 1
                    blnSkipThisEnd = True
                Else
                    ' Make sure strSuffix does not match strExceptionSuffixResidues
                    If CheckSequenceAgainstCleavageRuleMatchTestResidue(strSuffix, strExceptionSuffixResidues) Then
                        ' Match found; thus does not match cleavage rule
                        blnSkipThisEnd = True
                    End If
                End If
            End If
            
            If Not blnSkipThisEnd Then
                If CheckSequenceAgainstCleavageRuleMatchTestResidue(strTestResidue, strRuleResidues) Then
                    intRuleMatchCount = intRuleMatchCount + 1
                End If
            End If
        Next intEndToCheck
    
        If intRuleMatchCount = 2 Then
            blnMatchesCleavageRule = True
        ElseIf intRuleMatchCount >= 1 And blnAllowPartialCleavage Then
            blnMatchesCleavageRule = True
        End If
    End If

    CheckSequenceAgainstCleavageRule = blnMatchesCleavageRule

End Function

Private Function CheckSequenceAgainstCleavageRuleMatchTestResidue(strTestResidue As String, strRuleResidues As String) As Boolean
    ' Checks to see if strTestResidue matches one of the residues in strRuleResidues
    ' Used to test by Rule Residues and Exception Residues
    
    Dim intCharLocInRuleResidues As Integer
    Dim strCompareResidue As String
    Dim blnMatchFound As Boolean
    
    For intCharLocInRuleResidues = 1 To Len(strRuleResidues)
        strCompareResidue = Trim(Mid(strRuleResidues, intCharLocInRuleResidues, 1))
        If Len(strCompareResidue) > 0 Then
            If strTestResidue = strCompareResidue Then
                blnMatchFound = True
                Exit For
            End If
        End If
    Next intCharLocInRuleResidues
    
    CheckSequenceAgainstCleavageRuleMatchTestResidue = blnMatchFound
    
End Function

Public Function ComputeImmoniumMass(dblResidueMass As Double) As Double
    ComputeImmoniumMass = dblResidueMass - dblImmoniumMassDifference
End Function

Public Function LookupIonTypeString(eIonType As itIonTypeConstants) As String

    Select Case eIonType
    Case itIonTypeConstants.itAIon: LookupIonTypeString = "a"
    Case itIonTypeConstants.itBIon: LookupIonTypeString = "b"
    Case itIonTypeConstants.itYIon: LookupIonTypeString = "y"
    Case itIonTypeConstants.itCIon: LookupIonTypeString = "c"
    Case itIonTypeConstants.itZIon: LookupIonTypeString = "z"
    Case Else: LookupIonTypeString = ""
    End Select
    
End Function

Public Function RemoveAllResidues() As Long
    ' Removes all the residues
    ' Returns 0 on success, 1 on failure
    
    ReserveMemoryForResidues 50, False
    ResidueCount = 0
    mTotalMass = 0
    
    RemoveAllResidues = 0
End Function

Public Function RemoveAllModificationSymbols() As Long
    ' Removes all possible Modification Symbols
    ' Returns 0 on success, 1 on failure
    ' Removing all modifications will invalidate any modifications present in a sequence
    
    ReserveMemoryForModifications 10, False
    ModificationSymbolCount = 0
    
    RemoveAllModificationSymbols = 0
End Function

Private Function RemoveLeadingH(ByRef strWorkingSequence As String) As Boolean
    ' Returns True if a leading H is removed
    Dim lngAbbrevID As Long
    Dim blnHRemoved As Boolean
    
    blnHRemoved = False
    If UCase(Left(strWorkingSequence, 1)) = "H" And Len(strWorkingSequence) >= 4 Then
        ' If next character is not a character, then remove the H
        If Not IsCharacter(Mid(strWorkingSequence, 2, 1)) Then
            strWorkingSequence = Mid(strWorkingSequence, 3)
            blnHRemoved = True
        Else
            ' Otherwise, see if next three characters are letters
            If IsCharacter(Mid(strWorkingSequence, 2, 1)) And _
               IsCharacter(Mid(strWorkingSequence, 3, 1)) And _
               IsCharacter(Mid(strWorkingSequence, 4, 1)) Then
                ' Formula starts with 4 characters and the first is H, see if the first 3 characters are a valid amino acid code
                lngAbbrevID = GetAbbreviationIDInternal(Left(strWorkingSequence, 3), True)
                                
                If lngAbbrevID <= 0 Then
                    ' Doesn't start with a valid amino acid 3 letter abbreviation, so remove the initial H
                    strWorkingSequence = Mid(strWorkingSequence, 2)
                    blnHRemoved = True
                End If
            End If
        End If
    End If

    RemoveLeadingH = blnHRemoved
End Function

Private Function RemoveTrailingOH(ByRef strWorkingSequence As String) As Boolean
    ' Returns True if a trailing OH is removed
    Dim lngAbbrevID As Long
    Dim blnOHRemoved As Boolean
    Dim lngStringLength As Long
    
    blnOHRemoved = False
    lngStringLength = Len(strWorkingSequence)
    If UCase(Right(strWorkingSequence, 2)) = "OH" And lngStringLength >= 5 Then
        ' If previous character is not a character, then remove the OH
        If Not IsCharacter(Mid(strWorkingSequence, lngStringLength - 2, 1)) Then
            strWorkingSequence = Left(strWorkingSequence, lngStringLength - 3)
            blnOHRemoved = True
        Else
            ' Otherwise, see if previous three characters are letters
            If IsCharacter(Mid(strWorkingSequence, lngStringLength - 2, 1)) Then
                ' Formula ends with 3 characters and the last two are OH, see if the last 3 characters are a valid amino acid code
                lngAbbrevID = GetAbbreviationIDInternal(Mid(strWorkingSequence, lngStringLength - 2, 3), True)
                                
                If lngAbbrevID <= 0 Then
                    ' Doesn't end with a valid amino acid 3 letter abbreviation, so remove the trailing OH
                    strWorkingSequence = Left(strWorkingSequence, lngStringLength - 2)
                    blnOHRemoved = True
                End If
            End If
        End If
    End If

    RemoveTrailingOH = blnOHRemoved

End Function

Public Function RemoveModification(strModSymbol As String) As Long
    ' Returns 0 if found and removed; 1 if error
    
    Dim lngIndex As Long
    Dim blnRemoved As Boolean
    
    For lngIndex = 1 To ModificationSymbolCount
        If ModificationSymbols(lngIndex).Symbol = strModSymbol Then
            RemoveModificationByID lngIndex
            blnRemoved = True
        End If
    Next lngIndex

    If blnRemoved Then
        RemoveModification = 0
    Else
        RemoveModification = 1
    End If
End Function

Public Function RemoveModificationByID(ByVal lngModificationID As Long) As Long
    ' Returns 0 if found and removed; 1 if error
    
    Dim lngIndex As Long
    Dim blnRemoved As Boolean
    
    If lngModificationID >= 1 And lngModificationID <= ModificationSymbolCount Then
        For lngIndex = lngModificationID To ModificationSymbolCount - 1
            ModificationSymbols(lngIndex) = ModificationSymbols(lngIndex + 1)
        Next lngIndex
        ModificationSymbolCount = ModificationSymbolCount - 1
        blnRemoved = True
    Else
        blnRemoved = False
    End If
    
    If blnRemoved Then
        RemoveModificationByID = 0
    Else
        RemoveModificationByID = 1
    End If
    
End Function

Public Function RemoveResidue(lngResidueNumber As Long) As Long
    ' Returns 0 if found and removed; 1 if error
    
    Dim lngIndex As Long
    
    If lngResidueNumber >= 1 And lngResidueNumber <= ResidueCount Then
        For lngIndex = lngResidueNumber To ResidueCount - 1
            Residues(lngIndex) = Residues(lngIndex + 1)
        Next lngIndex
        ResidueCount = ResidueCount - 1
        RemoveResidue = 0
    Else
        RemoveResidue = 1
    End If

End Function

Private Sub ReserveMemoryForResidues(lngNewResidueCount As Long, blnPreserveContents As Boolean)
    ' Only reserves the memory if necessary
    ' Thus, do not use this sub to clear Residues()
    
    If lngNewResidueCount > ResidueCountDimmed Then
        ResidueCountDimmed = lngNewResidueCount + RESIDUE_DIM_CHUNK
        If blnPreserveContents Then
            ReDim Preserve Residues(ResidueCountDimmed)
        Else
            ReDim Residues(ResidueCountDimmed)
        End If
    End If
End Sub

Private Sub ReserveMemoryForModifications(lngNewModificationCount As Long, blnPreserveContents As Boolean)
    
    If lngNewModificationCount > ModificationSymbolCountDimmed Then
        ModificationSymbolCountDimmed = lngNewModificationCount + 10
        If blnPreserveContents Then
            ReDim Preserve ModificationSymbols(ModificationSymbolCountDimmed)
        Else
            ReDim ModificationSymbols(ModificationSymbolCountDimmed)
        End If
    End If
End Sub

Public Function SetCTerminus(strFormula As String, Optional strFollowingResidue As String = "", Optional blnUse3LetterCode As Boolean = True) As Long
    ' Returns 0 if success; 1 if error
    
    ' Typical N terminus mods
        ' Free Acid = OH
        ' Amide = NH2
    
    With mCTerminus
        .Formula = strFormula
        .Mass = ComputeFormulaWeight(.Formula)
        If .Mass < 0 Then
            .Mass = 0
            SetCTerminus = 1
        Else
            SetCTerminus = 0
        End If
        .PrecedingResidue = FillResidueStructureUsingSymbol("")
        .FollowingResidue = FillResidueStructureUsingSymbol(strFollowingResidue, blnUse3LetterCode)
    End With
    
    UpdateResidueMasses
End Function

Public Function SetCTerminusGroup(eCTerminusGroup As ctgCTerminusGroupConstants, Optional strFollowingResidue As String = "", Optional blnUse3LetterCode As Boolean = True) As Long
    ' Returns 0 if success; 1 if error
    Dim lngError As Long
    
    lngError = 0
    Select Case eCTerminusGroup
    Case ctgHydroxyl:   lngError = SetCTerminus("OH", strFollowingResidue, blnUse3LetterCode)
    Case ctgAmide:      lngError = SetCTerminus("NH2", strFollowingResidue, blnUse3LetterCode)
    Case ctgNone:       lngError = SetCTerminus("", strFollowingResidue, blnUse3LetterCode)
    Case Else:          lngError = 1
    End Select
    
    SetCTerminusGroup = lngError
    
End Function

Public Sub SetDefaultModificationSymbols()
    
    On Error GoTo SetDefaultModificationSymbolsErrorHandler
    
    RemoveAllModificationSymbols
        
    ' Add the symbol for phosphorylation
    SetModificationSymbol "*", dblPhosphorylationMass, True, "Phosphorylation [HPO3]"
    
    ' Define the other default modifications
    ' Valid Mod Symbols are ! # $ % & ' * + ? ^ _ ` ~
    
    SetModificationSymbol "+", 14.01565, False, "Methylation [CH2]"
    SetModificationSymbol "@", 15.99492, False, "Oxidation [O]"
    SetModificationSymbol "!", 57.02146, False, "Carbamidomethylation [C2H3NO]"
    SetModificationSymbol "&", 58.00548, False, "Carboxymethylation [CH2CO2]"
    SetModificationSymbol "#", 71.03711, False, "Acrylamide [CHCH2CONH2]"
    SetModificationSymbol "$", 227.127, False, "Cleavable ICAT [(^12C10)H17N3O3]"
    SetModificationSymbol "%", 236.127, False, "Cleavable ICAT [(^13C9)(^12C)H17N3O3]"
    SetModificationSymbol "~", 442.225, False, "ICAT D0 [C20H34N4O5S]"
    SetModificationSymbol "`", 450.274, False, "ICAT D8 [C20H26D8N4O5S]"
    
    Exit Sub

SetDefaultModificationSymbolsErrorHandler:
    GeneralErrorHandler "MWPeptideClass.SetDefaultModificationSymbols", Err.Number
    
End Sub

Public Sub SetDefaultOptions()
    Dim intIonIndex As Integer
    
On Error GoTo SetDefaultOptionsErrorHandler

    With mFragSpectrumOptions
        With .IntensityOptions
            .IonType(itIonTypeConstants.itAIon) = 20
            .IonType(itIonTypeConstants.itBIon) = 100
            .IonType(itIonTypeConstants.itYIon) = 100
            .IonType(itIonTypeConstants.itCIon) = 100
            .IonType(itIonTypeConstants.itZIon) = 100
            .BYIonShoulder = 50
            .NeutralLoss = 20
        End With
        
        ' A ions can have ammonia and phosphate loss, but not water loss
        With .IonTypeOptions(itIonTypeConstants.itAIon)
            .ShowIon = True
            .NeutralLossAmmonia = True
            .NeutralLossPhosphate = True
            .NeutralLossWater = False
        End With
        
        For intIonIndex = itIonTypeConstants.itBIon To itIonTypeConstants.itZIon
            With .IonTypeOptions(intIonIndex)
                .ShowIon = True
                .NeutralLossAmmonia = True
                .NeutralLossPhosphate = True
                .NeutralLossWater = True
            End With
        Next intIonIndex
        
        .DoubleChargeIonsShow = True
        .DoubleChargeIonsThreshold = 800
        
        .TripleChargeIonsShow = False
        .TripleChargeIonsThreshold = 900
    End With
    
    SetSymbolWaterLoss "-H2O"
    SetSymbolAmmoniaLoss "-NH3"
    SetSymbolPhosphoLoss "-H3PO4"
    
    UpdateStandardMasses

    SetDefaultModificationSymbols
    
Exit Sub

SetDefaultOptionsErrorHandler:
    GeneralErrorHandler "MWPeptideClass.SetDefaultOptions", Err.Number
    
End Sub

Public Sub SetFragmentationSpectrumOptions(udtNewFragSpectrumOptions As udtFragmentationSpectrumOptionsType)
    mFragSpectrumOptions = udtNewFragSpectrumOptions
End Sub

Public Function SetModificationSymbol(ByVal strModSymbol As String, ByVal dblModificationMass As Double, blnIndicatesPhosphorylation As Boolean, strComment As String) As Long
    ' Adds a new modification or updates an existing one (based on strModSymbol)
    ' Returns 0 if successful, otherwise, returns -1
    
    Dim strTestChar As String
    Dim lngIndex As Long, lngIndexToUse As Long, lngErrorID As Long
    
    lngErrorID = 0
    If Len(strModSymbol) < 1 Then
        lngErrorID = -1
    Else
        ' Make sure strModSymbol contains no letters, numbers, spaces, dashes, or periods
        For lngIndex = 1 To Len(strModSymbol)
            strTestChar = Mid(strModSymbol, lngIndex, 1)
            If Not IsModSymbolInternal(strTestChar) Then
                lngErrorID = -1
            End If
        Next lngIndex
        
        If lngErrorID = 0 Then
            ' See if the modification is alrady present
            lngIndexToUse = GetModificationSymbolID(strModSymbol)
            
            If lngIndexToUse = 0 Then
                ' Need to add the modification
                ModificationSymbolCount = ModificationSymbolCount + 1
                lngIndexToUse = ModificationSymbolCount
                ReserveMemoryForModifications ModificationSymbolCount, True
            End If
            
            With ModificationSymbols(lngIndexToUse)
                .Symbol = strModSymbol
                .ModificationMass = dblModificationMass
                .IndicatesPhosphorylation = blnIndicatesPhosphorylation
                .Comment = strComment
            End With
        End If
    End If

    SetModificationSymbol = lngErrorID

End Function

Public Function SetNTerminus(strFormula As String, Optional strPrecedingResidue As String = "", Optional blnUse3LetterCode As Boolean = True) As Long
    ' Returns 0 if success; 1 if error
    
    ' Typical N terminus mods
        ' Hydrogen = H
        ' Acetyl = C2OH3
        ' PyroGlu = C5O2NH6
        ' Carbamyl = CONH2
        ' PTC = C7H6NS

    With mNTerminus
        .Formula = strFormula
        .Mass = ComputeFormulaWeight(.Formula)
        If .Mass < 0 Then
            .Mass = 0
            SetNTerminus = 1
        Else
            SetNTerminus = 0
        End If
        .PrecedingResidue = FillResidueStructureUsingSymbol(strPrecedingResidue, blnUse3LetterCode)
        .FollowingResidue = FillResidueStructureUsingSymbol("")
    End With
    
    UpdateResidueMasses
End Function

Public Function SetNTerminusGroup(eNTerminusGroup As ntgNTerminusGroupConstants, Optional strPrecedingResidue As String = "", Optional blnUse3LetterCode As Boolean = True) As Long
    ' Returns 0 if success; 1 if error
    Dim lngError As Long
    
    lngError = 0
    Select Case eNTerminusGroup
    Case ntgHydrogen:           lngError = SetNTerminus("H", strPrecedingResidue, blnUse3LetterCode)
    Case ntgHydrogenPlusProton: lngError = SetNTerminus("HH", strPrecedingResidue, blnUse3LetterCode)
    Case ntgAcetyl:             lngError = SetNTerminus("C2OH3", strPrecedingResidue, blnUse3LetterCode)
    Case ntgPyroGlu:            lngError = SetNTerminus("C5O2NH6", strPrecedingResidue, blnUse3LetterCode)
    Case ntgCarbamyl:           lngError = SetNTerminus("CONH2", strPrecedingResidue, blnUse3LetterCode)
    Case ntgPTC:                lngError = SetNTerminus("C7H6NS", strPrecedingResidue, blnUse3LetterCode)
    Case ntgNone:               lngError = SetNTerminus("", strPrecedingResidue, blnUse3LetterCode)
    Case Else:                  lngError = 1
    End Select
    
    SetNTerminusGroup = lngError
    
End Function

Public Function SetResidue(ByVal lngResidueNumber As Long, strSymbol As String, Optional blnIs3LetterCode As Boolean = True, Optional blnPhosphorylated As Boolean = False) As Long
    ' Sets or adds a residue (must add residues in order)
    ' Returns the index of the modified residue, or the new index if added
    ' Returns -1 if a problem
    
    Dim lngIndexToUse As Long, str3LetterSymbol As String
    
    If Len(strSymbol) = 0 Then
        SetResidue = -1
        Exit Function
    End If
    
    If lngResidueNumber > ResidueCount Then
        ResidueCount = ResidueCount + 1
        ReserveMemoryForResidues ResidueCount, True
        lngIndexToUse = ResidueCount
    Else
        lngIndexToUse = lngResidueNumber
    End If
    
    With Residues(lngIndexToUse)
        If blnIs3LetterCode Then
            str3LetterSymbol = strSymbol
        Else
            str3LetterSymbol = GetAminoAcidSymbolConversionInternal(strSymbol, True)
        End If
        
        If Len(str3LetterSymbol) = 0 Then
            .Symbol = UNKNOWN_SYMBOL
        Else
            .Symbol = str3LetterSymbol
        End If
        
        .Phosphorylated = blnPhosphorylated
        If blnPhosphorylated Then
            ' Only Ser, Thr, or Tyr should be phosphorylated
            ' However, if the user sets other residues as phosphorylated, we'll allow that
            Debug.Assert .Symbol = "Ser" Or .Symbol = "Thr" Or .Symbol = "Tyr"
        End If
        
        .ModificationIDCount = 0
    End With
    
    UpdateResidueMasses
    
    SetResidue = lngIndexToUse
End Function

Public Function SetResidueModifications(ByVal lngResidueNumber As Long, ByVal intModificationCount As Integer, ByRef lngModificationIDsOneBased() As Long) As Long
    ' Sets the modifications for a specific residue
    ' Modification Symbols are defined using successive calls to SetModificationSymbol()
    
    ' Returns 0 if modifications set; returns 1 if an error
    
    Dim intIndex As Integer, lngNewModID As Long
    
    If lngResidueNumber >= 1 And lngResidueNumber <= ResidueCount And intModificationCount >= 0 Then
        With Residues(lngResidueNumber)
            If intModificationCount > MAX_MODIFICATIONS Then
                intModificationCount = MAX_MODIFICATIONS
            End If
            
            .ModificationIDCount = 0
            .Phosphorylated = False
            For intIndex = 1 To intModificationCount
                lngNewModID = lngModificationIDsOneBased(intIndex)
                If lngNewModID >= 1 And lngNewModID <= ModificationSymbolCount Then
                    .ModificationIDs(.ModificationIDCount) = lngNewModID
                    
                    ' Check for phosphorylation
                    If ModificationSymbols(lngNewModID).IndicatesPhosphorylation Then
                        .Phosphorylated = True
                    End If
                    
                    .ModificationIDCount = .ModificationIDCount + 1
                End If
            Next intIndex
        
        End With
        
        SetResidueModifications = 0
    Else
        SetResidueModifications = 1
    End If

End Function

Public Function SetSequence(ByVal strSequence As String, Optional eNTerminus As ntgNTerminusGroupConstants = ntgHydrogen, Optional eCTerminus As ctgCTerminusGroupConstants = ctgHydroxyl, Optional blnIs3LetterCode As Boolean = True, Optional bln1LetterCheckForPrefixAndSuffixResidues As Boolean = True, Optional bln3LetterCheckForPrefixHandSuffixOH As Boolean = True, Optional blnAddMissingModificationSymbols As Boolean = False) As Long
    ' If blnIs3LetterCode = false, then look for sequence of the form: R.ABCDEF.R
    ' If found, remove the leading and ending residues since these aren't for this peptide
    ' Returns 0 if success or 1 if an error
    ' Will return 0 even in strSequence is blank or if it contains no valid residues
    
    Dim lngSequenceStrLength As Long, lngIndex As Long, lngModSymbolLength As Long
    Dim str1LetterSymbol As String, str3LetterSymbol As String, strFirstChar As String
    
    strSequence = Trim(strSequence)
    
    lngSequenceStrLength = Len(strSequence)
    If lngSequenceStrLength = 0 Then Exit Function
    
    ' Clear any old residue information
    ResidueCount = 0
    ReserveMemoryForResidues ResidueCount, False
    
    If Not blnIs3LetterCode Then
        ' Sequence is 1 letter codes
        
        If bln1LetterCheckForPrefixAndSuffixResidues Then
            ' First look if sequence is in the form A.BCDEFG.Z or -.BCDEFG.Z or A.BCDEFG.-
            ' If so, then need to strip out the preceding A and Z residues since they aren't really part of the sequence
            If lngSequenceStrLength > 1 And InStr(strSequence, ".") Then
                If Mid(strSequence, 2, 1) = "." Then
                    strSequence = Mid(strSequence, 3)
                    lngSequenceStrLength = Len(strSequence)
                End If
            
                If Mid(strSequence, lngSequenceStrLength - 1, 1) = "." Then
                    strSequence = Left(strSequence, lngSequenceStrLength - 2)
                    lngSequenceStrLength = Len(strSequence)
                End If
                
                ' Also check for starting with a . or ending with a .
                If Left(strSequence, 1) = "." Then
                    strSequence = Mid(strSequence, 2)
                End If
                
                If Right(strSequence, 1) = "." Then
                    strSequence = Left(strSequence, Len(strSequence) - 1)
                End If
                
                lngSequenceStrLength = Len(strSequence)
            End If
            
        End If
        
        For lngIndex = 1 To lngSequenceStrLength
            str1LetterSymbol = Mid(strSequence, lngIndex, 1)
            If IsCharacter(str1LetterSymbol) Then
                ' Character found
                ' Look up 3 letter symbol
                ' If none is found, this will return an empty string
                str3LetterSymbol = GetAminoAcidSymbolConversionInternal(str1LetterSymbol, True)
                
                If Len(str3LetterSymbol) = 0 Then str3LetterSymbol = UNKNOWN_SYMBOL
                
                SetSequenceAddResidue str3LetterSymbol
                
                ' Look at following character(s), and record any modification symbols present
                lngModSymbolLength = CheckForModifications(Mid(strSequence, lngIndex + 1), ResidueCount, blnAddMissingModificationSymbols)
                
                lngIndex = lngIndex + lngModSymbolLength
            Else
                ' If . or - or space, then ignore it
                ' If a number, ignore it
                ' If anything else, then should have been skipped, or should be skipped
                If str1LetterSymbol = "." Or str1LetterSymbol = "-" Or str1LetterSymbol = " " Then
                    ' All is fine; we can skip this
                Else
                    ' Ignore it
                End If
            End If
        Next lngIndex
        
    Else
        ' Sequence is 3 letter codes
        lngIndex = 1
        
        If bln3LetterCheckForPrefixHandSuffixOH Then
            ' Look for a leading H or trailing OH, provided those don't match any of the amino acids
            RemoveLeadingH strSequence
            RemoveTrailingOH strSequence
            
            ' Recompute sequence length
            lngSequenceStrLength = Len(strSequence)
        End If
        
        Do While lngIndex <= lngSequenceStrLength - 2
            strFirstChar = Mid(strSequence, lngIndex, 1)
            If IsCharacter(strFirstChar) Then
                If IsCharacter(Mid(strSequence, lngIndex + 1, 1)) And _
                   IsCharacter(Mid(strSequence, lngIndex + 2, 1)) Then
                 
                    str3LetterSymbol = UCase(strFirstChar) & LCase((Mid(strSequence, lngIndex + 1, 2)))
                    
                    If GetAbbreviationIDInternal(str3LetterSymbol, True) = 0 Then
                        ' 3 letter symbol not found
                        ' Add anyway, but mark as Xxx
                        str3LetterSymbol = UNKNOWN_SYMBOL
                    End If
                    
                    SetSequenceAddResidue str3LetterSymbol
                    
                    ' Look at following character(s), and record any modification symbols present
                    lngModSymbolLength = CheckForModifications(Mid(strSequence, lngIndex + 3), ResidueCount, blnAddMissingModificationSymbols)
                    
                    lngIndex = lngIndex + 3
                    lngIndex = lngIndex + lngModSymbolLength
                
                Else
                    ' First letter is a character, but next two are not; ignore it
                    lngIndex = lngIndex + 1
                End If
            Else
                ' If . or - or space, then ignore it
                ' If a number, ignore it
                ' If anything else, then should have been skipped or should be skipped
                If strFirstChar = "." Or strFirstChar = "-" Or strFirstChar = " " Then
                    ' All is fine; we can skip this
                Else
                    ' Ignore it
                End If
                lngIndex = lngIndex + 1
            End If
        Loop
    End If
    
    ' By calling SetNTerminus and SetCTerminus, the UpdateResidueMasses() Sub will also be called
    mDelayUpdateResidueMass = True
    SetNTerminusGroup eNTerminus
    SetCTerminusGroup eCTerminus
    
    mDelayUpdateResidueMass = False
    UpdateResidueMasses
    
    SetSequence = 0
    Exit Function
    
SetSequenceErrorHandler:
    SetSequence = AssureNonZero(Err.Number)
End Function

Private Sub SetSequenceAddResidue(str3LetterSymbol As String)
    
    If Len(str3LetterSymbol) = 0 Then
        str3LetterSymbol = UNKNOWN_SYMBOL
    End If
    
    ResidueCount = ResidueCount + 1
    ReserveMemoryForResidues ResidueCount, True

    With Residues(ResidueCount)
        .Symbol = str3LetterSymbol
        .Phosphorylated = False
        .ModificationIDCount = 0
    End With

End Sub

Public Sub SetSymbolAmmoniaLoss(strNewSymbol As String)
    If Len(strNewSymbol) > 0 Then
        mAmmoniaLossSymbol = strNewSymbol
    End If
End Sub

Public Sub SetSymbolPhosphoLoss(strNewSymbol As String)
    If Len(strNewSymbol) > 0 Then
        mPhosphoLossSymbol = strNewSymbol
    End If
End Sub

Public Sub SetSymbolWaterLoss(strNewSymbol As String)
    If Len(strNewSymbol) > 0 Then
        mWaterLossSymbol = strNewSymbol
    End If
End Sub


Private Sub ShellSortFragSpectrum(ByRef FragSpectrumWork() As udtFragmentationSpectrumDataType, ByRef PointerArray() As Long, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long)
    ' Sort the list using a shell sort
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim lngPointerSwap As Long

    ' Sort PointerArray[lngLowIndex..lngHighIndex] by comparing FragSpectrumWork(PointerArray(x)).Mass

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
                If FragSpectrumWork(PointerArray(lngIndexCompare)).Mass <= FragSpectrumWork(lngPointerSwap).Mass Then Exit For
                PointerArray(lngIndexCompare + lngIncrement) = PointerArray(lngIndexCompare)
            Next lngIndexCompare
            PointerArray(lngIndexCompare + lngIncrement) = lngPointerSwap
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop

End Sub

Private Sub UpdateResidueMasses()
    Dim lngIndex As Long, lngAbbrevID As Long
    Dim lngValidResidueCount As Long
    Dim intModIndex As Integer
    Dim dblRunningTotal As Double
    Dim blnPhosphorylationMassAdded As Boolean
    Dim blnProtonatedNTerminus As Boolean
    
    If mDelayUpdateResidueMass Then Exit Sub
    
    ' The N-terminus ions are the basis for the running total
    dblRunningTotal = mNTerminus.Mass
    If UCase(mNTerminus.Formula) = "HH" Then
        ' ntgHydrogenPlusProton; since we add back in the proton below when computing the fragment masses,
        '  we need to subtract it out here
        ' However, we need to subtract out dblHydrogenMass, and not dblChargeCarrierMass since the current
        '  formula's mass was computed using two hydrogens, and not one hydrogen and one charge carrier
        blnProtonatedNTerminus = True
        dblRunningTotal = dblRunningTotal - dblHydrogenMass
    End If
    
    For lngIndex = 1 To ResidueCount
        With Residues(lngIndex)
            lngAbbrevID = GetAbbreviationIDInternal(.Symbol, True)
            
            If lngAbbrevID > 0 Then
                lngValidResidueCount = lngValidResidueCount + 1
                .Mass = GetAbbreviationMass(lngAbbrevID)
                
                blnPhosphorylationMassAdded = False
                
                ' Compute the mass, including the modifications
                .MassWithMods = .Mass
                For intModIndex = 1 To .ModificationIDCount
                    If .ModificationIDs(intModIndex) <= ModificationSymbolCount Then
                        .MassWithMods = .MassWithMods + ModificationSymbols(.ModificationIDs(intModIndex)).ModificationMass
                        If ModificationSymbols(.ModificationIDs(intModIndex)).IndicatesPhosphorylation Then
                            blnPhosphorylationMassAdded = True
                        End If
                    Else
                        ' Invalid ModificationID
                        Debug.Assert False
                    End If
                Next intModIndex
                
                If .Phosphorylated Then
                    ' Only add a mass if none of the .ModificationIDs has .IndicatesPhosphorylation = True
                    If Not blnPhosphorylationMassAdded Then
                        .MassWithMods = .MassWithMods + dblPhosphorylationMass
                    End If
                End If
                
                dblRunningTotal = dblRunningTotal + .MassWithMods
                
                .IonMass(itIonTypeConstants.itAIon) = dblRunningTotal - dblImmoniumMassDifference - dblChargeCarrierMass
                .IonMass(itIonTypeConstants.itBIon) = dblRunningTotal
                
                ' Add NH3 (ammonia) to the B ion mass to get the C ion mass
                .IonMass(itIonTypeConstants.itCIon) = .IonMass(itIonTypeConstants.itBIon) + dblNH3Mass
            Else
                .Mass = 0
                .MassWithMods = 0
                Erase .IonMass()
            End If
        End With
    Next lngIndex
    
    dblRunningTotal = dblRunningTotal + mCTerminus.Mass
    If blnProtonatedNTerminus Then
        dblRunningTotal = dblRunningTotal + dblChargeCarrierMass
    End If
    
    If lngValidResidueCount > 0 Then
        mTotalMass = dblRunningTotal
    Else
        mTotalMass = 0
    End If
    
    ' Now compute the y-ion and z-ion masses
    dblRunningTotal = mCTerminus.Mass + dblChargeCarrierMass
    
    For lngIndex = ResidueCount To 1 Step -1
        With Residues(lngIndex)
            If .IonMass(itIonTypeConstants.itAIon) > 0 Then
                dblRunningTotal = dblRunningTotal + .MassWithMods
                .IonMass(itIonTypeConstants.itYIon) = dblRunningTotal + dblChargeCarrierMass
                If lngIndex = 1 Then
                    ' Add the N-terminus mass to highest y ion
                    .IonMass(itIonTypeConstants.itYIon) = .IonMass(itIonTypeConstants.itYIon) + mNTerminus.Mass - dblChargeCarrierMass
                    If blnProtonatedNTerminus Then
                        ' ntgHydrogenPlusProton; since we add back in the proton below when computing the fragment masses,
                        '  we need to subtract it out here
                        ' However, we need to subtract out dblHydrogenMass, and not dblChargeCarrierMass since the current
                        '  formula's mass was computed using two hydrogens, and not one hydrogen and one charge carrier
                        .IonMass(itIonTypeConstants.itYIon) = .IonMass(itIonTypeConstants.itYIon) - dblHydrogenMass
                    End If
                End If
                
                ' Subtract NH2 (amide) from the Y ion mass to get the Z ion mass
                .IonMass(itIonTypeConstants.itZIon) = .IonMass(itIonTypeConstants.itYIon) - (dblNH3Mass - dblHydrogenMass)
            End If
        End With
    Next lngIndex
    
End Sub

Public Sub UpdateStandardMasses()
    Dim eElementModeSaved As emElementModeConstants

On Error GoTo UpdateStandardMassesErrorHandler

    eElementModeSaved = GetElementModeInternal()
    
    SetElementModeInternal 2        ' emIsotopicMass
    
    dblChargeCarrierMass = GetChargeCarrierMassInternal()
    
    ' Update standard mass values
    dblHOHMass = ComputeFormulaWeight("HOH")
    dblNH3Mass = ComputeFormulaWeight("NH3")
    dblH3PO4Mass = ComputeFormulaWeight("H3PO4")
    dblHydrogenMass = ComputeFormulaWeight("H")
    
    ' Phosphorylation is the loss of OH and the addition of H2PO4, for a net change of HPO3
    dblPhosphorylationMass = ComputeFormulaWeight("HPO3")
    
    ' The immonium mass is equal to the mass of CO minus the mass of H, thus typically 26.9871
    dblImmoniumMassDifference = ComputeFormulaWeight("CO") - dblHydrogenMass

    dblHistidineFW = ComputeFormulaWeight("His")
    dblPhenylalanineFW = ComputeFormulaWeight("Phe")
    dblTyrosineFW = ComputeFormulaWeight("Tyr")

    SetElementModeInternal eElementModeSaved
    
Exit Sub

UpdateStandardMassesErrorHandler:
    GeneralErrorHandler "MWPeptideClass.UpdateStandardMasses", Err.Number
    
End Sub

Private Sub Class_Initialize()
    On Error GoTo InitPeptideClassErrorHandler
    
    ResidueCountDimmed = 0
    ResidueCount = 0
    ReserveMemoryForResidues 50, False
    
    ModificationSymbolCountDimmed = 0
    ModificationSymbolCount = 0
    ReserveMemoryForModifications 10, False
    
    SetDefaultOptions

    Exit Sub

InitPeptideClassErrorHandler:
    GeneralErrorHandler "MWPeptideClass.Class_Initialize", Err.Number
End Sub

