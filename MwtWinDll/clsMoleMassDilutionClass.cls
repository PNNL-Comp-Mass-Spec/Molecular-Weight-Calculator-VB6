VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MWMoleMassDilutionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Molecular Weight Calculator routines with ActiveX Class interfaces: MWMoleMassDilutionClass
' Written by Matthew Monroe in Richland, WA

' <EnumStatements>
Public Enum acdAutoComputeDilutionModeConstants
    acdFindRequiredDilutionVolumes = 0
    acdFindRequiredTotalVolume
    acdFindFinalConcentration
    acdFindInitialConcentration
End Enum

Public Enum acqAutoComputeQuantityModeConstants
    acqFindAmount = 0
    acqFindVolume
    acqFindConcentration
End Enum

Private Const AMOUNT_UNITS_LIST_COUNT = 22
Private Const AMOUNT_UNITS_VOLUME_INDEX_START = 13
Private Const AMOUNT_UNITS_VOLUME_INDEX_END = 21
Public Enum uamUnitsAmountConstants
    uamMoles = 0
    uamMillimoles
    uamMicroMoles
    uamNanoMoles
    uamPicoMoles
    uamFemtoMoles
    uamAttoMoles
    uamKilograms
    uamGrams
    uamMilligrams
    uamMicrograms
    uamPounds
    uamOunces
    uamLiters
    uamDeciLiters
    uamMilliLiters
    uamMicroLiters
    uamNanoLiters
    uamPicoLiters
    uamGallons
    uamQuarts
    uamPints
End Enum

Public Enum uevUnitsExtendedVolumeConstants
    uevL = 0
    uevDL
    uevML
    uevUL
    uevNL
    uevPL
    uevGallons
    uevQuarts
    uevPints
End Enum

Public Enum ummcUnitsMoleMassConcentrationConstants
    ummcMolar = 0
    ummcMilliMolar
    ummcMicroMolar
    ummcNanoMolar
    ummcPicoMolar
    ummcFemtoMolar
    ummcAttoMolar
    ummcMgPerDL
    ummcMgPerML
    ummcUgPerML
    ummcNgPerML
    ummcUgPerUL
    ummcNgPerUL
End Enum


' </EnumStatements>

' <UDT's>
Private Type udtMoleMassQuantityType
    Amount As Double                    ' In Moles
    Volume As Double                    ' In L
    Concentration As Double             ' In Molar
    SampleMass As Double                ' In g
    SampleDensity As Double             ' In g/mL
End Type

Private Type udtMoleMassDilutionValuesType
    InitialConcentration As Double                  ' In Molar
    StockSolutionVolume As Double                   ' In L
    FinalConcentration As Double                    ' In Molar
    DilutingSolventVolume As Double                 ' In L
    TotalFinalVolume As Double                      ' In L
End Type
' </UDT's>

Private Const POUNDS_PER_KG As Single = 2.20462262
Private Const GALLONS_PER_L As Single = 0.264172052

Private mQuantity As udtMoleMassQuantityType
Private mDilutionValues As udtMoleMassDilutionValuesType

Private mAutoComputeDilution As Boolean                                   ' When true, automatically compute dilution results whenever any value changes
Private mAutoComputeDilutionMode As acdAutoComputeDilutionModeConstants   ' The value to compute when mAutoComputeDilution is true

Private mAutoComputeQuantity As Boolean                                   ' When true, automatically compute quantities whenever any value changes
Private mAutoComputeQuantityMode As acqAutoComputeQuantityModeConstants   ' The value to compute when mAutoComputeQuantity is true

Private Sub CheckAutoComputeDilution()
    ' Checks if AutoCompute Dilution is enabled
    ' If yes, calls the appropriate function
    
    If mAutoComputeDilution Then
        Select Case mAutoComputeDilutionMode
        Case acdFindRequiredTotalVolume:        ComputeDilutionTotalVolume
        Case acdFindFinalConcentration:         ComputeDilutionFinalConcentration
        Case acdFindInitialConcentration:       ComputeDilutionInitialConcentration
        Case Else
            ' Includes acdFindRequiredDilutionVolumes
            ComputeDilutionRequiredStockAndDilutingSolventVolumes
        End Select
    End If
End Sub

Private Sub CheckAutoComputeQuantity()
    ' Checks if AutoCompute Quantity is enabled
    ' If yes, calls the appropriate function
    
    If mAutoComputeQuantity Then
        Select Case mAutoComputeQuantityMode
        Case acqFindVolume:             ComputeQuantityVolume
        Case acqFindConcentration:      ComputeQuantityConcentration
        Case Else
            ' Includes acqFindAmount
            ComputeQuantityAmount
        End Select
    End If
End Sub

Public Function ComputeDilutionFinalConcentration(Optional eUnits As ummcUnitsMoleMassConcentrationConstants = ummcMolar) As Double
    ' Computes the Final Concentration, storing in .FinalConcentration, and returning it

    With mDilutionValues
        If .TotalFinalVolume <> 0 Then
            .FinalConcentration = .InitialConcentration * .StockSolutionVolume / .TotalFinalVolume
        Else
            .TotalFinalVolume = 0
        End If
        
        .DilutingSolventVolume = .TotalFinalVolume - .StockSolutionVolume
        If .DilutingSolventVolume < 0 Then .DilutingSolventVolume = -1
    End With
    
    ComputeDilutionFinalConcentration = ConvertConcentration(mDilutionValues.FinalConcentration, ummcMolar, eUnits)

End Function

Public Function ComputeDilutionInitialConcentration(Optional eUnits As ummcUnitsMoleMassConcentrationConstants = ummcMolar) As Double
    ' Computes the Initial Concentration, storing in .InitialConcentration, and returning it
    
    With mDilutionValues
        If .StockSolutionVolume <> 0 Then
            .InitialConcentration = .FinalConcentration * .TotalFinalVolume / .StockSolutionVolume
        Else
            .InitialConcentration = 0
        End If
        
        .DilutingSolventVolume = .TotalFinalVolume - .StockSolutionVolume
        If .DilutingSolventVolume < 0 Then .DilutingSolventVolume = -1
    End With
    
    ComputeDilutionInitialConcentration = ConvertConcentration(mDilutionValues.InitialConcentration, ummcMolar, eUnits)

End Function

Public Function ComputeDilutionRequiredStockAndDilutingSolventVolumes(Optional eStockSolutionUnits As uevUnitsExtendedVolumeConstants = uevML, Optional ByRef dblNewDilutingSolventVolume As Double, Optional eDilutingSolventUnits As uevUnitsExtendedVolumeConstants = uevML) As Double
    ' Computes the required dilution volumes using initial concentration, final concentration
    '  and total final volume, storing in .StockSolutionVolume and .DilutingSolventVolume,
    '  and returning .StockSolutionVolume
    ' In addition, returns the .DilutingSolventVolume using the ByRef parameter dblNewDilutingSolventVolume
    
    With mDilutionValues
        If .InitialConcentration <> 0 Then
            .StockSolutionVolume = .FinalConcentration * .TotalFinalVolume / .InitialConcentration
        Else
            .StockSolutionVolume = 0
        End If
        
        .DilutingSolventVolume = .TotalFinalVolume - .StockSolutionVolume
        
        If .DilutingSolventVolume < 0 Then
            .DilutingSolventVolume = -1
            .StockSolutionVolume = -1
        End If
    End With
    
    dblNewDilutingSolventVolume = ConvertVolumeExtended(mDilutionValues.DilutingSolventVolume, uevL, eDilutingSolventUnits)
    
    ComputeDilutionRequiredStockAndDilutingSolventVolumes = ConvertVolumeExtended(mDilutionValues.StockSolutionVolume, uevL, eStockSolutionUnits)
    
End Function

Public Function ComputeDilutionTotalVolume(Optional eTotalVolumeUnits As uevUnitsExtendedVolumeConstants = uevML, Optional ByRef dblNewDilutingSolventVolume As Double, Optional eDilutingSolventUnits As uevUnitsExtendedVolumeConstants = uevML) As Double
    ' Compute the total volume following the dilution, storing in .TotalFinalVolume, and returning it
    
    With mDilutionValues
        If .InitialConcentration > 0 And .FinalConcentration > 0 Then
            .TotalFinalVolume = .InitialConcentration * .StockSolutionVolume / .FinalConcentration
            If .TotalFinalVolume < 0 Then
                .TotalFinalVolume = 0
            End If
        Else
            .TotalFinalVolume = 0
        End If
        
        .DilutingSolventVolume = .TotalFinalVolume - .StockSolutionVolume
        If .DilutingSolventVolume < 0 Then .DilutingSolventVolume = -1
    End With
    
    dblNewDilutingSolventVolume = ConvertVolumeExtended(mDilutionValues.DilutingSolventVolume, uevL, eDilutingSolventUnits)
    
    ComputeDilutionTotalVolume = ConvertVolumeExtended(mDilutionValues.TotalFinalVolume, uevL, eTotalVolumeUnits)

End Function

Public Function ComputeQuantityAmount(Optional eUnits As uamUnitsAmountConstants = uamMoles) As Double
    ' Computes .Amount using .Volume and .Concentration, storing the result
    '  in .Amount and returning it
    
    With mQuantity
        .Amount = .Concentration * .Volume
    End With

    ComputeQuantityAmount = ConvertAmount(mQuantity.Amount, uamMoles, eUnits)
End Function

Public Function ComputeQuantityConcentration(Optional eUnits As ummcUnitsMoleMassConcentrationConstants = ummcMolar) As Double
    ' Computes .Concentration using .Amount and .Volume, storing the result
    '  in .Concentration and returning it
    
    With mQuantity
        If .Volume <> 0 Then
            .Concentration = .Amount / .Volume
        Else
            .Concentration = 0
        End If
    End With

    ComputeQuantityConcentration = ConvertConcentration(mQuantity.Concentration, ummcMolar, eUnits)
End Function

Public Function ComputeQuantityVolume(Optional eUnits As uevUnitsExtendedVolumeConstants = uevL) As Double
    ' Computes .Volume using .Amount and .Concentration, storing the result
    '  in .Volume and returning it
    
    With mQuantity
        If .Concentration <> 0 Then
            .Volume = .Amount / .Concentration
        Else
            .Volume = 0
        End If
    End With

    ComputeQuantityVolume = ConvertVolumeExtended(mQuantity.Volume, uevL, eUnits)
End Function

' This function uses .SampleMass and .SampleDensity if the units are mass and/or volume-based
Public Function ConvertAmount(dblAmountIn As Double, eCurrentUnits As uamUnitsAmountConstants, eNewUnits As uamUnitsAmountConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    Dim dblSampleMass As Double, dblSampleDensity As Double
    Dim eCurrentVolumeUnits As uevUnitsExtendedVolumeConstants
    Dim eNewVolumeUnits As uevUnitsExtendedVolumeConstants
    
    If eCurrentUnits = eNewUnits Then
        ' No conversion, simply return dblAmountIn
        ConvertAmount = dblAmountIn
        Exit Function
    End If
    
    If eCurrentUnits >= AMOUNT_UNITS_VOLUME_INDEX_START And eCurrentUnits <= AMOUNT_UNITS_VOLUME_INDEX_END And _
       eNewUnits >= AMOUNT_UNITS_VOLUME_INDEX_START And eNewUnits <= AMOUNT_UNITS_VOLUME_INDEX_END Then
        ' Converting from one volume unit to another volume unit
        ' No need to explicitly specify mass or density
    
        eCurrentVolumeUnits = eCurrentUnits - AMOUNT_UNITS_VOLUME_INDEX_START
        eNewVolumeUnits = eNewUnits - AMOUNT_UNITS_VOLUME_INDEX_START
        
        ConvertAmount = ConvertVolumeExtended(dblAmountIn, eCurrentVolumeUnits, eNewVolumeUnits)
    Else
    
        dblSampleMass = mQuantity.SampleMass
        dblSampleDensity = mQuantity.SampleDensity
        
        dblFactor = FactorAmount(eCurrentUnits, dblSampleMass, dblSampleDensity)
        If dblFactor = -1 Then
            ConvertAmount = -1
            Exit Function
        Else
            dblValue = dblAmountIn * dblFactor
        End If
        
        dblFactor = FactorAmount(eNewUnits, dblSampleMass, dblSampleDensity)
        If dblFactor = -1 Or dblFactor = 0 Then
            ConvertAmount = -1
        Else
            ConvertAmount = dblValue / dblFactor
        End If
    End If
    
End Function

' Duplicated function, in both MWCapillaryFlowClass and MWMoleMassDilutionClass
Public Function ConvertConcentration(dblConcentrationIn As Double, eCurrentUnits As ummcUnitsMoleMassConcentrationConstants, eNewUnits As ummcUnitsMoleMassConcentrationConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    Dim dblSampleMass As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertConcentration = dblConcentrationIn
        Exit Function
    End If
    
    dblSampleMass = mQuantity.SampleMass
    
    dblFactor = FactorConcentration(eCurrentUnits, dblSampleMass)
    If dblFactor = -1 Then
        ConvertConcentration = -1
        Exit Function
    Else
        dblValue = dblConcentrationIn * dblFactor
    End If
    
    dblFactor = FactorConcentration(eNewUnits, dblSampleMass)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertConcentration = -1
    Else
        ConvertConcentration = dblValue / dblFactor
    End If

End Function

Public Function ConvertVolumeExtended(dblVolume As Double, eCurrentUnits As uevUnitsExtendedVolumeConstants, eNewUnits As uevUnitsExtendedVolumeConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertVolumeExtended = dblVolume
        Exit Function
    End If
    
    dblFactor = FactorVolumeExtended(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertVolumeExtended = -1
        Exit Function
    Else
        dblValue = dblVolume * dblFactor
    End If
    
    dblFactor = FactorVolumeExtended(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertVolumeExtended = -1
    Else
        ConvertVolumeExtended = dblValue / dblFactor
    End If

End Function

' Multiplication factor for converting from eUnits to Moles
' dblSampleMass is required for mass-based units
' In addition is required for volume-based units
Private Function FactorAmount(eUnits As uamUnitsAmountConstants, Optional dblSampleMass As Double = -1, Optional dblSampleDensity As Double = 0) As Double
    Dim dblFactor As Double
    
    If dblSampleMass = 0 Then
        dblFactor = -1
    Else
        ' Determine the Amount multiplication dblFactor
        Select Case eUnits
        Case uamMoles:          dblFactor = 1
        Case uamMillimoles:     dblFactor = 1 / 1000#
        Case uamMicroMoles:     dblFactor = 1 / 1000000#
        Case uamNanoMoles:      dblFactor = 1 / 1000000000#
        Case uamPicoMoles:      dblFactor = 1 / 1000000000000#
        Case uamFemtoMoles:     dblFactor = 1 / 1E+15
        Case uamAttoMoles:      dblFactor = 1 / 1E+18
        Case uamKilograms:      dblFactor = 1000# / dblSampleMass
        Case uamGrams:          dblFactor = 1 / dblSampleMass
        Case uamMilligrams:     dblFactor = 1 / (dblSampleMass * 1000#)
        Case uamMicrograms:     dblFactor = 1 / (dblSampleMass * 1000000#)
        Case uamPounds:         dblFactor = 1000# / (dblSampleMass * POUNDS_PER_KG)
        Case uamOunces:         dblFactor = 1000# / (dblSampleMass * POUNDS_PER_KG * 16)
        Case uamLiters:         dblFactor = dblSampleDensity / dblSampleMass * FactorVolumeExtended(uevL)
        Case uamDeciLiters:     dblFactor = dblSampleDensity / dblSampleMass * FactorVolumeExtended(uevDL)
        Case uamMilliLiters:    dblFactor = dblSampleDensity / dblSampleMass * FactorVolumeExtended(uevML)
        Case uamMicroLiters:    dblFactor = dblSampleDensity / dblSampleMass * FactorVolumeExtended(uevUL)
        Case uamNanoLiters:     dblFactor = dblSampleDensity / dblSampleMass * FactorVolumeExtended(uevNL)
        Case uamPicoLiters:     dblFactor = dblSampleDensity / dblSampleMass * FactorVolumeExtended(uevPL)
        Case uamGallons:        dblFactor = dblSampleDensity / dblSampleMass * FactorVolumeExtended(uevGallons)
        Case uamQuarts:         dblFactor = dblSampleDensity / dblSampleMass * FactorVolumeExtended(uevQuarts)
        Case uamPints:          dblFactor = dblSampleDensity / dblSampleMass * FactorVolumeExtended(uevPints)
        Case Else:              dblFactor = -1
        End Select
    End If
    
    FactorAmount = dblFactor

End Function

' Multiplication factor for converting from eUnits to M
' dblSampleMass is required for mass-based units
' Duplicated function, in both MWCapillaryFlowClass and MWMoleMassDilutionClass
Private Function FactorConcentration(eUnits As ummcUnitsMoleMassConcentrationConstants, Optional ByVal dblSampleMass As Double = 0) As Double
    Dim dblFactor As Double
    
    If dblSampleMass = 0 Then
        dblFactor = -1
    Else
        Select Case eUnits
        Case ummcMolar:      dblFactor = 1#
        Case ummcMilliMolar: dblFactor = 1 / 1000#
        Case ummcMicroMolar: dblFactor = 1 / 1000000#
        Case ummcNanoMolar:  dblFactor = 1 / 1000000000#
        Case ummcPicoMolar:  dblFactor = 1 / 1000000000000#
        Case ummcFemtoMolar: dblFactor = 1 / 1E+15
        Case ummcAttoMolar:  dblFactor = 1 / 1E+18
        Case ummcMgPerDL:    dblFactor = 1 / dblSampleMass / 100#        '1/[(1 g / 1000 mg) * (1 / MW) * (10 dL/L)]
        Case ummcMgPerML:    dblFactor = 1 / dblSampleMass              '1/[(1 g / 1000 mg) * (1 / MW) * (1000 mL/L)]
        Case ummcUgPerML:    dblFactor = 1 / (dblSampleMass * 1000#)    '1/[(1 g / 1000000 ug) * (1 / MW) * (1000 mL/L)]
        Case ummcNgPerML:    dblFactor = 1 / (dblSampleMass * 1000000#) '1/[(1 g / 1000000000 ng) * (1 / MW) * (1000 mL/L)]
        Case ummcUgPerUL:    dblFactor = 1 / (dblSampleMass)            '1/[(1 g / 1000000 ug) * (1 / MW) * (1000000 uL/L)]
        Case ummcNgPerUL:    dblFactor = 1 / (dblSampleMass * 1000#)    '1/[(1 g / 1000000000 ng) * (1 / MW) * (1000000 uL/L)]
        Case Else:          dblFactor = -1
        End Select
    End If
    
    FactorConcentration = dblFactor
End Function

' Multiplication factor for converting from eUnits to mL
' An extended version of the FactorVolume function in MWCapillaryFlowClass
Private Function FactorVolumeExtended(eUnits As uevUnitsExtendedVolumeConstants) As Double
    
    ' Note: 4 quarts per gallon, 2 pints per quart
    Select Case eUnits
    Case uevL:       FactorVolumeExtended = 1 * 1000#
    Case uevDL:      FactorVolumeExtended = 1 * 100#
    Case uevML:      FactorVolumeExtended = 1#
    Case uevUL:      FactorVolumeExtended = 1 / 1000#
    Case uevNL:      FactorVolumeExtended = 1 / 1000000#
    Case uevPL:      FactorVolumeExtended = 1 / 1000000000#
    Case uevGallons: FactorVolumeExtended = 1000# / GALLONS_PER_L
    Case uevQuarts:  FactorVolumeExtended = 1000# / GALLONS_PER_L / 4#
    Case uevPints:   FactorVolumeExtended = 1000# / GALLONS_PER_L / 8#
    Case Else:       FactorVolumeExtended = -1
    End Select

End Function


' Get Subs
' Gets the most recently computed value
' If mAutoComputeDilution = False, then must manually call a Compute Sub to recompute the value
' Similarly, if mAutoComputeQuantity = False, then must manually call a Compute Sub to recompute the value

Public Function GetAutoComputeDilutionEnabled() As Boolean
    GetAutoComputeDilutionEnabled = mAutoComputeDilution
End Function

Public Function GetAutoComputeDilutionMode() As acdAutoComputeDilutionModeConstants
    GetAutoComputeDilutionMode = mAutoComputeDilutionMode
End Function

Public Function GetAutoComputeQuantityEnabled() As Boolean
    GetAutoComputeQuantityEnabled = mAutoComputeQuantity
End Function

Public Function GetAutoComputeQuantityMode() As acqAutoComputeQuantityModeConstants
    GetAutoComputeQuantityMode = mAutoComputeQuantityMode
End Function

Public Function GetDilutionFinalConcentration(Optional eUnits As ummcUnitsMoleMassConcentrationConstants = ummcMolar) As Double
     GetDilutionFinalConcentration = ConvertConcentration(mDilutionValues.FinalConcentration, ummcMolar, eUnits)
End Function

Public Function GetDilutionInitialConcentration(Optional eUnits As ummcUnitsMoleMassConcentrationConstants = ummcMolar) As Double
     GetDilutionInitialConcentration = ConvertConcentration(mDilutionValues.InitialConcentration, ummcMolar, eUnits)
End Function

Public Function GetDilutionTotalFinalVolume(Optional eUnits As uevUnitsExtendedVolumeConstants = uevML) As Double
    GetDilutionTotalFinalVolume = ConvertVolumeExtended(mDilutionValues.TotalFinalVolume, uevL, eUnits)
End Function

Public Function GetDilutionVolumeDilutingSolvent(Optional eUnits As uevUnitsExtendedVolumeConstants = uevML) As Double
    GetDilutionVolumeDilutingSolvent = ConvertVolumeExtended(mDilutionValues.DilutingSolventVolume, uevL, eUnits)
End Function

Public Function GetDilutionVolumeStockSolution(Optional eUnits As uevUnitsExtendedVolumeConstants = uevML) As Double
    GetDilutionVolumeStockSolution = ConvertVolumeExtended(mDilutionValues.StockSolutionVolume, uevL, eUnits)
End Function


Public Function GetQuantityAmount(Optional eUnits As uamUnitsAmountConstants = uamMoles) As Double
    GetQuantityAmount = ConvertAmount(mQuantity.Amount, uamMoles, eUnits)
End Function

Public Function GetQuantityConcentration(Optional eUnits As ummcUnitsMoleMassConcentrationConstants = ummcMolar) As Double
     GetQuantityConcentration = ConvertConcentration(mQuantity.Concentration, ummcMolar, eUnits)
End Function

Public Function GetQuantityVolume(Optional eUnits As uevUnitsExtendedVolumeConstants = uevML) As Double
    GetQuantityVolume = ConvertVolumeExtended(mQuantity.Volume, uevL, eUnits)
End Function

Public Function GetSampleDensity() As Double
    GetSampleDensity = mQuantity.SampleDensity
End Function

Public Function GetSampleMass() As Double
    GetSampleMass = mQuantity.SampleMass
End Function


' Set Subs
' If mAutoComputeDilution = False, then must manually call a Compute Sub to recompute the value
' Similarly, if mAutoComputeQuantity = False, then must manually call a Compute Sub to recompute the value

Public Sub SetAutoComputeDilutionEnabled(blnAutoCompute As Boolean)
    mAutoComputeDilution = blnAutoCompute
End Sub

Public Sub SetAutoComputeDilutionMode(eAutoComputeMode As acdAutoComputeDilutionModeConstants)
    If eAutoComputeMode >= acdFindRequiredDilutionVolumes And eAutoComputeMode <= acdFindFinalConcentration Then
        mAutoComputeDilutionMode = eAutoComputeMode
    End If
End Sub

Public Sub SetAutoComputeQuantityEnabled(blnAutoCompute As Boolean)
    mAutoComputeQuantity = blnAutoCompute
End Sub

Public Sub SetAutoComputeQuantityMode(eAutoComputeMode As acqAutoComputeQuantityModeConstants)
    If eAutoComputeMode >= acqFindAmount And eAutoComputeMode <= acqFindConcentration Then
        mAutoComputeQuantityMode = eAutoComputeMode
    End If
End Sub


Public Sub SetDilutionFinalConcentration(dblConcentration As Double, Optional eUnits As ummcUnitsMoleMassConcentrationConstants = ummcMolar)
    mDilutionValues.FinalConcentration = ConvertConcentration(dblConcentration, eUnits, ummcMolar)
    CheckAutoComputeDilution
End Sub

Public Sub SetDilutionInitialConcentration(dblConcentration As Double, Optional eUnits As ummcUnitsMoleMassConcentrationConstants = ummcMolar)
     mDilutionValues.InitialConcentration = ConvertConcentration(dblConcentration, eUnits, ummcMolar)
    CheckAutoComputeDilution
End Sub

Public Sub SetDilutionTotalFinalVolume(dblVolume As Double, Optional eUnits As uevUnitsExtendedVolumeConstants = uevML)
    mDilutionValues.TotalFinalVolume = ConvertVolumeExtended(dblVolume, eUnits, uevL)
    CheckAutoComputeDilution
End Sub

Public Sub SetDilutionVolumeDilutingSolvent(dblVolume As Double, Optional eUnits As uevUnitsExtendedVolumeConstants = uevML)
    mDilutionValues.DilutingSolventVolume = ConvertVolumeExtended(dblVolume, eUnits, uevL)
    CheckAutoComputeDilution
End Sub

Public Sub SetDilutionVolumeStockSolution(dblVolume As Double, Optional eUnits As uevUnitsExtendedVolumeConstants = uevML)
    mDilutionValues.StockSolutionVolume = ConvertVolumeExtended(dblVolume, eUnits, uevL)
    CheckAutoComputeDilution
End Sub


Public Sub SetQuantityAmount(dblAmount As Double, Optional eUnits As uamUnitsAmountConstants = uamMoles)
    mQuantity.Amount = ConvertAmount(dblAmount, eUnits, uamMoles)
    CheckAutoComputeQuantity
End Sub

Public Sub SetQuantityConcentration(dblConcentration As Double, Optional eUnits As ummcUnitsMoleMassConcentrationConstants = ummcMolar)
     mQuantity.Concentration = ConvertConcentration(dblConcentration, eUnits, ummcMolar)
     CheckAutoComputeQuantity
End Sub

Public Sub SetQuantityVolume(dblVolume As Double, Optional eUnits As uevUnitsExtendedVolumeConstants = uevML)
    mQuantity.Volume = ConvertVolumeExtended(dblVolume, eUnits, uevL)
End Sub

Public Function SetSampleDensity(dblDensityInGramsPerML As Double) As Double
    If dblDensityInGramsPerML >= 0 Then
        mQuantity.SampleDensity = dblDensityInGramsPerML
    Else
        mQuantity.SampleDensity = 0
    End If
    CheckAutoComputeQuantity
End Function

Public Function SetSampleMass(dblMassInGramsPerMole As Double) As Double
    If dblMassInGramsPerMole >= 0 Then
        mQuantity.SampleMass = dblMassInGramsPerMole
    Else
        mQuantity.SampleMass = 0
    End If
    CheckAutoComputeQuantity
End Function


Private Sub Class_Initialize()
    Me.SetAutoComputeDilutionEnabled False
    Me.SetAutoComputeQuantityEnabled False
    
    Me.SetAutoComputeDilutionMode acdFindRequiredDilutionVolumes
    
    Me.SetAutoComputeQuantityMode acqFindConcentration
    
    Me.SetQuantityAmount 1, uamMoles
    Me.SetQuantityVolume 100, uevML
    Me.SetQuantityConcentration 1, ummcMolar
    
    Me.SetDilutionInitialConcentration 10, ummcMolar
    Me.SetDilutionVolumeStockSolution 3, uevML
    Me.SetDilutionFinalConcentration 2, ummcMolar
    Me.SetDilutionVolumeDilutingSolvent 12, uevML
    Me.SetDilutionTotalFinalVolume 15, uevML
    
    ' Recompute
    Me.ComputeQuantityAmount
    Me.ComputeDilutionRequiredStockAndDilutingSolventVolumes
    
    Me.SetAutoComputeDilutionEnabled True
    Me.SetAutoComputeQuantityEnabled True

End Sub

Public Property Get AmountsUnitListCount() As Integer
    AmountsUnitListCount = AMOUNT_UNITS_LIST_COUNT
End Property

Public Property Get AmountsUnitListVolumeIndexStart() As Integer
    AmountsUnitListVolumeIndexStart = AMOUNT_UNITS_VOLUME_INDEX_START
End Property

Public Property Get AmountsUnitListVolumeIndexEnd() As Integer
    AmountsUnitListVolumeIndexEnd = AMOUNT_UNITS_VOLUME_INDEX_END
End Property

